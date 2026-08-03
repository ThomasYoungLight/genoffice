import { httpBodyDetail } from './http-error'
import { DIRECT_BASE_URLS } from './providers'
import type { AiProviderConfig, AiProviderId } from './types'

/**
 * Image generation with the user's own key.
 *
 * The suite's existing `generate_image` runs through the Genspark service,
 * which is no use to someone using their own provider — so this is the
 * provider-native path: OpenAI's `/v1/images/generations`, returning the bytes
 * rather than a hosted URL, so nothing leaves the machine but the prompt.
 *
 * Runs in the main process like every other keyed request; renderers ask over
 * IPC and get base64 back.
 */

/** stop waiting on a stalled request; image generation is slow but not this slow */
const IMAGE_TIMEOUT_MS = 180_000

/**
 * The model used when the settings file names none. Deliberately a plain
 * default rather than a hardcoded rule: `providers.<id>.imageModel` overrides
 * it, so a newer model needs no code change.
 */
export const DEFAULT_IMAGE_MODELS: Partial<Record<AiProviderId, string>> = {
  openai: 'gpt-image-2',
}

/** whether this provider has an image endpoint this build knows how to call */
export function providerGeneratesImages(provider: AiProviderId): boolean {
  return provider in DEFAULT_IMAGE_MODELS
}

export function imageModelFor(provider: AiProviderId, configured?: string | undefined): string {
  return configured?.trim() || DEFAULT_IMAGE_MODELS[provider] || ''
}

export interface ImageGenRequest {
  prompt: string
  /** e.g. '1024x1024', '1536x1024', '1024x1536', or 'auto' */
  size?: string | undefined
  /** 'low' | 'medium' | 'high' | 'auto' — the cost/latency dial */
  quality?: string | undefined
  /** cut the background out (png only); for logos and overlays */
  transparent?: boolean | undefined
}

export interface ImageGenResult {
  ok: boolean
  /** raw base64, no `data:` prefix — the shape the insert paths already take */
  base64?: string
  mime?: string
  /** what the model was actually asked for, for the activity line */
  model?: string
  error?: string
}

interface OpenAiImageResponse {
  data?: Array<{ b64_json?: string; url?: string }>
  error?: { message?: string }
}

/**
 * Generate one image. `config.model` is ignored — the chat model and the image
 * model are different things, and the caller resolves the latter with
 * `imageModelFor` before getting here.
 */
export async function generateProviderImage(
  provider: AiProviderId,
  config: AiProviderConfig,
  model: string,
  request: ImageGenRequest,
  signal?: AbortSignal,
): Promise<ImageGenResult> {
  if (!providerGeneratesImages(provider)) {
    return { ok: false, error: `${provider} has no image generation endpoint` }
  }
  if (!config.apiKey) return { ok: false, error: 'No API key configured' }
  if (!model) return { ok: false, error: 'No image model configured' }
  const prompt = request.prompt.trim()
  if (!prompt) return { ok: false, error: 'The prompt is empty' }

  const base = (config.baseUrl || DIRECT_BASE_URLS.openai).replace(/\/$/, '')
  const url = `${base}/images/generations`
  const timeout = AbortSignal.timeout(IMAGE_TIMEOUT_MS)
  const abort = signal ? AbortSignal.any([signal, timeout]) : timeout

  /**
   * The optional fields are the ones older image models reject (`dall-e-3`
   * knows none of them). Rather than branch on the model id — the same
   * mistake the chat path had to unlearn — the full request is tried and a 400
   * naming one of them is retried with the minimum every image API accepts.
   */
  const full: Record<string, unknown> = {
    model,
    prompt,
    n: 1,
    ...(request.size ? { size: request.size } : {}),
    ...(request.quality ? { quality: request.quality } : {}),
    ...(request.transparent ? { background: 'transparent', output_format: 'png' } : {}),
  }
  const minimal = { model, prompt, n: 1, ...(request.size ? { size: request.size } : {}) }

  const post = (body: unknown) =>
    fetch(url, {
      method: 'POST',
      signal: abort,
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${config.apiKey}` },
      body: JSON.stringify(body),
    })

  try {
    let response = await post(full)
    if (response.status === 400 && Object.keys(full).length > Object.keys(minimal).length) {
      const text = await response.clone().text()
      if (/quality|background|output_format|response_format/i.test(text)) {
        response = await post(minimal)
      }
    }
    if (!response.ok) {
      return { ok: false, error: await describeImageFailure(response) }
    }
    const json = (await response.json()) as OpenAiImageResponse
    const first = json.data?.[0]
    if (first?.b64_json) {
      return {
        ok: true,
        base64: first.b64_json,
        mime: request.transparent ? 'image/png' : 'image/png',
        model,
      }
    }
    // dall-e-3 and some compatible servers hand back a hosted URL instead
    if (first?.url) {
      const image = await fetch(first.url, { signal: abort })
      if (!image.ok) return { ok: false, error: `image download failed: HTTP ${image.status}` }
      const bytes = Buffer.from(await image.arrayBuffer())
      const mime = image.headers.get('content-type')?.split(';')[0] || 'image/png'
      return { ok: true, base64: bytes.toString('base64'), mime, model }
    }
    return { ok: false, error: json.error?.message ?? 'the provider returned no image' }
  } catch (err) {
    if (err instanceof Error && err.name === 'TimeoutError') {
      return { ok: false, error: `No image within ${IMAGE_TIMEOUT_MS / 1000}s` }
    }
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}

async function describeImageFailure(response: Response): Promise<string> {
  const detail = httpBodyDetail(await response.text().catch(() => ''))
  const hint =
    response.status === 401 || response.status === 403
      ? 'the API key was rejected'
      : response.status === 404
        ? 'the image model was not found for this account'
        : response.status === 429
          ? 'rate limited or out of quota'
          : ''
  return `HTTP ${response.status}${hint ? ` (${hint})` : ''}${detail ? `: ${detail}` : ''}`
}

/** image-capable model ids this key can reach, for the settings UI and for diagnosis */
export function isImageModelId(id: string): boolean {
  return /(^|-)image(-|$)|^dall-e|image-\d/i.test(id)
}
