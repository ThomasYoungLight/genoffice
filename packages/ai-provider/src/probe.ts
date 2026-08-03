import { httpBodyDetail } from './http-error'
import { isLocalCliProvider, resolveProviderWire } from './providers'
import type { AiProviderConfig, AiProviderId, ProviderProtocol } from './types'

/**
 * Two read-only checks the settings dialog runs against a provider:
 *
 * - `testProvider` — does this key actually work with this model at this
 *   endpoint? Cheaper and far clearer than finding out on the user's first
 *   real prompt, where a bad key surfaces as a failed document edit.
 * - `listProviderModels` — what can this key reach right now? A hardcoded
 *   catalog goes stale the moment a vendor ships, so the dialog offers to ask.
 *
 * Both run in the main process: renderers cannot reach these hosts (CORS) and
 * must never hold the key anyway.
 */

/** stop waiting on an unresponsive endpoint rather than hanging the dialog */
const PROBE_TIMEOUT_MS = 20_000

function withTimeout(signal?: AbortSignal): AbortSignal {
  const timeout = AbortSignal.timeout(PROBE_TIMEOUT_MS)
  return signal ? AbortSignal.any([signal, timeout]) : timeout
}

function authHeaders(protocol: ProviderProtocol, apiKey: string): Record<string, string> {
  switch (protocol) {
    case 'anthropic':
      return {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        // main-process fetch goes through Chromium's network stack, which
        // Anthropic rejects with 403 without this opt-in (same as chat/stream)
        'anthropic-dangerous-direct-browser-access': 'true',
      }
    case 'gemini':
      return { 'x-goog-api-key': apiKey }
    default:
      return { Authorization: `Bearer ${apiKey}` }
  }
}

/** turn a failed response into something a user can act on */
async function describeFailure(response: Response): Promise<string> {
  const detail = httpBodyDetail(await response.text().catch(() => ''))
  const hint =
    response.status === 401 || response.status === 403
      ? 'the API key was rejected'
      : response.status === 404
        ? 'the model or endpoint was not found'
        : response.status === 429
          ? 'rate limited or out of quota'
          : ''
  return `HTTP ${response.status}${hint ? ` (${hint})` : ''}${detail ? `: ${detail}` : ''}`
}

function describeNetworkError(err: unknown): string {
  if (err instanceof Error && err.name === 'TimeoutError') {
    return `No response within ${PROBE_TIMEOUT_MS / 1000}s`
  }
  // the real cause of a main-process fetch failure hides in `cause`
  const e = err as { message?: unknown; cause?: { code?: unknown; message?: unknown } } | null
  const cause = e?.cause ? ` (${String(e.cause.code || e.cause.message)})` : ''
  return `${e?.message ?? String(err)}${cause}`
}

export interface ProviderProbeResult {
  ok: boolean
  error?: string
}

/**
 * Send the smallest possible generation request: a one-token completion. It
 * costs a fraction of a cent and, unlike merely listing models, proves the key,
 * the endpoint and the chosen model id all work together.
 */
export async function testProvider(
  provider: AiProviderId,
  config: AiProviderConfig,
  signal?: AbortSignal,
): Promise<ProviderProbeResult> {
  if (!config.apiKey) return { ok: false, error: 'No API key configured' }
  if (!config.model) return { ok: false, error: 'No model configured' }

  let wire: ReturnType<typeof resolveProviderWire>
  try {
    wire = resolveProviderWire(provider, config)
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
  const base = wire.baseUrl.replace(/\/$/, '')
  const headers = {
    'Content-Type': 'application/json',
    ...authHeaders(wire.protocol, config.apiKey),
  }

  const request = (): { url: string; body: unknown } => {
    switch (wire.protocol) {
      case 'anthropic':
        return {
          url: `${base}/v1/messages`,
          body: { model: config.model, max_tokens: 1, messages: [{ role: 'user', content: 'hi' }] },
        }
      case 'gemini':
        return {
          url: `${base}/models/${config.model}:generateContent`,
          body: {
            contents: [{ role: 'user', parts: [{ text: 'hi' }] }],
            generationConfig: { maxOutputTokens: 1 },
          },
        }
      case 'openai-responses':
        // test the endpoint the app will actually use, not a sibling that may
        // accept a key/model pair this one rejects
        return {
          url: `${base}/responses`,
          body: {
            model: config.model,
            input: [{ role: 'user', content: [{ type: 'input_text', text: 'hi' }] }],
            max_output_tokens: 16,
          },
        }
      default:
        return {
          url: `${base}/chat/completions`,
          body: { model: config.model, max_tokens: 1, messages: [{ role: 'user', content: 'hi' }] },
        }
    }
  }

  const { url, body } = request()
  try {
    let response = await fetch(url, {
      method: 'POST',
      signal: withTimeout(signal),
      headers,
      body: JSON.stringify(body),
    })
    // Newer OpenAI models reject `max_tokens` in favour of
    // `max_completion_tokens`. That is a quirk of the probe, not a broken key,
    // so retry once without the cap rather than reporting a false failure.
    if (!response.ok && response.status === 400 && wire.protocol === 'openai') {
      const text = await response.clone().text()
      if (/max_tokens|max_completion_tokens/i.test(text)) {
        response = await fetch(url, {
          method: 'POST',
          signal: withTimeout(signal),
          headers,
          body: JSON.stringify({
            model: config.model,
            messages: [{ role: 'user', content: 'hi' }],
          }),
        })
      }
    }
    if (!response.ok) return { ok: false, error: await describeFailure(response) }
    return { ok: true }
  } catch (err) {
    return { ok: false, error: describeNetworkError(err) }
  }
}

/**
 * Families an OpenAI-compatible `/models` listing returns that cannot serve a
 * chat turn. Matching is deliberately conservative — anything unrecognised is
 * kept, since a wrongly hidden model is worse than one extra row, and the
 * model field accepts free text regardless.
 */
const NON_CHAT_MODEL =
  /embedding|whisper|^tts|-tts|dall-e|moderation|transcribe|^sora|-image|image-\d|-instruct/i

export interface ModelListResult {
  ok: boolean
  models?: string[]
  error?: string
}

/** Ask the provider which models this key can reach. */
export async function listProviderModels(
  provider: AiProviderId,
  config: AiProviderConfig,
  signal?: AbortSignal,
): Promise<ModelListResult> {
  if (provider === 'genspark') {
    // The proxy renames models to its own scheme and exposes no catalogue; the
    // curated list in AI_PROVIDERS is the source of truth for it.
    return { ok: false, error: 'Genspark does not publish a model list' }
  }
  if (isLocalCliProvider(provider)) {
    // a CLI has no endpoint to query: it publishes its catalogue over its own
    // machine protocol, which @genoffice/ai-cli speaks (listCliModels). Callers
    // branch before reaching here, so this is a guard rather than a path.
    return { ok: false, error: 'A local CLI is listed through its own protocol' }
  }
  if (!config.apiKey) return { ok: false, error: 'No API key configured' }

  let wire: ReturnType<typeof resolveProviderWire>
  try {
    wire = resolveProviderWire(provider, config)
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
  const base = wire.baseUrl.replace(/\/$/, '')
  const url =
    wire.protocol === 'anthropic'
      ? `${base}/v1/models?limit=1000`
      : wire.protocol === 'gemini'
        ? `${base}/models?pageSize=1000`
        : `${base}/models`

  try {
    const response = await fetch(url, {
      signal: withTimeout(signal),
      headers: authHeaders(wire.protocol, config.apiKey),
    })
    if (!response.ok) return { ok: false, error: await describeFailure(response) }
    const json: unknown = await response.json()
    return { ok: true, models: parseModelList(wire.protocol, json) }
  } catch (err) {
    return { ok: false, error: describeNetworkError(err) }
  }
}

function parseModelList(protocol: ProviderProtocol, json: unknown): string[] {
  const ids: string[] = []
  if (protocol === 'gemini') {
    const body = json as {
      models?: Array<{ name?: string; supportedGenerationMethods?: string[] }>
    }
    for (const model of body.models ?? []) {
      // the listing includes embedding and legacy endpoints; keep what can chat
      const methods = model.supportedGenerationMethods
      if (methods && !methods.includes('generateContent')) continue
      // ids come back namespaced as "models/gemini-2.5-pro"
      if (model.name) ids.push(model.name.replace(/^models\//, ''))
    }
  } else {
    // both the Anthropic and OpenAI listings use { data: [{ id }] }
    const body = json as { data?: Array<{ id?: string }> }
    for (const model of body.data ?? []) {
      if (!model.id) continue
      if (protocol !== 'anthropic' && NON_CHAT_MODEL.test(model.id)) continue
      ids.push(model.id)
    }
  }
  return [...new Set(ids)].sort()
}
