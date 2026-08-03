/**
 * AI IPC for the slides main process, extracted from slides-main.ts:
 * settings persistence, the streaming proxy (main process does the networking
 * to avoid renderer CORS), search tools, and the slides-only ai:* channels
 * (image generation, media analysis, style templates).
 */
import { app, ipcMain, safeStorage } from 'electron'
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { randomUUID } from 'node:crypto'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  isAiProviderId,
  isLocalCliProvider,
  generateProviderImage,
  listProviderModels,
  providerGeneratesImages,
  streamForProvider,
  testProvider,
  type AiProviderProbeRequest,
  type AiSettings,
  type AiStreamChunk,
  type AiStreamRequest,
  type GenSparkAccountStatus,
} from '@genoffice/ai-provider'
import {
  cliStatus,
  isCliProvider,
  listCliModels,
  streamAgentCli,
  testAgentCli,
} from '@genoffice/ai-cli'
import type { AgentToolCall } from '@genoffice/agent-core'
import { createAiSettingsStore, fetchWithSsrfGuard } from '@genoffice/electron-utils'
import {
  webSearch,
  imageSearch,
  gskApiKey,
  gskGenerateImage,
  gskAnalyzeMedia,
  gskLogin,
  gskLoginInfo,
  hasGskAuth,
} from '@genoffice/ai-search'
import { addPicture } from '@genoffice/pptx-engine'
import { EMU_PER_PX_96 } from '@genoffice/pptx-render'
import { tm } from './i18n-main'
import { pushHistory, rebuildSlide, sessions } from './session-state'

// ---- AI settings + streaming proxy (the main process does the networking to avoid renderer CORS; implementation shared via @genoffice/ai-provider) ----

const AI_SETTINGS_PATH = () => join(app.getPath('userData'), 'ai-settings.json')

function readJson<T>(path: string, fallback: T): T {
  try {
    if (existsSync(path)) return JSON.parse(readFileSync(path, 'utf-8')) as T
  } catch {
    /* Corrupted state file: fall back to defaults */
  }
  return fallback
}

function writeJson(path: string, value: unknown): void {
  mkdirSync(join(path, '..'), { recursive: true })
  writeFileSync(path, JSON.stringify(value, null, 2))
}

const activeAiStreams = new Map<string, AbortController>()

/** the same userData/ai-settings.json the other apps use, so one setup covers the suite */
const aiSettingsStore = createAiSettingsStore({ path: AI_SETTINGS_PATH, safeStorage })

/** genspark authenticates from the gsk login state, every other provider from the settings file */
function resolveAiApiKey(provider: AiSettings['provider']): string {
  return provider === 'genspark' ? gskApiKey() : aiSettingsStore.apiKeyFor(provider)
}

export function registerAiIpc(): void {
  ipcMain.handle('ai:get-settings', (): AiSettings => aiSettingsStore.forRenderer())

  // Genspark account (gsk login state): the auth source for AI features; when logged out the frontend uses this to guide login
  ipcMain.handle(
    'ai:gsk-status',
    async (_event, withEmail?: boolean): Promise<GenSparkAccountStatus> => {
      if (!hasGskAuth()) return { loggedIn: false }
      if (!withEmail) return { loggedIn: true }
      const info = await gskLoginInfo()
      return info?.email ? { loggedIn: true, email: info.email } : { loggedIn: true }
    },
  )

  ipcMain.handle('ai:gsk-login', () => {
    gskLogin()
  })

  ipcMain.handle('ai:set-settings', (_event, settings: AiSettings) => {
    aiSettingsStore.write(settings)
  })

  // Settings-dialog probes: verify a key up front, and refresh the model list
  // from the provider rather than relying on the built-in catalogue.
  ipcMain.handle('ai:test-provider', async (_event, request: AiProviderProbeRequest) => {
    if (!isAiProviderId(request.provider)) return { ok: false, error: 'Unknown provider' }
    const config = aiSettingsStore.probeConfigFor(request, resolveAiApiKey(request.provider))
    if (isCliProvider(request.provider)) return testAgentCli(request.provider, config)
    return testProvider(request.provider, config)
  })

  // Whether a locally installed agent CLI can be found, for the settings UI
  ipcMain.handle('ai:cli-status', async (_event, provider: string) =>
    isCliProvider(provider) ? cliStatus(provider) : { installed: false },
  )

  ipcMain.handle('ai:list-models', async (_event, request: AiProviderProbeRequest) => {
    if (!isAiProviderId(request.provider)) return { ok: false, error: 'Unknown provider' }
    if (isCliProvider(request.provider)) return listCliModels(request.provider)
    const config = aiSettingsStore.probeConfigFor(request, resolveAiApiKey(request.provider))
    return listProviderModels(request.provider, config)
  })

  ipcMain.handle('ai:stream', async (event, request: AiStreamRequest) => {
    const { requestId, settings, system, messages } = request
    const tools = request.tools ?? []
    const maxTokens = request.maxTokens ?? 8192
    const provider = isAiProviderId(settings.provider) ? settings.provider : 'genspark'
    // the renderer picks the provider and model; the key and endpoint come from
    // the main process (see AiSettingsStore.configFor)
    const config = aiSettingsStore.configFor(
      provider,
      settings.providers?.[provider]?.model ?? '',
      resolveAiApiKey(provider),
    )
    const send = (chunk: AiStreamChunk) => {
      if (!event.sender.isDestroyed()) event.sender.send('ai:stream-chunk', chunk)
    }
    // A CLI backend has neither here: it authenticates through its own login
    // and takes its model from its own config, so empty is the normal state.
    const selfConfigured = isLocalCliProvider(provider)
    if (!config.apiKey && !selfConfigured) {
      send({
        requestId,
        type: 'error',
        error: provider === 'genspark' ? tm('errAiNotConfigured') : tm('errNoApiKey', { provider }),
      })
      return
    }
    if (!config.model && !selfConfigured) {
      send({ requestId, type: 'error', error: tm('errNoModel') })
      return
    }
    const controller = new AbortController()
    activeAiStreams.set(requestId, controller)
    try {
      const callbacks = {
        signal: controller.signal,
        onDelta: (text: string) => send({ requestId, type: 'delta', text }),
        onToolCall: (toolCall: AgentToolCall) => send({ requestId, type: 'tool-call', toolCall }),
      }
      // a local CLI is a subprocess, not an endpoint, so it bypasses the HTTP path
      if (isCliProvider(provider)) {
        await streamAgentCli(provider, config, system, messages, tools, callbacks)
      } else {
        await streamForProvider(provider, config, system, messages, tools, maxTokens, callbacks)
      }
      send({ requestId, type: 'done' })
    } catch (err) {
      if (controller.signal.aborted) {
        send({ requestId, type: 'done' })
      } else {
        const msg = err instanceof Error ? err.message : String(err)
        console.error(`[ai-stream] ${requestId} (${provider}/${config.model}) failed:`, msg)
        send({ requestId, type: 'error', error: msg })
      }
    } finally {
      activeAiStreams.delete(requestId)
    }
  })

  ipcMain.handle('ai:stream-cancel', (_event, requestId: string) => {
    activeAiStreams.get(requestId)?.abort()
  })

  // Search tools (content + images), Serper with DuckDuckGo fallback
  ipcMain.handle('ai:web-search', async (_event, query: string, maxResults?: number) => {
    try {
      return await webSearch(String(query), typeof maxResults === 'number' ? maxResults : 6)
    } catch (err) {
      return { results: [], method: 'error', error: String(err) }
    }
  })

  ipcMain.handle('ai:image-search', async (_event, query: string, maxResults?: number) => {
    try {
      return await imageSearch(String(query), typeof maxResults === 'number' ? maxResults : 8)
    } catch (err) {
      return { images: [], method: 'error', error: String(err) }
    }
  })
}

// ── ai:* handlers unique to slides ──────────────────────────────────────
// Must be registered inside registerSlidesIpc (not registerAiIpc): in shell aggregate mode the
// generic ai:* channels are registered by docs-main.registerAiIpc, and slides' registerAiIpc is
// never called; docs does not have these channels, so putting them in the wrong place raises
// "No handler registered".
/**
 * Paths this process wrote itself, so a marker cannot be used to read an
 * arbitrary file: the insert handler only accepts one it issued. Same shape as
 * the cloud-page markers in slides-main.
 */
const issuedGeneratedImages = new Set<string>()
const GENERATED_IMAGE_PREFIX = 'genimg:'

/**
 * A generated image arrives as bytes, but every insert path downstream takes a
 * URL. Rather than a `file://` URL — which the SSRF guard rightly refuses, and
 * should keep refusing for anything the model dreamt up — the bytes go to a
 * temp file behind an opaque marker only this process can redeem. The base64
 * never enters the model's context, where it would swamp the conversation.
 */
function writeGeneratedImage(base64: string, mime: string): string {
  const ext = mime.includes('webp') ? 'webp' : mime.includes('jpeg') ? 'jpg' : 'png'
  const file = join(tmpdir(), `genoffice-image-${randomUUID()}.${ext}`)
  writeFileSync(file, Buffer.from(base64, 'base64'))
  issuedGeneratedImages.add(file)
  return GENERATED_IMAGE_PREFIX + file
}

/** bytes for an insert: a marker this process issued, or a guarded download */
async function readImageForInsert(url: string): Promise<{ bytes: Buffer; ext: string } | null> {
  if (url.startsWith(GENERATED_IMAGE_PREFIX)) {
    const path = url.slice(GENERATED_IMAGE_PREFIX.length)
    if (!issuedGeneratedImages.has(path)) return null
    const ext = path.endsWith('.webp') ? 'webp' : path.endsWith('.jpg') ? 'jpg' : 'png'
    return { bytes: readFileSync(path), ext }
  }
  // the URL originates from AI tool calls (prompt-injectable via image search
  // results), so refuse non-http schemes and private/link-local targets;
  // redirects are followed manually so every hop is validated
  const resp = await fetchWithSsrfGuard(url, { headers: { 'User-Agent': 'Mozilla/5.0' } })
  if (!resp || !resp.ok) return null
  const ct = resp.headers.get('content-type') ?? ''
  return {
    bytes: Buffer.from(await resp.arrayBuffer()),
    ext: ct.includes('png') ? 'png' : ct.includes('gif') ? 'gif' : 'jpg',
  }
}

export function registerSlidesOnlyAiIpc(): void {
  // gsk (Genspark CLI) capabilities: AI image generation / media analysis. Returns an error prompt when not logged in.
  ipcMain.handle(
    'ai:generate-image',
    async (
      _event,
      op: {
        prompt: string
        model?: string
        referenceImageUrls?: string[]
        aspectRatio?: string
        imageSize?: string
      },
    ) => {
      // The user's own provider comes first: they configured it, it bills to
      // their account, and the bytes stay on this machine. Genspark hosts the
      // result and hands back a URL instead, and serves the providers that
      // have no image endpoint of their own.
      const provider = aiSettingsStore.forRenderer().provider
      if (providerGeneratesImages(provider)) {
        const result = await generateProviderImage(
          provider,
          aiSettingsStore.configFor(provider, '', resolveAiApiKey(provider)),
          aiSettingsStore.imageModelFor(provider),
          {
            prompt: String(op.prompt ?? ''),
            size: op.imageSize ? String(op.imageSize) : undefined,
          },
        )
        if (!result.ok || !result.base64) return { error: result.error ?? 'generation failed' }
        try {
          return { url: writeGeneratedImage(result.base64, result.mime ?? 'image/png') }
        } catch (err) {
          return { error: err instanceof Error ? err.message : String(err) }
        }
      }
      if (!hasGskAuth()) return { error: tm('errNoImageProvider') }
      try {
        const r = await gskGenerateImage({
          prompt: String(op.prompt),
          model: op.model ? String(op.model) : undefined,
          referenceImageUrls: Array.isArray(op.referenceImageUrls)
            ? op.referenceImageUrls.map(String)
            : undefined,
          aspectRatio: op.aspectRatio ? String(op.aspectRatio) : undefined,
          imageSize: op.imageSize ? String(op.imageSize) : undefined,
        })
        return { url: r.url }
      } catch (err) {
        return { error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    'ai:analyze-media',
    async (_event, op: { mediaUrls: string[]; requirements: string }) => {
      if (!hasGskAuth()) return { error: tm('errGskCli') }
      try {
        const text = await gskAnalyzeMedia({
          mediaUrls: (op.mediaUrls ?? []).map(String),
          requirements: String(op.requirements ?? ''),
        })
        return { text }
      } catch (err) {
        return { error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  // Download an image from a URL and insert it into the given page (image search -> insert in one step; download in the main process avoids CORS)
  ipcMain.handle(
    'ai:insert-image-url',
    async (
      e,
      op: {
        slideIndex: number
        url: string
        xPx: number
        yPx: number
        wPx: number
        hPx: number
        fitWidthPx: number
      },
    ) => {
      const session = sessions.get(e.sender.id)
      if (!session) return null
      const slide = session.opened.deck.slides[op.slideIndex]
      if (!slide) return null
      try {
        const image = await readImageForInsert(String(op.url))
        if (!image) return null
        const buf = image.bytes
        const ext = image.ext
        const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
        const scale = op.fitWidthPx / baseWidthPx
        const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
        pushHistory(session)
        const el = addPicture(session.opened, slide, {
          bytes: new Uint8Array(buf),
          ext,
          offset: {
            x: toEmu(op.xPx),
            y: toEmu(op.yPx),
            cx: Math.max(1, toEmu(op.wPx)),
            cy: Math.max(1, toEmu(op.hPx)),
          },
        })
        if (!el) {
          session.undoStack.pop()
          return null
        }
        session.fitWidthPx = op.fitWidthPx
        const rebuilt = rebuildSlide(session, op.slideIndex)
        return rebuilt ? { slide: rebuilt, sourceId: el.id } : null
      } catch {
        return null
      }
    },
  )

  // ── Style Skill sidecar persistence: write a same-named .styleskill.json next to the draft (fail-open)
  ipcMain.handle(
    'ai:save-sidecar',
    async (
      event,
      data: { topic: string; styleSkill: string; createdAt: string },
    ): Promise<{ ok: boolean }> => {
      try {
        const session = sessions.get(event.sender.id)
        const draftPath = session?.path
        if (!draftPath || !draftPath.endsWith('.pptx')) return { ok: false }
        const sidecarPath = draftPath.replace(/\.pptx$/i, '.styleskill.json')
        writeFileSync(sidecarPath, JSON.stringify(data, null, 2))
        return { ok: true }
      } catch {
        return { ok: false }
      }
    },
  )

  // ── Style template save: stored in userData/style-templates/<name>.json
  const STYLE_TEMPLATES_DIR = () => join(app.getPath('userData'), 'style-templates')

  ipcMain.handle(
    'ai:save-style-template',
    (
      _event,
      name: string,
      data: { topic: string; styleSkill: string; createdAt: string },
    ): { ok: boolean; error?: string } => {
      try {
        const dir = STYLE_TEMPLATES_DIR()
        mkdirSync(dir, { recursive: true })
        // Filename: replace illegal characters in the name with _ then truncate to 64 chars
        const safeName = name.replace(/[/\\:*?"<>|]/g, '_').slice(0, 64)
        if (!safeName) return { ok: false, error: tm('errTplNameInvalid') }
        writeJson(join(dir, `${safeName}.json`), { ...data, name: safeName })
        return { ok: true }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  // ── Style template list
  ipcMain.handle(
    'ai:list-style-templates',
    (): Array<{ name: string; topic: string; createdAt: string }> => {
      try {
        const dir = STYLE_TEMPLATES_DIR()
        if (!existsSync(dir)) return []
        const files = readdirSync(dir).filter((f) => f.endsWith('.json'))
        return files
          .map((f) => {
            try {
              const raw = readJson<{
                name?: string
                topic?: string
                createdAt?: string
                styleSkill?: string
              }>(join(dir, f), {})
              return {
                name: raw.name ?? f.replace(/\.json$/, ''),
                topic: raw.topic ?? '',
                createdAt: raw.createdAt ?? '',
              }
            } catch {
              return null
            }
          })
          .filter(Boolean) as Array<{ name: string; topic: string; createdAt: string }>
      } catch {
        return []
      }
    },
  )

  // ── Style template load
  ipcMain.handle(
    'ai:load-style-template',
    (
      _event,
      name: string,
    ): { ok: boolean; styleSkill?: string; topic?: string; error?: string } => {
      try {
        const dir = STYLE_TEMPLATES_DIR()
        const safeName = name.replace(/[/\\:*?"<>|]/g, '_').slice(0, 64)
        const filePath = join(dir, `${safeName}.json`)
        if (!existsSync(filePath)) return { ok: false, error: tm('errTplMissing', { name }) }
        const raw = readJson<{ styleSkill?: string; topic?: string }>(filePath, {})
        if (!raw.styleSkill) return { ok: false, error: tm('errTplNoSkill', { name }) }
        return { ok: true, styleSkill: raw.styleSkill, topic: raw.topic ?? '' }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )
}
