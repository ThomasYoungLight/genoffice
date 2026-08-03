import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname } from 'node:path'
import {
  defaultAiSettings,
  imageModelFor,
  isAiProviderId,
  mergeAiApiKeys,
  redactAiApiKeys,
  resolveAiSettings,
  type AiProviderConfig,
  type AiProviderId,
  type AiProviderProbeRequest,
  type AiSettings,
  type LegacyAiSettings,
} from '@genoffice/ai-provider'
import { openApiKeys, sealApiKeys, type SafeStorageLike } from './secret-store'

/**
 * The `userData/ai-settings.json` store, shared by every app's main process
 * (docs, sheets and slides each register the same `ai:*` channels, and the
 * shell registers one set for all of its window types).
 *
 * Two rules the callers depend on:
 *
 * 1. **Keys stay in the main process.** `forRenderer()` blanks every API key
 *    and leaves a `hasApiKey` flag behind, and `apiKeyFor()` is what the
 *    stream/chat handlers use to fill the key in at request time. A renderer
 *    hosts document content, which is prompt-injectable, so it has no business
 *    holding the user's key — the genspark key has always worked this way and
 *    a bring-your-own key gets the same treatment.
 * 2. **Keys are encrypted at rest** through the OS keychain when it is
 *    available (see `secret-store`).
 */
export interface AiSettingsStore {
  /** full settings including decrypted keys; main-process use only */
  read(): AiSettings
  /** settings safe to hand a renderer: keys blanked, `hasApiKey` set */
  forRenderer(): AiSettings
  /** persist a renderer update, folding its redacted keys onto what is stored */
  write(incoming: AiSettings): AiSettings
  /** the key to authenticate a request with, '' when none is configured */
  apiKeyFor(provider: AiProviderId): string
  /**
   * The config a stream/chat request runs with: the stored provider settings,
   * except the model, which a renderer may legitimately vary per request
   * (slides forces a high-quality model for deck generation).
   *
   * The base URL deliberately does *not* come from the renderer. Pairing a
   * renderer-chosen endpoint with the stored key would hand a compromised
   * renderer a way to post that key to any host it likes.
   */
  configFor(provider: AiProviderId, model: string, apiKey: string): AiProviderConfig
  /**
   * The config a settings-dialog probe runs with. A key the user has just
   * typed but not saved brings its own endpoint — nothing stored is at risk,
   * and testing an unsaved custom endpoint is the whole point. Without a fresh
   * key the stored endpoint applies, exactly as in `configFor`.
   */
  probeConfigFor(request: AiProviderProbeRequest, storedKey: string): AiProviderConfig
  /**
   * Which image model this provider should generate with: the one named in the
   * settings file, else the build's default for that provider. Never taken
   * from a renderer — same reasoning as the base URL.
   */
  imageModelFor(provider: AiProviderId): string
}

export interface AiSettingsStoreOptions {
  /** resolved lazily: `app.getPath('userData')` is not available at import time */
  path: () => string
  /** Electron's `safeStorage`; null disables encryption (keys stay readable) */
  safeStorage: SafeStorageLike | null
  /** provider keys this build ships preconfigured, if any */
  defaultApiKeys?: Partial<Record<AiProviderId, string>>
}

export function createAiSettingsStore(options: AiSettingsStoreOptions): AiSettingsStore {
  const read = (): AiSettings => {
    let stored: Partial<AiSettings> & LegacyAiSettings = {}
    try {
      const raw: unknown = JSON.parse(readFileSync(options.path(), 'utf8'))
      if (raw && typeof raw === 'object') stored = raw as typeof stored
    } catch {
      // missing or corrupt settings file: start from defaults
    }
    const settings = resolveAiSettings(stored, defaultAiSettings(options.defaultApiKeys))
    return openApiKeys(options.safeStorage, settings)
  }

  return {
    read,

    forRenderer: () => redactAiApiKeys(read()),

    write: (incoming) => {
      const merged = mergeAiApiKeys(
        {
          // a renderer could name a provider this build has no client for
          provider: isAiProviderId(incoming.provider) ? incoming.provider : 'genspark',
          providers: incoming.providers,
          // rebuilt field by field rather than spread, so a renderer cannot
          // smuggle extra keys into the settings file; mergeAiApiKeys
          // re-validates the preset list before it is written
          presets: incoming.presets,
        },
        read(),
      )
      const path = options.path()
      mkdirSync(dirname(path), { recursive: true })
      writeFileSync(path, JSON.stringify(sealApiKeys(options.safeStorage, merged), null, 2))
      return merged
    },

    apiKeyFor: (provider) => read().providers[provider]?.apiKey ?? '',

    imageModelFor: (provider) => imageModelFor(provider, read().providers[provider]?.imageModel),

    configFor: (provider, model, apiKey) => ({
      apiKey,
      model,
      baseUrl: read().providers[provider]?.baseUrl,
    }),

    probeConfigFor: (request, storedKey) =>
      request.apiKey
        ? { apiKey: request.apiKey, model: request.model, baseUrl: request.baseUrl }
        : {
            apiKey: storedKey,
            model: request.model,
            baseUrl: read().providers[request.provider]?.baseUrl,
          },
  }
}
