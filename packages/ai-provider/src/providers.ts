import { responsesSupported } from './openai-responses'
import type {
  AiModelPreset,
  AiProviderConfig,
  AiProviderId,
  AiProviderMeta,
  AiSettings,
  LegacyAiSettings,
  ProviderProtocol,
} from './types'

/**
 * Genspark server-side LLM proxy endpoints. All three protocols share the
 * api_key from the gsk login; model ids follow the proxy's own naming scheme,
 * which differs from the official vendor ids.
 */
export const GENSPARK_LLM_BASE_URLS = {
  anthropic: 'https://www.genspark.ai/api/anthropic',
  gemini: 'https://www.genspark.ai/api/llm_proxy/gemini/v1beta',
  openai: 'https://www.genspark.ai/api/llm_proxy/v1',
} as const

/** each vendor's own endpoint, used when the user brings their own key */
export const DIRECT_BASE_URLS = {
  anthropic: 'https://api.anthropic.com',
  gemini: 'https://generativelanguage.googleapis.com/v1beta',
  deepseek: 'https://api.deepseek.com/v1',
  openai: 'https://api.openai.com/v1',
} as const

/**
 * Which protocol to speak and where to send it.
 *
 * Genspark is the interesting case: one account fronts three upstreams, so the
 * protocol follows the model id rather than the provider. Every caller
 * (streaming, one-shot chat, the connectivity probe) routes through here so
 * that rule lives in exactly one place.
 */
export function resolveProviderWire(
  provider: AiProviderId,
  config: Pick<AiProviderConfig, 'model' | 'baseUrl'>,
): { protocol: ProviderProtocol; baseUrl: string } {
  switch (provider) {
    case 'genspark':
      if (config.model.startsWith('claude')) {
        return { protocol: 'anthropic', baseUrl: GENSPARK_LLM_BASE_URLS.anthropic }
      }
      if (config.model.startsWith('gemini')) {
        return { protocol: 'gemini', baseUrl: GENSPARK_LLM_BASE_URLS.gemini }
      }
      return { protocol: 'openai', baseUrl: GENSPARK_LLM_BASE_URLS.openai }
    case 'anthropic':
      return { protocol: 'anthropic', baseUrl: DIRECT_BASE_URLS.anthropic }
    case 'gemini':
      return { protocol: 'gemini', baseUrl: DIRECT_BASE_URLS.gemini }
    case 'deepseek':
      return { protocol: 'openai', baseUrl: DIRECT_BASE_URLS.deepseek }
    case 'openai':
      // OpenAI's own endpoint speaks the Responses API, which is the only way
      // to use function tools with a reasoning model (Chat Completions makes
      // you turn reasoning off). Everything else here — DeepSeek, custom
      // endpoints, the Genspark proxy — speaks Chat Completions only.
      return {
        protocol: responsesSupported(DIRECT_BASE_URLS.openai) ? 'openai-responses' : 'openai',
        baseUrl: DIRECT_BASE_URLS.openai,
      }
    case 'custom':
      if (!config.baseUrl) throw new Error('A custom provider requires a Base URL')
      return { protocol: 'openai', baseUrl: config.baseUrl }
    case 'claude-cli':
    case 'codex-cli':
      // a subprocess, not an endpoint: the main process routes these to
      // @genoffice/ai-cli before reaching any HTTP path
      throw new Error(`${provider} runs as a local CLI and has no HTTP endpoint`)
    default:
      throw new Error(`Unknown provider: ${String(provider)}`)
  }
}

export const AI_PROVIDERS: AiProviderMeta[] = [
  {
    id: 'genspark',
    label: 'Genspark',
    models: [
      'claude-opus-4-7',
      'claude-opus-4-8',
      'claude-sonnet-4-6',
      'claude-haiku-4-5',
      'gpt-5.2',
      'gemini-3.1-pro-preview',
      'gemini-3-flash-preview',
    ],
    defaultModel: 'claude-opus-4-7',
    keyPlaceholder: 'Not required - sign in to Genspark',
  },
  {
    id: 'anthropic',
    label: 'Claude',
    models: [
      'claude-sonnet-5',
      'claude-opus-4-8',
      'claude-opus-4-7',
      'claude-sonnet-4-6',
      'claude-opus-4-6',
      'claude-opus-4-5-20251101',
      'claude-haiku-4-5-20251001',
      'claude-sonnet-4-5-20250929',
    ],
    defaultModel: 'claude-opus-4-7',
    keyPlaceholder: 'sk-ant-api03-...',
  },
  {
    id: 'gemini',
    label: 'Gemini',
    models: ['gemini-2.5-pro', 'gemini-2.5-flash', 'gemini-2.0-flash'],
    defaultModel: 'gemini-2.5-flash',
    keyPlaceholder: 'AIza...',
  },
  {
    id: 'deepseek',
    label: 'DeepSeek',
    models: ['deepseek-chat', 'deepseek-reasoner'],
    defaultModel: 'deepseek-chat',
    keyPlaceholder: 'sk-...',
  },
  {
    id: 'openai',
    label: 'OpenAI',
    models: ['gpt-4.1', 'gpt-4.1-mini', 'gpt-4o', 'gpt-4o-mini'],
    defaultModel: 'gpt-4.1-mini',
    keyPlaceholder: 'sk-...',
  },
  {
    id: 'custom',
    label: 'Custom',
    models: [],
    defaultModel: '',
    keyPlaceholder: 'API Key',
    needsBaseUrl: true,
  },
  // Locally installed agent CLIs. Model ids are the aliases each CLI accepts;
  // an empty model hands the choice to the CLI's own configuration, which is
  // the right default for someone who already set it up.
  {
    id: 'claude-cli',
    label: 'Claude Code',
    models: ['opus', 'sonnet', 'haiku'],
    defaultModel: '',
    keyPlaceholder: '',
    isLocalCli: true,
    modelOptional: true,
  },
  {
    id: 'codex-cli',
    label: 'Codex',
    models: ['gpt-5.2-codex', 'gpt-5.1-codex', 'gpt-5.2', 'gpt-5.1'],
    defaultModel: '',
    keyPlaceholder: '',
    isLocalCli: true,
    modelOptional: true,
  },
]

/** the provider ids this build knows how to talk to */
export const AI_PROVIDER_IDS: readonly AiProviderId[] = AI_PROVIDERS.map((meta) => meta.id)

export function isAiProviderId(value: unknown): value is AiProviderId {
  return typeof value === 'string' && (AI_PROVIDER_IDS as readonly string[]).includes(value)
}

/**
 * Whether the user must supply a key. Genspark takes its key from the gsk
 * login state, and the local CLI backends authenticate through their own
 * login — in both cases the settings file holds no key.
 */
export function providerRequiresApiKey(provider: AiProviderId): boolean {
  if (provider === 'genspark') return false
  return !AI_PROVIDERS.find((meta) => meta.id === provider)?.isLocalCli
}

/** backed by a CLI on this machine rather than an HTTP endpoint */
export function isLocalCliProvider(provider: AiProviderId): boolean {
  return !!AI_PROVIDERS.find((meta) => meta.id === provider)?.isLocalCli
}

/**
 * Fresh settings with every provider's default model and an empty key,
 * except providers listed in `defaultApiKeys` (e.g. an app-specific
 * preconfigured Anthropic key). Callers own that policy; this package
 * has no hardcoded keys.
 */
export function defaultAiSettings(
  defaultApiKeys?: Partial<Record<AiProviderId, string>>,
): AiSettings {
  const providers = {} as AiSettings['providers']
  for (const meta of AI_PROVIDERS) {
    providers[meta.id] = {
      apiKey: defaultApiKeys?.[meta.id] ?? '',
      model: meta.defaultModel,
      baseUrl: meta.needsBaseUrl ? '' : undefined,
    }
  }
  return { provider: 'genspark', providers, presets: [] }
}

/**
 * Keep only presets this build can act on. The settings file is user-editable
 * and may come from a newer build, so a preset naming an unknown provider (or
 * missing a field) is dropped rather than shown as a row that cannot work.
 */
export function sanitizePresets(value: unknown): AiModelPreset[] {
  if (!Array.isArray(value)) return []
  const seen = new Set<string>()
  const presets: AiModelPreset[] = []
  for (const entry of value) {
    const preset = entry as Partial<AiModelPreset>
    const name = typeof preset?.name === 'string' ? preset.name.trim() : ''
    if (!name || seen.has(name)) continue
    if (!isAiProviderId(preset.provider) || typeof preset.model !== 'string') continue
    seen.add(name)
    presets.push({ name, provider: preset.provider, model: preset.model })
  }
  return presets
}

/** the preset the current provider+model corresponds to, if any */
export function activePreset(settings: AiSettings): AiModelPreset | undefined {
  const model = settings.providers[settings.provider]?.model ?? ''
  return settings.presets?.find((p) => p.provider === settings.provider && p.model === model)
}

/**
 * Merge on-disk settings over freshly computed defaults, migrating the
 * pre-provider shape (a single OpenAI-compatible endpoint) into the
 * "custom" provider slot. `stored` is whatever the caller read from its
 * settings file (already JSON-parsed); this function does no file I/O.
 */
export function resolveAiSettings(
  stored: Partial<AiSettings> & LegacyAiSettings,
  defaults: AiSettings,
): AiSettings {
  if (!stored.providers) {
    if (stored.apiKey) {
      defaults.providers.custom = {
        apiKey: stored.apiKey,
        model: stored.model ?? '',
        baseUrl: stored.baseUrl ?? 'https://api.openai.com/v1',
      }
      // a legacy file only ever held a custom endpoint, so that is what it selected
      return { ...defaults, provider: 'custom' }
    }
    return defaults
  }
  return {
    // a settings file written by a newer build (or hand-edited) can name a
    // provider this build has no client for — fall back rather than fail later
    provider: isAiProviderId(stored.provider) ? stored.provider : defaults.provider,
    providers: { ...defaults.providers, ...stored.providers },
    presets: sanitizePresets(stored.presets),
  }
}

/**
 * Fold a renderer-supplied settings update onto what is already on disk.
 *
 * Renderers never receive API keys, so an incoming config carries a key only
 * when the user just typed one. An empty `apiKey` means "leave the stored key
 * alone" unless `hasApiKey` is explicitly false, which is the settings UI's
 * "remove key" action.
 */
export function mergeAiApiKeys(incoming: AiSettings, stored: AiSettings): AiSettings {
  const providers = {} as AiSettings['providers']
  for (const id of AI_PROVIDER_IDS) {
    const next = incoming.providers[id] ?? stored.providers[id]!
    const previousKey = stored.providers[id]?.apiKey ?? ''
    const apiKey = next.apiKey ? next.apiKey : next.hasApiKey === false ? '' : previousKey
    providers[id] = {
      apiKey,
      model: next.model,
      baseUrl: next.baseUrl,
      imageModel: next.imageModel,
    }
  }
  return {
    provider: incoming.provider,
    providers,
    // presets hold no secrets and the settings UI owns them outright, so the
    // incoming list replaces what is on disk (that is how a delete arrives)
    presets: sanitizePresets(incoming.presets),
  }
}

/**
 * Strip every API key out of settings on their way to a renderer, leaving a
 * `hasApiKey` flag so the settings UI can still say whether one is stored.
 */
export function redactAiApiKeys(settings: AiSettings): AiSettings {
  const providers = {} as AiSettings['providers']
  for (const id of AI_PROVIDER_IDS) {
    const config = settings.providers[id]!
    providers[id] = {
      apiKey: '',
      model: config.model,
      baseUrl: config.baseUrl,
      imageModel: config.imageModel,
      hasApiKey: !!config.apiKey,
    }
  }
  return { provider: settings.provider, providers, presets: settings.presets ?? [] }
}
