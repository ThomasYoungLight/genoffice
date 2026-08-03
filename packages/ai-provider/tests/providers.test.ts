import { describe, expect, it } from 'vitest'
import type { AiModelPreset } from '../src/types'
import {
  AI_PROVIDERS,
  activePreset,
  defaultAiSettings,
  isAiProviderId,
  mergeAiApiKeys,
  providerRequiresApiKey,
  redactAiApiKeys,
  resolveAiSettings,
  sanitizePresets,
} from '../src/providers'

describe('defaultAiSettings', () => {
  it('gives every provider its default model and an empty key by default', () => {
    const settings = defaultAiSettings()
    expect(settings.provider).toBe('genspark')
    for (const meta of AI_PROVIDERS) {
      expect(settings.providers[meta.id].apiKey).toBe('')
      expect(settings.providers[meta.id].model).toBe(meta.defaultModel)
    }
    expect(settings.providers.custom.baseUrl).toBe('')
    expect(settings.providers.anthropic.baseUrl).toBeUndefined()
  })

  it('applies caller-supplied default keys only to the listed providers', () => {
    const settings = defaultAiSettings({ anthropic: 'sk-ant-preset' })
    expect(settings.providers.anthropic.apiKey).toBe('sk-ant-preset')
    expect(settings.providers.gemini.apiKey).toBe('')
  })
})

describe('resolveAiSettings', () => {
  it('returns fresh defaults when nothing is stored', () => {
    const defaults = defaultAiSettings({ anthropic: 'sk-ant-preset' })
    expect(resolveAiSettings({}, defaults)).toEqual(defaults)
  })

  it('migrates the pre-provider single-endpoint shape into the custom provider', () => {
    const defaults = defaultAiSettings()
    const resolved = resolveAiSettings(
      { apiKey: 'legacy-key', model: 'legacy-model', baseUrl: 'https://legacy.example.com/v1' },
      defaults,
    )
    expect(resolved.providers.custom).toEqual({
      apiKey: 'legacy-key',
      model: 'legacy-model',
      baseUrl: 'https://legacy.example.com/v1',
    })
    // untouched providers keep their defaults
    expect(resolved.providers.anthropic).toEqual(defaults.providers.anthropic)
  })

  it('defaults the legacy base URL to the OpenAI endpoint when omitted', () => {
    const resolved = resolveAiSettings({ apiKey: 'legacy-key' }, defaultAiSettings())
    expect(resolved.providers.custom.baseUrl).toBe('https://api.openai.com/v1')
  })

  it('selects the custom provider a legacy file was configured for', () => {
    expect(resolveAiSettings({ apiKey: 'legacy-key' }, defaultAiSettings()).provider).toBe('custom')
    // a legacy file with no key configured nothing, so the default selection stands
    expect(resolveAiSettings({ model: 'gpt-4o' }, defaultAiSettings()).provider).toBe('genspark')
  })

  it('falls back to the default provider when the stored one is unknown', () => {
    const resolved = resolveAiSettings(
      { provider: 'some-future-provider' as never, providers: {} as never },
      defaultAiSettings(),
    )
    expect(resolved.provider).toBe('genspark')
  })

  it('merges stored multi-provider settings over the defaults, provider by provider', () => {
    const defaults = defaultAiSettings({ anthropic: 'preset-key' })
    const resolved = resolveAiSettings(
      {
        provider: 'gemini',
        providers: {
          gemini: { apiKey: 'stored-gemini-key', model: 'gemini-2.5-pro' },
        } as never,
      },
      defaults,
    )
    expect(resolved.provider).toBe('gemini')
    expect(resolved.providers.gemini).toEqual({
      apiKey: 'stored-gemini-key',
      model: 'gemini-2.5-pro',
    })
    // provider not mentioned in stored.providers keeps the computed default
    expect(resolved.providers.anthropic.apiKey).toBe('preset-key')
  })
})

describe('isAiProviderId / providerRequiresApiKey', () => {
  it('recognises exactly the providers this build ships a client for', () => {
    for (const meta of AI_PROVIDERS) expect(isAiProviderId(meta.id)).toBe(true)
    expect(isAiProviderId('mistral')).toBe(false)
    expect(isAiProviderId(undefined)).toBe(false)
  })

  it('exempts only genspark from needing a user-supplied key', () => {
    expect(providerRequiresApiKey('genspark')).toBe(false)
    expect(providerRequiresApiKey('anthropic')).toBe(true)
    expect(providerRequiresApiKey('custom')).toBe(true)
  })
})

describe('redactAiApiKeys', () => {
  it('replaces keys with a hasApiKey flag so renderers never see them', () => {
    const settings = defaultAiSettings({ anthropic: 'sk-ant-secret' })
    settings.provider = 'anthropic'
    const redacted = redactAiApiKeys(settings)
    expect(JSON.stringify(redacted)).not.toContain('sk-ant-secret')
    expect(redacted.providers.anthropic.apiKey).toBe('')
    expect(redacted.providers.anthropic.hasApiKey).toBe(true)
    expect(redacted.providers.gemini.hasApiKey).toBe(false)
    // non-secret fields survive so the settings UI can render them
    expect(redacted.provider).toBe('anthropic')
    expect(redacted.providers.anthropic.model).toBe(settings.providers.anthropic.model)
    expect(redacted.providers.custom.baseUrl).toBe('')
  })
})

describe('mergeAiApiKeys', () => {
  const stored = defaultAiSettings({ anthropic: 'sk-ant-stored', openai: 'sk-openai-stored' })

  it('keeps the stored key when the renderer sends the redacted placeholder', () => {
    const incoming = redactAiApiKeys(stored)
    incoming.provider = 'anthropic'
    const merged = mergeAiApiKeys(incoming, stored)
    expect(merged.provider).toBe('anthropic')
    expect(merged.providers.anthropic.apiKey).toBe('sk-ant-stored')
    expect(merged.providers.openai.apiKey).toBe('sk-openai-stored')
  })

  it('takes a newly typed key over the stored one', () => {
    const incoming = redactAiApiKeys(stored)
    incoming.providers.anthropic = { apiKey: 'sk-ant-new', model: 'claude-sonnet-5' }
    const merged = mergeAiApiKeys(incoming, stored)
    expect(merged.providers.anthropic.apiKey).toBe('sk-ant-new')
    expect(merged.providers.anthropic.model).toBe('claude-sonnet-5')
  })

  it('deletes the stored key when hasApiKey is explicitly false', () => {
    const incoming = redactAiApiKeys(stored)
    incoming.providers.anthropic = { ...incoming.providers.anthropic, hasApiKey: false }
    const merged = mergeAiApiKeys(incoming, stored)
    expect(merged.providers.anthropic.apiKey).toBe('')
    // clearing one provider leaves the others alone
    expect(merged.providers.openai.apiKey).toBe('sk-openai-stored')
  })

  it('never writes the hasApiKey hint back to disk', () => {
    const merged = mergeAiApiKeys(redactAiApiKeys(stored), stored)
    for (const config of Object.values(merged.providers)) {
      expect(config.hasApiKey).toBeUndefined()
    }
  })

  it('falls back to the stored config for a provider the renderer omitted', () => {
    const incoming = redactAiApiKeys(stored)
    delete (incoming.providers as Partial<typeof incoming.providers>).gemini
    const merged = mergeAiApiKeys(incoming, stored)
    expect(merged.providers.gemini).toEqual({
      apiKey: '',
      model: stored.providers.gemini.model,
      baseUrl: undefined,
    })
  })
})

describe('model presets', () => {
  // provider is typed loosely on purpose: half these cases are values a
  // hand-edited or newer settings file could contain
  const preset = (name: string, provider: string, model: string) =>
    ({ name, provider, model }) as AiModelPreset

  it('drops entries this build could not act on, and duplicate names', () => {
    expect(
      sanitizePresets([
        preset('fast', 'openai', 'gpt-4o-mini'),
        preset('fast', 'anthropic', 'claude-sonnet-5'), // duplicate name: first wins
        preset('  ', 'openai', 'gpt-4o'), // no name
        preset('future', 'some-future-provider', 'x'), // provider this build has no client for
        { name: 'no-model', provider: 'openai' }, // missing model
        'nonsense',
      ]),
    ).toEqual([{ name: 'fast', provider: 'openai', model: 'gpt-4o-mini' }])
  })

  it('treats a missing or malformed list as none', () => {
    expect(sanitizePresets(undefined)).toEqual([])
    expect(sanitizePresets({ fast: 'gpt-4o' })).toEqual([])
  })

  it('survives the round trip through the settings file and the renderer', () => {
    const stored = {
      provider: 'openai' as const,
      providers: {} as never,
      presets: [preset('deep', 'openai', 'gpt-5.6-sol')],
    }
    const resolved = resolveAiSettings(stored, defaultAiSettings())
    expect(resolved.presets).toEqual([{ name: 'deep', provider: 'openai', model: 'gpt-5.6-sol' }])
    // out to a renderer and back in with an edit, the way the settings UI saves
    expect(redactAiApiKeys(resolved).presets).toEqual(resolved.presets)
    const merged = mergeAiApiKeys(
      { ...resolved, presets: [...resolved.presets!, preset('fast', 'openai', 'gpt-4o-mini')] },
      resolved,
    )
    expect(merged.presets?.map((p) => p.name)).toEqual(['deep', 'fast'])
    // and a delete arrives as a shorter list
    expect(mergeAiApiKeys({ ...resolved, presets: [] }, merged).presets).toEqual([])
  })

  it('names the preset the current provider and model correspond to', () => {
    const settings = defaultAiSettings()
    settings.provider = 'openai'
    settings.providers.openai.model = 'gpt-4o-mini'
    settings.presets = [
      preset('deep', 'openai', 'gpt-5.6-sol'),
      preset('fast', 'openai', 'gpt-4o-mini'),
    ]
    expect(activePreset(settings)?.name).toBe('fast')

    // a model edited away from every preset is simply not on one
    settings.providers.openai.model = 'gpt-4.1'
    expect(activePreset(settings)).toBeUndefined()

    // same model, different provider: not a match
    settings.provider = 'custom'
    settings.providers.custom.model = 'gpt-4o-mini'
    expect(activePreset(settings)).toBeUndefined()
  })
})
