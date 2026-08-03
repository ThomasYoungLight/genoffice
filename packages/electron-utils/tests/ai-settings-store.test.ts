import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { createAiSettingsStore } from '../src/ai-settings-store'
import type { SafeStorageLike } from '../src/secret-store'

/** stand-in for the OS keychain: a reversible transform, not real crypto */
const keychain: SafeStorageLike = {
  isEncryptionAvailable: () => true,
  encryptString: (plain) => Buffer.from(`kc(${plain})`, 'utf8'),
  decryptString: (buf) => /^kc\((.*)\)$/s.exec(buf.toString('utf8'))![1]!,
}

let dir: string
/** under a directory that does not exist yet, so writes have to create it */
let path: string
/** already-existing directory, for tests that seed a settings file first */
let seededPath: string

const newStore = (safeStorage: SafeStorageLike | null = keychain) =>
  createAiSettingsStore({ path: () => path, safeStorage })

/** a store over a settings file written by hand, as an older build would leave it */
const storeOverSeeded = (contents: string) => {
  writeFileSync(seededPath, contents)
  return createAiSettingsStore({ path: () => seededPath, safeStorage: keychain })
}

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'genoffice-ai-settings-'))
  path = join(dir, 'nested', 'ai-settings.json')
  seededPath = join(dir, 'ai-settings.json')
})

afterEach(() => rmSync(dir, { recursive: true, force: true }))

describe('createAiSettingsStore', () => {
  it('starts from defaults when no file exists yet', () => {
    const settings = newStore().read()
    expect(settings.provider).toBe('genspark')
    expect(settings.providers.anthropic.apiKey).toBe('')
  })

  it('creates the settings directory on first write', () => {
    const store = newStore()
    const next = store.read()
    next.provider = 'anthropic'
    next.providers.anthropic = { apiKey: 'sk-ant-x', model: 'claude-sonnet-5' }
    store.write(next)
    expect(store.read().providers.anthropic.apiKey).toBe('sk-ant-x')
  })

  it('round-trips a provider choice and key through the file', () => {
    const store = newStore()
    store.write({
      provider: 'anthropic',
      providers: { ...store.read().providers, anthropic: { apiKey: 'sk-ant-x', model: 'm' } },
    })
    // a second store, as a different window's main process would see it
    const reopened = newStore()
    expect(reopened.read().provider).toBe('anthropic')
    expect(reopened.apiKeyFor('anthropic')).toBe('sk-ant-x')
  })

  it('encrypts the key on disk', () => {
    const store = newStore()
    store.write({
      provider: 'openai',
      providers: { ...store.read().providers, openai: { apiKey: 'sk-openai-x', model: 'gpt-4.1' } },
    })
    expect(readFileSync(path, 'utf8')).not.toContain('sk-openai-x')
  })

  it('blanks keys for the renderer but reports that one exists', () => {
    const store = newStore()
    store.write({
      provider: 'anthropic',
      providers: { ...store.read().providers, anthropic: { apiKey: 'sk-ant-x', model: 'm' } },
    })
    const forRenderer = store.forRenderer()
    expect(JSON.stringify(forRenderer)).not.toContain('sk-ant-x')
    expect(forRenderer.providers.anthropic.hasApiKey).toBe(true)
    expect(forRenderer.providers.gemini.hasApiKey).toBe(false)
  })

  it('keeps the stored key when the renderer writes back what it was given', () => {
    const store = newStore()
    store.write({
      provider: 'anthropic',
      providers: { ...store.read().providers, anthropic: { apiKey: 'sk-ant-x', model: 'm' } },
    })
    // the round trip a settings dialog performs: read redacted, change the
    // model, save. The key must survive.
    const edited = store.forRenderer()
    edited.providers.anthropic = { ...edited.providers.anthropic, model: 'claude-opus-4-8' }
    store.write(edited)
    expect(store.apiKeyFor('anthropic')).toBe('sk-ant-x')
    expect(store.read().providers.anthropic.model).toBe('claude-opus-4-8')
  })

  it('deletes the key when the dialog clears it', () => {
    const store = newStore()
    store.write({
      provider: 'anthropic',
      providers: { ...store.read().providers, anthropic: { apiKey: 'sk-ant-x', model: 'm' } },
    })
    const cleared = store.forRenderer()
    cleared.providers.anthropic = { ...cleared.providers.anthropic, hasApiKey: false }
    store.write(cleared)
    expect(store.apiKeyFor('anthropic')).toBe('')
  })

  it('rejects a provider id this build has no client for', () => {
    const store = newStore()
    store.write({ provider: 'wat' as never, providers: store.read().providers })
    expect(store.read().provider).toBe('genspark')
  })

  it('falls back to defaults on a corrupt settings file', () => {
    expect(storeOverSeeded('not json').read().provider).toBe('genspark')
  })

  it('reads a plaintext key written before keys were encrypted', () => {
    const legacy = storeOverSeeded(
      JSON.stringify({
        provider: 'openai',
        providers: { openai: { apiKey: 'sk-legacy-plain', model: 'gpt-4.1' } },
      }),
    )
    expect(legacy.apiKeyFor('openai')).toBe('sk-legacy-plain')
  })

  describe('endpoint rules', () => {
    const withCustom = () => {
      const store = newStore()
      store.write({
        provider: 'custom',
        providers: {
          ...store.read().providers,
          custom: { apiKey: 'sk-custom', model: 'm', baseUrl: 'https://saved.example/v1' },
        },
      })
      return store
    }

    it('lets a request pick the model but never the endpoint', () => {
      const store = withCustom()
      const config = store.configFor('custom', 'other-model', 'sk-custom')
      expect(config.model).toBe('other-model')
      // a renderer must not be able to aim the stored key at another host
      expect(config.baseUrl).toBe('https://saved.example/v1')
    })

    it('probes the stored endpoint when reusing the stored key', () => {
      const store = withCustom()
      const config = store.probeConfigFor(
        { provider: 'custom', model: 'm', baseUrl: 'https://attacker.example/v1' },
        'sk-custom',
      )
      expect(config.apiKey).toBe('sk-custom')
      expect(config.baseUrl).toBe('https://saved.example/v1')
    })

    it('probes the typed endpoint when the key was typed alongside it', () => {
      const store = withCustom()
      const config = store.probeConfigFor(
        {
          provider: 'custom',
          model: 'm',
          baseUrl: 'https://new.example/v1',
          apiKey: 'sk-just-typed',
        },
        'sk-custom',
      )
      // nothing stored is exposed: both the key and the host came from the user
      expect(config.apiKey).toBe('sk-just-typed')
      expect(config.baseUrl).toBe('https://new.example/v1')
    })
  })

  it('round-trips model presets, which live alongside the keys but are not secret', () => {
    const store = newStore()
    const base = store.read()
    store.write({
      ...base,
      provider: 'openai',
      presets: [
        { name: 'fast', provider: 'openai', model: 'gpt-4o-mini' },
        { name: 'deep', provider: 'openai', model: 'gpt-5.6-sol' },
      ],
    })
    expect(store.read().presets?.map((preset) => preset.name)).toEqual(['fast', 'deep'])
    // and they reach the renderer, which is where the picker reads them
    expect(store.forRenderer().presets).toHaveLength(2)
  })

  it('reports no key rather than failing when the keychain cannot decrypt', () => {
    const store = newStore()
    store.write({
      provider: 'anthropic',
      providers: { ...store.read().providers, anthropic: { apiKey: 'sk-ant-x', model: 'm' } },
    })
    // same file opened where the OS offers no keyring at all
    expect(newStore(null).apiKeyFor('anthropic')).toBe('')
  })
})
