import { describe, expect, it } from 'vitest'
import {
  isSealed,
  openApiKeys,
  openSecret,
  sealApiKeys,
  sealSecret,
  type SafeStorageLike,
} from '../src/secret-store'

/** stand-in for the OS keychain: a reversible transform, not real crypto */
function fakeSafeStorage(available = true): SafeStorageLike {
  return {
    isEncryptionAvailable: () => available,
    encryptString: (plain) => Buffer.from(`kc(${plain})`, 'utf8'),
    decryptString: (buf) => {
      const text = buf.toString('utf8')
      const match = /^kc\((.*)\)$/s.exec(text)
      if (!match) throw new Error('not encrypted by this keychain')
      return match[1]!
    },
  }
}

describe('sealSecret / openSecret', () => {
  it('round-trips a key through the keychain', () => {
    const safeStorage = fakeSafeStorage()
    const sealed = sealSecret(safeStorage, 'sk-ant-secret')
    expect(sealed).not.toContain('sk-ant-secret')
    expect(isSealed(sealed)).toBe(true)
    expect(openSecret(safeStorage, sealed)).toBe('sk-ant-secret')
  })

  it('keeps an empty key empty instead of sealing nothing', () => {
    expect(sealSecret(fakeSafeStorage(), '')).toBe('')
    expect(openSecret(fakeSafeStorage(), '')).toBe('')
  })

  it('falls back to a marked plaintext value when encryption is unavailable', () => {
    const safeStorage = fakeSafeStorage(false)
    const sealed = sealSecret(safeStorage, 'sk-plain')
    expect(isSealed(sealed)).toBe(false)
    expect(sealed).not.toBe('sk-plain')
    expect(openSecret(safeStorage, sealed)).toBe('sk-plain')
  })

  it('falls back to plaintext when the keychain throws', () => {
    const throwing: SafeStorageLike = {
      isEncryptionAvailable: () => true,
      encryptString: () => {
        throw new Error('keyring locked')
      },
      decryptString: () => '',
    }
    expect(openSecret(throwing, sealSecret(throwing, 'sk-plain'))).toBe('sk-plain')
  })

  it('reads a bare key from a settings file that predates the encoding', () => {
    expect(openSecret(fakeSafeStorage(), 'sk-ant-legacy')).toBe('sk-ant-legacy')
  })

  it('reports an undecryptable value as no key rather than throwing', () => {
    const sealed = sealSecret(fakeSafeStorage(), 'sk-ant-secret')
    // a different machine: same prefix, keychain cannot open the payload
    const otherMachine: SafeStorageLike = {
      isEncryptionAvailable: () => true,
      encryptString: (plain) => Buffer.from(plain, 'utf8'),
      decryptString: () => {
        throw new Error('decryption failed')
      },
    }
    expect(openSecret(otherMachine, sealed)).toBe('')
  })

  it('reports a sealed value as no key when the keychain is gone entirely', () => {
    const sealed = sealSecret(fakeSafeStorage(), 'sk-ant-secret')
    expect(openSecret(fakeSafeStorage(false), sealed)).toBe('')
    expect(openSecret(null, sealed)).toBe('')
  })
})

describe('sealApiKeys / openApiKeys', () => {
  const settings = {
    provider: 'anthropic',
    providers: {
      anthropic: { apiKey: 'sk-ant-x', model: 'claude-opus-4-7' },
      openai: { apiKey: '', model: 'gpt-4.1-mini' },
      custom: { apiKey: 'sk-custom', model: 'm', baseUrl: 'https://example.com/v1' },
    },
  }

  it('round-trips every provider key and leaves the rest of the shape alone', () => {
    const safeStorage = fakeSafeStorage()
    const sealed = sealApiKeys(safeStorage, settings)
    expect(JSON.stringify(sealed)).not.toContain('sk-ant-x')
    expect(JSON.stringify(sealed)).not.toContain('sk-custom')
    expect(sealed.provider).toBe('anthropic')
    expect(sealed.providers.custom!.baseUrl).toBe('https://example.com/v1')
    expect(openApiKeys(safeStorage, sealed)).toEqual(settings)
  })

  it('does not mutate the settings it was given', () => {
    sealApiKeys(fakeSafeStorage(), settings)
    expect(settings.providers.anthropic.apiKey).toBe('sk-ant-x')
  })
})
