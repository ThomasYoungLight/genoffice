/**
 * At-rest protection for the model API keys the user brings.
 *
 * Electron's `safeStorage` hands the key material to the OS keychain
 * (Keychain on macOS, DPAPI on Windows), which is the only place it can be
 * kept out of a plain-text file. This module stays free of an Electron import
 * so it remains unit-testable: the caller passes `safeStorage` in as a
 * `SafeStorageLike`.
 *
 * Encryption is not always available (a Linux session without a keyring, a
 * first run before `app.whenReady`). Rather than refuse to store the key we
 * fall back to plain text and mark it as such in the file, so the value can be
 * upgraded on a later write and so an inspection of the file is honest about
 * what it holds.
 */

/** the surface of Electron's `safeStorage` this module uses */
export interface SafeStorageLike {
  isEncryptionAvailable(): boolean
  encryptString(plainText: string): Buffer
  decryptString(encrypted: Buffer): string
}

/** prefix on an OS-encrypted value; the rest of the string is base64 ciphertext */
const ENCRYPTED_PREFIX = 'enc:v1:'
/** prefix on a value we had to leave readable, so its state is not ambiguous */
const PLAINTEXT_PREFIX = 'plain:v1:'

/**
 * Encode a secret for the settings file. Empty input stays empty (an absent
 * key is not a secret), so callers can pass through cleared fields untouched.
 */
export function sealSecret(safeStorage: SafeStorageLike | null, value: string): string {
  if (!value) return ''
  try {
    if (safeStorage?.isEncryptionAvailable()) {
      return ENCRYPTED_PREFIX + safeStorage.encryptString(value).toString('base64')
    }
  } catch {
    // keychain refused (locked session, missing keyring): fall through to plain
  }
  return PLAINTEXT_PREFIX + value
}

/**
 * Decode a value written by `sealSecret`. Anything without a known prefix is
 * treated as a bare key from a settings file that predates this encoding.
 * A value that cannot be decrypted (different machine, rotated keychain entry)
 * comes back empty, which surfaces as "no key configured" rather than a crash.
 */
export function openSecret(safeStorage: SafeStorageLike | null, stored: string): string {
  if (!stored) return ''
  if (stored.startsWith(PLAINTEXT_PREFIX)) return stored.slice(PLAINTEXT_PREFIX.length)
  if (!stored.startsWith(ENCRYPTED_PREFIX)) return stored
  const ciphertext = stored.slice(ENCRYPTED_PREFIX.length)
  try {
    if (!safeStorage?.isEncryptionAvailable()) return ''
    return safeStorage.decryptString(Buffer.from(ciphertext, 'base64'))
  } catch {
    return ''
  }
}

/** whether a stored value is protected by the OS keychain rather than readable */
export function isSealed(stored: string): boolean {
  return stored.startsWith(ENCRYPTED_PREFIX)
}

/** shape the AI settings file shares with anything else holding per-key secrets */
interface SecretBearingSettings {
  providers: Record<string, { apiKey: string }>
}

/** seal every `providers[*].apiKey` for writing to disk */
export function sealApiKeys<T extends SecretBearingSettings>(
  safeStorage: SafeStorageLike | null,
  settings: T,
): T {
  return mapApiKeys(settings, (value) => sealSecret(safeStorage, value))
}

/** reverse of `sealApiKeys`, applied to what was read off disk */
export function openApiKeys<T extends SecretBearingSettings>(
  safeStorage: SafeStorageLike | null,
  settings: T,
): T {
  return mapApiKeys(settings, (value) => openSecret(safeStorage, value))
}

function mapApiKeys<T extends SecretBearingSettings>(settings: T, fn: (key: string) => string): T {
  const providers: Record<string, { apiKey: string }> = {}
  for (const [id, config] of Object.entries(settings.providers ?? {})) {
    providers[id] = { ...config, apiKey: fn(config.apiKey ?? '') }
  }
  return { ...settings, providers: providers as T['providers'] }
}
