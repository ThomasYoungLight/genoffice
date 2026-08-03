import { execFile } from 'node:child_process'
import { delimiter, join } from 'node:path'
import { accessSync, constants } from 'node:fs'
import { homedir } from 'node:os'
import type { AiProviderId } from '@genoffice/ai-provider'

/** the provider ids backed by a locally installed agent CLI rather than an HTTP endpoint */
export const CLI_PROVIDERS = ['claude-cli', 'codex-cli'] as const
export type CliProviderId = (typeof CLI_PROVIDERS)[number]

export function isCliProvider(provider: AiProviderId | string): provider is CliProviderId {
  return (CLI_PROVIDERS as readonly string[]).includes(provider)
}

interface CliSpec {
  /** executable name to look for on PATH */
  bin: string
  /** environment variable that overrides the discovered path (dev + unusual installs) */
  envOverride: string
  /** where the vendor's installer puts it, for GUI launches with a thin PATH */
  extraDirs: string[]
}

const SPECS: Record<CliProviderId, CliSpec> = {
  'claude-cli': {
    bin: 'claude',
    envOverride: 'GENOFFICE_CLAUDE_CLI_PATH',
    extraDirs: ['.local/bin', '.claude/local', '.bun/bin', '.npm-global/bin'],
  },
  'codex-cli': {
    bin: 'codex',
    envOverride: 'GENOFFICE_CODEX_CLI_PATH',
    extraDirs: ['.local/bin', '.codex/bin', '.bun/bin', '.npm-global/bin'],
  },
}

export interface CliStatus {
  installed: boolean
  /** absolute path to the executable when found */
  path?: string
  /** whatever `--version` printed, trimmed to one line */
  version?: string
}

/**
 * Locate a CLI.
 *
 * `which` is not enough: a macOS app launched from Finder inherits a minimal
 * PATH that usually omits `~/.local/bin`, so an app that only consulted PATH
 * would report "not installed" for a CLI the user runs daily in their terminal.
 * The vendors' usual install directories are therefore searched too.
 */
export function resolveCliPath(provider: CliProviderId): string | null {
  const spec = SPECS[provider]
  const override = process.env[spec.envOverride]
  if (override) return override

  const dirs = [
    ...(process.env.PATH ?? '').split(delimiter).filter(Boolean),
    ...spec.extraDirs.map((d) => join(homedir(), d)),
    '/usr/local/bin',
    '/opt/homebrew/bin',
  ]
  const names = process.platform === 'win32' ? [`${spec.bin}.cmd`, `${spec.bin}.exe`] : [spec.bin]
  for (const dir of dirs) {
    for (const name of names) {
      const candidate = join(dir, name)
      try {
        accessSync(candidate, constants.X_OK)
        return candidate
      } catch {
        // not here; keep looking
      }
    }
  }
  return null
}

const VERSION_TIMEOUT_MS = 10_000

/**
 * Whether the CLI is present, and what version.
 *
 * Deliberately does *not* report login state. Claude Code keeps its
 * credentials in the OS keychain and Codex in its own config, both private
 * formats that would rot the moment either changes. The dialog's Test button
 * makes a real call, which is the only honest way to answer "will this work".
 */
export async function cliStatus(provider: CliProviderId): Promise<CliStatus> {
  const path = resolveCliPath(provider)
  if (!path) return { installed: false }
  const version = await new Promise<string | undefined>((resolve) => {
    execFile(path, ['--version'], { timeout: VERSION_TIMEOUT_MS }, (err, stdout) => {
      if (err) return resolve(undefined)
      resolve(stdout.trim().split('\n')[0]?.trim() || undefined)
    })
  })
  // it answered --version, so it is a working executable rather than a stub
  return version ? { installed: true, path, version } : { installed: true, path }
}
