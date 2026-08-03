import { describe, expect, it } from 'vitest'
import { defaultAiSettings, type AiSettings } from '@genoffice/ai-provider'
import { isAgentConfigured } from '../src/renderer/agent-configured'

/**
 * The switch between the real agent and the offline regex planner. A false
 * negative here is quiet and confusing: a plain instruction comes back as the
 * micro-DSL hint instead of doing anything.
 */
function settings(patch: (settings: AiSettings) => void): AiSettings {
  const value = defaultAiSettings()
  patch(value)
  return value
}

describe('isAgentConfigured', () => {
  it('accepts a stored key, which a renderer only ever sees as a flag', () => {
    // this is what `ai:get-settings` hands the renderer: key blanked, flag set
    const redacted = settings((s) => {
      s.provider = 'openai'
      s.providers.openai = { apiKey: '', model: 'gpt-4o-mini', hasApiKey: true }
    })
    expect(isAgentConfigured(redacted)).toBe(true)
  })

  it('accepts a key the user just typed but has not saved', () => {
    const typed = settings((s) => {
      s.provider = 'openai'
      s.providers.openai = { apiKey: 'sk-typed', model: 'gpt-4o-mini' }
    })
    expect(isAgentConfigured(typed)).toBe(true)
  })

  it('rejects a provider that needs a key and has none', () => {
    const bare = settings((s) => {
      s.provider = 'openai'
      s.providers.openai = { apiKey: '', model: 'gpt-4o-mini' }
    })
    expect(isAgentConfigured(bare)).toBe(false)
  })

  it('accepts the backends that authenticate somewhere other than the settings file', () => {
    for (const provider of ['genspark', 'claude-cli', 'codex-cli'] as const) {
      expect(isAgentConfigured(settings((s) => (s.provider = provider)))).toBe(true)
    }
  })

  it('still requires a model from providers that need one', () => {
    const noModel = settings((s) => {
      s.provider = 'anthropic'
      s.providers.anthropic = { apiKey: 'sk-ant', model: '' }
    })
    expect(isAgentConfigured(noModel)).toBe(false)
    // ...but a CLI backend's empty model means "whatever the CLI is set to"
    expect(isAgentConfigured(settings((s) => (s.provider = 'claude-cli')))).toBe(true)
  })

  it('is false before the settings file has been read', () => {
    expect(isAgentConfigured(null)).toBe(false)
  })
})
