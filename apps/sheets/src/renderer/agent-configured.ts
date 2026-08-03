import { AI_PROVIDERS, providerRequiresApiKey, type AiSettings } from '@genoffice/ai-provider'

/**
 * Whether an AI backend is set up well enough to run the agent.
 *
 * Sheets keeps a deterministic regex planner ("set A1 to 42") for the offline
 * case, and this is the switch between the two — so getting it wrong does not
 * produce an error, it silently answers a real instruction with the micro-DSL
 * hint. Hence the care, and the tests.
 */
export function isAgentConfigured(settings: AiSettings | null): boolean {
  if (!settings) return false
  const config = settings.providers[settings.provider]
  if (!config) return false
  // a local CLI backend with no model runs whatever its own config selects
  const meta = AI_PROVIDERS.find((provider) => provider.id === settings.provider)
  if (!config.model && !meta?.modelOptional) return false
  // Genspark's key never lands in the settings file (the main process injects
  // it from the gsk login state) and a CLI backend authenticates through its
  // own login, so neither has one to find here.
  if (!providerRequiresApiKey(settings.provider)) return true
  // A renderer is never given the key itself — `ai:get-settings` blanks it and
  // reports `hasApiKey` instead — so a saved key shows up only in that flag.
  return !!config.apiKey || !!config.hasApiKey
}
