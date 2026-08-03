import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'

export type AiProviderId =
  | 'genspark'
  | 'anthropic'
  | 'gemini'
  | 'deepseek'
  | 'openai'
  | 'custom'
  /** locally installed agent CLIs, driven as subprocesses (see @genoffice/ai-cli) */
  | 'claude-cli'
  | 'codex-cli'

/** Genspark account status (gsk login state; the sole auth source for AI features) */
export interface GenSparkAccountStatus {
  loggedIn: boolean
  email?: string
}

export interface AiProviderConfig {
  /**
   * The model API key. Only ever populated inside the Electron main process:
   * `ai:get-settings` blanks it before the settings reach a renderer, and the
   * main process re-reads the real key from disk when it makes the request, so
   * a key never sits in renderer memory (same rule the genspark key follows).
   */
  apiKey: string
  model: string
  /** only used by the custom (OpenAI-compatible) provider */
  baseUrl?: string | undefined
  /**
   * Model for image generation, which is a different model from the chat one.
   * Unset means the provider's default (see DEFAULT_IMAGE_MODELS); naming it
   * here is how a user reaches a newer image model without a new build.
   */
  imageModel?: string | undefined
  /**
   * Whether a key is stored on disk for this provider. Set by the main process
   * for the renderer's benefit (so the settings UI can show "key saved" without
   * receiving it). On the way back in it carries the user's intent: `apiKey: ''`
   * with `hasApiKey: true` means "keep the stored key", with `hasApiKey: false`
   * means "delete it".
   */
  hasApiKey?: boolean | undefined
}

export interface AiProviderMeta {
  id: AiProviderId
  label: string
  models: string[]
  defaultModel: string
  keyPlaceholder: string
  needsBaseUrl?: boolean
  /**
   * Backed by a CLI installed on this machine instead of an HTTP endpoint: no
   * API key, and the settings UI shows install state rather than a key field.
   */
  isLocalCli?: boolean
  /** an empty model is valid and means "whatever the backend defaults to" */
  modelOptional?: boolean
}

/**
 * A named provider+model pairing, so two models from the same provider can be
 * one click apart ("fast" → openai/gpt-4o-mini, "deep" → openai/gpt-5.6-sol).
 * A provider slot holds one model, which is what presets exist to work around.
 *
 * Deliberately no id: the name is the identity, so applying a preset is a plain
 * settings edit and "which preset am I on" is derived by matching the current
 * provider and model rather than stored and kept in sync.
 */
export interface AiModelPreset {
  name: string
  provider: AiProviderId
  model: string
}

export interface AiSettings {
  provider: AiProviderId
  providers: Record<AiProviderId, AiProviderConfig>
  /** user-named provider+model pairings, in the order they were saved */
  presets?: AiModelPreset[] | undefined
}

/** pre-provider settings shape (single OpenAI-compatible endpoint); migrated into "custom" */
export interface LegacyAiSettings {
  baseUrl?: string
  apiKey?: string
  model?: string
}

/**
 * The wire format a request has to speak, independent of who is hosting it.
 * `openai` is Chat Completions; `openai-responses` is OpenAI's newer
 * /v1/responses, which the openai provider uses (see openai-responses.ts).
 */
export type ProviderProtocol = 'anthropic' | 'gemini' | 'openai' | 'openai-responses'

/** whether a locally installed agent CLI (Claude Code, Codex) can be found */
export interface CliStatus {
  installed: boolean
  path?: string
  version?: string
}

/** renderer → main: check a provider, optionally with a key the user just typed */
export interface AiProviderProbeRequest {
  provider: AiProviderId
  model: string
  /** only for the custom provider; ignored unless `apiKey` is also supplied */
  baseUrl?: string | undefined
  /**
   * A key typed into the settings dialog but not saved yet. Omitted when
   * testing what is already stored, in which case the main process uses the
   * saved key *and* the saved base URL — never a renderer-supplied endpoint,
   * which would otherwise be a way to post the stored key to any host.
   */
  apiKey?: string | undefined
}

export interface AiChatRequest {
  settings: AiSettings
  system: string
  user: string
}

export interface AiChatResponse {
  ok: boolean
  content?: string
  error?: string
}

export interface AiStreamRequest {
  requestId: string
  settings: AiSettings
  system: string
  messages: AgentMessage[]
  tools?: AgentToolDef[]
  maxTokens?: number
}

export interface AiStreamChunk {
  requestId: string
  type: 'delta' | 'tool-call' | 'done' | 'error'
  text?: string
  /** complete parsed tool call (emitted once its arguments finish streaming) */
  toolCall?: AgentToolCall
  error?: string
}
