import type { AgentMessage, AgentToolDef } from '@genoffice/agent-core'
import { httpBodyDetail } from './http-error'
import { postWithQuirkRetry } from './openai-quirks'
import { sseLines, type StreamCallbacks } from './stream'
import type { AiChatResponse, AiProviderConfig } from './types'

/**
 * OpenAI's Responses API (`/v1/responses`).
 *
 * Chat Completions refuses function tools on the reasoning models unless
 * reasoning is switched off — it says so in the 400 and points here. Turning
 * reasoning off on a reasoning model is the wrong trade for document editing,
 * so the openai provider talks to this endpoint instead and lets the model
 * reason with tools in hand.
 *
 * Only the openai provider is routed here. DeepSeek, local servers, the
 * Genspark proxy and anything behind "custom" speak Chat Completions and stay
 * on that client; an endpoint that turns out not to know /v1/responses falls
 * back to it too (see `responsesUnavailable`).
 *
 * The shapes differ enough to need their own client: `input` instead of
 * `messages`, `instructions` instead of a system message, `max_output_tokens`,
 * tools flattened rather than nested under `function`, and its own SSE events.
 */

/** endpoints that answered 404 — one probe each, then Chat Completions forever */
const responsesUnavailable = new Set<string>()

export function responsesSupported(baseUrl: string): boolean {
  return !responsesUnavailable.has(baseUrl)
}

/** test seam: the set is process-wide and would leak between cases */
export function resetResponsesSupport(): void {
  responsesUnavailable.clear()
}

/**
 * Conversation → Responses `input` items.
 *
 * Tool calls and their results are top-level items here rather than fields on
 * a message, paired by `call_id` — which is the id this package already hands
 * the agent loop, so it round-trips untouched.
 */
export function responsesInput(messages: AgentMessage[]): unknown[] {
  const out: unknown[] = []
  for (const m of messages) {
    if (m.role === 'user') {
      out.push({
        role: 'user',
        content: [
          ...(m.text ? [{ type: 'input_text', text: m.text }] : []),
          ...(m.images ?? []).map((img) => ({
            type: 'input_image',
            image_url: `data:${img.mime};base64,${img.base64}`,
          })),
        ],
      })
    } else if (m.role === 'assistant') {
      if (m.text) out.push({ role: 'assistant', content: [{ type: 'output_text', text: m.text }] })
      for (const call of m.toolCalls ?? []) {
        out.push({
          type: 'function_call',
          call_id: call.id,
          name: call.name,
          arguments: JSON.stringify(call.input),
        })
      }
    } else {
      for (const r of m.results) {
        out.push({ type: 'function_call_output', call_id: r.id, output: r.output })
      }
    }
  }
  return out
}

function responsesTools(tools: AgentToolDef[]): unknown[] {
  // flat, unlike Chat Completions' { type, function: { … } }
  return tools.map((t) => ({
    type: 'function',
    name: t.name,
    description: t.description,
    parameters: t.inputSchema,
  }))
}

/** one streamed event, as far as this client cares */
interface ResponsesEvent {
  type?: string
  delta?: string
  item_id?: string
  item?: { type?: string; call_id?: string; name?: string; arguments?: string }
  response?: {
    error?: { message?: string }
    status?: string
    incomplete_details?: { reason?: string }
  }
  message?: string
  error?: { message?: string }
}

export async function streamOpenAiResponses(
  baseUrl: string,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
): Promise<{ unsupported: true } | void> {
  const base = baseUrl.replace(/\/$/, '')
  const response = await postWithQuirkRetry(
    `${base}/responses`,
    `${base}/responses|${config.model}`,
    { Authorization: `Bearer ${config.apiKey}` },
    (quirks) => ({
      model: config.model,
      instructions: system,
      input: responsesInput(messages),
      max_output_tokens: maxTokens,
      ...(tools.length > 0 ? { tools: responsesTools(tools) } : {}),
      // reasoning effort is deliberately not set: the model's own default is
      // the right one, and sending it to a non-reasoning model is a 400
      ...(quirks.noTemperature ? {} : { temperature: 0.3 }),
      stream: true,
    }),
    cb.signal,
  )
  // an endpoint that does not know this route is not an error, it is an older
  // or third-party server — the caller retries on Chat Completions
  if (response.status === 404) {
    responsesUnavailable.add(baseUrl)
    return { unsupported: true }
  }
  if (!response.ok || !response.body) {
    throw new Error(`HTTP ${response.status}: ${httpBodyDetail(await response.text())}`)
  }

  /**
   * Tool calls arrive as an item announced up front, then argument fragments,
   * then a completed item carrying the whole argument string. The completed
   * item is authoritative; the accumulated fragments are the fallback for a
   * stream that ends without one.
   */
  const pending = new Map<string, { callId: string; name: string; json: string }>()
  const emitted = new Set<string>()
  let sawText = false
  /** why the model stopped early, when it did */
  let incomplete: string | undefined
  const emit = (itemId: string, callId: string, name: string, json: string) => {
    if (emitted.has(callId)) return
    emitted.add(callId)
    pending.delete(itemId)
    let input: Record<string, unknown> = {}
    let inputError: string | undefined
    try {
      if (json.trim()) input = JSON.parse(json) as Record<string, unknown>
    } catch (err) {
      inputError = `${err instanceof Error ? err.message : String(err)}; raw: ${json.slice(0, 500)}`
    }
    cb.onToolCall({ id: callId, name, input, ...(inputError ? { inputError } : {}) })
  }

  for await (const line of sseLines(response.body)) {
    if (!line.startsWith('data:')) continue
    const payload = line.slice(5).trim()
    if (!payload || payload === '[DONE]') continue
    let event: ResponsesEvent
    try {
      event = JSON.parse(payload) as ResponsesEvent
    } catch {
      continue
    }
    switch (event.type) {
      case 'response.output_text.delta':
        if (event.delta) {
          sawText = true
          cb.onDelta(event.delta)
        }
        break
      case 'response.incomplete':
        // reasoning tokens count against max_output_tokens here, so a long
        // think can eat the whole budget and end the turn with nothing said
        incomplete = event.response?.incomplete_details?.reason ?? 'incomplete'
        break
      case 'response.output_item.added':
        if (event.item?.type === 'function_call' && event.item.call_id) {
          pending.set(event.item_id ?? event.item.call_id, {
            callId: event.item.call_id,
            name: event.item.name ?? '',
            json: '',
          })
        }
        break
      case 'response.function_call_arguments.delta': {
        const item = event.item_id ? pending.get(event.item_id) : undefined
        if (item) item.json += event.delta ?? ''
        break
      }
      case 'response.output_item.done':
        if (event.item?.type === 'function_call' && event.item.call_id) {
          emit(
            event.item_id ?? event.item.call_id,
            event.item.call_id,
            event.item.name ?? '',
            event.item.arguments ?? pending.get(event.item_id ?? '')?.json ?? '',
          )
        }
        break
      case 'response.failed':
      case 'error':
        throw new Error(
          event.response?.error?.message ?? event.error?.message ?? event.message ?? 'stream error',
        )
      default:
        break
    }
  }
  // anything the stream announced but never completed
  for (const [itemId, item] of pending) emit(itemId, item.callId, item.name, item.json)
  // A turn that said nothing and asked for nothing reaches the caller as an
  // empty answer, which reads like the app broke. Say what the endpoint said.
  if (!sawText && emitted.size === 0) {
    throw new Error(
      incomplete === 'max_output_tokens'
        ? 'the model used its whole output budget on reasoning and produced no answer'
        : `the model returned no content${incomplete ? ` (${incomplete})` : ''}`,
    )
  }
  return undefined
}

/** one-shot, no tools — the compaction summariser and the deck planner use this */
export async function chatOpenAiResponses(
  baseUrl: string,
  config: AiProviderConfig,
  system: string,
  user: string,
): Promise<AiChatResponse | { unsupported: true }> {
  const base = baseUrl.replace(/\/$/, '')
  const response = await postWithQuirkRetry(
    `${base}/responses`,
    `${base}/responses|${config.model}`,
    { Authorization: `Bearer ${config.apiKey}` },
    (quirks) => ({
      model: config.model,
      instructions: system,
      input: [{ role: 'user', content: [{ type: 'input_text', text: user }] }],
      ...(quirks.noTemperature ? {} : { temperature: 0.3 }),
    }),
    undefined,
  )
  if (response.status === 404) {
    responsesUnavailable.add(baseUrl)
    return { unsupported: true }
  }
  if (!response.ok) {
    return { ok: false, error: `HTTP ${response.status}: ${httpBodyDetail(await response.text())}` }
  }
  const json = (await response.json()) as {
    output_text?: string
    output?: Array<{ type?: string; content?: Array<{ type?: string; text?: string }> }>
  }
  // output_text is the convenience field; fall back to walking the items, which
  // also skips the reasoning item a reasoning model puts first
  const content =
    json.output_text ??
    json.output
      ?.filter((item) => item.type === 'message')
      .flatMap((item) => item.content ?? [])
      .filter((part) => part.type === 'output_text')
      .map((part) => part.text ?? '')
      .join('')
  if (!content) return { ok: false, error: 'AI returned an empty response' }
  return { ok: true, content }
}
