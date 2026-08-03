/**
 * Request fields an OpenAI endpoint refuses, learned by asking it.
 *
 * The newer models renamed `max_tokens` to `max_completion_tokens`, accept
 * only the default temperature, and (on /v1/chat/completions) refuse function
 * tools while reasoning is on. None of that can be decided from the model id —
 * the naming is not a reliable signal, and these code paths also serve
 * DeepSeek, local servers and any OpenAI-compatible endpoint the user points
 * "custom" at, most of which still want the old fields.
 *
 * So the endpoint is asked: a request that trips over one of these comes back
 * as a 400 naming the parameter, and it is retried with that field fixed. The
 * answer is remembered per endpoint+model, so only the first request of a
 * session pays for it. Nothing has streamed at that point, so the retry is
 * invisible.
 */
export interface OpenAiQuirks {
  /** send `max_completion_tokens` instead of `max_tokens` */
  maxCompletionTokens?: boolean
  /** omit `temperature` entirely */
  noTemperature?: boolean
  /**
   * send `reasoning_effort: 'none'`. Chat Completions refuses function tools
   * on a reasoning model otherwise; the alternative it offers is /v1/responses,
   * which the openai provider now uses (see openai-responses.ts) — this is the
   * fallback for endpoints that only speak Chat Completions.
   */
  noReasoning?: boolean
}

const learned = new Map<string, OpenAiQuirks>()

/**
 * The vendor phrases the complaint both ways round — "Unsupported parameter:
 * 'max_tokens'" and "'temperature' does not support 0.3 with this model" — so
 * the parameter name is matched near the rejection on either side.
 */
function rejects(body: string, param: string): boolean {
  const said = '(unsupported|not supported|does not support)'
  return new RegExp(`(${said}[^\\n]{0,40}${param})|(${param}[^\\n]{0,60}${said})`, 'i').test(body)
}

/** which quirk, if any, this 400 body is complaining about */
export function quirkFromError(body: string): keyof OpenAiQuirks | null {
  if (/max_completion_tokens/i.test(body) || rejects(body, 'max_tokens')) {
    return 'maxCompletionTokens'
  }
  if (rejects(body, 'reasoning_effort')) return 'noReasoning'
  if (rejects(body, 'temperature')) return 'noTemperature'
  return null
}

export function quirksFor(key: string): OpenAiQuirks {
  return learned.get(key) ?? {}
}

/**
 * POST, and retry while the endpoint keeps naming a field we know how to drop
 * or rename. At most one retry per quirk, so a genuinely broken request still
 * surfaces instead of looping.
 */
export async function postWithQuirkRetry(
  url: string,
  key: string,
  headers: Record<string, string>,
  body: (quirks: OpenAiQuirks) => unknown,
  signal: AbortSignal | undefined,
): Promise<Response> {
  const post = () =>
    fetch(url, {
      method: 'POST',
      ...(signal ? { signal } : {}),
      headers: { 'Content-Type': 'application/json', ...headers },
      body: JSON.stringify(body(quirksFor(key))),
    })

  let response = await post()
  for (let attempt = 0; attempt < 3 && response.status === 400; attempt++) {
    const text = await response.clone().text()
    const quirk = quirkFromError(text)
    if (!quirk || quirksFor(key)[quirk]) break
    learned.set(key, { ...quirksFor(key), [quirk]: true })
    response = await post()
  }
  return response
}

/** test seam: the learned map is process-wide and would leak between cases */
export function resetOpenAiQuirks(): void {
  learned.clear()
}
