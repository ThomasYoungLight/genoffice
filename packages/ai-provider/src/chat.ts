import { httpBodyDetail } from './http-error'
import { chatOpenAiResponses } from './openai-responses'
import { resolveProviderWire } from './providers'
import type { AiChatResponse, AiProviderConfig, AiProviderId } from './types'

async function chatAnthropic(
  config: AiProviderConfig,
  system: string,
  user: string,
  baseUrl = 'https://api.anthropic.com',
): Promise<AiChatResponse> {
  const response = await fetch(`${baseUrl.replace(/\/$/, '')}/v1/messages`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': config.apiKey,
      'anthropic-version': '2023-06-01',
      // Fetch in the Electron main process goes through Chromium's network stack; this header avoids 403.
      'anthropic-dangerous-direct-browser-access': 'true',
    },
    body: JSON.stringify({
      model: config.model,
      max_tokens: 8192,
      system,
      messages: [{ role: 'user', content: user }],
    }),
  })
  if (!response.ok) {
    return {
      ok: false,
      error: `Claude HTTP ${response.status}: ${httpBodyDetail(await response.text())}`,
    }
  }
  const json = (await response.json()) as { content?: Array<{ type: string; text?: string }> }
  const content = json.content
    ?.filter((c) => c.type === 'text')
    .map((c) => c.text ?? '')
    .join('')
  if (!content) return { ok: false, error: 'Claude returned an empty response' }
  return { ok: true, content }
}

async function chatGemini(
  config: AiProviderConfig,
  system: string,
  user: string,
  baseUrl = 'https://generativelanguage.googleapis.com/v1beta',
): Promise<AiChatResponse> {
  const url = `${baseUrl.replace(/\/$/, '')}/models/${config.model}:generateContent`
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-goog-api-key': config.apiKey },
    body: JSON.stringify({
      systemInstruction: { parts: [{ text: system }] },
      contents: [{ role: 'user', parts: [{ text: user }] }],
      generationConfig: { temperature: 0.3 },
    }),
  })
  if (!response.ok) {
    return {
      ok: false,
      error: `Gemini HTTP ${response.status}: ${httpBodyDetail(await response.text())}`,
    }
  }
  const json = (await response.json()) as {
    candidates?: Array<{ content?: { parts?: Array<{ text?: string }> } }>
  }
  const content = json.candidates?.[0]?.content?.parts?.map((p) => p.text ?? '').join('')
  if (!content) return { ok: false, error: 'Gemini returned an empty response' }
  return { ok: true, content }
}

async function chatOpenAiCompatible(
  baseUrl: string,
  config: AiProviderConfig,
  system: string,
  user: string,
): Promise<AiChatResponse> {
  const response = await fetch(`${baseUrl.replace(/\/$/, '')}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${config.apiKey}`,
    },
    body: JSON.stringify({
      model: config.model,
      messages: [
        { role: 'system', content: system },
        { role: 'user', content: user },
      ],
      temperature: 0.3,
    }),
  })
  if (!response.ok) {
    return { ok: false, error: `HTTP ${response.status}: ${httpBodyDetail(await response.text())}` }
  }
  const json = (await response.json()) as { choices?: Array<{ message?: { content?: string } }> }
  const content = json.choices?.[0]?.message?.content
  if (!content) return { ok: false, error: 'AI returned an empty response' }
  return { ok: true, content }
}

/** route a one-shot (non-streaming, non-tool-calling) chat call by provider id */
export async function chatForProvider(
  provider: AiProviderId,
  config: AiProviderConfig,
  system: string,
  user: string,
): Promise<AiChatResponse> {
  let wire: { protocol: string; baseUrl: string }
  try {
    wire = resolveProviderWire(provider, config)
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
  switch (wire.protocol) {
    case 'anthropic':
      return chatAnthropic(config, system, user, wire.baseUrl)
    case 'gemini':
      return chatGemini(config, system, user, wire.baseUrl)
    case 'openai-responses': {
      const result = await chatOpenAiResponses(wire.baseUrl, config, system, user)
      // endpoint does not know /v1/responses; serve this call the old way
      if (!('unsupported' in result)) return result
      return chatOpenAiCompatible(wire.baseUrl, config, system, user)
    }
    default:
      return chatOpenAiCompatible(wire.baseUrl, config, system, user)
  }
}
