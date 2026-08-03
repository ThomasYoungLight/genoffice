import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { listProviderModels, testProvider } from '../src/probe'
import { errorResponse, jsonResponse } from './test-utils'
import type { AiProviderConfig } from '../src/types'

const config = (over: Partial<AiProviderConfig> = {}): AiProviderConfig => ({
  apiKey: 'sk-test',
  model: 'test-model',
  ...over,
})

/** the last fetch call, as [url, init] */
const lastCall = () => {
  const mock = globalThis.fetch as unknown as ReturnType<typeof vi.fn>
  const call = mock.mock.calls.at(-1) as [string, RequestInit]
  return { url: call[0], init: call[1] }
}

beforeEach(() => {
  vi.stubGlobal('fetch', vi.fn())
})
afterEach(() => vi.unstubAllGlobals())

const mockOnce = (response: Response) => {
  ;(globalThis.fetch as unknown as ReturnType<typeof vi.fn>).mockResolvedValueOnce(response)
}

describe('testProvider', () => {
  it('refuses before making a request when there is no key or model', async () => {
    expect(await testProvider('anthropic', config({ apiKey: '' }))).toEqual({
      ok: false,
      error: 'No API key configured',
    })
    expect(await testProvider('anthropic', config({ model: '' }))).toEqual({
      ok: false,
      error: 'No model configured',
    })
    expect(globalThis.fetch).not.toHaveBeenCalled()
  })

  it('sends a one-token Anthropic message and reports success', async () => {
    mockOnce(jsonResponse({ content: [{ type: 'text', text: 'hi' }] }))
    expect(await testProvider('anthropic', config({ model: 'claude-sonnet-5' }))).toEqual({
      ok: true,
    })
    const { url, init } = lastCall()
    expect(url).toBe('https://api.anthropic.com/v1/messages')
    expect((init.headers as Record<string, string>)['x-api-key']).toBe('sk-test')
    expect(JSON.parse(init.body as string).max_tokens).toBe(1)
  })

  it('caps the Gemini probe to one output token', async () => {
    mockOnce(jsonResponse({ candidates: [] }))
    await testProvider('gemini', config({ model: 'gemini-2.5-flash' }))
    const { url, init } = lastCall()
    expect(url).toContain('/models/gemini-2.5-flash:generateContent')
    expect(JSON.parse(init.body as string).generationConfig.maxOutputTokens).toBe(1)
  })

  it('explains a rejected key rather than dumping a status code', async () => {
    mockOnce(errorResponse(401, 'invalid x-api-key'))
    const result = await testProvider('anthropic', config())
    expect(result.ok).toBe(false)
    expect(result.error).toContain('the API key was rejected')
    expect(result.error).toContain('401')
  })

  it('explains an unknown model', async () => {
    mockOnce(errorResponse(404, 'model not found'))
    expect((await testProvider('openai', config())).error).toContain(
      'the model or endpoint was not found',
    )
  })

  it('explains being out of quota', async () => {
    mockOnce(errorResponse(429, 'quota exceeded'))
    expect((await testProvider('openai', config())).error).toContain('rate limited or out of quota')
  })

  it('tests the openai provider against /v1/responses, the endpoint it will use', async () => {
    mockOnce(jsonResponse({ output_text: 'hi' }))
    expect(await testProvider('openai', config())).toEqual({ ok: true })
    expect(lastCall().url).toBe('https://api.openai.com/v1/responses')
    const body = JSON.parse(lastCall().init.body as string)
    expect(body.max_output_tokens).toBe(16)
    expect(body.max_tokens).toBeUndefined()
    expect(body.input[0].content[0]).toEqual({ type: 'input_text', text: 'hi' })
  })

  it('retries without the token cap when a Chat Completions model rejects max_tokens', async () => {
    mockOnce(errorResponse(400, 'Unsupported parameter: max_tokens is not supported'))
    mockOnce(jsonResponse({ choices: [{ message: { content: 'hi' } }] }))
    expect(await testProvider('deepseek', config())).toEqual({ ok: true })
    expect(globalThis.fetch).toHaveBeenCalledTimes(2)
    expect(JSON.parse(lastCall().init.body as string).max_tokens).toBeUndefined()
  })

  it('does not retry a 400 that has nothing to do with the token cap', async () => {
    mockOnce(errorResponse(400, 'invalid request: bad content'))
    expect((await testProvider('deepseek', config())).ok).toBe(false)
    expect(globalThis.fetch).toHaveBeenCalledTimes(1)
  })

  it('reports a custom provider with no base URL instead of throwing', async () => {
    const result = await testProvider('custom', config())
    expect(result).toEqual({ ok: false, error: 'A custom provider requires a Base URL' })
  })

  it('surfaces the underlying cause of a network failure', async () => {
    ;(globalThis.fetch as unknown as ReturnType<typeof vi.fn>).mockRejectedValueOnce(
      Object.assign(new Error('fetch failed'), { cause: { code: 'ENOTFOUND' } }),
    )
    expect((await testProvider('openai', config())).error).toContain('ENOTFOUND')
  })
})

describe('listProviderModels', () => {
  it('has nothing to list for the genspark proxy', async () => {
    const result = await listProviderModels('genspark', config())
    expect(result.ok).toBe(false)
    expect(globalThis.fetch).not.toHaveBeenCalled()
  })

  it('reads the Anthropic catalogue', async () => {
    mockOnce(jsonResponse({ data: [{ id: 'claude-opus-4-8' }, { id: 'claude-sonnet-5' }] }))
    const result = await listProviderModels('anthropic', config())
    expect(result.models).toEqual(['claude-opus-4-8', 'claude-sonnet-5'])
    expect(lastCall().url).toContain('/v1/models')
  })

  it('unwraps Gemini names and keeps only models that can chat', async () => {
    mockOnce(
      jsonResponse({
        models: [
          { name: 'models/gemini-2.5-pro', supportedGenerationMethods: ['generateContent'] },
          { name: 'models/text-embedding-004', supportedGenerationMethods: ['embedContent'] },
        ],
      }),
    )
    expect((await listProviderModels('gemini', config())).models).toEqual(['gemini-2.5-pro'])
  })

  it('drops non-chat families from an OpenAI listing', async () => {
    mockOnce(
      jsonResponse({
        data: [
          { id: 'gpt-4o' },
          { id: 'text-embedding-3-small' },
          { id: 'whisper-1' },
          { id: 'dall-e-3' },
          { id: 'tts-1' },
          { id: 'omni-moderation-latest' },
          // seen in a real OpenAI listing: an image model and a completion model
          { id: 'chatgpt-image-latest' },
          { id: 'gpt-3.5-turbo-instruct' },
          { id: 'gpt-4.1-mini' },
        ],
      }),
    )
    expect((await listProviderModels('openai', config())).models).toEqual([
      'gpt-4.1-mini',
      'gpt-4o',
    ])
  })

  it('keeps the current chat families a real listing returns', async () => {
    mockOnce(
      jsonResponse({
        data: [
          { id: 'gpt-5.4' },
          { id: 'gpt-5.2-chat-latest' },
          { id: 'gpt-4o-search-preview' },
          { id: 'gpt-5-codex' },
        ],
      }),
    )
    expect((await listProviderModels('openai', config())).models).toEqual([
      'gpt-4o-search-preview',
      'gpt-5-codex',
      'gpt-5.2-chat-latest',
      'gpt-5.4',
    ])
  })

  it('keeps unrecognised ids: hiding a real model is worse than one extra row', async () => {
    mockOnce(jsonResponse({ data: [{ id: 'some-new-frontier-model' }] }))
    expect((await listProviderModels('openai', config())).models).toEqual([
      'some-new-frontier-model',
    ])
  })

  it('de-duplicates and sorts', async () => {
    mockOnce(jsonResponse({ data: [{ id: 'b' }, { id: 'a' }, { id: 'b' }] }))
    expect((await listProviderModels('deepseek', config())).models).toEqual(['a', 'b'])
  })

  it('hits the custom endpoint the caller supplied', async () => {
    mockOnce(jsonResponse({ data: [{ id: 'local-llama' }] }))
    const result = await listProviderModels(
      'custom',
      config({ baseUrl: 'https://host.example/v1/' }),
    )
    expect(result.models).toEqual(['local-llama'])
    expect(lastCall().url).toBe('https://host.example/v1/models')
  })

  it('reports a rejected key', async () => {
    mockOnce(errorResponse(401, 'bad key'))
    const result = await listProviderModels('openai', config())
    expect(result.ok).toBe(false)
    expect(result.error).toContain('the API key was rejected')
  })

  it('tolerates a listing in an unexpected shape', async () => {
    mockOnce(jsonResponse({ unexpected: true }))
    expect(await listProviderModels('openai', config())).toEqual({ ok: true, models: [] })
  })
})
