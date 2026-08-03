import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AgentToolCall } from '@genoffice/agent-core'
import { sseLines, streamForProvider } from '../src/stream'
import { quirkFromError, resetOpenAiQuirks } from '../src/openai-quirks'
import { errorResponse, okResponse, sseStream } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
  // the learned-quirk map is process-wide; without this a case would inherit
  // what an earlier one taught the same endpoint+model
  resetOpenAiQuirks()
})

function collector() {
  const deltas: string[] = []
  const toolCalls: AgentToolCall[] = []
  return {
    deltas,
    toolCalls,
    cb: {
      signal: new AbortController().signal,
      onDelta: (text: string) => deltas.push(text),
      onToolCall: (call: AgentToolCall) => toolCalls.push(call),
    },
  }
}

describe('sseLines', () => {
  it('splits a stream into lines, including a trailing line with no newline', async () => {
    const encoder = new TextEncoder()
    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(encoder.encode('data: a\ndata: b\n'))
        controller.enqueue(encoder.encode('data: c')) // no trailing newline
        controller.close()
      },
    })
    const lines: string[] = []
    for await (const line of sseLines(body)) lines.push(line)
    expect(lines).toEqual(['data: a', 'data: b', 'data: c'])
  })
})

describe('streamForProvider: anthropic', () => {
  it('emits text deltas and a completed tool call', async () => {
    const body = sseStream([
      'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"hello "}}',
      'data: {"type":"content_block_start","index":1,"content_block":{"type":"tool_use","id":"t1","name":"do_thing"}}',
      'data: {"type":"content_block_delta","index":1,"delta":{"type":"input_json_delta","partial_json":"{\\"a\\":"}}',
      'data: {"type":"content_block_delta","index":1,"delta":{"type":"input_json_delta","partial_json":"1}"}}',
      'data: {"type":"content_block_stop","index":1}',
      'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"world"}}',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { deltas, toolCalls, cb } = collector()
    await streamForProvider(
      'anthropic',
      { apiKey: 'k', model: 'claude-sonnet-5' },
      'system',
      [{ role: 'user', text: 'hi' }],
      [],
      100,
      cb,
    )
    expect(deltas.join('')).toBe('hello world')
    expect(toolCalls).toEqual([{ id: 't1', name: 'do_thing', input: { a: 1 } }])
  })

  it('repairs unescaped quotes inside tool input string values', async () => {
    const partial = JSON.stringify({
      type: 'content_block_delta',
      index: 1,
      delta: { type: 'input_json_delta', partial_json: '{"topic": "from "future" to "present""}' },
    })
    const body = sseStream([
      'data: {"type":"content_block_start","index":1,"content_block":{"type":"tool_use","id":"t1","name":"gen"}}',
      `data: ${partial}`,
      'data: {"type":"content_block_stop","index":1}',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { toolCalls, cb } = collector()
    await streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(toolCalls).toEqual([
      { id: 't1', name: 'gen', input: { topic: 'from "future" to "present"' } },
    ])
  })

  it('unparseable tool input becomes inputError instead of killing the stream', async () => {
    const partial = JSON.stringify({
      type: 'content_block_delta',
      index: 1,
      delta: { type: 'input_json_delta', partial_json: '{"a": 1,' }, // truncated JSON, unrepairable
    })
    const body = sseStream([
      'data: {"type":"content_block_start","index":1,"content_block":{"type":"tool_use","id":"t1","name":"gen"}}',
      `data: ${partial}`,
      'data: {"type":"content_block_stop","index":1}',
      'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"after"}}',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { deltas, toolCalls, cb } = collector()
    await streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(toolCalls).toHaveLength(1)
    expect(toolCalls[0]!.input).toEqual({})
    expect(toolCalls[0]!.inputError).toContain('raw: {"a": 1,')
    expect(deltas.join('')).toBe('after') // the stream was not interrupted
  })

  it('throws on a non-ok HTTP response', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response('bad key', { status: 401 })))
    const { cb } = collector()
    await expect(
      streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow(/Claude HTTP 401/)
  })

  it('replaces an HTML error body (e.g. a gateway block page) with a readable note', async () => {
    const html =
      '<!doctype html>\n<html>\n<head><title>Genspark</title></head><body>app shell</body></html>'
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response(html, { status: 403 })))
    const { cb } = collector()
    await expect(
      streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow(/Claude HTTP 403: .*web page instead of an API response/)
  })
})

describe('streamForProvider: gemini', () => {
  it('emits text and a whole (non-partial) function call', async () => {
    const body = sseStream([
      'data: {"candidates":[{"content":{"parts":[{"text":"hi there"}]}}]}',
      'data: {"candidates":[{"content":{"parts":[{"functionCall":{"name":"set_cell","args":{"a1":"42"}}}]}}]}',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { deltas, toolCalls, cb } = collector()
    await streamForProvider(
      'gemini',
      { apiKey: 'k', model: 'gemini-2.5-flash' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(deltas.join('')).toBe('hi there')
    expect(toolCalls).toHaveLength(1)
    expect(toolCalls[0]).toMatchObject({ name: 'set_cell', input: { a1: '42' } })
  })
})

describe('streamForProvider: openai-compatible', () => {
  it('reassembles fragmented tool call arguments and flushes on finish_reason', async () => {
    const body = sseStream([
      'data: {"choices":[{"delta":{"content":"partial "}}]}',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"c1","function":{"name":"replace"}}]}}]}',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"{\\"x\\":1}"}}]}}]}',
      'data: {"choices":[{"delta":{},"finish_reason":"tool_calls"}]}',
      'data: [DONE]',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { deltas, toolCalls, cb } = collector()
    await streamForProvider(
      'deepseek',
      { apiKey: 'k', model: 'deepseek-chat' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(deltas.join('')).toBe('partial ')
    expect(toolCalls).toEqual([{ id: 'c1', name: 'replace', input: { x: 1 } }])
  })

  it('routes deepseek and openai to their fixed base URLs', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream(['data: [DONE]'])))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'deepseek',
      { apiKey: 'k', model: 'deepseek-chat' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(fetchMock).toHaveBeenCalledWith(
      'https://api.deepseek.com/v1/chat/completions',
      expect.anything(),
    )
  })

  it('uses the configured base URL for the custom provider', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream(['data: [DONE]'])))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'custom',
      { apiKey: 'k', model: 'm', baseUrl: 'https://my-endpoint.example.com/v1' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(fetchMock).toHaveBeenCalledWith(
      'https://my-endpoint.example.com/v1/chat/completions',
      expect.anything(),
    )
  })

  /**
   * Bodies copied from real 400s off gpt-5-class models, which reject fields
   * every earlier model required.
   */
  describe('parameter quirks', () => {
    const REJECTS_MAX_TOKENS = JSON.stringify({
      error: {
        message:
          "Unsupported parameter: 'max_tokens' is not supported with this model. Use 'max_completion_tokens' instead.",
        type: 'invalid_request_error',
        param: 'max_tokens',
        code: 'unsupported_parameter',
      },
    })
    const REJECTS_TEMPERATURE = JSON.stringify({
      error: {
        message:
          "Unsupported value: 'temperature' does not support 0.3 with this model. Only the default (1) is supported.",
        type: 'invalid_request_error',
        param: 'temperature',
        code: 'unsupported_value',
      },
    })
    const bodyOf = (call: unknown): Record<string, unknown> =>
      JSON.parse((call as { body: string }).body) as Record<string, unknown>

    const REJECTS_REASONING = JSON.stringify({
      error: {
        message:
          "Function tools with reasoning_effort are not supported for gpt-5.6-sol in /v1/chat/completions. To use function tools, use /v1/responses or set reasoning_effort to 'none'.",
        type: 'invalid_request_error',
        param: 'reasoning_effort',
      },
    })

    it('classifies the rejections it knows how to work around', () => {
      expect(quirkFromError(REJECTS_MAX_TOKENS)).toBe('maxCompletionTokens')
      expect(quirkFromError(REJECTS_TEMPERATURE)).toBe('noTemperature')
      expect(quirkFromError(REJECTS_REASONING)).toBe('noReasoning')
      expect(quirkFromError('{"error":{"message":"Incorrect API key provided"}}')).toBeNull()
    })

    it('turns reasoning off when the model refuses tools with it on', async () => {
      const fetchMock = vi
        .fn()
        .mockResolvedValueOnce(errorResponse(400, REJECTS_REASONING))
        .mockImplementation(() => Promise.resolve(okResponse(sseStream(['data: [DONE]']))))
      vi.stubGlobal('fetch', fetchMock)
      const { cb } = collector()
      const tool = { name: 'edit', description: 'd', inputSchema: { type: 'object' as const } }
      await streamForProvider(
        'deepseek',
        { apiKey: 'k', model: 'quirk-e' },
        'sys',
        [],
        [tool],
        100,
        cb,
      )

      expect(fetchMock).toHaveBeenCalledTimes(2)
      expect(bodyOf(fetchMock.mock.calls[0]![1]).reasoning_effort).toBeUndefined()
      expect(bodyOf(fetchMock.mock.calls[1]![1]).reasoning_effort).toBe('none')
    })

    it('stacks quirks when an endpoint objects to more than one field', async () => {
      const fetchMock = vi
        .fn()
        .mockResolvedValueOnce(errorResponse(400, REJECTS_MAX_TOKENS))
        .mockResolvedValueOnce(errorResponse(400, REJECTS_REASONING))
        .mockImplementation(() => Promise.resolve(okResponse(sseStream(['data: [DONE]']))))
      vi.stubGlobal('fetch', fetchMock)
      const { cb } = collector()
      await streamForProvider('deepseek', { apiKey: 'k', model: 'quirk-f' }, 'sys', [], [], 100, cb)

      expect(fetchMock).toHaveBeenCalledTimes(3)
      const last = bodyOf(fetchMock.mock.calls[2]![1])
      expect(last.max_completion_tokens).toBe(100)
      expect(last.reasoning_effort).toBe('none')
    })

    it('retries with max_completion_tokens when the model rejects max_tokens', async () => {
      const fetchMock = vi
        .fn()
        .mockResolvedValueOnce(errorResponse(400, REJECTS_MAX_TOKENS))
        .mockImplementation(() => Promise.resolve(okResponse(sseStream(['data: [DONE]']))))
      vi.stubGlobal('fetch', fetchMock)
      const { cb } = collector()
      // a model id of its own, so the learned quirk cannot leak between tests
      await streamForProvider('deepseek', { apiKey: 'k', model: 'quirk-a' }, 'sys', [], [], 100, cb)

      expect(fetchMock).toHaveBeenCalledTimes(2)
      const first = bodyOf(fetchMock.mock.calls[0]![1])
      const second = bodyOf(fetchMock.mock.calls[1]![1])
      expect(first.max_tokens).toBe(100)
      expect(second.max_tokens).toBeUndefined()
      expect(second.max_completion_tokens).toBe(100)
    })

    it('remembers the quirk, so the next turn asks correctly the first time', async () => {
      const fetchMock = vi
        .fn()
        .mockResolvedValueOnce(errorResponse(400, REJECTS_MAX_TOKENS))
        .mockImplementation(() => Promise.resolve(okResponse(sseStream(['data: [DONE]']))))
      vi.stubGlobal('fetch', fetchMock)
      const { cb } = collector()
      await streamForProvider('deepseek', { apiKey: 'k', model: 'quirk-b' }, 'sys', [], [], 100, cb)
      await streamForProvider('deepseek', { apiKey: 'k', model: 'quirk-b' }, 'sys', [], [], 100, cb)

      expect(fetchMock).toHaveBeenCalledTimes(3)
      expect(bodyOf(fetchMock.mock.calls[2]![1]).max_completion_tokens).toBe(100)
    })

    it('drops temperature when the model only accepts its default', async () => {
      const fetchMock = vi
        .fn()
        .mockResolvedValueOnce(errorResponse(400, REJECTS_TEMPERATURE))
        .mockImplementation(() => Promise.resolve(okResponse(sseStream(['data: [DONE]']))))
      vi.stubGlobal('fetch', fetchMock)
      const { cb } = collector()
      await streamForProvider('deepseek', { apiKey: 'k', model: 'quirk-c' }, 'sys', [], [], 100, cb)

      expect(fetchMock).toHaveBeenCalledTimes(2)
      expect(bodyOf(fetchMock.mock.calls[0]![1]).temperature).toBe(0.3)
      expect(bodyOf(fetchMock.mock.calls[1]![1]).temperature).toBeUndefined()
    })

    it('surfaces a 400 it cannot work around instead of retrying blindly', async () => {
      const fetchMock = vi
        .fn()
        .mockResolvedValue(errorResponse(400, '{"error":{"message":"Invalid model"}}'))
      vi.stubGlobal('fetch', fetchMock)
      const { cb } = collector()
      await expect(
        streamForProvider('deepseek', { apiKey: 'k', model: 'quirk-d' }, 'sys', [], [], 100, cb),
      ).rejects.toThrow(/HTTP 400/)
      expect(fetchMock).toHaveBeenCalledTimes(1)
    })
  })

  it('rejects the custom provider without a base URL, without ever calling fetch', async () => {
    const fetchMock = vi.fn()
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await expect(
      streamForProvider('custom', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow(/Base URL/)
    expect(fetchMock).not.toHaveBeenCalled()
  })
})

describe('streamForProvider: genspark', () => {
  it('routes claude models to the Anthropic-compatible proxy endpoint', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream([])))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'genspark',
      { apiKey: 'gsk-k', model: 'claude-opus-4-7' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(fetchMock).toHaveBeenCalledWith(
      'https://www.genspark.ai/api/anthropic/v1/messages',
      expect.objectContaining({ headers: expect.objectContaining({ 'x-api-key': 'gsk-k' }) }),
    )
  })

  it('routes gemini models to the Gemini proxy with header auth', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream([])))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'genspark',
      { apiKey: 'gsk-k', model: 'gemini-3-flash-preview' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(fetchMock).toHaveBeenCalledWith(
      'https://www.genspark.ai/api/llm_proxy/gemini/v1beta/models/gemini-3-flash-preview:streamGenerateContent?alt=sse',
      expect.objectContaining({ headers: expect.objectContaining({ 'x-goog-api-key': 'gsk-k' }) }),
    )
  })

  it('routes other models to the OpenAI-compatible proxy', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream(['data: [DONE]'])))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'genspark',
      { apiKey: 'gsk-k', model: 'gpt-5.2' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(fetchMock).toHaveBeenCalledWith(
      'https://www.genspark.ai/api/llm_proxy/v1/chat/completions',
      expect.anything(),
    )
  })
})

it('rejects an unknown provider id', async () => {
  const { cb } = collector()
  await expect(
    streamForProvider('unknown' as never, { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
  ).rejects.toThrow(/Unknown provider/)
})
