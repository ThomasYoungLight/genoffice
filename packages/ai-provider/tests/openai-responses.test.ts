import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'
import { chatForProvider } from '../src/chat'
import { resetOpenAiQuirks } from '../src/openai-quirks'
import { resetResponsesSupport } from '../src/openai-responses'
import { streamForProvider } from '../src/stream'
import { errorResponse, jsonResponse, okResponse, sseStream } from './test-utils'

/**
 * The openai provider talks to /v1/responses; everything else that speaks the
 * OpenAI protocol (DeepSeek, custom endpoints, the Genspark proxy) stays on
 * Chat Completions. Event shapes here follow real /v1/responses streams.
 */

beforeEach(() => {
  resetResponsesSupport()
  resetOpenAiQuirks()
})
afterEach(() => {
  vi.unstubAllGlobals()
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

const TOOL: AgentToolDef = {
  name: 'replace_blocks',
  description: 'replace document blocks',
  inputSchema: { type: 'object', properties: { id: { type: 'string' } } },
}

const config = { apiKey: 'k', model: 'gpt-5.6-sol' }

function stub(...responses: Response[]) {
  const fetchMock = vi.fn()
  for (const response of responses) fetchMock.mockResolvedValueOnce(response)
  fetchMock.mockImplementation(() => Promise.resolve(okResponse(sseStream([]))))
  vi.stubGlobal('fetch', fetchMock)
  return fetchMock
}

const bodyOf = (call: unknown) => JSON.parse((call as { body: string }).body) as Record<string, any>

describe('streamForProvider: openai → Responses API', () => {
  it('posts to /v1/responses in the shape that endpoint expects', async () => {
    const fetchMock = stub()
    const { cb } = collector()
    const messages: AgentMessage[] = [
      { role: 'user', text: 'fix the heading' },
      {
        role: 'assistant',
        text: 'on it',
        toolCalls: [{ id: 'c1', name: 'replace_blocks', input: { id: 'b7' } }],
      },
      { role: 'tool', results: [{ id: 'c1', name: 'replace_blocks', output: 'done' }] },
    ]
    await expect(
      streamForProvider('openai', config, 'you are an editor', messages, [TOOL], 4096, cb),
    ).rejects.toThrow() // empty stub stream; this case is about the request

    expect(fetchMock.mock.calls[0]![0]).toBe('https://api.openai.com/v1/responses')
    const body = bodyOf(fetchMock.mock.calls[0]![1])
    // the system prompt is a field, not a message
    expect(body.instructions).toBe('you are an editor')
    expect(body.max_output_tokens).toBe(4096)
    expect(body.max_tokens).toBeUndefined()
    // tools are flat here, not nested under `function`
    expect(body.tools).toEqual([
      {
        type: 'function',
        name: 'replace_blocks',
        description: 'replace document blocks',
        parameters: TOOL.inputSchema,
      },
    ])
    // reasoning is left to the model's own default — the whole point of moving
    expect(body.reasoning).toBeUndefined()
    expect(body.reasoning_effort).toBeUndefined()
    // tool call and result are top-level items paired by call_id
    expect(body.input).toEqual([
      { role: 'user', content: [{ type: 'input_text', text: 'fix the heading' }] },
      { role: 'assistant', content: [{ type: 'output_text', text: 'on it' }] },
      { type: 'function_call', call_id: 'c1', name: 'replace_blocks', arguments: '{"id":"b7"}' },
      { type: 'function_call_output', call_id: 'c1', output: 'done' },
    ])
  })

  it('sends images as input_image data URLs', async () => {
    const fetchMock = stub()
    const { cb } = collector()
    await expect(
      streamForProvider(
        'openai',
        config,
        'sys',
        [{ role: 'user', text: 'what is this', images: [{ base64: 'AAA', mime: 'image/png' }] }],
        [],
        100,
        cb,
      ),
    ).rejects.toThrow()
    expect(bodyOf(fetchMock.mock.calls[0]![1]).input[0].content).toEqual([
      { type: 'input_text', text: 'what is this' },
      { type: 'input_image', image_url: 'data:image/png;base64,AAA' },
    ])
  })

  it('streams text deltas and a completed tool call', async () => {
    stub(
      okResponse(
        sseStream([
          'data: {"type":"response.created"}',
          'data: {"type":"response.output_text.delta","delta":"Rewriting "}',
          'data: {"type":"response.output_text.delta","delta":"the heading"}',
          'data: {"type":"response.output_item.added","item_id":"i1","item":{"type":"function_call","call_id":"c9","name":"replace_blocks"}}',
          'data: {"type":"response.function_call_arguments.delta","item_id":"i1","delta":"{\\"id\\":"}',
          'data: {"type":"response.function_call_arguments.delta","item_id":"i1","delta":"\\"b7\\"}"}',
          'data: {"type":"response.output_item.done","item_id":"i1","item":{"type":"function_call","call_id":"c9","name":"replace_blocks","arguments":"{\\"id\\":\\"b7\\"}"}}',
          'data: {"type":"response.completed"}',
        ]),
      ),
    )
    const { deltas, toolCalls, cb } = collector()
    await streamForProvider('openai', config, 'sys', [], [TOOL], 100, cb)
    expect(deltas.join('')).toBe('Rewriting the heading')
    expect(toolCalls).toEqual([{ id: 'c9', name: 'replace_blocks', input: { id: 'b7' } }])
  })

  it('falls back to the accumulated fragments when no completed item arrives', async () => {
    stub(
      okResponse(
        sseStream([
          'data: {"type":"response.output_item.added","item_id":"i1","item":{"type":"function_call","call_id":"c9","name":"replace_blocks"}}',
          'data: {"type":"response.function_call_arguments.delta","item_id":"i1","delta":"{\\"id\\":\\"b7\\"}"}',
          'data: {"type":"response.completed"}',
        ]),
      ),
    )
    const { toolCalls, cb } = collector()
    await streamForProvider('openai', config, 'sys', [], [TOOL], 100, cb)
    expect(toolCalls).toEqual([{ id: 'c9', name: 'replace_blocks', input: { id: 'b7' } }])
  })

  it('emits a tool call once, even though the item is both completed and pending', async () => {
    stub(
      okResponse(
        sseStream([
          'data: {"type":"response.output_item.added","item_id":"i1","item":{"type":"function_call","call_id":"c9","name":"t"}}',
          'data: {"type":"response.output_item.done","item_id":"i1","item":{"type":"function_call","call_id":"c9","name":"t","arguments":"{}"}}',
        ]),
      ),
    )
    const { toolCalls, cb } = collector()
    await streamForProvider('openai', config, 'sys', [], [TOOL], 100, cb)
    expect(toolCalls).toHaveLength(1)
  })

  it('reports unparseable tool input instead of killing the stream', async () => {
    stub(
      okResponse(
        sseStream([
          'data: {"type":"response.output_item.done","item_id":"i1","item":{"type":"function_call","call_id":"c9","name":"t","arguments":"{\\"a\\":"}}',
          'data: {"type":"response.output_text.delta","delta":"after"}',
        ]),
      ),
    )
    const { deltas, toolCalls, cb } = collector()
    await streamForProvider('openai', config, 'sys', [], [TOOL], 100, cb)
    expect(toolCalls[0]!.inputError).toContain('raw: {"a":')
    expect(deltas.join('')).toBe('after')
  })

  it('throws on a failure event', async () => {
    stub(
      okResponse(
        sseStream([
          'data: {"type":"response.failed","response":{"error":{"message":"context length exceeded"}}}',
        ]),
      ),
    )
    const { cb } = collector()
    await expect(streamForProvider('openai', config, 'sys', [], [], 100, cb)).rejects.toThrow(
      /context length exceeded/,
    )
  })

  it('says why a turn produced nothing instead of returning an empty answer', async () => {
    stub(
      okResponse(
        sseStream([
          'data: {"type":"response.incomplete","response":{"incomplete_details":{"reason":"max_output_tokens"}}}',
        ]),
      ),
    )
    const { cb } = collector()
    await expect(streamForProvider('openai', config, 'sys', [], [], 100, cb)).rejects.toThrow(
      /whole output budget on reasoning/,
    )
  })

  it('drops temperature when the model only accepts its default', async () => {
    const fetchMock = stub(
      errorResponse(
        400,
        '{"error":{"message":"Unsupported value: \'temperature\' does not support 0.3 with this model."}}',
      ),
    )
    const { cb } = collector()
    await expect(streamForProvider('openai', config, 'sys', [], [], 100, cb)).rejects.toThrow()
    expect(bodyOf(fetchMock.mock.calls[0]![1]).temperature).toBe(0.3)
    expect(bodyOf(fetchMock.mock.calls[1]![1]).temperature).toBeUndefined()
  })
})

describe('endpoints that do not know /v1/responses', () => {
  it('falls back to Chat Completions for the turn, and stops asking afterwards', async () => {
    const fetchMock = stub(errorResponse(404, 'Unknown request URL'))
    const { cb } = collector()
    await streamForProvider('openai', config, 'sys', [], [], 100, cb)
    expect(fetchMock.mock.calls[0]![0]).toBe('https://api.openai.com/v1/responses')
    expect(fetchMock.mock.calls[1]![0]).toBe('https://api.openai.com/v1/chat/completions')

    // the next turn goes straight to Chat Completions
    await streamForProvider('openai', config, 'sys', [], [], 100, cb)
    expect(fetchMock.mock.calls[2]![0]).toBe('https://api.openai.com/v1/chat/completions')
  })
})

describe('chatForProvider: openai → Responses API', () => {
  it('reads the reply out of output_text', async () => {
    stub(jsonResponse({ output_text: 'a summary' }))
    expect(await chatForProvider('openai', config, 'sys', 'summarise')).toEqual({
      ok: true,
      content: 'a summary',
    })
  })

  it('walks the output items when output_text is absent, skipping the reasoning item', async () => {
    stub(
      jsonResponse({
        output: [
          { type: 'reasoning', summary: [] },
          { type: 'message', content: [{ type: 'output_text', text: 'a summary' }] },
        ],
      }),
    )
    expect(await chatForProvider('openai', config, 'sys', 'summarise')).toEqual({
      ok: true,
      content: 'a summary',
    })
  })

  it('leaves other OpenAI-protocol providers on Chat Completions', async () => {
    const fetchMock = stub(jsonResponse({ choices: [{ message: { content: 'hi' } }] }))
    await chatForProvider('deepseek', { apiKey: 'k', model: 'deepseek-chat' }, 'sys', 'hi')
    expect(fetchMock.mock.calls[0]![0]).toBe('https://api.deepseek.com/v1/chat/completions')
  })
})
