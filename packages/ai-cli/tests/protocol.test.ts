import { describe, expect, it } from 'vitest'
import {
  parseToolCall,
  serializeConversation,
  stripToolCall,
  toolProtocolPrompt,
} from '../src/protocol'
import type { AgentToolDef } from '@genoffice/agent-core'

const TOOLS: AgentToolDef[] = [
  {
    name: 'replace_block',
    description: 'Replace a paragraph',
    inputSchema: { type: 'object', properties: { blockId: { type: 'string' } } },
  },
  { name: 'read_document', description: 'Read the document', inputSchema: { type: 'object' } },
]
const NAMES = TOOLS.map((t) => t.name)

describe('toolProtocolPrompt', () => {
  it('lists every tool with its schema', () => {
    const prompt = toolProtocolPrompt(TOOLS)
    expect(prompt).toContain('replace_block')
    expect(prompt).toContain('read_document')
    expect(prompt).toContain('"blockId"')
  })

  it('adds nothing when the turn offers no tools', () => {
    expect(toolProtocolPrompt([])).toBe('')
  })
})

describe('parseToolCall', () => {
  it('reads the fenced form the prompt asks for', () => {
    const call = parseToolCall(
      '```tool_call\n{"tool": "replace_block", "input": {"blockId": "b7", "text": "hi"}}\n```',
      NAMES,
    )
    expect(call?.name).toBe('replace_block')
    expect(call?.input).toEqual({ blockId: 'b7', text: 'hi' })
  })

  it('tolerates a json-tagged or untagged fence', () => {
    expect(parseToolCall('```json\n{"tool":"read_document","input":{}}\n```', NAMES)?.name).toBe(
      'read_document',
    )
    expect(parseToolCall('```\n{"tool":"read_document","input":{}}\n```', NAMES)?.name).toBe(
      'read_document',
    )
  })

  it('tolerates a reply that drops the fence entirely', () => {
    expect(parseToolCall('{"tool":"read_document","input":{}}', NAMES)?.name).toBe('read_document')
  })

  it('accepts the OpenAI spelling models drift into', () => {
    const call = parseToolCall(
      '```tool_call\n{"name":"read_document","arguments":{"a":1}}\n```',
      NAMES,
    )
    expect(call?.name).toBe('read_document')
    expect(call?.input).toEqual({ a: 1 })
  })

  it('finds the call even when the model adds prose around it', () => {
    const call = parseToolCall(
      'Let me look at the document first.\n\n```tool_call\n{"tool":"read_document","input":{}}\n```\n\nThen I will edit it.',
      NAMES,
    )
    expect(call?.name).toBe('read_document')
  })

  it('treats a plain answer as an answer, not a call', () => {
    expect(parseToolCall('The document has three paragraphs.', NAMES)).toBeNull()
  })

  it('does not mistake a code block shown to the user for a call', () => {
    // the model is explaining JSON, not invoking anything
    expect(parseToolCall('```json\n{"blockId": "b7", "text": "hi"}\n```', NAMES)).toBeNull()
  })

  it('refuses a tool that was never offered', () => {
    expect(parseToolCall('```tool_call\n{"tool":"rm_rf","input":{}}\n```', NAMES)).toBeNull()
  })

  it('survives malformed JSON instead of throwing', () => {
    expect(parseToolCall('```tool_call\n{"tool": "read_document",,,}\n```', NAMES)).toBeNull()
  })

  it('defaults a missing or non-object input to an empty object', () => {
    expect(parseToolCall('```tool_call\n{"tool":"read_document"}\n```', NAMES)?.input).toEqual({})
    expect(
      parseToolCall('```tool_call\n{"tool":"read_document","input":[1,2]}\n```', NAMES)?.input,
    ).toEqual({})
  })
})

describe('stripToolCall', () => {
  it('removes the block so its plumbing never reaches the transcript', () => {
    const text = 'Reading it now.\n\n```tool_call\n{"tool":"read_document","input":{}}\n```'
    expect(stripToolCall(text)).toBe('Reading it now.')
  })

  it('leaves an ordinary answer untouched', () => {
    expect(stripToolCall('Just an answer.')).toBe('Just an answer.')
  })
})

describe('serializeConversation', () => {
  it('flattens roles into one prompt, since a CLI takes a single string', () => {
    const prompt = serializeConversation([
      { role: 'user', text: 'Summarise this' },
      { role: 'assistant', text: '', toolCalls: [{ id: '1', name: 'read_document', input: {} }] },
      { role: 'tool', results: [{ id: '1', name: 'read_document', output: 'Hello world' }] },
    ])
    expect(prompt).toContain('## User\nSummarise this')
    expect(prompt).toContain('read_document')
    expect(prompt).toContain('## Tool results')
    expect(prompt).toContain('Hello world')
  })

  it('marks a failed tool result as an error', () => {
    const prompt = serializeConversation([
      {
        role: 'tool',
        results: [{ id: '1', name: 'edit', output: 'no such block', isError: true }],
      },
    ])
    expect(prompt).toContain('(error)')
  })

  it('says images were dropped rather than pretending the model saw them', () => {
    const prompt = serializeConversation([
      { role: 'user', text: 'What is this?', images: [{ base64: 'x', mime: 'image/png' }] },
    ])
    expect(prompt).toContain('image attachment')
    expect(prompt).toContain('text-only')
  })
})
