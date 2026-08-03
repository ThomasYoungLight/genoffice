import { describe, expect, it } from 'vitest'
import { extractError, extractText } from '../src/stream'
import { isCliProvider } from '../src/detect'

/**
 * Line shapes captured from real runs of `claude --print --output-format
 * stream-json` and `codex exec --json`, so the parsers are pinned to what the
 * CLIs actually emit rather than to what their docs describe.
 */

describe('extractText: claude-cli', () => {
  it('reads a text delta out of the wrapped vendor event', () => {
    const line = JSON.stringify({
      type: 'stream_event',
      event: { type: 'content_block_delta', delta: { type: 'text_delta', text: 'Hello' } },
    })
    expect(extractText('claude-cli', line)).toBe('Hello')
  })

  it('ignores the hook, init and rate-limit chatter that surrounds the answer', () => {
    for (const event of [
      { type: 'system', subtype: 'hook_started' },
      { type: 'system', subtype: 'init' },
      { type: 'rate_limit_event' },
      { type: 'assistant' },
      { type: 'result', subtype: 'success', result: 'Hello' },
    ]) {
      expect(extractText('claude-cli', JSON.stringify(event))).toBeNull()
    }
  })

  it('ignores a non-text delta', () => {
    const line = JSON.stringify({
      type: 'stream_event',
      event: {
        type: 'content_block_delta',
        delta: { type: 'input_json_delta', partial_json: '{' },
      },
    })
    expect(extractText('claude-cli', line)).toBeNull()
  })
})

describe('extractText: codex-cli', () => {
  it('reads a completed agent message', () => {
    const line = JSON.stringify({
      type: 'item.completed',
      item: { id: 'item_3', type: 'agent_message', text: 'PONG' },
    })
    expect(extractText('codex-cli', line)).toBe('PONG')
  })

  it('ignores thread and turn bookkeeping', () => {
    for (const event of [
      { type: 'thread.started', thread_id: 'x' },
      { type: 'turn.started' },
      { type: 'turn.completed', usage: {} },
      { type: 'item.completed', item: { type: 'error', message: 'a warning' } },
    ]) {
      expect(extractText('codex-cli', JSON.stringify(event))).toBeNull()
    }
  })
})

describe('extractText: robustness', () => {
  it('skips the plain-text logging both CLIs interleave with their JSON', () => {
    expect(extractText('codex-cli', 'ERROR rmcp::transport::worker: worker quit')).toBeNull()
    expect(extractText('claude-cli', 'Reading additional input from stdin...')).toBeNull()
  })

  it('skips JSON that is not an object', () => {
    expect(extractText('claude-cli', '"a string"')).toBeNull()
    expect(extractText('codex-cli', 'null')).toBeNull()
  })
})

describe('extractError', () => {
  it('reports a failed claude turn', () => {
    const line = JSON.stringify({
      type: 'result',
      is_error: true,
      result: 'Credit balance too low',
    })
    expect(extractError('claude-cli', line)).toBe('Credit balance too low')
  })

  it('does not treat a successful claude result as an error', () => {
    const line = JSON.stringify({ type: 'result', is_error: false, result: 'Hello' })
    expect(extractError('claude-cli', line)).toBeNull()
  })

  it('collects codex error items', () => {
    const line = JSON.stringify({
      type: 'item.completed',
      item: { type: 'error', message: 'Model metadata not found' },
    })
    expect(extractError('codex-cli', line)).toBe('Model metadata not found')
  })
})

describe('isCliProvider', () => {
  it('recognises only the CLI-backed providers', () => {
    expect(isCliProvider('claude-cli')).toBe(true)
    expect(isCliProvider('codex-cli')).toBe(true)
    expect(isCliProvider('anthropic')).toBe(false)
    expect(isCliProvider('genspark')).toBe(false)
  })
})
