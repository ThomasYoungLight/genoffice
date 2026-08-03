import { describe, expect, it } from 'vitest'
import { parseClaudeReply, parseCodexReply } from '../src/models'

/**
 * Line shapes captured from real runs of the two machine protocols — Claude
 * Code's SDK control channel and Codex's app-server JSON-RPC — so the parsers
 * are pinned to what the CLIs actually answer.
 */

describe('parseClaudeReply', () => {
  const response = (payload: unknown) =>
    JSON.stringify({
      type: 'control_response',
      response: { subtype: 'success', request_id: 'genoffice-list-models', response: payload },
    })

  it('reads the --model values out of the catalogue', () => {
    const line = response({
      models: [
        { value: 'opus[1m]', resolvedModel: 'claude-opus-5[1m]', displayName: 'Opus (1M context)' },
        { value: 'sonnet', resolvedModel: 'claude-sonnet-5', displayName: 'Sonnet' },
        { value: 'haiku', resolvedModel: 'claude-haiku-4-5-20251001', displayName: 'Haiku' },
      ],
    })
    expect(parseClaudeReply(line)).toEqual({ models: ['opus[1m]', 'sonnet', 'haiku'] })
  })

  it('drops the "default" entry, which an empty model field already means', () => {
    const line = response({
      models: [{ value: 'default', displayName: 'Default (recommended)' }, { value: 'sonnet' }],
    })
    expect(parseClaudeReply(line)).toEqual({ models: ['sonnet'] })
  })

  it('ignores the hook and init chatter that surrounds the answer', () => {
    for (const event of [
      { type: 'system', subtype: 'hook_started' },
      { type: 'system', subtype: 'init' },
      { type: 'assistant' },
      // a control response to somebody else's request
      { type: 'control_response', response: { subtype: 'success', request_id: 'other' } },
    ]) {
      expect(parseClaudeReply(JSON.stringify(event))).toBeNull()
    }
    expect(parseClaudeReply('Loading plugins…')).toBeNull()
  })

  it('surfaces a rejected request', () => {
    const line = JSON.stringify({
      type: 'control_response',
      response: {
        subtype: 'error',
        request_id: 'genoffice-list-models',
        error: 'Unknown control request',
      },
    })
    expect(parseClaudeReply(line)).toEqual({ error: 'Unknown control request' })
  })
})

describe('parseCodexReply', () => {
  const result = (data: unknown[]) => JSON.stringify({ id: 2, result: { data, nextCursor: null } })

  it('reads the model ids out of the catalogue', () => {
    const line = result([
      { id: 'gpt-5.5', model: 'gpt-5.5', displayName: 'GPT-5.5', hidden: false, isDefault: true },
      { id: 'gpt-5.4', model: 'gpt-5.4', displayName: 'gpt-5.4', hidden: false, isDefault: false },
    ])
    expect(parseCodexReply(line)).toEqual({ models: ['gpt-5.5', 'gpt-5.4'] })
  })

  it('skips models the CLI hides from its own picker', () => {
    const line = result([
      { id: 'gpt-5.5', hidden: false },
      { id: 'gpt-5.1-codex-old', hidden: true },
    ])
    expect(parseCodexReply(line)).toEqual({ models: ['gpt-5.5'] })
  })

  it('ignores the handshake reply and unsolicited notifications', () => {
    for (const event of [
      { id: 1, result: { userAgent: 'genoffice/0.143.0', codexHome: '/Users/x/.codex' } },
      { method: 'remoteControl/status/changed', params: { status: 'disabled' } },
    ]) {
      expect(parseCodexReply(JSON.stringify(event))).toBeNull()
    }
    expect(parseCodexReply('not json')).toBeNull()
  })

  it('surfaces a JSON-RPC error', () => {
    const line = JSON.stringify({ id: 2, error: { code: -32601, message: 'Method not found' } })
    expect(parseCodexReply(line)).toEqual({ error: 'Method not found' })
  })
})
