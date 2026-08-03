import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'

/**
 * Making a coding-agent CLI behave like a tool-calling model.
 *
 * Claude Code and Codex are agents, not model endpoints. Two consequences
 * shape everything here:
 *
 * 1. **They accept no caller-supplied tool definitions.** Their own tools are
 *    fixed (file edits, shell) and are switched off entirely — the app's tools
 *    edit a document held in the renderer, which no subprocess can reach. So
 *    the tool contract is carried in the prompt and the call is parsed back out
 *    of the reply. Less robust than a native `tools` parameter, hence the
 *    defensive parsing below, but it is the only mechanism both CLIs share.
 * 2. **Each invocation is a fresh session.** There is no server-side thread to
 *    append to, so the whole conversation is re-serialised into one prompt per
 *    turn. The agent loop already resends full history, so this costs prompt
 *    tokens (which both CLIs cache) rather than correctness.
 */

/** fence the model is asked to emit a call in, and that `parseToolCall` looks for */
const TOOL_FENCE = 'tool_call'

/**
 * Describe the tools and the reply format. Appended to the caller's system
 * prompt so the app's own instructions still lead.
 */
export function toolProtocolPrompt(tools: AgentToolDef[]): string {
  if (tools.length === 0) return ''
  const catalogue = tools
    .map(
      (t) =>
        `### ${t.name}\n${t.description}\nInput JSON Schema:\n${JSON.stringify(t.inputSchema)}`,
    )
    .join('\n\n')
  return `

# Tool calling

You are driving an application through tools. You cannot invoke them yourself:
to call one, reply with ONLY this fenced block and nothing else — no prose
before it, no explanation after it.

\`\`\`${TOOL_FENCE}
{"tool": "<tool name>", "input": { ...arguments matching the tool's schema... }}
\`\`\`

Rules:
- One call per reply. You will be given the result and can then call again.
- The block must contain a single valid JSON object. Do not add comments.
- When you are finished and want to answer the user, reply with plain text and
  no fenced block.
- Never invent a tool that is not listed below.

## Available tools

${catalogue}`
}

/**
 * Flatten the conversation into one prompt. Roles are marked with headings
 * because a CLI takes a single string, not a role array.
 */
export function serializeConversation(messages: AgentMessage[]): string {
  const parts: string[] = []
  for (const message of messages) {
    if (message.role === 'user') {
      // images cannot cross a CLI's stdin; note them so the model is not left
      // silently reasoning about attachments it was never shown
      const images = message.images?.length
        ? `\n[${message.images.length} image attachment(s) omitted: this backend is text-only]`
        : ''
      parts.push(`## User\n${message.text}${images}`)
    } else if (message.role === 'assistant') {
      const calls = (message.toolCalls ?? [])
        .map(
          (c) => `\`\`\`${TOOL_FENCE}\n${JSON.stringify({ tool: c.name, input: c.input })}\n\`\`\``,
        )
        .join('\n')
      parts.push(`## Assistant\n${[message.text, calls].filter(Boolean).join('\n')}`)
    } else {
      const results = message.results
        .map((r) => `### Result of ${r.name}${r.isError ? ' (error)' : ''}\n${r.output}`)
        .join('\n\n')
      parts.push(`## Tool results\n${results}`)
    }
  }
  return parts.join('\n\n')
}

/** the fenced form, and a bare object as a fallback for models that drop the fence */
const FENCED = new RegExp(`\`\`\`(?:${TOOL_FENCE}|json)?\\s*\\n([\\s\\S]*?)\\n?\`\`\``, 'g')

/**
 * Pull a tool call out of a reply.
 *
 * Returns null when the reply is ordinary prose, which is the signal that the
 * turn is an answer rather than a call. Only objects carrying a `tool` string
 * count, so a model that legitimately shows the user a JSON code block does not
 * get mistaken for calling something.
 */
export function parseToolCall(text: string, knownTools: string[]): AgentToolCall | null {
  for (const match of text.matchAll(FENCED)) {
    const call = asToolCall(match[1] ?? '', knownTools)
    if (call) return call
  }
  // some replies drop the fence and emit the bare object
  const trimmed = text.trim()
  if (trimmed.startsWith('{')) return asToolCall(trimmed, knownTools)
  return null
}

function asToolCall(json: string, knownTools: string[]): AgentToolCall | null {
  let parsed: unknown
  try {
    parsed = JSON.parse(json.trim())
  } catch {
    return null
  }
  if (!parsed || typeof parsed !== 'object') return null
  const record = parsed as { tool?: unknown; name?: unknown; input?: unknown; arguments?: unknown }
  // accept `name`/`arguments` too: models drift toward the OpenAI spelling
  const name = typeof record.tool === 'string' ? record.tool : record.name
  if (typeof name !== 'string' || !knownTools.includes(name)) return null
  const rawInput = record.input ?? record.arguments ?? {}
  const input =
    rawInput && typeof rawInput === 'object' && !Array.isArray(rawInput)
      ? (rawInput as Record<string, unknown>)
      : {}
  return { id: `cli_${name}_${knownTools.indexOf(name)}_${json.length}`, name, input }
}

/**
 * Strip a tool-call block out of the visible text. A model that ignores "no
 * prose" and explains itself around the call should still not leak the raw
 * JSON into the chat transcript.
 */
export function stripToolCall(text: string): string {
  return text.replace(FENCED, '').trim()
}
