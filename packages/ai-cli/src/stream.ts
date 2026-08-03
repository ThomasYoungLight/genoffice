import { spawn } from 'node:child_process'
import type { AgentMessage, AgentToolDef } from '@genoffice/agent-core'
import type { AiProviderConfig } from '@genoffice/ai-provider'
import { resolveCliPath, type CliProviderId } from './detect'
import { parseToolCall, serializeConversation, toolProtocolPrompt } from './protocol'

/**
 * Run a locally installed agent CLI as if it were a streaming model.
 *
 * Both CLIs are driven in their non-interactive JSON modes with their own
 * tools switched off, so they behave as a plain text generator over the user's
 * existing subscription — no API key, no per-token billing from this app. The
 * app's tools travel through the prompt protocol (see `protocol.ts`).
 */

/** callbacks match @genoffice/ai-provider's StreamCallbacks so mains can swap implementations */
export interface CliStreamCallbacks {
  onDelta: (text: string) => void
  onToolCall: (call: { id: string; name: string; input: Record<string, unknown> }) => void
  signal: AbortSignal
}

/** a CLI turn can involve model thinking time well past a normal HTTP timeout */
const CLI_TIMEOUT_MS = 10 * 60_000

function buildArgs(provider: CliProviderId, config: AiProviderConfig, system: string): string[] {
  if (provider === 'claude-cli') {
    return [
      '--print',
      // wraps the vendor's own streaming events, so deltas arrive incrementally
      '--output-format',
      'stream-json',
      '--include-partial-messages',
      '--verbose',
      // its built-in file/shell tools are irrelevant here and actively unwanted:
      // the document being edited lives in the renderer, not on disk
      '--tools',
      '',
      '--system-prompt',
      system,
      ...(config.model ? ['--model', config.model] : []),
    ]
  }
  return [
    'exec',
    '--json',
    // the CLI is used as a generator, not an agent let loose on the filesystem
    '--sandbox',
    'read-only',
    '--skip-git-repo-check',
    ...(config.model ? ['--model', config.model] : []),
  ]
}

/** Codex takes no system-prompt flag, so its instructions lead the prompt instead */
function buildPrompt(provider: CliProviderId, system: string, conversation: string): string {
  return provider === 'claude-cli' ? conversation : `${system}\n\n---\n\n${conversation}`
}

/**
 * Pull assistant text out of one line of the CLI's JSON output.
 *
 * Claude Code re-emits the vendor's raw streaming events, so text arrives as
 * incremental deltas. Codex reports completed items, so its text arrives whole.
 * Everything else on the stream (hook chatter, rate-limit notices, usage) is
 * ignored.
 */
export function extractText(provider: CliProviderId, line: string): string | null {
  let event: unknown
  try {
    event = JSON.parse(line)
  } catch {
    // CLIs interleave plain-text logging with their JSON; skip it
    return null
  }
  if (!event || typeof event !== 'object') return null

  if (provider === 'claude-cli') {
    const e = event as {
      type?: string
      event?: { type?: string; delta?: { type?: string; text?: string } }
    }
    if (e.type !== 'stream_event') return null
    const inner = e.event
    if (inner?.type !== 'content_block_delta') return null
    return inner.delta?.type === 'text_delta' ? (inner.delta.text ?? null) : null
  }

  const e = event as { type?: string; item?: { type?: string; text?: string } }
  if (e.type !== 'item.completed') return null
  return e.item?.type === 'agent_message' ? (e.item.text ?? null) : null
}

/** an error the CLI reported on its own event stream */
export function extractError(provider: CliProviderId, line: string): string | null {
  let event: unknown
  try {
    event = JSON.parse(line)
  } catch {
    return null
  }
  const e = event as {
    type?: string
    subtype?: string
    is_error?: boolean
    result?: unknown
    item?: { type?: string; message?: string }
  }
  if (provider === 'claude-cli') {
    if (e.type === 'result' && e.is_error) return String(e.result ?? 'the CLI reported an error')
    return null
  }
  // Codex reports non-fatal warnings as error items too, so only a turn that
  // produced no text at all should surface them; the caller decides that
  if (e.type === 'item.completed' && e.item?.type === 'error') return e.item.message ?? null
  return null
}

export async function streamAgentCli(
  provider: CliProviderId,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  cb: CliStreamCallbacks,
): Promise<void> {
  const bin = resolveCliPath(provider)
  if (!bin) throw new Error(`${provider} is not installed or could not be found on PATH`)

  const fullSystem = system + toolProtocolPrompt(tools)
  const prompt = buildPrompt(provider, fullSystem, serializeConversation(messages))
  const child = spawn(bin, buildArgs(provider, config, fullSystem), {
    stdio: ['pipe', 'pipe', 'pipe'],
    // a coding CLI started in the app's cwd would pick up that directory's
    // CLAUDE.md/AGENTS.md; the home directory keeps the run neutral
    cwd: process.env.HOME || undefined,
    env: { ...process.env, NO_COLOR: '1' },
  })

  const killTimer = setTimeout(() => child.kill('SIGKILL'), CLI_TIMEOUT_MS)
  const onAbort = () => child.kill('SIGTERM')
  cb.signal.addEventListener('abort', onAbort, { once: true })

  // the prompt goes over stdin: an argv-sized conversation would blow the
  // platform's argument limit on any real document
  child.stdin.end(prompt)

  let assistantText = ''
  const errors: string[] = []
  let stderr = ''

  try {
    await new Promise<void>((resolve, reject) => {
      let buffer = ''
      child.stdout.setEncoding('utf8')
      child.stdout.on('data', (chunk: string) => {
        buffer += chunk
        const lines = buffer.split('\n')
        buffer = lines.pop() ?? ''
        for (const line of lines) handleLine(line)
      })
      child.stderr.setEncoding('utf8')
      child.stderr.on('data', (chunk: string) => {
        stderr += chunk
      })
      child.on('error', reject)
      child.on('close', (code) => {
        if (buffer) handleLine(buffer)
        if (cb.signal.aborted) return resolve()
        if (code !== 0) {
          const detail = errors.join('; ') || stderr.trim().split('\n').slice(-3).join(' ')
          return reject(new Error(`${bin} exited with code ${code}${detail ? `: ${detail}` : ''}`))
        }
        resolve()
      })

      function handleLine(line: string): void {
        const trimmed = line.trim()
        if (!trimmed) return
        const error = extractError(provider, trimmed)
        if (error) errors.push(error)
        const text = extractText(provider, trimmed)
        if (text === null) return
        assistantText += text
        // Text is buffered rather than forwarded as it arrives: a tool call can
        // only be recognised once its JSON block is complete, and streaming the
        // raw block into the chat would show the user the protocol's plumbing.
        // Prose-only replies are flushed below, so the cost is that a CLI turn
        // appears at once instead of typing out.
      }
    })
  } finally {
    clearTimeout(killTimer)
    cb.signal.removeEventListener('abort', onAbort)
  }

  if (cb.signal.aborted) return

  const call = parseToolCall(
    assistantText,
    tools.map((t) => t.name),
  )
  if (call) {
    // Prose alongside a call is off-protocol, and in practice it is a false
    // completion claim: the model narrates "Done! I replaced b7" in the same
    // breath as asking for the edit that has not happened yet. A native
    // tool-calling API separates preamble from result; this one cannot, so the
    // narration is dropped rather than shown to the user as fact. The tool chip
    // the panel renders already says what is running.
    cb.onToolCall(call)
    return
  }
  if (assistantText.trim()) {
    cb.onDelta(assistantText)
    return
  }
  // no text and no call: surface whatever the CLI complained about
  throw new Error(errors.join('; ') || stderr.trim() || 'the CLI produced no output')
}
