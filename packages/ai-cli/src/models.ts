import { spawn } from 'node:child_process'
import type { ModelListResult } from '@genoffice/ai-provider'
import { resolveCliPath, type CliProviderId } from './detect'

/**
 * Ask a locally installed agent CLI which models it can run.
 *
 * Neither CLI has a `models` subcommand, but both expose the catalogue their
 * own picker uses over the machine protocol they already speak — and that
 * catalogue is the honest one: it reflects the user's subscription and the
 * CLI's version, so it lists what `--model` will actually accept today rather
 * than what was true when this app was built.
 *
 * - Claude Code: the SDK control protocol on stdin, `subtype: "list_models"`.
 * - Codex: the app-server's JSON-RPC `model/list`.
 *
 * Both are read-only queries. Nothing is generated, no turn is started, and
 * the subprocess is killed the moment its answer arrives.
 */

/** the CLIs load plugins/hooks on startup, so first response can take a few seconds */
const LIST_TIMEOUT_MS = 45_000

/** one parsed line: the answer, a reported failure, or nothing of interest */
type Reply = { models: string[] } | { error: string } | null

/** writes a request line to the subprocess */
type Send = (payload: unknown) => void

/**
 * Claude Code's control protocol needs `--print` with stream-json on both
 * sides; without a user message it simply waits, which is what we want — the
 * control request is answered from the same loop.
 */
const CLAUDE_ARGS = [
  '--print',
  '--input-format',
  'stream-json',
  '--output-format',
  'stream-json',
  // stream-json output is rejected in print mode without it
  '--verbose',
]

const CLAUDE_REQUEST_ID = 'genoffice-list-models'

export function parseClaudeReply(line: string): Reply {
  let event: unknown
  try {
    event = JSON.parse(line)
  } catch {
    // the CLI interleaves plain-text logging with its JSON
    return null
  }
  const e = event as {
    type?: string
    response?: {
      subtype?: string
      request_id?: string
      error?: string
      response?: { models?: Array<{ value?: string }> }
    }
  }
  if (e.type !== 'control_response' || e.response?.request_id !== CLAUDE_REQUEST_ID) return null
  if (e.response.subtype !== 'success') {
    return { error: e.response.error || 'the CLI rejected the model list request' }
  }
  const models: string[] = []
  for (const model of e.response.response?.models ?? []) {
    // `value` is what --model takes; "default" is the CLI's own choice, which
    // this app already expresses by leaving the model field empty
    if (model.value && model.value !== 'default') models.push(model.value)
  }
  return { models }
}

const CODEX_INIT_ID = 1
const CODEX_LIST_ID = 2

export function parseCodexReply(line: string): Reply {
  let event: unknown
  try {
    event = JSON.parse(line)
  } catch {
    return null
  }
  const e = event as {
    id?: number
    error?: { message?: string }
    result?: { data?: Array<{ id?: string; model?: string; hidden?: boolean }> }
  }
  if (e.id !== CODEX_LIST_ID) return null
  if (e.error) return { error: e.error.message || 'the CLI rejected the model list request' }
  const models: string[] = []
  for (const model of e.result?.data ?? []) {
    // `hidden` marks models the CLI keeps out of its own picker (deprecated or
    // internal); matching that keeps the two lists saying the same thing
    if (model.hidden) continue
    const id = model.id ?? model.model
    if (id) models.push(id)
  }
  return { models }
}

/** Ask a CLI for the models it can run. */
export async function listCliModels(provider: CliProviderId): Promise<ModelListResult> {
  const bin = resolveCliPath(provider)
  if (!bin) {
    return {
      ok: false,
      error: `${provider === 'claude-cli' ? 'Claude Code' : 'Codex'} was not found. Install it and make sure it runs in a terminal.`,
    }
  }

  const claude = provider === 'claude-cli'
  const args = claude ? CLAUDE_ARGS : ['app-server']
  const parse = claude ? parseClaudeReply : parseCodexReply
  const open = (send: Send) => {
    if (claude) {
      send({
        type: 'control_request',
        request_id: CLAUDE_REQUEST_ID,
        request: { subtype: 'list_models' },
      })
      return
    }
    // the app-server expects a handshake first; requests may be pipelined
    send({
      jsonrpc: '2.0',
      id: CODEX_INIT_ID,
      method: 'initialize',
      params: { clientInfo: { name: 'genoffice', title: null, version: '1' }, capabilities: null },
    })
    // limit well past the size of any real picker, so paging never comes up
    send({ jsonrpc: '2.0', id: CODEX_LIST_ID, method: 'model/list', params: { limit: 100 } })
  }

  const reply = await queryCli(bin, args, open, parse)
  if ('error' in reply) return { ok: false, error: reply.error }
  if (!reply.models.length) return { ok: false, error: 'the CLI reported no models' }
  // Deliberately not sorted, unlike the HTTP catalogues: these lists are short
  // and the CLI returns them best-first, which is more useful than alphabetical.
  return { ok: true, models: [...new Set(reply.models)] }
}

/**
 * Run a CLI in its machine-protocol mode, send the opening request(s), and
 * resolve as soon as a line parses into an answer.
 */
function queryCli(
  bin: string,
  args: string[],
  open: (send: Send) => void,
  parse: (line: string) => Reply,
): Promise<{ models: string[] } | { error: string }> {
  return new Promise((resolve) => {
    const child = spawn(bin, args, {
      stdio: ['pipe', 'pipe', 'pipe'],
      // same reasoning as streamAgentCli: started in the app's cwd, a coding
      // CLI would pick up that directory's CLAUDE.md / AGENTS.md
      cwd: process.env.HOME || undefined,
      env: { ...process.env, NO_COLOR: '1' },
    })

    let settled = false
    const finish = (result: { models: string[] } | { error: string }) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      child.kill('SIGKILL')
      resolve(result)
    }
    const timer = setTimeout(
      () => finish({ error: `${bin} did not answer within ${LIST_TIMEOUT_MS / 1000}s` }),
      LIST_TIMEOUT_MS,
    )

    let buffer = ''
    let stderr = ''
    child.stdout.setEncoding('utf8')
    child.stdout.on('data', (chunk: string) => {
      buffer += chunk
      const lines = buffer.split('\n')
      buffer = lines.pop() ?? ''
      for (const line of lines) {
        if (!line.trim()) continue
        const reply = parse(line)
        if (reply) finish(reply)
      }
    })
    child.stderr.setEncoding('utf8')
    child.stderr.on('data', (chunk: string) => {
      stderr += chunk
    })
    child.on('error', (err) => finish({ error: err.message }))
    child.on('close', () =>
      finish({
        error: stderr.trim().split('\n').slice(-2).join(' ') || 'the CLI produced no list',
      }),
    )

    try {
      open((payload) => child.stdin.write(`${JSON.stringify(payload)}\n`))
    } catch (err) {
      finish({ error: err instanceof Error ? err.message : String(err) })
    }
  })
}
