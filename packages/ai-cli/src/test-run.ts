import type { AiProviderConfig } from '@genoffice/ai-provider'
import { cliStatus, type CliProviderId } from './detect'
import { streamAgentCli } from './stream'

/**
 * The settings dialog's Test, for a CLI backend.
 *
 * Unlike an HTTP provider there is no key to validate — what can go wrong is
 * that the CLI is missing, not logged in, or on a version whose flags have
 * moved. All three show up as a failed round trip, so the check is simply to
 * ask it to say one word.
 */
export async function testAgentCli(
  provider: CliProviderId,
  config: AiProviderConfig,
): Promise<{ ok: boolean; error?: string }> {
  const status = await cliStatus(provider)
  if (!status.installed) {
    return {
      ok: false,
      error: `${provider === 'claude-cli' ? 'Claude Code' : 'Codex'} was not found. Install it and make sure it runs in a terminal.`,
    }
  }

  const controller = new AbortController()
  let reply = ''
  try {
    await streamAgentCli(
      provider,
      config,
      'You are a connectivity check. Reply with exactly: OK',
      [{ role: 'user', text: 'Reply with exactly: OK' }],
      [],
      {
        signal: controller.signal,
        onDelta: (text) => {
          reply += text
        },
        // no tools were offered, so a call here would be the model inventing one
        onToolCall: () => {},
      },
    )
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
  if (!reply.trim()) return { ok: false, error: 'the CLI returned an empty reply' }
  return { ok: true }
}
