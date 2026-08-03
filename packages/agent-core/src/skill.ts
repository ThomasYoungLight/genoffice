import type { AgentToolCall, AgentToolDef, ToolExecution } from './types'

/**
 * A skill packages one capability domain for the agent loop: its system
 * prompt section, its tools, per-turn context, and the tool executor.
 * AI Docs ships a docx skill; Excel / PPT skills plug in the same way.
 */
export interface AgentSkill {
  id: string
  /** system prompt section describing this skill's rules and tools */
  systemPrompt: string
  tools: AgentToolDef[]
  /**
   * Fresh context sections attached to every user turn (e.g. document
   * skeleton + selection). Return '' when there is nothing to attach.
   */
  buildContext?(): string
  /**
   * signal: aborted when the user hits stop. Long-running tools (e.g.
   * generate_deck with internal LLM calls) should check signal.aborted in
   * their loops and stop promptly.
   */
  executeTool(call: AgentToolCall, signal?: AbortSignal): ToolExecution | Promise<ToolExecution>
}

/**
 * Merge several skills into one (tool names must be globally unique).
 * `intro` becomes the shared preamble of the combined system prompt.
 */
export function composeSkills(id: string, intro: string, skills: AgentSkill[]): AgentSkill {
  // name collisions are a programming error, so check once up front
  const seen = new Set<string>()
  for (const skill of skills) {
    for (const tool of skill.tools) {
      if (seen.has(tool.name)) throw new Error(`duplicate tool name: ${tool.name}`)
      seen.add(tool.name)
    }
  }
  return {
    id,
    systemPrompt: [intro, ...skills.map((s) => s.systemPrompt)].filter(Boolean).join('\n\n'),
    /**
     * Recomputed per read rather than flattened once: a skill may vary the
     * tools it offers over the life of a conversation (slides withholds the
     * Genspark-only tools until that account is known to be available), and a
     * composed skill is built when the panel mounts, well before such a check
     * can have answered.
     */
    get tools(): AgentToolDef[] {
      return skills.flatMap((s) => s.tools)
    },
    buildContext: () =>
      skills
        .map((s) => s.buildContext?.() ?? '')
        .filter(Boolean)
        .join('\n\n'),
    executeTool: (call, signal) => {
      // resolved against the current tool lists for the same reason
      const skill = skills.find((s) => s.tools.some((tool) => tool.name === call.name))
      if (!skill) {
        return { output: `Unknown tool: ${call.name}`, isError: true, summary: call.name }
      }
      return skill.executeTool(call, signal)
    },
  }
}
