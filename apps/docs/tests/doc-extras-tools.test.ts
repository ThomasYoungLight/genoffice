/**
 * Footnotes, watermark and bibliography sources.
 *
 * All three live in React state beside the editor and are written into the
 * package on save; the References and Design ribbons have driven them for a
 * long time. The agent had no tool for any of them, so a document it wrote
 * could not carry a footnote or a citation. The accessor goes through the same
 * functions the ribbon calls, so an AI-inserted note is indistinguishable from
 * a hand-inserted one.
 */
import { describe, it, expect, beforeEach } from 'vitest'
import { executeTool, type DocExtras } from '../src/renderer/ai/tools'
import type { AgentToolCall } from '../src/shared/ipc'
import type { Editor } from '@tiptap/core'
import type { NoteInfo, SourceInfo } from '@genoffice/docx-engine'

const editor = {} as Editor
const numIds = { bullet: null, ordered: null }

function makeExtras() {
  const state = {
    footnote: [] as NoteInfo[],
    endnote: [] as NoteInfo[],
    watermark: null as string | null,
    sources: [] as SourceInfo[],
  }
  let seq = 0
  const extras: DocExtras = {
    notes: (kind) => state[kind],
    addNote: (kind, text) => {
      const id = `n${++seq}`
      state[kind] = [...state[kind], { id, text }]
      return id
    },
    editNote: (kind, id, text) => {
      if (!state[kind].some((n) => n.id === id)) return false
      state[kind] = state[kind].map((n) => (n.id === id ? { ...n, text } : n))
      return true
    },
    deleteNote: (kind, id) => {
      if (!state[kind].some((n) => n.id === id)) return false
      state[kind] = state[kind].filter((n) => n.id !== id)
      return true
    },
    watermark: () => state.watermark,
    setWatermark: (text) => {
      state.watermark = text
    },
    sources: () => state.sources,
    setSources: (list) => {
      state.sources = [...list]
    },
  }
  return { extras, state }
}

let extras: DocExtras
let state: ReturnType<typeof makeExtras>['state']
beforeEach(() => {
  const made = makeExtras()
  extras = made.extras
  state = made.state
})

const run = (name: string, input: Record<string, unknown>) =>
  executeTool(editor, { id: 't', name, input } as AgentToolCall, numIds, undefined, extras) as {
    output: string
    isError?: boolean
    mutated?: boolean
  }

describe('manage_notes', () => {
  it('adds a footnote and reports its id back for later edits', () => {
    const r = run('manage_notes', { action: 'add', text: 'Figures are unaudited.' })
    expect(r.isError).toBeUndefined()
    expect(r.mutated).toBe(true)
    expect(state.footnote).toHaveLength(1)
    expect(r.output).toContain(state.footnote[0]!.id)
  })

  it('keeps footnotes and endnotes apart', () => {
    run('manage_notes', { action: 'add', kind: 'footnote', text: 'a' })
    run('manage_notes', { action: 'add', kind: 'endnote', text: 'b' })
    expect(state.footnote.map((n) => n.text)).toEqual(['a'])
    expect(state.endnote.map((n) => n.text)).toEqual(['b'])
    expect(run('manage_notes', { action: 'list', kind: 'endnote' }).output).toContain('b')
  })

  it('edits and deletes by id', () => {
    const id =
      state.footnote[0]?.id ??
      (run('manage_notes', { action: 'add', text: 'first' }), state.footnote[0]!.id)
    expect(run('manage_notes', { action: 'edit', id, text: 'corrected' }).isError).toBeUndefined()
    expect(state.footnote[0]!.text).toBe('corrected')
    expect(run('manage_notes', { action: 'delete', id }).mutated).toBe(true)
    expect(state.footnote).toHaveLength(0)
  })

  it('refuses an id the document does not have, rather than adding a new note', () => {
    run('manage_notes', { action: 'add', text: 'only one' })
    const r = run('manage_notes', { action: 'edit', id: 'nope', text: 'x' })
    expect(r.isError).toBe(true)
    expect(state.footnote).toHaveLength(1)
    expect(state.footnote[0]!.text).toBe('only one')
  })

  it('refuses an empty note and an edit with no id', () => {
    expect(run('manage_notes', { action: 'add', text: '   ' }).isError).toBe(true)
    expect(run('manage_notes', { action: 'edit', text: 'x' }).isError).toBe(true)
    expect(state.footnote).toHaveLength(0)
  })

  it('says the document has none rather than returning an empty string', () => {
    expect(run('manage_notes', { action: 'list' }).output).toContain('no footnotes')
  })
})

describe('set_watermark', () => {
  it('sets and clears', () => {
    expect(run('set_watermark', { text: 'DRAFT' }).mutated).toBe(true)
    expect(state.watermark).toBe('DRAFT')
    run('set_watermark', { text: '' })
    expect(state.watermark).toBeNull()
  })
})

describe('manage_sources', () => {
  const src = {
    action: 'add',
    tag: 'Wang2024',
    type: 'JournalArticle',
    author: 'Wang, L.',
    title: 'Async review latency',
    year: '2024',
  }

  it('adds a source and lists it by tag', () => {
    expect(run('manage_sources', src).mutated).toBe(true)
    expect(state.sources).toHaveLength(1)
    expect(run('manage_sources', { action: 'list' }).output).toContain('Wang2024')
  })

  it('refuses a duplicate tag, because a citation resolves by it', () => {
    run('manage_sources', src)
    const r = run('manage_sources', { ...src, title: 'Something else' })
    expect(r.isError).toBe(true)
    expect(state.sources).toHaveLength(1)
  })

  it('requires at least a tag and a title', () => {
    expect(run('manage_sources', { action: 'add', tag: 'X' }).isError).toBe(true)
    expect(run('manage_sources', { action: 'add', title: 'No tag' }).isError).toBe(true)
    expect(state.sources).toHaveLength(0)
  })

  it('omits optional fields rather than storing empty strings', () => {
    run('manage_sources', src)
    expect(state.sources[0]!.publisher).toBeUndefined()
    expect(state.sources[0]!.url).toBeUndefined()
  })

  it('removes by tag and reports an unknown one', () => {
    run('manage_sources', src)
    expect(run('manage_sources', { action: 'remove', tag: 'Wang2024' }).mutated).toBe(true)
    expect(state.sources).toHaveLength(0)
    expect(run('manage_sources', { action: 'remove', tag: 'ghost' }).isError).toBe(true)
  })
})

describe('a host without the accessor', () => {
  it('reports the tool as unavailable instead of throwing', () => {
    const r = executeTool(
      editor,
      { id: 't', name: 'set_watermark', input: { text: 'X' } } as AgentToolCall,
      numIds,
    ) as { isError?: boolean; output: string }
    expect(r.isError).toBe(true)
    expect(r.output).toContain('not available')
  })
})
