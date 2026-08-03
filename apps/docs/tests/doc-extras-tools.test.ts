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
import type { HeaderFooter, NoteInfo, SectionSettings, SourceInfo } from '@genoffice/docx-engine'

const editor = {} as Editor
const numIds = { bullet: null, ordered: null }

function makeExtras() {
  const state = {
    footnote: [] as NoteInfo[],
    endnote: [] as NoteInfo[],
    watermark: null as string | null,
    sources: [] as SourceInfo[],
    header: null as HeaderFooter | null,
    footer: null as HeaderFooter | null,
    // A4 portrait with 2.54 cm margins, the shape a fresh document has
    section: {
      pageWidth: 11906,
      pageHeight: 16838,
      orientation: 'portrait',
      marginTop: 1440,
      marginRight: 1440,
      marginBottom: 1440,
      marginLeft: 1440,
      pageBorder: false,
      columns: 1,
    } as SectionSettings,
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
    headerFooter: (kind) => state[kind],
    setHeaderFooter: (kind, value) => {
      state[kind] = value
    },
    pageSetup: () => state.section,
    setPageSetup: (next) => {
      state.section = next
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

describe('set_page_setup', () => {
  it('reports the current setup when asked for no change', () => {
    const r = run('set_page_setup', {})
    expect(r.mutated).toBeFalsy()
    expect(r.output).toContain('portrait')
    expect(r.output).toContain('21') // A4 width in cm
  })

  it('turns the page landscape by swapping the sides, not by relabelling', () => {
    run('set_page_setup', { orientation: 'landscape' })
    expect(state.section.orientation).toBe('landscape')
    expect(state.section.pageWidth).toBeGreaterThan(state.section.pageHeight)
  })

  it('keeps the orientation when only the paper size changes', () => {
    run('set_page_setup', { orientation: 'landscape' })
    run('set_page_setup', { pageSize: 'letter' })
    expect(state.section.orientation).toBe('landscape')
    expect(state.section.pageWidth).toBeGreaterThan(state.section.pageHeight)
  })

  it('converts margins from centimetres', () => {
    run('set_page_setup', { marginLeftCm: 2 })
    expect(state.section.marginLeft).toBe(1134)
  })

  it('refuses a margin that would leave no page', () => {
    const r = run('set_page_setup', { marginTopCm: 40 })
    expect(r.isError).toBe(true)
    expect(state.section.marginTop).toBe(1440)
  })

  it('refuses a paper size it does not know, and lists the ones it does', () => {
    const r = run('set_page_setup', { pageSize: 'a2' })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('a4')
  })

  it('sets columns and rejects a silly count', () => {
    run('set_page_setup', { columns: 2 })
    expect(state.section.columns).toBe(2)
    expect(run('set_page_setup', { columns: 0 }).isError).toBe(true)
  })
})

describe('set_header_footer', () => {
  it('reads the current footer before it is set', () => {
    const r = run('set_header_footer', { kind: 'footer' })
    expect(r.mutated).toBeFalsy()
    expect(r.output).toMatch(/no footer/i)
  })

  it('sets footer text with an automatic page number', () => {
    run('set_header_footer', { kind: 'footer', text: 'Confidential', pageNumber: true })
    expect(state.footer).toEqual({ text: 'Confidential', pageNumber: true })
  })

  it('keeps the existing text when only the page number is toggled', () => {
    run('set_header_footer', { kind: 'footer', text: 'Draft' })
    run('set_header_footer', { kind: 'footer', pageNumber: true })
    expect(state.footer?.text).toBe('Draft')
    expect(state.footer?.pageNumber).toBe(true)
  })

  it('refuses a page number in the header, where Word does not put one', () => {
    const r = run('set_header_footer', { kind: 'header', pageNumber: true })
    expect(r.isError).toBe(true)
    expect(state.header).toBeNull()
  })

  it('treats empty text as clearing', () => {
    run('set_header_footer', { kind: 'header', text: 'Title' })
    const r = run('set_header_footer', { kind: 'header', text: '' })
    expect(state.header?.text).toBe('')
    expect(r.output).toMatch(/cleared/i)
  })

  it('rejects a kind that is neither header nor footer', () => {
    expect(run('set_header_footer', { kind: 'sidebar', text: 'x' }).isError).toBe(true)
  })
})
