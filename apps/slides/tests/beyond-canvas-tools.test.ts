/**
 * Tools for the parts of a deck that are not on the canvas: speaker notes,
 * hyperlinks, transitions and animations.
 *
 * The editor and the pptx engine have supported all four for a long time —
 * `slides:set-notes`, `slides:set-link`, `slides:set-transition`,
 * `slides:set-animations` — but no agent tool reached them, so a generated
 * deck arrived with no notes and no links. These cover the wiring and, more
 * importantly, the validation: the model supplies element ids and enum values
 * from memory, and a wrong one must be a refusal with a reason rather than a
 * silently dropped write.
 */
import { describe, it, expect, vi, beforeEach } from 'vitest'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'
import type { RenderSlide } from '@genoffice/pptx-render'
import type { AgentToolCall } from '../src/shared/ipc'

const slideWith = (ids: string[]) =>
  ({
    widthPx: 1280,
    heightPx: 720,
    background: null,
    nodes: ids.map((id, i) => ({
      id,
      type: 'text',
      sourceId: id,
      box: { x: 40, y: 40 + i * 120, w: 400, h: 80, rot: 0, flipH: false, flipV: false },
      fill: { kind: 'none' },
      stroke: null,
      paragraphs: [{ runs: [{ text: id }] }],
    })),
  }) as unknown as RenderSlide

const DECK = [slideWith(['title_1', 'body_1']), slideWith(['title_2'])]

function makeAccess() {
  const appliedSlides: number[] = []
  const access: DeckAccess = {
    getSlides: () => DECK,
    getCurrent: () => 0,
    getSelectedIds: () => [],
    applySlide: (idx) => {
      appliedSlides.push(idx)
    },
    applyDeck: () => {},
    fitWidthPx: 1280,
  }
  return { access, appliedSlides }
}

const call = (name: string, input: Record<string, unknown>): AgentToolCall => ({
  id: 't',
  name,
  input,
})

const run = (name: string, input: Record<string, unknown>) =>
  createSlidesSkill(makeAccess().access).executeTool!(call(name, input))

beforeEach(() => {
  ;(window as any).slidesApi = {
    setNotes: vi.fn(async () => true),
    getNotes: vi.fn(async () => ''),
    getSlideLinks: vi.fn(async () => []),
    setLink: vi.fn(async () => DECK[0]),
    setTransition: vi.fn(async () => true),
    setAnimations: vi.fn(async () => true),
  }
})

describe('set_slide_notes', () => {
  it('writes the notes and reports it as a document change', async () => {
    const r = await run('set_slide_notes', { slideIndex: 1, text: 'Lead with the 18% number.' })
    expect(r.isError).toBeUndefined()
    expect(r.mutated).toBe(true)
    expect((window as any).slidesApi.setNotes).toHaveBeenCalledWith({
      slideIndex: 1,
      text: 'Lead with the 18% number.',
    })
  })

  it('treats an empty string as "clear", not as a mistake', async () => {
    const r = await run('set_slide_notes', { slideIndex: 0, text: '' })
    expect(r.isError).toBeUndefined()
    expect(r.output).toContain('Cleared')
  })

  it('refuses a page that does not exist', async () => {
    const r = await run('set_slide_notes', { slideIndex: 9, text: 'x' })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('out of range')
    expect((window as any).slidesApi.setNotes).not.toHaveBeenCalled()
  })
})

describe('set_element_link', () => {
  it('links to a url', async () => {
    const r = await run('set_element_link', {
      slideIndex: 0,
      sourceId: 'title_1',
      url: 'https://example.com/pricing',
    })
    expect(r.mutated).toBe(true)
    expect((window as any).slidesApi.setLink).toHaveBeenCalledWith({
      slideIndex: 0,
      sourceId: 'title_1',
      target: { kind: 'url', url: 'https://example.com/pricing' },
    })
  })

  it('links to another page, for an agenda entry', async () => {
    await run('set_element_link', { slideIndex: 0, sourceId: 'body_1', targetSlideIndex: 1 })
    expect((window as any).slidesApi.setLink).toHaveBeenCalledWith(
      expect.objectContaining({ target: { kind: 'slide', slideIndex: 1 } }),
    )
  })

  it('clears the link when given no target', async () => {
    await run('set_element_link', { slideIndex: 0, sourceId: 'title_1' })
    expect((window as any).slidesApi.setLink).toHaveBeenCalledWith(
      expect.objectContaining({ target: null }),
    )
  })

  it('refuses an element that is not on the page, and a non-http url', async () => {
    const missing = await run('set_element_link', {
      slideIndex: 0,
      sourceId: 'nope',
      url: 'https://x.com',
    })
    expect(missing.isError).toBe(true)
    expect(missing.output).toContain('not found')

    const scheme = await run('set_element_link', {
      slideIndex: 0,
      sourceId: 'title_1',
      url: 'javascript:alert(1)',
    })
    expect(scheme.isError).toBe(true)
    expect(scheme.output).toContain('http')
    expect((window as any).slidesApi.setLink).not.toHaveBeenCalled()
  })

  it('refuses a jump to a page that does not exist', async () => {
    const r = await run('set_element_link', {
      slideIndex: 0,
      sourceId: 'title_1',
      targetSlideIndex: 7,
    })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('out of range')
  })
})

describe('set_slide_transition', () => {
  it('applies to the whole deck with -1', async () => {
    const r = await run('set_slide_transition', { slideIndex: -1, kind: 'fade' })
    expect(r.mutated).toBe(true)
    expect(r.output).toContain('all 2 pages')
    expect((window as any).slidesApi.setTransition).toHaveBeenCalledWith({
      slideIndex: -1,
      kind: 'fade',
    })
  })

  it('refuses an effect the format does not have', async () => {
    const r = await run('set_slide_transition', { slideIndex: 0, kind: 'swoosh' })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('swoosh')
    expect((window as any).slidesApi.setTransition).not.toHaveBeenCalled()
  })
})

describe('set_slide_animations', () => {
  it('keeps the given order and fills the defaults the model omits', async () => {
    const r = await run('set_slide_animations', {
      slideIndex: 0,
      items: [
        { sourceId: 'body_1', effect: 'fade', trigger: 'afterPrev', paragraph: 2 },
        { sourceId: 'title_1', effect: 'wipe' },
      ],
    })
    expect(r.mutated).toBe(true)
    const op = (window as any).slidesApi.setAnimations.mock.calls[0][0]
    expect(op.items.map((i: { sourceId: string }) => i.sourceId)).toEqual(['body_1', 'title_1'])
    expect(op.items[0]).toEqual({
      sourceId: 'body_1',
      effect: 'fade',
      trigger: 'afterPrev',
      durationMs: 500,
      delayMs: 0,
      paragraph: 2,
    })
    // an omitted trigger becomes onClick rather than undefined reaching the engine
    expect(op.items[1].trigger).toBe('onClick')
  })

  it('clears with an empty list', async () => {
    const r = await run('set_slide_animations', { slideIndex: 0, items: [] })
    expect(r.output).toContain('Cleared')
  })

  it('refuses the whole list when one entry is wrong, rather than writing a partial one', async () => {
    const badEffect = await run('set_slide_animations', {
      slideIndex: 0,
      items: [
        { sourceId: 'title_1', effect: 'fade' },
        { sourceId: 'body_1', effect: 'explode' },
      ],
    })
    expect(badEffect.isError).toBe(true)
    expect(badEffect.output).toContain('explode')

    const badTarget = await run('set_slide_animations', {
      slideIndex: 0,
      items: [{ sourceId: 'ghost', effect: 'fade' }],
    })
    expect(badTarget.isError).toBe(true)
    expect(badTarget.output).toContain('ghost')
    expect((window as any).slidesApi.setAnimations).not.toHaveBeenCalled()
  })
})

describe('read_slide', () => {
  it('reports the notes so the model can tell a noted page from an empty one', async () => {
    ;(window as any).slidesApi.getNotes = vi.fn(async () => 'Open with the churn number.')
    ;(window as any).slidesApi.getSlideLinks = vi.fn(async () => [
      { sourceId: 'title_1', target: { kind: 'url', url: 'https://x' } },
    ])
    const r = await run('read_slide', { slideIndex: 0 })
    expect(r.output).toContain('Open with the churn number.')
    expect(r.output).toContain('title_1')
  })

  it('says so explicitly when a page has none', async () => {
    const r = await run('read_slide', { slideIndex: 0 })
    expect(r.output).toContain('(none)')
  })
})

describe('manage_sections', () => {
  const SECTIONS = [
    { id: '{A}', name: 'Intro', slideIndices: [0] },
    { id: '{B}', name: 'Detail', slideIndices: [1] },
  ]

  beforeEach(() => {
    Object.assign((window as any).slidesApi, {
      getSections: vi.fn(async () => SECTIONS),
      addSection: vi.fn(async () => SECTIONS),
      renameSection: vi.fn(async () => SECTIONS),
      removeSection: vi.fn(async () => SECTIONS),
      moveSection: vi.fn(async () => ({ slides: DECK, sections: SECTIONS })),
    })
  })

  it('lists sections with the ids a later call needs', async () => {
    const r = await run('manage_sections', { action: 'list' })
    expect(r.mutated).toBe(false)
    expect(r.output).toContain('{A}')
    expect(r.output).toContain('pages 1')
  })

  it('adds a section at a page', async () => {
    const r = await run('manage_sections', { action: 'add', atSlideIndex: 1, name: 'Detail' })
    expect(r.mutated).toBe(true)
    expect((window as any).slidesApi.addSection).toHaveBeenCalledWith({
      atSlideIndex: 1,
      name: 'Detail',
    })
  })

  it('refuses an unnamed section and an out-of-range page', async () => {
    expect(
      (await run('manage_sections', { action: 'add', atSlideIndex: 0, name: '  ' })).isError,
    ).toBe(true)
    expect(
      (await run('manage_sections', { action: 'add', atSlideIndex: 9, name: 'X' })).isError,
    ).toBe(true)
    expect((window as any).slidesApi.addSection).not.toHaveBeenCalled()
  })

  it('rejects rename and remove without an id, rather than guessing one', async () => {
    expect((await run('manage_sections', { action: 'rename', name: 'X' })).isError).toBe(true)
    expect((await run('manage_sections', { action: 'remove' })).isError).toBe(true)
  })

  it('applies the reordered deck when a section moves', async () => {
    const r = await run('manage_sections', { action: 'move', id: '{B}', dir: 'up' })
    expect(r.mutated).toBe(true)
    expect((window as any).slidesApi.moveSection).toHaveBeenCalledWith({ id: '{B}', dir: 'up' })
  })

  it('says so on an unknown action instead of doing nothing quietly', async () => {
    const r = await run('manage_sections', { action: 'reticulate' })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('reticulate')
  })
})

describe('manage_comments', () => {
  const COMMENTS = [
    {
      authorId: 1,
      author: 'Sam',
      initials: 'S',
      dt: '2026-01-01T00:00:00Z',
      idx: 1,
      text: 'Check this number',
    },
  ]

  beforeEach(() => {
    Object.assign((window as any).slidesApi, {
      getComments: vi.fn(async () => COMMENTS),
      addComment: vi.fn(async () => COMMENTS),
      deleteComment: vi.fn(async () => COMMENTS),
    })
  })

  it('lists comments with the keys removal needs', async () => {
    const r = await run('manage_comments', { action: 'list', slideIndex: 0 })
    expect(r.output).toContain('authorId=1')
    expect(r.output).toContain('idx=1')
    expect(r.output).toContain('Check this number')
    expect(r.mutated).toBe(false)
  })

  it('reports an empty page plainly', async () => {
    ;(window as any).slidesApi.getComments = vi.fn(async () => [])
    const r = await run('manage_comments', { action: 'list', slideIndex: 1 })
    expect(r.output).toContain('no comments')
  })

  it('adds a comment', async () => {
    const r = await run('manage_comments', { action: 'add', slideIndex: 0, text: 'Needs a source' })
    expect(r.mutated).toBe(true)
    expect((window as any).slidesApi.addComment).toHaveBeenCalledWith({
      slideIndex: 0,
      text: 'Needs a source',
    })
  })

  it('will not add an empty comment or remove without both keys', async () => {
    expect(
      (await run('manage_comments', { action: 'add', slideIndex: 0, text: ' ' })).isError,
    ).toBe(true)
    expect(
      (await run('manage_comments', { action: 'remove', slideIndex: 0, authorId: 1 })).isError,
    ).toBe(true)
    expect((window as any).slidesApi.deleteComment).not.toHaveBeenCalled()
  })
})

/**
 * insert_diagram: Mermaid → native shapes and arrows. SmartArt offers seven
 * fixed layouts and an arbitrary graph is not one of them, so this is the only
 * way the agent can draw a process or a decision tree.
 */
describe('insert_diagram', () => {
  beforeEach(() => {
    let n = 0
    Object.assign((window as any).slidesApi, {
      addElement: vi.fn(async () => ({ slide: DECK[0], sourceId: `el_${++n}` })),
      flipElements: vi.fn(async () => DECK[0]),
    })
  })

  it('draws a node per box and an arrow per edge', async () => {
    const r = await run('insert_diagram', {
      slideIndex: 0,
      mermaid: 'flowchart TD; A[Submit] --> B{Approved?}; B --> C[Ship]',
    })
    expect(r.isError).toBeUndefined()
    expect(r.mutated).toBe(true)
    const calls = (window as any).slidesApi.addElement.mock.calls.map((c: any[]) => c[0].kind)
    expect(calls.filter((k: string) => k !== 'lineArrow')).toEqual(['rect', 'diamond', 'rect'])
    expect(calls.filter((k: string) => k === 'lineArrow')).toHaveLength(2)
  })

  it('flips an arrow that runs backwards inside its own box', async () => {
    // two children: the left one is reached by a right-to-left arrow
    await run('insert_diagram', { slideIndex: 0, mermaid: 'flowchart TD; A --> B; A --> C' })
    const flips = (window as any).slidesApi.flipElements.mock.calls
    expect(flips.length).toBe(1)
    expect(flips[0][0].axis).toBe('h')
  })

  it('refuses a diagram type it would otherwise draw wrong', async () => {
    const r = await run('insert_diagram', {
      slideIndex: 0,
      mermaid: 'sequenceDiagram\n A->>B: hello',
    })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('sequenceDiagram')
    expect((window as any).slidesApi.addElement).not.toHaveBeenCalled()
  })

  it('says which arrow labels it could not draw rather than dropping them silently', async () => {
    const r = await run('insert_diagram', {
      slideIndex: 0,
      mermaid: 'flowchart TD; A -->|yes| B',
    })
    expect(r.output).toContain('1 arrow label')
  })

  it('refuses a page that does not exist', async () => {
    const r = await run('insert_diagram', { slideIndex: 9, mermaid: 'flowchart TD; A --> B' })
    expect(r.isError).toBe(true)
    expect((window as any).slidesApi.addElement).not.toHaveBeenCalled()
  })
})
