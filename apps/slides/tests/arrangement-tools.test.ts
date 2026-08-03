/**
 * Arrangement and deck-level tools: grouping, stacking order, vertical text
 * anchoring, page order, and the deck footer.
 *
 * All five drove the editor over IPC long before any of them had a tool, which
 * is why a deck the agent built could not be reordered and a shape it drew last
 * always covered the text. The wiring is the easy part; what these cover is the
 * validation, because the model supplies ids and enum values from memory, and
 * the two operations that renumber things (group, move) have to say so — a
 * silent re-id is how the next tool call lands on the wrong element.
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
      box: { x: 40, y: 40 + i * 100, w: 400, h: 80, rot: 0, flipH: false, flipV: false },
      fill: { kind: 'none' },
      stroke: null,
      paragraphs: [{ runs: [{ text: id }] }],
    })),
  }) as unknown as RenderSlide

const DECK = [slideWith(['title_1', 'body_1', 'icon_1']), slideWith(['title_2'])]

function makeAccess() {
  const slidesApplied: number[] = []
  const decksApplied: number[] = []
  const access: DeckAccess = {
    getSlides: () => DECK,
    getCurrent: () => 0,
    getSelectedIds: () => [],
    applySlide: (idx) => {
      slidesApplied.push(idx)
    },
    applyDeck: (_slides, idx) => {
      decksApplied.push(idx ?? -1)
    },
    fitWidthPx: 1280,
  }
  return { access, slidesApplied, decksApplied }
}

const call = (name: string, input: Record<string, unknown>): AgentToolCall => ({
  id: 't',
  name,
  input,
})

let harness: ReturnType<typeof makeAccess>
const run = (name: string, input: Record<string, unknown>) =>
  createSlidesSkill(harness.access).executeTool!(call(name, input))
const api = () => (window as any).slidesApi

beforeEach(() => {
  harness = makeAccess()
  ;(window as any).slidesApi = {
    groupElements: vi.fn(async () => ({ slide: DECK[0], groupId: 'grp_9' })),
    reorderElement: vi.fn(async () => DECK[0]),
    setTextAnchor: vi.fn(async () => DECK[0]),
    moveSlide: vi.fn(async () => ({ slides: DECK, sections: [] })),
    applyHeaderFooter: vi.fn(async () => DECK),
  }
})

describe('group_elements', () => {
  it('groups and warns that every id on the page changed', async () => {
    const r = await run('group_elements', { slideIndex: 0, sourceIds: ['title_1', 'body_1'] })
    expect(r.isError).toBeUndefined()
    expect(r.mutated).toBe(true)
    expect(r.output).toContain('grp_9')
    expect(r.output).toMatch(/ids on this page changed/i)
    expect(harness.slidesApplied).toEqual([0])
  })

  it('refuses fewer than two elements rather than making a group of one', async () => {
    const r = await run('group_elements', { slideIndex: 0, sourceIds: ['title_1'] })
    expect(r.isError).toBe(true)
    expect(api().groupElements).not.toHaveBeenCalled()
  })

  it('names the ids it could not find instead of grouping the rest', async () => {
    const r = await run('group_elements', { slideIndex: 0, sourceIds: ['title_1', 'ghost'] })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('ghost')
    expect(r.output).not.toContain('title_1,')
    expect(api().groupElements).not.toHaveBeenCalled()
  })

  it('refuses a page that does not exist', async () => {
    const r = await run('group_elements', { slideIndex: 7, sourceIds: ['a', 'b'] })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('out of range')
  })
})

describe('reorder_element', () => {
  it('sends the direction through', async () => {
    const r = await run('reorder_element', { slideIndex: 0, sourceId: 'icon_1', dir: 'back' })
    expect(r.isError).toBeUndefined()
    expect(api().reorderElement).toHaveBeenCalledWith({
      slideIndex: 0,
      sourceId: 'icon_1',
      dir: 'back',
    })
  })

  it('rejects a direction that is not one of the four', async () => {
    const r = await run('reorder_element', { slideIndex: 0, sourceId: 'icon_1', dir: 'up' })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('front')
    expect(api().reorderElement).not.toHaveBeenCalled()
  })

  it('rejects an element that is not on the page', async () => {
    const r = await run('reorder_element', { slideIndex: 0, sourceId: 'ghost', dir: 'front' })
    expect(r.isError).toBe(true)
    expect(api().reorderElement).not.toHaveBeenCalled()
  })
})

describe('set_text_anchor', () => {
  it('anchors text vertically', async () => {
    const r = await run('set_text_anchor', {
      slideIndex: 0,
      sourceId: 'title_1',
      anchor: 'middle',
    })
    expect(r.isError).toBeUndefined()
    expect(api().setTextAnchor).toHaveBeenCalledWith({
      slideIndex: 0,
      sourceId: 'title_1',
      anchor: 'middle',
    })
  })

  it('rejects an anchor value it does not know', async () => {
    const r = await run('set_text_anchor', { slideIndex: 0, sourceId: 'title_1', anchor: 'center' })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('middle')
    expect(api().setTextAnchor).not.toHaveBeenCalled()
  })
})

describe('move_slide', () => {
  it('moves a page and warns that later page numbers shifted', async () => {
    const r = await run('move_slide', { fromIndex: 1, toIndex: 0 })
    expect(r.isError).toBeUndefined()
    expect(r.mutated).toBe(true)
    expect(r.output).toMatch(/shifted/i)
    expect(harness.decksApplied).toEqual([0])
  })

  it('treats a move to the same place as a no-op, not a change', async () => {
    const r = await run('move_slide', { fromIndex: 1, toIndex: 1 })
    expect(r.isError).toBeUndefined()
    expect(r.mutated).toBe(false)
    expect(api().moveSlide).not.toHaveBeenCalled()
  })

  it('refuses an out-of-range target', async () => {
    const r = await run('move_slide', { fromIndex: 0, toIndex: 5 })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('toIndex')
    expect(api().moveSlide).not.toHaveBeenCalled()
  })
})

describe('set_header_footer', () => {
  it('applies only the fields that were sent', async () => {
    const r = await run('set_header_footer', { footer: 'Q3 review', slideNum: true })
    expect(r.isError).toBeUndefined()
    const op = api().applyHeaderFooter.mock.calls[0][0]
    expect(op.footer).toBe('Q3 review')
    expect(op.slideNum).toBe(true)
    // an untouched field must stay untouched, not be reset to a default
    expect('date' in op).toBe(false)
    expect(r.output).toContain('Q3 review')
  })

  it('treats an empty footer as removal', async () => {
    const r = await run('set_header_footer', { footer: '' })
    expect(api().applyHeaderFooter.mock.calls[0][0].footer).toBeNull()
    expect(r.output).toMatch(/removed/i)
  })

  it('refuses a call that would change nothing', async () => {
    const r = await run('set_header_footer', {})
    expect(r.isError).toBe(true)
    expect(api().applyHeaderFooter).not.toHaveBeenCalled()
  })
})
