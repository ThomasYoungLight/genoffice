/**
 * The raw OOXML escape hatch, at the tool layer.
 *
 * The engine owns the validation (unique match, well-formed result, slides
 * still parse) and is tested there. What matters here is that the tool does not
 * soften any of it: a refusal from the engine reaches the model as an error with
 * the engine's own wording — those messages tell it what to do differently, and
 * a tool that summarised them to "raw edit failed" would leave it guessing.
 */
import { describe, it, expect, vi, beforeEach } from 'vitest'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'
import type { RenderSlide } from '@genoffice/pptx-render'
import type { AgentToolCall } from '../src/shared/ipc'

const slide = () =>
  ({
    widthPx: 1280,
    heightPx: 720,
    background: null,
    nodes: [],
  }) as unknown as RenderSlide

const DECK = [slide(), slide()]

function makeAccess() {
  const applied: RenderSlide[][] = []
  const access: DeckAccess = {
    getSlides: () => DECK,
    getCurrent: () => 0,
    getSelectedIds: () => [],
    applySlide: () => {},
    applyDeck: (slides) => {
      applied.push(slides)
    },
    fitWidthPx: 1280,
  }
  return { access, applied }
}

const call = (name: string, input: Record<string, unknown>): AgentToolCall => ({
  id: 't',
  name,
  input,
})

let harness: ReturnType<typeof makeAccess>
const run = (name: string, input: Record<string, unknown>) =>
  createSlidesSkill(harness.access).executeTool!(call(name, input))

beforeEach(() => {
  harness = makeAccess()
  ;(window as any).slidesApi = {
    rawParts: vi.fn(async () => [
      { path: 'ppt/presentation.xml', ref: '/presentation', bytes: 3210 },
      { path: 'ppt/slides/slide1.xml', ref: '/slide[1]', bytes: 8400 },
      { path: 'ppt/theme/theme1.xml', bytes: 12000 },
    ]),
    rawGet: vi.fn(async () => ({ ok: true, path: 'ppt/slides/slide1.xml', xml: '<p:sld/>' })),
    rawSet: vi.fn(async () => ({
      ok: true,
      path: 'ppt/slides/slide1.xml',
      reparsedSlides: 1,
      slides: DECK,
    })),
  }
})

describe('read_raw_xml', () => {
  it('lists the parts when asked for none, naming the short forms', async () => {
    const r = await run('read_raw_xml', {})
    expect(r.isError).toBeUndefined()
    expect(r.mutated).toBe(false)
    expect(r.output).toContain('/slide[1]')
    expect(r.output).toContain('ppt/theme/theme1.xml')
    expect((window as any).slidesApi.rawGet).not.toHaveBeenCalled()
  })

  it('returns the part text with its resolved path', async () => {
    const r = await run('read_raw_xml', { part: '/slide[1]' })
    expect(r.output).toContain('ppt/slides/slide1.xml')
    expect(r.output).toContain('<p:sld/>')
    expect((window as any).slidesApi.rawGet).toHaveBeenCalledWith('/slide[1]')
  })

  it('passes a read failure back verbatim', async () => {
    ;(window as any).slidesApi.rawGet = vi.fn(async () => ({
      ok: false,
      error: 'No part matches "/slide[9]"',
    }))
    const r = await run('read_raw_xml', { part: '/slide[9]' })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('/slide[9]')
  })

  it('never reports a read as a document change', async () => {
    expect((await run('read_raw_xml', { part: '/theme' })).mutated).toBe(false)
  })
})

describe('edit_raw_xml', () => {
  it('applies the edit and re-renders the deck', async () => {
    const r = await run('edit_raw_xml', {
      part: '/slide[1]',
      find: '<a:t>Old</a:t>',
      replace: '<a:t>New</a:t>',
    })
    expect(r.isError).toBeUndefined()
    expect(r.mutated).toBe(true)
    expect((window as any).slidesApi.rawSet).toHaveBeenCalledWith({
      ref: '/slide[1]',
      find: '<a:t>Old</a:t>',
      replace: '<a:t>New</a:t>',
      fitWidthPx: 1280,
    })
    expect(harness.applied).toHaveLength(1)
  })

  it('tells the model to verify the result, since raw XML bypasses the model', async () => {
    const r = await run('edit_raw_xml', { part: '/slide[1]', find: 'a', replace: 'b' })
    expect(r.output).toMatch(/tell the user/i)
  })

  it('passes the ambiguous-match refusal through with its count', async () => {
    ;(window as any).slidesApi.rawSet = vi.fn(async () => ({
      ok: false,
      error: 'find occurs 7 times in "ppt/slides/slide1.xml"; it must match exactly once',
    }))
    const r = await run('edit_raw_xml', { part: '/slide[1]', find: '<a:t>', replace: 'x' })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('7 times')
    expect(r.mutated).toBe(false)
    expect(harness.applied).toHaveLength(0)
  })

  it('passes the malformed-XML refusal through', async () => {
    ;(window as any).slidesApi.rawSet = vi.fn(async () => ({
      ok: false,
      error: 'The result is not well-formed XML: Unclosed tag (line 1)',
    }))
    const r = await run('edit_raw_xml', { part: '/slide[1]', find: 'a', replace: '<b>' })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('well-formed')
  })

  it('refuses an empty find before reaching the document', async () => {
    const r = await run('edit_raw_xml', { part: '/slide[1]', find: '', replace: 'x' })
    expect(r.isError).toBe(true)
    expect((window as any).slidesApi.rawSet).not.toHaveBeenCalled()
  })

  it('refuses a missing part reference', async () => {
    const r = await run('edit_raw_xml', { find: 'a', replace: 'b' })
    expect(r.isError).toBe(true)
    expect((window as any).slidesApi.rawSet).not.toHaveBeenCalled()
  })

  it('treats an empty replace as a deletion, not as a missing argument', async () => {
    const r = await run('edit_raw_xml', { part: '/slide[1]', find: '<a:effectLst/>', replace: '' })
    expect(r.isError).toBeUndefined()
    expect((window as any).slidesApi.rawSet).toHaveBeenCalledWith(
      expect.objectContaining({ replace: '' }),
    )
  })
})
