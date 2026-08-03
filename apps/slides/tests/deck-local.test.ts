/**
 * Local deck generation: the page loop that runs when there is no cloud
 * service. What matters here is what the user ends up with — every planned
 * page present, in order, with the deck it replaced gone — and that one bad
 * page does not take the deck down with it.
 */
import { describe, expect, it, vi } from 'vitest'
import {
  generateDeckLocally,
  pageContentFrom,
  type LocalDeckArgs,
} from '../src/renderer/ai/deck-local'
import { CANVAS, type RenderedPage } from '../src/renderer/ai/deck-layout'

const plan = (n: number): Array<Record<string, unknown>> =>
  Array.from({ length: n }, (_, i) => ({
    title: `Page ${i + 1}`,
    type: i === 0 ? 'cover' : 'content',
    brief: `brief ${i + 1}`,
    layout: i === 0 ? 'cover_dark_minimal' : 'three_column_cards',
    image_queries: [],
  }))

function harness(over: Partial<LocalDeckArgs> = {}) {
  const landed: Array<{ page: RenderedPage; first: boolean; replaceExisting: boolean }> = []
  const asked: Array<Record<string, unknown>> = []
  const args: LocalDeckArgs = {
    pages: plan(3),
    coreHook: 'hook',
    styleSkill: 'clean editorial, light',
    insertMode: 'replace',
    planPageContent: async (a) => {
      asked.push({ ...a })
      return {
        ok: true,
        content: {
          title: `content for ${a.title}`,
          bullets: ['one', 'two'],
          cards: [
            { heading: 'A', body: 'a' },
            { heading: 'B', body: 'b' },
            { heading: 'C', body: 'c' },
          ],
        },
      }
    },
    renderLocalPage: async ({ page, first, replaceExisting }) => {
      landed.push({ page, first, replaceExisting })
      return { ok: true }
    },
    ...over,
  }
  return { args, landed, asked }
}

const textOf = (page: RenderedPage) =>
  page.elements
    .flatMap((e) => (e.kind === 'text' ? e.paragraphs : []))
    .flatMap((p) => p.runs.map((r) => r.text))
    .join(' ')

describe('generateDeckLocally', () => {
  it('lands every planned page, in order', async () => {
    const { args, landed } = harness()
    const result = await generateDeckLocally(args)

    expect(result.landed).toBe(3)
    expect(result.doneFlags).toEqual([true, true, true])
    expect(landed.map((l) => textOf(l.page).includes('content for Page 1'))).toEqual([
      true,
      false,
      false,
    ])
    // only the first page carries the instruction to clear the old deck
    expect(landed.map((l) => l.first)).toEqual([true, false, false])
    expect(landed.every((l) => l.replaceExisting)).toBe(true)
  })

  it('asks for the slots the normalized layout actually has', async () => {
    const { args, asked } = harness()
    await generateDeckLocally(args)
    expect(asked.map((a) => a.layout)).toEqual(['cover', 'cards', 'cards'])
    expect(asked.map((a) => a.pageIndex)).toEqual([1, 2, 3])
    expect(asked.every((a) => a.totalPages === 3)).toBe(true)
  })

  it('tells the content step when a page already has an image', async () => {
    const pages = plan(2)
    pages[1]!.layout = 'left_text_right_image'
    pages[1]!.image_queries = ['https://example.com/a.jpg']
    const { args, asked, landed } = harness({ pages })
    await generateDeckLocally(args)

    expect(asked.map((a) => a.hasImage)).toEqual([false, true])
    expect(landed[1]!.page.elements.some((e) => e.kind === 'image')).toBe(true)
  })

  it('accepts a locally generated image marker, not just http urls', async () => {
    const pages = plan(1)
    pages[0]!.type = 'content'
    pages[0]!.layout = 'left_text_right_image'
    pages[0]!.image_queries = ['genimg:/tmp/x.png']
    const { args, landed } = harness({ pages })
    await generateDeckLocally(args)
    expect(
      landed[0]!.page.elements.some((e) => e.kind === 'image' && e.url === 'genimg:/tmp/x.png'),
    ).toBe(true)
  })

  it('does not waste a searched image on a layout with nowhere to put it', async () => {
    // generic layout + an image → the image gets a slot
    const generic = plan(1)
    generic[0]!.type = 'content'
    generic[0]!.layout = 'unknown_variant'
    generic[0]!.image_queries = ['https://example.com/a.jpg']
    const withSlot = harness({ pages: generic })
    await generateDeckLocally(withSlot.args)
    expect(withSlot.landed[0]!.page.elements.some((e) => e.kind === 'image')).toBe(true)

    // a structural layout keeps its structure; the photo is the thing that gives
    const cards = plan(1)
    cards[0]!.type = 'content'
    cards[0]!.layout = 'three_column_cards'
    cards[0]!.image_queries = ['https://example.com/a.jpg']
    const structural = harness({ pages: cards })
    await generateDeckLocally(structural.args)
    expect(structural.landed[0]!.page.elements.some((e) => e.kind === 'image')).toBe(false)
  })

  it('retries the content call once, then lands the page with its title rather than skipping it', async () => {
    const planPageContent = vi
      .fn<LocalDeckArgs['planPageContent']>()
      .mockResolvedValueOnce({ ok: false, error: 'bad json' })
      .mockResolvedValueOnce({ ok: true, content: { title: 'recovered' } })
      .mockResolvedValue({ ok: false, error: 'model is down' })
    const { args, landed } = harness({ pages: plan(2), planPageContent })
    const result = await generateDeckLocally(args)

    expect(planPageContent).toHaveBeenCalledTimes(4) // two attempts per page
    expect(result.landed).toBe(2)
    expect(textOf(landed[0]!.page)).toContain('recovered')
    // page 2 exhausted both attempts: it is on the canvas, but reported as unfinished
    expect(textOf(landed[1]!.page)).toContain('Page 2')
    expect(result.errors[0]).toBeUndefined()
    expect(result.errors[1]).toContain('model is down')
    expect(result.errors[1]).toContain('title only')
  })

  it('keeps going when one page cannot be drawn, and reports it', async () => {
    const renderLocalPage = vi
      .fn<LocalDeckArgs['renderLocalPage']>()
      .mockResolvedValueOnce({ ok: false, error: 'canvas said no' })
      .mockResolvedValue({ ok: true })
    const { args } = harness({ renderLocalPage })
    const result = await generateDeckLocally(args)

    expect(result.landed).toBe(2)
    expect(result.doneFlags).toEqual([false, true, true])
    expect(result.errors[0]).toBe('canvas said no')
    // the deck still has to be replaced — by the first page that did land
    expect(renderLocalPage.mock.calls.map((c) => c[0]!.first)).toEqual([true, true, false])
  })

  it('stops between pages when the user stops, keeping what landed', async () => {
    const controller = new AbortController()
    const drawn: RenderedPage[] = []
    const { args } = harness({
      pages: plan(4),
      signal: controller.signal,
      renderLocalPage: async ({ page }) => {
        drawn.push(page)
        controller.abort()
        return { ok: true }
      },
    })
    const landed = drawn
    const result = await generateDeckLocally(args)

    expect(result.cancelled).toBe(true)
    expect(landed).toHaveLength(1)
    expect(result.landed).toBe(1)
  })

  it('follows the style hint into the theme', async () => {
    const light = harness({ pages: plan(1) })
    await generateDeckLocally(light.args)
    const dark = harness({ pages: plan(1), styleSkill: 'dark, high contrast, minimal' })
    await generateDeckLocally(dark.args)

    expect(light.landed[0]!.page.background).not.toBe(dark.landed[0]!.page.background)
    for (const l of [...light.landed, ...dark.landed]) {
      for (const el of l.page.elements) {
        expect(el.x + el.w).toBeLessThanOrEqual(CANVAS.w)
        expect(el.y + el.h).toBeLessThanOrEqual(CANVAS.h)
      }
    }
  })
})

describe('pageContentFrom', () => {
  it('keeps the good fields of a partly malformed answer', () => {
    const content = pageContentFrom(
      {
        title: '  Real title  ',
        bullets: ['keep', '', 42, '  trim  '],
        cards: [{ heading: 'A', body: 'a' }, 'nonsense', { heading: '', body: '' }],
        kpis: [{ value: '18%', label: 'growth' }],
        figure: { value: '', caption: 'orphan caption' },
        source: 'Source: filings',
      },
      'fallback',
    )
    expect(content.title).toBe('Real title')
    expect(content.bullets).toEqual(['keep', 'trim'])
    expect(content.cards).toEqual([{ heading: 'A', body: 'a' }])
    expect(content.kpis).toEqual([{ value: '18%', label: 'growth' }])
    // a figure with no value is not a figure; the layout would render "—"
    expect(content.figure).toBeUndefined()
    expect(content.source).toBe('Source: filings')
  })

  it('falls back to the planned title and caps runaway lists', () => {
    const content = pageContentFrom(
      {
        bullets: Array.from({ length: 12 }, (_, i) => `b${i}`),
        cards: Array.from({ length: 8 }, (_, i) => ({ heading: `h${i}`, body: 'x' })),
        kpis: Array.from({ length: 8 }, (_, i) => ({ value: `${i}`, label: 'l' })),
      },
      'planned title',
    )
    expect(content.title).toBe('planned title')
    expect(content.bullets).toHaveLength(6)
    expect(content.cards).toHaveLength(3)
    expect(content.kpis).toHaveLength(4)
  })

  it('survives junk instead of an object', () => {
    expect(pageContentFrom(null, 'title').title).toBe('title')
    expect(pageContentFrom('a string', 'title').bullets).toBeUndefined()
    expect(pageContentFrom({}, 'title', 'https://x/y.png').imageUrl).toBe('https://x/y.png')
  })
})
