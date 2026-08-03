/**
 * Local deck generation: the page loop that runs when there is no cloud
 * service. What matters here is what the user ends up with — every planned
 * page present, in order, with the deck it replaced gone — and that one bad
 * page does not take the deck down with it.
 */
import { describe, expect, it, vi } from 'vitest'
import {
  chartFrom,
  generateDeckLocally,
  pageContentFrom,
  tightenContent,
  type LocalDeckArgs,
} from '../src/renderer/ai/deck-local'
import { CANVAS, type PageContent, type RenderedPage } from '../src/renderer/ai/deck-layout'

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

/**
 * The layout engine sizes boxes from an estimate of how wide text will be; the
 * renderer knows. Every fitting bug in this feature came from that gap, so a
 * landed page is measured and rebuilt with less content when the measurement
 * disagrees.
 */
describe('render → measure → tighten', () => {
  const wordy = {
    title: 'A page whose content the renderer will say does not fit in the space available',
    bullets: [
      'A bullet long enough that the estimate and the renderer are likely to disagree about it',
      'A second bullet of similar length, also pushing against the bottom of its box',
      'A third for good measure, because three is what the layout expects to receive',
      'A fourth that only survives the first tightening pass',
    ],
    cards: [
      { heading: 'A heading longer than the slot', body: 'A body sentence of a fair length.' },
    ],
  }

  it('rebuilds the page in place until the renderer stops complaining', async () => {
    const rendered: Array<{ replaceIndex?: number; chars: number }> = []
    let round = 0
    const bulletPage = plan(1)
    bulletPage[0]!.type = 'content'
    bulletPage[0]!.layout = 'bullets'
    const { args } = harness({
      pages: bulletPage,
      planPageContent: async () => ({ ok: true, content: wordy }),
      renderLocalPage: async ({ page, replaceIndex }) => {
        const chars = page.elements
          .flatMap((e) => (e.kind === 'text' ? e.paragraphs : []))
          .flatMap((p) => p.runs.map((r) => r.text.length))
          .reduce((a, b) => a + b, 0)
        rendered.push({ ...(replaceIndex !== undefined ? { replaceIndex } : {}), chars })
        // the renderer reports overflow on the first attempt only
        return { ok: true, slideIndex: 3, issues: round++ === 0 ? ['Text overflow: 40px'] : [] }
      },
    })
    const result = await generateDeckLocally(args)

    expect(rendered).toHaveLength(2)
    // the retry rebuilds the page that already landed, rather than appending
    expect(rendered[0]!.replaceIndex).toBeUndefined()
    expect(rendered[1]!.replaceIndex).toBe(3)
    // and it carries less text than the attempt the renderer rejected
    expect(rendered[1]!.chars).toBeLessThan(rendered[0]!.chars)
    expect(result.landed).toBe(1)
    expect(result.errors[0]).toBeUndefined()
  })

  it('gives up after two rebuilds and says the page is still imperfect', async () => {
    let calls = 0
    const { args } = harness({
      pages: plan(1),
      planPageContent: async () => ({ ok: true, content: wordy }),
      renderLocalPage: async () => {
        calls++
        return { ok: true, slideIndex: 0, issues: ['Overlap: title and body intersect'] }
      },
    })
    const result = await generateDeckLocally(args)

    expect(calls).toBe(3) // first attempt + MAX_TIGHTEN rebuilds
    expect(result.landed).toBe(1) // the page stays: imperfect beats absent
    expect(result.errors[0]).toContain('still imperfect')
    expect(result.errors[0]).toContain('Overlap')
  })

  it('does not rebuild a page the renderer is happy with', async () => {
    let calls = 0
    const { args } = harness({
      pages: plan(2),
      renderLocalPage: async () => {
        calls++
        return { ok: true, slideIndex: calls - 1, issues: [] }
      },
    })
    await generateDeckLocally(args)
    expect(calls).toBe(2)
  })

  it('treats a bridge that reports nothing as nothing to fix', async () => {
    // older/other implementations may not return issues at all
    let calls = 0
    const { args } = harness({
      pages: plan(1),
      renderLocalPage: async () => {
        calls++
        return { ok: true }
      },
    })
    const result = await generateDeckLocally(args)
    expect(calls).toBe(1)
    expect(result.landed).toBe(1)
  })
})

describe('tightenContent', () => {
  const full: PageContent = {
    title: 'A title of moderate length that is still perfectly reasonable for one slide',
    subtitle: 'A supporting line that says a little more about what the page is claiming here',
    bullets: ['one two three four five six seven eight nine ten', 'b', 'c', 'd', 'e'],
    cards: [{ heading: 'A heading that is rather long for a card', body: 'x'.repeat(200) }],
    figure: { value: '42%', caption: 'y'.repeat(120) },
    source: 'Source: a long provenance line naming several documents at once, in detail',
  }

  it('changes nothing at level zero', () => {
    expect(tightenContent(full, 0)).toBe(full)
  })

  it('sheds text before it sheds items, and marks what it cut', () => {
    const t1 = tightenContent(full, 1)
    expect(t1.bullets).toHaveLength(4)
    expect(t1.cards![0]!.body.length).toBeLessThan(full.cards![0]!.body.length)
    expect(t1.cards![0]!.body.endsWith('…')).toBe(true)
    // provenance survives the first pass — it is evidence, not filler
    expect(t1.source).toBe(full.source)

    const t2 = tightenContent(full, 2)
    expect(t2.bullets).toHaveLength(3)
    expect(t2.cards![0]!.body.length).toBeLessThan(t1.cards![0]!.body.length)
    expect(t2.figure!.caption.length).toBeLessThan(t1.figure!.caption.length)
  })

  it('cuts on a word boundary rather than mid-word', () => {
    const long =
      'alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma'
    const t = tightenContent({ ...full, bullets: [long] }, 2)
    const line = t.bullets![0]!
    expect(line.endsWith('…')).toBe(true)
    // the cut lands between words: the last token is a whole one from the input
    const last = line.replace('…', '').trim().split(' ').pop()!
    expect(long.split(' ')).toContain(last)
  })

  it('leaves already-short content alone', () => {
    const short: PageContent = { title: 'Short', bullets: ['a', 'b'] }
    expect(tightenContent(short, 2).title).toBe('Short')
    expect(tightenContent(short, 2).bullets).toEqual(['a', 'b'])
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

/**
 * The KPI value slot is 52pt and holds a figure. A real generation filled it
 * with metric names — "Acknowledged <15 min" — which pass a digit check and
 * render as huge clipped stubs.
 */
describe('KPI values are figures, not metric names', () => {
  const kpiPage = (kpis: Array<{ value: string; label: string }>) =>
    pageContentFrom({ title: 'Metrics', kpis }, 'fallback')

  it('accepts the shapes a figure actually takes', () => {
    const c = kpiPage([
      { value: '18%', label: 'growth' },
      { value: '$4.2M', label: 'pipeline' },
      { value: '2.1 days', label: 'lead time' },
      { value: '99.95%', label: 'uptime' },
    ])
    expect(c.kpis).toHaveLength(4)
    expect(c.cards).toBeUndefined()
  })

  it('rejects a metric name wearing a number, and keeps it as a card', () => {
    const c = kpiPage([
      { value: 'Acknowledged <15 min', label: 'Response performance' },
      { value: 'Pages/shift P90', label: 'Alert load and toil' },
      { value: 'Error-budget burn 20%', label: 'Reliability' },
    ])
    expect(c.kpis).toBeUndefined()
    // the content survives in a slot that can hold a phrase
    expect(c.cards).toEqual([
      { heading: 'Acknowledged <15 min', body: 'Response performance' },
      { heading: 'Pages/shift P90', body: 'Alert load and toil' },
      { heading: 'Error-budget burn 20%', body: 'Reliability' },
    ])
  })

  it('does not overwrite real cards with salvaged ones', () => {
    const c = pageContentFrom(
      {
        title: 'Both',
        cards: [{ heading: 'Real', body: 'card' }],
        kpis: [
          { value: 'Not a figure at all', label: 'x' },
          { value: 'Also not one here', label: 'y' },
        ],
      },
      'fallback',
    )
    expect(c.cards).toEqual([{ heading: 'Real', body: 'card' }])
  })

  it('still drops icon-style values that carry no number', () => {
    const c = kpiPage([
      { value: '◯', label: 'Deep work' },
      { value: '◇', label: 'Decisions' },
    ])
    expect(c.kpis).toBeUndefined()
  })
})

/**
 * Charts on generated pages. The engine has always built native pptx charts;
 * the local generator emitted only rectangles, text and images, so a data page
 * came out as pipe-joined figures in a card.
 */
describe('chartFrom', () => {
  const good = {
    kind: 'bar',
    title: 'Incidents by quarter',
    categories: ['Q1', 'Q2', 'Q3'],
    series: [{ name: 'Incidents', values: [18, 12, 9] }],
    figures: 'document',
  }

  it('accepts a consistent chart', () => {
    const c = chartFrom(good, true)
    expect(c?.kind).toBe('bar')
    expect(c?.series[0]!.values).toEqual([18, 12, 9])
  })

  it('rejects a series that does not line up with the categories', () => {
    // a short series silently shifts the axis rather than failing
    expect(chartFrom({ ...good, series: [{ name: 'x', values: [1, 2] }] }, true)).toBeUndefined()
    expect(
      chartFrom({ ...good, series: [{ name: 'x', values: [1, null, 3] }] }, true),
    ).toBeUndefined()
  })

  it('rejects a chart with nothing to plot', () => {
    expect(chartFrom({ ...good, categories: ['only'] }, true)).toBeUndefined()
    expect(chartFrom({ ...good, series: [] }, true)).toBeUndefined()
    expect(chartFrom({ ...good, kind: 'sankey' }, true)).toBeUndefined()
    expect(chartFrom(null, true)).toBeUndefined()
  })

  it('downgrades a provenance claim the deck cannot support', () => {
    // no reference material: "document" is not something the model can know
    expect(chartFrom(good, false)?.figures).toBe('sample')
    expect(chartFrom({ ...good, figures: 'search' }, false)?.figures).toBe('sample')
    // with material, the claim stands
    expect(chartFrom(good, true)?.figures).toBe('document')
    // an unrecognised claim is treated as illustrative
    expect(chartFrom({ ...good, figures: 'trust me' }, true)?.figures).toBe('sample')
  })

  it('keeps at most four series and twelve categories', () => {
    const many = {
      ...good,
      categories: Array.from({ length: 20 }, (_, i) => `c${i}`),
      series: Array.from({ length: 8 }, (_, i) => ({
        name: `s${i}`,
        values: Array.from({ length: 20 }, () => 1),
      })),
    }
    const c = chartFrom(many, true)!
    expect(c.categories).toHaveLength(12)
    expect(c.series).toHaveLength(4)
    expect(c.series[0]!.values).toHaveLength(12)
  })
})

describe('generateDeckLocally with a chart page', () => {
  it('emits a chart element and reports illustrative figures', async () => {
    const pages = plan(1)
    pages[0]!.type = 'data'
    pages[0]!.layout = 'chart_with_insight'
    const { args, landed } = harness({
      pages,
      planPageContent: async () => ({
        ok: true,
        content: {
          title: 'Incidents fell through the year',
          bullets: ['Down 50% since Q1'],
          chart: {
            kind: 'line',
            categories: ['Q1', 'Q2', 'Q3'],
            series: [{ name: 'Incidents', values: [18, 12, 9] }],
            figures: 'document',
          },
        },
      }),
    })
    const result = await generateDeckLocally(args)

    const chart = landed[0]!.page.elements.find((e) => e.kind === 'chart')
    expect(chart).toBeDefined()
    if (chart?.kind !== 'chart') throw new Error('expected a chart element')
    expect(chart.chart.kind).toBe('line')
    // no context was supplied, so the "document" claim is downgraded and reported
    expect(chart.chart.figures).toBe('sample')
    expect(result.illustrative).toEqual([1])
  })

  it('falls back when the model gives a chart page no usable data', async () => {
    const pages = plan(1)
    pages[0]!.type = 'data'
    pages[0]!.layout = 'chart_with_insight'
    const { args, landed } = harness({
      pages,
      planPageContent: async () => ({
        ok: true,
        content: { title: 'No data', bullets: ['a', 'b'], chart: { kind: 'bar' } },
      }),
    })
    const result = await generateDeckLocally(args)
    expect(landed[0]!.page.elements.some((e) => e.kind === 'chart')).toBe(false)
    expect(result.illustrative).toEqual([])
    // the page still lands, carrying its bullets
    expect(result.landed).toBe(1)
  })
})
