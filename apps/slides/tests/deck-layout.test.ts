import { describe, expect, it } from 'vitest'
import {
  CANVAS,
  DARK_THEME,
  estimateTextHeight,
  LIGHT_THEME,
  layoutPage,
  normalizeLayout,
  themeFor,
  type ElementSpec,
  type LayoutId,
  type PageContent,
} from '../src/renderer/ai/deck-layout'

/**
 * These are the failures that make a generated deck look generated: text off
 * the page, boxes on top of each other, invisible text. All three are
 * checkable without rendering, so they are checked for every layout rather
 * than spotted one deck at a time.
 */

const CONTENT: PageContent = {
  title: 'Q3 delivered momentum, but Q4 needs focus',
  subtitle: 'Board review — November',
  bullets: [
    'Revenue grew 18% quarter over quarter',
    'Churn fell to 2.1%, the lowest in six quarters',
    'Two enterprise deals slipped into Q4',
  ],
  cards: [
    { heading: 'Growth', body: 'New logos up 24%, driven by the self-serve funnel.' },
    { heading: 'Risk', body: 'Two deals slipped; both are procurement-blocked, not lost.' },
    { heading: 'Focus', body: 'Ship onboarding v2 before the January renewal window.' },
  ],
  kpis: [
    { value: '18%', label: 'QoQ revenue growth' },
    { value: '2.1%', label: 'Monthly churn' },
    { value: '$4.2M', label: 'Pipeline added' },
  ],
  figure: { value: '18%', caption: 'quarter-over-quarter revenue growth' },
  source: 'Source: internal finance reporting, October 2026',
  imageUrl: 'https://example.com/photo.jpg',
}

const ALL: LayoutId[] = [
  'cover',
  'bullets',
  'image_right',
  'cards',
  'big_number',
  'kpis',
  'comparison',
  'closing',
]

const boxes = (elements: ElementSpec[]) => elements.filter((e) => e.kind !== 'rect')

function overlaps(a: ElementSpec, b: ElementSpec): boolean {
  return a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h
}

describe('layoutPage geometry', () => {
  it.each(ALL)('keeps every element inside the canvas: %s', (layout) => {
    for (const theme of [LIGHT_THEME, DARK_THEME]) {
      for (const el of layoutPage(layout, CONTENT, theme).elements) {
        expect(el.x).toBeGreaterThanOrEqual(0)
        expect(el.y).toBeGreaterThanOrEqual(0)
        expect(el.x + el.w).toBeLessThanOrEqual(CANVAS.w)
        expect(el.y + el.h).toBeLessThanOrEqual(CANVAS.h)
        expect(el.w).toBeGreaterThan(0)
        expect(el.h).toBeGreaterThan(0)
      }
    }
  })

  it.each(ALL)('never lets two text or image boxes overlap: %s', (layout) => {
    const content = boxes(layoutPage(layout, CONTENT, LIGHT_THEME).elements)
    for (let i = 0; i < content.length; i++) {
      for (let j = i + 1; j < content.length; j++) {
        expect(overlaps(content[i]!, content[j]!), `${layout}: element ${i} overlaps ${j}`).toBe(
          false,
        )
      }
    }
  })

  it.each(ALL)('puts the page title somewhere on the page: %s', (layout) => {
    const text = layoutPage(layout, CONTENT, LIGHT_THEME)
      .elements.flatMap((e) => (e.kind === 'text' ? e.paragraphs : []))
      .flatMap((p) => p.runs.map((r) => r.text))
      .join(' ')
    expect(text).toContain(CONTENT.title)
  })
})

/**
 * A slide text box does not clip: overrunning text draws over what is below it
 * and off the bottom of the page. The first generated deck ended with a page
 * whose five-step procedure arrived in a one-line slot and spilled off the
 * canvas, so overlong content is checked for every layout, not just short.
 */
describe('layoutPage with more content than the slot expects', () => {
  const FLOOD: PageContent = {
    title:
      'Run the two-week writing-first experiment across every engineering team before the next planning cycle begins',
    subtitle:
      '1. Name one owner and one shared place for proposals and decisions 2. Replace recurring status and routine decision meetings with written updates 3. Use one template: context, decision needed, options, recommendation, owner, deadline 4. Allow at least one full working day for async review 5. Review calendar fragmentation, reopened questions and decision clarity',
    bullets: [
      'Fragmentation — deep work is split into unusable intervals, and the recovery cost after each interruption is far larger than the meeting itself',
      'Forced simultaneity — presence is treated as participation, which quietly excludes anyone outside the organiser’s time zone',
      'Decisions evaporate — what was agreed lives in someone’s memory rather than in a document anyone can find later',
    ],
    cards: [
      {
        heading: 'A considerably longer heading than the slot was designed for',
        body: 'Two-by-two grid on #EEE8DC: horizontal axis lower to higher ambiguity, vertical axis lower to higher urgency, with each quadrant carrying a worked example and a recommended medium for that combination of the two axes.',
      },
      {
        heading: 'Second column',
        body: 'Another long body that keeps going well past the point where a card of this size could reasonably hold it, continuing for several more clauses to be sure.',
      },
      {
        heading: 'Third column',
        body: 'A third body of similar length, so the cards layout is compared at the same column width as the short-content case rather than a wider two-column one.',
      },
    ],
    kpis: [{ value: '18%', label: 'A label considerably longer than five words was requested' }],
    figure: {
      value: '42%',
      caption:
        'A caption that runs on well past twelve words, describing the figure in far more detail than the space allows',
    },
    source:
      'Source: a provenance line that is itself rather long, naming several documents at once',
  }

  it.each(ALL)('keeps flooded content inside the canvas: %s', (layout) => {
    for (const theme of [LIGHT_THEME, DARK_THEME]) {
      for (const el of layoutPage(layout, FLOOD, theme).elements) {
        expect(el.x).toBeGreaterThanOrEqual(0)
        expect(el.y).toBeGreaterThanOrEqual(0)
        expect(el.x + el.w).toBeLessThanOrEqual(CANVAS.w)
        expect(el.y + el.h).toBeLessThanOrEqual(CANVAS.h)
      }
    }
  })

  it.each(ALL)('keeps flooded text inside its own box: %s', (layout) => {
    for (const el of layoutPage(layout, FLOOD, LIGHT_THEME).elements) {
      if (el.kind !== 'text') continue
      const lines = el.paragraphs.map((p) => p.runs.map((r) => r.text).join(''))
      const size = el.paragraphs[0]?.runs[0]?.fontSize ?? 12
      const spacing = el.paragraphs[0]?.lineSpacingPct ? el.paragraphs[0].lineSpacingPct / 100 : 1
      // the same estimate the layout fits against, so this asserts the fit ran
      expect(estimateTextHeight(lines, size, el.w) * spacing).toBeLessThanOrEqual(el.h + 1)
    }
  })

  it.each(ALL)('still shows no overlap when content overflows: %s', (layout) => {
    const content = boxes(layoutPage(layout, FLOOD, LIGHT_THEME).elements)
    for (let i = 0; i < content.length; i++) {
      for (let j = i + 1; j < content.length; j++) {
        expect(overlaps(content[i]!, content[j]!), `${layout}: ${i} overlaps ${j}`).toBe(false)
      }
    }
  })

  it('shrinks within the ladder before giving up any content', () => {
    const titleSize = (content: PageContent) => {
      const el = layoutPage('bullets', content, LIGHT_THEME).elements.find(
        (e) =>
          e.kind === 'text' &&
          e.paragraphs[0]?.runs[0]?.text.startsWith(content.title.slice(0, 12)),
      )
      return el?.kind === 'text' ? (el.paragraphs[0]?.runs[0]?.fontSize ?? 0) : 0
    }
    expect(titleSize(CONTENT)).toBe(36)
    // the long title uses the ladder rather than wrapping out of its box
    expect(titleSize(FLOOD)).toBeLessThan(36)
  })

  it('clips at the type floor instead of shrinking under it', () => {
    const body = layoutPage('cards', FLOOD, LIGHT_THEME).elements.find(
      (e) => e.kind === 'text' && e.paragraphs.some((p) => p.runs[0]?.text.includes('Two-by-two')),
    )
    expect(body?.kind).toBe('text')
    if (body?.kind !== 'text') return
    const run = body.paragraphs[0]!.runs[0]!
    // body text stays at the floor…
    expect(run.fontSize).toBe(18)
    // …and the overflow is taken out of the text, visibly
    expect(run.text.endsWith('…')).toBe(true)
    expect(run.text.length).toBeLessThan(FLOOD.cards![0]!.body.length)
  })
})

describe('layoutPage content handling', () => {
  it('falls back to bullets when the layout has no content to fill it', () => {
    const bare: PageContent = { title: 'Only a title', bullets: ['one point'] }
    for (const layout of ['cards', 'kpis', 'comparison', 'big_number'] as LayoutId[]) {
      const rendered = layoutPage(layout, bare, LIGHT_THEME)
      const text = rendered.elements
        .flatMap((e) => (e.kind === 'text' ? e.paragraphs : []))
        .flatMap((p) => p.runs.map((r) => r.text))
      expect(text).toContain('one point')
    }
  })

  it('survives a page with nothing but a title', () => {
    for (const layout of ALL) {
      const rendered = layoutPage(layout, { title: 'Bare' }, LIGHT_THEME)
      expect(rendered.elements.length).toBeGreaterThan(0)
      for (const el of rendered.elements) expect(el.x + el.w).toBeLessThanOrEqual(CANVAS.w)
    }
  })

  it('caps runaway lists so a page stays a page', () => {
    const many: PageContent = {
      title: 'Too much',
      bullets: Array.from({ length: 20 }, (_, i) => `point ${i}`),
      cards: Array.from({ length: 9 }, (_, i) => ({ heading: `c${i}`, body: 'x' })),
      kpis: Array.from({ length: 9 }, (_, i) => ({ value: `${i}`, label: 'metric' })),
    }
    const paras = (layout: LayoutId) =>
      layoutPage(layout, many, LIGHT_THEME).elements.flatMap((e) =>
        e.kind === 'text' ? e.paragraphs : [],
      )
    expect(paras('bullets').filter((p) => p.bullet === 'char')).toHaveLength(6)
    // three cards: heading + body each, plus the title block
    expect(paras('cards')).toHaveLength(1 + 3 * 2)
    expect(paras('kpis')).toHaveLength(1 + 4 * 2)
  })

  it('shows a placeholder panel when an image page has no image', () => {
    const withImage = layoutPage('image_right', CONTENT, LIGHT_THEME)
    expect(withImage.elements.some((e) => e.kind === 'image')).toBe(true)

    const without = layoutPage('image_right', { ...CONTENT, imageUrl: undefined }, LIGHT_THEME)
    expect(without.elements.some((e) => e.kind === 'image')).toBe(false)
    // the right half is filled rather than left as a hole
    expect(without.elements.filter((e) => e.kind === 'rect' && e.w > 400)).toHaveLength(1)
  })

  it('carries a provenance line onto pages that have one, and only those', () => {
    const has = (c: PageContent) =>
      layoutPage('bullets', c, LIGHT_THEME)
        .elements.flatMap((e) => (e.kind === 'text' ? e.paragraphs : []))
        .some((p) => p.runs.some((r) => r.text.startsWith('Source:')))
    expect(has(CONTENT)).toBe(true)
    expect(has({ ...CONTENT, source: undefined })).toBe(false)
  })
})

/**
 * Delivery floor, from the OfficeCLI pptx skill (Apache-2.0, iOfficeAI): a deck
 * is read from across a room, so a title carries at 36pt and body at 18pt, and
 * text on a dark fill has to clear a real brightness gap rather than merely
 * differ from it. Captions, sources and short KPI sublabels are the exceptions
 * that skill names.
 */
describe('typographic and contrast floor', () => {
  const CAPTION_SLOTS = [13, 14, 15, 12, 11]
  const textRuns = (content: PageContent, layout: LayoutId, theme = LIGHT_THEME) =>
    layoutPage(layout, content, theme)
      .elements.filter((e) => e.kind === 'text')
      .flatMap((e) => (e.kind === 'text' ? e.paragraphs : []))
      .flatMap((p) => p.runs)

  it.each(ALL)('never sets body text below the readable floor: %s', (layout) => {
    for (const run of textRuns(CONTENT, layout)) {
      const size = run.fontSize ?? 0
      // captions/sources/sublabels are allowed under 18; nothing else is
      if (CAPTION_SLOTS.includes(size)) continue
      expect(size, `${layout}: "${run.text.slice(0, 30)}" at ${size}pt`).toBeGreaterThanOrEqual(18)
    }
  })

  it('gives every page title the weight of a title', () => {
    for (const layout of ALL) {
      const title = textRuns(CONTENT, layout).find((r) => r.text === CONTENT.title)
      expect(title?.fontSize ?? 0, layout).toBeGreaterThanOrEqual(36)
      expect(title?.bold, layout).toBe(true)
    }
  })

  it('clears a real brightness gap between text and the surface behind it', () => {
    // (R×299 + G×587 + B×114) / 1000
    const brightness = (hex: string) => {
      const n = parseInt(hex.replace('#', ''), 16)
      return (((n >> 16) & 255) * 299 + ((n >> 8) & 255) * 587 + (n & 255) * 114) / 1000
    }
    for (const theme of [LIGHT_THEME, DARK_THEME]) {
      for (const layout of ALL) {
        const page = layoutPage(layout, CONTENT, theme)
        for (const el of page.elements) {
          if (el.kind !== 'text') continue
          const behind = page.elements
            .filter(
              (r) =>
                r.kind === 'rect' &&
                r.x <= el.x &&
                r.y <= el.y &&
                r.x + r.w >= el.x + el.w &&
                r.y + r.h >= el.y + el.h,
            )
            .pop()
          const surface = behind?.kind === 'rect' ? behind.fill : page.background
          for (const p of el.paragraphs) {
            for (const r of p.runs) {
              const gap = Math.abs(brightness(r.color ?? '#000000') - brightness(surface))
              expect(gap, `${layout}: ${r.color} on ${surface}`).toBeGreaterThan(90)
            }
          }
        }
      }
    }
  })
})

describe('theme and layout selection', () => {
  it('never puts text in the colour of whatever is behind it', () => {
    for (const theme of [LIGHT_THEME, DARK_THEME]) {
      for (const layout of ALL) {
        const rendered = layoutPage(layout, CONTENT, theme)
        for (const el of rendered.elements) {
          if (el.kind !== 'text') continue
          // the surface this text actually sits on: the last rect that covers
          // its origin (later elements paint over earlier ones), else the page
          const behind = rendered.elements
            .filter(
              (r) =>
                r.kind === 'rect' &&
                r.x <= el.x &&
                r.y <= el.y &&
                r.x + r.w >= el.x + el.w &&
                r.y + r.h >= el.y + el.h,
            )
            .pop()
          const surface = behind?.kind === 'rect' ? behind.fill : rendered.background
          for (const p of el.paragraphs) {
            for (const r of p.runs) {
              expect(r.color, `${layout}: text on ${surface}`).not.toBe(surface)
            }
          }
        }
      }
    }
  })

  it('maps the planner vocabulary onto a layout it can draw', () => {
    expect(normalizeLayout('cover_full_image_overlay')).toBe('cover')
    expect(normalizeLayout('cover_dark_minimal')).toBe('cover')
    expect(normalizeLayout('three_column_cards')).toBe('cards')
    expect(normalizeLayout('timeline_horizontal')).toBe('cards')
    expect(normalizeLayout('hero_big_number')).toBe('big_number')
    expect(normalizeLayout('kpi_cards_row')).toBe('kpis')
    expect(normalizeLayout('two_column_comparison')).toBe('comparison')
    expect(normalizeLayout('two_by_two_grid')).toBe('comparison')
    expect(normalizeLayout('left_text_right_image')).toBe('image_right')
    expect(normalizeLayout('full_image_text_overlay')).toBe('image_right')
    expect(normalizeLayout('closing_cta')).toBe('closing')
    // an unknown variant still renders rather than failing the page
    expect(normalizeLayout('some_future_variant')).toBe('bullets')
    // the page type decides when the layout name says nothing
    expect(normalizeLayout('', 'cover')).toBe('cover')
    expect(normalizeLayout(undefined, 'closing')).toBe('closing')
  })

  it('picks the dark theme only when the style asks for it', () => {
    expect(themeFor('dark, minimal, high contrast')).toBe(DARK_THEME)
    expect(themeFor('深色背景，极简')).toBe(DARK_THEME)
    expect(themeFor('clean editorial, plenty of white space')).toBe(LIGHT_THEME)
    expect(themeFor(undefined)).toBe(LIGHT_THEME)
  })
})

/**
 * A generated deck ended a comparison column mid-clause — "…assigned within 10
 * minutes: 83% (15 of 18); …" — which reads as damage rather than editing.
 */
describe('clipping ends somewhere a reader can stop', () => {
  const bodyOf = (text: string) => {
    const page = layoutPage(
      'comparison',
      {
        title: 'Two ways to look at it',
        cards: [
          { heading: 'Left', body: text },
          { heading: 'Right', body: 'short' },
        ],
      },
      LIGHT_THEME,
    )
    return page.elements
      .filter((e) => e.kind === 'text')
      .flatMap((e) => (e.kind === 'text' ? e.paragraphs : []))
      .map((p) => p.runs[0]!.text)
      .find((t) => t.startsWith(text.slice(0, 10)))!
  }

  it('stops at the last full sentence when there is one', () => {
    const long =
      'Median MTTA is four minutes. An incident commander is assigned within ten minutes in 83% of cases. ' +
      'Corrective actions close on time in 66% of cases, which is the number the team is trying to move this quarter. ' +
      'The remainder slip into the following month and are re-triaged at the weekly review, usually without much ceremony. ' +
      'That backlog is the single largest contributor to repeat incidents across the last three quarters of data.'
    const out = bodyOf(long)
    expect(out.length).toBeLessThan(long.length)
    expect(out.endsWith('.')).toBe(true)
    expect(out.endsWith('…')).toBe(false)
    // and it is a real prefix of the original, not a mangled one
    expect(long.startsWith(out)).toBe(true)
  })

  it('falls back to a word boundary when no sentence fits', () => {
    const noStops = `${'word '.repeat(200)}end`
    const out = bodyOf(noStops)
    expect(out.endsWith('…')).toBe(true)
    expect(out.replace('…', '').trimEnd().endsWith('word')).toBe(true)
  })

  it('leaves text that fits completely alone', () => {
    const short = 'Two deals slipped; both are procurement-blocked.'
    expect(bodyOf(short)).toBe(short)
  })
})

/**
 * A CJK glyph occupies the full em square — close to twice a Latin character.
 * A single blended width ratio under-counted every Chinese, Japanese and
 * Korean deck by roughly 40%, so boxes were sized for text that could not fit
 * them. The app ships nineteen locales; this is not an edge case.
 */
describe('full-width text', () => {
  const CJK: PageContent = {
    title: '远程工程团队为什么应该用异步写作代替会议',
    subtitle: '面向工程负责人的季度评审材料，二零二六年十一月',
    bullets: [
      '会议把深度工作切成无法使用的碎片，恢复成本远高于会议本身',
      '同步出席被当作参与，把不在同一时区的同事排除在外',
      '口头达成的决定没有记录，几周后没有人能复述当时的理由',
    ],
    cards: [
      { heading: '增长', body: '自助渠道带动新客户增长百分之二十四，续约窗口前还会继续提升。' },
      { heading: '风险', body: '两笔企业订单顺延，都卡在采购流程，并非丢单。' },
      { heading: '重点', body: '在一月续约窗口之前完成新版引导流程。' },
    ],
    kpis: [
      { value: '18%', label: '环比收入增长' },
      { value: '2.1%', label: '月度流失率' },
    ],
    figure: { value: '18%', caption: '环比收入增长，连续六个季度最高' },
    source: '来源：内部财务报表，二零二六年十月',
  }

  it('measures a wide glyph as wider than a latin one', () => {
    // same character count, very different rendered width
    const latin = estimateTextHeight(['aaaaaaaaaaaaaaaaaaaa'], 24, 300)
    const wide = estimateTextHeight(['一二三四五六七八九十一二三四五六七八九十'], 24, 300)
    expect(wide).toBeGreaterThan(latin)
  })

  it.each(ALL)('keeps a CJK page inside the canvas: %s', (layout) => {
    for (const theme of [LIGHT_THEME, DARK_THEME]) {
      for (const el of layoutPage(layout, CJK, theme).elements) {
        expect(el.x).toBeGreaterThanOrEqual(0)
        expect(el.y).toBeGreaterThanOrEqual(0)
        expect(el.x + el.w).toBeLessThanOrEqual(CANVAS.w)
        expect(el.y + el.h).toBeLessThanOrEqual(CANVAS.h)
      }
    }
  })

  it.each(ALL)('keeps CJK text inside its own box: %s', (layout) => {
    for (const el of layoutPage(layout, CJK, LIGHT_THEME).elements) {
      if (el.kind !== 'text') continue
      const lines = el.paragraphs.map((p) => p.runs.map((r) => r.text).join(''))
      const size = el.paragraphs[0]?.runs[0]?.fontSize ?? 12
      const spacing = el.paragraphs[0]?.lineSpacingPct ? el.paragraphs[0].lineSpacingPct / 100 : 1
      expect(estimateTextHeight(lines, size, el.w) * spacing).toBeLessThanOrEqual(el.h + 1)
    }
  })

  it.each(ALL)('does not overlap on a CJK page: %s', (layout) => {
    const content = boxes(layoutPage(layout, CJK, LIGHT_THEME).elements)
    for (let i = 0; i < content.length; i++) {
      for (let j = i + 1; j < content.length; j++) {
        expect(overlaps(content[i]!, content[j]!), `${layout}: ${i} overlaps ${j}`).toBe(false)
      }
    }
  })

  it('treats halfwidth katakana as narrow, not full width', () => {
    // FF61-FF9F is halfwidth despite sitting in the fullwidth forms block
    const half = estimateTextHeight(['ｱｲｳｴｵｱｲｳｴｵｱｲｳｴｵｱｲｳｴｵ'], 24, 300)
    const full = estimateTextHeight(['アイウエオアイウエオアイウエオアイウエオ'], 24, 300)
    expect(half).toBeLessThan(full)
  })
})
