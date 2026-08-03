import { describe, expect, it } from 'vitest'
import {
  CANVAS,
  DARK_THEME,
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
