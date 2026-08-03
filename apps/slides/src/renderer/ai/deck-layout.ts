import type { EditParagraph, EditRun } from '../../shared/ipc'

/**
 * Local deck layout: page plan + content → positioned elements.
 *
 * Cloud generation writes HTML per page and converts it; without a Genspark
 * account that path is closed, so this builds pages out of the native
 * elements the deck engine already has. The split is deliberate:
 *
 * - the **model** supplies content (title, bullets, card text, figures),
 *   which is what it is good at;
 * - this module supplies **geometry and type scale**, which it is not — models
 *   place boxes by guesswork and produce collisions, drift and unreadable
 *   contrast.
 *
 * Everything here is pure: a plan goes in, element specs come out. No IPC, no
 * DOM, no async — so the layouts can be tested for the things that actually go
 * wrong (overflow off-canvas, overlapping boxes, unreadable pairings) instead
 * of being eyeballed one deck at a time.
 */

/** the deck engine's own working canvas, in px at 96dpi (16:9) */
export const CANVAS = { w: 1280, h: 720 } as const

/** page margins and the rhythm everything else is measured against */
const M = { x: 84, top: 76, bottom: 64, gap: 28 } as const

export interface Theme {
  /** page background */
  bg: string
  /** panels and cards sitting on the background */
  surface: string
  /** primary body/heading colour, must read on `bg` */
  text: string
  /** secondary text: labels, captions, sources */
  muted: string
  /** one accent, used for emphasis and the big numbers */
  accent: string
  /** text placed on top of `accent` */
  onAccent: string
  headFont: string
  bodyFont: string
}

/**
 * Two themes rather than a generated palette: a deck's job is to be readable,
 * and every pairing here is checked for contrast once instead of gambling on
 * whatever hex the model invents. `styleHint` only picks between them.
 */
export const LIGHT_THEME: Theme = {
  bg: '#FFFFFF',
  surface: '#F3F5F9',
  text: '#12161C',
  muted: '#5B6472',
  accent: '#1A54C9',
  onAccent: '#FFFFFF',
  headFont: 'Calibri',
  bodyFont: 'Calibri',
}

export const DARK_THEME: Theme = {
  bg: '#12161C',
  surface: '#1E2530',
  text: '#F4F6FA',
  muted: '#9AA6B8',
  accent: '#6EA8FF',
  onAccent: '#0B1017',
  headFont: 'Calibri',
  bodyFont: 'Calibri',
}

/** Pick a theme from the free-text style the planner produced. */
export function themeFor(styleHint: string | undefined): Theme {
  return /\bdark\b|深色|暗色|black background|midnight/i.test(styleHint ?? '')
    ? DARK_THEME
    : LIGHT_THEME
}

/** What this module can draw. The planner's vocabulary is mapped onto it. */
export type LayoutId =
  'cover' | 'bullets' | 'image_right' | 'cards' | 'big_number' | 'kpis' | 'comparison' | 'closing'

/**
 * The planner names 15 variants (six cover treatments, and so on). Rendering
 * all of them separately would be fifteen chances to get spacing wrong; they
 * collapse onto the eight shapes above, which differ in *structure* rather
 * than decoration. An unknown name falls back to bullets, which can render
 * any page that has a title and some text.
 */
export function normalizeLayout(planned: string | undefined, type?: string): LayoutId {
  const l = (planned ?? '').toLowerCase()
  const t = (type ?? '').toLowerCase()
  if (l.startsWith('cover') || t === 'cover') return 'cover'
  if (l.startsWith('closing') || t === 'closing') return 'closing'
  if (l.includes('kpi')) return 'kpis'
  if (l.includes('big_number') || l.includes('hero_big')) return 'big_number'
  if (l.includes('comparison') || l.includes('two_by_two')) return 'comparison'
  if (l.includes('card') || l.includes('three_column') || l.includes('timeline')) return 'cards'
  if (l.includes('image')) return 'image_right'
  return 'bullets'
}

/** Content slots. A layout uses what it needs and tolerates the rest missing. */
export interface PageContent {
  title: string
  subtitle?: string | undefined
  bullets?: string[] | undefined
  cards?: Array<{ heading: string; body: string }> | undefined
  kpis?: Array<{ value: string; label: string }> | undefined
  /** the one figure a big_number page exists for */
  figure?: { value: string; caption: string } | undefined
  /** provenance line for anything with numbers in it */
  source?: string | undefined
  /** resolved image URL (or a generated-image marker) */
  imageUrl?: string | undefined
}

export type ElementSpec =
  | { kind: 'rect'; x: number; y: number; w: number; h: number; fill: string }
  | { kind: 'text'; x: number; y: number; w: number; h: number; paragraphs: EditParagraph[] }
  | { kind: 'image'; x: number; y: number; w: number; h: number; url: string }

export interface RenderedPage {
  background: string
  elements: ElementSpec[]
}

// ── text helpers ────────────────────────────────────────
// Font sizes are points; the geometry below assumes ~1.35 line height, which
// is what the engine renders at these sizes.

interface TextOpts {
  size: number
  color: string
  bold?: boolean
  align?: EditParagraph['align']
  font?: string
  /** rendered as a bulleted paragraph list rather than plain lines */
  bullet?: boolean
  lineSpacingPct?: number
}

function run(text: string, o: TextOpts): EditRun {
  return {
    text,
    fontSize: o.size,
    color: o.color,
    ...(o.bold ? { bold: true } : {}),
    ...(o.font ? { fontFamily: o.font } : {}),
  }
}

function para(text: string, o: TextOpts): EditParagraph {
  return {
    runs: [run(text, o)],
    align: o.align ?? 'left',
    ...(o.bullet ? { bullet: 'char' as const, bulletChar: '•' } : {}),
    ...(o.lineSpacingPct ? { lineSpacingPct: o.lineSpacingPct } : {}),
  }
}

/** one text element from a list of lines sharing a style */
function textBox(
  x: number,
  y: number,
  w: number,
  h: number,
  lines: string[],
  o: TextOpts,
): ElementSpec {
  return { kind: 'text', x, y, w, h, paragraphs: lines.map((line) => para(line, o)) }
}

/**
 * Height a text box needs, from a crude character-per-line estimate. Exact
 * measurement would need the shaping engine; this only has to be close enough
 * that boxes do not overlap, and it errs generous.
 */
export function estimateTextHeight(lines: string[], sizePt: number, widthPx: number): number {
  const charW = sizePt * 0.55
  const perLine = Math.max(1, Math.floor(widthPx / charW))
  const rows = lines.reduce((n, line) => n + Math.max(1, Math.ceil(line.length / perLine)), 0)
  return Math.ceil(rows * sizePt * 1.35 * 1.34) // pt → px at 96dpi, with leading
}

/** Trim a list so a page stays a page rather than a document. */
function cap<T>(items: T[] | undefined, max: number): T[] {
  return (items ?? []).slice(0, max)
}

// ── layouts ─────────────────────────────────────────────

function coverPage(c: PageContent, th: Theme): RenderedPage {
  const elements: ElementSpec[] = []
  // an accent band anchors the title block instead of it floating in space
  elements.push({ kind: 'rect', x: 0, y: 0, w: 14, h: CANVAS.h, fill: th.accent })
  const titleH = estimateTextHeight([c.title], 54, CANVAS.w - M.x * 2 - 120)
  const top = Math.max(M.top, (CANVAS.h - titleH - (c.subtitle ? 90 : 0)) / 2)
  elements.push(
    textBox(M.x, top, CANVAS.w - M.x * 2 - 60, titleH, [c.title], {
      size: 54,
      color: th.text,
      bold: true,
      font: th.headFont,
    }),
  )
  if (c.subtitle) {
    elements.push(
      textBox(M.x, top + titleH + 22, CANVAS.w - M.x * 2 - 200, 64, [c.subtitle], {
        size: 22,
        color: th.muted,
        font: th.bodyFont,
      }),
    )
  }
  return { background: th.bg, elements }
}

function bulletsPage(c: PageContent, th: Theme): RenderedPage {
  const elements = [titleBlock(c, th)].flat()
  const bodyTop = M.top + 96
  const lines = cap(c.bullets, 6)
  if (lines.length) {
    const w = CANVAS.w - M.x * 2
    elements.push(
      textBox(M.x, bodyTop, w, CANVAS.h - bodyTop - M.bottom, lines, {
        size: 20,
        color: th.text,
        font: th.bodyFont,
        bullet: true,
        lineSpacingPct: 150,
      }),
    )
  }
  return withSource(c, th, { background: th.bg, elements })
}

function imageRightPage(c: PageContent, th: Theme): RenderedPage {
  const elements = [titleBlock(c, th)].flat()
  const bodyTop = M.top + 96
  const colW = (CANVAS.w - M.x * 2 - M.gap * 2) / 2
  const lines = cap(c.bullets, 5)
  if (lines.length) {
    elements.push(
      textBox(M.x, bodyTop, colW, CANVAS.h - bodyTop - M.bottom, lines, {
        size: 19,
        color: th.text,
        font: th.bodyFont,
        bullet: true,
        lineSpacingPct: 150,
      }),
    )
  }
  const imgX = M.x + colW + M.gap * 2
  const imgW = CANVAS.w - imgX - M.x
  const imgH = CANVAS.h - bodyTop - M.bottom
  if (c.imageUrl) {
    elements.push({ kind: 'image', x: imgX, y: bodyTop, w: imgW, h: imgH, url: c.imageUrl })
  } else {
    // a filled panel rather than a gap: an empty right half reads as a bug
    elements.push({ kind: 'rect', x: imgX, y: bodyTop, w: imgW, h: imgH, fill: th.surface })
  }
  return withSource(c, th, { background: th.bg, elements })
}

function cardsPage(c: PageContent, th: Theme): RenderedPage {
  const elements = [titleBlock(c, th)].flat()
  const cards = cap(c.cards, 3)
  const top = M.top + 110
  const h = 300
  const w = (CANVAS.w - M.x * 2 - M.gap * (cards.length - 1)) / Math.max(1, cards.length)
  cards.forEach((card, i) => {
    const x = M.x + i * (w + M.gap)
    elements.push({ kind: 'rect', x, y: top, w, h, fill: th.surface })
    elements.push({ kind: 'rect', x, y: top, w: 6, h, fill: th.accent })
    elements.push(
      textBox(x + 28, top + 26, w - 56, 74, [card.heading], {
        size: 22,
        color: th.text,
        bold: true,
        font: th.headFont,
      }),
    )
    elements.push(
      textBox(x + 28, top + 104, w - 56, h - 130, [card.body], {
        size: 16,
        color: th.muted,
        font: th.bodyFont,
        lineSpacingPct: 140,
      }),
    )
  })
  return withSource(c, th, { background: th.bg, elements })
}

function bigNumberPage(c: PageContent, th: Theme): RenderedPage {
  const elements = [titleBlock(c, th)].flat()
  const figure = c.figure ?? { value: '—', caption: c.subtitle ?? '' }
  const top = M.top + 150
  elements.push(
    textBox(M.x, top, CANVAS.w - M.x * 2, 190, [figure.value], {
      size: 128,
      color: th.accent,
      bold: true,
      align: 'center',
      font: th.headFont,
    }),
  )
  if (figure.caption) {
    elements.push(
      textBox(M.x + 100, top + 210, CANVAS.w - M.x * 2 - 200, 90, [figure.caption], {
        size: 22,
        color: th.muted,
        align: 'center',
        font: th.bodyFont,
      }),
    )
  }
  return withSource(c, th, { background: th.bg, elements })
}

function kpisPage(c: PageContent, th: Theme): RenderedPage {
  const elements = [titleBlock(c, th)].flat()
  const kpis = cap(c.kpis, 4)
  const top = M.top + 130
  const h = 210
  const w = (CANVAS.w - M.x * 2 - M.gap * (kpis.length - 1)) / Math.max(1, kpis.length)
  kpis.forEach((kpi, i) => {
    const x = M.x + i * (w + M.gap)
    elements.push({ kind: 'rect', x, y: top, w, h, fill: th.surface })
    elements.push(
      textBox(x + 20, top + 34, w - 40, 92, [kpi.value], {
        size: 52,
        color: th.accent,
        bold: true,
        align: 'center',
        font: th.headFont,
      }),
    )
    elements.push(
      textBox(x + 20, top + 132, w - 40, 60, [kpi.label], {
        size: 15,
        color: th.muted,
        align: 'center',
        font: th.bodyFont,
      }),
    )
  })
  return withSource(c, th, { background: th.bg, elements })
}

function comparisonPage(c: PageContent, th: Theme): RenderedPage {
  const elements = [titleBlock(c, th)].flat()
  const cols = cap(c.cards, 2)
  const top = M.top + 110
  const h = CANVAS.h - top - M.bottom - 20
  const w = (CANVAS.w - M.x * 2 - M.gap) / Math.max(1, cols.length)
  cols.forEach((col, i) => {
    const x = M.x + i * (w + M.gap)
    // the second column carries the accent, so the contrast is visible at a glance
    const accent = i === 1
    elements.push({ kind: 'rect', x, y: top, w, h, fill: accent ? th.accent : th.surface })
    elements.push(
      textBox(x + 30, top + 28, w - 60, 70, [col.heading], {
        size: 24,
        color: accent ? th.onAccent : th.text,
        bold: true,
        font: th.headFont,
      }),
    )
    elements.push(
      textBox(x + 30, top + 104, w - 60, h - 130, [col.body], {
        size: 17,
        color: accent ? th.onAccent : th.muted,
        font: th.bodyFont,
        lineSpacingPct: 145,
      }),
    )
  })
  return withSource(c, th, { background: th.bg, elements })
}

function closingPage(c: PageContent, th: Theme): RenderedPage {
  const elements: ElementSpec[] = []
  const titleH = estimateTextHeight([c.title], 46, CANVAS.w - M.x * 2)
  const top = (CANVAS.h - titleH - 80) / 2
  elements.push(
    textBox(M.x, top, CANVAS.w - M.x * 2, titleH, [c.title], {
      size: 46,
      color: th.text,
      bold: true,
      align: 'center',
      font: th.headFont,
    }),
  )
  if (c.subtitle) {
    elements.push(
      textBox(M.x, top + titleH + 20, CANVAS.w - M.x * 2, 70, [c.subtitle], {
        size: 20,
        color: th.muted,
        align: 'center',
        font: th.bodyFont,
      }),
    )
  }
  return { background: th.bg, elements }
}

/** Shared heading: title, optional deck, and a rule under it. */
function titleBlock(c: PageContent, th: Theme): ElementSpec[] {
  const w = CANVAS.w - M.x * 2
  return [
    textBox(M.x, M.top, w, 62, [c.title], {
      size: 32,
      color: th.text,
      bold: true,
      font: th.headFont,
    }),
    { kind: 'rect', x: M.x, y: M.top + 66, w: 64, h: 4, fill: th.accent },
  ]
}

/** Provenance line, bottom-left, on any page whose content carries one. */
function withSource(c: PageContent, th: Theme, page: RenderedPage): RenderedPage {
  if (!c.source) return page
  page.elements.push(
    textBox(M.x, CANVAS.h - M.bottom + 8, CANVAS.w - M.x * 2, 34, [c.source], {
      size: 12,
      color: th.muted,
      font: th.bodyFont,
    }),
  )
  return page
}

const LAYOUTS: Record<LayoutId, (c: PageContent, th: Theme) => RenderedPage> = {
  cover: coverPage,
  bullets: bulletsPage,
  image_right: imageRightPage,
  cards: cardsPage,
  big_number: bigNumberPage,
  kpis: kpisPage,
  comparison: comparisonPage,
  closing: closingPage,
}

/**
 * Render one page. A layout whose content never arrived (no cards on a cards
 * page) degrades to bullets rather than emitting an empty frame.
 */
export function layoutPage(layout: LayoutId, content: PageContent, theme: Theme): RenderedPage {
  const empty =
    (layout === 'cards' && !content.cards?.length) ||
    (layout === 'comparison' && !content.cards?.length) ||
    (layout === 'kpis' && !content.kpis?.length) ||
    (layout === 'big_number' && !content.figure)
  const chosen = empty ? 'bullets' : layout
  return LAYOUTS[chosen](content, theme)
}
