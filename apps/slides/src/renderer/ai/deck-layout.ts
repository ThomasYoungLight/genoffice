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

/** height reserved for the page title, and where the body starts under it */
const TITLE_H = 96
const BODY_TOP = M.top + TITLE_H + 24

/**
 * Type floor. A slide title carries from the back of the room at 36pt and a
 * body line at 18pt; below that a deck reads as a document projected. Fitting
 * shrinks within these ladders and then drops or clips content, rather than
 * shrinking on past the floor — an unreadable slide that fits is not a fit.
 * Captions and short KPI sublabels are the deliberate exceptions.
 */
const SIZES = {
  pageTitle: [36, 32, 28, 24],
  body: [20, 19, 18],
  cardBody: [18],
  caption: [15, 14, 13],
  source: [12, 11],
} as const

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
  // The layout name is the specific instruction and the page type is the
  // category, so the name is read first: a page typed "cover" that asks for
  // left_text_right_image wants the image.
  if (l.startsWith('cover')) return 'cover'
  if (l.startsWith('closing')) return 'closing'
  if (l.includes('kpi')) return 'kpis'
  if (l.includes('big_number') || l.includes('hero_big')) return 'big_number'
  if (l.includes('comparison') || l.includes('two_by_two')) return 'comparison'
  if (l.includes('card') || l.includes('three_column') || l.includes('timeline')) return 'cards'
  if (l.includes('image')) return 'image_right'
  const t = (type ?? '').toLowerCase()
  if (t === 'cover') return 'cover'
  if (t === 'closing') return 'closing'
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
 * Advance width in px per point of font size.
 *
 * `latin` is measured against what the renderer actually draws, not against
 * Calibri's nominal metrics: a 46pt line of mixed-case text came out at ~0.68,
 * and 0.55 (the typographic figure) made a two-line title report as one, which
 * overlapped whatever sat under it. 0.72 keeps the error on the safe side — a
 * box slightly too tall costs nothing, a box too short collides.
 *
 * `wide` is one em: a CJK glyph occupies the full square, near twice a Latin
 * one. Averaging the two into a single ratio silently under-counts every
 * Chinese, Japanese and Korean deck by about 40% — the app ships nineteen
 * locales, so that is not an edge case. Measured per character instead.
 */
const CHAR_W = { latin: 0.72, wide: 1.33 } as const

/**
 * Whether a character occupies a full-width cell. Ranges follow the same
 * set OfficeCLI uses for its own overflow check (Apache-2.0, iOfficeAI):
 * unified ideographs and extension A, compatibility ideographs, CJK symbols
 * and punctuation, fullwidth forms, kana, Hangul syllables, bopomofo.
 * Halfwidth katakana (FF61–FF9F) is deliberately not in it.
 */
function isWideChar(code: number): boolean {
  return (
    (code >= 0x4e00 && code <= 0x9fff) ||
    (code >= 0x3400 && code <= 0x4dbf) ||
    (code >= 0xf900 && code <= 0xfaff) ||
    (code >= 0x3000 && code <= 0x303f) ||
    (code >= 0xff01 && code <= 0xff60) ||
    (code >= 0x3040 && code <= 0x309f) ||
    (code >= 0x30a0 && code <= 0x30ff) ||
    (code >= 0xac00 && code <= 0xd7af) ||
    (code >= 0x3100 && code <= 0x312f)
  )
}

/** How many rendered rows a single line wraps into inside `widthPx`. */
function wrapRows(line: string, sizePt: number, widthPx: number): number {
  if (!line) return 1
  let rows = 1
  let w = 0
  for (const ch of line) {
    const cw = sizePt * (isWideChar(ch.codePointAt(0) ?? 0) ? CHAR_W.wide : CHAR_W.latin)
    if (w + cw > widthPx && w > 0) {
      rows++
      w = cw
    } else {
      w += cw
    }
  }
  return rows
}

/**
 * Height a text box needs. Exact measurement would need the shaping engine;
 * this only has to be close enough that boxes do not overlap, and it must err
 * generous.
 */
export function estimateTextHeight(lines: string[], sizePt: number, widthPx: number): number {
  const rows = lines.reduce((n, line) => n + wrapRows(line, sizePt, widthPx), 0)
  return Math.ceil(rows * sizePt * 1.35 * 1.34) // pt → px at 96dpi, with leading
}

/** Trim a list so a page stays a page rather than a document. */
function cap<T>(items: T[] | undefined, max: number): T[] {
  return (items ?? []).slice(0, max)
}

/**
 * Make text fit its box.
 *
 * A slide text box does not clip: text longer than the box keeps drawing, over
 * whatever is below it and off the bottom of the page. Nothing upstream can
 * fully prevent it — the model is told how long each slot should be, but a
 * five-step procedure will still arrive in a one-line slot sooner or later.
 * So fitting is enforced here, where the box size is known.
 *
 * Two steps, in the order that costs the reader least: shrink through the
 * given sizes, then drop lines, then clip the last one. The result always
 * fits, so a page can be crowded but never broken.
 */
function fitText(
  lines: string[],
  widthPx: number,
  heightPx: number,
  sizes: readonly number[],
): { size: number; lines: string[] } {
  const size = sizes.find((s) => estimateTextHeight(lines, s, widthPx) <= heightPx) ?? sizes.at(-1)!
  let out = lines
  while (out.length > 1 && estimateTextHeight(out, size, widthPx) > heightPx) out = out.slice(0, -1)
  const first = out[0]
  if (first && estimateTextHeight(out, size, widthPx) > heightPx) {
    // one line, still too tall: keep the characters that fit
    const rows = Math.max(1, Math.floor(heightPx / (size * 1.35 * 1.34)))
    out = [clipAt(first, charBudget(first, size, widthPx * rows))]
  }
  return { size, lines: out }
}

/** How many characters of `text` fit in `budgetPx` of total advance width. */
function charBudget(text: string, sizePt: number, budgetPx: number): number {
  let w = 0
  let n = 0
  for (const ch of text) {
    w += sizePt * (isWideChar(ch.codePointAt(0) ?? 0) ? CHAR_W.wide : CHAR_W.latin)
    if (w > budgetPx) break
    n++
  }
  return Math.max(8, n - 1)
}

/**
 * Cut a run of text to fit, ending somewhere a reader can stop.
 *
 * A cut mid-clause ("…assigned within 10 minutes: 83% (15 of 18); …") reads
 * as damage. Ending on the last complete sentence inside the budget reads as
 * an editor's choice, so a sentence boundary in the back half of the space
 * wins over the extra half-sentence it costs. Failing that, a word boundary
 * with an ellipsis, which at least does not sever a word.
 */
function clipAt(text: string, max: number): string {
  const room = text.slice(0, max)
  const sentence = Math.max(
    room.lastIndexOf('. '),
    room.lastIndexOf('! '),
    room.lastIndexOf('? '),
    // a trailing terminator with nothing after it is also a clean stop
    /[.!?]$/.test(room.trimEnd()) ? room.trimEnd().length - 1 : -1,
  )
  if (sentence > max * 0.5) return room.slice(0, sentence + 1).trimEnd()
  const space = room.lastIndexOf(' ')
  return `${(space > max * 0.6 ? room.slice(0, space) : room).trimEnd()}…`
}

// ── layouts ─────────────────────────────────────────────

function coverPage(c: PageContent, th: Theme): RenderedPage {
  const elements: ElementSpec[] = []
  // an accent band anchors the title block instead of it floating in space
  elements.push({ kind: 'rect', x: 0, y: 0, w: 14, h: CANVAS.h, fill: th.accent })
  const titleW = CANVAS.w - M.x * 2 - 60
  // a cover title gets the space it needs, down to a size that still reads big
  const title = fitText([c.title], titleW, 300, [54, 46, 40, 34])
  const titleH = estimateTextHeight(title.lines, title.size, titleW)
  const subW = CANVAS.w - M.x * 2 - 200
  const sub = c.subtitle ? fitText([c.subtitle], subW, 96, [22, 20, 18]) : null
  const subH = sub ? estimateTextHeight(sub.lines, sub.size, subW) : 0
  const top = Math.max(M.top, (CANVAS.h - titleH - (sub ? subH + 22 : 0)) / 2)
  elements.push(
    textBox(M.x, top, titleW, titleH, title.lines, {
      size: title.size,
      color: th.text,
      bold: true,
      font: th.headFont,
    }),
  )
  if (sub) {
    elements.push(
      textBox(M.x, top + titleH + 22, subW, subH, sub.lines, {
        size: sub.size,
        color: th.muted,
        font: th.bodyFont,
      }),
    )
  }
  return { background: th.bg, elements }
}

function bulletsPage(c: PageContent, th: Theme): RenderedPage {
  const elements = [titleBlock(c, th)].flat()
  const bodyTop = BODY_TOP
  const lines = cap(c.bullets, 6)
  if (lines.length) {
    const w = CANVAS.w - M.x * 2
    const h = CANVAS.h - bodyTop - (c.source ? M.bottom + 20 : M.bottom)
    // bullets carry 1.5 line spacing, so the fit budget is proportionally smaller
    const fit = fitText(lines, w, h / 1.5, SIZES.body)
    elements.push(
      textBox(M.x, bodyTop, w, h, fit.lines, {
        size: fit.size,
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
  const bodyTop = BODY_TOP
  const colW = (CANVAS.w - M.x * 2 - M.gap * 2) / 2
  const lines = cap(c.bullets, 5)
  if (lines.length) {
    const h = CANVAS.h - bodyTop - (c.source ? M.bottom + 20 : M.bottom)
    const fit = fitText(lines, colW, h / 1.5, SIZES.body)
    elements.push(
      textBox(M.x, bodyTop, colW, h, fit.lines, {
        size: fit.size,
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
  const top = BODY_TOP
  // tall enough to hold a full sentence at the 18pt floor; the alternative is
  // shrinking under the floor, which is what the floor exists to prevent
  const h = 380
  const w = (CANVAS.w - M.x * 2 - M.gap * (cards.length - 1)) / Math.max(1, cards.length)
  cards.forEach((card, i) => {
    const x = M.x + i * (w + M.gap)
    // A solid panel and nothing else: the card with a coloured left-border
    // stripe is the other well-known machine-made-slide tell. The heading
    // carries the accent instead, which separates the cards just as well.
    elements.push({ kind: 'rect', x, y: top, w, h, fill: th.surface })
    const head = fitText([card.heading], w - 56, 74, [22, 20, 18])
    elements.push(
      textBox(x + 28, top + 26, w - 56, 74, head.lines, {
        size: head.size,
        color: th.accent,
        bold: true,
        font: th.headFont,
      }),
    )
    const body = fitText([card.body], w - 56, (h - 130) / 1.4, SIZES.cardBody)
    elements.push(
      textBox(x + 28, top + 104, w - 56, h - 130, body.lines, {
        size: body.size,
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
  const top = BODY_TOP + 20
  // 128pt occupies ~232px of line box; the box is sized for it, and a figure
  // too long to sit on one line steps down instead of running over the caption
  const value = fitText([figure.value], CANVAS.w - M.x * 2, 240, [128, 108, 88, 70])
  elements.push(
    textBox(M.x, top, CANVAS.w - M.x * 2, 240, value.lines, {
      size: value.size,
      color: th.accent,
      bold: true,
      align: 'center',
      font: th.headFont,
    }),
  )
  if (figure.caption) {
    const capW = CANVAS.w - M.x * 2 - 200
    const fit = fitText([figure.caption], capW, 90, [22, 20, 18])
    elements.push(
      textBox(M.x + 100, top + 254, capW, 90, fit.lines, {
        size: fit.size,
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
  const top = BODY_TOP + 10
  const h = 210
  const w = (CANVAS.w - M.x * 2 - M.gap * (kpis.length - 1)) / Math.max(1, kpis.length)
  kpis.forEach((kpi, i) => {
    const x = M.x + i * (w + M.gap)
    elements.push({ kind: 'rect', x, y: top, w, h, fill: th.surface })
    // 52pt needs ~100px of line box; a longer figure ("$4.2M/quarter") shrinks
    const value = fitText([kpi.value], w - 40, 100, [52, 44, 36, 30])
    elements.push(
      textBox(x + 20, top + 30, w - 40, 100, value.lines, {
        size: value.size,
        color: th.accent,
        bold: true,
        align: 'center',
        font: th.headFont,
      }),
    )
    const label = fitText([kpi.label], w - 40, 60, SIZES.caption)
    elements.push(
      textBox(x + 20, top + 138, w - 40, 60, label.lines, {
        size: label.size,
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
  const top = BODY_TOP
  const h = CANVAS.h - top - M.bottom - 20
  const w = (CANVAS.w - M.x * 2 - M.gap) / Math.max(1, cols.length)
  cols.forEach((col, i) => {
    const x = M.x + i * (w + M.gap)
    // the second column carries the accent, so the contrast is visible at a glance
    const accent = i === 1
    elements.push({ kind: 'rect', x, y: top, w, h, fill: accent ? th.accent : th.surface })
    const head = fitText([col.heading], w - 60, 70, [24, 21, 18])
    elements.push(
      textBox(x + 30, top + 28, w - 60, 70, head.lines, {
        size: head.size,
        color: accent ? th.onAccent : th.text,
        bold: true,
        font: th.headFont,
      }),
    )
    const body = fitText([col.body], w - 60, (h - 130) / 1.45, SIZES.cardBody)
    elements.push(
      textBox(x + 30, top + 104, w - 60, h - 130, body.lines, {
        size: body.size,
        color: accent ? th.onAccent : th.muted,
        font: th.bodyFont,
        lineSpacingPct: 145,
      }),
    )
  })
  return withSource(c, th, { background: th.bg, elements })
}

/**
 * Closing page: title, an optional line under it, and — because a closing page
 * is often "here is what to do on Monday" — an optional short list. Without
 * the list a procedural ending has nowhere to go but the subtitle, which is
 * how a five-step plan ends up crushed into one line.
 */
function closingPage(c: PageContent, th: Theme): RenderedPage {
  const elements: ElementSpec[] = []
  const w = CANVAS.w - M.x * 2
  const title = fitText([c.title], w, 200, [46, 40, 34])
  const titleH = estimateTextHeight(title.lines, title.size, w)
  const steps = cap(c.bullets, 4)
  const sub = c.subtitle ? fitText([c.subtitle], w, 76, [20, 18]) : null
  const subH = sub ? estimateTextHeight(sub.lines, sub.size, w) : 0
  const listW = w - 220
  const list = steps.length ? fitText(steps, listW, 200 / 1.5, SIZES.cardBody) : null
  const listH = list ? estimateTextHeight(list.lines, list.size, listW) * 1.5 : 0

  const blockH = titleH + (sub ? subH + 20 : 0) + (list ? listH + 28 : 0)
  const top = Math.max(M.top, (CANVAS.h - blockH) / 2)
  elements.push(
    textBox(M.x, top, w, titleH, title.lines, {
      size: title.size,
      color: th.text,
      bold: true,
      align: 'center',
      font: th.headFont,
    }),
  )
  let y = top + titleH
  if (sub) {
    elements.push(
      textBox(M.x, y + 20, w, subH, sub.lines, {
        size: sub.size,
        color: th.muted,
        align: 'center',
        font: th.bodyFont,
      }),
    )
    y += subH + 20
  }
  if (list) {
    elements.push(
      textBox(M.x + 110, y + 28, listW, listH, list.lines, {
        size: list.size,
        color: th.text,
        font: th.bodyFont,
        bullet: true,
        lineSpacingPct: 150,
      }),
    )
  }
  return { background: th.bg, elements }
}

/** Shared heading: title, optional deck, and a rule under it. */
function titleBlock(c: PageContent, th: Theme): ElementSpec[] {
  const w = CANVAS.w - M.x * 2
  // Fixed height: everything below is positioned against it, so a long title
  // shrinks to fit rather than pushing the page down. No rule or accent bar
  // under it — a stripe below a heading is the most recognisable
  // machine-made-slide tell; the whitespace above the body separates them.
  const fit = fitText([c.title], w, TITLE_H, [36, 32, 28, 24])
  return [
    textBox(M.x, M.top, w, TITLE_H, fit.lines, {
      size: fit.size,
      color: th.text,
      bold: true,
      font: th.headFont,
    }),
  ]
}

/** Provenance line, bottom-left, on any page whose content carries one. */
function withSource(c: PageContent, th: Theme, page: RenderedPage): RenderedPage {
  if (!c.source) return page
  const w = CANVAS.w - M.x * 2
  const fit = fitText([c.source], w, 34, SIZES.source)
  page.elements.push(
    textBox(M.x, CANVAS.h - M.bottom + 8, w, 34, fit.lines, {
      size: fit.size,
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
  // Degrade to whatever the page does have rather than always to bullets: a
  // KPI page whose figures were rejected often still has cards, and rendering
  // those beats rendering a title over an empty page.
  const fallback: LayoutId = content.cards?.length ? 'cards' : 'bullets'
  return LAYOUTS[empty ? fallback : layout](content, theme)
}
