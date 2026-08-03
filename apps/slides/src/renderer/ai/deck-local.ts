import {
  layoutPage,
  normalizeLayout,
  themeFor,
  type DeckChart,
  type LayoutId,
  type PageContent,
  type RenderedPage,
} from './deck-layout'

/**
 * Local deck generation: the page loop that runs when cloud page generation is
 * unavailable.
 *
 * The cloud path asks a service to write HTML per page and converts it. Without
 * a Genspark account that path is closed, so this one keeps the same stages
 * (style → outline → images, all already local) and replaces only the last one:
 * per page, ask the model for *content* in the slots the chosen layout has,
 * then hand the positioned result to the renderer. Geometry stays in
 * deck-layout.ts — see the note there on why the model is not asked for it.
 *
 * Pages are generated one at a time rather than in a concurrent batch: the
 * content call is small, and a page can only be appended once the page before
 * it exists, so ordering would cost more than the overlap saves.
 */

/** What the caller must supply; a subset of DeckAccess, kept narrow for testing. */
export interface LocalDeckDeps {
  planPageContent(args: {
    pageIndex: number
    totalPages: number
    title: string
    brief: string
    layout: LayoutId
    coreHook: string
    styleSkill: string
    topic?: string | undefined
    context?: string | undefined
    hasImage: boolean
    signal?: AbortSignal | undefined
  }): Promise<{ ok: boolean; content?: Record<string, unknown>; error?: string }>
  renderLocalPage(args: {
    page: RenderedPage
    /** first page of this run: with replaceExisting, the old deck goes away once it lands */
    first: boolean
    replaceExisting: boolean
    deckName?: string | undefined
    /** retry: rebuild the page already sitting at this index instead of appending */
    replaceIndex?: number | undefined
  }): Promise<{
    ok: boolean
    error?: string
    /**
     * What the layout engine got wrong, measured by the renderer rather than
     * estimated: overflow, overlap, off-canvas. Empty means the page is sound.
     */
    issues?: string[]
    /** where the page landed, so a retry can rebuild it in place */
    slideIndex?: number
  }>
}

export interface LocalDeckArgs extends LocalDeckDeps {
  /** planner output, one entry per page (loose shape, validated at point of use) */
  pages: Array<Record<string, unknown>>
  coreHook: string
  styleSkill: string
  topic?: string | undefined
  context?: string | undefined
  deckName?: string | undefined
  insertMode: 'replace' | 'append'
  signal?: AbortSignal | undefined
  /** progress hook; `index` is 0-based */
  onPage?(index: number, status: 'running' | 'done' | 'error', error?: string): void
}

export interface LocalDeckResult {
  landed: number
  /** 1-based pages whose chart carries illustrative rather than measured numbers */
  illustrative: number[]
  /** aligned with `pages`: whether each page made it onto the canvas */
  doneFlags: boolean[]
  /** aligned with `pages`: last failure reason, when there was one */
  errors: Array<string | undefined>
  cancelled: boolean
}

/**
 * Rebuild rounds allowed per page when the render disagrees with the layout.
 * The OfficeCLI pptx skill caps its own fix-verify loop at three for the same
 * reason: past that, a page that still reports issues is usually oscillating
 * rather than converging, and the honest move is to report it.
 */
const MAX_TIGHTEN = 2

/**
 * Longest string still readable as a figure in the 52pt KPI slot. "$4.2M",
 * "18%", "2.1 days" pass; "Acknowledged <15 min" is a metric name wearing a
 * number and does not.
 */
const KPI_VALUE_MAX = 12

const str = (v: unknown): string => (typeof v === 'string' ? v.trim() : '')

/** Non-empty strings from an array field, capped — models pad lists when unsure. */
function strList(v: unknown, max: number): string[] {
  if (!Array.isArray(v)) return []
  return v.map(str).filter(Boolean).slice(0, max)
}

function pairList<K extends string, V extends string>(
  v: unknown,
  a: K,
  b: V,
  max: number,
): Array<Record<K | V, string>> {
  if (!Array.isArray(v)) return []
  const out: Array<Record<K | V, string>> = []
  for (const item of v) {
    if (!item || typeof item !== 'object') continue
    const rec = item as Record<string, unknown>
    const first = str(rec[a])
    const second = str(rec[b])
    // a heading with no body still reads as a card; neither means an empty box
    if (!first && !second) continue
    out.push({ [a]: first, [b]: second } as Record<K | V, string>)
    if (out.length === max) break
  }
  return out
}

/**
 * Model output → PageContent. Every field is optional and independently
 * validated: a page that comes back with only a title still renders (the
 * layout degrades), and one bad field does not cost the page.
 */
export function pageContentFrom(
  raw: unknown,
  fallbackTitle: string,
  imageUrl?: string,
  hasContext = false,
): PageContent {
  const r = (raw && typeof raw === 'object' ? raw : {}) as Record<string, unknown>
  const figureRaw =
    r.figure && typeof r.figure === 'object' ? (r.figure as Record<string, unknown>) : null
  const figureValue = figureRaw ? str(figureRaw.value) : ''
  const bullets = strList(r.bullets, 6)
  const cards = pairList(r.cards, 'heading', 'body', 3)
  /**
   * A KPI page exists to show figures, and the value slot is 52pt: it holds a
   * figure, not a phrase. Requiring a digit is not enough on its own — a
   * metric *name* like "Acknowledged <15 min" contains one, and renders as a
   * huge clipped stub ("Acknowled…"). A real figure is short, so length is the
   * test that actually separates them. Below four valid ones the page degrades
   * to a layout that can carry prose.
   */
  const kpiItems = pairList(r.kpis, 'value', 'label', 4)
  const kpis = kpiItems.filter((k) => /\d/.test(k.value) && k.value.length <= KPI_VALUE_MAX)
  // A rejected KPI set is a list of metric names with descriptions, which is
  // what a card is. Keep the content and let it render as one, rather than
  // dropping the page's whole substance because the slot was wrong.
  const salvaged =
    kpis.length === 0 && kpiItems.length >= 2 && cards.length === 0
      ? kpiItems.slice(0, 3).map((k) => ({ heading: k.value, body: k.label }))
      : []
  const cardList = cards.length ? cards : salvaged
  const chart = chartFrom(r.chart, hasContext)
  return {
    title: str(r.title) || fallbackTitle,
    ...(str(r.subtitle) ? { subtitle: str(r.subtitle) } : {}),
    ...(bullets.length ? { bullets } : {}),
    ...(cardList.length ? { cards: cardList } : {}),
    ...(kpis.length ? { kpis } : {}),
    ...(figureValue ? { figure: { value: figureValue, caption: str(figureRaw?.caption) } } : {}),
    ...(chart ? { chart } : {}),
    ...(str(r.source) ? { source: str(r.source) } : {}),
    ...(imageUrl ? { imageUrl } : {}),
  }
}

const CHART_KINDS = new Set(['bar', 'barH', 'barStacked', 'line', 'area', 'pie', 'doughnut'])

/**
 * Chart data out of model output.
 *
 * Rejected unless it is internally consistent — every series the same length
 * as the categories, at least one real number — because a chart with a ragged
 * series renders as a broken axis rather than an error. `figures` records
 * where the numbers came from and defaults to 'sample': a deck generated
 * without source material must not present invented numbers as measured ones,
 * which is the same rule the add_chart tool enforces.
 */
export function chartFrom(raw: unknown, hasContext: boolean): DeckChart | undefined {
  if (!raw || typeof raw !== 'object') return undefined
  const r = raw as Record<string, unknown>
  const kind = str(r.kind)
  const categories = strList(r.categories, 12)
  if (!CHART_KINDS.has(kind) || categories.length < 2) return undefined
  const series: Array<{ name: string; values: number[] }> = []
  for (const item of Array.isArray(r.series) ? r.series : []) {
    if (!item || typeof item !== 'object') continue
    const rec = item as Record<string, unknown>
    const values = (Array.isArray(rec.values) ? rec.values : [])
      .map((v) => (typeof v === 'number' && Number.isFinite(v) ? v : Number.NaN))
      .slice(0, categories.length)
    // a series shorter than the categories would silently shift the axis
    if (values.length !== categories.length || values.some((v) => Number.isNaN(v))) continue
    series.push({ name: str(rec.name) || `Series ${series.length + 1}`, values })
    if (series.length === 4) break
  }
  if (series.length === 0) return undefined
  const declared = str(r.figures)
  const figures: DeckChart['figures'] =
    declared === 'document' || declared === 'search' ? (hasContext ? declared : 'sample') : 'sample'
  const title = str(r.title)
  return {
    kind: kind as DeckChart['kind'],
    ...(title ? { title } : {}),
    categories,
    series,
    figures,
  }
}

/** Cut a string at a word boundary, marking that something was removed. */
function shorten(s: string, max: number): string {
  if (s.length <= max) return s
  const cut = s.slice(0, max)
  const space = cut.lastIndexOf(' ')
  return `${(space > max * 0.6 ? cut.slice(0, space) : cut).trimEnd()}…`
}

/**
 * Less content, for a page the renderer says does not fit.
 *
 * The type floor rules out shrinking the way out of an overflow, so the answer
 * is to carry less: shorter lines first, then fewer of them. Each level is
 * roughly a third off, which converges in the two rounds the loop allows
 * without gutting the page on the first retry.
 */
export function tightenContent(c: PageContent, level: number): PageContent {
  if (level <= 0) return c
  /**
   * Proportional, not a fixed cap. A cap only bites on text that happens to
   * exceed it, so a page overflowing by a little could come back from a
   * "tightened" round completely unchanged and loop without converging.
   * Trimming a fraction always reduces something there is something to reduce.
   */
  const keep = level === 1 ? 0.75 : 0.55
  const trim = (s: string, floor = 24) =>
    s.length <= floor ? s : shorten(s, Math.max(floor, Math.round(s.length * keep)))
  // a list sheds one item per round before its lines get any shorter
  const bullets = c.bullets?.slice(0, Math.max(2, c.bullets.length - level)).map((b) => trim(b))
  const cards = c.cards
    ?.slice(0, level > 1 ? Math.max(1, c.cards.length - 1) : c.cards.length)
    .map((card) => ({ heading: trim(card.heading, 16), body: trim(card.body) }))
  return {
    ...c,
    title: trim(c.title, 40),
    ...(c.subtitle ? { subtitle: trim(c.subtitle, 32) } : {}),
    ...(bullets?.length ? { bullets } : {}),
    ...(cards?.length ? { cards } : {}),
    ...(c.kpis?.length ? { kpis: c.kpis.map((k) => ({ ...k, label: trim(k.label, 18) })) } : {}),
    ...(c.figure ? { figure: { ...c.figure, caption: trim(c.figure.caption, 28) } } : {}),
    // a source line is provenance, not filler — it is the last thing to go
    ...(c.source && level > 1 ? { source: trim(c.source, 40) } : {}),
  }
}

/** The already-resolved image URL for a page, if Step 1.5 found one. */
function imageFor(page: Record<string, unknown>): string | undefined {
  if (!Array.isArray(page.image_queries)) return undefined
  for (const q of page.image_queries as unknown[]) {
    const s = String(q).trim()
    if (/^(https?:\/\/|genimg:)/i.test(s)) return s
  }
  return undefined
}

export async function generateDeckLocally(args: LocalDeckArgs): Promise<LocalDeckResult> {
  const { pages, signal } = args
  const total = pages.length
  const doneFlags = new Array<boolean>(total).fill(false)
  const errors = new Array<string | undefined>(total).fill(undefined)
  const theme = themeFor(args.styleSkill)
  const illustrative: number[] = []
  let landed = 0

  for (let i = 0; i < total; i++) {
    if (signal?.aborted) return { landed, doneFlags, errors, illustrative, cancelled: true }
    const plan = pages[i] ?? {}
    const title = str(plan.title) || `Page ${i + 1}`
    const imageUrl = imageFor(plan)
    // A page the planner wanted a photo on, rendered by the layout that has
    // nowhere to put one, is a wasted image search. Only the generic bullets
    // layout is overridden: cards/kpis/comparison were chosen for a structure
    // the content needs, and that structure outranks the photo.
    const planned = normalizeLayout(str(plan.layout), str(plan.type))
    const layout = imageUrl && planned === 'bullets' ? 'image_right' : planned
    args.onPage?.(i, 'running')

    // One retry: the content call is cheap and its usual failure is a
    // malformed JSON body, which a second attempt normally gets right.
    let content: PageContent | null = null
    let lastErr = 'content generation failed'
    for (let attempt = 0; attempt < 2 && !content; attempt++) {
      if (signal?.aborted) return { landed, doneFlags, errors, illustrative, cancelled: true }
      const r = await args.planPageContent({
        pageIndex: i + 1,
        totalPages: total,
        title,
        brief: str(plan.brief),
        layout,
        coreHook: args.coreHook,
        styleSkill: args.styleSkill,
        topic: args.topic,
        context: args.context,
        hasImage: !!imageUrl,
        signal,
      })
      if (r.ok && r.content) content = pageContentFrom(r.content, title, imageUrl, !!args.context)
      else lastErr = r.error ?? lastErr
    }
    // A page whose content call failed twice is still worth landing: the
    // planner already produced a title, and a titled page in the right place
    // beats a gap the user has to notice and fill.
    const degraded = !content
    const base = content ?? { title, ...(imageUrl ? { imageUrl } : {}) }

    /**
     * Render, then look. The layout engine sizes boxes from an estimate of how
     * wide the text will be; the renderer knows. Every fitting bug in this
     * feature so far has been the gap between the two, so the page is measured
     * once it exists and rebuilt with less content if the measurement
     * disagrees — two rounds, then take what we have rather than seesaw.
     */
    let r = { ok: false } as Awaited<ReturnType<LocalDeckDeps['renderLocalPage']>>
    let unfixed: string[] = []
    for (let round = 0; round <= MAX_TIGHTEN; round++) {
      if (signal?.aborted) return { landed, doneFlags, errors, illustrative, cancelled: true }
      r = await args.renderLocalPage({
        page: layoutPage(layout, tightenContent(base, round), theme),
        // "first" means first to land, not first planned: if page 1 failed, the
        // deck must still be replaced by whichever page arrives first
        first: landed === 0,
        replaceExisting: args.insertMode === 'replace',
        deckName: args.deckName,
        ...(round > 0 ? { replaceIndex: r.slideIndex } : {}),
      })
      unfixed = r.issues ?? []
      // a page that could not be drawn at all is not a fitting problem
      if (!r.ok || unfixed.length === 0 || r.slideIndex === undefined) break
    }

    if (r.ok) {
      landed += 1
      doneFlags[i] = true
      if (base.chart?.figures === 'sample') illustrative.push(i + 1)
      errors[i] = degraded
        ? `${lastErr} (page landed with its title only)`
        : unfixed.length
          ? `layout still imperfect after ${MAX_TIGHTEN} retries: ${unfixed[0]}`
          : undefined
      args.onPage?.(i, 'done', errors[i])
    } else {
      errors[i] = r.error ?? 'page could not be added to the deck'
      args.onPage?.(i, 'error', errors[i])
    }
  }
  return { landed, doneFlags, errors, illustrative, cancelled: signal?.aborted === true }
}
