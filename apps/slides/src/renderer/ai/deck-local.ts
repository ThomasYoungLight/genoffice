import {
  layoutPage,
  normalizeLayout,
  themeFor,
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
  }): Promise<{ ok: boolean; error?: string }>
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
  /** aligned with `pages`: whether each page made it onto the canvas */
  doneFlags: boolean[]
  /** aligned with `pages`: last failure reason, when there was one */
  errors: Array<string | undefined>
  cancelled: boolean
}

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
): PageContent {
  const r = (raw && typeof raw === 'object' ? raw : {}) as Record<string, unknown>
  const figureRaw =
    r.figure && typeof r.figure === 'object' ? (r.figure as Record<string, unknown>) : null
  const figureValue = figureRaw ? str(figureRaw.value) : ''
  const bullets = strList(r.bullets, 6)
  const cards = pairList(r.cards, 'heading', 'body', 3)
  // A KPI page exists to show figures. Asked for them with none to hand, models
  // fill the slot with icons or single words, which renders as a row of huge
  // meaningless glyphs — worse than the layout the page degrades to instead.
  const kpis = pairList(r.kpis, 'value', 'label', 4).filter((k) => /\d/.test(k.value))
  return {
    title: str(r.title) || fallbackTitle,
    ...(str(r.subtitle) ? { subtitle: str(r.subtitle) } : {}),
    ...(bullets.length ? { bullets } : {}),
    ...(cards.length ? { cards } : {}),
    ...(kpis.length ? { kpis } : {}),
    ...(figureValue ? { figure: { value: figureValue, caption: str(figureRaw?.caption) } } : {}),
    ...(str(r.source) ? { source: str(r.source) } : {}),
    ...(imageUrl ? { imageUrl } : {}),
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
  let landed = 0

  for (let i = 0; i < total; i++) {
    if (signal?.aborted) return { landed, doneFlags, errors, cancelled: true }
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
      if (signal?.aborted) return { landed, doneFlags, errors, cancelled: true }
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
      if (r.ok && r.content) content = pageContentFrom(r.content, title, imageUrl)
      else lastErr = r.error ?? lastErr
    }
    // A page whose content call failed twice is still worth landing: the
    // planner already produced a title, and a titled page in the right place
    // beats a gap the user has to notice and fill.
    const degraded = !content
    const page = layoutPage(layout, content ?? { title, ...(imageUrl ? { imageUrl } : {}) }, theme)

    const r = await args.renderLocalPage({
      page,
      // "first" means first to land, not first planned: if page 1 failed, the
      // deck must still be replaced by whichever page arrives first
      first: landed === 0,
      replaceExisting: args.insertMode === 'replace',
      deckName: args.deckName,
    })
    if (r.ok) {
      landed += 1
      doneFlags[i] = true
      errors[i] = degraded ? `${lastErr} (page landed with its title only)` : undefined
      args.onPage?.(i, 'done', errors[i])
    } else {
      errors[i] = r.error ?? 'page could not be added to the deck'
      args.onPage?.(i, 'error', errors[i])
    }
  }
  return { landed, doneFlags, errors, cancelled: signal?.aborted === true }
}
