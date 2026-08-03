/**
 * Raw OOXML access — the escape hatch for what the model layer cannot express.
 *
 * Everything else in this engine works on the parsed model: elements, fills,
 * paragraphs. Some things have no model — a gradient stop list on a layout, an
 * `<a:effectLst>` nobody parses, a content-type override — and without a way to
 * reach them the answer to "can you fix this one attribute" is no.
 *
 * The cost of an escape hatch is that it can write a file PowerPoint refuses to
 * open, and the failure is silent until someone opens it. So a raw edit is not a
 * passthrough:
 *
 *   1. the part must resolve, exist, and be XML;
 *   2. `find` must occur exactly once, so an edit meant for one place cannot
 *      land in five;
 *   3. the result must be well-formed XML;
 *   4. the affected slides must still parse.
 *
 * A failure at any step restores the original bytes and returns a message. That
 * catches malformed output and structural damage; it does not catch XML that is
 * well-formed and parseable but schema-invalid, which is the residual risk this
 * tool cannot design away — hence the narrow find/replace shape rather than
 * "here is a new part".
 *
 * Undo is the caller's job: the session snapshot already covers archive entries
 * and the parsed slides, so a raw edit rolls back like any other.
 */
import { XMLValidator } from 'fast-xml-parser'
import { parseSlideFromArchive, type OpenedPptx } from './index'
import type { Slide } from './types'

/** Largest part this will hand back or edit; beyond it the agent should use the model tools. */
const MAX_PART_BYTES = 400_000

export interface RawPartInfo {
  /** zip-internal path, e.g. ppt/slides/slide1.xml */
  path: string
  /** the short name this part answers to, when it has one */
  ref?: string
  bytes: number
}

export type RawResult<T> = ({ ok: true } & T) | { ok: false; error: string }

const decoder = new TextDecoder()
const encoder = new TextEncoder()

/**
 * Resolve a part reference to a zip path. Two forms, both of which the model
 * writes naturally: a short name like `/slide[2]` or `/theme`, and a literal
 * zip path. Short names are resolved through the presentation's slide order, so
 * they stay correct after a reorder — the file numbering does not.
 */
export function resolvePartPath(opened: OpenedPptx, ref: string): string | null {
  const trimmed = ref.trim().replace(/^\/+/, '')
  if (!trimmed) return null
  // a literal path wins when it exists, so anything read out of a rels file works
  if (opened.archive.has(trimmed)) return trimmed
  const { slidePaths } = opened.archive.readPresentation()

  const indexed = /^(slide|slideLayout|slideMaster|notesSlide)\[(\d+)\]$/.exec(trimmed)
  if (indexed) {
    const kind = indexed[1]!
    const n = Number(indexed[2])
    if (!Number.isInteger(n) || n < 1) return null
    if (kind === 'slide') return slidePaths[n - 1] ?? null
    if (kind === 'notesSlide') {
      const slidePath = slidePaths[n - 1]
      if (!slidePath) return null
      for (const rel of opened.archive.readRels(slidePath).values()) {
        if (rel.type.endsWith('/notesSlide')) return resolveRelative(slidePath, rel.target)
      }
      return null
    }
    // layouts and masters are not ordered by the presentation; use file numbering
    const path = `ppt/${kind}s/${kind}${n}.xml`
    return opened.archive.has(path) ? path : null
  }

  if (trimmed === 'presentation') return 'ppt/presentation.xml'
  if (trimmed === 'theme')
    return opened.archive.has('ppt/theme/theme1.xml') ? 'ppt/theme/theme1.xml' : null
  if (trimmed === 'contentTypes') return '[Content_Types].xml'
  return null
}

/** Resolve a rels target against the part that declared it. */
function resolveRelative(fromPart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const dir = fromPart.slice(0, fromPart.lastIndexOf('/'))
  const parts = `${dir}/${target}`.split('/')
  const out: string[] = []
  for (const segment of parts) {
    if (segment === '.' || segment === '') continue
    if (segment === '..') out.pop()
    else out.push(segment)
  }
  return out.join('/')
}

/** The XML parts worth naming, in the order a reader would look for them. */
export function listRawParts(opened: OpenedPptx): RawPartInfo[] {
  const { slidePaths } = opened.archive.readPresentation()
  const refs = new Map<string, string>()
  refs.set('ppt/presentation.xml', '/presentation')
  slidePaths.forEach((path, i) => refs.set(path, `/slide[${i + 1}]`))

  const out: RawPartInfo[] = []
  for (const [path, bytes] of opened.archive.entries) {
    if (!path.endsWith('.xml') && !path.endsWith('.rels')) continue
    const ref = refs.get(path)
    out.push({ path, bytes: bytes.length, ...(ref ? { ref } : {}) })
  }
  return out.sort((a, b) => a.path.localeCompare(b.path))
}

export function readRawPart(
  opened: OpenedPptx,
  ref: string,
): RawResult<{ path: string; xml: string }> {
  const path = resolvePartPath(opened, ref)
  if (!path) return { ok: false, error: `No part matches "${ref}"` }
  const bytes = opened.archive.readBytes(path)
  if (!bytes) return { ok: false, error: `Part "${path}" is not in the package` }
  if (bytes.length > MAX_PART_BYTES) {
    return {
      ok: false,
      error: `Part "${path}" is ${Math.round(bytes.length / 1024)} KB, over the ${MAX_PART_BYTES / 1024} KB raw limit — use the model tools for a part this size`,
    }
  }
  return { ok: true, path, xml: decoder.decode(bytes) }
}

/** Slides whose parsed form depends on a part: itself, or all of them for shared parts. */
function affectedSlides(opened: OpenedPptx, path: string): string[] {
  const { slidePaths } = opened.archive.readPresentation()
  if (slidePaths.includes(path)) return [path]
  // layouts, masters, themes and the presentation part feed every slide's
  // inheritance, so a change to one is a change to all of them
  return slidePaths
}

/**
 * Replace one exact occurrence of `find` with `replace` in a part, then prove
 * the package still parses before keeping the result.
 */
export function editRawPart(
  opened: OpenedPptx,
  ref: string,
  find: string,
  replace: string,
): RawResult<{ path: string; reparsedSlides: number }> {
  const path = resolvePartPath(opened, ref)
  if (!path) return { ok: false, error: `No part matches "${ref}"` }
  const before = opened.archive.readBytes(path)
  if (!before) return { ok: false, error: `Part "${path}" is not in the package` }
  if (before.length > MAX_PART_BYTES) {
    return {
      ok: false,
      error: `Part "${path}" is too large to edit raw (${Math.round(before.length / 1024)} KB)`,
    }
  }
  if (!find) return { ok: false, error: 'find must not be empty' }

  const xml = decoder.decode(before)
  const first = xml.indexOf(find)
  if (first < 0) {
    return {
      ok: false,
      error: `find does not occur in "${path}" — read the part first and copy the text exactly`,
    }
  }
  if (xml.indexOf(find, first + find.length) >= 0) {
    const count = xml.split(find).length - 1
    return {
      ok: false,
      error: `find occurs ${count} times in "${path}"; it must match exactly once — include more surrounding text`,
    }
  }

  const next = xml.slice(0, first) + replace + xml.slice(first + find.length)
  const valid = XMLValidator.validate(next)
  if (valid !== true) {
    return {
      ok: false,
      error: `The result is not well-formed XML: ${valid.err.msg} (line ${valid.err.line})`,
    }
  }

  opened.archive.entries.set(path, encoder.encode(next))
  const targets = affectedSlides(opened, path)
  const reparsed: Array<{ index: number; slide: Slide }> = []
  for (const slidePath of targets) {
    const index = opened.deck.slides.findIndex((s) => s.path === slidePath)
    if (index < 0) continue
    // the parser throws on structural damage as readily as it returns null, and
    // either way the original bytes go back: a part that no longer parses would
    // leave the session rendering one document and saving another
    let slide: Slide | null
    try {
      slide = parseSlideFromArchive(opened.archive, slidePath)
    } catch {
      slide = null
    }
    if (!slide) {
      opened.archive.entries.set(path, before)
      return {
        ok: false,
        error: `The edit left "${slidePath}" unparseable; the part was restored unchanged`,
      }
    }
    reparsed.push({ index, slide })
  }
  for (const { index, slide } of reparsed) opened.deck.slides[index] = slide
  return { ok: true, path, reparsedSlides: reparsed.length }
}
