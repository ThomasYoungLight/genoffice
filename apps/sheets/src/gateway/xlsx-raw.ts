/**
 * Raw OOXML access for workbooks — the escape hatch for what the DSL cannot
 * express.
 *
 * Everything else in this gateway works through the workbook model: cells,
 * formats, filters, charts. Some things have no model — a `<fileVersion>`
 * attribute, a custom part, an `<extLst>` nobody reads — and without a way to
 * reach them the answer to "can you fix this one attribute" is no.
 *
 * The cost of an escape hatch is that it can write a workbook Excel refuses to
 * open, and the failure is silent until someone opens it. So a raw edit is not
 * a passthrough. Three of the four gates live here, on plain strings:
 *
 *   1. the part must resolve and exist;
 *   2. `find` must occur exactly once, so an edit meant for one place cannot
 *      land in five;
 *   3. the result must be well-formed XML.
 *
 * The fourth — the workbook must still read — cannot be done on a string. The
 * slides engine re-parses affected slides in memory; sheets has no in-memory
 * parse to re-run, so the caller round-trips the candidate through the sidecar
 * (see `verifyRawEdit` in the main process) and rolls back on failure. Without
 * that gate an edit to styles.xml produces a workbook Excel rejects, and we
 * would not find out until the user did.
 *
 * None of this catches XML that is well-formed and parseable but
 * schema-invalid, which is the residual risk the feature cannot design away —
 * hence the narrow find/replace shape rather than "here is a new part".
 */
import { XMLValidator } from 'fast-xml-parser'

/** Largest part this will hand back or edit; beyond it, use the model tools. */
export const RAW_MAX_PART_BYTES = 400_000

export interface RawPartInfo {
  /** zip-internal entry name, e.g. xl/worksheets/sheet1.xml */
  path: string
  /** the short name this part answers to, when it has one */
  ref?: string
  bytes: number
}

export type RawResult<T> = ({ ok: true } & T) | { ok: false; error: string }

/** What `listRawParts` and `resolveRawPartRef` need to know about the package. */
export interface RawArchiveView {
  /** every entry name, with its uncompressed size */
  readonly entries: readonly { readonly name: string; readonly uncompressedSize: number }[]
  /** worksheet part paths in the workbook's own sheet order, not file numbering */
  readonly sheetPaths: readonly string[]
}

/**
 * Worksheet part paths in workbook order.
 *
 * File numbering is not sheet order: sheet3.xml can be the first tab, and after
 * a reorder or a delete the numbering says nothing at all. `/sheet[2]` has to
 * mean the second tab, so it resolves the way Excel does — through the sheet
 * list in workbook.xml and the relationship it names.
 */
export function sheetPathsInOrder(workbookXml: string, relsXml: string): string[] {
  const targetById = new Map<string, string>()
  for (const rel of relsXml.match(/<Relationship\b[^>]*\/?>/g) ?? []) {
    // attribute order varies by producer, so match each independently
    const id = /\bId="([^"]+)"/.exec(rel)?.[1]
    const target = /\bTarget="([^"]+)"/.exec(rel)?.[1]
    if (id && target) targetById.set(id, target)
  }
  const out: string[] = []
  for (const sheet of workbookXml.match(/<sheet\b[^>]*\/?>/g) ?? []) {
    const rid = /\br:id="([^"]+)"/.exec(sheet)?.[1]
    const target = rid ? targetById.get(rid) : undefined
    if (!target) continue
    out.push(
      target.startsWith('/')
        ? target.slice(1)
        : `xl/${target.replace(/^\/?xl\//, '').replace(/^\.\//, '')}`,
    )
  }
  return out
}

/** Short names for the parts worth naming, keyed by entry name. */
function shortNames(view: RawArchiveView): Map<string, string> {
  const refs = new Map<string, string>()
  refs.set('xl/workbook.xml', '/workbook')
  refs.set('xl/styles.xml', '/styles')
  refs.set('xl/sharedStrings.xml', '/sharedStrings')
  refs.set('xl/theme/theme1.xml', '/theme')
  refs.set('[Content_Types].xml', '/contentTypes')
  view.sheetPaths.forEach((path, i) => refs.set(path, `/sheet[${i + 1}]`))
  return refs
}

/**
 * Resolve a part reference to an entry name. Two forms, both of which a model
 * writes naturally: a short name like `/sheet[2]` or `/styles`, and a literal
 * entry name copied out of a listing or a rels file.
 */
export function resolveRawPartRef(view: RawArchiveView, ref: string): string | null {
  const trimmed = ref.trim().replace(/^\/+/, '')
  if (!trimmed) return null
  const names = new Set(view.entries.map((e) => e.name))
  // a literal name wins when it exists, so anything read out of a rels file works
  if (names.has(trimmed)) return trimmed
  // [Content_Types].xml survives the leading-slash strip as a literal too
  if (names.has(`/${trimmed}`)) return `/${trimmed}`

  const indexed = /^sheet\[(\d+)\]$/.exec(trimmed)
  if (indexed) {
    const n = Number(indexed[1])
    if (!Number.isInteger(n) || n < 1) return null
    return view.sheetPaths[n - 1] ?? null
  }
  for (const [path, short] of shortNames(view)) {
    if (short === `/${trimmed}` && names.has(path)) return path
  }
  return null
}

/** The XML parts worth listing, sorted by entry name. */
export function listRawParts(view: RawArchiveView): RawPartInfo[] {
  const refs = shortNames(view)
  return view.entries
    .filter((e) => e.name.endsWith('.xml') || e.name.endsWith('.rels'))
    .map((e) => {
      const ref = refs.get(e.name)
      return { path: e.name, bytes: e.uncompressedSize, ...(ref ? { ref } : {}) }
    })
    .sort((a, b) => a.path.localeCompare(b.path))
}

/**
 * Gates 2 and 3: replace one exact occurrence of `find`, and prove the result
 * is still well-formed XML. Pure, so the interesting failures are testable
 * without a workbook.
 */
export function applyRawEdit(
  path: string,
  xml: string,
  find: string,
  replace: string,
): RawResult<{ xml: string }> {
  if (!find) return { ok: false, error: 'find must not be empty' }

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
  return { ok: true, xml: next }
}

/** Gate 1's size half, shared by read and edit so they refuse the same parts. */
export function checkRawPartSize(path: string, bytes: number): string | null {
  if (bytes <= RAW_MAX_PART_BYTES) return null
  return (
    `Part "${path}" is ${Math.round(bytes / 1024)} KB, over the ` +
    `${Math.round(RAW_MAX_PART_BYTES / 1024)} KB raw limit — use the model tools for a part this size`
  )
}
