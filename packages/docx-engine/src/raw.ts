/**
 * Raw OOXML access for documents — the escape hatch for what the block model
 * cannot express.
 *
 * The gates are the slides ones (`pptx-engine/src/raw.ts`): the part must
 * resolve and be small enough, `find` must occur exactly once, the result must
 * be well-formed XML, and the document must still parse. What differs is one
 * refusal that has no equivalent there.
 *
 * **`word/document.xml` is readable but not writable.** `saveDocx` rebuilds its
 * body from the editor's blocks, so bytes written there are discarded at save —
 * the tool would report success and change nothing, which is worse than
 * refusing. Every other part is written verbatim through
 * `SaveOptions.partOverrides`, and reads are unrestricted because reading the
 * body is genuinely useful for working out what to change elsewhere.
 *
 * Residual risk, unchanged from slides: well-formed, parseable XML can still be
 * schema-invalid OOXML that Word refuses or repairs. Validation cannot close
 * that, which is why the shape stays a narrow find/replace.
 */
import JSZip from 'jszip'
import { XMLValidator } from 'fast-xml-parser'
import { parseDocx } from './parse'

/** Largest part this will hand back or edit; beyond it, use the model tools. */
export const RAW_MAX_PART_BYTES = 400_000

/** The body is regenerated from blocks at save, so a raw write there is a lie. */
export const RAW_BODY_PART = 'word/document.xml'

/**
 * Parts a raw edit may write. Everything outside it is either regenerated from
 * the model (the body) or has no business being hand-edited (media, embedded
 * workbooks). Deliberately a prefix/exact list rather than "anything XML":
 * the escape hatch exists for the document properties our model does not
 * represent, which is exactly this set.
 */
const WRITABLE = [
  /^word\/styles\.xml$/,
  /^word\/numbering\.xml$/,
  /^word\/settings\.xml$/,
  /^word\/theme\/theme\d*\.xml$/,
  /^word\/(header|footer)\d*\.xml$/,
  /^\[Content_Types\]\.xml$/,
  /^word\/_rels\/document\.xml\.rels$/,
]

export interface RawPartInfo {
  path: string
  ref?: string
  bytes: number
  /** false when the part may be read but not written */
  writable: boolean
}

export type RawResult<T> = ({ ok: true } & T) | { ok: false; error: string }

const SHORT_NAMES: Record<string, string> = {
  'word/document.xml': '/document',
  'word/styles.xml': '/styles',
  'word/numbering.xml': '/numbering',
  'word/settings.xml': '/settings',
  'word/theme/theme1.xml': '/theme',
  'word/comments.xml': '/comments',
  '[Content_Types].xml': '/contentTypes',
}

export function isRawWritable(path: string): boolean {
  return path !== RAW_BODY_PART && WRITABLE.some((re) => re.test(path))
}

/** Resolve a short name or a literal entry name against the package. */
export function resolveRawPartPath(names: ReadonlySet<string>, ref: string): string | null {
  const trimmed = ref.trim()
  if (!trimmed) return null
  if (names.has(trimmed)) return trimmed
  const stripped = trimmed.replace(/^\/+/, '')
  if (names.has(stripped)) return stripped
  const short = `/${stripped}`
  for (const [path, name] of Object.entries(SHORT_NAMES)) {
    if (name === short && names.has(path)) return path
  }
  // /header[1] and /footer[2], numbered as the files are — unlike slides there
  // is no document-level order to resolve them through
  const indexed = /^(header|footer)\[(\d+)\]$/.exec(stripped)
  if (indexed) {
    const path = `word/${indexed[1]}${Number(indexed[2])}.xml`
    return names.has(path) ? path : null
  }
  return null
}

/** Open a package for raw work; the caller keeps the zip for the edit. */
export async function openRawPackage(bytes: Uint8Array): Promise<JSZip> {
  return JSZip.loadAsync(bytes)
}

export function listRawParts(zip: JSZip): RawPartInfo[] {
  const out: RawPartInfo[] = []
  for (const [path, file] of Object.entries(zip.files)) {
    if (file.dir) continue
    if (!path.endsWith('.xml') && !path.endsWith('.rels')) continue
    const ref = SHORT_NAMES[path]
    out.push({
      path,
      bytes: 0,
      writable: isRawWritable(path),
      ...(ref ? { ref } : {}),
    })
  }
  return out.sort((a, b) => a.path.localeCompare(b.path))
}

export async function readRawPart(
  zip: JSZip,
  ref: string,
): Promise<RawResult<{ path: string; xml: string }>> {
  const names = new Set(Object.keys(zip.files))
  const path = resolveRawPartPath(names, ref)
  if (!path) return { ok: false, error: `No part matches "${ref}"` }
  const file = zip.file(path)
  if (!file) return { ok: false, error: `Part "${path}" is not in the package` }
  const xml = await file.async('string')
  if (xml.length > RAW_MAX_PART_BYTES) {
    return {
      ok: false,
      error:
        `Part "${path}" is ${Math.round(xml.length / 1024)} KB, over the ` +
        `${Math.round(RAW_MAX_PART_BYTES / 1024)} KB raw limit — use the model tools for a part this size`,
    }
  }
  return { ok: true, path, xml }
}

/**
 * Apply a unique find/replace to a part and prove the document still parses.
 * Returns the bytes to hand to `SaveOptions.partOverrides`; nothing is mutated
 * in place, so a rejected edit leaves the caller's state untouched by
 * construction rather than by remembering to roll back.
 */
export async function editRawPart(
  zip: JSZip,
  originalBytes: Uint8Array,
  ref: string,
  find: string,
  replace: string,
  existingOverrides: ReadonlyMap<string, Uint8Array> = new Map(),
): Promise<RawResult<{ path: string; bytes: Uint8Array }>> {
  const names = new Set(Object.keys(zip.files))
  const path = resolveRawPartPath(names, ref)
  if (!path) return { ok: false, error: `No part matches "${ref}"` }
  if (path === RAW_BODY_PART) {
    return {
      ok: false,
      error:
        `"${path}" is rebuilt from the document's blocks when the file is saved, so a raw edit ` +
        'to it would be discarded. Use apply_commands or replace_blocks for body content; ' +
        'reading this part is still fine.',
    }
  }
  if (!isRawWritable(path)) {
    return {
      ok: false,
      error: `"${path}" is not one of the parts a raw edit may write (styles, numbering, settings, theme, headers, footers, content types, document rels)`,
    }
  }
  if (!find) return { ok: false, error: 'find must not be empty' }

  // an earlier accepted edit is the base for this one, so two edits to the same
  // part compose instead of the second reverting the first
  const pending = existingOverrides.get(path)
  const before =
    pending !== undefined ? new TextDecoder().decode(pending) : await zip.file(path)!.async('string')
  if (before.length > RAW_MAX_PART_BYTES) {
    return { ok: false, error: `Part "${path}" is too large to edit raw` }
  }

  const first = before.indexOf(find)
  if (first < 0) {
    return {
      ok: false,
      error: `find does not occur in "${path}" — read the part first and copy the text exactly`,
    }
  }
  if (before.indexOf(find, first + find.length) >= 0) {
    const count = before.split(find).length - 1
    return {
      ok: false,
      error: `find occurs ${count} times in "${path}"; it must match exactly once — include more surrounding text`,
    }
  }

  const next = before.slice(0, first) + replace + before.slice(first + find.length)
  const valid = XMLValidator.validate(next)
  if (valid !== true) {
    return {
      ok: false,
      error: `The result is not well-formed XML: ${valid.err.msg} (line ${valid.err.line})`,
    }
  }
  const bytes = new TextEncoder().encode(next)

  // Fourth gate: rebuild the package with the edit in place and parse it. The
  // slides version re-parses affected slides; here the whole document is the
  // unit, and styles or numbering can break it as readily as the body.
  const candidate = await openRawPackage(originalBytes)
  for (const [name, content] of existingOverrides) candidate.file(name, content)
  candidate.file(path, bytes)
  try {
    const parsed = await parseDocx(await candidate.generateAsync({ type: 'uint8array' }))
    if (parsed.blocks.length === 0) {
      return { ok: false, error: `The edit left the document with no content; it was discarded` }
    }
  } catch (error: unknown) {
    const reason = error instanceof Error ? error.message : String(error)
    return { ok: false, error: `The edit left the document unparseable, so it was discarded: ${reason}` }
  }
  return { ok: true, path, bytes }
}
