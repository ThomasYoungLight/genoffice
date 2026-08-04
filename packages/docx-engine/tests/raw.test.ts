/**
 * Raw OOXML access for documents.
 *
 * The gate that has no counterpart in slides is the one worth most of this
 * file: `word/document.xml` is readable but not writable, because saveDocx
 * rebuilds its body from the editor's blocks. Without the refusal the tool
 * would report success and change nothing — a silent no-op is the one outcome
 * an escape hatch must never produce.
 */
import { describe, it, expect } from 'vitest'
import {
  buildBlankDocx,
  editRawPart,
  isRawWritable,
  listRawParts,
  openRawPackage,
  parseDocx,
  readRawPart,
  resolveRawPartPath,
  saveDocx,
  type SaveBlock,
} from '../src/index'

const blank = async () => {
  const bytes = await buildBlankDocx()
  return { bytes, zip: await openRawPackage(bytes) }
}
const originalOrder = (doc: Awaited<ReturnType<typeof parseDocx>>): SaveBlock[] =>
  doc.blocks
    .filter((b) => !b.hidden && b.docxIndex !== null)
    .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))

describe('part resolution', () => {
  it('resolves short names and literal entry names', async () => {
    const { zip } = await blank()
    const names = new Set(Object.keys(zip.files))
    expect(resolveRawPartPath(names, '/styles')).toBe('word/styles.xml')
    expect(resolveRawPartPath(names, '/document')).toBe('word/document.xml')
    expect(resolveRawPartPath(names, 'word/settings.xml')).toBe('word/settings.xml')
    expect(resolveRawPartPath(names, '/nope')).toBeNull()
  })

  it('marks the body readable but not writable', async () => {
    const { zip } = await blank()
    const parts = listRawParts(zip)
    expect(parts.find((p) => p.path === 'word/document.xml')?.writable).toBe(false)
    expect(parts.find((p) => p.path === 'word/styles.xml')?.writable).toBe(true)
    expect(isRawWritable('word/media/image1.png')).toBe(false)
  })
})

describe('reads', () => {
  it('reads the body even though it cannot be written', async () => {
    const { zip } = await blank()
    const result = await readRawPart(zip, '/document')
    expect(result.ok).toBe(true)
    expect(result.ok && result.xml).toContain('<w:body>')
  })
})

describe('edits', () => {
  it('refuses the body and names the tools that own it', async () => {
    const { bytes, zip } = await blank()
    const result = await editRawPart(zip, bytes, '/document', '<w:body>', '<w:body>')
    expect(result.ok).toBe(false)
    expect(!result.ok && result.error).toContain('rebuilt from')
    expect(!result.ok && result.error).toContain('apply_commands')
  })

  it('refuses a part outside the writable set', async () => {
    const { bytes, zip } = await blank()
    const result = await editRawPart(zip, bytes, 'docProps/core.xml', 'x', 'y')
    expect(result.ok).toBe(false)
  })

  it('refuses an ambiguous find', async () => {
    const { bytes, zip } = await blank()
    const result = await editRawPart(zip, bytes, '/settings', '<w:', '<w:')
    expect(result.ok).toBe(false)
    expect(!result.ok && result.error).toMatch(/occurs \d+ times/)
  })

  it('refuses a replacement that is not well-formed', async () => {
    const { bytes, zip } = await blank()
    const result = await editRawPart(
      zip,
      bytes,
      '/settings',
      '<w:compat>',
      '<w:compat><w:unclosed>',
    )
    expect(result.ok).toBe(false)
    expect(!result.ok && result.error).toContain('not well-formed')
  })

  it('accepts a valid edit and the override survives the save', async () => {
    const { bytes, zip } = await blank()
    const result = await editRawPart(
      zip,
      bytes,
      '/settings',
      '<w:compat>',
      '<w:compat><w:doNotExpandShiftReturn/>',
    )
    expect(result.ok).toBe(true)
    if (!result.ok) return

    const parsed = await parseDocx(bytes)
    const saved = await saveDocx(parsed, originalOrder(parsed), {
      partOverrides: new Map([[result.path, result.bytes]]),
    })
    const out = await openRawPackage(saved)
    expect(await out.file('word/settings.xml')!.async('string')).toContain(
      '<w:doNotExpandShiftReturn/>',
    )
    // the document still parses after the round trip
    expect((await parseDocx(saved)).blocks.length).toBeGreaterThan(0)
  })

  it('composes a second edit on top of the first', async () => {
    const { bytes, zip } = await blank()
    const first = await editRawPart(zip, bytes, '/settings', '<w:compat>', '<w:compat><w:a/>')
    expect(first.ok).toBe(true)
    if (!first.ok) return
    const overrides = new Map([[first.path, first.bytes]])
    const second = await editRawPart(zip, bytes, '/settings', '<w:a/>', '<w:a/><w:b/>', overrides)
    expect(second.ok).toBe(true)
    if (!second.ok) return
    const text = new TextDecoder().decode(second.bytes)
    expect(text).toContain('<w:a/>')
    expect(text).toContain('<w:b/>')
  })

  it('never lets an override reach the body, even if one is handed to the save', async () => {
    // belt and braces: the tool refuses first, but the save must not honour it
    // either, or a caller reaching saveDocx directly could still corrupt itself
    const bytes = await buildBlankDocx()
    const parsed = await parseDocx(bytes)
    const saved = await saveDocx(parsed, originalOrder(parsed), {
      partOverrides: new Map([['word/document.xml', new TextEncoder().encode('<nonsense/>')]]),
    })
    expect((await parseDocx(saved)).blocks.length).toBeGreaterThan(0)
  })
})
