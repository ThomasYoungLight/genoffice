/**
 * Raw OOXML access.
 *
 * The point of these is the refusals. A raw edit is the one operation here that
 * can produce a file PowerPoint will not open, so what matters is not that a
 * good edit lands — it is that an ambiguous match, malformed output or a part
 * that stops parsing all leave the document exactly as it was, and say why.
 */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  editRawPart,
  listRawParts,
  openPptx,
  readRawPart,
  resolvePartPath,
  savePptx,
} from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))
const open = () => openPptx(fx('01_standard_business.pptx'))

const ok = <T>(r: { ok: true } & T): T => r
const xmlOf = (r: ReturnType<typeof readRawPart>) => {
  if (!r.ok) throw new Error(r.error)
  return r.xml
}

describe('resolvePartPath', () => {
  it('resolves short names through the presentation order, not the file numbering', async () => {
    const opened = await open()
    const { slidePaths } = opened.archive.readPresentation()
    expect(resolvePartPath(opened, '/slide[1]')).toBe(slidePaths[0])
    expect(resolvePartPath(opened, '/slide[2]')).toBe(slidePaths[1])
    expect(resolvePartPath(opened, '/presentation')).toBe('ppt/presentation.xml')
  })

  it('accepts a literal zip path', async () => {
    const opened = await open()
    expect(resolvePartPath(opened, 'ppt/presentation.xml')).toBe('ppt/presentation.xml')
    expect(resolvePartPath(opened, '/ppt/presentation.xml')).toBe('ppt/presentation.xml')
  })

  it('returns null for a part that is not there', async () => {
    const opened = await open()
    expect(resolvePartPath(opened, '/slide[999]')).toBeNull()
    expect(resolvePartPath(opened, '/nonsense')).toBeNull()
    expect(resolvePartPath(opened, 'ppt/slides/slide404.xml')).toBeNull()
  })
})

describe('listRawParts', () => {
  it('names the parts that have short names and includes the rest', async () => {
    const opened = await open()
    const parts = listRawParts(opened)
    expect(parts.find((p) => p.path === 'ppt/presentation.xml')?.ref).toBe('/presentation')
    expect(parts.find((p) => p.ref === '/slide[1]')).toBeTruthy()
    expect(parts.some((p) => p.path.endsWith('.rels'))).toBe(true)
    expect(parts.every((p) => p.bytes > 0)).toBe(true)
  })
})

describe('readRawPart', () => {
  it('returns the part text', async () => {
    const opened = await open()
    const xml = xmlOf(readRawPart(opened, '/slide[1]'))
    expect(xml).toContain('<p:sld')
    expect(xml).toContain('spTree')
  })

  it('names the part it could not find', async () => {
    const opened = await open()
    const result = readRawPart(opened, '/slide[99]')
    expect(result.ok).toBe(false)
    expect(!result.ok && result.error).toContain('/slide[99]')
  })
})

describe('editRawPart', () => {
  it('applies a unique match and reparses the slide', async () => {
    const opened = await open()
    const before = xmlOf(readRawPart(opened, '/slide[1]'))
    const marker = /<a:t>([^<]{4,})<\/a:t>/.exec(before)
    const original = marker![1]!
    const result = editRawPart(
      opened,
      '/slide[1]',
      `<a:t>${original}</a:t>`,
      '<a:t>Edited raw</a:t>',
    )
    expect(ok(result as { ok: true; reparsedSlides: number }).reparsedSlides).toBe(1)
    expect(xmlOf(readRawPart(opened, '/slide[1]'))).toContain('Edited raw')
    // the parsed model followed the bytes, rather than still describing the old text
    const texts = opened.deck.slides[0]!.elements.flatMap((e) =>
      'text' in e
        ? ((e.text?.paragraphs ?? []).flatMap((p) => p.runs.map((r) => r.text)) ?? [])
        : [],
    )
    expect(texts.join(' ')).toContain('Edited raw')
  })

  it('survives a save/reopen round trip', async () => {
    const opened = await open()
    const before = xmlOf(readRawPart(opened, '/slide[1]'))
    const original = /<a:t>([^<]{4,})<\/a:t>/.exec(before)![1]!
    editRawPart(opened, '/slide[1]', `<a:t>${original}</a:t>`, '<a:t>Round tripped</a:t>')
    const reopened = await openPptx(await savePptx(opened))
    expect(xmlOf(readRawPart(reopened, '/slide[1]'))).toContain('Round tripped')
  })

  it('refuses a match that occurs more than once, and counts them', async () => {
    const opened = await open()
    const result = editRawPart(opened, '/slide[1]', '<a:t>', '<a:t>x')
    expect(result.ok).toBe(false)
    expect(!result.ok && result.error).toMatch(/occurs \d+ times/)
  })

  it('refuses a match that is not there', async () => {
    const opened = await open()
    const result = editRawPart(opened, '/slide[1]', '<not:in-this-file/>', '')
    expect(result.ok).toBe(false)
    expect(!result.ok && result.error).toContain('does not occur')
  })

  it('refuses an empty find rather than inserting at position zero', async () => {
    const opened = await open()
    expect(editRawPart(opened, '/slide[1]', '', '<junk/>').ok).toBe(false)
  })

  it('rejects malformed output and leaves the part untouched', async () => {
    const opened = await open()
    const before = xmlOf(readRawPart(opened, '/slide[1]'))
    const result = editRawPart(opened, '/slide[1]', '</p:sld>', '<p:unclosed>')
    expect(result.ok).toBe(false)
    expect(!result.ok && result.error).toContain('well-formed')
    expect(xmlOf(readRawPart(opened, '/slide[1]'))).toBe(before)
  })

  it('restores the original bytes when the edit leaves a slide unparseable', async () => {
    const opened = await open()
    const before = xmlOf(readRawPart(opened, '/slide[1]'))
    // well-formed, so validation passes, but the slide loses its shape tree
    const spTree = /<p:spTree>[\s\S]*<\/p:spTree>/.exec(before)
    const result = spTree
      ? editRawPart(opened, '/slide[1]', spTree[0], '<p:notASpTree/>')
      : { ok: false as const, error: 'no spTree' }
    expect(result.ok).toBe(false)
    expect(xmlOf(readRawPart(opened, '/slide[1]'))).toBe(before)
    expect(opened.deck.slides[0]!.elements.length).toBeGreaterThan(0)
  })

  it('reparses every slide when a shared part changes', async () => {
    const opened = await open()
    const themePath = resolvePartPath(opened, '/theme')
    expect(themePath).toBeTruthy()
    const theme = xmlOf(readRawPart(opened, '/theme'))
    const accent = /<a:srgbClr val="([0-9A-Fa-f]{6})"\/><\/a:accent1>/.exec(theme)
    const result = accent
      ? editRawPart(opened, '/theme', accent[0], '<a:srgbClr val="FF0000"/></a:accent1>')
      : editRawPart(opened, '/theme', '<a:accent1>', '<a:accent1>')
    if (accent) {
      expect(ok(result as { ok: true; reparsedSlides: number }).reparsedSlides).toBe(
        opened.deck.slides.length,
      )
      expect(xmlOf(readRawPart(opened, '/theme'))).toContain('FF0000')
    }
  })

  it('names a part it cannot resolve instead of editing something else', async () => {
    const opened = await open()
    const result = editRawPart(opened, '/slide[99]', 'a', 'b')
    expect(result.ok).toBe(false)
    expect(!result.ok && result.error).toContain('/slide[99]')
  })
})
