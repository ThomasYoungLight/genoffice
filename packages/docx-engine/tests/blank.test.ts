/**
 * The blank document we hand to Word.
 *
 * Word decides a document's feature level from the `compatibilityMode` compat
 * setting. A package with no settings part is not invalid and does not warn —
 * Word simply assumes an old format, disables newer features and writes
 * "Compatibility Mode" in the title bar. Every document this product created
 * did that, and no test noticed because nothing here had ever been opened in
 * Word.
 */
import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { buildBlankDocx, parseDocx } from '../src/index'

const openBlank = async () => JSZip.loadAsync(await buildBlankDocx())
const textOf = async (zip: JSZip, path: string) => (await zip.file(path)?.async('string')) ?? ''

describe('buildBlankDocx', () => {
  it('declares a modern compatibility mode, so Word does not downgrade the document', async () => {
    const settings = await textOf(await openBlank(), 'word/settings.xml')
    expect(settings).toContain('<w:compat>')
    expect(settings).toContain('w:name="compatibilityMode"')
    expect(settings).toContain('w:val="15"')
  })

  it('registers the settings part in both places a part has to be registered', async () => {
    // a part that is written but not declared is invisible to Word, and a
    // declared part that is missing makes the package invalid
    const zip = await openBlank()
    expect(await textOf(zip, '[Content_Types].xml')).toContain('PartName="/word/settings.xml"')
    expect(await textOf(zip, 'word/_rels/document.xml.rels')).toContain('Target="settings.xml"')
  })

  it('still parses', async () => {
    const parsed = await parseDocx(await buildBlankDocx())
    expect(parsed.blocks.length).toBeGreaterThan(0)
  })
})
