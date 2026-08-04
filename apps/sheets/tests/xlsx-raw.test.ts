/**
 * The raw-OOXML gates that can be tested without a workbook.
 *
 * The point of the find/replace shape is that a bad edit is refused rather
 * than written, so these are mostly tests of refusal. The fourth gate — the
 * workbook must still read — needs the sidecar and lives with the main-process
 * wiring.
 */
import { describe, it, expect } from 'vitest'
import {
  applyRawEdit,
  checkRawPartSize,
  listRawParts,
  resolveRawPartRef,
  sheetPathsInOrder,
  type RawArchiveView,
} from '../src/gateway/xlsx-raw'

const VIEW: RawArchiveView = {
  entries: [
    { name: '[Content_Types].xml', uncompressedSize: 1200 },
    { name: 'xl/workbook.xml', uncompressedSize: 900 },
    { name: 'xl/_rels/workbook.xml.rels', uncompressedSize: 600 },
    { name: 'xl/styles.xml', uncompressedSize: 4000 },
    { name: 'xl/sharedStrings.xml', uncompressedSize: 300 },
    { name: 'xl/theme/theme1.xml', uncompressedSize: 7000 },
    { name: 'xl/worksheets/sheet1.xml', uncompressedSize: 2000 },
    { name: 'xl/worksheets/sheet2.xml', uncompressedSize: 2500 },
    { name: 'xl/media/image1.png', uncompressedSize: 50_000 },
  ],
  // deliberately not file order: sheet2 is the first tab
  sheetPaths: ['xl/worksheets/sheet2.xml', 'xl/worksheets/sheet1.xml'],
}

describe('sheetPathsInOrder', () => {
  const RELS =
    '<Relationships>' +
    '<Relationship Id="rId1" Target="worksheets/sheet3.xml"/>' +
    '<Relationship Target="worksheets/sheet1.xml" Id="rId2"/>' +
    '</Relationships>'

  it('follows the workbook sheet order, not the file numbering', () => {
    const wb =
      '<workbook><sheets>' +
      '<sheet name="Summary" sheetId="7" r:id="rId1"/>' +
      '<sheet name="Data" sheetId="3" r:id="rId2"/>' +
      '</sheets></workbook>'
    // sheet3.xml is the first tab; numbering would have said sheet1
    expect(sheetPathsInOrder(wb, RELS)).toEqual([
      'xl/worksheets/sheet3.xml',
      'xl/worksheets/sheet1.xml',
    ])
  })

  it('reads Target before Id, because producers disagree on attribute order', () => {
    const wb = '<workbook><sheets><sheet name="Data" r:id="rId2"/></sheets></workbook>'
    expect(sheetPathsInOrder(wb, RELS)).toEqual(['xl/worksheets/sheet1.xml'])
  })

  it('handles an absolute target', () => {
    const wb = '<workbook><sheets><sheet name="S" r:id="rId9"/></sheets></workbook>'
    const rels = '<Relationships><Relationship Id="rId9" Target="/xl/worksheets/sheetA.xml"/></Relationships>'
    expect(sheetPathsInOrder(wb, rels)).toEqual(['xl/worksheets/sheetA.xml'])
  })
})

describe('resolveRawPartRef', () => {
  it('resolves a sheet index through workbook order', () => {
    expect(resolveRawPartRef(VIEW, '/sheet[1]')).toBe('xl/worksheets/sheet2.xml')
    expect(resolveRawPartRef(VIEW, 'sheet[2]')).toBe('xl/worksheets/sheet1.xml')
  })

  it('resolves the named singletons', () => {
    expect(resolveRawPartRef(VIEW, '/workbook')).toBe('xl/workbook.xml')
    expect(resolveRawPartRef(VIEW, '/styles')).toBe('xl/styles.xml')
    expect(resolveRawPartRef(VIEW, '/sharedStrings')).toBe('xl/sharedStrings.xml')
    expect(resolveRawPartRef(VIEW, '/theme')).toBe('xl/theme/theme1.xml')
    expect(resolveRawPartRef(VIEW, '/contentTypes')).toBe('[Content_Types].xml')
  })

  it('takes a literal entry name, so a rels target can be pasted straight in', () => {
    expect(resolveRawPartRef(VIEW, 'xl/_rels/workbook.xml.rels')).toBe('xl/_rels/workbook.xml.rels')
    expect(resolveRawPartRef(VIEW, '[Content_Types].xml')).toBe('[Content_Types].xml')
  })

  it('returns null rather than guessing', () => {
    expect(resolveRawPartRef(VIEW, '/sheet[3]')).toBeNull()
    expect(resolveRawPartRef(VIEW, '/sheet[0]')).toBeNull()
    expect(resolveRawPartRef(VIEW, '/nope')).toBeNull()
    expect(resolveRawPartRef(VIEW, '')).toBeNull()
    expect(resolveRawPartRef(VIEW, 'xl/worksheets/sheet9.xml')).toBeNull()
  })
})

describe('listRawParts', () => {
  it('lists XML and rels only, with short names where they exist', () => {
    const parts = listRawParts(VIEW)
    expect(parts.some((p) => p.path.endsWith('.png'))).toBe(false)
    expect(parts.find((p) => p.path === 'xl/worksheets/sheet2.xml')?.ref).toBe('/sheet[1]')
    expect(parts.find((p) => p.path === 'xl/styles.xml')?.ref).toBe('/styles')
    // a part with no short name is still listed, by its entry name
    expect(parts.find((p) => p.path === 'xl/_rels/workbook.xml.rels')?.ref).toBeUndefined()
    expect(parts.map((p) => p.path)).toEqual([...parts.map((p) => p.path)].sort())
  })
})

describe('applyRawEdit', () => {
  const XML = '<root><a v="1"/><b v="2"/><a v="3"/></root>'

  it('replaces the one occurrence it was given', () => {
    const r = applyRawEdit('p.xml', XML, '<b v="2"/>', '<b v="9"/>')
    expect(r.ok && r.xml).toContain('<b v="9"/>')
  })

  it('refuses an ambiguous find rather than editing five places', () => {
    const r = applyRawEdit('p.xml', XML, '<a', '<z')
    expect(r.ok).toBe(false)
    expect(!r.ok && r.error).toContain('occurs 2 times')
  })

  it('refuses a find that is not there, and says to copy it exactly', () => {
    const r = applyRawEdit('p.xml', XML, '<missing/>', '')
    expect(r.ok).toBe(false)
    expect(!r.ok && r.error).toContain('copy the text exactly')
  })

  it('refuses an empty find, which would match at position zero', () => {
    expect(applyRawEdit('p.xml', XML, '', 'x').ok).toBe(false)
  })

  it('refuses a replacement that breaks well-formedness', () => {
    const r = applyRawEdit('p.xml', XML, '<b v="2"/>', '<b v="2">')
    expect(r.ok).toBe(false)
    expect(!r.ok && r.error).toContain('not well-formed')
  })

  it('allows a deletion, which is a replacement with nothing', () => {
    const r = applyRawEdit('p.xml', XML, '<b v="2"/>', '')
    expect(r.ok && r.xml).toBe('<root><a v="1"/><a v="3"/></root>')
  })
})

describe('checkRawPartSize', () => {
  it('passes a part under the cap and names the cap when over', () => {
    expect(checkRawPartSize('xl/styles.xml', 1000)).toBeNull()
    expect(checkRawPartSize('xl/styles.xml', 500_000)).toContain('over the 391 KB raw limit')
  })
})
