/**
 * The pure and lightly-coupled helpers inside univer-sync.ts.
 *
 * The file is the largest in the application layer and most of it is bound to
 * a live Univer instance, but these pieces are not, and they carry decisions
 * that are wrong in ways nothing downstream would catch: a link target that
 * quietly becomes a relative path, a column width converted with the wrong
 * rounding, a protect guard that lets an unindexed sheet through.
 */
import { describe, expect, it } from 'vitest'

import {
  cellValueBounds,
  characterWidthToPixels,
  nextSessionPivotName,
  nextSessionTableName,
  normalizeLinkTarget,
  protectSheetGuard,
} from '../src/renderer/univer-sync'
import { fakeLazyState } from './helpers/fake-univer'

describe('normalizeLinkTarget', () => {
  it.each([
    ['https://example.com', 'https://example.com'],
    ['http://example.com/a?b=c', 'http://example.com/a?b=c'],
    ['mailto:someone@example.com', 'mailto:someone@example.com'],
    ['HTTPS://EXAMPLE.COM', 'HTTPS://EXAMPLE.COM'],
  ])('passes an explicit scheme through unchanged: %s', (input, expected) => {
    expect(normalizeLinkTarget(input)).toBe(expected)
  })

  it('promotes a bare domain to https rather than leaving it relative', () => {
    // Left alone, Excel resolves "example.com" against the file's own
    // location, producing a link to a folder that does not exist.
    expect(normalizeLinkTarget('example.com')).toBe('https://example.com')
    expect(normalizeLinkTarget('docs.example.co.uk/guide')).toBe(
      'https://docs.example.co.uk/guide',
    )
  })

  it('marks an in-workbook reference with # so it is not treated as a URL', () => {
    expect(normalizeLinkTarget('Sheet1!A1')).toBe('#Sheet1!A1')
    expect(normalizeLinkTarget('#Sheet1!A1')).toBe('#Sheet1!A1')
  })

  it('handles a quoted sheet name, which is how a name with spaces arrives', () => {
    expect(normalizeLinkTarget("'My Sheet'!B2")).toBe("#'My Sheet'!B2")
  })

  it('trims surrounding whitespace before deciding', () => {
    expect(normalizeLinkTarget('  https://example.com  ')).toBe('https://example.com')
  })

  it.each([
    ['empty', ''],
    ['whitespace only', '   '],
    ['a bare word', 'notalink'],
    ['a sentence', 'click here for the report'],
    ['over the 2083-character limit', `https://e.com/${'a'.repeat(2100)}`],
  ])('refuses %s rather than guessing', (_label, input) => {
    expect(normalizeLinkTarget(input)).toBeNull()
  })
})

describe('characterWidthToPixels', () => {
  it('maps zero width to zero, not to the padding constant', () => {
    // A hidden column is width 0; adding padding would make it reappear.
    expect(characterWidthToPixels(0)).toBe(0)
  })

  it('adds the cell padding for any non-zero width', () => {
    expect(characterWidthToPixels(1)).toBeGreaterThan(5)
  })

  it('matches Excel for the default width of 8.43 characters', () => {
    // Excel's own default column is 64px; the conversion has to land there or
    // every unset column drifts.
    expect(characterWidthToPixels(8.43)).toBe(64)
  })

  it('increases monotonically with width', () => {
    const widths = [1, 2, 5, 10, 20, 50].map(characterWidthToPixels)
    expect(widths).toEqual([...widths].sort((a, b) => a - b))
  })
})

describe('session naming', () => {
  it('starts at Table1 and Pivot1 on an empty journal', () => {
    const journal = fakeLazyState().editJournal as never
    expect(nextSessionTableName(journal)).toBe('Table1')
    expect(nextSessionPivotName(journal)).toBe('Pivot1')
  })

  it('skips a number already taken, case-insensitively', () => {
    // Excel table names are case-insensitive; handing out "Table2" when
    // "table2" exists produces a file that opens with a repair prompt.
    const state = fakeLazyState()
    ;(state.editJournal.tableAdds as unknown[]).push({ name: 'table2' })
    expect(nextSessionTableName(state.editJournal as never)).not.toBe('table2')
  })

  it('keeps handing out fresh names as the journal grows', () => {
    const state = fakeLazyState()
    const seen = new Set<string>()
    for (let i = 0; i < 5; i += 1) {
      const name = nextSessionTableName(state.editJournal as never)
      expect(seen.has(name.toLowerCase())).toBe(false)
      seen.add(name.toLowerCase())
      ;(state.editJournal.tableAdds as unknown[]).push({ name })
    }
  })
})

describe('protectSheetGuard', () => {
  it('refuses a sheet that was removed this session', () => {
    const state = fakeLazyState()
    state.editJournal.sheets.removed.add('sheet-1')
    expect(protectSheetGuard(state as never, 'sheet-1', true)).toMatch(/unknown sheet/i)
  })

  it('refuses a sheet that has not finished indexing', () => {
    // Protection state comes from the file; without it we would be guessing at
    // whether a password is already set.
    const state = fakeLazyState()
    expect(protectSheetGuard(state as never, 'sheet-1', true)).toBeTruthy()
  })

  it('allows protecting a sheet added this session, which has no file state', () => {
    const state = fakeLazyState()
    state.editJournal.sheets.added.add('sheet-1')
    expect(protectSheetGuard(state as never, 'sheet-1', true)).toBeNull()
  })

  it('allows protecting an indexed sheet', () => {
    const state = fakeLazyState()
    state.sheetProtections.set('sheet-1', { protected: false, hasPassword: false })
    expect(protectSheetGuard(state as never, 'sheet-1', true)).toBeNull()
  })

  it('refuses to unprotect a sheet whose password we do not have', () => {
    // Removing protection we cannot authenticate would silently drop the
    // password from the saved file.
    const state = fakeLazyState()
    state.sheetProtections.set('sheet-1', { protected: true, hasPassword: true })
    expect(protectSheetGuard(state as never, 'sheet-1', false)).toBeTruthy()
  })

  it('allows unprotecting a sheet with no password', () => {
    const state = fakeLazyState()
    state.sheetProtections.set('sheet-1', { protected: true, hasPassword: false })
    expect(protectSheetGuard(state as never, 'sheet-1', false)).toBeNull()
  })
})

describe('cellValueBounds', () => {
  it('returns null for anything that is not a cell matrix', () => {
    for (const input of [null, undefined, 42, 'text', true]) {
      expect(cellValueBounds(input)).toBeNull()
    }
  })

  it('returns null for an empty matrix', () => {
    expect(cellValueBounds({})).toBeNull()
  })

  it('finds the extent of a sparse matrix', () => {
    const bounds = cellValueBounds({ 0: { 0: { v: 1 } }, 5: { 3: { v: 2 } } })
    expect(bounds).toMatchObject({ startRow: 0, endRow: 5, startColumn: 0, endColumn: 3 })
  })

  it('does not assume the matrix starts at the origin', () => {
    const bounds = cellValueBounds({ 10: { 4: { v: 1 } } })
    expect(bounds).toMatchObject({ startRow: 10, endRow: 10, startColumn: 4, endColumn: 4 })
  })
})
