/**
 * The read-only views the AI skill layer is built on: what the agent is told
 * about a workbook before it edits one.
 *
 * These matter more than "read helpers" suggests. Every write the agent makes
 * is reasoned from this output, so a wrong or missing line here becomes a
 * wrong edit later, and the failure surfaces far from its cause. They had no
 * tests; the file sat at 0%.
 *
 * The functions were extracted from App.tsx specifically so their runtime
 * arrives as an argument, which is what makes this possible without Univer.
 */
import { describe, expect, it } from 'vitest'

import {
  getActiveSheetInfo,
  readCells,
  readFormats,
  readSheetFeatures,
} from '../src/renderer/ai/workbook-readers'
import {
  fakeLazyState,
  fakeRuntime,
  fakeSheet,
  fakeWorkbook,
  ref,
  type FakeSheet,
} from './helpers/fake-univer'

/// Lazy mode by default: that is what an opened file uses, and it is the path
/// the agent actually reads through.
function context(sheets?: FakeSheet[]) {
  const list = sheets ?? [fakeSheet()]
  const workbook = fakeWorkbook(list)
  return {
    univerRef: ref(fakeRuntime(workbook)),
    lazyWorkbookRef: ref(
      fakeLazyState(list.map((s) => ({ id: s.getSheetId(), name: s.getSheetName() }))),
    ),
    adapterRef: ref({ getSnapshot: () => ({ sheets: [] }) }),
  } as never
}

describe('readSheetFeatures', () => {
  it('says so plainly when no workbook is open', () => {
    const ctx = { univerRef: ref(fakeRuntime(null)), lazyWorkbookRef: ref(null) } as never
    expect(readSheetFeatures(ctx)).toBe('No workbook is currently open.')
  })

  it('names the unknown sheet rather than falling back to the active one', () => {
    // Silently reading a different sheet would be worse than an error: the
    // agent would act on the wrong data and never know.
    expect(readSheetFeatures(context(), 'sheet-missing')).toBe('Unknown sheet: sheet-missing')
  })

  it('reports the sheet it did read, by name and id', () => {
    const report = readSheetFeatures(context([fakeSheet({ id: 's7', name: 'Budget' })]))
    expect(report).toContain('Budget')
    expect(report).toContain('s7')
  })

  it('reports "none" for absent features instead of omitting the section', () => {
    // An omitted section reads as "unknown"; an explicit none is actionable.
    const report = readSheetFeatures(context())
    expect(report).toMatch(/AutoFilter/i)
    expect(report.toLowerCase()).toContain('none')
  })

  it('reports an AutoFilter with its range when one exists', () => {
    const sheet = fakeSheet({ filter: { startRow: 0, startColumn: 0, endRow: 6, endColumn: 2 } })
    expect(readSheetFeatures(context([sheet]))).toMatch(/A1:C7/)
  })

  it('reports freeze panes', () => {
    const sheet = fakeSheet({ freeze: { xSplit: 1, ySplit: 2 } })
    expect(readSheetFeatures(context([sheet]))).toMatch(/frozen/i)
  })

  it('reports merged ranges through the active-sheet view', () => {
    const sheet = fakeSheet({ merges: [{ row: 0, column: 0, width: 2, height: 1 }] })
    expect(getActiveSheetInfo(context([sheet])).merges).toHaveLength(1)
  })

  it('survives a facade method that throws, and still reports the rest', () => {
    // The file's own contract: "each section is independent so one facade
    // failure never hides the rest".
    const sheet = fakeSheet()
    const broken = {
      ...sheet,
      getMergedRanges: () => {
        throw new Error('univer exploded')
      },
    } as unknown as FakeSheet
    const report = readSheetFeatures(context([broken]))
    expect(report).toContain('Sheet1')
    expect(report.length).toBeGreaterThan(20)
  })
})

describe('getActiveSheetInfo', () => {
  it('reports mode none when nothing is open', () => {
    const ctx = {
      univerRef: ref(fakeRuntime(null)),
      lazyWorkbookRef: ref(null),
      adapterRef: ref({ getSnapshot: () => ({ sheets: [] }) }),
    } as never
    expect(getActiveSheetInfo(ctx).mode).toBe('none')
  })

  it('lists every sheet, not just the active one', () => {
    const info = getActiveSheetInfo(
      context([fakeSheet({ id: 'a', name: 'One' }), fakeSheet({ id: 'b', name: 'Two' })]),
    )
    expect(info.sheets.map((s) => s.name)).toEqual(['One', 'Two'])
  })

  it('carries the current selection so the agent can act on "this range"', () => {
    expect(getActiveSheetInfo(context()).selection).toBeTruthy()
  })
})

describe('readCells', () => {
  it('returns nothing for an empty request rather than the whole sheet', () => {
    expect(readCells(context(), [])).toEqual({})
  })

  it('reads scattered addresses independently', () => {
    const sheet = fakeSheet({ cells: { '0:0': 'Region', '1:1': 120 } })
    const cells = readCells(context([sheet]), ['A1', 'B2'])
    expect(cells['A1']?.value).toBe('Region')
    expect(cells['B2']?.value).toBe(120)
  })

  it('carries a formula alongside its value when the cell has one', () => {
    const sheet = fakeSheet({ cells: { '0:0': 3 }, formulas: { '0:0': '=1+2' } })
    expect(readCells(context([sheet]), ['A1'])['A1']).toMatchObject({ formula: '=1+2' })
  })

  it('does not invent a formula key for a plain value', () => {
    const sheet = fakeSheet({ cells: { '0:0': 3 } })
    expect(readCells(context([sheet]), ['A1'])['A1']).not.toHaveProperty('formula')
  })
})

describe('readFormats', () => {
  it('omits cells that carry no formatting, rather than padding the reply', () => {
    // The agent pays for every token of this; a page of "no formatting" entries
    // is worse than silence.
    expect(readFormats(context(), ['A1', 'B2'])).toEqual({})
  })

  it('reports the formatting a cell does carry', () => {
    const sheet = fakeSheet({ styles: { '0:0': { bl: 1, it: 1 } } })
    expect(readFormats(context([sheet]), ['A1'])['A1']).toMatchObject({
      bold: true,
      italic: true,
    })
  })

  it('reports a number format', () => {
    const sheet = fakeSheet({ numberFormats: { '0:0': '0.00%' } })
    expect(readFormats(context([sheet]), ['A1'])['A1']).toMatchObject({ numberFormat: '0.00%' })
  })

  it('returns nothing when no workbook is open', () => {
    const ctx = {
      univerRef: ref(fakeRuntime(null)),
      lazyWorkbookRef: ref(null),
      adapterRef: ref({ getSnapshot: () => ({ sheets: [] }) }),
    } as never
    expect(readFormats(ctx, ['A1'])).toEqual({})
  })
})
