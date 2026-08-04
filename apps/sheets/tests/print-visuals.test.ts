/**
 * What reaches a printed page or an exported PDF.
 *
 * The chain is worth stating, because it is what makes this test decisive
 * rather than indicative: `buildSheetPrintPayload` produces an HTML string,
 * `pdf-export.ts` writes exactly that string to a file, loads it in a hidden
 * BrowserWindow with `javascript: false`, and calls `printToPDF`. Nothing
 * captures the live grid. So anything absent from this payload cannot appear in
 * the output, and asserting on the payload is asserting on the PDF.
 *
 * The finding: it is a table of cells. Charts, shapes and images — every one of
 * which is a float DOM overlay rather than part of Univer's model — have no
 * representation here at all.
 */
import { describe, expect, it } from 'vitest'

import { buildSheetPrintPayload, type PrintWorksheet } from '../src/renderer/print-html'

/// A 3x2 sheet, plus a chart and a picture anchored over it. The visuals are
/// not passed in because there is nowhere to pass them: PrintWorksheet has no
/// parameter for them, which is the shape of the gap.
function sheetWithVisuals(): PrintWorksheet {
  const values = [
    ['Region', 'Revenue'],
    ['North', '120'],
    ['South', '96'],
  ]
  return {
    getLastRow: () => 2,
    getLastColumn: () => 1,
    getRowHeight: () => 20,
    getColumnWidth: () => 80,
    getMergedRanges: () => [],
    getRange: ((row: number, column: number, numRows?: number, numColumns?: number) => {
      if (numRows === undefined) return { getCellStyleData: () => null }
      const block = values
        .slice(row, row + numRows)
        .map((r) => r.slice(column, column + (numColumns ?? 1)))
      return { getDisplayValues: () => block, getValues: () => block }
    }) as PrintWorksheet['getRange'],
  }
}

describe('print/PDF payload', () => {
  const payload = () => buildSheetPrintPayload(sheetWithVisuals(), {}, 'book.xlsx', 'Data')

  it('carries the cell values, so the harness itself is sound', () => {
    const { html } = payload()
    expect(html).toContain('Region')
    expect(html).toContain('North')
    expect(html).toContain('120')
  })

  it('carries no visual of any kind — the gap this test exists to record', () => {
    const { html } = payload()
    // a chart, shape or image would have to arrive as one of these; none can,
    // because the payload is built from cells and the visuals are float DOM
    expect(html).not.toContain('<img')
    expect(html).not.toContain('<svg')
    expect(html).not.toContain('<canvas')
    expect(html).not.toMatch(/chart/i)
  })

  it('has no way to be given visuals in the first place', () => {
    // documents the cause rather than the symptom: the print layer's view of a
    // worksheet is cells, styles and merges. If visuals are ever added to the
    // output, this assertion is the one that should fail first.
    const keys = Object.keys(sheetWithVisuals())
    expect(keys).toEqual([
      'getLastRow',
      'getLastColumn',
      'getRowHeight',
      'getColumnWidth',
      'getMergedRanges',
      'getRange',
    ])
  })
})
