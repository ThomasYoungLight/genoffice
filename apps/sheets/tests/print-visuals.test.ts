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
 * This file used to record the gap: charts, shapes and images are float DOM
 * overlays rather than part of Univer's model, and had no representation here
 * at all. They now arrive as a `PrintVisualLayer` — markup captured from the
 * live DOM plus a twoCellAnchor — and hang off the td for their marker cell,
 * so Chromium's pagination carries each one to the page its cell lands on.
 */
import { describe, expect, it } from 'vitest'

import {
  buildSheetPrintPayload,
  type PrintVisual,
  type PrintWorksheet,
} from '../src/renderer/print-html'

const EMU_PER_POINT = 12700

function sheet(): PrintWorksheet {
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

/// A chart two columns wide and two rows tall, anchored at B2 with no
/// intra-cell offset.
function chartAt(fromRow: number, fromColumn: number): PrintVisual {
  return {
    fromRow,
    fromColumn,
    fromRowOffset: 0,
    fromColumnOffset: 0,
    toRow: fromRow + 2,
    toColumn: fromColumn + 2,
    toRowOffset: 0,
    toColumnOffset: 0,
    html: '<figure class="xlsx-chart"><svg class="chart-svg"></svg></figure>',
  }
}

const build = (visuals: readonly PrintVisual[] = [], css = ''): string =>
  buildSheetPrintPayload(sheet(), {}, 'book.xlsx', 'Data', { visuals, css }).html

describe('print/PDF payload', () => {
  it('carries the cell values, so the harness itself is sound', () => {
    const html = build()
    expect(html).toContain('Region')
    expect(html).toContain('120')
  })

  it('carries no visual markup when there are no visuals', () => {
    const html = build()
    expect(html).not.toContain('<svg')
    expect(html).not.toContain('class="pv"')
  })

  it('places a chart over the cell it is anchored to', () => {
    const html = build([chartAt(1, 1)])
    expect(html).toContain('xlsx-chart')
    // 2 columns x 80px and 2 rows x 20px, at 0.75pt per px.
    expect(html).toContain('width:120pt;height:30pt')
    expect(html).toContain('left:0pt;top:0pt')
  })

  it('hangs the box off the anchor cell, not off the table', () => {
    // Pagination and repeated header rows both shift the table origin; a box
    // parented to its own td rides along with the cell instead of drifting.
    const html = build([chartAt(1, 1)])
    expect(html).toMatch(/<td[^>]*position:relative;overflow:visible[^>]*>120<div class="pv"/)
    // and no other cell is disturbed
    expect(html).toMatch(/<td(?![^>]*position:relative)[^>]*>North<\/td>/)
  })

  it('honours the intra-cell EMU offsets of a twoCellAnchor', () => {
    const visual = { ...chartAt(1, 1), fromColumnOffset: 6 * EMU_PER_POINT }
    const html = build([visual])
    expect(html).toContain('left:6pt;top:0pt')
    // the `to` marker did not move, so the frame is 6pt narrower
    expect(html).toContain('width:114pt')
  })

  it('falls back to the cell box for a degenerate anchor (sparklines)', () => {
    const spark: PrintVisual = {
      fromRow: 1,
      fromColumn: 1,
      fromRowOffset: 0,
      fromColumnOffset: 0,
      toRow: 1,
      toColumn: 1,
      toRowOffset: 0,
      toColumnOffset: 0,
      html: '<svg class="sparkline-svg"></svg>',
    }
    expect(build([spark])).toContain('width:60pt;height:15pt')
  })

  it('grows the used range to cover a visual beyond the data', () => {
    // The data stops at B5; a chart at D2:F5 is still part of what Excel
    // considers used, so the table has to reach far enough to host it.
    const html = build([chartAt(1, 3)])
    expect(html).toContain('class="pv"')
    // the `to` marker sits on F's left boundary with a zero offset, so the
    // chart ends at E — the table runs A..E, five <col> entries
    expect((html.match(/<col style/g) ?? []).length).toBe(5)
  })

  it('drops visuals anchored outside the printed area', () => {
    const outside = buildSheetPrintPayload(
      sheet(),
      { printArea: 'A1:B2' },
      'book.xlsx',
      'Data',
      { visuals: [chartAt(2, 0)], css: '' },
    ).html
    expect(outside).not.toContain('class="pv"')
  })

  it('inlines the captured stylesheet but cannot let it close the style block', () => {
    const html = build([chartAt(1, 1)], '.xlsx-chart { border: 1px solid #c9cdd1 }')
    expect(html).toContain('.xlsx-chart { border: 1px solid #c9cdd1 }')
    const escaped = build([chartAt(1, 1)], '</style><script>alert(1)</script>')
    expect(escaped).not.toContain('</style><script>')
  })
})
