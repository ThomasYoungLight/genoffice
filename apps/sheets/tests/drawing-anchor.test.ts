/**
 * oneCellAnchor drawings.
 *
 * Such a drawing gives a single marker and an extent — "start at D2, be
 * 415x227px" — instead of two markers. The parser can only report `to ==
 * from`, a zero span, so before this the drawing collapsed into a sliver of
 * its top-left cell: on screen, in print, and in the frame a drag started
 * from. openpyxl writes oneCellAnchor by default and Excel writes it for
 * "move but don't size with cells", so this is the common case, not an exotic
 * one.
 */
import { describe, expect, it } from 'vitest'

import {
  type DrawingAnchorLike,
  EMU_PER_PIXEL,
  resolveAnchorExtent,
} from '../src/domain/drawing-anchor'

/// 80px columns and 20px rows, the defaults the print tests also assume.
const SIZES = {
  columnWidth: () => 80,
  rowHeight: () => 20,
  maxColumn: 16383,
  maxRow: 1048575,
}

/// What the sidecar reports for `<oneCellAnchor><from>D2</from><ext .../>`:
/// both markers on the from cell, plus the extent.
function oneCell(widthPx: number, heightPx: number): DrawingAnchorLike {
  return {
    fromRow: 1,
    fromColumn: 3,
    fromRowOffset: 0,
    fromColumnOffset: 0,
    toRow: 1,
    toColumn: 3,
    toRowOffset: 0,
    toColumnOffset: 0,
    extWidthEmu: widthPx * EMU_PER_PIXEL,
    extHeightEmu: heightPx * EMU_PER_PIXEL,
  }
}

describe('resolveAnchorExtent', () => {
  it('walks the extent across real cells to recover the second marker', () => {
    // 200px from the left edge of D = two whole 80px columns plus 40 into F.
    const resolved = resolveAnchorExtent(oneCell(200, 50), SIZES)
    expect(resolved.toColumn).toBe(5)
    expect(resolved.toColumnOffset).toBe(40 * EMU_PER_PIXEL)
    // 50px from the top of row 2 = two whole 20px rows plus 10 into row 4.
    expect(resolved.toRow).toBe(3)
    expect(resolved.toRowOffset).toBe(10 * EMU_PER_PIXEL)
  })

  it('carries the from-marker offset, so the extent measures from the edge', () => {
    // the sidecar reads both markers off the same <from> node, so an offset
    // appears on each — that is still degenerate, and still needs resolving
    const offset = {
      ...oneCell(200, 50),
      fromColumnOffset: 60 * EMU_PER_PIXEL,
      toColumnOffset: 60 * EMU_PER_PIXEL,
    }
    const resolved = resolveAnchorExtent(offset, SIZES)
    // starts 60 into D, so only 20 of D is left; +E +F = 180, and the last
    // 20 land in G
    expect(resolved.toColumn).toBe(6)
    expect(resolved.toColumnOffset).toBe(20 * EMU_PER_PIXEL)
  })

  it('lands exactly on a boundary when the extent is a whole number of cells', () => {
    const resolved = resolveAnchorExtent(oneCell(160, 40), SIZES)
    expect(resolved.toColumn).toBe(5)
    expect(resolved.toColumnOffset).toBe(0)
    expect(resolved.toRow).toBe(3)
    expect(resolved.toRowOffset).toBe(0)
  })

  it('leaves a twoCellAnchor completely alone', () => {
    // no extent at all: both markers came from the file
    const two: DrawingAnchorLike = {
      fromRow: 1,
      fromColumn: 3,
      fromRowOffset: 0,
      fromColumnOffset: 0,
      toRow: 12,
      toColumn: 8,
      toRowOffset: 0,
      toColumnOffset: 0,
    }
    expect(resolveAnchorExtent(two, SIZES)).toBe(two)
  })

  it('does not second-guess an anchor that already has a real span', () => {
    // both a span and an extent: the markers win, they are more specific
    const both = { ...oneCell(200, 50), toColumn: 8, toRow: 12 }
    const resolved = resolveAnchorExtent(both, SIZES)
    expect(resolved.toColumn).toBe(8)
    expect(resolved.toRow).toBe(12)
  })

  it('leaves a sizeless anchor degenerate rather than inventing a frame', () => {
    // absoluteAnchor and malformed oneCellAnchor both land here; install falls
    // back to filling the cell range, which is the old behaviour
    const sizeless = { ...oneCell(0, 0), extWidthEmu: undefined, extHeightEmu: undefined }
    expect(resolveAnchorExtent(sizeless, SIZES)).toBe(sizeless)
  })

  it('respects uneven column widths instead of assuming a default', () => {
    const uneven = {
      ...SIZES,
      columnWidth: (index: number) => (index === 3 ? 300 : 80),
    }
    // 300px eats all of D on its own, so 320 ends 20 into E
    const resolved = resolveAnchorExtent(oneCell(320, 20), uneven)
    expect(resolved.toColumn).toBe(4)
    expect(resolved.toColumnOffset).toBe(20 * EMU_PER_PIXEL)
  })
})
