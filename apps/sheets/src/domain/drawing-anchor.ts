/**
 * xlsx drawing anchor arithmetic.
 *
 * An anchor edge is a cell index plus an offset inside that cell (`<xdr:col>`
 * + `<xdr:colOff>`), so every geometric question about a drawing — how wide is
 * it, where does dragging it by 40px put it — is a walk across real row
 * heights and column widths rather than plain arithmetic.
 *
 * Offsets are EMU in the file and pixels here; 9525 EMU per CSS pixel at 96dpi.
 */

export const EMU_PER_PIXEL = 9525

/// One edge of a drawing anchor: a cell index plus a pixel offset inside it.
export interface AnchorMarker {
  index: number
  offset: number
}

export const markerFrom = (index: number, offsetEmu: number): AnchorMarker => ({
  index,
  offset: offsetEmu / EMU_PER_PIXEL,
})

/// Move a marker by a pixel delta, carrying across real row/column sizes.
/// Clamps at the sheet start and inside the last row/column.
export function walkMarker(
  marker: AnchorMarker,
  delta: number,
  sizeOf: (index: number) => number,
  maxIndex: number,
): AnchorMarker {
  let index = Math.min(marker.index, maxIndex)
  let offset = marker.offset + delta
  while (offset < 0 && index > 0) {
    index -= 1
    offset += sizeOf(index)
  }
  if (offset < 0) offset = 0
  while (index < maxIndex && offset >= sizeOf(index)) {
    offset -= sizeOf(index)
    index += 1
  }
  if (index >= maxIndex) offset = Math.min(offset, sizeOf(maxIndex))
  return { index, offset }
}

/// Pixel distance between two markers (negative when `to` sits before `from`).
export function markerSpan(
  from: AnchorMarker,
  to: AnchorMarker,
  sizeOf: (index: number) => number,
): number {
  const span = to.offset - from.offset
  const low = Math.min(from.index, to.index)
  const high = Math.max(from.index, to.index)
  let cells = 0
  for (let index = low; index < high; index += 1) cells += sizeOf(index)
  return span + (from.index <= to.index ? cells : -cells)
}

/// The eight numbers of a twoCellAnchor, plus the extent a oneCellAnchor
/// carries in place of its second marker.
export interface DrawingAnchorLike {
  readonly fromRow: number
  readonly fromColumn: number
  readonly fromRowOffset: number
  readonly fromColumnOffset: number
  readonly toRow: number
  readonly toColumn: number
  readonly toRowOffset: number
  readonly toColumnOffset: number
  readonly extWidthEmu?: number | undefined
  readonly extHeightEmu?: number | undefined
}

export interface AnchorSizes {
  readonly columnWidth: (index: number) => number
  readonly rowHeight: (index: number) => number
  readonly maxColumn: number
  readonly maxRow: number
}

/// Gives a oneCellAnchor the second marker it never had.
///
/// Such a drawing states its size as an extent — "415px wide" — and the parser
/// can only report `to == from`, a zero span, which renders and prints as a
/// sliver of a single cell. Walking the extent across the real column widths
/// and row heights recovers the frame the file actually asked for.
///
/// A twoCellAnchor already has both markers and passes through untouched, as
/// does an anchor whose extent is missing or whose span is already non-zero.
export function resolveAnchorExtent<T extends DrawingAnchorLike>(
  anchor: T,
  sizes: AnchorSizes,
): T {
  const { extWidthEmu, extHeightEmu } = anchor
  if (extWidthEmu === undefined && extHeightEmu === undefined) return anchor
  const degenerate =
    anchor.toRow === anchor.fromRow &&
    anchor.toColumn === anchor.fromColumn &&
    anchor.toRowOffset === anchor.fromRowOffset &&
    anchor.toColumnOffset === anchor.fromColumnOffset
  if (!degenerate) return anchor
  const across = (
    index: number,
    offsetEmu: number,
    extentEmu: number | undefined,
    sizeOf: (index: number) => number,
    maxIndex: number,
  ): AnchorMarker => {
    const start = markerFrom(index, offsetEmu)
    if (extentEmu === undefined || extentEmu <= 0) return start
    return walkMarker(start, extentEmu / EMU_PER_PIXEL, sizeOf, maxIndex)
  }
  const toX = across(
    anchor.fromColumn,
    anchor.fromColumnOffset,
    extWidthEmu,
    sizes.columnWidth,
    sizes.maxColumn,
  )
  const toY = across(
    anchor.fromRow,
    anchor.fromRowOffset,
    extHeightEmu,
    sizes.rowHeight,
    sizes.maxRow,
  )
  return {
    ...anchor,
    toColumn: toX.index,
    toColumnOffset: Math.round(toX.offset * EMU_PER_PIXEL),
    toRow: toY.index,
    toRowOffset: Math.round(toY.offset * EMU_PER_PIXEL),
  }
}
