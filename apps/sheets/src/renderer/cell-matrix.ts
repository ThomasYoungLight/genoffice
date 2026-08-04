/**
 * Turning a range of file cells into the matrix Univer renders.
 *
 * `patchWorksheetRangeInner` did two unrelated things: decide what the user
 * should see, and push it into a live worksheet. The first is where every
 * visible bug lives — a numeric-looking string coerced to a number, a link
 * colour overriding the one the file specifies, a formula shown as literal
 * text because a cached value was missing, a table stripe painted over an
 * explicit fill. The second is two method calls.
 *
 * Split apart, the decisions are a plain function from file data to cell data,
 * so a test can assert on the matrix itself rather than fishing the values
 * back out of a worksheet double.
 *
 * Nothing here touches a worksheet, the journal or the runtime.
 */
import {
  BooleanNumber,
  BorderStyleTypes,
  CellValueType,
  HorizontalAlign,
  VerticalAlign,
  WrapStrategy,
  type ICellData,
  type IRange,
  type IStyleData,
} from '@univerjs/core'

import type { RichRun } from '../domain/workbook.types'
import type {
  WorkbookCellStyle,
  WorkbookFile,
  WorkbookRangeResult,
} from '../shared/desktop-api'
import { INDENT_STEP_PX } from './selection-format'

/** Excel's own hyperlink blue, used only where the file specifies nothing. */
const LINK_COLOR = '#0563C1'

type SheetTables = WorkbookFile['sheets'][number]['tables']
type SheetFreeze = WorkbookFile['sheets'][number]['freeze']

/**
 * The part of a previously-patched range that may be cleared before the next
 * one is written, or null if there is nothing to clear.
 *
 * Frozen rows and columns stay on screen while the viewport scrolls away from
 * them, and later patches do not include them — evicting them would leave the
 * frozen strip blank.
 */
export function evictionRange(
  previousRange: IRange | undefined,
  freeze: SheetFreeze,
): IRange | null {
  if (!previousRange) return null
  const startRow = Math.max(previousRange.startRow, freeze?.frozenRows ?? 0)
  const startColumn = Math.max(previousRange.startColumn, freeze?.frozenColumns ?? 0)
  if (startRow > previousRange.endRow || startColumn > previousRange.endColumn) return null
  return { startRow, startColumn, endRow: previousRange.endRow, endColumn: previousRange.endColumn }
}

export interface CellMatrixInput {
  readonly range: IRange
  readonly cells: WorkbookRangeResult['cells']
  readonly styles: readonly WorkbookCellStyle[]
  readonly hyperlinks: WorkbookRangeResult['hyperlinks']
  readonly tables: SheetTables
  readonly useFormulas: boolean
  /**
   * Cells covered by a legacy CSE array formula but not its master. Their
   * cached values are dead and would block the master's spill with #SPILL!,
   * so they keep their style and let the engine fill the content.
   */
  readonly arrayFollowers?: ReadonlySet<string> | undefined
}

/** Builds the `ICellData` matrix for a range, table banding included. */
export function buildCellMatrix(input: CellMatrixInput): ICellData[][] {
  const { range, styles, useFormulas, arrayFollowers } = input
  const linkedCells = new Set(input.hyperlinks.map((link) => `${link.row}:${link.column}`))
  const rows = range.endRow - range.startRow + 1
  const columns = range.endColumn - range.startColumn + 1
  const matrix: ICellData[][] = Array.from({ length: rows }, () =>
    Array.from({ length: columns }, () => ({})),
  )
  for (const cell of input.cells) {
    if (
      cell.row < range.startRow ||
      cell.row > range.endRow ||
      cell.column < range.startColumn ||
      cell.column > range.endColumn
    ) {
      continue
    }
    const row = matrix[cell.row - range.startRow]
    if (!row) continue
    const style = cell.styleIndex === undefined ? undefined : styles[cell.styleIndex]
    if (useFormulas && arrayFollowers?.has(`${cell.row}:${cell.column}`)) {
      row[cell.column - range.startColumn] = style ? { s: toUniverStyle(style) } : {}
      continue
    }
    const displayValue = cell.value ?? cell.formula ?? ''
    const isLink = linkedCells.has(`${cell.row}:${cell.column}`)
    const multiline = typeof displayValue === 'string' && displayValue.includes('\n')
    row[cell.column - range.startColumn] = {
      ...cellContent(cell, displayValue, useFormulas, multiline),
      ...(style || isLink || multiline
        ? { s: cellStyle(style, isLink, multiline) }
        : {}),
    }
  }
  applyTableBanding(matrix, range, input.tables)
  return matrix
}

function cellContent(
  cell: WorkbookRangeResult['cells'][number],
  displayValue: string | number | boolean,
  useFormulas: boolean,
  multiline: boolean,
): ICellData {
  if (cell.rich && typeof displayValue === 'string') {
    return { p: toRichTextDocument(displayValue, cell.rich) }
  }
  if (useFormulas && cell.formula) {
    // No cached value: leave `v` unset so the engine computes it instead of
    // showing the formula text as a literal.
    return cell.value === null || cell.value === undefined
      ? { f: cell.formula }
      : { f: cell.formula, v: cell.value }
  }
  // Bare `v` renders only the first line; the doc model keeps all of them.
  if (typeof displayValue === 'string' && multiline) {
    return { p: toRichTextDocument(displayValue) }
  }
  // Explicit string typing: bare `v` lets Univer coerce numeric-looking text
  // ("007", phone numbers) into numbers.
  if (typeof displayValue === 'string' && displayValue !== '') {
    return { v: displayValue, t: CellValueType.STRING }
  }
  return { v: displayValue }
}

function cellStyle(
  style: WorkbookCellStyle | undefined,
  isLink: boolean,
  multiline: boolean,
): IStyleData {
  return {
    // Link blue/underline is a fallback only: a colour or underline the file
    // specifies must win (#161).
    ...(isLink ? { cl: { rgb: LINK_COLOR }, ul: { s: BooleanNumber.TRUE } } : {}),
    ...(style ? toUniverStyle(style) : {}),
    // Excel shows manual line breaks even without wrapText.
    ...(multiline ? { tb: WrapStrategy.WRAP } : {}),
  }
}

/**
 * Approximates Excel table styles (header band + row stripes) for cells that
 * carry no explicit fill of their own.
 */
export function applyTableBanding(
  matrix: ICellData[][],
  range: IRange,
  tables: SheetTables,
): void {
  for (const table of tables) {
    const rowStart = Math.max(range.startRow, table.range.startRow)
    const rowEnd = Math.min(range.endRow, table.range.endRow)
    const columnStart = Math.max(range.startColumn, table.range.startColumn)
    const columnEnd = Math.min(range.endColumn, table.range.endColumn)
    if (rowStart > rowEnd || columnStart > columnEnd) continue
    // Colors are resolved sidecar-side from the workbook's real theme accents
    // (Light/Medium/Dark variant rules); the literals are a last-resort fallback.
    const headerFill = table.headerFill
    const headerFont = table.headerFontColor ?? '#FFFFFF'
    const stripeFill = table.stripeFill ?? '#D9E1F2'
    const dataStartRow = table.range.startRow + table.headerRowCount
    for (let row = rowStart; row <= rowEnd; row += 1) {
      const isHeader = row < dataStartRow
      const isStripe = !isHeader && table.showRowStripes && (row - dataStartRow) % 2 === 1
      if (!isHeader && !isStripe) continue
      for (let column = columnStart; column <= columnEnd; column += 1) {
        const cell = matrix[row - range.startRow]?.[column - range.startColumn]
        if (!cell) continue
        const style = (cell.s ?? {}) as IStyleData
        if (style.bg) continue
        cell.s = isHeader
          ? {
              ...style,
              ...(headerFill ? { bg: { rgb: headerFill } } : {}),
              cl: { rgb: headerFill ? headerFont : (table.headerFontColor ?? '#333333') },
              bl: BooleanNumber.TRUE,
            }
          : { ...style, bg: { rgb: stripeFill } }
      }
    }
  }
}

/**
 * `runs` is deliberately looser than WorkbookRichRun: the DSL lets a run set
 * only what it changes ("bold the second word"), and every flag is read
 * truthily below, so requiring all four would force callers to invent values
 * for flags they have no opinion about.
 */
export function toRichTextDocument(
  text: string,
  runs: readonly RichRun[] = [],
): ICellData['p'] {
  const textRuns = []
  let cursor = 0
  for (const run of runs) {
    const end = cursor + run.text.length
    textRuns.push({
      st: cursor,
      ed: end,
      ts: {
        ...(run.family ? { ff: run.family } : {}),
        ...(run.size ? { fs: run.size } : {}),
        ...(run.bold ? { bl: BooleanNumber.TRUE } : {}),
        ...(run.italic ? { it: BooleanNumber.TRUE } : {}),
        ...(run.underline ? { ul: { s: BooleanNumber.TRUE } } : {}),
        ...(run.strikethrough ? { st: { s: BooleanNumber.TRUE } } : {}),
        ...(run.color ? { cl: { rgb: run.color } } : {}),
      },
    })
    cursor = end
  }
  // Univer document streams use \r as paragraph break and \n as section
  // break; a raw \n would split the cell into sections and drop later lines.
  // 1:1 replacement, so textRun offsets stay valid.
  const dataStream = `${text.replace(/\n/g, '\r')}\r\n`
  const paragraphs: Array<{ startIndex: number }> = []
  for (let i = 0; i < dataStream.length; i += 1) {
    if (dataStream[i] === '\r') paragraphs.push({ startIndex: i })
  }
  return {
    id: 'rich-cell',
    body: {
      dataStream,
      textRuns,
      paragraphs,
      sectionBreaks: [{ startIndex: dataStream.length - 1 }],
    },
    documentStyle: {},
  }
}

export function toUniverStyle(style: WorkbookCellStyle): IStyleData {
  const diagonal = style.borderDiagonal ? toUniverBorder(style.borderDiagonal) : undefined
  const borders = {
    ...(style.borderTop ? { t: toUniverBorder(style.borderTop) } : {}),
    ...(style.borderBottom ? { b: toUniverBorder(style.borderBottom) } : {}),
    ...(style.borderLeft ? { l: toUniverBorder(style.borderLeft) } : {}),
    ...(style.borderRight ? { r: toUniverBorder(style.borderRight) } : {}),
    ...(diagonal && style.diagonalDown ? { tl_br: diagonal } : {}),
    ...(diagonal && style.diagonalUp ? { bl_tr: diagonal } : {}),
  }
  return {
    ...(style.fontFamily ? { ff: style.fontFamily } : {}),
    ...(style.fontSize ? { fs: style.fontSize } : {}),
    bl: style.bold ? BooleanNumber.TRUE : BooleanNumber.FALSE,
    it: style.italic ? BooleanNumber.TRUE : BooleanNumber.FALSE,
    ...(style.underline ? { ul: { s: BooleanNumber.TRUE } } : {}),
    ...(style.strikethrough ? { st: { s: BooleanNumber.TRUE } } : {}),
    ...(style.wrapText ? { tb: WrapStrategy.WRAP } : {}),
    ...(style.fontColor ? { cl: { rgb: style.fontColor } } : {}),
    ...(style.fillColor ? { bg: { rgb: style.fillColor } } : {}),
    ...(style.numberFormat ? { n: { pattern: style.numberFormat } } : {}),
    ...(Object.keys(borders).length > 0 ? { bd: borders } : {}),
    ...(mapHorizontalAlignment(style.horizontalAlignment) === undefined
      ? {}
      : { ht: mapHorizontalAlignment(style.horizontalAlignment) }),
    ...(mapVerticalAlignment(style.verticalAlignment) === undefined
      ? {}
      : { vt: mapVerticalAlignment(style.verticalAlignment) }),
    ...(style.indent ? { pd: { l: style.indent * INDENT_STEP_PX } } : {}),
  }
}

function toUniverBorder(edge: NonNullable<WorkbookCellStyle['borderTop']>): {
  s: BorderStyleTypes
  cl: { rgb: string }
} {
  return {
    s: mapBorderStyle(edge.style),
    cl: { rgb: edge.color ?? '#000000' },
  }
}

export function mapBorderStyle(style: string): BorderStyleTypes {
  switch (style) {
    case 'hair':
      return BorderStyleTypes.HAIR
    case 'dotted':
      return BorderStyleTypes.DOTTED
    case 'dashed':
      return BorderStyleTypes.DASHED
    case 'dashDot':
      return BorderStyleTypes.DASH_DOT
    case 'dashDotDot':
      return BorderStyleTypes.DASH_DOT_DOT
    case 'double':
      return BorderStyleTypes.DOUBLE
    case 'medium':
      return BorderStyleTypes.MEDIUM
    case 'mediumDashed':
      return BorderStyleTypes.MEDIUM_DASHED
    case 'mediumDashDot':
      return BorderStyleTypes.MEDIUM_DASH_DOT
    case 'mediumDashDotDot':
      return BorderStyleTypes.MEDIUM_DASH_DOT_DOT
    case 'slantDashDot':
      return BorderStyleTypes.SLANT_DASH_DOT
    case 'thick':
      return BorderStyleTypes.THICK
    default:
      return BorderStyleTypes.THIN
  }
}

export function mapHorizontalAlignment(value: string | undefined): HorizontalAlign | undefined {
  if (value === 'left') return HorizontalAlign.LEFT
  if (value === 'center') return HorizontalAlign.CENTER
  if (value === 'right') return HorizontalAlign.RIGHT
  if (value === 'justify') return HorizontalAlign.JUSTIFIED
  if (value === 'distributed') return HorizontalAlign.DISTRIBUTED
  return undefined
}

export function mapVerticalAlignment(value: string | undefined): VerticalAlign | undefined {
  if (value === 'top') return VerticalAlign.TOP
  if (value === 'center') return VerticalAlign.MIDDLE
  if (value === 'bottom') return VerticalAlign.BOTTOM
  return undefined
}
