/**
 * buildCellMatrix — what the user actually sees in a loaded range.
 *
 * Every wrong answer here is visible on screen and none of it throws: a
 * leading-zero part number silently becomes a number, a formula shows as
 * literal text because its cached value was absent, a table stripe paints
 * over the fill the file specified, the hyperlink blue overrides a colour the
 * author chose. These were unreachable while the logic lived inside a
 * function that also drove a live worksheet.
 */
import {
  BooleanNumber,
  BorderStyleTypes,
  CellValueType,
  HorizontalAlign,
  VerticalAlign,
  WrapStrategy,
  type ICellData,
  type IStyleData,
} from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import {
  applyTableBanding,
  buildCellMatrix,
  evictionRange,
  mapBorderStyle,
  mapHorizontalAlignment,
  mapVerticalAlignment,
  toRichTextDocument,
  toUniverStyle,
} from '../src/renderer/cell-matrix'

const range = (startRow = 0, endRow = 2, startColumn = 0, endColumn = 2) => ({
  startRow,
  endRow,
  startColumn,
  endColumn,
})

type Cell = Parameters<typeof buildCellMatrix>[0]['cells'][number]

/// ICellData.s is a style object or a style id; only the object form is built here.
const styleOf = (written: ICellData | undefined): IStyleData | undefined =>
  written?.s as IStyleData | undefined

const cell = (row: number, column: number, extra: Partial<Cell> = {}): Cell =>
  ({ row, column, value: null, ...extra }) as Cell

function build(overrides: Partial<Parameters<typeof buildCellMatrix>[0]> = {}) {
  return buildCellMatrix({
    range: range(),
    cells: [],
    styles: [],
    hyperlinks: [],
    tables: [],
    useFormulas: false,
    ...overrides,
  })
}

const table = (extra: Record<string, unknown> = {}) =>
  ({
    range: { startRow: 0, startColumn: 0, endRow: 4, endColumn: 2 },
    headerRowCount: 1,
    showRowStripes: true,
    showColumnStripes: false,
    ...extra,
  }) as never

describe('evictionRange', () => {
  it('clears nothing when no range was patched before', () => {
    expect(evictionRange(undefined, null)).toBeNull()
  })

  it('clears the whole previous range when nothing is frozen', () => {
    expect(evictionRange(range(5, 10, 2, 4), null)).toEqual(range(5, 10, 2, 4))
  })

  it('spares frozen rows, which later patches never re-send', () => {
    // Scroll away from a frozen header and it is not in the next viewport
    // patch; evicting it would leave the frozen strip blank.
    expect(evictionRange(range(0, 40, 0, 5), { frozenRows: 2, frozenColumns: 0 })).toMatchObject({
      startRow: 2,
    })
  })

  it('spares frozen columns too', () => {
    expect(evictionRange(range(0, 40, 0, 5), { frozenRows: 0, frozenColumns: 3 })).toMatchObject({
      startColumn: 3,
    })
  })

  it('leaves the previous range alone when it starts past the freeze', () => {
    expect(evictionRange(range(20, 40, 8, 12), { frozenRows: 2, frozenColumns: 3 })).toEqual(
      range(20, 40, 8, 12),
    )
  })

  it('clears nothing when the previous range is entirely frozen', () => {
    // Otherwise the arithmetic produces a negative height and Univer clears an
    // arbitrary rectangle.
    expect(evictionRange(range(0, 1, 0, 5), { frozenRows: 2, frozenColumns: 0 })).toBeNull()
    expect(evictionRange(range(0, 5, 0, 1), { frozenRows: 0, frozenColumns: 2 })).toBeNull()
  })
})

describe('buildCellMatrix: shape', () => {
  it('fills the whole requested rectangle, not just the cells supplied', () => {
    // Univer's setValues writes exactly what it is given; a short matrix would
    // leave whatever the previous range left behind.
    const matrix = build({ range: range(0, 3, 0, 4) })
    expect(matrix).toHaveLength(4)
    expect(matrix[0]).toHaveLength(5)
    expect(matrix[3]?.[4]).toEqual({})
  })

  it('positions cells relative to the range origin, not the sheet origin', () => {
    const matrix = build({
      range: range(10, 12, 5, 7),
      cells: [cell(11, 6, { value: 'here' })],
    })
    expect(matrix[1]?.[1]).toMatchObject({ v: 'here' })
    expect(matrix[0]?.[0]).toEqual({})
  })

  it('drops cells outside the range instead of writing out of bounds', () => {
    const matrix = build({
      range: range(5, 7, 5, 7),
      cells: [
        cell(4, 5, { value: 'above' }),
        cell(8, 5, { value: 'below' }),
        cell(5, 4, { value: 'left' }),
        cell(5, 8, { value: 'right' }),
      ],
    })
    expect(JSON.stringify(matrix)).not.toMatch(/above|below|left|right/)
  })
})

describe('buildCellMatrix: values', () => {
  it('types text explicitly so numeric-looking strings stay text', () => {
    // Without CellValueType.STRING, Univer coerces "007" to 7 and the user's
    // part number loses its leading zeros on screen and on save.
    expect(build({ cells: [cell(0, 0, { value: '007' })] })[0]?.[0]).toEqual({
      v: '007',
      t: CellValueType.STRING,
    })
  })

  it('leaves a real number as a number', () => {
    const written = build({ cells: [cell(0, 0, { value: 42 })] })[0]?.[0]
    expect(written).toEqual({ v: 42 })
  })

  it('writes an empty string without the string type marker', () => {
    expect(build({ cells: [cell(0, 0, { value: '' })] })[0]?.[0]).toEqual({ v: '' })
  })

  it('shows the cached value when formulas are off', () => {
    const written = build({ cells: [cell(0, 0, { value: 3, formula: 'A1+A2' })] })[0]?.[0]
    expect(written).toEqual({ v: 3 })
  })

  it('writes formula and cached value together when formulas are on', () => {
    const written = build({
      useFormulas: true,
      cells: [cell(0, 0, { value: 3, formula: 'A1+A2' })],
    })[0]?.[0]
    expect(written).toMatchObject({ f: 'A1+A2', v: 3 })
  })

  it('omits v for a formula with no cached value, so the engine computes it', () => {
    // Setting v to the fallback '' would show a blank cell forever; setting it
    // to the formula text would show "=A1+A2" as a literal.
    const written = build({
      useFormulas: true,
      cells: [cell(0, 0, { value: null, formula: 'A1+A2' })],
    })[0]?.[0]
    expect(written).toEqual({ f: 'A1+A2' })
    expect(written).not.toHaveProperty('v')
  })

  it('falls back to the formula text when formulas are off and nothing is cached', () => {
    const written = build({ cells: [cell(0, 0, { value: null, formula: 'A1+A2' })] })[0]?.[0]
    expect(written).toMatchObject({ v: 'A1+A2' })
  })
})

describe('buildCellMatrix: multiline and rich text', () => {
  it('uses the document model for a manual line break', () => {
    // A bare v renders only the first line.
    const written = build({ cells: [cell(0, 0, { value: 'a\nb' })] })[0]?.[0]
    expect(written).toHaveProperty('p')
    expect(written).not.toHaveProperty('v')
  })

  it('wraps a multiline cell even when the file sets no wrapText', () => {
    // Excel shows manual breaks regardless; without this the second line is
    // clipped rather than wrapped.
    expect(build({ cells: [cell(0, 0, { value: 'a\nb' })] })[0]?.[0]?.s).toMatchObject({
      tb: WrapStrategy.WRAP,
    })
  })

  it('uses the document model when the cell carries rich runs', () => {
    const written = build({
      cells: [cell(0, 0, { value: 'hello', rich: [{ text: 'hello', bold: true }] as never })],
    })[0]?.[0]
    expect(written).toHaveProperty('p')
  })

  it('prefers rich runs over the formula, which has no run information', () => {
    const written = build({
      useFormulas: true,
      cells: [cell(0, 0, { value: 'x', formula: 'A1', rich: [{ text: 'x' }] as never })],
    })[0]?.[0]
    expect(written).toHaveProperty('p')
    expect(written).not.toHaveProperty('f')
  })
})

describe('buildCellMatrix: hyperlinks', () => {
  const link = { row: 0, column: 0, target: 'https://example.com' } as never

  it('gives an unstyled link cell the Excel blue and an underline', () => {
    const s = build({ cells: [cell(0, 0, { value: 'go' })], hyperlinks: [link] })[0]?.[0]?.s
    expect(s).toMatchObject({ cl: { rgb: '#0563C1' }, ul: { s: BooleanNumber.TRUE } })
  })

  it('lets the file’s own colour win over the link blue (#161)', () => {
    // The blue is a fallback for links the file leaves unstyled, not a
    // stylesheet override — an author who coloured the link meant it.
    const s = build({
      cells: [cell(0, 0, { value: 'go', styleIndex: 0 })],
      styles: [{ fontColor: '#FF0000' } as never],
      hyperlinks: [link],
    })[0]?.[0]?.s
    expect(s).toMatchObject({ cl: { rgb: '#FF0000' } })
  })

  it('only styles the cell the link points at', () => {
    const matrix = build({
      cells: [cell(0, 0, { value: 'go' }), cell(0, 1, { value: 'plain' })],
      hyperlinks: [link],
    })
    expect(matrix[0]?.[1]?.s).toBeUndefined()
  })

  it('leaves a plain cell with no style object at all', () => {
    expect(build({ cells: [cell(0, 0, { value: 'plain' })] })[0]?.[0]).toEqual({
      v: 'plain',
      t: CellValueType.STRING,
    })
  })
})

describe('buildCellMatrix: CSE array followers', () => {
  const followers = new Set(['0:1'])

  it('drops a follower’s dead cached value so the master can spill', () => {
    // Leaving the stale value in place makes the master's spill collide with
    // it and the whole range renders #SPILL!.
    const matrix = build({
      useFormulas: true,
      arrayFollowers: followers,
      cells: [cell(0, 1, { value: 'stale' })],
    })
    expect(matrix[0]?.[1]).toEqual({})
  })

  it('keeps the follower’s style, which the engine does not supply', () => {
    const matrix = build({
      useFormulas: true,
      arrayFollowers: followers,
      cells: [cell(0, 1, { value: 'stale', styleIndex: 0 })],
      styles: [{ bold: true } as never],
    })
    expect(matrix[0]?.[1]?.s).toMatchObject({ bl: BooleanNumber.TRUE })
    expect(matrix[0]?.[1]).not.toHaveProperty('v')
  })

  it('keeps cached values when formulas are off, since nothing will recompute', () => {
    const matrix = build({
      useFormulas: false,
      arrayFollowers: followers,
      cells: [cell(0, 1, { value: 'cached' })],
    })
    expect(matrix[0]?.[1]).toMatchObject({ v: 'cached' })
  })

  it('leaves the master itself alone', () => {
    const matrix = build({
      useFormulas: true,
      arrayFollowers: followers,
      cells: [cell(0, 0, { value: 1, formula: 'TRANSPOSE(D1:E1)' })],
    })
    expect(matrix[0]?.[0]).toMatchObject({ f: 'TRANSPOSE(D1:E1)' })
  })
})

describe('table banding', () => {
  it('bolds the header row and fills it', () => {
    const matrix = build({
      range: range(0, 4, 0, 2),
      tables: [table({ headerFill: '#4472C4' })],
    })
    expect(matrix[0]?.[0]?.s).toMatchObject({
      bg: { rgb: '#4472C4' },
      bl: BooleanNumber.TRUE,
    })
  })

  it('stripes alternate data rows, starting after the header', () => {
    const matrix = build({ range: range(0, 4, 0, 2), tables: [table()] })
    // Header is row 0; data starts at row 1, so row 2 is the first stripe.
    expect(styleOf(matrix[1]?.[0])?.bg).toBeUndefined()
    expect(matrix[2]?.[0]?.s).toMatchObject({ bg: { rgb: '#D9E1F2' } })
    expect(styleOf(matrix[3]?.[0])?.bg).toBeUndefined()
  })

  it('skips stripes when the table has them turned off', () => {
    const matrix = build({
      range: range(0, 4, 0, 2),
      tables: [table({ showRowStripes: false })],
    })
    expect(styleOf(matrix[2]?.[0])?.bg).toBeUndefined()
  })

  it('never paints over a fill the file specified', () => {
    // Banding approximates a table style; an explicit fill is data.
    const matrix = build({
      range: range(0, 4, 0, 2),
      tables: [table()],
      cells: [cell(2, 0, { value: 'x', styleIndex: 0 })],
      styles: [{ fillColor: '#FFFF00' } as never],
    })
    expect(matrix[2]?.[0]?.s).toMatchObject({ bg: { rgb: '#FFFF00' } })
  })

  it('clips a table that only partly overlaps the loaded range', () => {
    // The viewport patch covers rows 2-4 of a table that starts at row 0; the
    // header is off-screen and must not be painted onto row 2.
    const matrix = build({
      range: range(2, 4, 0, 2),
      tables: [table()],
    })
    expect(styleOf(matrix[0]?.[0])?.bl).toBeUndefined()
  })

  it('ignores a table that does not intersect the range at all', () => {
    const matrix = build({
      range: range(50, 52, 0, 2),
      tables: [table()],
    })
    expect(matrix[0]?.[0]).toEqual({})
  })

  it('supports a multi-row header', () => {
    const matrix = build({
      range: range(0, 4, 0, 2),
      tables: [table({ headerRowCount: 2 })],
    })
    expect(matrix[1]?.[0]?.s).toMatchObject({ bl: BooleanNumber.TRUE })
  })

  it('uses a dark default font when the theme supplies no header fill', () => {
    // White-on-white is the failure this guards.
    const matrix = build({ range: range(0, 4, 0, 2), tables: [table()] })
    expect(matrix[0]?.[0]?.s).toMatchObject({ cl: { rgb: '#333333' } })
  })

  it('tolerates a matrix smaller than the table it is told to band', () => {
    const matrix = [[{}]]
    expect(() =>
      applyTableBanding(matrix, range(0, 10, 0, 10), [table({ headerRowCount: 0 })]),
    ).not.toThrow()
  })
})

describe('toUniverStyle', () => {
  it('states bold and italic explicitly, so unsetting them takes effect', () => {
    // Omitting the key leaves the previous cell's bold in place when the
    // set-range-values mutation merges styles.
    expect(toUniverStyle({} as never)).toMatchObject({
      bl: BooleanNumber.FALSE,
      it: BooleanNumber.FALSE,
    })
  })

  it('maps fonts, colours and number format', () => {
    expect(
      toUniverStyle({
        fontFamily: 'Calibri',
        fontSize: 11,
        fontColor: '#112233',
        fillColor: '#FFEEDD',
        numberFormat: '0.00%',
      } as never),
    ).toMatchObject({
      ff: 'Calibri',
      fs: 11,
      cl: { rgb: '#112233' },
      bg: { rgb: '#FFEEDD' },
      n: { pattern: '0.00%' },
    })
  })

  it('omits the border block entirely when there are no borders', () => {
    expect(toUniverStyle({ bold: true } as never)).not.toHaveProperty('bd')
  })

  it('maps each edge to its own slot', () => {
    const style = toUniverStyle({
      borderTop: { style: 'thin' },
      borderBottom: { style: 'thick' },
      borderLeft: { style: 'dotted' },
      borderRight: { style: 'double' },
    } as never)
    expect(Object.keys(style.bd ?? {}).sort()).toEqual(['b', 'l', 'r', 't'])
  })

  it('defaults a border with no colour to black rather than leaving it unset', () => {
    const style = toUniverStyle({ borderTop: { style: 'thin' } } as never)
    expect(style.bd?.t).toMatchObject({ cl: { rgb: '#000000' } })
  })

  it('only emits a diagonal for the direction the file marks', () => {
    const down = toUniverStyle({
      borderDiagonal: { style: 'thin' },
      diagonalDown: true,
    } as never)
    expect(down.bd).toHaveProperty('tl_br')
    expect(down.bd).not.toHaveProperty('bl_tr')
  })

  it('maps an up diagonal to the other slot', () => {
    const up = toUniverStyle({ borderDiagonal: { style: 'thin' }, diagonalUp: true } as never)
    expect(up.bd).toHaveProperty('bl_tr')
    expect(up.bd).not.toHaveProperty('tl_br')
  })

  it('emits both diagonals for a crossed cell', () => {
    const both = toUniverStyle({
      borderDiagonal: { style: 'thin' },
      diagonalUp: true,
      diagonalDown: true,
    } as never)
    expect(both.bd).toHaveProperty('bl_tr')
    expect(both.bd).toHaveProperty('tl_br')
  })

  it('emits no diagonal at all when neither direction is set', () => {
    const style = toUniverStyle({ borderDiagonal: { style: 'thin' } } as never)
    expect(style).not.toHaveProperty('bd')
  })

  it('converts indent levels to pixels', () => {
    const one = toUniverStyle({ indent: 1 } as never)
    const two = toUniverStyle({ indent: 2 } as never)
    expect((two.pd?.l ?? 0)).toBe((one.pd?.l ?? 0) * 2)
  })

  it('omits padding for indent 0, which is not the same as an indent', () => {
    expect(toUniverStyle({ indent: 0 } as never)).not.toHaveProperty('pd')
  })

  it('carries alignment through', () => {
    expect(
      toUniverStyle({ horizontalAlignment: 'right', verticalAlignment: 'top' } as never),
    ).toMatchObject({ ht: HorizontalAlign.RIGHT, vt: VerticalAlign.TOP })
  })

  it('leaves alignment unset for "general", so Excel’s own rules apply', () => {
    // Numbers right-align and text left-aligns under "general"; pinning it to
    // LEFT here would left-align every numeric column in the workbook.
    const style = toUniverStyle({
      horizontalAlignment: 'general',
      verticalAlignment: undefined,
    } as never)
    expect(style).not.toHaveProperty('ht')
    expect(style).not.toHaveProperty('vt')
  })

  it('carries underline, strikethrough and wrap', () => {
    expect(
      toUniverStyle({ underline: true, strikethrough: true, wrapText: true } as never),
    ).toMatchObject({
      ul: { s: BooleanNumber.TRUE },
      st: { s: BooleanNumber.TRUE },
      tb: WrapStrategy.WRAP,
    })
  })
})

describe('style enum mapping', () => {
  it.each([
    ['hair', BorderStyleTypes.HAIR],
    ['dotted', BorderStyleTypes.DOTTED],
    ['dashed', BorderStyleTypes.DASHED],
    ['dashDot', BorderStyleTypes.DASH_DOT],
    ['dashDotDot', BorderStyleTypes.DASH_DOT_DOT],
    ['double', BorderStyleTypes.DOUBLE],
    ['medium', BorderStyleTypes.MEDIUM],
    ['mediumDashed', BorderStyleTypes.MEDIUM_DASHED],
    ['mediumDashDot', BorderStyleTypes.MEDIUM_DASH_DOT],
    ['mediumDashDotDot', BorderStyleTypes.MEDIUM_DASH_DOT_DOT],
    ['slantDashDot', BorderStyleTypes.SLANT_DASH_DOT],
    ['thick', BorderStyleTypes.THICK],
    ['thin', BorderStyleTypes.THIN],
  ])('maps the %s border', (name, expected) => {
    expect(mapBorderStyle(name)).toBe(expected)
  })

  it('falls back to thin for a style it does not know', () => {
    // A border the file asks for and we cannot name should still be a border.
    expect(mapBorderStyle('someFutureStyle')).toBe(BorderStyleTypes.THIN)
  })

  it.each([
    ['left', HorizontalAlign.LEFT],
    ['center', HorizontalAlign.CENTER],
    ['right', HorizontalAlign.RIGHT],
    ['justify', HorizontalAlign.JUSTIFIED],
    ['distributed', HorizontalAlign.DISTRIBUTED],
  ])('maps horizontal %s', (name, expected) => {
    expect(mapHorizontalAlignment(name)).toBe(expected)
  })

  it.each([
    ['top', VerticalAlign.TOP],
    ['center', VerticalAlign.MIDDLE],
    ['bottom', VerticalAlign.BOTTOM],
  ])('maps vertical %s', (name, expected) => {
    expect(mapVerticalAlignment(name)).toBe(expected)
  })

  it('returns undefined for an unset alignment rather than guessing left/top', () => {
    // Excel's "general" alignment right-aligns numbers and left-aligns text;
    // forcing LEFT here would break every numeric column.
    expect(mapHorizontalAlignment(undefined)).toBeUndefined()
    expect(mapHorizontalAlignment('general')).toBeUndefined()
    expect(mapVerticalAlignment(undefined)).toBeUndefined()
  })
})

describe('toRichTextDocument', () => {
  it('gives each run the offsets of its own text', () => {
    const p = toRichTextDocument('abcdef', [{ text: 'abc' }, { text: 'def' }] as never)
    expect(p?.body?.textRuns).toMatchObject([
      { st: 0, ed: 3 },
      { st: 3, ed: 6 },
    ])
  })

  it('only sets the flags a run actually declares', () => {
    // The DSL lets a run say "bold the second word" without inventing values
    // for italic, size and colour.
    const ts = toRichTextDocument('x', [{ text: 'x', bold: true }] as never)?.body?.textRuns?.[0]
      ?.ts
    expect(ts).toEqual({ bl: BooleanNumber.TRUE })
  })

  it('carries every supported run property through', () => {
    const ts = toRichTextDocument('x', [
      {
        text: 'x',
        family: 'Arial',
        size: 14,
        bold: true,
        italic: true,
        underline: true,
        strikethrough: true,
        color: '#00FF00',
      },
    ] as never)?.body?.textRuns?.[0]?.ts
    expect(ts).toMatchObject({
      ff: 'Arial',
      fs: 14,
      bl: BooleanNumber.TRUE,
      it: BooleanNumber.TRUE,
      ul: { s: BooleanNumber.TRUE },
      st: { s: BooleanNumber.TRUE },
      cl: { rgb: '#00FF00' },
    })
  })

  it('converts newlines to paragraph breaks, not section breaks', () => {
    // A raw \n is a section break in a Univer document stream: it splits the
    // cell and everything after the first line disappears.
    const p = toRichTextDocument('a\nb\nc')
    expect(p?.body?.dataStream).toBe('a\rb\rc\r\n')
    expect(p?.body?.paragraphs).toHaveLength(3)
  })

  it('keeps the stream the same length so run offsets stay valid', () => {
    const p = toRichTextDocument('ab\ncd', [{ text: 'ab\ncd' }] as never)
    const run = p?.body?.textRuns?.[0]
    expect(run?.ed).toBe(5)
    expect(p?.body?.dataStream?.startsWith('ab\rcd')).toBe(true)
  })

  it('always closes with a section break at the end of the stream', () => {
    const p = toRichTextDocument('a')
    expect(p?.body?.sectionBreaks).toEqual([{ startIndex: (p?.body?.dataStream?.length ?? 0) - 1 }])
  })

  it('handles empty text without producing an invalid document', () => {
    const p = toRichTextDocument('')
    expect(p?.body?.dataStream).toBe('\r\n')
    expect(p?.body?.paragraphs).toHaveLength(1)
  })
})
