/**
 * A hand-built stand-in for the slice of Univer's facade the renderer's
 * application layer touches.
 *
 * Those modules were deliberately extracted from App.tsx so that "every
 * function receives its runtime and state explicitly" — which makes them
 * testable without a DOM, a canvas or a real Univer instance, provided
 * something supplies that runtime. This is that something.
 *
 * It is a double, not an emulator: methods return what they were configured to
 * return and record what they were called with. Anything a test does not
 * configure returns the empty/neutral value, so a test states only the part of
 * the world it cares about.
 */

import { parseAddress } from '../../src/domain/cell-address'

export interface FakeRangeOptions {
  values?: unknown[][]
  displayValues?: string[][]
  formulas?: (string | null)[][]
  a1?: string
}

export interface FakeSheetOptions {
  id?: string
  name?: string
  rows?: number
  columns?: number
  /** cell values keyed 'row:column' */
  cells?: Record<string, unknown>
  formulas?: Record<string, string>
  /** IStyleData per cell, keyed 'row:column' */
  styles?: Record<string, unknown>
  numberFormats?: Record<string, string>
  merges?: { row: number; column: number; width: number; height: number }[]
  freeze?: { xSplit: number; ySplit: number } | null
  filter?: { startRow: number; startColumn: number; endRow: number; endColumn: number } | null
  notes?: { row: number; column: number; note: string }[]
  conditionalRules?: unknown[]
  dataValidations?: unknown[]
  columnWidth?: number
  rowHeight?: number
}

export function fakeSheet(options: FakeSheetOptions = {}) {
  const {
    id = 'sheet-1',
    name = 'Sheet1',
    rows = 100,
    columns = 26,
    cells = {},
    formulas = {},
    styles = {},
    numberFormats = {},
    merges = [],
    freeze = null,
    filter = null,
    notes = [],
    conditionalRules = [],
    dataValidations = [],
    columnWidth = 80,
    rowHeight = 20,
  } = options

  const cellAt = (row: number, column: number): unknown => cells[`${row}:${column}`] ?? null

  const range = (row: number, column: number, numRows = 1, numColumns = 1) => ({
    getRow: () => row,
    getColumn: () => column,
    getWidth: () => numColumns,
    getHeight: () => numRows,
    getValue: () => cellAt(row, column),
    getValues: () =>
      Array.from({ length: numRows }, (_, r) =>
        Array.from({ length: numColumns }, (_, c) => cellAt(row + r, column + c)),
      ),
    getDisplayValues: () =>
      Array.from({ length: numRows }, (_, r) =>
        Array.from({ length: numColumns }, (_, c) => {
          const value = cellAt(row + r, column + c)
          return value === null || value === undefined ? '' : String(value)
        }),
      ),
    getFormula: () => formulas[`${row}:${column}`] ?? '',
    getFormulas: () =>
      Array.from({ length: numRows }, (_, r) =>
        Array.from({ length: numColumns }, (_, c) => formulas[`${row + r}:${column + c}`] ?? ''),
      ),
    getCellStyleData: () => styles[`${row}:${column}`] ?? null,
    getNumberFormat: () => numberFormats[`${row}:${column}`] ?? '',
    getA1Notation: () => `R${row + 1}C${column + 1}`,
    getRange: () => ({
      startRow: row,
      startColumn: column,
      endRow: row + numRows - 1,
      endColumn: column + numColumns - 1,
    }),
  })

  const sheet = {
    getSheetId: () => id,
    getSheetName: () => name,
    getMaxRows: () => rows,
    getMaxColumns: () => columns,
    getColumnWidth: () => columnWidth,
    getRowHeight: () => rowHeight,
    getZoom: () => 1,
    setRowCount: () => undefined,
    setColumnCount: () => undefined,
    getMergedRanges: () =>
      merges.map((m) => ({
        getRow: () => m.row,
        getColumn: () => m.column,
        getWidth: () => m.width,
        getHeight: () => m.height,
        getA1Notation: () => `R${m.row + 1}C${m.column + 1}:R${m.row + m.height}C${m.column + m.width}`,
      })),
    getFreeze: () => freeze,
    getFilter: () =>
      filter
        ? {
            getRange: () => ({ getRange: () => filter }),
            getColumnFilterCriteria: () => null,
          }
        : null,
    getNotes: () => notes,
    getConditionalFormattingRules: () => conditionalRules,
    getDataValidations: () => dataValidations,
    getSheet: () => ({ getConfig: () => ({ defaultColumnWidth: columnWidth, defaultRowHeight: rowHeight }) }),
    // A1 strings resolve to the cell they name — several readers address the
    // grid that way, and a double that ignored it would silently read A1.
    getRange: (a: number | string, b?: number, c?: number, d?: number) => {
      if (typeof a !== 'string') return range(a, b ?? 0, c ?? 1, d ?? 1)
      const [start = a] = a.replace(/\$/g, '').split(':')
      const { row, column } = parseAddress(start)
      return range(row, column, 1, 1)
    },
    addFloatDomToRange: () => ({ dispose: () => undefined }),
  }
  return sheet
}

export type FakeSheet = ReturnType<typeof fakeSheet>

export function fakeWorkbook(sheets: FakeSheet[] = [fakeSheet()], definedNames: unknown[] = []) {
  const active = sheets[0]
  return {
    getActiveSheet: () => active,
    getSheets: () => sheets,
    getSheetBySheetId: (id: string) => sheets.find((s) => s.getSheetId() === id) ?? null,
    getDefinedNames: () => definedNames,
    getActiveRange: () => active?.getRange(0, 0, 1, 1) ?? null,
  }
}

/** The `univerRef.current` shape the renderer modules expect. */
export function fakeRuntime(workbook: ReturnType<typeof fakeWorkbook> | null = fakeWorkbook()) {
  return {
    univerAPI: {
      getActiveWorkbook: () => workbook,
    },
  }
}

/**
 * The minimum LazyWorkbookState the read/write helpers dereference. Lazy mode
 * is the one that matters: it is what a real opened file uses, while the
 * in-memory adapter path only serves the blank/demo workbook.
 */
export function fakeLazyState(
  sheets: { id: string; name: string; rowCount?: number; columnCount?: number }[] = [
    { id: 'sheet-1', name: 'Sheet1' },
  ],
) {
  return {
    file: {
      sessionId: '3f2504e0-4f89-11d3-9a0c-0305e82c3301',
      name: 'Book1.xlsx',
      // Collection fields the readers dereference without guarding; a real
      // WorkbookFile always has them, so the double must too.
      sheets: sheets.map((s) => ({
        rowCount: 100,
        columnCount: 26,
        hidden: false,
        pivotTables: [],
        pivotRanges: [],
        tables: [],
        comments: [],
        columnWidths: [],
        sparklineGroups: [],
        ...s,
      })),
      visuals: [],
    },
    generation: 1,
    loadedRanges: new Map<string, unknown>(),
    loadingKeys: new Map(),
    retryTimers: new Map(),
    appliedMerges: new Map(),
    appliedRowKeys: new Map(),
    appliedCfSheets: new Set(),
    appliedFilterSheets: new Set(),
    appliedDvSheets: new Set(),
    sheetProtections: new Map(),
    flags: { preloadComplete: true },
    editJournal: {
      visualAdds: [],
      visualEdits: new Map(),
      chartEdits: new Map(),
      pageSetup: new Map(),
      sheets: { hidden: new Map(), added: [], removed: new Set(), renamed: new Map() },
      sheetProtection: new Map(),
      definedNames: new Map(),
      cells: new Map(),
    },
  }
}

/** A ref wrapper, since every context takes `{ current }` refs rather than values. */
export const ref = <T,>(current: T): { readonly current: T } => ({ current })
