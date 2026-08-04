/**
 * Pivot and visual insertion actions — the other two modules the ribbon
 * dispatcher hands off to.
 *
 * The pivot side is mostly selection interpretation: what range did the user
 * mean, which fields does it offer, are they standing inside an existing
 * pivot. Getting that wrong builds a correct pivot over the wrong data, which
 * looks like a working feature.
 *
 * The visual side inserts charts, shapes and images into the journal, and its
 * guards are about not inserting into a workbook that cannot hold them.
 */
import { describe, expect, it, vi } from 'vitest'

import {
  getSourceRange,
  isSelectionInPivot,
  pivotFieldOptions,
} from '../src/renderer/pivot-actions'
import { handleInsertShape } from '../src/renderer/visual-actions'
import { fakeLazyState, fakeRuntime, fakeSheet, fakeWorkbook, ref } from './helpers/fake-univer'

/// A workbook whose active range is a specific block of the sheet.
function withSelection(
  row: number,
  column: number,
  height: number,
  width: number,
  cells: Record<string, unknown> = {},
) {
  const sheet = fakeSheet({ cells })
  const workbook = fakeWorkbook([sheet])
  const selection = sheet.getRange(row, column, height, width)
  return fakeRuntime({ ...workbook, getActiveRange: () => selection } as never)
}

function pivotContext(runtime: unknown, lazy: unknown = fakeLazyState()) {
  return {
    univerRef: ref(runtime),
    lazyWorkbookRef: ref(lazy),
    setMessage: vi.fn(),
    setPendingEdits: vi.fn(),
  } as never
}

describe('getSourceRange reads the selection back as a reference', () => {
  it('describes a multi-cell selection in A1 notation', () => {
    expect(getSourceRange(pivotContext(withSelection(0, 0, 5, 3)))).toBe('A1:C5')
  })

  it('describes a selection that does not start at the origin', () => {
    expect(getSourceRange(pivotContext(withSelection(2, 1, 3, 2)))).toBe('B3:C5')
  })

  it('describes a single cell as a degenerate range rather than failing', () => {
    expect(getSourceRange(pivotContext(withSelection(0, 0, 1, 1)))).toBe('A1:A1')
  })

  it('returns empty with nothing selected, so the dialog can prompt', () => {
    expect(getSourceRange(pivotContext(fakeRuntime(null)))).toBe('')
  })
})

describe('pivotFieldOptions offers the headers of the selection', () => {
  const headers = { '0:0': 'Region', '0:1': 'Quarter', '0:2': 'Amount' }

  it('reads the first row as field labels, with their column indexes', () => {
    expect(pivotFieldOptions(pivotContext(withSelection(0, 0, 5, 3, headers)))).toEqual([
      { label: 'Region', colIndex: 0 },
      { label: 'Quarter', colIndex: 1 },
      { label: 'Amount', colIndex: 2 },
    ])
  })

  it('offers nothing for a selection with no data rows under the header', () => {
    // One row is a header with nothing to aggregate; offering fields would
    // invite a pivot over an empty body.
    expect(pivotFieldOptions(pivotContext(withSelection(0, 0, 1, 3, headers)))).toEqual([])
  })

  it('offers nothing when there is no selection at all', () => {
    expect(pivotFieldOptions(pivotContext(fakeRuntime(null)))).toEqual([])
  })

  it('keeps the column index relative to the sheet, not the selection', () => {
    // The pivot reads its source by absolute column; a selection-relative
    // index would silently aggregate the wrong column.
    expect(
      pivotFieldOptions(pivotContext(withSelection(0, 2, 5, 1, { '0:2': 'Amount' }))),
    ).toEqual([{ label: 'Amount', colIndex: 2 }])
  })
})

describe('isSelectionInPivot', () => {
  it('is false when nothing is open', () => {
    expect(isSelectionInPivot(pivotContext(fakeRuntime(null), null))).toBe(false)
  })

  it('is false for a plain selection on a sheet with no pivots', () => {
    expect(isSelectionInPivot(pivotContext(withSelection(0, 0, 2, 2)))).toBe(false)
  })
})

describe('handleInsertShape', () => {
  function visualContext(runtime: unknown, lazy: unknown = fakeLazyState()) {
    const setMessage = vi.fn()
    const ctx = {
      adapterRef: ref({ getSnapshot: () => ({ sheets: [] }) }),
      univerRef: ref(runtime),
      lazyWorkbookRef: ref(lazy),
      visualDisposablesRef: ref([]),
      visualInstallTimerRef: ref(null),
      chartEditRef: ref(vi.fn()),
      chartVectorRef: ref(vi.fn()),
      shapeEditRef: ref(vi.fn()),
      setMessage,
      setPendingEdits: vi.fn(),
    } as never
    return { ctx, setMessage }
  }

  it('does not throw with no workbook open', () => {
    const { ctx } = visualContext(fakeRuntime(null), null)
    expect(() => handleInsertShape(ctx, 'rect', false)).not.toThrow()
  })

  it('inserts a shape into the journal', () => {
    const lazy = fakeLazyState()
    const { ctx } = visualContext(fakeRuntime(fakeWorkbook([fakeSheet()])), lazy)
    handleInsertShape(ctx, 'rect', false)
    expect((lazy.editJournal.visualAdds as unknown[]).length).toBeGreaterThan(0)
  })

  it('inserts a text box through the same path', () => {
    const lazy = fakeLazyState()
    const { ctx } = visualContext(fakeRuntime(fakeWorkbook([fakeSheet()])), lazy)
    handleInsertShape(ctx, 'rect', true)
    expect((lazy.editJournal.visualAdds as unknown[]).length).toBeGreaterThan(0)
  })

  it.each(['rect', 'roundRect', 'ellipse', 'triangle', 'diamond', 'hexagon'])(
    'accepts the %s shape the ribbon offers',
    (shapeType) => {
      const { ctx } = visualContext(fakeRuntime(fakeWorkbook([fakeSheet()])))
      expect(() => handleInsertShape(ctx, shapeType, false)).not.toThrow()
    },
  )
})
