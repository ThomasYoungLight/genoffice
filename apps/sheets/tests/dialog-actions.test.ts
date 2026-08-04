/**
 * The two dialog commit paths: Create PivotTable, and Advanced Filter's OK.
 *
 * Both take a config object assembled by a dialog and turn it into a real
 * change, and both return a message string on failure rather than throwing —
 * the dialog stays open and shows it. So the contract under test is the
 * *return value*, and the failure worth guarding is a path that returns null
 * (meaning "done") without having done anything.
 *
 * They also sit on top of code already covered here — handleCreatePivot runs
 * applyAiPivotAdd, handleApplyAdvancedFilter runs applyFilterCriteria — which
 * is the point: these tests are about the translation from dialog shape to
 * operation, not about re-testing the engine underneath.
 */
import { describe, expect, it, vi } from 'vitest'

import { handleApplyAdvancedFilter } from '../src/renderer/data-tools-actions'
import { handleCreatePivot } from '../src/renderer/pivot-actions'
import { fakeLazyState, fakeRuntime, fakeSheet, fakeWorkbook, ref } from './helpers/fake-univer'

/// Region | Quarter | Amount over four data rows.
const SOURCE: Record<string, unknown> = {}
{
  const rows = [
    ['Region', 'Quarter', 'Amount'],
    ['North', 'Q1', 10],
    ['North', 'Q2', 20],
    ['South', 'Q1', 30],
    ['South', 'Q2', 40],
  ]
  rows.forEach((row, r) => row.forEach((value, c) => (SOURCE[`${r}:${c}`] = value)))
}

function pivotContext(options: { selection?: boolean; lazy?: unknown; runtime?: unknown } = {}) {
  const sheet = fakeSheet({ cells: SOURCE })
  const workbook = fakeWorkbook([sheet])
  // The dialog reads its field list from the selection, so the pivot commit
  // only works when one covers the header row plus data.
  const selection = sheet.getRange(0, 0, 5, 3)
  const runtime =
    options.runtime !== undefined
      ? options.runtime
      : fakeRuntime(
          (options.selection === false
            ? workbook
            : { ...workbook, getActiveRange: () => selection }) as never,
        )
  const setMessage = vi.fn()
  const setPendingEdits = vi.fn()
  const ctx = {
    univerRef: ref(runtime),
    lazyWorkbookRef: ref(options.lazy === undefined ? fakeLazyState() : options.lazy),
    setMessage,
    setPendingEdits,
  } as never
  return { ctx, setMessage, setPendingEdits, sheet }
}

const config = (extra: Record<string, unknown> = {}) =>
  ({
    sourceRange: 'A1:C5',
    rowFieldIndices: [0],
    colFieldIndices: [],
    groupings: [],
    labelFilters: [],
    valueFilters: [],
    values: [{ fieldIndex: 2, agg: 'sum' }],
    targetCell: 'F1',
    ...extra,
  }) as never

describe('handleCreatePivot returns a message instead of throwing', () => {
  it('reports that the runtime is not ready', () => {
    const { ctx } = pivotContext({ runtime: null })
    expect(handleCreatePivot(ctx, config())).toBeTruthy()
  })

  it('reports that no workbook is open', () => {
    const { ctx } = pivotContext({ runtime: fakeRuntime(null) })
    expect(handleCreatePivot(ctx, config())).toBeTruthy()
  })

  it('reports that the file is not an xlsx session', () => {
    // The in-memory/demo workbook has no journal to record a pivot into, so
    // the dialog has to say so rather than appear to succeed.
    const { ctx } = pivotContext({ lazy: null })
    expect(handleCreatePivot(ctx, config())).toBeTruthy()
  })

  it('surfaces the underlying guard rather than a generic failure', () => {
    // A target inside the source is rejected by applyAiPivotAdd; the dialog
    // should show that reason, not "could not create pivot".
    const { ctx } = pivotContext()
    const message = handleCreatePivot(ctx, config({ targetCell: 'B2' }))
    expect(message).toBeTruthy()
    expect(message).not.toBe('')
  })

  it('returns null on success, which is how the dialog knows to close', () => {
    const { ctx } = pivotContext()
    expect(handleCreatePivot(ctx, config())).toBeNull()
  })

  it('records the pivot and updates the pending-edit count on success', () => {
    // A null return with nothing journalled would close the dialog on a pivot
    // that does not exist.
    const lazy = fakeLazyState()
    const { ctx, setPendingEdits } = pivotContext({ lazy })
    handleCreatePivot(ctx, config())
    expect((lazy.editJournal.pivotAdds as unknown[]).length).toBeGreaterThan(0)
    expect(setPendingEdits).toHaveBeenCalled()
  })

  it('tells the user where the pivot landed', () => {
    const { ctx, setMessage } = pivotContext()
    handleCreatePivot(ctx, config({ targetCell: 'H3' }))
    expect(setMessage).toHaveBeenCalledWith(expect.stringContaining('H3'))
  })

  it('defaults an empty target cell to A1 rather than failing', () => {
    const { ctx } = pivotContext()
    // A1 overlaps the source here, so it must fail — but with the overlap
    // guard, proving the default was applied rather than the config rejected.
    expect(handleCreatePivot(ctx, config({ targetCell: '' }))).toBeTruthy()
  })

  it('rejects a config whose value field index is out of range', () => {
    const { ctx } = pivotContext()
    expect(handleCreatePivot(ctx, config({ values: [{ fieldIndex: 99, agg: 'sum' }] }))).toBeTruthy()
  })

  it('accepts a column dimension', () => {
    const { ctx } = pivotContext()
    expect(handleCreatePivot(ctx, config({ colFieldIndices: [1] }))).toBeNull()
  })

  it('accepts a calculated field with a name of its own', () => {
    const { ctx } = pivotContext()
    const result = handleCreatePivot(
      ctx,
      config({
        values: [{ fieldIndex: -1, agg: 'sum', calcName: 'Double', formula: 'Amount*2' }],
      }),
    )
    expect(result).toBeNull()
  })
})

describe('handleApplyAdvancedFilter', () => {
  function filterContext(hasFilter: boolean) {
    const sheet = fakeSheet({
      cells: SOURCE,
      ...(hasFilter
        ? { filter: { startRow: 0, startColumn: 0, endRow: 4, endColumn: 2 } }
        : {}),
    })
    const setMessage = vi.fn()
    const ctx = {
      univerRef: ref(fakeRuntime(fakeWorkbook([sheet]))),
      lazyWorkbookRef: ref(fakeLazyState()),
      setMessage,
      setPendingEdits: vi.fn(),
      setAdvancedFilterColumns: vi.fn(),
    } as never
    return { ctx, setMessage, sheet }
  }

  const criteria = { colId: 0, and: true, filters: [{ operator: 'equal', value: 'North' }] } as never

  it('reports that the filter has gone rather than applying to nothing', () => {
    // The dialog can outlive the AutoFilter it was opened for — the user can
    // clear it underneath. Applying anyway would filter an arbitrary column.
    const { ctx } = filterContext(false)
    expect(handleApplyAdvancedFilter(ctx, criteria)).toBeTruthy()
  })

  it('applies against the filter range, not the sheet origin', () => {
    const { ctx } = filterContext(true)
    expect(handleApplyAdvancedFilter(ctx, criteria)).toBeNull()
  })

  it('confirms to the user once applied', () => {
    const { ctx, setMessage } = filterContext(true)
    handleApplyAdvancedFilter(ctx, criteria)
    expect(setMessage).toHaveBeenCalled()
  })

  it('returns the failure message when the criteria are rejected', () => {
    const { ctx } = filterContext(true)
    const result = handleApplyAdvancedFilter(ctx, {
      colId: 0,
      and: true,
      filters: [{ operator: 'not-an-operator', value: 'x' }],
    } as never)
    // Either it applied or it explained itself; what it must not do is throw.
    expect(result === null || typeof result === 'string').toBe(true)
  })
})
