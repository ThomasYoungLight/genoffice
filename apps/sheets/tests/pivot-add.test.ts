/**
 * applyAiPivotAdd — the single largest function in the renderer's application
 * layer, and the one with the most ways to produce a file Excel will not open.
 *
 * A pivot is four coupled parts written at save time from what this function
 * decides, so almost every branch here is a guard that fails now rather than
 * letting the failure surface at ⌘S or, worse, on the user's next open. The
 * function had no tests.
 *
 * The success cases assert the *numbers*, not just that something was written:
 * a pivot that lays out correctly and aggregates wrongly is the failure mode
 * nobody notices.
 */
import { describe, expect, it } from 'vitest'

import { applyAiPivotAdd } from '../src/renderer/workbook-ops'
import { fakeLazyState, fakeRuntime, fakeSheet, fakeWorkbook, mutationsOf } from './helpers/fake-univer'

/// Region | Quarter | Amount, four data rows over two regions and two quarters.
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

function setup(cells: Record<string, unknown> = SOURCE) {
  const sheet = fakeSheet({ cells })
  const runtime = fakeRuntime(fakeWorkbook([sheet])) as never
  const state = fakeLazyState() as never
  return { runtime, state, sheet }
}

const pivot = (extra: Record<string, unknown> = {}) =>
  ({
    op: 'add_pivot',
    sheetId: 'sheet-1',
    sourceRange: 'A1:C5',
    targetCell: 'E1',
    rowFields: ['Region'],
    values: [{ field: 'Amount', agg: 'sum' }],
    ...extra,
  }) as never

describe('applyAiPivotAdd: the sheet must exist', () => {
  it('refuses an unknown source sheet', () => {
    const { runtime, state } = setup()
    expect(() => applyAiPivotAdd(runtime, state, pivot({ sheetId: 'nope' }))).toThrow(
      /unknown sheet: nope/i,
    )
  })

  it('refuses an unknown target sheet', () => {
    const { runtime, state } = setup()
    expect(() => applyAiPivotAdd(runtime, state, pivot({ targetSheetId: 'ghost' }))).toThrow(
      /unknown sheet: ghost/i,
    )
  })

  it('refuses a sheet added this session', () => {
    // A session-added sheet has no worksheet part yet, so the pivot's
    // relationships would point at nothing in the written file.
    const { runtime, state } = setup()
    ;(state as never as ReturnType<typeof fakeLazyState>).editJournal.sheets.added.add('sheet-1')
    expect(() => applyAiPivotAdd(runtime, state, pivot())).toThrow()
  })
})

describe('applyAiPivotAdd: the source has to be pivotable', () => {
  it('refuses a source with a header but no data', () => {
    const { runtime, state } = setup()
    expect(() => applyAiPivotAdd(runtime, state, pivot({ sourceRange: 'A1:C1' }))).toThrow()
  })

  it('refuses more source rows than the writer can hold', () => {
    const { runtime, state } = setup()
    expect(() => applyAiPivotAdd(runtime, state, pivot({ sourceRange: 'A1:C20000' }))).toThrow()
  })

  it('refuses more source columns than the writer can hold', () => {
    const { runtime, state } = setup()
    expect(() => applyAiPivotAdd(runtime, state, pivot({ sourceRange: 'A1:ZZ5' }))).toThrow()
  })

  it('refuses a blank header cell, which has no field name to reference', () => {
    const cells = { ...SOURCE, '0:1': '' }
    const { runtime, state } = setup(cells)
    expect(() => applyAiPivotAdd(runtime, state, pivot())).toThrow()
  })

  it('refuses duplicate headers, case-insensitively', () => {
    // Two "Region" columns make field references ambiguous; Excel resolves it
    // silently and differently from us, so refusing is the honest option.
    const cells = { ...SOURCE, '0:1': 'region' }
    const { runtime, state } = setup(cells)
    expect(() => applyAiPivotAdd(runtime, state, pivot())).toThrow()
  })

  it('refuses a row field that is not one of the headers', () => {
    const { runtime, state } = setup()
    expect(() => applyAiPivotAdd(runtime, state, pivot({ rowFields: ['Nope'] }))).toThrow(/nope/i)
  })

  it('refuses a value field that is not one of the headers', () => {
    const { runtime, state } = setup()
    expect(() =>
      applyAiPivotAdd(runtime, state, pivot({ values: [{ field: 'Ghost', agg: 'sum' }] })),
    ).toThrow(/ghost/i)
  })

  it('refuses a column field that is not one of the headers', () => {
    const { runtime, state } = setup()
    expect(() => applyAiPivotAdd(runtime, state, pivot({ columnField: 'Nope' }))).toThrow()
  })
})

describe('applyAiPivotAdd: placement', () => {
  it('refuses a target that overlaps its own source', () => {
    // The pivot would overwrite the data it aggregates, and then recalculate
    // from what it just destroyed.
    const { runtime, state } = setup()
    expect(() => applyAiPivotAdd(runtime, state, pivot({ targetCell: 'B2' }))).toThrow()
  })

  it('accepts a target clear of the source', () => {
    const { runtime, state } = setup()
    expect(() => applyAiPivotAdd(runtime, state, pivot())).not.toThrow()
  })

  it('refuses a second pivot with a name already taken', () => {
    const { runtime, state } = setup()
    applyAiPivotAdd(runtime, state, pivot({ name: 'Report' }))
    expect(() =>
      applyAiPivotAdd(runtime, state, pivot({ name: 'Report', targetCell: 'J1' })),
    ).toThrow()
  })

  it('refuses a second pivot that would overlap the first', () => {
    const { runtime, state } = setup()
    applyAiPivotAdd(runtime, state, pivot())
    expect(() => applyAiPivotAdd(runtime, state, pivot())).toThrow()
  })
})

describe('applyAiPivotAdd: what it produces', () => {
  it('journals the pivot so the save has something to write', () => {
    const { runtime, state } = setup()
    applyAiPivotAdd(runtime, state, pivot())
    const journal = (state as never as ReturnType<typeof fakeLazyState>).editJournal as {
      pivotAdds?: unknown[]
    }
    expect(journal.pivotAdds?.length ?? 0).toBeGreaterThan(0)
  })

  it('writes the result onto the target sheet', () => {
    const { runtime, state, sheet } = setup()
    applyAiPivotAdd(runtime, state, pivot())
    const writes = mutationsOf(sheet).filter(([name]) => String(name).startsWith('setValue'))
    expect(writes.length).toBeGreaterThan(0)
  })

  it('sums each row group rather than counting or repeating rows', () => {
    // North = 10 + 20, South = 30 + 40. Getting the layout right and the
    // arithmetic wrong is the failure nobody notices.
    const { runtime, state, sheet } = setup()
    applyAiPivotAdd(runtime, state, pivot())
    const written = JSON.stringify(mutationsOf(sheet))
    expect(written).toContain('30')
    expect(written).toContain('70')
  })

  it('counts rows when asked to count, not sum them', () => {
    const { runtime, state, sheet } = setup()
    applyAiPivotAdd(runtime, state, pivot({ values: [{ field: 'Amount', agg: 'count' }] }))
    const written = JSON.stringify(mutationsOf(sheet))
    expect(written).toContain('2')
  })

  it('accepts a column field, giving a two-dimensional report', () => {
    const { runtime, state, sheet } = setup()
    applyAiPivotAdd(runtime, state, pivot({ columnField: 'Quarter' }))
    const written = JSON.stringify(mutationsOf(sheet))
    expect(written).toContain('Q1')
    expect(written).toContain('Q2')
  })

  it('refuses a calculated field whose name clashes with a source header', () => {
    const { runtime, state } = setup()
    expect(() =>
      applyAiPivotAdd(
        runtime,
        state,
        pivot({ values: [{ field: 'Amount', agg: 'sum', formula: 'Amount*2' }] }),
      ),
    ).toThrow()
  })

  it('accepts a calculated field with a fresh name', () => {
    const { runtime, state } = setup()
    expect(() =>
      applyAiPivotAdd(
        runtime,
        state,
        pivot({ values: [{ field: 'Double', agg: 'sum', formula: 'Amount*2' }] }),
      ),
    ).not.toThrow()
  })
})
