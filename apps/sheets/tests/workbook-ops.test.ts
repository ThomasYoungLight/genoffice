/**
 * Table and pivot operation appliers — the code that turns a validated DSL
 * operation into calls against the live grid and entries in the edit journal.
 *
 * Two halves worth separating. The guards decide whether an operation happens
 * at all, and each rejects something that would otherwise produce a file Excel
 * dislikes (a table with no data rows, a duplicate table name). The pure
 * helpers turn a pivot definition into labels, where an off-by-one shows up as
 * a mislabelled report rather than an error.
 */
import { describe, expect, it } from 'vitest'

import {
  applyAiTableAdd,
  applyAiTableRowAdd,
  pivotColLineLabel,
  pivotMemberCaption,
} from '../src/renderer/workbook-ops'
import {
  fakeLazyState,
  fakeRuntime,
  fakeSheet,
  fakeWorkbook,
  mutationsOf,
  type FakeSheet,
} from './helpers/fake-univer'

function setup(sheet: FakeSheet = fakeSheet()) {
  const runtime = fakeRuntime(fakeWorkbook([sheet])) as never
  const state = fakeLazyState() as never
  return { runtime, state, sheet }
}

const addTable = (extra: Record<string, unknown> = {}) =>
  ({ op: 'add_table', sheetId: 'sheet-1', range: 'A1:C5', ...extra }) as never

describe('applyAiTableAdd guards', () => {
  it('refuses an unknown sheet by name rather than falling back to the active one', () => {
    const { runtime, state } = setup()
    expect(() => applyAiTableAdd(runtime, state, addTable({ sheetId: 'nope' }))).toThrow(
      /unknown sheet: nope/i,
    )
  })

  it('refuses a table with a header row but no data rows', () => {
    // Excel treats the first row as the header, so a single-row range is a
    // table of nothing — valid to write and useless to open.
    const { runtime, state } = setup()
    expect(() => applyAiTableAdd(runtime, state, addTable({ range: 'A1:C1' }))).toThrow()
  })

  it('refuses an absurdly wide table instead of building it', () => {
    const { runtime, state } = setup()
    expect(() => applyAiTableAdd(runtime, state, addTable({ range: 'A1:AZZ100' }))).toThrow()
  })

  it('accepts a well-formed range and journals the table', () => {
    const { runtime, state } = setup()
    applyAiTableAdd(runtime, state, addTable())
    expect((state as never as ReturnType<typeof fakeLazyState>).editJournal.tableAdds).toHaveLength(
      1,
    )
  })

  it('refuses a second table with the same name, case-insensitively', () => {
    // Excel table names are case-insensitive and must be unique in a workbook;
    // two "Sales" tables makes a file that opens with a repair prompt.
    const { runtime, state } = setup()
    applyAiTableAdd(runtime, state, addTable({ name: 'Sales' }))
    expect(() =>
      applyAiTableAdd(runtime, state, addTable({ name: 'sales', range: 'E1:G5' })),
    ).toThrow()
  })

  it('names a table itself when the operation does not', () => {
    const { runtime, state } = setup()
    applyAiTableAdd(runtime, state, addTable())
    const [entry] = (state as never as ReturnType<typeof fakeLazyState>).editJournal
      .tableAdds as { name: string }[]
    expect(entry?.name).toBeTruthy()
  })
})

describe('applyAiTableRowAdd', () => {
  const addRow = (extra: Record<string, unknown> = {}) =>
    ({ op: 'add_table_row', sheetId: 'sheet-1', tableName: 'Sales', ...extra }) as never

  function withTable() {
    const { runtime, state, sheet } = setup()
    applyAiTableAdd(runtime, state, addTable({ name: 'Sales' }))
    return { runtime, state, sheet }
  }

  it('refuses a table that is not in the session journal', () => {
    // Only session-added tables can be grown this way; a table from the file
    // has to go through the gateway, and silently doing nothing would be worse.
    const { runtime, state } = setup()
    expect(() => applyAiTableRowAdd(runtime, state, addRow())).toThrow(/not found/i)
  })

  it('appends at the end of the data region by default', () => {
    const { runtime, state, sheet } = withTable()
    applyAiTableRowAdd(runtime, state, addRow())
    expect(mutationsOf(sheet).some(([name]) => name === 'insertRowsBefore')).toBe(true)
  })

  it('grows the journalled area so the table keeps describing itself', () => {
    // The saved table XML comes from this area; if it does not grow, the new
    // rows sit outside the table in the written file.
    const { runtime, state } = withTable()
    const journal = (state as never as ReturnType<typeof fakeLazyState>).editJournal
    const before = (journal.tableAdds as { area: { endRow: number } }[])[0]!.area.endRow
    applyAiTableRowAdd(runtime, state, addRow({ count: 3 }))
    const after = (journal.tableAdds as { area: { endRow: number } }[])[0]!.area.endRow
    expect(after).toBe(before + 3)
  })

  it('matches the table name case-insensitively, as Excel does', () => {
    const { runtime, state } = withTable()
    expect(() => applyAiTableRowAdd(runtime, state, addRow({ tableName: 'SALES' }))).not.toThrow()
  })
})

describe('pivot label helpers', () => {
  const definition = {
    fields: [{ sharedItems: ['North', 'South'] }, { sharedItems: ['Q1', 'Q2'] }],
    fieldItems: [
      [{ x: 0 }, { x: 1 }, { x: null }],
      [{ x: 0 }, { x: 1 }],
    ],
    colFields: [1],
    colLines: [],
  } as never

  it('resolves a member to its shared-item caption', () => {
    expect(pivotMemberCaption(definition, 0, 0)).toBe('North')
    expect(pivotMemberCaption(definition, 0, 1)).toBe('South')
  })

  it('returns empty for a member with no shared-item index', () => {
    // x === null is how a pivot marks a blank/derived member; guessing a
    // caption for it would put a wrong label in the report.
    expect(pivotMemberCaption(definition, 0, 2)).toBe('')
  })

  it('returns empty rather than throwing for an out-of-range field or member', () => {
    expect(pivotMemberCaption(definition, 99, 0)).toBe('')
    expect(pivotMemberCaption(definition, 0, 99)).toBe('')
  })

  it('labels a grand-total column line', () => {
    expect(pivotColLineLabel(definition, { t: 'grand' } as never)).toBe('Grand Total')
  })

  it('labels a blank column line as empty', () => {
    expect(pivotColLineLabel(definition, { t: 'blank' } as never)).toBe('')
  })
})
