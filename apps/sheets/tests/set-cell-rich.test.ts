/**
 * set_cell_rich — a cell whose text is not uniformly formatted.
 *
 * The whole write path for rich text already existed, driven by what a user
 * types; the DSL was the only thing that could not reach it, so this is mostly
 * about the operation being translated into the same shape the manual path
 * produces. What matters:
 *
 *   - `value` keeps the whole string, so every reader that only wants text —
 *     search, expectedValue, the plan summary — is unaffected;
 *   - `rich` carries the run structure alongside it;
 *   - a run may set only what it changes, because "bold the second word" should
 *     not have to restate the font.
 */
import { describe, it, expect } from 'vitest'
import { workbookOperationSchema } from '../src/domain/workbook-dsl'
import { InMemoryWorkbookAdapter } from '../src/domain/in-memory-workbook'

const RUNS = [{ text: 'Revenue ' }, { text: 'up 12%', bold: true, color: '#107C41' }]

const parse = (op: unknown) => workbookOperationSchema.safeParse(op)
const richOp = (extra: Record<string, unknown> = {}) => ({
  op: 'set_cell_rich',
  sheetId: 's1',
  address: 'A1',
  runs: RUNS,
  ...extra,
})

describe('set_cell_rich schema', () => {
  it('accepts runs that set only what they change', () => {
    const result = parse(richOp())
    expect(result.success).toBe(true)
  })

  it('needs at least one run — an empty rich cell is just clear_cell', () => {
    expect(parse(richOp({ runs: [] })).success).toBe(false)
  })

  it('caps the run count, so one cell cannot carry a document', () => {
    expect(parse(richOp({ runs: Array.from({ length: 65 }, () => ({ text: 'x' })) })).success).toBe(
      false,
    )
  })

  it('rejects a colour Excel would not accept', () => {
    expect(parse(richOp({ runs: [{ text: 'x', color: 'green' }] })).success).toBe(false)
    expect(parse(richOp({ runs: [{ text: 'x', color: '#107C41' }] })).success).toBe(true)
  })

  it('rejects an unknown run field rather than dropping it silently', () => {
    expect(parse(richOp({ runs: [{ text: 'x', bolded: true }] })).success).toBe(false)
  })
})

describe('set_cell_rich planning', () => {
  const adapter = (): InMemoryWorkbookAdapter =>
    new InMemoryWorkbookAdapter({
      revision: 0,
      sheets: [{ id: 's1', name: 'Sheet1', cells: {} }],
    })

  it('puts the joined text in value and the runs in rich', () => {
    const plan = adapter().plan({
      dslVersion: 1,
      transactionId: 't1',
      baseRevision: 0,
      summary: 'rich cell',
      operations: [richOp()],
    })
    const change = plan.cellChanges[0]!
    expect(change.address).toBe('A1')
    // the text is the concatenation, in order
    expect(change.after.value).toBe('Revenue up 12%')
    expect(change.after.rich).toEqual(RUNS)
  })

  it('does not clear the cell, which is what an unhandled op would have done', () => {
    const plan = adapter().plan({
      dslVersion: 1,
      transactionId: 't2',
      baseRevision: 0,
      summary: 'rich cell',
      operations: [richOp()],
    })
    expect(plan.cellChanges[0]!.after.value).not.toBeNull()
  })
})
