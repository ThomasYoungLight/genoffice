/**
 * The gate between what the agent proposes and what reaches the workbook.
 *
 * Everything here is a refusal path, and each one exists because the
 * alternative is worse than an error: a plan that applies cleanly and then
 * fails at ⌘S, or a find_replace that silently misses rows because they were
 * never loaded. The file had no tests, which is a poor place for that to be
 * true — it is the only thing standing between a confident wrong answer from a
 * model and the user's file.
 */
import { describe, expect, it, vi } from 'vitest'

import { proposeOperations } from '../src/renderer/plan-operations'
import { fakeLazyState, fakeRuntime, fakeSheet, fakeWorkbook, ref } from './helpers/fake-univer'

function context(overrides: { lazy?: unknown; sheet?: ReturnType<typeof fakeSheet> } = {}) {
  const sheet = overrides.sheet ?? fakeSheet()
  const setPreview = vi.fn()
  const autoApplySafePlan = vi.fn()
  const ctx = {
    adapterRef: ref({ getSnapshot: () => ({ sheets: [] }) }),
    univerRef: ref(fakeRuntime(fakeWorkbook([sheet]))),
    lazyWorkbookRef: ref(overrides.lazy === undefined ? fakeLazyState() : overrides.lazy),
    lazyPreviewRef: ref(null),
    setPreview,
    autoApplySafePlan,
  } as never
  return { ctx, setPreview, autoApplySafePlan }
}

const setCell = (value: string) => [
  { op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value } as const,
]

describe('proposeOperations', () => {
  it('refuses when no workbook is open rather than queueing work into the void', () => {
    const ctx = {
      adapterRef: ref({ getSnapshot: () => ({ sheets: [] }) }),
      univerRef: ref(fakeRuntime(null)),
      lazyWorkbookRef: ref(fakeLazyState()),
      lazyPreviewRef: ref(null),
      setPreview: vi.fn(),
      autoApplySafePlan: vi.fn(),
    } as never
    expect(proposeOperations(ctx, setCell('x'), 'set A1')).toEqual({
      ok: false,
      error: 'No workbook is open.',
    })
  })

  it('rejects operations that are not valid DSL', () => {
    const { ctx } = context()
    const result = proposeOperations(ctx, [{ op: 'not_a_real_op' }] as never, 'nonsense')
    expect(result.ok).toBe(false)
  })

  it('accepts a well-formed operation and produces a plan', () => {
    const { ctx } = context()
    const result = proposeOperations(ctx, setCell('hello'), 'set A1')
    expect(result.ok).toBe(true)
    if (result.ok) expect(result.plan).toBeTruthy()
  })

  describe('find_replace has to be able to see the cells it edits', () => {
    const findReplace = [
      {
        op: 'find_replace',
        sheetId: 'sheet-1',
        range: 'A1:C10',
        find: 'a',
        replace: 'b',
      } as const,
    ]

    it('refuses a range that has not been loaded, and says what to do instead', () => {
      // Unloaded cells read as empty, so the replace would report success while
      // silently skipping rows — the worst available outcome.
      const { ctx } = context()
      const result = proposeOperations(ctx, findReplace, 'replace')
      expect(result.ok).toBe(false)
      if (!result.ok) {
        expect(result.error).toMatch(/not fully loaded/i)
        expect(result.error).toMatch(/read_range|set_range/)
      }
    })

    it('allows it once the range is loaded', () => {
      const lazy = fakeLazyState()
      lazy.loadedRanges.set('sheet-1', {
        startRow: 0,
        endRow: 100,
        startColumn: 0,
        endColumn: 20,
      })
      const { ctx } = context({ lazy })
      expect(proposeOperations(ctx, findReplace, 'replace').ok).toBe(true)
    })

    it('allows it in formula mode, where the whole grid is present', () => {
      const lazy = { ...fakeLazyState(), formulaMode: true }
      const { ctx } = context({ lazy })
      expect(proposeOperations(ctx, findReplace, 'replace').ok).toBe(true)
    })

    it('refuses a range that only partly overlaps what is loaded', () => {
      const lazy = fakeLazyState()
      lazy.loadedRanges.set('sheet-1', { startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 })
      const { ctx } = context({ lazy })
      expect(proposeOperations(ctx, findReplace, 'replace').ok).toBe(false)
    })
  })

  describe('edit_chart fails closed here rather than at save time', () => {
    const editChart = (extra: Record<string, unknown> = {}) => [
      { op: 'edit_chart', chartPath: 'xl/charts/chart1.xml', title: 'T', ...extra } as never,
    ]

    const withChart = (chart: unknown) => {
      const lazy = fakeLazyState()
      ;(lazy.file.visuals as unknown[]).push({
        id: 'v1',
        kind: 'chart',
        sheetId: 'sheet-1',
        chartPath: 'xl/charts/chart1.xml',
        chart,
      })
      return lazy
    }

    it('refuses a chart it cannot find', () => {
      const { ctx } = context()
      const result = proposeOperations(ctx, editChart(), 'retitle')
      expect(result.ok).toBe(false)
      if (!result.ok) expect(result.error).toMatch(/unknown chart/i)
    })

    it('refuses a type change on a chart that cannot be converted', () => {
      // The save path fails closed on multi-plot charts. Catching it here is
      // the difference between an error now and Apply succeeding then ⌘S failing.
      const lazy = withChart({ chartTypes: ['barChart', 'lineChart'], series: [], title: '' })
      const { ctx } = context({ lazy })
      const result = proposeOperations(ctx, editChart({ chartType: 'pie' }), 'convert')
      expect(result.ok).toBe(false)
      if (!result.ok) expect(result.error).toMatch(/cannot be converted/i)
    })

    it('refuses axis titles on a pie chart, which has no axes', () => {
      const lazy = withChart({ chartTypes: ['pieChart'], series: [], title: '' })
      const { ctx } = context({ lazy })
      const result = proposeOperations(
        ctx,
        editChart({ axisTitles: { category: 'x' } }),
        'axis titles',
      )
      expect(result.ok).toBe(false)
      if (!result.ok) expect(result.error).toMatch(/no axes/i)
    })

    it('allows a plain retitle of a convertible chart', () => {
      const lazy = withChart({ chartTypes: ['barChart'], series: [], title: 'old' })
      const { ctx } = context({ lazy })
      expect(proposeOperations(ctx, editChart(), 'retitle').ok).toBe(true)
    })

    it('finds a session-added chart by visual id, not just by chart part path', () => {
      // A chart added this session has no part yet; its id doubles as the path.
      const lazy = fakeLazyState()
      ;(lazy.editJournal.visualAdds as unknown[]).push({
        id: 'added-chart-1',
        kind: 'chart',
        sheetId: 'sheet-1',
        chart: { chartTypes: ['barChart'], series: [], title: 'x' },
      })
      const { ctx } = context({ lazy })
      const result = proposeOperations(
        ctx,
        [{ op: 'edit_chart', chartPath: 'added-chart-1', title: 'T' } as never],
        'retitle',
      )
      expect(result.ok).toBe(true)
    })
  })
})
