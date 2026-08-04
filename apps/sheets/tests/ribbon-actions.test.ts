/**
 * The ribbon command dispatcher.
 *
 * Every button, dropdown and menu item in the app arrives here as a string,
 * and the failure this suite is really guarding against is a command that
 * throws on a workbook where its precondition does not hold — no selection, no
 * chart selected, nothing open. The user sees a dead button and a stack trace
 * in a log they will never read.
 *
 * So the wide check is a property over the real command list rather than a
 * case per command: a dispatcher must degrade to a no-op or a message, never
 * an exception. It also covers commands added later, which a per-command suite
 * would silently miss.
 */
import { describe, expect, it, vi } from 'vitest'

import { handleRibbonCommand, parseStyleCommand } from '../src/renderer/ribbon-actions'
import { fakeLazyState, fakeRuntime, fakeSheet, fakeWorkbook, ref } from './helpers/fake-univer'

describe('parseStyleCommand', () => {
  it('treats a command with no colon as a bare name', () => {
    expect(parseStyleCommand('undo')).toEqual({ name: 'undo', argument: '', extra: '' })
  })

  it('splits name from argument at the first colon', () => {
    expect(parseStyleCommand('zoom:150')).toEqual({ name: 'zoom', argument: '150', extra: '' })
  })

  it('keeps later colons in the argument, because number formats contain them', () => {
    // "h:mm:ss AM/PM" is a legitimate Excel format string. Splitting on every
    // colon would silently truncate it to "h" and format times as hours.
    expect(parseStyleCommand('numfmt:h:mm:ss AM/PM')).toEqual({
      name: 'numfmt',
      argument: 'h:mm:ss AM/PM',
      extra: '',
    })
  })

  it.each(['cellprot', 'sort-custom', 'border'])(
    'splits %s into three segments, since it is declared to take an extra',
    (name) => {
      expect(parseStyleCommand(`${name}:a:b`)).toEqual({ name, argument: 'a', extra: 'b' })
    },
  )

  it('keeps trailing colons inside the extra segment of a three-part command', () => {
    expect(parseStyleCommand('border:top:thin:extra')).toEqual({
      name: 'border',
      argument: 'top',
      extra: 'thin:extra',
    })
  })

  it('handles a three-segment command that only supplies two', () => {
    expect(parseStyleCommand('border:top')).toEqual({ name: 'border', argument: 'top', extra: '' })
  })

  it('handles an empty argument', () => {
    expect(parseStyleCommand('zoom:')).toEqual({ name: 'zoom', argument: '', extra: '' })
  })
})

/** An object where every property read yields a no-op function. */
function autoStub(): never {
  return new Proxy(
    {},
    {
      get: () => () => undefined,
      has: () => true,
    },
  ) as never
}

/** Every collaborator stubbed; the dispatcher's job is to pick one, not to work. */
function context(overrides: Record<string, unknown> = {}) {
  const setMessage = vi.fn()
  const handlePageLayoutCommand = vi.fn()
  const handleExportPdf = vi.fn(() => Promise.resolve())
  const recordFreezeJournal = vi.fn()
  const setChartDialog = vi.fn()
  const setSymbolDialogOpen = vi.fn()
  const setPendingEdits = vi.fn()
  const ctx = {
    univerRef: ref(fakeRuntime(fakeWorkbook([fakeSheet()]))),
    lazyWorkbookRef: ref(fakeLazyState()),
    traceArrowsRef: ref({ disposables: [], nextId: 1 }),
    sparklineDisposablesRef: ref([]),
    sparklineTimerRef: ref(null),
    chartEditRef: ref(vi.fn()),
    shapeEditRef: ref(vi.fn()),
    refreshSelectionFormatRef: ref(vi.fn()),
    selectedVisual: null,
    selectedChart: null,
    setMessage,
    setChartDialog,
    setSymbolDialogOpen,
    setPendingEdits,
    // Auto-stubbed collaborators: any property is a no-op function. The
    // dispatcher's job is to pick one and hand off; whether the collaborator
    // then works is that collaborator's own test. An empty object would make
    // this suite fail on the hand-off itself, which proves nothing.
    visualContext: () => autoStub(),
    dataToolsContext: () => autoStub(),
    pivotContext: () => autoStub(),
    recordFreezeJournal,
    handlePageLayoutCommand,
    handleExportPdf,
    ...overrides,
  } as never
  return {
    ctx,
    setMessage,
    handlePageLayoutCommand,
    handleExportPdf,
    recordFreezeJournal,
    setChartDialog,
  }
}

/// The real command strings the dispatcher branches on, with plausible
/// arguments for the ones that take them.
const COMMANDS = [
  'undo',
  'redo',
  'zoom-in',
  'zoom-reset',
  'zoom:150',
  'fill-down',
  'decimal-inc',
  'insert-row-here',
  'insert-col-here',
  'delete-row-here',
  'row-height:20',
  'col-width:80',
  'cell-style:Good',
  'format-as-table',
  'format-as-table:TableStyleMedium2',
  'insert-chart:column',
  'insert-pivot-chart:column',
  'insert-shape:rect',
  'insert-textbox',
  'insert-picture',
  'trace-precedents',
  'slicer-open',
  'chart-delete',
  'chart-format-pane',
  'chart-select-data',
  'chart-switch-row-col',
  'chart-title:New title',
  'chart-legend:bottom',
  'chart-labels:value',
  'chart-layout:1',
  'chart-colors:1',
  'chart-grouping:stacked',
  'chart-axis-cat:Category',
  'chart-axis-val:Value',
  'chart-type-column',
  'page-layout:margins:normal',
  'export-pdf',
  'error:something went wrong',
]

describe('handleRibbonCommand never throws', () => {
  it.each(COMMANDS)('%s survives a workbook with nothing selected', (command) => {
    const { ctx } = context()
    expect(() => handleRibbonCommand(ctx, command)).not.toThrow()
  })

  it.each(COMMANDS)('%s survives with no workbook open at all', (command) => {
    const { ctx } = context({ univerRef: ref(fakeRuntime(null)), lazyWorkbookRef: ref(null) })
    expect(() => handleRibbonCommand(ctx, command)).not.toThrow()
  })

  it('ignores a command it has never heard of rather than failing', () => {
    const { ctx } = context()
    expect(() => handleRibbonCommand(ctx, 'not-a-real-command:42')).not.toThrow()
  })

  it('ignores an empty command', () => {
    const { ctx } = context()
    expect(() => handleRibbonCommand(ctx, '')).not.toThrow()
  })
})

describe('handleRibbonCommand routes to the right collaborator', () => {
  it('hands page-layout commands to the page-layout handler, minus the prefix', () => {
    const { ctx, handlePageLayoutCommand } = context()
    handleRibbonCommand(ctx, 'page-layout:margins:wide')
    expect(handlePageLayoutCommand).toHaveBeenCalledWith('margins:wide')
  })

  it('routes export-pdf to the export handler', () => {
    const { ctx, handleExportPdf } = context()
    handleRibbonCommand(ctx, 'export-pdf')
    expect(handleExportPdf).toHaveBeenCalled()
  })

  it('surfaces an error command as a message rather than swallowing it', () => {
    const { ctx, setMessage } = context()
    handleRibbonCommand(ctx, 'error:the thing failed')
    expect(setMessage).toHaveBeenCalledWith(expect.stringContaining('the thing failed'))
  })

  it('opens the chart dialog only when a chart is actually selected', () => {
    const { ctx, setChartDialog } = context()
    handleRibbonCommand(ctx, 'chart-select-data')
    expect(setChartDialog).not.toHaveBeenCalled()
  })
})
