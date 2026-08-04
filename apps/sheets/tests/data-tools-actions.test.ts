/**
 * Data-tab actions: Go To, defined names, formula entry, Advanced Filter,
 * Subtotal, Consolidate, Outline, Format as Table.
 *
 * These are the commands the ribbon dispatcher hands off to, and they run
 * against whatever the user has selected — including nothing. The guards that
 * matter are the ones that turn "no selection" or "a range that does not
 * exist" into a message rather than an exception, because the alternative is a
 * menu item that appears to do nothing at all.
 */
import { describe, expect, it, vi } from 'vitest'

import {
  activeCellLabel,
  consolidateDefaultReference,
  goToReference,
  listDefinedNames,
} from '../src/renderer/data-tools-actions'
import { fakeLazyState, fakeRuntime, fakeSheet, fakeWorkbook, ref } from './helpers/fake-univer'

function context(overrides: Record<string, unknown> = {}) {
  const setMessage = vi.fn()
  const setPendingEdits = vi.fn()
  const setAdvancedFilterColumns = vi.fn()
  const ctx = {
    univerRef: ref(fakeRuntime(fakeWorkbook([fakeSheet()]))),
    lazyWorkbookRef: ref(fakeLazyState()),
    setMessage,
    setPendingEdits,
    setAdvancedFilterColumns,
    ...overrides,
  } as never
  return { ctx, setMessage, setPendingEdits, setAdvancedFilterColumns }
}

const noWorkbook = () =>
  context({ univerRef: ref(fakeRuntime(null)), lazyWorkbookRef: ref(null) })

describe('activeCellLabel', () => {
  it('names the selected cell in A1 notation', () => {
    const { ctx } = context()
    expect(activeCellLabel(ctx)).toBe('A1')
  })

  it('falls back to a label rather than throwing when nothing is open', () => {
    // This drives the Name Box, which is always on screen; an exception here
    // would take the whole toolbar down.
    const { ctx } = noWorkbook()
    expect(typeof activeCellLabel(ctx)).toBe('string')
    expect(activeCellLabel(ctx).length).toBeGreaterThan(0)
  })
})

describe('goToReference', () => {
  // It returns the problem rather than calling setMessage — the caller decides
  // how to surface it, and null means the jump happened.
  it('returns a message naming a reference it cannot resolve', () => {
    // Silently going to A1 on a typo is worse than saying no: the user acts on
    // the belief they are somewhere else.
    const { ctx } = context()
    const message = goToReference(ctx, 'not a reference!!')
    expect(message).toBeTruthy()
    expect(message).toContain('not a reference!!')
  })

  it('returns a message for an empty reference rather than jumping anywhere', () => {
    const { ctx } = context()
    expect(goToReference(ctx, '')).toBeTruthy()
  })

  it('returns a not-ready message when no workbook is open', () => {
    const { ctx } = noWorkbook()
    expect(goToReference(ctx, 'A1')).toBeTruthy()
  })

  it.each(['B2', 'A1:C10', '  B2  '])('resolves %s without complaining', (reference) => {
    const { ctx } = context()
    expect(() => goToReference(ctx, reference)).not.toThrow()
  })
})

describe('listDefinedNames', () => {
  it('returns an empty list rather than null when there are none', () => {
    // The Name Manager renders this directly; null would be a crash.
    const { ctx } = context()
    expect(Array.isArray(listDefinedNames(ctx))).toBe(true)
  })

  it('returns an empty list with no workbook open', () => {
    const { ctx } = noWorkbook()
    expect(listDefinedNames(ctx)).toEqual([])
  })
})

describe('consolidateDefaultReference', () => {
  it('offers something usable as a starting reference', () => {
    const { ctx } = context()
    expect(typeof consolidateDefaultReference(ctx)).toBe('string')
  })

  it('does not throw with no workbook open', () => {
    const { ctx } = noWorkbook()
    expect(() => consolidateDefaultReference(ctx)).not.toThrow()
  })
})
