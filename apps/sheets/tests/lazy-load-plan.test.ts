/**
 * The decisions extracted out of `loadRange`.
 *
 * These are the parts of lazy loading that are easy to get subtly wrong and
 * impossible to notice: an off-by-one in the indexed-row count, a retry that
 * loops forever because it re-checks "already loaded", a sub-range computed
 * from a null that paints blanks over cells still in flight. None of it was
 * reachable by a test while it lived inside a function that also drove a
 * sidecar, a worksheet and two timers.
 */
import { describe, expect, it } from 'vitest'

import {
  availableSubRange,
  containsRange,
  indexingProgress,
  LOAD_RETRY_MS,
  MAX_POLL_ATTEMPTS,
  planFollowUp,
  planRangeLoad,
  requestKeyFor,
} from '../src/renderer/lazy-load-plan'

const range = (startRow: number, endRow: number, startColumn = 0, endColumn = 5) => ({
  startRow,
  endRow,
  startColumn,
  endColumn,
})

describe('containsRange', () => {
  it('is true for an identical range', () => {
    expect(containsRange(range(0, 10), range(0, 10))).toBe(true)
  })

  it('is true for a strict subset', () => {
    expect(containsRange(range(0, 10, 0, 9), range(2, 8, 1, 4))).toBe(true)
  })

  it('is false when the request extends past the bottom', () => {
    expect(containsRange(range(0, 10), range(0, 11))).toBe(false)
  })

  it('is false when the request extends past the right edge', () => {
    expect(containsRange(range(0, 10, 0, 5), range(0, 10, 0, 6))).toBe(false)
  })

  it('is false for a disjoint range', () => {
    expect(containsRange(range(0, 10), range(20, 30))).toBe(false)
  })
})

describe('planRangeLoad', () => {
  it('loads when nothing is on screen yet', () => {
    expect(planRangeLoad({ range: range(0, 50), loaded: undefined, inFlightKey: undefined, isRetry: false })).toEqual(
      { kind: 'load', requestKey: requestKeyFor(range(0, 50)) },
    )
  })

  it('skips a range already covered by what is loaded', () => {
    const decision = planRangeLoad({
      range: range(10, 20),
      loaded: range(0, 100),
      inFlightKey: undefined,
      isRetry: false,
    })
    expect(decision).toEqual({ kind: 'skip', reason: 'already-loaded' })
  })

  it('loads a range that only partly overlaps what is loaded', () => {
    const decision = planRangeLoad({
      range: range(90, 120),
      loaded: range(0, 100),
      inFlightKey: undefined,
      isRetry: false,
    })
    expect(decision.kind).toBe('load')
  })

  it('skips a request identical to one already in flight', () => {
    const wanted = range(0, 50)
    const decision = planRangeLoad({
      range: wanted,
      loaded: undefined,
      inFlightKey: requestKeyFor(wanted),
      isRetry: false,
    })
    expect(decision).toEqual({ kind: 'skip', reason: 'in-flight' })
  })

  it('loads when a different request is in flight', () => {
    const decision = planRangeLoad({
      range: range(0, 50),
      loaded: undefined,
      inFlightKey: requestKeyFor(range(60, 90)),
      isRetry: false,
    })
    expect(decision.kind).toBe('load')
  })

  it('always proceeds on a retry, whatever the state says', () => {
    // A retry exists because the last attempt could not satisfy the range.
    // Honouring "already loaded" or "in flight" here would spin it forever.
    for (const state of [
      { loaded: range(0, 100), inFlightKey: undefined },
      { loaded: undefined, inFlightKey: requestKeyFor(range(10, 20)) },
      { loaded: range(0, 100), inFlightKey: requestKeyFor(range(10, 20)) },
    ]) {
      expect(planRangeLoad({ range: range(10, 20), ...state, isRetry: true }).kind).toBe('load')
    }
  })

  it('gives different ranges different request keys', () => {
    expect(requestKeyFor(range(0, 10))).not.toBe(requestKeyFor(range(0, 11)))
    expect(requestKeyFor(range(0, 10, 0, 5))).not.toBe(requestKeyFor(range(0, 10, 0, 6)))
  })
})

describe('availableSubRange', () => {
  it('returns nothing when the sheet has not been indexed at all', () => {
    // null is "no rows yet", not "row 0". Treating it as 0 would paint a row
    // of blanks over cells that simply have not arrived.
    expect(availableSubRange(range(0, 50), null)).toBeNull()
  })

  it('clips the request to what has been indexed', () => {
    expect(availableSubRange(range(0, 50), 20)).toEqual(range(0, 20))
  })

  it('returns the whole request when indexing is ahead of it', () => {
    expect(availableSubRange(range(0, 50), 500)).toEqual(range(0, 50))
  })

  it('returns nothing when indexing has not reached the top of the request', () => {
    // Scrolled to row 900 with 100 rows indexed: there is nothing to show yet.
    expect(availableSubRange(range(900, 950), 100)).toBeNull()
  })

  it('returns a single row when indexing has reached exactly the first one', () => {
    expect(availableSubRange(range(10, 50), 10)).toEqual(range(10, 10))
  })

  it('keeps the requested columns untouched', () => {
    expect(availableSubRange(range(0, 50, 3, 8), 20)).toMatchObject({
      startColumn: 3,
      endColumn: 8,
    })
  })
})

describe('planFollowUp', () => {
  const base = {
    indexingComplete: false,
    requestedEndRow: 50,
    availableEndRow: 20,
    waitForRequestedRange: false,
    waitAttempt: 0,
  }

  it('is done once indexing completes, even mid-poll', () => {
    expect(planFollowUp({ ...base, indexingComplete: true, waitForRequestedRange: true })).toEqual({
      kind: 'done',
    })
  })

  it('schedules a timer when nobody is blocked on the range', () => {
    expect(planFollowUp(base)).toEqual({ kind: 'retry-later', delayMs: LOAD_RETRY_MS })
  })

  it('polls when the caller is waiting and the range is short', () => {
    expect(planFollowUp({ ...base, waitForRequestedRange: true })).toEqual({
      kind: 'poll',
      delayMs: LOAD_RETRY_MS,
      nextAttempt: 1,
    })
  })

  it('stops polling once the range is satisfied, even mid-index', () => {
    const followUp = planFollowUp({
      ...base,
      waitForRequestedRange: true,
      availableEndRow: 50,
    })
    expect(followUp.kind).toBe('retry-later')
  })

  it('gives up polling at the attempt cap and falls back to the timer', () => {
    // Polling blocks the caller, so it has to be bounded; the timer path does
    // not, so it can keep going until the stream finishes.
    expect(
      planFollowUp({ ...base, waitForRequestedRange: true, waitAttempt: MAX_POLL_ATTEMPTS }).kind,
    ).toBe('retry-later')
  })

  it('still polls on the last attempt below the cap', () => {
    expect(
      planFollowUp({ ...base, waitForRequestedRange: true, waitAttempt: MAX_POLL_ATTEMPTS - 1 })
        .kind,
    ).toBe('poll')
  })

  it('counts attempts upward so the cap is reachable', () => {
    const first = planFollowUp({ ...base, waitForRequestedRange: true, waitAttempt: 3 })
    expect(first).toMatchObject({ nextAttempt: 4 })
  })

  it('polls when nothing at all is available yet', () => {
    expect(
      planFollowUp({ ...base, waitForRequestedRange: true, availableEndRow: null }).kind,
    ).toBe('poll')
  })
})

describe('indexingProgress', () => {
  it('counts rows, not the index of the last one', () => {
    // indexedThroughRow is a 0-based row index; reporting it directly is off
    // by one on every sheet in the product.
    expect(indexingProgress({ isActiveSheet: true, indexedThroughRow: 99, fileEndRow: 500 })).toEqual(
      { show: true, rowsIndexed: 100 },
    )
  })

  it('reports zero rows when nothing is indexed yet', () => {
    expect(
      indexingProgress({ isActiveSheet: true, indexedThroughRow: null, fileEndRow: 500 }),
    ).toEqual({ show: true, rowsIndexed: 0 })
  })

  it('stays quiet on a sheet the user is not looking at', () => {
    expect(
      indexingProgress({ isActiveSheet: false, indexedThroughRow: 10, fileEndRow: 500 }).show,
    ).toBe(false)
  })

  it('stops reporting once indexing has passed the end of the file', () => {
    expect(
      indexingProgress({ isActiveSheet: true, indexedThroughRow: 500, fileEndRow: 500 }).show,
    ).toBe(false)
  })
})
