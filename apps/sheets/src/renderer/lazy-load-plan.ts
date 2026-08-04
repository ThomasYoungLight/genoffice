/**
 * The decisions inside lazy range loading, separated from the effects.
 *
 * `loadRange` interleaved three judgements with the I/O that acts on them —
 * whether a request is worth making at all, how much of it the sidecar can
 * satisfy yet, and what to do when it cannot satisfy all of it. Those
 * judgements carry the off-by-ones and the retry cap; the I/O around them is
 * straightforward. Untangled, the judgements are ordinary functions of their
 * inputs and can be tested without a Univer instance or a sidecar.
 *
 * Nothing here touches the runtime, the journal or the clock.
 */

/** A rectangle of the grid; matches Univer's IRange for the fields used here. */
export interface GridRange {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

export function containsRange(container: GridRange, requested: GridRange): boolean {
  return (
    container.startRow <= requested.startRow &&
    container.endRow >= requested.endRow &&
    container.startColumn <= requested.startColumn &&
    container.endColumn >= requested.endColumn
  )
}

/** Identifies one in-flight request for a sheet, so a duplicate can be dropped. */
export function requestKeyFor(range: GridRange): string {
  return `${range.startRow}:${range.endRow}:${range.startColumn}:${range.endColumn}`
}

export type LoadDecision =
  /// Already on screen, or the same request is in flight — do nothing.
  | { readonly kind: 'skip'; readonly reason: 'already-loaded' | 'in-flight' }
  | { readonly kind: 'load'; readonly requestKey: string }

/**
 * Whether a range request is worth making.
 *
 * A retry always proceeds: it exists precisely because the previous attempt
 * could not satisfy the range, so the "already loaded" and "in flight" checks
 * would send it in a circle.
 */
export function planRangeLoad(input: {
  readonly range: GridRange
  readonly loaded: GridRange | undefined
  readonly inFlightKey: string | undefined
  readonly isRetry: boolean
}): LoadDecision {
  const requestKey = requestKeyFor(input.range)
  if (input.isRetry) return { kind: 'load', requestKey }
  if (input.loaded && containsRange(input.loaded, input.range)) {
    return { kind: 'skip', reason: 'already-loaded' }
  }
  if (input.inFlightKey === requestKey) return { kind: 'skip', reason: 'in-flight' }
  return { kind: 'load', requestKey }
}

/**
 * How much of the requested range the sidecar has indexed so far.
 *
 * `indexedThroughScreen === null` means nothing is indexed yet, which is not
 * the same as "row 0 is indexed" — returning a range in that case would paint
 * a row of blanks over cells that simply have not arrived.
 */
export function availableSubRange(
  range: GridRange,
  indexedThroughScreen: number | null,
): GridRange | null {
  if (indexedThroughScreen === null) return null
  const endRow = Math.min(indexedThroughScreen, range.endRow)
  if (endRow < range.startRow) return null
  return { ...range, endRow }
}

export type LoadFollowUp =
  /// The sheet finished indexing; nothing more to wait for.
  | { readonly kind: 'done' }
  /// The caller asked to wait for the full range and it is still incomplete:
  /// sleep briefly and ask again on the same turn.
  | { readonly kind: 'poll'; readonly delayMs: number; readonly nextAttempt: number }
  /// Still indexing, but nobody is blocked on it: come back on a timer.
  | { readonly kind: 'retry-later'; readonly delayMs: number }

/** How long to wait before looking again, in either mode. */
export const LOAD_RETRY_MS = 250
/** Polling gives up after this many turns and falls back to the timer. */
export const MAX_POLL_ATTEMPTS = 20

/**
 * What to do once a partial result is in.
 *
 * The distinction that matters: polling blocks the caller and so must be
 * bounded, while the timer path does not and so can continue indefinitely
 * until the stream completes.
 */
export function planFollowUp(input: {
  readonly indexingComplete: boolean
  readonly requestedEndRow: number
  readonly availableEndRow: number | null
  readonly waitForRequestedRange: boolean
  readonly waitAttempt: number
}): LoadFollowUp {
  if (input.indexingComplete) return { kind: 'done' }
  const satisfied = input.availableEndRow !== null && input.availableEndRow >= input.requestedEndRow
  if (input.waitForRequestedRange && input.waitAttempt < MAX_POLL_ATTEMPTS && !satisfied) {
    return { kind: 'poll', delayMs: LOAD_RETRY_MS, nextAttempt: input.waitAttempt + 1 }
  }
  return { kind: 'retry-later', delayMs: LOAD_RETRY_MS }
}

/**
 * Whether to tell the user indexing is still running, and how many rows are
 * done. `indexedThroughRow` is the last row indexed, so the count is one more
 * — reporting the index itself is off by one on every sheet.
 */
export function indexingProgress(input: {
  readonly isActiveSheet: boolean
  readonly indexedThroughRow: number | null
  readonly fileEndRow: number
}): { readonly show: boolean; readonly rowsIndexed: number } {
  const rowsIndexed = (input.indexedThroughRow ?? -1) + 1
  const show =
    input.isActiveSheet &&
    (input.indexedThroughRow === null || input.indexedThroughRow < input.fileEndRow)
  return { show, rowsIndexed }
}
