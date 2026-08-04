/**
 * The w:date format.
 *
 * Word writes local wall-clock digits with a Z suffix and renders the digits as
 * local time whatever the designator says, keeping the real instant in a
 * separate w16du:dateUtc. Verified by letting Word author a comment and a
 * tracked insertion at 13:49 local in UTC+8 and reading back
 * `w:date="2026-08-04T13:49:00Z" w16du:dateUtc="2026-08-04T05:49:00Z"`.
 *
 * Writing a true UTC timestamp instead — which is what an ISO string is —
 * displays every comment and tracked change shifted by the author's UTC offset.
 */
import { describe, it, expect } from 'vitest'
import { wordTimestamp } from '../src/index'

describe('wordTimestamp', () => {
  it('writes the local wall clock, not UTC, and still ends in Z', () => {
    // a local time whose UTC form lands on a different day, so a UTC timestamp
    // cannot coincidentally match
    const t = new Date(2026, 7, 4, 1, 5, 9)
    expect(wordTimestamp(t)).toBe('2026-08-04T01:05:09Z')
  })

  it('agrees with Word on the shape of the string', () => {
    expect(wordTimestamp(new Date())).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/)
  })

  it('reads back as the same wall clock a reader would show', () => {
    const t = new Date(2026, 11, 31, 23, 59, 58)
    const shown = wordTimestamp(t).replace(/Z$/, '')
    expect(new Date(shown).getHours()).toBe(23)
    expect(new Date(shown).getDate()).toBe(31)
  })
})
