/**
 * The deck footer, and the partial update that used to erase it.
 *
 * applyHeaderFooter clears the whole dt/ftr/sldNum family before writing the
 * enabled parts back, so an omitted field reads as a deletion. The dialog always
 * sends all three and never noticed; the agent sends only the field it was asked
 * to change, so "add slide numbers" silently dropped the footer. The skill-level
 * test asserted the tool omits untouched fields and stopped at the IPC mock —
 * the other side of that boundary was where the omission became a delete.
 */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import {
  applyHeaderFooter,
  mergeHeaderFooter,
  openPptx,
  readHeaderFooter,
  savePptx,
} from '../src/index'

const fx = (name: string) => readFileSync(join(__dirname, 'fixtures', name))
const ALL = { footer: 'Q3 review', slideNum: true, date: '4 August 2026' }

describe('applyHeaderFooter', () => {
  it('writes all three placeholders and reads them back after a save', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    expect(applyHeaderFooter(opened, ALL)).toBe(true)
    const reopened = await openPptx(await savePptx(opened))
    expect(readHeaderFooter(reopened.deck.slides[0]!)).toEqual(ALL)
    // every slide, not just the first
    expect(readHeaderFooter(reopened.deck.slides[2]!).footer).toBe('Q3 review')
  })

  it('clears the fields it is not given — the behaviour mergeHeaderFooter exists to absorb', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    applyHeaderFooter(opened, ALL)
    applyHeaderFooter(opened, { date: '5 August 2026' })
    expect(readHeaderFooter(opened.deck.slides[0]!)).toEqual({
      footer: null,
      slideNum: false,
      date: '5 August 2026',
    })
  })

  it('keeps the untouched fields when the caller merges first', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    applyHeaderFooter(opened, ALL)
    const current = readHeaderFooter(opened.deck.slides[0]!)
    applyHeaderFooter(opened, mergeHeaderFooter({ date: '5 August 2026' }, current))
    expect(readHeaderFooter(opened.deck.slides[0]!)).toEqual({
      footer: 'Q3 review',
      slideNum: true,
      date: '5 August 2026',
    })
  })
})

describe('mergeHeaderFooter', () => {
  const current = { footer: 'Keep me', slideNum: true, date: '4 August 2026' }

  it('inherits every field the caller left out', () => {
    expect(mergeHeaderFooter({}, current)).toEqual(current)
  })

  it('still lets null mean remove, which is how the tool clears a footer', () => {
    expect(mergeHeaderFooter({ footer: null }, current).footer).toBeNull()
    expect(mergeHeaderFooter({ slideNum: false }, current).slideNum).toBe(false)
  })

  it('takes the caller value when one is given', () => {
    expect(mergeHeaderFooter({ footer: 'New' }, current).footer).toBe('New')
  })
})
