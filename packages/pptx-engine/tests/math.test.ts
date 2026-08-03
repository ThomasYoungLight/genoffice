/**
 * Equations on a slide.
 *
 * Two things have to hold and neither is obvious. PowerPoint only typesets the
 * maths if the `mc:Choice` block is shaped exactly right, and every other
 * reader — including our own Konva canvas, which draws runs and knows nothing
 * about OMML — falls through to `mc:Fallback`, so the fallback has to carry
 * something worth showing. And the paragraph has to survive our own parser:
 * a slide that stops parsing after an insert would leave the session rendering
 * one document and saving another.
 */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  addEquation,
  equationParagraphXml,
  hasEquation,
  openPptx,
  readRawPart,
  savePptx,
} from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))
const open = () => openPptx(fx('01_standard_business.pptx'))

describe('equationParagraphXml', () => {
  it('wraps OMML so maths-aware readers choose it and others fall back', () => {
    const { xml } = equationParagraphXml('E = mc^2')
    expect(xml).toContain('<mc:AlternateContent')
    expect(xml).toContain('Requires="a14"')
    expect(xml).toContain('<a14:m>')
    expect(xml).toContain('<m:oMath>')
    expect(xml).toContain('<mc:Fallback>')
  })

  it('declares every namespace it uses, so no slide-root edit is needed', () => {
    const { xml } = equationParagraphXml('x^2')
    for (const prefix of ['mc', 'a14', 'm']) {
      expect(xml).toContain(`xmlns:${prefix}=`)
    }
  })

  it('puts the source in the fallback, so a reader without maths shows something', () => {
    const { xml, fallbackText } = equationParagraphXml('\\frac{a}{b}')
    expect(fallbackText).toBe('\\frac{a}{b}')
    expect(xml).toContain('<a:t>\\frac{a}{b}</a:t>')
  })

  it('escapes a formula containing XML metacharacters', () => {
    const { xml } = equationParagraphXml('a < b')
    expect(xml).toContain('a &lt; b')
    expect(xml).not.toContain('<a:t>a < b</a:t>')
  })

  it('builds real OMML for fractions, roots and superscripts', () => {
    expect(equationParagraphXml('\\frac{1}{2}').xml).toContain('<m:f>')
    expect(equationParagraphXml('\\sqrt{2}').xml).toContain('<m:rad>')
    expect(equationParagraphXml('x^2').xml).toContain('<m:sSup>')
  })

  it('honours alignment on both the paragraph and the maths', () => {
    const { xml } = equationParagraphXml('x', { align: 'left' })
    expect(xml).toContain('algn="l"')
    expect(xml).toContain('m:val="left"')
  })

  it('refuses an empty formula', () => {
    expect(() => equationParagraphXml('   ')).toThrow()
  })

  it('reports unparseable LaTeX by naming where it stopped', () => {
    expect(() => equationParagraphXml('\\frac{1}')).toThrow()
  })
})

describe('an equation on a real slide', () => {
  it('lands in the saved slide as maths, not as a run of text', async () => {
    const opened = await open()
    const { element, fallbackText } = addEquation(opened.deck.slides[0]!, {
      latex: 'E = mc^2',
      offset: { x: 914400, y: 914400, cx: 3657600, cy: 914400 },
    })
    expect(element.type).toBe('text')
    expect(fallbackText).toBe('E = mc^2')

    const reopened = await openPptx(await savePptx(opened))
    const saved = readRawPart(reopened, '/slide[1]')
    if (!saved.ok) throw new Error(saved.error)
    expect(hasEquation(saved.xml)).toBe(true)
    expect(saved.xml).toContain('<m:oMath>')
    expect(saved.xml).toContain('<mc:Fallback>')
  })

  it('leaves the slide parseable, so the session and the file agree', async () => {
    const opened = await open()
    const before = opened.deck.slides[0]!.elements.length
    addEquation(opened.deck.slides[0]!, {
      latex: '\\frac{a}{b}',
      offset: { x: 0, y: 0, cx: 1828800, cy: 457200 },
    })
    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.deck.slides.length).toBe(opened.deck.slides.length)
    expect(reopened.deck.slides[0]!.elements.length).toBe(before + 1)
  })

  it('shows the source on our own canvas, which cannot draw maths', async () => {
    const opened = await open()
    const { element } = addEquation(opened.deck.slides[0]!, {
      latex: '\\sqrt{x}',
      offset: { x: 0, y: 0, cx: 1828800, cy: 457200 },
    })
    expect(element.text?.paragraphs[0]?.runs[0]?.text).toBe('\\sqrt{x}')
  })

  it('does not touch the deck when the formula will not parse', async () => {
    const opened = await open()
    const before = opened.deck.slides[0]!.elements.length
    expect(() =>
      addEquation(opened.deck.slides[0]!, {
        latex: '\\frac{1}',
        offset: { x: 0, y: 0, cx: 100, cy: 100 },
      }),
    ).toThrow()
    expect(opened.deck.slides[0]!.elements.length).toBe(before)
  })
})
