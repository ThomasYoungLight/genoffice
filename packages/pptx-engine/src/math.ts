/**
 * Equations on a slide.
 *
 * A slide equation is OMML — the same markup Word uses — carried inside a
 * DrawingML paragraph rather than a Word one. PowerPoint wraps it in
 * `mc:AlternateContent`: a reader that understands the 2010 drawing extensions
 * takes `mc:Choice` and typesets the maths, anything else takes `mc:Fallback`.
 * The fallback is therefore not decoration — without it an older reader shows
 * an empty paragraph and the formula is simply gone. It holds the LaTeX source,
 * which is the most useful thing to show and is also what our own canvas
 * renders, since the Konva renderer draws runs and not maths.
 *
 * The LaTeX parser is not reimplemented here: `@genoffice/omml` has the one the
 * docs `<formula>` channel already uses, so the same input produces the same
 * maths in both apps and a formula moved from a document to a deck stays the
 * same formula.
 *
 * Namespaces are declared on the elements that use them rather than on the
 * slide root, which keeps this to a paragraph insert — the same approach the
 * morph transition in generate.ts takes.
 */
import { latexToOmml } from '@genoffice/omml'
import { escapeXmlText } from './xml-utils'
import { addElement } from './insert'
import type { EmuRect, Slide, TextElement } from './types'

const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const A14_NS = 'http://schemas.microsoft.com/office/drawing/2010/main'
const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

export interface EquationParagraph {
  /** an `<a:p>` for a shape's txBody */
  xml: string
  /** what a reader without maths support shows, and what our canvas draws */
  fallbackText: string
}

/**
 * Build the paragraph for one display equation.
 *
 * Throws when the LaTeX cannot be parsed; the message comes from the shared
 * parser and names the fragment it stopped on, which is what a model needs in
 * order to fix the input rather than retry the same string.
 */
export function equationParagraphXml(
  latex: string,
  opts: { align?: 'left' | 'center' | 'right'; fontSizePt?: number } = {},
): EquationParagraph {
  const source = latex.trim()
  if (!source) throw new Error('The equation is empty')
  const omml = latexToOmml(source)
  const align = opts.align ?? 'center'
  const algn = align === 'left' ? 'l' : align === 'right' ? 'r' : 'ctr'
  const size = opts.fontSizePt ? ` sz="${Math.round(opts.fontSizePt * 100)}"` : ''
  const fallback = `<a:r><a:rPr lang="en-US"${size} dirty="0"/><a:t>${escapeXmlText(source)}</a:t></a:r>`
  const xml =
    `<a:p><a:pPr algn="${algn}"/>` +
    `<mc:AlternateContent xmlns:mc="${MC_NS}">` +
    `<mc:Choice xmlns:a14="${A14_NS}" Requires="a14">` +
    `<a14:m><m:oMathPara xmlns:m="${MATH_NS}">` +
    `<m:oMathParaPr><m:jc m:val="${align === 'center' ? 'center' : align}"/></m:oMathParaPr>` +
    `<m:oMath>${omml}</m:oMath></m:oMathPara></a14:m></mc:Choice>` +
    `<mc:Fallback>${fallback}</mc:Fallback>` +
    `</mc:AlternateContent></a:p>`
  return { xml, fallbackText: source }
}

/** True when a text body carries at least one equation. */
export function hasEquation(xml: string): boolean {
  return /<a14:m\b/.test(xml)
}

/**
 * Add a text box holding one equation and return the new element.
 *
 * The element's text model holds the LaTeX source rather than the maths: the
 * canvas renders runs, and this way what we draw matches what a reader without
 * maths support sees. That also bounds the damage if the element is later
 * rewritten with a text edit — the equation degrades to its source instead of
 * to nothing, because a text edit regenerates the body from the run model and
 * cannot preserve the AlternateContent block.
 */
export function addEquation(
  slide: Slide,
  opts: {
    latex: string
    offset: EmuRect
    align?: 'left' | 'center' | 'right'
    fontSizePt?: number
  },
): { element: TextElement; fallbackText: string } {
  const { xml, fallbackText } = equationParagraphXml(opts.latex, {
    ...(opts.align ? { align: opts.align } : {}),
    ...(opts.fontSizePt ? { fontSizePt: opts.fontSizePt } : {}),
  })
  const element = addElement(slide, {
    kind: 'textbox',
    offset: opts.offset,
    rawParagraphsXml: xml,
    paragraphs: [
      {
        runs: [{ text: fallbackText, ...(opts.fontSizePt ? { fontSize: opts.fontSizePt } : {}) }],
        ...(opts.align ? { align: opts.align } : {}),
      },
    ],
  })
  return { element, fallbackText }
}
