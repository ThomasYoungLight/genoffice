/**
 * The shared OMML layer.
 *
 * This package exists because Word and PowerPoint carry the same maths markup,
 * and the two engines were about to grow separate copies of the parser. Its
 * contract is therefore narrower than either engine's: no `w:` and no `a:`,
 * just OMML in and out. These tests pin that — a `w:` creeping back in means
 * the docx-specific wrapper has leaked down a layer.
 */
import { describe, it, expect } from 'vitest'
import { latexToOmml, ommlToLatex, ommlToMathML } from '../src/index'

describe('latexToOmml', () => {
  it('builds the structures a formula is actually made of', () => {
    expect(latexToOmml('\\frac{a}{b}')).toContain('<m:f>')
    expect(latexToOmml('\\sqrt{x}')).toContain('<m:rad>')
    expect(latexToOmml('x^2')).toContain('<m:sSup>')
    expect(latexToOmml('x_i')).toContain('<m:sSub>')
    expect(latexToOmml('\\sum_{i=1}^{n} i')).toContain('<m:nary>')
  })

  it('emits no format-specific markup, which is the point of this package', () => {
    const omml = latexToOmml('\\frac{\\alpha}{\\beta} + x^2')
    expect(omml).not.toMatch(/<w:/)
    expect(omml).not.toMatch(/<a:/)
  })

  it('maps Greek names to the characters themselves', () => {
    expect(latexToOmml('\\alpha')).toContain('α')
    expect(latexToOmml('\\Omega')).toContain('Ω')
  })

  it('rejects a formula it cannot finish reading', () => {
    expect(() => latexToOmml('\\frac{1}')).toThrow()
    expect(() => latexToOmml('{')).toThrow()
  })

  it('round-trips through the reader for a simple formula', () => {
    const back = ommlToLatex(`<m:oMath>${latexToOmml('x^2')}</m:oMath>`)
    expect(back).toContain('x')
  })
})

describe('ommlToMathML', () => {
  it('produces MathML a browser can render', () => {
    const mathml = ommlToMathML(`<m:oMath>${latexToOmml('\\frac{a}{b}')}</m:oMath>`)
    expect(mathml).toContain('<mfrac>')
  })
})
