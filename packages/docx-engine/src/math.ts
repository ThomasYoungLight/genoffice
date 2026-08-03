/**
 * Equations in a Word document.
 *
 * OMML itself is not a Word format — PowerPoint carries the same markup — so
 * the LaTeX parser, the MathML converter and the OMML readers live in
 * `@genoffice/omml` and are re-exported here. What is genuinely docx is the
 * paragraph the maths sits in, which is the one function below.
 */
export {
  latexToOmml,
  mathTokensOf,
  ommlFragmentsOf,
  ommlToLatex,
  ommlToMathML,
} from '@genoffice/omml'

/** OMML for a display formula paragraph created by the editor */
export function mathParagraphXml(
  omml: string,
  align: 'left' | 'center' | 'right' = 'center',
): string {
  const jc = align === 'center' ? '' : `<w:pPr><w:jc w:val="${align}"/></w:pPr>`
  return (
    `<w:p>${jc}<m:oMathPara><m:oMathParaPr><m:jc m:val="${align === 'center' ? 'center' : align}"/></m:oMathParaPr>` +
    `<m:oMath>${omml}</m:oMath></m:oMathPara></w:p>`
  )
}
