/**
 * XML helpers for docx.
 *
 * The generic half — the parser and the node accessors — lives in
 * `@genoffice/omml`, because the math module there needs exactly the same
 * helpers and two copies would drift. They are re-exported so every
 * `from './xml-utils'` import in this package keeps working. What stays below
 * is the part that only means something in a Word document.
 */
export {
  attrsOf,
  childrenOf,
  escapeXmlAttr,
  escapeXmlText,
  findChild,
  findChildren,
  nameOf,
  serializeXNode,
  textOf,
  xmlParser,
  type XNode,
} from '@genoffice/omml'

import { attrsOf, childrenOf, findChild, nameOf, type XNode } from '@genoffice/omml'

/**
 * Direct children with `name`, looking through w:sdt → w:sdtContent wrappers
 * (nested sdt included). Structured document tags may wrap table rows, cells
 * or paragraphs at any level; for display purposes the wrapper is transparent
 * (research-report templates wrap every field in an sdt).
 */
export function childrenThroughSdt(node: XNode, name: string): XNode[] {
  const out: XNode[] = []
  const visit = (n: XNode): void => {
    for (const child of childrenOf(n)) {
      const cn = nameOf(child)
      if (cn === name) out.push(child)
      else if (cn === 'w:sdt') {
        const content = findChild(child, 'w:sdtContent')
        if (content) visit(content)
      }
    }
  }
  visit(node)
  return out
}

/** OOXML boolean property: present => true unless w:val says otherwise */
export function boolProp(parent: XNode, name: string): boolean {
  const child = findChild(parent, name)
  if (!child) return false
  const val = attrsOf(child)['w:val']
  if (val === undefined) return true
  return !['0', 'false', 'none', 'off'].includes(val.toLowerCase())
}

/**
 * w:u is NOT an OOXML boolean (CT_OnOff) — it is CT_Underline, where the
 * underline pattern lives entirely in w:val. A <w:u> with no w:val (e.g.
 * `<w:u w:color="415461"/>` as emitted by Pages/LibreOffice) means no
 * underline, matching how Word renders it.
 */
export function underlineProp(parent: XNode): boolean {
  const child = findChild(parent, 'w:u')
  if (!child) return false
  const val = attrsOf(child)['w:val']
  return val !== undefined && val !== 'none'
}
