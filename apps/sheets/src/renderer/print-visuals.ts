/**
 * Captures the floating visuals the grid is showing — charts, images, shapes,
 * sparklines — as static markup the print/PDF layout can place over its table.
 *
 * It reads the live DOM rather than re-rendering the React components, because
 * the on-screen ones have already resolved the two things a static render
 * could not: image bytes fetched asynchronously through `readWorkbookMedia`,
 * and chart series hydrated from the grid for files written without a
 * numCache. What you see is what prints.
 *
 * Each visual marks itself with `data-print-anchor` (see WorkbookVisuals);
 * that attribute carries the xlsx twoCellAnchor, which is the only thing the
 * layout needs to position the fragment against its own re-laid-out table.
 */
import type { PrintVisual, PrintVisualLayer } from './print-html'

/// Editing affordances: real to the grid, meaningless on paper.
const CHROME = [
  '.shape-delete-button',
  '.shape-handle',
  '.shape-resize-outline',
  '.shape-text-editor',
  '.chart-edit-button',
  '.chart-editor',
  '.chart-menu-backdrop',
].join(',')

/// Selection and drag state are transient, not part of the document.
const TRANSIENT_CLASSES = ['selected', 'chart-el-selected', 'shape-drag-ghost']

/// A visual still fetching its bytes or reporting a failure has nothing to
/// print; better a blank space than the word "Loading…" in a PDF.
const NOT_READY = '.xlsx-visual-loading, .xlsx-visual-error'

/// State-dependent rules describe how the grid responds to a pointer, which a
/// printed page never has. Matching them structurally would bake a hover into
/// the output.
const STATEFUL = /:(hover|focus|active|focus-visible|focus-within|target)\b|::/

const MAX_VISUALS = 200

export function capturePrintVisuals(root: ParentNode = document): PrintVisualLayer {
  const visuals: PrintVisual[] = []
  const captured: Element[] = []
  for (const host of Array.from(root.querySelectorAll('[data-print-anchor]'))) {
    if (visuals.length >= MAX_VISUALS) break
    const anchor = parseAnchor(host.getAttribute('data-print-anchor'))
    if (!anchor) continue
    const source = host.firstElementChild
    if (!source || source.querySelector(NOT_READY) || source.matches(NOT_READY)) continue
    const clone = strip(source.cloneNode(true) as Element)
    if (!clone) continue
    captured.push(source)
    visuals.push({ ...anchor, html: clone.outerHTML })
  }
  return { visuals, css: visuals.length === 0 ? '' : collectCss(captured) }
}

/// The eight numbers of a twoCellAnchor, in the order WorkbookVisuals writes
/// them: from(row, column, rowOffset, columnOffset) then to(...).
function parseAnchor(value: string | null): Omit<PrintVisual, 'html'> | null {
  if (value === null) return null
  const parts = value.split(',').map(Number)
  if (parts.length !== 8 || parts.some((part) => !Number.isFinite(part))) return null
  const [fromRow = 0, fromColumn = 0, fromRowOffset = 0, fromColumnOffset = 0] = parts
  const [, , , , toRow = 0, toColumn = 0, toRowOffset = 0, toColumnOffset = 0] = parts
  return {
    fromRow,
    fromColumn,
    fromRowOffset,
    fromColumnOffset,
    toRow,
    toColumn,
    toRowOffset,
    toColumnOffset,
  }
}

/// Removes editing chrome and transient state from a detached clone, then
/// unwraps the editable wrapper so what remains is the visual itself — which
/// is sized 100%/100% and therefore fills the box the layout gives it.
function strip(clone: Element): Element | null {
  for (const node of Array.from(clone.querySelectorAll(CHROME))) node.remove()
  for (const node of [clone, ...Array.from(clone.querySelectorAll('*'))]) {
    node.classList.remove(...TRANSIENT_CLASSES)
    node.removeAttribute('tabindex')
    node.removeAttribute('title')
    node.removeAttribute('contenteditable')
    node.removeAttribute('data-print-anchor')
  }
  if (!clone.classList.contains('shape-editable')) return clone
  const inner = clone.firstElementChild
  return inner ?? null
}

/// Lifts the rules the captured fragments depend on out of the page's own
/// stylesheets, so the exported HTML stays self-contained without a copy of
/// styles.css that could drift from it.
function collectCss(nodes: readonly Element[]): string {
  const subtree: Element[] = []
  for (const node of nodes) subtree.push(node, ...Array.from(node.querySelectorAll('*')))
  const rules: string[] = []
  for (const sheet of Array.from(document.styleSheets)) {
    let list: CSSRuleList
    try {
      list = sheet.cssRules
    } catch {
      // A cross-origin sheet cannot be read; nothing of ours lives in one.
      continue
    }
    for (const rule of Array.from(list)) {
      // Top level only: a rule inside `@media` is conditional on a viewport
      // or an input device, and printToPDF renders under print media where
      // neither is what the screen had.
      if (!(rule instanceof CSSStyleRule)) continue
      const selector = rule.selectorText
      if (selector === ':root' || selector === 'html' || selector.startsWith(':root,')) {
        // Custom properties the captured rules resolve against.
        rules.push(rule.cssText)
        continue
      }
      if (STATEFUL.test(selector)) continue
      if (matchesAny(selector, subtree)) rules.push(rule.cssText)
    }
  }
  return rules.join('\n')
}

function matchesAny(selector: string, subtree: readonly Element[]): boolean {
  for (const part of selector.split(',')) {
    const trimmed = part.trim()
    if (trimmed === '') continue
    try {
      if (subtree.some((node) => node.matches(trimmed))) return true
    } catch {
      // An unsupported selector cannot match anything we captured.
    }
  }
  return false
}
