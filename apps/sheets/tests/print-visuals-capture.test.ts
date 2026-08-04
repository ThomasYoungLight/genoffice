/**
 * @vitest-environment jsdom
 *
 * The DOM half of printing visuals. print-visuals.test.ts feeds
 * buildSheetPrintPayload markup it wrote by hand, so it proves the layout
 * places a fragment correctly but says nothing about where fragments come
 * from. This covers that: what capturePrintVisuals lifts off the live grid,
 * and — more to the point — what it refuses to lift.
 *
 * Only this file needs a DOM, so the environment is set per-file rather than
 * flipping the whole sheets suite (which is node, and much faster for it).
 */
import { beforeEach, describe, expect, it } from 'vitest'

import { capturePrintVisuals } from '../src/renderer/print-visuals'

const ANCHOR = '1,3,0,0,12,8,0,0'

function mount(html: string): HTMLElement {
  document.body.innerHTML = html
  return document.body
}

/// The shape of what installWorkbookVisuals renders: a display:contents
/// marker carrying the anchor, wrapping the visual itself.
function marker(inner: string, anchor = ANCHOR): string {
  return `<div style="display:contents" data-print-anchor="${anchor}">${inner}</div>`
}

describe('capturePrintVisuals', () => {
  beforeEach(() => {
    document.body.innerHTML = ''
    for (const sheet of Array.from(document.querySelectorAll('style'))) sheet.remove()
  })

  it('reads the anchor off the marker and the markup off the visual', () => {
    const { visuals } = capturePrintVisuals(mount(marker('<figure class="xlsx-chart">x</figure>')))
    expect(visuals).toHaveLength(1)
    expect(visuals[0]).toMatchObject({
      fromRow: 1,
      fromColumn: 3,
      toRow: 12,
      toColumn: 8,
      fromRowOffset: 0,
      toColumnOffset: 0,
    })
    expect(visuals[0]?.html).toBe('<figure class="xlsx-chart">x</figure>')
  })

  it('captures an image with its resolved bytes, not a fetch that has to be redone', () => {
    // The whole reason capture reads the DOM rather than re-rendering: by the
    // time the user prints, readWorkbookMedia has already run and the src is
    // a data URI. A static re-render would have to redo it asynchronously.
    const { visuals } = capturePrintVisuals(
      mount(marker('<img class="xlsx-image" src="data:image/png;base64,AAAA" alt="logo">')),
    )
    expect(visuals[0]?.html).toContain('src="data:image/png;base64,AAAA"')
  })

  it('captures a drawn shape, svg and text together', () => {
    const shape =
      '<div class="xlsx-shape-drawn"><svg viewBox="0 0 100 100">' +
      '<rect fill="#DDEBF7"></rect></svg><span class="shape-text">Step 1</span></div>'
    const { visuals } = capturePrintVisuals(mount(marker(shape)))
    expect(visuals[0]?.html).toContain('<svg')
    expect(visuals[0]?.html).toContain('Step 1')
  })

  it('unwraps the editable wrapper and drops every editing affordance', () => {
    const editable = marker(
      '<div class="shape-editable selected" tabindex="0" title="drag me">' +
        '<figure class="xlsx-chart"><button class="chart-edit-button">✎</button>bars</figure>' +
        '<button class="shape-delete-button">✕</button>' +
        '<span class="shape-handle handle-se" data-corner="se"></span>' +
        '</div>',
    )
    const html = capturePrintVisuals(mount(editable)).visuals[0]?.html ?? ''
    expect(html.startsWith('<figure class="xlsx-chart"')).toBe(true)
    expect(html).toContain('bars')
    expect(html).not.toContain('shape-delete-button')
    expect(html).not.toContain('chart-edit-button')
    expect(html).not.toContain('shape-handle')
    // selection is transient state, not part of the document
    expect(html).not.toContain('selected')
    expect(html).not.toContain('tabindex')
  })

  it('skips a visual that has not resolved — a PDF should not say "Loading…"', () => {
    const pending = marker('<div class="xlsx-visual-loading">Loading…</div>')
    const failed = marker('<div class="xlsx-visual-error">Image failed</div>', '2,0,0,0,4,2,0,0')
    const { visuals } = capturePrintVisuals(mount(pending + failed))
    expect(visuals).toHaveLength(0)
  })

  it('ignores a marker whose anchor is not eight finite numbers', () => {
    const bad = marker('<figure class="xlsx-chart">x</figure>', '1,3,0,0')
    const worse = marker('<figure class="xlsx-chart">y</figure>', 'a,b,c,d,e,f,g,h')
    expect(capturePrintVisuals(mount(bad + worse)).visuals).toHaveLength(0)
  })

  it('collects the rules its fragments need and leaves the rest of the page behind', () => {
    const style = document.createElement('style')
    style.textContent = [
      ':root { --excel-green: #107c41; }',
      '.xlsx-chart { border: 1px solid #c9cdd1; }',
      '.ribbon-button { padding: 4px; }',
    ].join('\n')
    document.head.appendChild(style)
    const { css } = capturePrintVisuals(mount(marker('<figure class="xlsx-chart">x</figure>')))
    expect(css).toContain('--excel-green')
    // cssText normalises the hex to rgb(), so match the rule, not the literal
    expect(css).toContain('.xlsx-chart')
    expect(css).toContain('rgb(201, 205, 209)')
    // nothing about the ribbon belongs in an exported page
    expect(css).not.toContain('ribbon-button')
    style.remove()
  })

  it('leaves out rules that describe a pointer the printed page will never have', () => {
    const style = document.createElement('style')
    style.textContent =
      '.xlsx-chart { border: 1px solid #c9cdd1; }\n.xlsx-chart:hover { border-color: red; }'
    document.head.appendChild(style)
    const { css } = capturePrintVisuals(mount(marker('<figure class="xlsx-chart">x</figure>')))
    expect(css).toContain('rgb(201, 205, 209)')
    expect(css).not.toContain(':hover')
    expect(css).not.toContain('red')
    style.remove()
  })

  it('returns nothing at all when the grid has no visuals', () => {
    expect(capturePrintVisuals(mount('<div class="univer-grid"></div>'))).toEqual({
      visuals: [],
      css: '',
    })
  })
})
