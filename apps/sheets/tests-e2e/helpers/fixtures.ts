/**
 * Fixtures for the end-to-end runs, generated rather than committed so what
 * they contain is readable in the diff.
 *
 * They are written by openpyxl on purpose. Two of the three real bugs this
 * project has hit came from files written by openpyxl and pandas rather than
 * Excel — a `oneCellAnchor` whose size lives in `<ext>`, and drawing XML with
 * no `xdr:` prefix — so an Excel-authored fixture would have missed both.
 */
import { execFileSync } from 'node:child_process'
import path from 'node:path'

function build(script: string, target: string): string {
  execFileSync('python3', ['-c', script, target], { encoding: 'utf8' })
  return target
}

/** Data, styles, a formula with no cached value, links, a table, frozen panes. */
export function buildGridFixture(dir: string): string {
  return build(
    `
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

wb = Workbook(); ws = wb.active; ws.title = "Render"
rows = [["Part","Note","Qty"],["007","alpha",1],["0080","beta",2],
        ["0900","gamma",3],["1000","delta",4],["1100","epsilon",5]]
for r,row in enumerate(rows, start=1):
    for c,v in enumerate(row, start=1): ws.cell(r,c,v)
for r in range(2,7): ws.cell(r,1).number_format = '@'
t = Table(displayName="T1", ref="A1:C6")
t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
ws.add_table(t)
ws["E1"] = "Formula:"; ws["E2"] = "=SUM(C2:C6)"
ws["E4"] = "line one\\nline two"
ws["E6"] = "red link"; ws["E6"].hyperlink = "https://example.com"
ws["E6"].font = Font(color="FFFF0000")
ws["G5"] = "filled"; ws["G5"].fill = PatternFill("solid", fgColor="FFFFFF00")
ws["G3"] = "right"; ws["G3"].alignment = Alignment(horizontal="right")
ws.freeze_panes = "B2"
wb.save(sys.argv[1])
`,
    path.join(dir, 'grid.xlsx'),
  )
}

/**
 * A picture on a `oneCellAnchor` — openpyxl's default, and the shape whose
 * size lives in `<ext cx cy>` rather than in a second marker.
 */
export function buildOneCellFixture(dir: string): string {
  return build(
    `
import sys
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import io, os, tempfile

png = os.path.join(tempfile.mkdtemp(), 'dot.png')
PILImage.new('RGB', (240, 160), (60, 110, 200)).save(png)

wb = Workbook(); ws = wb.active; ws.title = "Anchored"
for r in range(1, 25):
    ws.cell(r, 1, f"row {r}"); ws.cell(r, 2, r * 3)
img = Image(png)
ws.add_image(img, "D4")          # openpyxl writes this as a oneCellAnchor
wb.save(sys.argv[1])
`,
    path.join(dir, 'onecell.xlsx'),
  )
}

/** Enough rows that loading streams in blocks rather than arriving at once. */
export function buildLargeFixture(dir: string, rows = 4000): string {
  return build(
    `
import sys
from openpyxl import Workbook
wb = Workbook(); ws = wb.active; ws.title = "Big"
ws.append(["Row","Value"])
for r in range(1, ${rows} + 1):
    ws.append([f"row {r}", r * 3])
wb.save(sys.argv[1])
`,
    path.join(dir, 'large.xlsx'),
  )
}
