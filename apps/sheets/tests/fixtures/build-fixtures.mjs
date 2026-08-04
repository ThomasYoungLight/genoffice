/**
 * Builds the xlsx fixtures the E2E cases in docs/testing/e2e-test-cases.md use.
 *
 * Generated rather than committed, so what each file contains is readable here
 * instead of being a binary nobody can diff. Needs python3 with openpyxl and
 * pillow (`pip3 install openpyxl pillow`); openpyxl is the point rather than an
 * accident — it writes oneCellAnchor drawings and bare-namespace drawing parts,
 * which is exactly the shape Excel never produces and our bugs lived in.
 *
 *   node tests/fixtures/build-fixtures.mjs [outDir]
 */
import { execFileSync } from 'node:child_process'
import { mkdirSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const outDir = process.argv[2] ?? join(tmpdir(), 'genoffice-sheets-fixtures')
mkdirSync(outDir, { recursive: true })

const script = `
import struct, zlib, os, sys
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, IconSetRule, CellIsRule
from openpyxl.styles import PatternFill

out = sys.argv[1]

def png(w, h, rgb):
    raw = b''.join(b'\\x00' + bytes(rgb) * w for _ in range(h))
    def chunk(t, d):
        c = t + d
        return struct.pack('>I', len(d)) + c + struct.pack('>I', zlib.crc32(c))
    return (b'\\x89PNG\\r\\n\\x1a\\n'
            + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b''))

logo = os.path.join(out, 'logo.png')
open(logo, 'wb').write(png(160, 80, (0xC0, 0x39, 0x2B)))

def revenue_sheet(ws):
    for r in [('Region','Revenue'),('North',120),('South',96),('East',143),('West',88)]:
        ws.append(r)

def bar(ws, title, first, last):
    c = BarChart(); c.title = title
    c.add_data(Reference(ws, min_col=2, min_row=first, max_row=last), titles_from_data=(first == 1))
    c.set_categories(Reference(ws, min_col=1, min_row=max(first, 2), max_row=last))
    return c

def two_cell(fc, fr, tc, tr):
    return TwoCellAnchor(editAs='twoCell',
        _from=AnchorMarker(col=fc, colOff=0, row=fr, rowOff=0),
        to=AnchorMarker(col=tc, colOff=0, row=tr, rowOff=0))

# --- twocell.xlsx: what Excel itself writes -------------------------------
wb = Workbook(); ws = wb.active; ws.title = 'Data'; revenue_sheet(ws)
c = bar(ws, 'Revenue by Region', 1, 5); c.anchor = two_cell(3, 1, 8, 12); ws.add_chart(c)
img = XLImage(logo); img.anchor = two_cell(3, 14, 6, 19); ws.add_image(img)
wb.save(os.path.join(out, 'twocell.xlsx'))

# --- onecell.xlsx: what openpyxl/pandas write by default ------------------
wb = Workbook(); ws = wb.active; ws.title = 'Data'; revenue_sheet(ws)
c = bar(ws, 'Revenue by Region', 1, 5); c.height = 6; c.width = 11
ws.add_chart(c, 'D2')                      # oneCellAnchor + <ext>
ws.add_image(XLImage(logo), 'D16')
wb.save(os.path.join(out, 'onecell.xlsx'))

# --- multipage.xlsx: visuals at four depths, several pages ----------------
wb = Workbook(); ws = wb.active; ws.title = 'Data'
ws.append(('Row','Value'))
for i in range(1, 121): ws.append((f'row {i}', i * 3 % 97))
for title, first, last, fr, tr in [('TOP', 2, 9, 1, 11),
                                   ('MIDDLE', 40, 47, 44, 54),
                                   ('BOTTOM', 100, 107, 99, 109)]:
    c = bar(ws, title, first, last); c.anchor = two_cell(3, fr, 8, tr); ws.add_chart(c)
img = XLImage(logo); img.anchor = two_cell(3, 74, 6, 79); ws.add_image(img)
wb.save(os.path.join(out, 'multipage.xlsx'))

# --- cf.xlsx: one rule of each kind ---------------------------------------
wb = Workbook(); ws = wb.active; ws.title = 'CF'
ws.append(('Region','Score','Bar','Icon','Flag'))
for r in [('North',95,95,95,120),('South',41,41,41,40),('East',78,78,78,143),
          ('West',12,12,12,20),('Mid',63,63,63,88),('Far',88,88,88,101)]:
    ws.append(r)
ws.conditional_formatting.add('B2:B7', ColorScaleRule(
    start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50,
    mid_color='FFEB84', end_type='max', end_color='63BE7B'))
ws.conditional_formatting.add('C2:C7', DataBarRule(start_type='min', end_type='max', color='638EC6'))
ws.conditional_formatting.add('D2:D7', IconSetRule('3TrafficLights1', 'percent', [0, 33, 67]))
ws.conditional_formatting.add('E2:E7', CellIsRule(
    operator='greaterThan', formula=['100'],
    fill=PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')))
wb.save(os.path.join(out, 'cf.xlsx'))

# --- filtered.xlsx: an AutoFilter to drive filter cases -------------------
wb = Workbook(); ws = wb.active; ws.title = 'Data'
for r in [('Region','Rep','Revenue'),('North','Ana',120),('South','Bo',96),
          ('East','Cy',143),('West','Di',88),('North','Eve',77),('South','Fay',54)]:
    ws.append(r)
ws.auto_filter.ref = 'A1:C7'
wb.save(os.path.join(out, 'filtered.xlsx'))

print(out)
`

try {
  const written = execFileSync('python3', ['-c', script, outDir], { encoding: 'utf8' }).trim()
  console.log(`fixtures written to ${written}`)
} catch (error) {
  console.error('Fixture build failed. Needs: pip3 install openpyxl pillow')
  console.error(error instanceof Error ? error.message : String(error))
  process.exitCode = 1
}
