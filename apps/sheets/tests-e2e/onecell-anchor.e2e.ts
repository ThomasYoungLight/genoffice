/**
 * E3/E4 — a `oneCellAnchor` picture sizes correctly, and survives being moved.
 *
 * This is the regression net under the two worst bugs this project has had,
 * both of them in files written by openpyxl rather than Excel:
 *
 *   1. `parse_anchor` fell back to `to = from` for an anchor with one marker,
 *      ignoring the `<ext cx cy>` that carries its size. Every such drawing
 *      rendered as a sliver.
 *   2. The drawing rewriter hard-coded the `xdr:` prefix. openpyxl makes
 *      spreadsheetDrawing the default namespace and writes elements bare, so
 *      moving a picture in one of those files failed the save outright with
 *      "Drawing anchor #1 was not found".
 *
 * Both were invisible to unit tests: the first needed a real render, the
 * second a real save. Hence these.
 */
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import path from 'node:path'

import { afterAll, afterEach, beforeAll, describe, expect, it } from 'vitest'

import { buildOneCellFixture } from './helpers/fixtures'
import {
  gotoCell,
  launchSheets,
  openWorkbook,
  packageProblems,
  readAnchors,
  saveWorkbook,
  typeInCell,
  workingCopy,
  type SheetsApp,
} from './helpers/sheets-app'

/** The fixture's picture: 240x160 px, written as 2286000 x 1524000 EMU. */
const SOURCE_WIDTH_PX = 240
const SOURCE_HEIGHT_PX = 160
/** Zoom and sub-pixel rounding move the rendered box by a few px. */
const TOLERANCE_PX = 12

let fixtureDir: string
let onecell: string
let session: SheetsApp | undefined

beforeAll(() => {
  fixtureDir = mkdtempSync(path.join(tmpdir(), 'genoffice-e2e-anchor-'))
  onecell = buildOneCellFixture(fixtureDir)
})

afterAll(() => {
  rmSync(fixtureDir, { recursive: true, force: true })
})

afterEach(async () => {
  await session?.close()
  session = undefined
})

/** The rendered box of the workbook's one picture. */
async function pictureBox(session: SheetsApp): Promise<{
  x: number
  y: number
  width: number
  height: number
}> {
  const image = session.page.locator('img.xlsx-image').first()
  await image.waitFor({ timeout: 30_000 })
  const box = await image.boundingBox()
  if (!box) throw new Error('picture has no box')
  return box
}

describe('E3: a oneCellAnchor picture renders at its declared size', () => {
  it('is the size the <ext> says, not a sliver', async () => {
    // The sliver bug produced a box a few pixels tall. Asserting "greater than
    // zero" would have passed against it, so this asserts the actual size.
    session = await launchSheets({ workbook: onecell })
    await openWorkbook(session)
    const box = await pictureBox(session)

    expect(Math.abs(box.width - SOURCE_WIDTH_PX)).toBeLessThan(TOLERANCE_PX)
    expect(Math.abs(box.height - SOURCE_HEIGHT_PX)).toBeLessThan(TOLERANCE_PX)
  })

  it('keeps the source aspect ratio', async () => {
    // A height taken from the wrong marker still gives a plausible-looking
    // box; the ratio is what catches it.
    session = await launchSheets({ workbook: onecell })
    await openWorkbook(session)
    const box = await pictureBox(session)

    expect(box.width / box.height).toBeCloseTo(SOURCE_WIDTH_PX / SOURCE_HEIGHT_PX, 1)
  })
})

describe('E4: saving preserves the anchor', () => {
  it('leaves the picture alone when an unrelated cell is edited', async () => {
    // The app declines a no-op save ("No edits to save yet"), so the edit is
    // both what makes the save happen and the point: rewriting the drawing as
    // a twoCellAnchor would still open and still look right — until a row
    // above it is resized and the picture stretches with it.
    const target = workingCopy(fixtureDir, onecell, 'untouched.xlsx')
    session = await launchSheets({ workbook: target })
    await openWorkbook(session)
    await gotoCell(session.page, 'A1')
    await typeInCell(session.page, 'edited elsewhere')
    await saveWorkbook(session, target)

    const anchors = readAnchors(target)
    expect(anchors).toHaveLength(1)
    expect(anchors[0]).toMatchObject({
      kind: 'OneCellAnchor',
      fromCol: 3,
      fromRow: 3,
      extCx: 2286000,
      extCy: 1524000,
    })
  })

  it('moves the anchor when the picture is dragged, and still saves', async () => {
    // The `xdr:` bug failed exactly here: the file has no such prefix, the
    // rewriter looked for one, and the save reported a missing anchor.
    const target = workingCopy(fixtureDir, onecell, 'moved.xlsx')
    session = await launchSheets({ workbook: target })
    await openWorkbook(session)

    const before = await pictureBox(session)
    await session.page.mouse.move(before.x + before.width / 2, before.y + before.height / 2)
    await session.page.mouse.down()
    // Two steps: a single move can be coalesced away before the drag starts.
    await session.page.mouse.move(before.x + before.width / 2 + 60, before.y + before.height / 2 + 40)
    await session.page.mouse.move(before.x + before.width / 2 + 160, before.y + before.height / 2 + 120)
    await session.page.mouse.up()
    await session.page.waitForTimeout(800)

    await saveWorkbook(session, target)

    const anchors = readAnchors(target)
    expect(anchors).toHaveLength(1)
    const moved = anchors[0] as Record<string, number | string>
    // Down and to the right of D4, and still one-cell anchored at its own size.
    expect(moved.kind).toBe('OneCellAnchor')
    expect(Number(moved.fromCol)).toBeGreaterThan(3)
    expect(Number(moved.fromRow)).toBeGreaterThan(3)
    expect(moved.extCx).toBe(2286000)
    expect(moved.extCy).toBe(1524000)
  })

  it('writes a package with no missing parts after a move', async () => {
    const target = workingCopy(fixtureDir, onecell, 'moved-package.xlsx')
    session = await launchSheets({ workbook: target })
    await openWorkbook(session)

    const before = await pictureBox(session)
    await session.page.mouse.move(before.x + before.width / 2, before.y + before.height / 2)
    await session.page.mouse.down()
    await session.page.mouse.move(before.x + before.width / 2 + 80, before.y + before.height / 2 + 60)
    await session.page.mouse.move(before.x + before.width / 2 + 140, before.y + before.height / 2 + 100)
    await session.page.mouse.up()
    await session.page.waitForTimeout(800)
    await saveWorkbook(session, target)

    expect(packageProblems(target)).toEqual([])
  })

  it('reopens a moved picture at the same size it was saved with', async () => {
    // Closes the loop: a move that writes a good anchor but a bad extent would
    // pass every check above and shrink the picture on the next open.
    const target = workingCopy(fixtureDir, onecell, 'reopen.xlsx')
    session = await launchSheets({ workbook: target })
    await openWorkbook(session)

    const before = await pictureBox(session)
    await session.page.mouse.move(before.x + before.width / 2, before.y + before.height / 2)
    await session.page.mouse.down()
    await session.page.mouse.move(before.x + before.width / 2 + 70, before.y + before.height / 2 + 50)
    await session.page.mouse.move(before.x + before.width / 2 + 150, before.y + before.height / 2 + 110)
    await session.page.mouse.up()
    await session.page.waitForTimeout(800)
    await saveWorkbook(session, target)
    await session.close()

    session = await launchSheets({ workbook: target })
    await openWorkbook(session)
    const after = await pictureBox(session)

    expect(Math.abs(after.width - before.width)).toBeLessThan(TOLERANCE_PX)
    expect(Math.abs(after.height - before.height)).toBeLessThan(TOLERANCE_PX)
  })
})
