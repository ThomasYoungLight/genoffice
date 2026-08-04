/**
 * E1 — a workbook opens, renders, is edited, and saves to a file another
 * program can read.
 *
 * This is the case that was marked *partial* in the manual log: the Excel half
 * had only ever been run against a visual edit, never a cell edit. What runs
 * here is the cell-edit path, end to end through the real app, with the
 * result read back by openpyxl rather than by our own parser — so a bug that
 * writes and reads the same wrong bytes cannot pass.
 */
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import path from 'node:path'

import { afterAll, afterEach, beforeAll, describe, expect, it } from 'vitest'

import { buildGridFixture } from './helpers/fixtures'
import {
  currentStatus,
  gotoCell,
  launchSheets,
  openWorkbook,
  packageProblems,
  readCells,
  saveWorkbook,
  typeInCell,
  workingCopy,
  type SheetsApp,
} from './helpers/sheets-app'

let fixtureDir: string
let grid: string
let session: SheetsApp | undefined

beforeAll(() => {
  fixtureDir = mkdtempSync(path.join(tmpdir(), 'genoffice-e2e-fixtures-'))
  grid = buildGridFixture(fixtureDir)
})

afterAll(() => {
  rmSync(fixtureDir, { recursive: true, force: true })
})

afterEach(async () => {
  await session?.close()
  session = undefined
})

describe('E1: open, edit, save, reopen', () => {
  it('renders the workbook and reports it fully loaded', async () => {
    session = await launchSheets({ workbook: grid })
    await openWorkbook(session)
    expect(await currentStatus(session.page)).toMatch(/fully loaded/i)
  })

  it('writes an edited cell to a file openpyxl can read', async () => {
    // The copy lives beside the fixture, so the fixture itself is never written to.
    const target = workingCopy(fixtureDir, grid, 'edited.xlsx')
    session = await launchSheets({ workbook: target })
    await openWorkbook(session)
    await gotoCell(session.page, 'B3')
    await typeInCell(session.page, 'EDITED')
    await saveWorkbook(session, target)

    expect(readCells(target, 'Render', ['B3'])).toEqual({ B3: 'EDITED' })
  })

  it('leaves the rest of the sheet alone when one cell changes', async () => {
    // A save that rebuilds the worksheet from what happens to be loaded is the
    // way lazy loading destroys data: rows the user never scrolled to vanish.
    const target = workingCopy(fixtureDir, grid, 'untouched.xlsx')
    session = await launchSheets({ workbook: target })
    await openWorkbook(session)
    await gotoCell(session.page, 'B3')
    await typeInCell(session.page, 'EDITED')
    await saveWorkbook(session, target)

    expect(readCells(target, 'Render', ['A2', 'A6', 'B6', 'C6', 'G5'])).toEqual({
      A2: '007',
      A6: '1100',
      B6: 'epsilon',
      C6: 5,
      G5: 'filled',
    })
  })

  it('keeps a leading zero typed into a text-formatted cell', async () => {
    // A2:A6 carry the '@' text format. Rendering "007" correctly is not the
    // same as surviving a save: the value leaves through the journal, where it
    // can be re-typed as a number and lose the zeros the format exists for.
    const target = workingCopy(fixtureDir, grid, 'leading-zero.xlsx')
    session = await launchSheets({ workbook: target })
    await openWorkbook(session)
    await gotoCell(session.page, 'A3')
    await typeInCell(session.page, '0042')
    await saveWorkbook(session, target)

    const read = readCells(target, 'Render', ['A2', 'A3'])
    expect(read.A2).toBe('007')
    expect(String(read.A3)).toBe('0042')
  })

  it('treats a leading zero in a General cell as a number, as Excel does', async () => {
    // The counterpart, pinned so the case above cannot be "fixed" by forcing
    // every typed value to text. A7 has no format; 0042 there means 42.
    const target = workingCopy(fixtureDir, grid, 'general-cell.xlsx')
    session = await launchSheets({ workbook: target })
    await openWorkbook(session)
    await gotoCell(session.page, 'A7')
    await typeInCell(session.page, '0042')
    await saveWorkbook(session, target)

    expect(readCells(target, 'Render', ['A7'])).toEqual({ A7: 42 })
  })

  it('produces a package with no missing parts or undeclared content types', async () => {
    // The three faults that make Excel offer to repair a file. Necessary, not
    // sufficient — the manual Excel pass is what settles it.
    const target = workingCopy(fixtureDir, grid, 'package.xlsx')
    session = await launchSheets({ workbook: target })
    await openWorkbook(session)
    await gotoCell(session.page, 'B3')
    await typeInCell(session.page, 'EDITED')
    await saveWorkbook(session, target)

    expect(packageProblems(target)).toEqual([])
  })

  it('reopens the saved file and shows the edit', async () => {
    const target = workingCopy(fixtureDir, grid, 'reopen.xlsx')
    session = await launchSheets({ workbook: target })
    await openWorkbook(session)
    await gotoCell(session.page, 'B3')
    await typeInCell(session.page, 'REOPENED')
    await saveWorkbook(session, target)
    await session.close()

    session = await launchSheets({ workbook: target })
    await openWorkbook(session)
    expect(await currentStatus(session.page)).toMatch(/fully loaded/i)
    expect(readCells(target, 'Render', ['B3'])).toEqual({ B3: 'REOPENED' })
  })
})
