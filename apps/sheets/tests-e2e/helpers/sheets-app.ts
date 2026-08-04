/**
 * Driving the real Sheets app for end-to-end tests.
 *
 * These tests launch the built Electron binary, act through the UI, and check
 * the result with a reader that is not ours — the standing rule for this
 * project is that no format-writing feature is done until the file has been
 * opened by the application that has to read it. openpyxl is the stand-in for
 * that reader in CI; Excel itself still has to be run by hand.
 *
 * Nothing here reaches into renderer internals. The only privileged channel
 * used is `menu:action`, which is exactly what the real application menu
 * sends, so a test drives the same path a user does.
 */
import { execFileSync } from 'node:child_process'
import { copyFileSync, mkdirSync, mkdtempSync, rmSync, statSync } from 'node:fs'
import { tmpdir } from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

import { _electron, type ElectronApplication, type Page } from 'playwright-core'

const HERE = path.dirname(fileURLToPath(import.meta.url))
/** apps/sheets */
export const APP_ROOT = path.resolve(HERE, '../..')
const REPO_ROOT = path.resolve(APP_ROOT, '../..')

const ELECTRON = path.join(
  REPO_ROOT,
  'node_modules/electron/dist/Electron.app/Contents/MacOS/Electron',
)
const MAIN = path.join(APP_ROOT, 'out/main/index.js')
const SIDECAR = path.join(APP_ROOT, 'native/xlsx-engine/target/release/xlsx-sidecar')

/** Opening a large workbook goes through the sidecar and a streaming index. */
const LOAD_TIMEOUT_MS = 90_000

export interface SheetsApp {
  readonly app: ElectronApplication
  readonly page: Page
  /** Scratch directory for this run; removed on close. */
  readonly workDir: string
  close(): Promise<void>
}

/**
 * Launches the app against a private userData directory.
 *
 * The isolation is not incidental: without it a test would read the developer's
 * own settings and, worse, its autosave-recovery cases would offer to restore
 * the developer's unsaved work.
 */
export async function launchSheets(options: { workbook?: string } = {}): Promise<SheetsApp> {
  const workDir = mkdtempSync(path.join(tmpdir(), 'genoffice-e2e-'))
  const userData = path.join(workDir, 'userData')
  mkdirSync(userData, { recursive: true })

  const app = await _electron.launch({
    executablePath: ELECTRON,
    args: [MAIN, `--user-data-dir=${userData}`],
    cwd: APP_ROOT,
    env: {
      ...process.env,
      XLSX_SIDECAR_PATH: SIDECAR,
      // Consumed by the main process when the renderer asks to open, so the
      // native file dialog never appears. The click path is unchanged.
      ...(options.workbook ? { XLSX_OPEN_PATH: options.workbook } : {}),
    },
  })
  const page = await app.firstWindow()
  await page.waitForLoadState('domcontentloaded')

  return {
    app,
    page,
    workDir,
    async close() {
      await app.close().catch(() => undefined)
      rmSync(workDir, { recursive: true, force: true })
    },
  }
}

/** Sends what the application menu sends. */
export async function menuAction(app: ElectronApplication, action: string): Promise<void> {
  await app.evaluate(({ webContents }, sent) => {
    for (const contents of webContents.getAllWebContents()) {
      contents.send('menu:action', sent)
    }
  }, action)
}

/** Triggers Open and waits for the workbook to finish streaming in. */
export async function openWorkbook(session: SheetsApp): Promise<void> {
  await menuAction(session.app, 'open')
  await waitForStatus(session.page, /fully loaded/i)
}

export async function waitForStatus(page: Page, text: RegExp): Promise<string> {
  const status = page.locator('.workbook-status')
  await status.filter({ hasText: text }).waitFor({ timeout: LOAD_TIMEOUT_MS })
  return (await status.first().textContent()) ?? ''
}

/** Reads the status line without waiting, for asserting what did *not* happen. */
export async function currentStatus(page: Page): Promise<string> {
  return (await page.locator('.workbook-status').first().textContent()) ?? ''
}

/**
 * Moves the selection with the Name Box, the way a user jumps to a cell.
 *
 * Clicking the grid would mean computing pixel coordinates from column widths,
 * which makes the test a measurement of our own layout maths rather than of
 * the thing it is trying to check.
 */
export async function gotoCell(page: Page, ref: string): Promise<void> {
  const nameBox = page.locator('input.name-box')
  await nameBox.click()
  await nameBox.fill(ref)
  await nameBox.press('Enter')
  // The grid takes focus back asynchronously; typing too early is swallowed.
  await page.waitForTimeout(300)
}

/** Types into the selected cell and commits, as a user would. */
export async function typeInCell(page: Page, value: string): Promise<void> {
  await page.keyboard.type(value, { delay: 25 })
  await page.keyboard.press('Enter')
  await page.waitForTimeout(400)
}

/**
 * Saves and waits for the bytes on disk to change.
 *
 * Not for the status line: the save message is replaced by the reload's own
 * message almost immediately, so a test that waits for "Saved …" races the UI
 * and reports a timeout for a save that worked. The file is the thing being
 * asserted about anyway.
 */
export async function saveWorkbook(
  session: SheetsApp,
  file: string,
  timeoutMs = 60_000,
): Promise<void> {
  const before = statSync(file)
  await menuAction(session.app, 'save')
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    await session.page.waitForTimeout(250)
    const now = statSync(file)
    if (now.mtimeMs !== before.mtimeMs || now.size !== before.size) {
      // The writer replaces the package; give the final rename a moment to
      // settle before a reader opens it.
      await session.page.waitForTimeout(400)
      return
    }
  }
  throw new Error(
    `save did not change ${path.basename(file)} within ${timeoutMs}ms ` +
      `(status: ${await currentStatus(session.page)})`,
  )
}

/**
 * Copies a fixture so a test can write to it without destroying the original.
 *
 * `dir` must outlive the app session — a session's own workDir is removed on
 * close, which would take the copy with it before a relaunch could open it.
 */
export function workingCopy(dir: string, source: string, name?: string): string {
  const target = path.join(dir, name ?? path.basename(source))
  copyFileSync(source, target)
  return target
}

/**
 * Reads cells back with openpyxl — deliberately not our own parser, so a bug
 * that writes and reads the same wrong bytes cannot pass.
 */
export function readCells(
  file: string,
  sheet: string,
  addresses: readonly string[],
): Record<string, unknown> {
  const script = `
import json, sys
from openpyxl import load_workbook
wb = load_workbook(sys.argv[1])
ws = wb[sys.argv[2]]
print(json.dumps({a: ws[a].value for a in sys.argv[3:]}, default=str))
`
  const out = execFileSync('python3', ['-c', script, file, sheet, ...addresses], {
    encoding: 'utf8',
  })
  return JSON.parse(out) as Record<string, unknown>
}

/** Every drawing anchor in a sheet, as openpyxl understands them. */
export function readAnchors(file: string): Array<Record<string, unknown>> {
  const script = `
import json, sys
from openpyxl import load_workbook
wb = load_workbook(sys.argv[1])
out = []
for ws in wb.worksheets:
    for img in getattr(ws, '_images', []):
        a = img.anchor
        out.append({
            'kind': type(a).__name__,
            'fromCol': a._from.col, 'fromRow': a._from.row,
            'toCol': getattr(getattr(a, 'to', None), 'col', None),
            'toRow': getattr(getattr(a, 'to', None), 'row', None),
            'extCx': getattr(getattr(a, 'ext', None), 'cx', None),
            'extCy': getattr(getattr(a, 'ext', None), 'cy', None),
        })
print(json.dumps(out))
`
  return JSON.parse(execFileSync('python3', ['-c', script, file], { encoding: 'utf8' })) as Array<
    Record<string, unknown>
  >
}

/**
 * Whether Excel would show a repair prompt.
 *
 * We cannot ask Excel from CI, so this checks the things that provoke one and
 * that our own writer can plausibly get wrong: an unreadable zip, a part the
 * content types do not declare, or a relationship pointing at a missing part.
 * A clean result here is necessary, not sufficient — the manual Excel pass in
 * `docs/testing/e2e-test-cases.md` is what actually settles it.
 */
export function packageProblems(file: string): string[] {
  const script = `
import json, sys, zipfile, posixpath
from xml.etree import ElementTree as ET
problems = []
try:
    z = zipfile.ZipFile(sys.argv[1])
except Exception as exc:
    print(json.dumps([f'unreadable zip: {exc}'])); raise SystemExit
bad = z.testzip()
if bad: problems.append(f'corrupt entry: {bad}')
names = set(z.namelist())
CT = '{http://schemas.openxmlformats.org/package/2006/content-types}'
try:
    ct = ET.fromstring(z.read('[Content_Types].xml'))
except Exception as exc:
    problems.append(f'content types unreadable: {exc}'); ct = None
if ct is not None:
    defaults = {e.get('Extension','').lower() for e in ct.findall(CT+'Default')}
    overrides = {e.get('PartName','').lstrip('/') for e in ct.findall(CT+'Override')}
    for n in names:
        if n.startswith('_rels/') or n.endswith('/') or n == '[Content_Types].xml': continue
        if '/_rels/' in n: continue
        ext = n.rsplit('.', 1)[-1].lower() if '.' in n else ''
        if n not in overrides and ext not in defaults:
            problems.append(f'part not declared in [Content_Types].xml: {n}')
R = '{http://schemas.openxmlformats.org/package/2006/relationships}'
for n in [x for x in names if x.endswith('.rels')]:
    base = posixpath.dirname(posixpath.dirname(n))
    try:
        rels = ET.fromstring(z.read(n))
    except Exception as exc:
        problems.append(f'{n} unreadable: {exc}'); continue
    for rel in rels.findall(R+'Relationship'):
        if rel.get('TargetMode') == 'External': continue
        target = rel.get('Target','')
        if target.startswith('/'):
            resolved = target.lstrip('/')
        else:
            resolved = posixpath.normpath(posixpath.join(base, target)).lstrip('/')
        if resolved not in names:
            problems.append(f'{n} -> missing part {resolved}')
print(json.dumps(problems))
`
  return JSON.parse(execFileSync('python3', ['-c', script, file], { encoding: 'utf8' })) as string[]
}
