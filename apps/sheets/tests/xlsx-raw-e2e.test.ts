/**
 * The raw OOXML escape hatch against the real sidecar.
 *
 * The three string gates are unit-tested in xlsx-raw.test.ts. What needs a real
 * workbook is the fourth: an edit that leaves the package unreadable has to be
 * refused, not saved. This mirrors what the main process does — resolve, edit,
 * verify by making the sidecar open the candidate, then save — because the gate
 * is only worth anything if it sits between the edit and the file.
 */
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import JSZip from 'jszip'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'

import { saveWorkbookViaSidecar } from '../src/gateway/xlsx-package-io'
import {
  applyRawEdit,
  listRawParts,
  resolveRawPartRef,
  sheetPathsInOrder,
  type RawArchiveView,
} from '../src/gateway/xlsx-raw'
import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import { buildEditFixture } from './fixture-builder'

describe('raw OOXML edits against the sidecar', () => {
  let directory: string
  let client: XlsxSidecarClient
  let sourcePath: string

  beforeAll(async () => {
    directory = await mkdtemp(join(tmpdir(), 'xlsx-raw-e2e-'))
    client = new XlsxSidecarClient(sidecarBinaryPath())
    sourcePath = join(directory, 'source.xlsx')
    await writeFile(sourcePath, await buildEditFixture())
  })

  afterAll(async () => {
    client.stop()
    await rm(directory, { recursive: true, force: true })
  })

  const view = async (): Promise<RawArchiveView> => {
    const entries = (
      (await client.archiveManifest(sourcePath)) as {
        entries: { name: string; uncompressedSize: number }[]
      }
    ).entries
    const read = async (name: string): Promise<string> => {
      const zip = await JSZip.loadAsync(await readFile(sourcePath))
      return (await zip.file(name)?.async('text')) ?? ''
    }
    return {
      entries,
      sheetPaths: sheetPathsInOrder(
        await read('xl/workbook.xml'),
        await read('xl/_rels/workbook.xml.rels'),
      ),
    }
  }

  /** The main process's fourth gate, in the shape the handler uses it. */
  const stillReads = async (rawParts: Map<string, string>): Promise<string | null> => {
    const workDir = await mkdtemp(join(directory, 'check-'))
    try {
      const replacements: { name: string; contentPath: string }[] = []
      let i = 0
      for (const [name, xml] of rawParts) {
        const contentPath = join(workDir, `raw-${i++}.bin`)
        await writeFile(contentPath, xml, 'utf8')
        replacements.push({ name, contentPath })
      }
      const candidate = join(workDir, 'candidate.xlsx')
      await client.saveArchive({
        sourcePath,
        targetPath: candidate,
        replacements,
        removals: [],
        additions: [],
      })
      const opened = (await client.open(candidate)) as {
        sessionId: string
        sheets: { id: string }[]
      }
      try {
        const first = opened.sheets[0]
        if (first) {
          await client.readRange({
            sessionId: opened.sessionId,
            sheetId: first.id,
            range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
          })
        }
      } finally {
        await client.close(opened.sessionId)
      }
      return null
    } catch (error: unknown) {
      return error instanceof Error ? error.message : String(error)
    } finally {
      await rm(workDir, { recursive: true, force: true })
    }
  }

  it('resolves short names against the real package', async () => {
    const v = await view()
    expect(resolveRawPartRef(v, '/sheet[1]')).toBe('xl/worksheets/sheet1.xml')
    expect(resolveRawPartRef(v, '/styles')).toBe('xl/styles.xml')
    expect(listRawParts(v).find((p) => p.path === 'xl/workbook.xml')?.ref).toBe('/workbook')
  })

  it('accepts an edit the workbook still reads, and it survives the save', async () => {
    const v = await view()
    const path = resolveRawPartRef(v, '/sheet[1]')!
    const zip = await JSZip.loadAsync(await readFile(sourcePath))
    const before = (await zip.file(path)?.async('text')) ?? ''

    const edited = applyRawEdit(path, before, '<sheetData>', '<sheetPr><tabColor rgb="FF00B050"/></sheetPr>\n  <sheetData>')
    expect(edited.ok).toBe(true)
    if (!edited.ok) return

    const rawParts = new Map([[path, edited.xml]])
    expect(await stillReads(rawParts)).toBeNull()

    const targetPath = join(directory, 'accepted.xlsx')
    await saveWorkbookViaSidecar({ client, sourcePath, targetPath, edits: [], rawParts })
    const saved = await (await JSZip.loadAsync(await readFile(targetPath))).file(path)?.async('text')
    expect(saved).toContain('<tabColor rgb="FF00B050"/>')
  })

  it('refuses an edit that leaves the workbook unreadable', async () => {
    const v = await view()
    const path = resolveRawPartRef(v, '/workbook')!
    const zip = await JSZip.loadAsync(await readFile(sourcePath))
    const before = (await zip.file(path)?.async('text')) ?? ''

    // well-formed XML, and still nonsense as a workbook: the sheet list now
    // points at a relationship that does not exist. Only the reader catches it.
    const edited = applyRawEdit(path, before, 'r:id="rId1"', 'r:id="rIdNope"')
    expect(edited.ok).toBe(true)
    if (!edited.ok) return

    expect(await stillReads(new Map([[path, edited.xml]]))).not.toBeNull()
  })
})

function sidecarBinaryPath(): string {
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  return fileURLToPath(
    new URL(`../native/xlsx-engine/target/release/${executable}`, import.meta.url),
  )
}
