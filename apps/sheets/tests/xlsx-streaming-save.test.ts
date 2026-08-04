import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import JSZip from 'jszip'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'

import type { ArchiveEntry } from '../src/gateway/xlsx-package-io'
import { assertManifestPreserved, saveWorkbookViaSidecar } from '../src/gateway/xlsx-package-io'
import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import { buildEditFixture } from './fixture-builder'

describe('saveWorkbookViaSidecar', () => {
  let directory: string
  let client: XlsxSidecarClient

  beforeAll(async () => {
    directory = await mkdtemp(join(tmpdir(), 'xlsx-streaming-save-'))
    client = new XlsxSidecarClient(sidecarBinaryPath())
  })

  afterAll(async () => {
    client.stop()
    await rm(directory, { recursive: true, force: true })
  })

  /**
   * Raw OOXML edits are an overlay applied at save. The two things that can go
   * wrong are the part not being written at all (the planner never touched it,
   * so it is absent from plan.replaced) and a raw edit racing a model edit to
   * the same part, where one silently discards the other.
   */
  it('writes a raw part the planner never touched, and preserves the rest', async () => {
    const sourcePath = join(directory, 'raw-source.xlsx')
    const targetPath = join(directory, 'raw-saved.xlsx')
    const sourceBuffer = await buildEditFixture()
    await writeFile(sourcePath, sourceBuffer)

    const result = await saveWorkbookViaSidecar({
      client,
      sourcePath,
      targetPath,
      edits: [],
      rawParts: new Map([
        ['xl/styles.xml', '<?xml version="1.0"?>\n<styleSheet xmlns="x"><marker/></styleSheet>'],
      ]),
    })

    expect(result.touchedEntries).toContain('xl/styles.xml')
    const savedZip = await JSZip.loadAsync(await readFile(targetPath))
    expect(await savedZip.file('xl/styles.xml')?.async('text')).toContain('<marker/>')
    // the escape hatch must not disturb anything it was not pointed at
    const sourceZip = await JSZip.loadAsync(sourceBuffer)
    expect(Object.keys(savedZip.files).sort()).toEqual(Object.keys(sourceZip.files).sort())
    expect(await savedZip.file('customXml/item1.xml')?.async('text')).toBe(
      await sourceZip.file('customXml/item1.xml')?.async('text'),
    )
  })

  it('composes a raw edit with a model edit to the same part', async () => {
    const sourcePath = join(directory, 'raw-both-source.xlsx')
    const targetPath = join(directory, 'raw-both-saved.xlsx')
    await writeFile(sourcePath, await buildEditFixture())

    // the planner reads through the overlay, so its rewrite of sheet1 starts
    // from the raw text instead of the bytes on disk
    const rawSheet =
      '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n' +
      '  <sheetPr><tabColor rgb="FFFF0000"/></sheetPr>\n' +
      '  <sheetData>\n' +
      '    <row r="1"><c r="A1" t="s" s="1"><v>0</v></c><c r="C1"><v>5</v></c></row>\n' +
      '  </sheetData>\n' +
      '</worksheet>'

    await saveWorkbookViaSidecar({
      client,
      sourcePath,
      targetPath,
      edits: [{ sheetName: 'Data', row: 0, column: 0, writeValue: true, cell: { value: 'World' } }],
      rawParts: new Map([['xl/worksheets/sheet1.xml', rawSheet]]),
    })

    const saved = await (await JSZip.loadAsync(await readFile(targetPath)))
      .file('xl/worksheets/sheet1.xml')
      ?.async('text')
    // both survive: neither overwrote the other
    expect(saved).toContain('<tabColor rgb="FFFF0000"/>')
    expect(saved).toContain('World')
  })

  it('saves a cell edit while raw-copying every untouched entry byte-for-byte', async () => {
    const sourcePath = join(directory, 'source.xlsx')
    const targetPath = join(directory, 'saved.xlsx')
    const sourceBuffer = await buildEditFixture()
    await writeFile(sourcePath, sourceBuffer)

    const result = await saveWorkbookViaSidecar({
      client,
      sourcePath,
      targetPath,
      edits: [
        {
          sheetName: 'Data',
          row: 0,
          column: 0,
          writeValue: true,
          cell: { value: 'World' },
        },
      ],
    })

    expect(result.touchedEntries).toEqual(['xl/workbook.xml', 'xl/worksheets/sheet1.xml'])
    expect(result.removedEntries).toEqual([])
    expect(result.addedEntries).toEqual([])

    const sourceZip = await JSZip.loadAsync(sourceBuffer)
    const savedZip = await JSZip.loadAsync(await readFile(targetPath))
    expect(Object.keys(savedZip.files).sort()).toEqual(Object.keys(sourceZip.files).sort())

    const savedSheet = await savedZip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(savedSheet).toContain('<is><t xml:space="preserve">World</t></is>')
    const savedWorkbook = await savedZip.file('xl/workbook.xml')?.async('text')
    expect(savedWorkbook).toContain('fullCalcOnLoad="1"')

    for (const path of Object.keys(sourceZip.files)) {
      if (sourceZip.files[path]?.dir) continue
      if (path === 'xl/worksheets/sheet1.xml' || path === 'xl/workbook.xml') continue
      const [sourceBytes, savedBytes] = await Promise.all([
        sourceZip.file(path)?.async('nodebuffer'),
        savedZip.file(path)?.async('nodebuffer'),
      ])
      expect(savedBytes?.equals(sourceBytes ?? Buffer.of())).toBe(true)
    }
  })

  it('reopens its own output: a second edit chains on the saved file', async () => {
    const firstPath = join(directory, 'chain-1.xlsx')
    const secondPath = join(directory, 'chain-2.xlsx')
    await writeFile(firstPath, await buildEditFixture())

    await saveWorkbookViaSidecar({
      client,
      sourcePath: firstPath,
      targetPath: secondPath,
      edits: [{ sheetName: 'Data', row: 4, column: 0, writeValue: true, cell: { value: 41 } }],
    })
    await saveWorkbookViaSidecar({
      client,
      sourcePath: secondPath,
      targetPath: secondPath,
      edits: [{ sheetName: 'Data', row: 5, column: 0, writeValue: true, cell: { value: 42 } }],
    })

    const savedZip = await JSZip.loadAsync(await readFile(secondPath))
    const sheet = await savedZip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(sheet).toContain('<c r="A5"><v>41</v></c>')
    expect(sheet).toContain('<c r="A6"><v>42</v></c>')
  })
})

describe('assertManifestPreserved', () => {
  const entry = (name: string, crc32 = 1, size = 10): ArchiveEntry => ({
    name,
    crc32,
    compressedSize: size,
    uncompressedSize: size,
  })
  const plan = (
    replaced: string[] = [],
    added: string[] = [],
    removedEntries: string[] = [],
  ): Parameters<typeof assertManifestPreserved>[0] => ({
    replaced: new Map(replaced.map((name) => [name, ''])),
    added: new Map(added.map((name) => [name, ''])),
    removedEntries,
  })

  it('accepts identical untouched entries and declared changes', () => {
    expect(() =>
      assertManifestPreserved(
        plan(['b.xml'], ['c.xml'], ['d.xml']),
        [entry('a.xml'), entry('b.xml'), entry('d.xml')],
        [entry('a.xml'), entry('b.xml', 99, 12), entry('c.xml', 7)],
      ),
    ).not.toThrow()
  })

  it('rejects an undeclared content change', () => {
    expect(() => assertManifestPreserved(plan(), [entry('a.xml', 1)], [entry('a.xml', 2)])).toThrow(
      'unexpectedly modify a.xml',
    )
  })

  it('rejects dropped, surviving-despite-removal, and surprise entries', () => {
    expect(() => assertManifestPreserved(plan(), [entry('a.xml')], [])).toThrow('drop a.xml')
    expect(() =>
      assertManifestPreserved(plan([], [], ['a.xml']), [entry('a.xml')], [entry('a.xml')]),
    ).toThrow('should have removed a.xml')
    expect(() =>
      assertManifestPreserved(plan(), [entry('a.xml')], [entry('a.xml'), entry('extra.xml')]),
    ).toThrow('unexpectedly create extra.xml')
  })

  it('rejects a missing declared addition', () => {
    expect(() =>
      assertManifestPreserved(plan([], ['new.xml']), [entry('a.xml')], [entry('a.xml')]),
    ).toThrow('should have created new.xml')
  })
})

function sidecarBinaryPath(): string {
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  return fileURLToPath(
    new URL(`../native/xlsx-engine/target/release/${executable}`, import.meta.url),
  )
}
