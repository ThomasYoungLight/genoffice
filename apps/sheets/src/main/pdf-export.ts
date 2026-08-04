/// PDF export: renders the print HTML (laid out by the renderer) in a hidden
/// scripting-disabled window and writes webContents.printToPDF's output where
/// the save dialog points.
///
/// renderPreview reuses that window for a different consumer: the agent, which
/// otherwise has no way to see what it built. Same HTML, captured as a PNG
/// rather than printed.

import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { BrowserWindow, dialog } from 'electron'

import type { IpcMainInvokeEvent } from 'electron'
import type {
  WorkbookExportPdfRequest,
  WorkbookExportPdfResult,
  WorkbookRenderPreviewRequest,
  WorkbookRenderPreviewResult,
} from '../shared/desktop-api'

export async function exportPdf(
  event: IpcMainInvokeEvent,
  request: WorkbookExportPdfRequest,
): Promise<WorkbookExportPdfResult> {
  const parent = BrowserWindow.fromWebContents(event.sender)
  const dialogOptions = {
    defaultPath: request.fileName,
    filters: [{ name: 'PDF', extensions: ['pdf'] }],
  }
  const selection = parent
    ? await dialog.showSaveDialog(parent, dialogOptions)
    : await dialog.showSaveDialog(dialogOptions)
  if (selection.canceled || !selection.filePath) return { canceled: true }

  const workDir = await mkdtemp(join(tmpdir(), 'ai-excel-pdf-'))
  const htmlPath = join(workDir, 'print.html')
  const window = new BrowserWindow({
    show: false,
    webPreferences: { sandbox: true, javascript: false },
  })
  try {
    await writeFile(htmlPath, request.html, 'utf8')
    await window.loadFile(htmlPath)
    const pdf = await window.webContents.printToPDF({
      landscape: request.landscape,
      pageSize: request.pageSize,
      margins: request.margins,
      scale: request.scale,
      printBackground: true,
    })
    await writeFile(selection.filePath, pdf)
    return { canceled: false, path: selection.filePath }
  } finally {
    window.destroy()
    await rm(workDir, { recursive: true, force: true })
  }
}

/// Lays the print HTML out at a fixed width and captures it as a PNG. No
/// dialog and no file on disk: the agent gets bytes back and nothing is left
/// behind. Scripting stays disabled, exactly as for the PDF path — the HTML is
/// built from workbook content and must never execute.
export async function renderPreview(
  request: WorkbookRenderPreviewRequest,
): Promise<WorkbookRenderPreviewResult> {
  const workDir = await mkdtemp(join(tmpdir(), 'ai-excel-preview-'))
  const htmlPath = join(workDir, 'preview.html')
  const window = new BrowserWindow({
    show: false,
    width: request.width,
    height: request.maxHeight,
    webPreferences: { sandbox: true, javascript: false, offscreen: true },
  })
  try {
    await writeFile(htmlPath, request.html, 'utf8')
    await window.loadFile(htmlPath)
    // The table decides its own height; measuring it needs the document, and
    // scripting is off — so resize to the cap and let the capture clip. A
    // shorter sheet simply leaves white space, which reads fine.
    const image = await window.webContents.capturePage()
    const size = image.getSize()
    const png = image.toPNG()
    return {
      base64: png.toString('base64'),
      width: size.width,
      height: size.height,
      truncated: size.height >= request.maxHeight,
    }
  } finally {
    window.destroy()
    await rm(workDir, { recursive: true, force: true })
  }
}
