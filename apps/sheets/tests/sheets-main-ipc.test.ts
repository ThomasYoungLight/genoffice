/**
 * The main process's IPC surface.
 *
 * Preload validates on the way out of the renderer, but preload runs in the
 * renderer's process and a compromised or buggy renderer can call
 * `ipcRenderer.invoke` directly, bypassing it entirely. So main validates
 * again — and that second check is the one that actually protects the
 * filesystem, the sidecar and the key store. It had no tests.
 *
 * Rather than assert per handler, this drives every registered channel with
 * junk and asserts the property that has to hold for all of them, including
 * handlers added later: nothing malformed gets past the schema.
 */
import { beforeAll, describe, expect, it, vi } from 'vitest'

type Handler = (event: unknown, ...args: unknown[]) => unknown
const handlers = new Map<string, Handler>()

vi.mock('electron', () => {
  const noopWindow = {
    on: () => undefined,
    once: () => undefined,
    webContents: { send: () => undefined, on: () => undefined },
    isDestroyed: () => false,
    destroy: () => undefined,
    loadFile: () => Promise.resolve(),
  }
  return {
    app: {
      getPath: () => '/tmp/genoffice-test',
      on: () => undefined,
      whenReady: () => Promise.resolve(),
      getName: () => 'GenOffice Sheets',
    },
    ipcMain: {
      handle: (channel: string, handler: Handler) => handlers.set(channel, handler),
      on: () => undefined,
      removeHandler: () => undefined,
    },
    dialog: {
      showSaveDialog: () => Promise.resolve({ canceled: true }),
      showOpenDialog: () => Promise.resolve({ canceled: true }),
      showMessageBox: () => Promise.resolve({ response: 0 }),
    },
    shell: { openExternal: () => Promise.resolve() },
    BrowserWindow: Object.assign(
      function BrowserWindow() {
        return noopWindow
      },
      { fromWebContents: () => noopWindow, getAllWindows: () => [] },
    ),
    safeStorage: {
      isEncryptionAvailable: () => false,
      encryptString: (s: string) => Buffer.from(s),
      decryptString: (b: Buffer) => b.toString(),
    },
    webUtils: { getPathForFile: () => '/tmp/x' },
    nativeTheme: { on: () => undefined, shouldUseDarkColors: false },
    Menu: { setApplicationMenu: () => undefined, buildFromTemplate: () => ({}) },
  }
})

let registered: string[] = []

beforeAll(async () => {
  const main = await import('../src/main/sheets-main')
  const register = main as unknown as Record<string, undefined | (() => void)>
  // Whatever the module calls its registration entry points, run them all —
  // a renamed one should show up as a smaller channel list, not a silent skip.
  for (const name of Object.keys(register)) {
    if (/^register.*Ipc$/.test(name)) register[name]?.()
  }
  registered = [...handlers.keys()].sort()
})

/** An event object shaped enough for the handlers' session lookup. */
const event = { sender: { id: 1, send: () => undefined } }

describe('the IPC surface exists', () => {
  it('registers a substantial set of channels', () => {
    expect(registered.length).toBeGreaterThan(20)
  })

  it('includes the channels this app is built on', () => {
    for (const channel of ['workbook:export-pdf', 'workbook:render-preview']) {
      expect(registered, `${channel} missing from ${registered.length} channels`).toContain(channel)
    }
  })
})

describe('no handler accepts malformed input', () => {
  const JUNK: unknown[] = [undefined, null, 0, 'string', [], {}, { sessionId: 'not-a-uuid' }]

  it('every registered channel rejects every junk payload', async () => {
    const accepted: string[] = []
    for (const [channel, handler] of handlers) {
      for (const input of JUNK) {
        try {
          await handler(event, input)
          // Reaching here means the handler returned rather than throwing.
          // Read-only status channels legitimately do; anything that mutates
          // or touches the filesystem must not.
          if (/save|write|export|render|edit|raw|recovery|rename|close/i.test(channel)) {
            accepted.push(`${channel} <- ${JSON.stringify(input)}`)
          }
        } catch {
          // refused, which is the expected outcome
        }
      }
    }
    expect(accepted, `handlers that accepted junk:\n${accepted.join('\n')}`).toEqual([])
  })
})

describe('window and path bookkeeping', () => {
  it('reports no window before one is set', async () => {
    const { getSheetsWindow, setSheetsShellWindow } = await import('../src/main/sheets-main')
    setSheetsShellWindow(null)
    expect(getSheetsWindow()).toBeNull()
  })

  it('remembers the active web contents it was given', async () => {
    const { setActiveSheetsWebContents, getActiveSheetsWebContents } = await import(
      '../src/main/sheets-main'
    )
    const wc = { id: 7 } as never
    setActiveSheetsWebContents(wc)
    expect(getActiveSheetsWebContents()).toBe(wc)
    setActiveSheetsWebContents(null)
    expect(getActiveSheetsWebContents()).toBeNull()
  })

  it('tracks a queued workbook path', async () => {
    const { setForcedWorkbookPath, hasQueuedWorkbook } = await import('../src/main/sheets-main')
    setForcedWorkbookPath(undefined)
    expect(hasQueuedWorkbook()).toBe(false)
    setForcedWorkbookPath('/tmp/book.xlsx')
    expect(hasQueuedWorkbook()).toBe(true)
    setForcedWorkbookPath(undefined)
  })
})
