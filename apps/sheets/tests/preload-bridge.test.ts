/**
 * The preload bridge — case E9 in docs/testing/e2e-test-cases.md.
 *
 * This is the only validation between an untrusted renderer and a main process
 * that touches the filesystem, spawns a sidecar and holds API keys. It had no
 * tests at all, which matters more than the number suggests: it validates by
 * allow-list, so a field it does not name is dropped in silence with no error
 * anywhere. That is exactly how the oneCellAnchor extent went missing — the
 * parser emitted it, the schema accepted it, and this layer quietly deleted it
 * on the way past.
 *
 * So these tests assert two things per method: that malformed input is refused
 * rather than forwarded, and that valid input arrives at the channel intact.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest'

/// The boundary validates session ids as UUIDs, so the fixtures must be real ones.
const SESSION = '3f2504e0-4f89-11d3-9a0c-0305e82c3301'

const invoke = vi.fn()
const exposed = new Map<string, Record<string, (...args: never[]) => unknown>>()

vi.mock('electron', () => ({
  contextBridge: {
    exposeInMainWorld: (key: string, value: Record<string, (...a: never[]) => unknown>) => {
      exposed.set(key, value)
    },
  },
  ipcRenderer: {
    invoke: (...args: unknown[]) => invoke(...args),
    on: () => undefined,
    off: () => undefined,
    send: () => undefined,
  },
  webUtils: { getPathForFile: () => '/tmp/x' },
}))

await import('../src/preload/index')

/** One bridge method by name; throws if the bridge never exposed it. */
function api(name: string): (...args: never[]) => Promise<unknown> {
  const bridge = exposed.get('desktopApi') as
    | Record<string, ((...args: never[]) => Promise<unknown>) | undefined>
    | undefined
  const method = bridge?.[name]
  if (!method) throw new Error(`desktopApi has no method ${name}`)
  return method
}

/** The channel and payload the bridge forwarded, or null when it refused. */
function forwarded(): { channel: string; payload: unknown } | null {
  const call = invoke.mock.calls.at(-1)
  return call ? { channel: call[0] as string, payload: call[1] } : null
}

beforeEach(() => {
  invoke.mockReset()
})

describe('preload bridge: it exists and is shaped as expected', () => {
  it('exposes desktopApi with the methods the renderer calls', () => {
    expect(exposed.get('desktopApi')).toBeTruthy()
    for (const name of [
      'readWorkbookRange',
      'saveWorkbookEdits',
      'exportPdf',
      'renderPreview',
      'readRawPart',
      'editRawPart',
      'readWorkbookMedia',
    ]) {
      expect(typeof api(name)).toBe('function')
    }
  })
})

describe('renderPreview', () => {
  const valid = { html: '<html></html>', width: 1000, maxHeight: 4000 }

  it('forwards a valid request unchanged', async () => {
    invoke.mockResolvedValue({ base64: 'AAA', width: 100, height: 50, truncated: false })
    await api('renderPreview')(valid as never)
    expect(forwarded()).toEqual({ channel: 'workbook:render-preview', payload: valid })
  })

  it.each([
    ['empty html', { ...valid, html: '' }],
    ['non-string html', { ...valid, html: 42 }],
    ['width below the floor', { ...valid, width: 10 }],
    ['width above the ceiling', { ...valid, width: 99_999 }],
    ['fractional width', { ...valid, width: 800.5 }],
    ['maxHeight above the ceiling', { ...valid, maxHeight: 100_000 }],
    ['not an object', 'nope'],
  ])('refuses %s without forwarding', async (_label, request) => {
    await expect(api('renderPreview')(request as never)).rejects.toThrow()
    expect(invoke).not.toHaveBeenCalled()
  })

  it('refuses a malformed response rather than handing it to the renderer', async () => {
    invoke.mockResolvedValue({ base64: 'AAA', width: 'wide', height: 50, truncated: false })
    await expect(api('renderPreview')(valid as never)).rejects.toThrow(/preview render response/i)
  })

  it('refuses a response with an empty image', async () => {
    invoke.mockResolvedValue({ base64: '', width: 10, height: 10, truncated: false })
    await expect(api('renderPreview')(valid as never)).rejects.toThrow()
  })
})

describe('exportPdf', () => {
  const valid = {
    fileName: 'book.pdf',
    html: '<html></html>',
    landscape: false,
    pageSize: 'A4',
    margins: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7 },
    scale: 1,
  }

  it('forwards a valid request', async () => {
    invoke.mockResolvedValue({ canceled: true })
    await api('exportPdf')(valid as never)
    expect(forwarded()?.channel).toBe('workbook:export-pdf')
  })

  it.each([
    ['scale above the ceiling', { ...valid, scale: 5 }],
    ['scale below the floor', { ...valid, scale: 0 }],
    ['a margin beyond 3in', { ...valid, margins: { ...valid.margins, top: 9 } }],
    ['an unknown page size', { ...valid, pageSize: 'A0' }],
    ['a non-boolean landscape', { ...valid, landscape: 'yes' }],
    ['an empty file name', { ...valid, fileName: '' }],
  ])('refuses %s', async (_label, request) => {
    await expect(api('exportPdf')(request as never)).rejects.toThrow()
    expect(invoke).not.toHaveBeenCalled()
  })

  it('accepts a custom page size in inches', async () => {
    invoke.mockResolvedValue({ canceled: true })
    await api('exportPdf')({ ...valid, pageSize: { width: 7.25, height: 10.5 } } as never)
    expect(forwarded()?.channel).toBe('workbook:export-pdf')
  })
})

describe('readWorkbookRange', () => {
  it('forwards a valid request', async () => {
    invoke.mockResolvedValue({
      cells: [],
      rows: [],
      merges: [],
      hyperlinks: [],
      conditionalRules: [],
      dataValidations: [],
      autoFilter: null,
      indexedThroughRow: null,
      indexingComplete: true,
    })
    await api('readWorkbookRange')({
      sessionId: SESSION,
      sheetId: 'sheet-1',
      range: { startRow: 0, endRow: 9, startColumn: 0, endColumn: 4 },
    } as never)
    expect(forwarded()?.channel).toContain('workbook')
  })

  it('refuses a response missing its arrays rather than handing it on', async () => {
    invoke.mockResolvedValue({ cells: [], rows: [] })
    await expect(
      api('readWorkbookRange')({
        sessionId: SESSION,
        sheetId: 'sheet-1',
        range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
      } as never),
    ).rejects.toThrow(/range response/i)
  })

  it.each([
    ['a missing session', { sheetId: 'a', range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 } }],
    ['a missing range', { sessionId: SESSION, sheetId: 'a' }],
    ['a non-uuid session', { sessionId: 'nope', sheetId: 'a', range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 } }],
  ])('refuses %s', async (_label, request) => {
    await expect(api('readWorkbookRange')(request as never)).rejects.toThrow()
    expect(invoke).not.toHaveBeenCalled()
  })
})

describe('raw OOXML bridge', () => {
  it('forwards a well-formed raw edit', async () => {
    invoke.mockResolvedValue({ ok: true })
    await api('editRawPart')({
      sessionId: SESSION,
      ref: 'xl/worksheets/sheet1.xml',
      find: '<a/>',
      replace: '<b/>',
    } as never)
    expect(forwarded()?.channel).toContain('raw')
  })

  it.each([
    ['a missing session', { ref: 'x.xml', find: 'a', replace: 'b' }],
    ['a non-uuid session', { sessionId: 'session-1', ref: 'x.xml', find: 'a', replace: 'b' }],
    ['an empty part ref', { sessionId: SESSION, ref: '', find: 'a', replace: 'b' }],
    ['a non-string replace', { sessionId: SESSION, ref: 'x.xml', find: 'a', replace: 7 }],
  ])('refuses %s', async (_label, request) => {
    await expect(api('editRawPart')(request as never)).rejects.toThrow()
    expect(invoke).not.toHaveBeenCalled()
  })
})

describe('readWorkbookMedia', () => {
  it('refuses a response that is not a media payload', async () => {
    invoke.mockResolvedValue({ mediaType: 'image/png' })
    await expect(
      api('readWorkbookMedia')({ sessionId: SESSION, visualId: 'v1' } as never),
    ).rejects.toThrow()
  })

  it('passes a valid media payload through', async () => {
    invoke.mockResolvedValue({ mediaType: 'image/png', base64: 'AAAA' })
    const result = await api('readWorkbookMedia')({ sessionId: SESSION, visualId: 'v1' } as never)
    expect(result).toMatchObject({ mediaType: 'image/png', base64: 'AAAA' })
  })
})

/**
 * The rest of the boundary. These are grouped by what an attacker or a bug in
 * the renderer could reach: the shell, the filesystem, and the key store.
 */
describe('openExternal is the shell, so the scheme allow-list is the whole defence', () => {
  it.each([
    'file:///etc/passwd',
    'javascript:alert(1)',
    'data:text/html,<script>x</script>',
    'vscode://x',
    'HTTP-not-a-url',
    '',
  ])('refuses %s', async (url) => {
    await expect(api('openExternal')(url as never)).rejects.toThrow(/http/i)
    expect(invoke).not.toHaveBeenCalled()
  })

  it.each(['https://example.com/a', 'http://localhost:5173/x'])('allows %s', async (url) => {
    invoke.mockResolvedValue(undefined)
    await api('openExternal')(url as never)
    expect(forwarded()?.payload).toBe(url)
  })
})

describe('readLocalImage', () => {
  it('refuses a request with no path', async () => {
    await expect(api('readLocalImage')({} as never)).rejects.toThrow(/local image/i)
    expect(invoke).not.toHaveBeenCalled()
  })

  it('refuses a response that is not an image payload', async () => {
    invoke.mockResolvedValue({ mediaType: 'image/png' })
    await expect(api('readLocalImage')({ path: '/tmp/a.png' } as never)).rejects.toThrow()
  })
})

describe('attachment paths', () => {
  it('refuses a non-array', async () => {
    await expect(api('addAttachmentPaths')('/tmp/a.txt' as never)).rejects.toThrow(/attachment/i)
    expect(invoke).not.toHaveBeenCalled()
  })

  it.each([
    ['an empty path', ['']],
    ['a non-string entry', [42]],
    ['an over-long path', ['/'.padEnd(1100, 'a')]],
  ])('refuses %s', async (_label, paths) => {
    await expect(api('addAttachmentPaths')(paths as never)).rejects.toThrow()
    expect(invoke).not.toHaveBeenCalled()
  })
})

describe('autoRenameWorkbook', () => {
  it('refuses a non-uuid session', async () => {
    await expect(api('autoRenameWorkbook')('session-1' as never, 'Book1' as never)).rejects.toThrow(
      /session/i,
    )
    expect(invoke).not.toHaveBeenCalled()
  })

  it.each([
    ['an empty name', ''],
    ['an over-long name', 'x'.repeat(101)],
  ])('refuses %s', async (_label, name) => {
    await expect(api('autoRenameWorkbook')(SESSION as never, name as never)).rejects.toThrow(/name/i)
    expect(invoke).not.toHaveBeenCalled()
  })
})

describe('writeWorkbookRecovery takes a full save request, not a name', () => {
  it('refuses anything that is not one', async () => {
    await expect(api('writeWorkbookRecovery')({ sessionId: SESSION } as never)).rejects.toThrow(
      /save request/i,
    )
    expect(invoke).not.toHaveBeenCalled()
  })
})

describe('AI bridge responses are validated too', () => {
  it('refuses a provider-test response that is not a verdict', async () => {
    invoke.mockResolvedValue({ message: 'fine' })
    await expect(api('aiTestProvider')({ provider: 'openai' } as never)).rejects.toThrow(
      /provider test/i,
    )
  })

  it('refuses a genspark status response with no loggedIn flag', async () => {
    invoke.mockResolvedValue({ email: 'x@y.z' })
    await expect(api('aiGskStatus')(true as never)).rejects.toThrow(/account status/i)
  })

  it('refuses a stream cancel with no request id', async () => {
    await expect(api('aiStreamCancel')('' as never)).rejects.toThrow(/stream request id/i)
    expect(invoke).not.toHaveBeenCalled()
  })

  it('falls back to "not installed" rather than trusting a junk CLI status', async () => {
    invoke.mockResolvedValue('nonsense')
    await expect(api('aiCliStatus')('claude' as never)).resolves.toMatchObject({ installed: false })
  })
})
