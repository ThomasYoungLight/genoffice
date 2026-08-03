import type { Editor } from '@tiptap/core'
import type {
  ChartDisplay,
  CommentInfo,
  HeaderFooter,
  NewChart,
  NoteInfo,
  SectionSettings,
  SourceInfo,
} from '@genoffice/docx-engine'
import type { AgentToolCall, AgentToolDef } from '../../shared/ipc'
import { t } from '../i18n/locale'
import { executeCommands, type Command, type CommandEnvelope } from './commands'
import {
  blockRangePositions,
  buildDocumentContext,
  insertBlocksAfter,
  parseHtmlFragment,
  replaceBlockRange,
  serializeRangeToHtml,
  type AiTrack,
  type NumIds,
} from './protocol'

/**
 * Local agent tools: a document-context reader plus a script-style execution
 * channel. The "execution backend" is the in-process ProseMirror doc, split
 * into three safe primitives plus the deterministic command engine.
 */

const READ_MAX_CHARS = 24_000

export const AGENT_TOOLS: AgentToolDef[] = [
  {
    name: 'get_document_context',
    description:
      'Get the latest state of the current document: block list (index|type|content preview), full-text stats (word/character counts) and the current selection. Block indexes change after modifications; call this when you need up-to-date indexes.',
    inputSchema: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'read_blocks',
    description:
      'Read the full content of a block range (restricted HTML). Previews in the block list are truncated; you must read the full original text with this tool before rewriting.',
    inputSchema: {
      type: 'object',
      properties: {
        startBlockIndex: { type: 'integer', description: 'start block index (0-based, inclusive)' },
        endBlockIndex: { type: 'integer', description: 'end block index (inclusive)' },
      },
      required: ['startBlockIndex', 'endBlockIndex'],
    },
  },
  {
    name: 'insert_content',
    description:
      'Insert new content at a given position (restricted HTML, may contain multiple blocks). For writing/continuing/generating new content; to rewrite existing content use replace_blocks.',
    inputSchema: {
      type: 'object',
      properties: {
        html: { type: 'string', description: 'restricted HTML fragment to insert' },
        afterBlockIndex: {
          type: 'integer',
          description:
            'insert after this block index; -1 = start of document; omitted = after the block containing the cursor',
        },
      },
      required: ['html'],
    },
  },
  {
    name: 'replace_blocks',
    description:
      'Replace a block range with new content (restricted HTML). For rewriting/translating/condensing/expanding existing content; the new block count may differ from the old.',
    inputSchema: {
      type: 'object',
      properties: {
        startBlockIndex: { type: 'integer', description: 'start block index (0-based, inclusive)' },
        endBlockIndex: { type: 'integer', description: 'end block index (inclusive)' },
        html: { type: 'string', description: 'replacement restricted HTML fragment' },
      },
      required: ['startBlockIndex', 'endBlockIndex', 'html'],
    },
  },
  {
    name: 'apply_commands',
    description:
      'Execute formatting/structure/batch commands (batchUpdate style, see the command guide in the system prompt): text style, paragraph format, heading level, find & replace, delete/move blocks, list conversion, image properties.',
    inputSchema: {
      type: 'object',
      properties: {
        commands: {
          type: 'array',
          description: 'array of commands executed in order; each command is a single-key object',
          items: { type: 'object' },
        },
      },
      required: ['commands'],
    },
  },
  {
    name: 'set_page_setup',
    description:
      'Set the page itself: paper size, orientation, margins and text columns. This is the tool for "make it landscape", "A4", "narrower margins", "two columns" — none of which can be done by editing content. Margins are in centimetres. Omit any field to leave it as it is; call with no fields to read the current setup.',
    inputSchema: {
      type: 'object',
      properties: {
        pageSize: {
          type: 'string',
          enum: ['a4', 'a3', 'a5', 'letter', 'legal', 'tabloid'],
          description: 'Named paper size; the orientation is applied on top of it',
        },
        orientation: { type: 'string', enum: ['portrait', 'landscape'] },
        marginTopCm: { type: 'number' },
        marginRightCm: { type: 'number' },
        marginBottomCm: { type: 'number' },
        marginLeftCm: { type: 'number' },
        columns: { type: 'integer', description: 'Number of text columns; 1 is normal' },
      },
      required: [],
    },
  },
  {
    name: 'set_header_footer',
    description:
      'Set the running header or footer — the line repeated on every page. text is the content; pageNumber adds an automatic page number (footer only), which is what "add page numbers" means. Call with only kind to read the current value. A header is not a heading at the top of the first page: use insert_content for that.',
    inputSchema: {
      type: 'object',
      properties: {
        kind: { type: 'string', enum: ['header', 'footer'] },
        text: { type: 'string', description: 'Line of text; empty string clears it' },
        pageNumber: { type: 'boolean', description: 'Append an automatic page number (footer)' },
      },
      required: ['kind'],
    },
  },
  {
    name: 'insert_page_break',
    description:
      'Start a new page at a given point. Use it to keep a section starting at the top of a page; do not fake it with empty paragraphs, which move as soon as anything above them changes.',
    inputSchema: {
      type: 'object',
      properties: {
        afterBlockIndex: {
          type: 'integer',
          description: 'Break after this block index; -1 = start of document',
        },
      },
      required: ['afterBlockIndex'],
    },
  },
  {
    name: 'manage_comments',
    description:
      'Review comments anchored to the text — the margin notes a reviewer leaves, not footnotes (those are manage_notes). action=list returns every comment with its id, author and whether it is resolved; add attaches a new one to a block range; reply threads under an existing comment; resolve marks one done or reopens it; delete removes it. Use add when the user asks you to "comment on" or "flag" something rather than change it.',
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['list', 'add', 'reply', 'resolve', 'delete'] },
        text: { type: 'string', description: 'Comment body, for add and reply' },
        startBlockIndex: { type: 'integer', description: 'Block range to anchor to, for add' },
        endBlockIndex: { type: 'integer', description: 'Defaults to startBlockIndex' },
        id: { type: 'string', description: 'Comment id from list, for reply/resolve/delete' },
        done: { type: 'boolean', description: 'resolve: true marks done, false reopens' },
      },
      required: ['action'],
    },
  },
  {
    name: 'insert_shape',
    description:
      'Insert a floating text box or a preset shape at the cursor — a callout, a highlighted aside, an arrow. These float over the page rather than sitting in the text flow, so use them for annotation and emphasis, not for body content, which belongs in insert_content.',
    inputSchema: {
      type: 'object',
      properties: {
        kind: {
          type: 'string',
          enum: ['textbox', 'shape'],
          description: 'A text box, or a preset geometric shape',
        },
        preset: {
          type: 'string',
          description:
            "Shape geometry when kind is 'shape': rect, roundRect, ellipse, triangle, diamond, star5, rightArrow, and the other OOXML presets. Ignored for a text box.",
        },
      },
      required: ['kind'],
    },
  },
  {
    name: 'web_search',
    description:
      'Search the web for textual information (references/data/facts). Use when you need up-to-date information or are unsure about a fact. Returns titles/links/snippets.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'search keywords' },
        maxResults: { type: 'integer', description: 'maximum number of results, default 6' },
      },
      required: ['query'],
    },
  },
  {
    name: 'image_search',
    description:
      'Search for images. Returns a list of image imageUrl entries; after picking one, insert it into the document with insert_image.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'image search keywords (English works better)' },
        maxResults: { type: 'integer', description: 'maximum number of results, default 8' },
      },
      required: ['query'],
    },
  },
  {
    name: 'insert_image',
    description:
      'Download a direct image link (an imageUrl from image_search) and insert it into the document (at the cursor / end of document).',
    inputSchema: {
      type: 'object',
      properties: {
        url: { type: 'string', description: 'direct image link' },
        maxWidthPx: { type: 'integer', description: 'maximum width (px), default 480' },
      },
      required: ['url'],
    },
  },
  {
    name: 'generate_image',
    description:
      'Generate an illustration from a description and insert it into the document (at the cursor / end of document). Use for diagrams, covers and illustrations that no web image search would find; for a photo of something real, prefer image_search + insert_image.',
    inputSchema: {
      type: 'object',
      properties: {
        prompt: {
          type: 'string',
          description:
            'what the image should show, in detail: subject, composition, style, colours. English works better.',
        },
        size: {
          type: 'string',
          enum: ['1024x1024', '1536x1024', '1024x1536'],
          description: 'square, landscape or portrait; default square',
        },
        transparent: {
          type: 'boolean',
          description: 'cut out the background (for logos and icons placed over text)',
        },
        maxWidthPx: {
          type: 'integer',
          description: 'maximum width in the document (px), default 480',
        },
      },
      required: ['prompt'],
    },
  },
  {
    name: 'insert_chart',
    description:
      'Insert a chart (saved as a native Word chart). Data must be real: from the document content or web_search results — do not make up numbers.',
    inputSchema: {
      type: 'object',
      properties: {
        kind: { type: 'string', enum: ['bar', 'line', 'pie'], description: 'chart type' },
        title: { type: 'string', description: 'chart title' },
        categories: {
          type: 'array',
          items: { type: 'string' },
          description: 'category (x axis / sector) labels',
        },
        series: {
          type: 'array',
          description:
            'data series; values has the same length as categories, use null for missing data. Pie charts use only the first series',
          items: {
            type: 'object',
            properties: {
              name: { type: 'string' },
              values: { type: 'array', items: { type: ['number', 'null'] } },
            },
            required: ['values'],
          },
        },
        afterBlockIndex: {
          type: 'integer',
          description:
            'insert after this block index; -1 = start of document; omitted = after the block containing the cursor',
        },
      },
      required: ['kind', 'categories', 'series'],
    },
  },
  {
    name: 'edit_chart',
    description:
      'Edit the data of an existing chart in the document (title/category labels/series names/values). Chart blocks in the block list can be edited; ' +
      'the category count and the number of values per series must match the original chart (data points cannot be added or removed).',
    inputSchema: {
      type: 'object',
      properties: {
        blockIndex: { type: 'integer', description: 'block index of the chart' },
        title: { type: 'string', description: 'new title (omit to keep)' },
        categories: {
          type: 'array',
          items: { type: ['string', 'null'] },
          description:
            'new category labels, same length as the original; pass null for positions to keep',
        },
        series: {
          type: 'array',
          description: 'series to change',
          items: {
            type: 'object',
            properties: {
              index: { type: 'integer', description: 'series index (0-based)' },
              name: { type: 'string', description: 'new series name (omit to keep)' },
              values: {
                type: 'array',
                items: { type: ['number', 'null'] },
                description:
                  'new values, same length as the original series; pass null for positions to keep',
              },
            },
            required: ['index'],
          },
        },
      },
      required: ['blockIndex'],
    },
  },
  {
    name: 'manage_notes',
    description:
      'Footnotes and endnotes. action=list returns the existing ones with their ids; add inserts a reference marker at the cursor and creates the note; edit and delete take an id from list (deleting renumbers the rest). Use a footnote for an aside or a source on the page it belongs to, an endnote when the document collects them at the end — not for content that belongs in the body.',
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['list', 'add', 'edit', 'delete'] },
        kind: { type: 'string', enum: ['footnote', 'endnote'], description: 'default footnote' },
        text: { type: 'string', description: 'add / edit: the note text' },
        id: { type: 'string', description: 'edit / delete: id from action=list' },
      },
      required: ['action'],
    },
  },
  {
    name: 'set_watermark',
    description:
      'Set or remove the page watermark — the diagonal text behind the body, for DRAFT / CONFIDENTIAL and the like. Pass an empty string to remove it. One or two words; a watermark is a status marker, not a notice.',
    inputSchema: {
      type: 'object',
      properties: { text: { type: 'string', description: 'watermark text; empty removes it' } },
      required: ['text'],
    },
  },
  {
    name: 'manage_sources',
    description:
      'Bibliography sources. action=list returns them with their tags; add registers one. A source is only worth adding when the document cites it — pair it with a citation in the text. tag is the short key you cite by (e.g. "Wang2024"); type is Word\'s source type: Book, JournalArticle, InternetSite, Report, ConferenceProceedings.',
    inputSchema: {
      type: 'object',
      properties: {
        action: { type: 'string', enum: ['list', 'add', 'remove'] },
        tag: { type: 'string' },
        type: { type: 'string' },
        author: { type: 'string' },
        title: { type: 'string' },
        year: { type: 'string' },
        publisher: { type: 'string' },
        url: { type: 'string' },
      },
      required: ['action'],
    },
  },
]

export interface ToolExecution {
  /** result text fed back to the model */
  output: string
  isError?: boolean
  /** true when the tool changed the document */
  mutated: boolean
  /** short human-readable label for the chat activity chip */
  summary: string
}

const fail = (summary: string, output: string): ToolExecution => ({
  output,
  isError: true,
  mutated: false,
  summary,
})

function clampRange(
  editor: Editor,
  start: unknown,
  end: unknown,
): { start: number; end: number } | null {
  const count = editor.state.doc.childCount
  if (!Number.isInteger(start) || !Number.isInteger(end)) return null
  const s = Number(start)
  const e = Number(end)
  if (s < 0 || e < s || s >= count) return null
  return { start: s, end: Math.min(e, count - 1) }
}

/** Read the natural size of a dataURL image. */
function imageSizeOf(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => resolve({ width: img.naturalWidth || 1, height: img.naturalHeight || 1 })
    img.onerror = () => reject(new Error('image load failed'))
    img.src = dataUrl
  })
}

/** Async tools: web search / image search / insert web image. */
async function executeAsyncTool(editor: Editor, call: AgentToolCall): Promise<ToolExecution> {
  switch (call.name) {
    case 'web_search': {
      const query = String(call.input.query ?? '').trim()
      if (!query) return fail(t('aiSumWebSearch'), 'query must not be empty')
      const r = await window.desktop.webSearch(query, Number(call.input.maxResults) || 6)
      const lines: string[] = []
      if (r.answer) lines.push(`Direct answer: ${r.answer}\n`)
      r.results.forEach((it, i) =>
        lines.push(`${i + 1}. ${it.title}\n   ${it.url}\n   ${it.snippet}`),
      )
      return {
        output: lines.join('\n') || '(no results)',
        mutated: false,
        summary: t('aiSumWebSearchDone', { query, count: r.results.length }),
      }
    }
    case 'image_search': {
      const query = String(call.input.query ?? '').trim()
      if (!query) return fail(t('aiSumImageSearch'), 'query must not be empty')
      const r = await window.desktop.imageSearch(query, Number(call.input.maxResults) || 8)
      const lines = r.images.map(
        (im, i) =>
          `${i + 1}. ${im.title || '(untitled)'} [${im.width ?? '?'}x${im.height ?? '?'}]\n   ${im.imageUrl}`,
      )
      return {
        output: lines.join('\n') || '(no images)',
        mutated: false,
        summary: t('aiSumImageSearchDone', { query, count: r.images.length }),
      }
    }
    case 'generate_image': {
      const prompt = String(call.input.prompt ?? '').trim()
      if (!prompt) return fail(t('aiSumGenerateImage'), 'the prompt is empty')
      const generated = await window.desktop.generateImage({
        prompt,
        size: call.input.size ? String(call.input.size) : undefined,
        transparent: call.input.transparent === true,
      })
      if (!generated.ok || !generated.base64) {
        return fail(t('aiSumGenerateImage'), generated.error ?? 'generation failed')
      }
      return insertImageNode(
        editor,
        { base64: generated.base64, mime: generated.mime ?? 'image/png' },
        Number(call.input.maxWidthPx) || 480,
        t('aiSumGenerateImage'),
        'Image (generated)',
      )
    }
    case 'insert_image': {
      const url = String(call.input.url ?? '')
      if (!/^https?:\/\//.test(url)) return fail(t('aiSumInsertImage'), 'invalid url')
      const fetched = await window.desktop.fetchImage(url)
      if (!fetched)
        return fail(t('aiSumInsertImage'), 'download failed (the image may not be accessible)')
      return insertImageNode(
        editor,
        fetched,
        Number(call.input.maxWidthPx) || 480,
        t('aiSumInsertWebImage'),
        'Image (web)',
      )
    }
    default:
      return fail(t('aiSumUnknownTool'), call.name)
  }
}

/**
 * Decode, scale to fit, and insert as a protected image block. Shared by
 * insert_image (downloaded) and generate_image (produced by the provider) —
 * the bytes arrive the same way, only their origin differs.
 */
async function insertImageNode(
  editor: Editor,
  image: { base64: string; mime: string },
  maxWidthPx: number,
  summary: string,
  label: string,
): Promise<ToolExecution> {
  const dataUrl = `data:${image.mime};base64,${image.base64}`
  try {
    const natural = await imageSizeOf(dataUrl)
    const scale = Math.min(1, maxWidthPx / natural.width)
    const w = Math.round(natural.width * scale)
    const h = Math.round(natural.height * scale)
    editor
      .chain()
      .focus()
      .insertContent({
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'image',
          label,
          imageDataUrl: dataUrl,
          imageWidthPx: w,
          imageHeightPx: h,
          genImage: { base64: image.base64, mime: image.mime, widthPx: w, heightPx: h },
        },
      })
      .run()
    return { output: `Inserted the image (${w}×${h}px).`, mutated: true, summary }
  } catch {
    return { output: 'the image could not be decoded', isError: true, mutated: false, summary }
  }
}

/**
 * The parts of a .docx that do not live in the editor's document tree.
 *
 * Footnotes, endnotes, the watermark and the bibliography sources sit in React
 * state beside the editor and are written into the package on save. The
 * References and Design ribbons have driven all four for a long time; the
 * agent could reach none of them, so a document it wrote could not carry a
 * footnote or a citation.
 *
 * Optional: a host that does not supply it has no tool for those parts, rather
 * than a tool that fails when called.
 */
export interface DocExtras {
  notes(kind: NoteKind): NoteInfo[]
  /** insert a note at the cursor, returning its new id */
  addNote(kind: NoteKind, text: string): string
  editNote(kind: NoteKind, id: string, text: string): boolean
  deleteNote(kind: NoteKind, id: string): boolean
  watermark(): string | null
  setWatermark(text: string | null): void
  sources(): SourceInfo[]
  setSources(list: SourceInfo[]): void
  /** the running header/footer of the current section */
  headerFooter(kind: 'header' | 'footer'): HeaderFooter | null
  setHeaderFooter(kind: 'header' | 'footer', value: HeaderFooter): void
  /** page size, orientation, margins and columns of the current section */
  pageSetup(): SectionSettings | null
  setPageSetup(next: SectionSettings): void
  comments(): CommentInfo[]
  /** comment on the editor's current selection, which the caller positions first */
  addComment(text: string): boolean
  replyToComment(parentId: string, text: string): void
  resolveComment(id: string, done: boolean): void
  deleteComment(id: string): void
  /** insert a floating text box or preset shape at the cursor */
  insertTextbox(): void
  insertShape(preset: string): void
}

export type NoteKind = 'footnote' | 'endnote'

/** Twips per unit, for the page-setup tool's human-facing measurements. */
const TWIPS = { cm: 567, inch: 1440 } as const

/** Named page sizes, in twips, portrait. */
const PAGE_SIZES: Record<string, { w: number; h: number }> = {
  a4: { w: 11906, h: 16838 },
  a3: { w: 16838, h: 23811 },
  a5: { w: 8391, h: 11906 },
  letter: { w: 12240, h: 15840 },
  legal: { w: 12240, h: 20160 },
  tabloid: { w: 15840, h: 24480 },
}

export function executeTool(
  editor: Editor,
  call: AgentToolCall,
  numIds: NumIds,
  track?: AiTrack,
  extras?: DocExtras,
): ToolExecution | Promise<ToolExecution> {
  // async tools (search/image insertion) take a separate Promise branch; the other sync tools keep returning synchronously (doesn't break existing tests).
  if (
    call.name === 'web_search' ||
    call.name === 'image_search' ||
    call.name === 'insert_image' ||
    call.name === 'generate_image'
  ) {
    return executeAsyncTool(editor, call)
  }
  switch (call.name) {
    case 'get_document_context':
      return {
        output: buildDocumentContext(editor),
        mutated: false,
        summary: t('aiSumReadDocContext'),
      }

    case 'read_blocks': {
      const range = clampRange(editor, call.input.startBlockIndex, call.input.endBlockIndex)
      if (!range) return fail(t('aiSumReadBlocks'), 'block index invalid or out of range')
      const html = serializeRangeToHtml(editor, range.start, range.end)
      const clipped =
        html.length > READ_MAX_CHARS ? html.slice(0, READ_MAX_CHARS) + '\n…(truncated)' : html
      return {
        output: clipped || '(range is empty)',
        mutated: false,
        summary: t('aiSumReadBlocksRange', { start: range.start, end: range.end }),
      }
    }

    case 'insert_content': {
      const html = String(call.input.html ?? '')
      let nodes: ReturnType<typeof parseHtmlFragment>
      try {
        nodes = parseHtmlFragment(html, numIds)
      } catch (e) {
        return fail(t('aiSumInsertContent'), e instanceof Error ? e.message : String(e))
      }
      if (nodes.length === 0)
        return fail(t('aiSumInsertContent'), 'html did not parse into any content blocks')
      const count = editor.state.doc.childCount
      const docIsEmpty = editor.state.doc.textContent.trim() === ''
      if (docIsEmpty) {
        // the blank template's single empty paragraph gets replaced
        replaceBlockRange(editor, 0, count - 1, nodes, track)
        return {
          output: `Inserted ${nodes.length} block(s) (the document was empty). Block indexes have changed; use get_document_context if needed.`,
          mutated: true,
          summary: t('aiSumInsertedBlocks', { count: nodes.length }),
        }
      }
      const cursorScope = call.input.afterBlockIndex === undefined
      const after = cursorScope
        ? getCursorBlockIndex(editor)
        : Math.min(Math.max(Number(call.input.afterBlockIndex), -1), count - 1)
      if (!Number.isInteger(after)) return fail(t('aiSumInsertContent'), 'invalid afterBlockIndex')
      // -1 hits blockRangePositions' 0/0 default, i.e. insert at doc start
      insertBlocksAfter(editor, after, nodes, track)
      return {
        output: `Inserted ${nodes.length} block(s) after block ${after}. Subsequent block indexes have shifted; use get_document_context if needed.`,
        mutated: true,
        summary: t('aiSumInsertedBlocks', { count: nodes.length }),
      }
    }

    case 'replace_blocks': {
      const range = clampRange(editor, call.input.startBlockIndex, call.input.endBlockIndex)
      if (!range) return fail(t('aiSumReplaceContent'), 'block index invalid or out of range')
      let nodes: ReturnType<typeof parseHtmlFragment>
      try {
        nodes = parseHtmlFragment(String(call.input.html ?? ''), numIds)
      } catch (e) {
        return fail(t('aiSumReplaceContent'), e instanceof Error ? e.message : String(e))
      }
      if (nodes.length === 0)
        return fail(t('aiSumReplaceContent'), 'html did not parse into any content blocks')
      replaceBlockRange(editor, range.start, range.end, nodes, track)
      return {
        output: `Replaced blocks ${range.start}-${range.end} with ${nodes.length} block(s). Block indexes have changed; use get_document_context if needed.`,
        mutated: true,
        summary: t('aiSumReplacedBlocks', { start: range.start, end: range.end }),
      }
    }

    case 'insert_chart': {
      const kind = String(call.input.kind ?? '') as NewChart['kind']
      if (!['bar', 'line', 'pie'].includes(kind))
        return fail(t('aiSumInsertChart'), 'kind must be one of bar/line/pie')
      const categories = Array.isArray(call.input.categories)
        ? (call.input.categories as unknown[]).map((c) => String(c ?? ''))
        : []
      const seriesIn = Array.isArray(call.input.series)
        ? (call.input.series as ReadonlyArray<{ name?: unknown; values?: unknown } | null>)
        : []
      if (!categories.length || !seriesIn.length)
        return fail(t('aiSumInsertChart'), 'categories and series must not be empty')
      const series = seriesIn.map((s, i) => {
        const values: unknown[] = Array.isArray(s?.values) ? s.values : []
        return {
          name: String(s?.name ?? `Series ${i + 1}`),
          values: categories.map((_, j) => {
            const v = values[j] ?? null
            const n = Number(v)
            return v != null && Number.isFinite(n) ? n : null
          }),
        }
      })
      const title = String(call.input.title ?? '').trim() || 'Chart title'
      const spec: NewChart = { kind, title, categories, series }
      const display: ChartDisplay = { partPath: '', kind, title, categories, series }
      const count = editor.state.doc.childCount
      const after =
        call.input.afterBlockIndex === undefined
          ? getCursorBlockIndex(editor)
          : Math.min(Math.max(Number(call.input.afterBlockIndex), -1), count - 1)
      if (!Number.isInteger(after)) return fail(t('aiSumInsertChart'), 'invalid afterBlockIndex')
      const { to } = blockRangePositions(editor, after, after)
      editor
        .chain()
        .insertContentAt(to, {
          type: 'docProtected',
          attrs: {
            docxIndex: null,
            blockType: 'chart',
            label: 'Chart',
            genChart: spec,
            chartDisplay: display,
          },
        })
        .run()
      return {
        output: `Inserted a ${kind} chart "${title}" (${categories.length} categories × ${series.length} series).`,
        mutated: true,
        summary: t('aiSumInsertedChart', { title }),
      }
    }

    case 'edit_chart': {
      const idx = Number(call.input.blockIndex)
      if (!Number.isInteger(idx) || idx < 0 || idx >= editor.state.doc.childCount) {
        return fail(t('aiSumEditChart'), 'blockIndex invalid or out of range')
      }
      const node = editor.state.doc.child(idx)
      // native docx charts are passthrough blocks + chartDisplay; AI/UI-created ones are blockType 'chart'
      const display =
        node.type.name === 'docProtected' ? (node.attrs.chartDisplay as ChartDisplay | null) : null
      if (!display)
        return fail(
          t('aiSumEditChart'),
          `block ${idx} is not a chart (or the chart has no editable data cache)`,
        )
      const isNative = display.partPath !== '' // native chart part from docx: data-point structure is immutable
      const next: ChartDisplay = {
        ...display,
        categories: [...display.categories],
        series: display.series.map((s) => ({ ...s, values: [...s.values] })),
      }
      if (call.input.title !== undefined) next.title = String(call.input.title)
      if (call.input.categories !== undefined) {
        const cats = call.input.categories
        if (!Array.isArray(cats) || cats.length !== display.categories.length) {
          return fail(
            t('aiSumEditChart'),
            `categories must match the original category count (${display.categories.length})`,
          )
        }
        cats.forEach((c, i) => {
          if (c != null) next.categories[i] = String(c)
        })
      }
      const serIn = Array.isArray(call.input.series)
        ? (call.input.series as ReadonlyArray<{
            index?: unknown
            name?: unknown
            values?: unknown
          } | null>)
        : []
      for (const s of serIn) {
        const si = Number(s?.index)
        const orig = display.series[si]
        if (!Number.isInteger(si) || !orig)
          return fail(
            t('aiSumEditChart'),
            `series index ${s?.index} is invalid (${display.series.length} series in total)`,
          )
        if (s?.name !== undefined) next.series[si]!.name = String(s.name)
        if (s?.values !== undefined) {
          const values: unknown[] | null = Array.isArray(s.values) ? s.values : null
          if (!values || values.length !== orig.values.length) {
            return fail(
              t('aiSumEditChart'),
              `values of series ${si} must match the original series length (${orig.values.length})`,
            )
          }
          for (let j = 0; j < values.length; j++) {
            const v = values[j]
            if (v == null) continue
            const n = Number(v)
            if (!Number.isFinite(n))
              return fail(t('aiSumEditChart'), `value ${j} of series ${si} is not a number`)
            if (isNative && orig.values[j] == null) {
              return fail(
                t('aiSumEditChart'),
                `data point ${j} of series ${si} is empty in the original chart; empty points of a native chart cannot be written`,
              )
            }
            next.series[si]!.values[j] = n
          }
        }
      }
      // generated charts (genChart) update the spec in sync; the chart part is rebuilt from the new data on save
      const gen = node.attrs.genChart as NewChart | null
      const nextGen: NewChart | null = gen
        ? {
            ...gen,
            title: next.title ?? gen.title,
            categories: [...next.categories],
            series: next.series.map((s) => ({ name: s.name ?? '', values: [...s.values] })),
          }
        : null
      const { from } = blockRangePositions(editor, idx, idx)
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(from, undefined, {
          ...node.attrs,
          chartDisplay: next,
          genChart: nextGen,
        }),
      )
      return {
        output: `Updated the data of chart "${next.title ?? ''}" (changes are written back to the chart on save).`,
        mutated: true,
        summary: t('aiSumEditedChart', { index: idx }),
      }
    }

    case 'apply_commands': {
      const commands = call.input.commands
      if (!Array.isArray(commands) || commands.length === 0) {
        return fail(t('aiSumApplyCommands'), 'commands must be a non-empty array')
      }
      const envelope: CommandEnvelope = { commands: commands as Command[] }
      const outcome = executeCommands(editor, envelope, { numIds, track })
      if (!outcome.ok)
        return fail(t('aiSumApplyCommands'), outcome.error ?? 'command execution failed')
      const changed = outcome.results.reduce((sum, r) => sum + r.changed, 0)
      return {
        output: outcome.summary,
        mutated: changed > 0,
        summary: outcome.summary,
      }
    }

    case 'manage_notes': {
      if (!extras) return fail(call.name, 'Notes are not available in this window')
      const kind: NoteKind = call.input.kind === 'endnote' ? 'endnote' : 'footnote'
      const action = String(call.input.action ?? '')
      const list = () => extras.notes(kind)
      const render = (notes: NoteInfo[]) =>
        notes.length
          ? notes.map((n, i) => `[${i + 1}] id=${n.id}: ${n.text}`).join('\n')
          : `The document has no ${kind}s.`
      if (action === 'list')
        return { output: render(list()), mutated: false, summary: t('aiSumNotes') }
      if (action === 'add') {
        const text = String(call.input.text ?? '').trim()
        if (!text) return fail(t('aiFailNotes'), 'A note needs text')
        const id = extras.addNote(kind, text)
        return {
          output: `Added ${kind} id=${id} at the cursor.\n${render(list())}`,
          mutated: true,
          summary: t('aiSumNotes'),
        }
      }
      const id = String(call.input.id ?? '')
      if (!id) return fail(t('aiFailNotes'), `${action} needs the note id from action=list`)
      if (action === 'edit') {
        const text = String(call.input.text ?? '').trim()
        if (!text) return fail(t('aiFailNotes'), 'A note needs text')
        if (!extras.editNote(kind, id, text)) return fail(t('aiFailNotes'), `No ${kind} id=${id}`)
        return { output: render(list()), mutated: true, summary: t('aiSumNotes') }
      }
      if (action === 'delete') {
        if (!extras.deleteNote(kind, id)) return fail(t('aiFailNotes'), `No ${kind} id=${id}`)
        return {
          output: `Deleted ${kind} id=${id}; the remaining markers were renumbered.\n${render(list())}`,
          mutated: true,
          summary: t('aiSumNotes'),
        }
      }
      return fail(t('aiFailNotes'), `Unknown action "${action}"`)
    }

    case 'set_page_setup': {
      if (!extras) return fail(call.name, 'Page setup is not available in this window')
      const current = extras.pageSetup()
      if (!current) return fail(call.name, 'This document has no section settings to change')
      const cm = (twips: number) => Math.round((twips / TWIPS.cm) * 100) / 100
      const describe = (s: SectionSettings) =>
        `${s.orientation}, ${cm(s.pageWidth)}\u00d7${cm(s.pageHeight)} cm, margins ` +
        `T${cm(s.marginTop)} R${cm(s.marginRight)} B${cm(s.marginBottom)} L${cm(s.marginLeft)} cm` +
        (s.columns > 1 ? `, ${s.columns} columns` : '')
      const fields = [
        'pageSize',
        'orientation',
        'marginTopCm',
        'marginRightCm',
        'marginBottomCm',
        'marginLeftCm',
        'columns',
      ]
      if (!fields.some((f) => f in call.input)) {
        return {
          output: `Current page setup: ${describe(current)}.`,
          mutated: false,
          summary: t('aiSumPageSetup'),
        }
      }
      const next: SectionSettings = { ...current }
      const sizeKey = String(call.input.pageSize ?? '')
      if (sizeKey) {
        const size = PAGE_SIZES[sizeKey]
        if (!size)
          return fail(
            call.name,
            `Unknown pageSize "${sizeKey}"; use one of ${Object.keys(PAGE_SIZES).join(', ')}`,
          )
        next.pageWidth = size.w
        next.pageHeight = size.h
        // a named size is portrait; the existing orientation still applies unless
        // the same call changes it, so swap here and let the check below re-swap
        next.orientation = 'portrait'
      }
      const orientation =
        String(call.input.orientation ?? '') || (sizeKey ? current.orientation : '')
      if (orientation === 'landscape' || orientation === 'portrait') {
        const long = Math.max(next.pageWidth, next.pageHeight)
        const short = Math.min(next.pageWidth, next.pageHeight)
        next.orientation = orientation
        next.pageWidth = orientation === 'landscape' ? long : short
        next.pageHeight = orientation === 'landscape' ? short : long
      }
      const margins: Array<[string, keyof SectionSettings]> = [
        ['marginTopCm', 'marginTop'],
        ['marginRightCm', 'marginRight'],
        ['marginBottomCm', 'marginBottom'],
        ['marginLeftCm', 'marginLeft'],
      ]
      for (const [input, field] of margins) {
        if (!(input in call.input)) continue
        const value = Number(call.input[input])
        if (!Number.isFinite(value) || value < 0 || value > 10)
          return fail(call.name, `${input} must be between 0 and 10 cm`)
        ;(next[field] as number) = Math.round(value * TWIPS.cm)
      }
      if ('columns' in call.input) {
        const columns = Number(call.input.columns)
        if (!Number.isInteger(columns) || columns < 1 || columns > 6)
          return fail(call.name, 'columns must be a whole number between 1 and 6')
        next.columns = columns
      }
      extras.setPageSetup(next)
      return {
        output: `Page setup: ${describe(next)}.`,
        mutated: true,
        summary: t('aiSumPageSetup'),
      }
    }

    case 'set_header_footer': {
      if (!extras) return fail(call.name, 'Headers and footers are not available in this window')
      const kind = String(call.input.kind ?? '')
      if (kind !== 'header' && kind !== 'footer')
        return fail(call.name, "kind must be 'header' or 'footer'")
      const current = extras.headerFooter(kind)
      if (!('text' in call.input) && !('pageNumber' in call.input)) {
        return {
          output: current?.text
            ? `Current ${kind}: "${current.text}"${current.pageNumber ? ' + page number' : ''}.`
            : `The document has no ${kind}.`,
          mutated: false,
          summary: t('aiSumHeaderFooter'),
        }
      }
      const pageNumber =
        'pageNumber' in call.input ? Boolean(call.input.pageNumber) : (current?.pageNumber ?? false)
      if (pageNumber && kind === 'header')
        return fail(call.name, 'An automatic page number belongs in the footer, not the header')
      const text = 'text' in call.input ? String(call.input.text ?? '') : (current?.text ?? '')
      extras.setHeaderFooter(kind, { text, ...(pageNumber ? { pageNumber: true } : {}) })
      return {
        output:
          text || pageNumber
            ? `Set the ${kind} to "${text}"${pageNumber ? ' with an automatic page number' : ''}.`
            : `Cleared the ${kind}.`,
        mutated: true,
        summary: t('aiSumHeaderFooter'),
      }
    }

    case 'insert_page_break': {
      // a page break is a property of the paragraph that starts the new page,
      // not a block of its own — so it is set on the block after the break and
      // moves with that block, which is the whole reason not to fake it with
      // empty paragraphs
      const after = Number(call.input.afterBlockIndex)
      const total = editor.state.doc.childCount
      if (!Number.isInteger(after) || after < -1 || after >= total)
        return fail(call.name, `afterBlockIndex out of range (-1 to ${total - 1})`)
      const target = after + 1
      if (target >= total)
        return fail(call.name, 'There is no block after that one to start a new page with')
      const { from } = blockRangePositions(editor, target, target)
      const node = editor.state.doc.nodeAt(from)
      if (!node) return fail(call.name, 'The block after the break could not be found')
      if (node.attrs.pageBreakBefore)
        return {
          output: `Block ${target} already starts a new page.`,
          mutated: false,
          summary: t('aiSumPageBreak'),
        }
      const tr = editor.state.tr.setNodeMarkup(from, undefined, {
        ...node.attrs,
        pageBreakBefore: true,
      })
      editor.view.dispatch(tr)
      return {
        output: `Block ${target} now starts a new page.`,
        mutated: true,
        summary: t('aiSumPageBreak'),
      }
    }

    case 'manage_comments': {
      if (!extras) return fail(call.name, 'Comments are not available in this window')
      const action = String(call.input.action ?? '')
      const list = extras.comments()
      const render = (items: CommentInfo[]) =>
        items.length
          ? items
              .map(
                (c) =>
                  `${c.id} | ${c.author}${c.done ? ' | resolved' : ''}${c.parentId ? ` | reply to ${c.parentId}` : ''} | ${c.text}`,
              )
              .join('\n')
          : 'The document has no comments.'
      if (action === 'list')
        return { output: render(list), mutated: false, summary: t('aiSumComments') }

      if (action === 'add') {
        const text = String(call.input.text ?? '').trim()
        if (!text) return fail(call.name, 'A comment needs text')
        const total = editor.state.doc.childCount
        const start = Number(call.input.startBlockIndex)
        if (!Number.isInteger(start) || start < 0 || start >= total)
          return fail(call.name, `startBlockIndex out of range (0-${total - 1})`)
        const end = Number(call.input.endBlockIndex ?? start)
        if (!Number.isInteger(end) || end < start || end >= total)
          return fail(call.name, `endBlockIndex out of range (${start}-${total - 1})`)
        // the comment attaches to whatever is selected, so the range has to be
        // the selection before the commit runs
        const { from, to } = blockRangePositions(editor, start, end)
        editor.commands.setTextSelection({ from, to })
        if (!extras.addComment(text)) return fail(call.name, 'The comment could not be anchored')
        return {
          output: `Commented on block${end > start ? `s ${start}-${end}` : ` ${start}`}.`,
          mutated: true,
          summary: t('aiSumComments'),
        }
      }

      const id = String(call.input.id ?? '')
      if (!id) return fail(call.name, `${action} needs a comment id from list`)
      if (!list.some((c) => c.id === id)) return fail(call.name, `No comment with id ${id}`)
      if (action === 'reply') {
        const text = String(call.input.text ?? '').trim()
        if (!text) return fail(call.name, 'A reply needs text')
        extras.replyToComment(id, text)
        return { output: `Replied to ${id}.`, mutated: true, summary: t('aiSumComments') }
      }
      if (action === 'resolve') {
        const done = call.input.done === undefined ? true : Boolean(call.input.done)
        extras.resolveComment(id, done)
        return {
          output: `${done ? 'Resolved' : 'Reopened'} ${id}.`,
          mutated: true,
          summary: t('aiSumComments'),
        }
      }
      if (action === 'delete') {
        extras.deleteComment(id)
        return { output: `Deleted ${id}.`, mutated: true, summary: t('aiSumComments') }
      }
      return fail(call.name, `Unknown action "${action}"`)
    }

    case 'insert_shape': {
      if (!extras) return fail(call.name, 'Shapes are not available in this window')
      const kind = String(call.input.kind ?? '')
      if (kind === 'textbox') {
        extras.insertTextbox()
        return { output: 'Inserted a text box.', mutated: true, summary: t('aiSumShape') }
      }
      if (kind !== 'shape') return fail(call.name, "kind must be 'textbox' or 'shape'")
      const preset = String(call.input.preset ?? 'rect').trim() || 'rect'
      if (!/^[a-zA-Z][a-zA-Z0-9]*$/.test(preset))
        return fail(call.name, `"${preset}" is not a preset geometry name`)
      extras.insertShape(preset)
      return { output: `Inserted a ${preset} shape.`, mutated: true, summary: t('aiSumShape') }
    }

    case 'set_watermark': {
      if (!extras) return fail(call.name, 'The watermark is not available in this window')
      const text = String(call.input.text ?? '').trim()
      extras.setWatermark(text || null)
      return {
        output: text ? `Watermark set to "${text}".` : 'Watermark removed.',
        mutated: true,
        summary: t('aiSumWatermark'),
      }
    }

    case 'manage_sources': {
      if (!extras) return fail(call.name, 'Sources are not available in this window')
      const action = String(call.input.action ?? '')
      const render = (list: SourceInfo[]) =>
        list.length
          ? list.map((s) => `${s.tag} — ${s.author}, ${s.title} (${s.year}) [${s.type}]`).join('\n')
          : 'The document has no bibliography sources.'
      if (action === 'list')
        return { output: render(extras.sources()), mutated: false, summary: t('aiSumSources') }
      if (action === 'add') {
        const tag = String(call.input.tag ?? '').trim()
        const title = String(call.input.title ?? '').trim()
        if (!tag || !title)
          return fail(t('aiFailSources'), 'A source needs at least a tag and a title')
        const current = extras.sources()
        if (current.some((s) => s.tag === tag))
          return fail(t('aiFailSources'), `A source tagged "${tag}" already exists`)
        const publisher = String(call.input.publisher ?? '').trim()
        const url = String(call.input.url ?? '').trim()
        extras.setSources([
          ...current,
          {
            tag,
            type: String(call.input.type ?? 'Book').trim() || 'Book',
            author: String(call.input.author ?? '').trim(),
            title,
            year: String(call.input.year ?? '').trim(),
            ...(publisher ? { publisher } : {}),
            ...(url ? { url } : {}),
          },
        ])
        return { output: render(extras.sources()), mutated: true, summary: t('aiSumSources') }
      }
      if (action === 'remove') {
        const tag = String(call.input.tag ?? '').trim()
        const current = extras.sources()
        if (!current.some((s) => s.tag === tag))
          return fail(t('aiFailSources'), `No source tagged "${tag}"`)
        extras.setSources(current.filter((s) => s.tag !== tag))
        return { output: render(extras.sources()), mutated: true, summary: t('aiSumSources') }
      }
      return fail(t('aiFailSources'), `Unknown action "${action}"`)
    }

    default:
      return fail(call.name, `unknown tool: ${call.name}`)
  }
}

/** top-level index of the block containing the caret (doc end as fallback) */
function getCursorBlockIndex(editor: Editor): number {
  const { from } = editor.state.selection
  let result = editor.state.doc.childCount - 1
  let index = 0
  editor.state.doc.forEach((node, offset) => {
    if (from >= offset && from <= offset + node.nodeSize) result = index
    index++
  })
  return result
}
