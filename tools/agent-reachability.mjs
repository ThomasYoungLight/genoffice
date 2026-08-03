#!/usr/bin/env node
/**
 * Which document capabilities can the AI actually reach?
 *
 * Every app here has more capability than its agent can call: the editor
 * drives an operation over IPC, the engine implements it, and no tool exposes
 * it. That gap is invisible in review — nothing is broken, a feature simply
 * never gets used — and it is how speaker notes, hyperlinks, animations,
 * footnotes and three conditional formats sat unreachable for as long as they
 * did. This prints the diff so it has to be a decision rather than an
 * oversight.
 *
 *   node tools/agent-reachability.mjs           # report
 *   node tools/agent-reachability.mjs --check   # non-zero exit if a NEW gap appeared
 *
 * The baseline below is the set of operations we have consciously decided the
 * agent does not need. Adding an IPC operation without either a tool or a
 * baseline entry fails --check.
 */
import { readFileSync, readdirSync, existsSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')

/**
 * Operations deliberately out of the agent's reach, with the reason. Most are
 * UI-only concerns (window management, clipboard, presenter mode) or things
 * that would be actively unsafe to automate.
 */
const BASELINE = {
  slides: {
    'ai-snapshot-restore': 'undo plumbing for the agent itself',
    'audience-ready': 'presenter UI state',
    'chart-color-schemes': 'palette list for the UI picker',
    'clipboard-external': 'OS clipboard',
    'consume-pending-open': 'app startup',
    'export-pdf': 'user action with a file dialog',
    'files-add': 'attachment intake, handled by the files skill',
    'files-pick': 'opens a native dialog',
    'files-read-image': 'attachment intake',
    'get-chart-data': 'read path used by edit_chart',
    'get-layouts': 'used internally by add_slide_with_layout',
    'get-render-slides': 'canvas state, already in the deck outline',
    'get-shape-keys': 'morph pairing internals',
    'get-slide-size': 'in the deck outline',
    'has-slide-clipboard': 'UI enablement',
    'history-batch-begin': 'undo grouping, applied around tool calls',
    'history-batch-end': 'undo grouping, applied around tool calls',
    'is-dirty': 'save state',
    'media-data': 'playback',
    'native-clipboard': 'OS clipboard',
    'new-blank': 'app-level',
    open: 'app-level',
    'open-path': 'app-level',
    'pick-export-dir': 'native dialog',
    'pick-export-pdf-path': 'native dialog',
    'presenter-end': 'presenter mode',
    'presenter-start': 'presenter mode',
    'presenter-swap': 'presenter mode',
    recent: 'app-level',
    redo: 'user action',
    save: 'user action',
    'save-as': 'native dialog',
    undo: 'user action',
    'copy-elements': 'clipboard; the agent duplicates instead',
    'paste-elements': 'clipboard',
    'copy-slide': 'clipboard',
    'paste-slide': 'clipboard',
    'repaste-slide': 'clipboard',
    'duplicate-elements': 'covered by add_* tools',
    'cloud-gen-status': 'capability probe used by generate_deck',
    'set-advance-times': 'rehearsal timing, recorded by presenting',
    'set-hidden': 'reachable via execute_slide_script',
    'set-slide-size': 'deck-level setup, not content',
    'batch-edit-transform': 'covered by execute_slide_script',
    'edit-transform': 'covered by set_element_transform',
    'edit-fill': 'covered by set_element_fill',
    'edit-stroke': 'covered by set_element_stroke',
    'edit-connector-endpoints': 'geometry, via execute_slide_script',
    'edit-picture-opacity': 'via execute_slide_script',
    'edit-picture-src-rect': 'cropping, a manual gesture',
    'add-ink': 'handwriting input',
    'add-image-bytes': 'used by insert_web_image / generate_image',
    'add-media-bytes': 'used by media insertion',
    'add-slide-with-layout': 'covered by add_slide',
    'apply-theme': 'deck-level setup',
    'find-replace': 'covered by execute_slide_script',
    'get-animations': 'read side of set_slide_animations',
    'get-comments': 'read side of manage_comments',
    'get-header-footer': 'read side of apply-header-footer',
    'get-link': 'read side of set_element_link',
    'get-notes': 'read side of set_slide_notes, surfaced by read_slide',
    'get-run-links': 'read side of set_element_link',
    'get-sections': 'read side of manage_sections',
    'get-slide-links': 'surfaced by read_slide',
    'get-transition': 'read side of set_slide_transition',
    'set-sections': 'covered by manage_sections',
    'master-close': 'master editing is a mode, not an operation',
    'master-delete-element': 'master editing mode',
    'master-edit-fill': 'master editing mode',
    'master-edit-stroke': 'master editing mode',
    'master-edit-text': 'master editing mode',
    'master-edit-transform': 'master editing mode',
    'master-enter': 'master editing mode',
    'master-open': 'master editing mode',
    'apply-header-footer': 'covered by set_header_footer',
    'insert-image': 'used by insert_web_image',
    'move-slide': 'covered by move_slide',
    'add-slide': 'covered by the add_slide tool',
    'add-blank-slide': 'used by local deck generation',
    'add-element': 'used by add_text_box / add_shape',
    'add-table': 'covered by add_table',
    'add-chart': 'covered by add_chart',
    'add-smartart': 'covered by add_smartart',
    'add-comment': 'covered by manage_comments',
    'add-section': 'covered by manage_sections',
    'delete-comment': 'covered by manage_comments',
    'delete-element': 'covered by delete_element',
    'delete-slide': 'covered by delete_slide',
    'edit-background': 'covered by set_slide_background',
    'edit-table-cell': 'covered by edit_table_cell',
    'edit-table-style': 'covered by edit_table_style',
    'edit-text': 'covered by set_element_text',
    'group-elements': 'covered by group_elements',
    'ungroup-element': 'covered by ungroup_element',
    'flip-elements': 'via execute_slide_script',
    'reorder-element': 'covered by reorder_element',
    'move-section': 'covered by manage_sections',
    'remove-section': 'covered by manage_sections',
    'rename-section': 'covered by manage_sections',
    'set-animations': 'covered by set_slide_animations',
    'set-element-font': 'covered by set_element_style',
    'set-element-paragraph-format': 'covered by set_element_text',
    'set-link': 'covered by set_element_link',
    'set-notes': 'covered by set_slide_notes',
    'set-slide-layout': 'deck-level setup',
    'set-table-cell-anchor': 'covered by edit_table_style',
    'set-table-col-width': 'covered by edit_table_structure',
    'set-table-row-height': 'covered by edit_table_structure',
    'set-transition': 'covered by set_slide_transition',
    'table-merge': 'covered by edit_table_structure',
    'table-structure': 'covered by edit_table_structure',
    'edit-chart': 'covered by edit_chart',
    'insert-model3d': 'opens a native file dialog; the agent has no model file to insert',
    'set-text-anchor': 'covered by set_text_anchor',
    'cloud-page-generate': 'the cloud page call behind generate_deck',
    'html-to-pptx': 'the generation pipeline behind generate_deck',
    'edit-image-fill': 'opens a native file dialog',
    'insert-media': 'opens a native file dialog',
    'export-images': 'export with a directory dialog',
    'files-add-pasted-image': 'attachment intake',
    'files-read': 'attachment intake, used by read_attachment',
    print: 'user action',
    'add-equation': 'covered by add_equation',
    'raw-get': 'covered by read_raw_xml',
    'raw-parts': 'the part listing read_raw_xml returns with no argument',
    'raw-set': 'covered by edit_raw_xml',
  },
  docs: {
    'consume-new-blank': 'app startup',
    'consume-pending-open': 'app startup',
    open: 'app-level',
    'open-path': 'app-level',
    'pick-image': 'native dialog; the agent uses insert_image',
    print: 'user action',
    recent: 'app-level',
    save: 'user action',
    'save-as': 'native dialog',
    'save-new': 'native dialog',
    'write-recovery': 'crash recovery',
    'export-pdf': 'user action with a file dialog',
    'print-pdf-buffer': 'printing',
    'save-merged-pdf': 'native dialog',
  },
  sheets: {
    'consume-new-blank': 'app startup',
  },
}

/** every ipcMain.handle('<app>:<op>') in an app's main process */
function ipcOps(app) {
  const dir = join(ROOT, 'apps', app, 'src', 'main')
  if (!existsSync(dir)) return []
  const ops = new Set()
  for (const file of readdirSync(dir)) {
    if (!file.endsWith('.ts')) continue
    const src = readFileSync(join(dir, file), 'utf8')
    // prettier wraps a long registration, putting the channel on its own line —
    // matching only same-line handles quietly hid every one of those
    for (const m of src.matchAll(/ipcMain\.handle\(\s*'([a-z0-9-]+):([a-z0-9-]+)'/g)) {
      if (m[1] === app) ops.add(m[2])
    }
  }
  return [...ops].sort()
}

/** every tool the app's agent skill declares */
function agentTools(app, files) {
  const names = new Set()
  for (const rel of files) {
    const path = join(ROOT, 'apps', app, rel)
    if (!existsSync(path)) continue
    const src = readFileSync(path, 'utf8')
    for (const m of src.matchAll(/name: '([a-z_]+)'/g)) names.add(m[1])
  }
  return [...names].sort()
}

const APPS = [
  { app: 'slides', toolFiles: ['src/renderer/ai/slides-skill.ts'] },
  { app: 'docs', toolFiles: ['src/renderer/ai/tools.ts'] },
  { app: 'sheets', toolFiles: ['src/renderer/ai/tools.ts'] },
  { app: 'pdf', toolFiles: ['src/renderer/ai/tools.ts'] },
]

const check = process.argv.includes('--check')
let unexplained = 0

for (const { app, toolFiles } of APPS) {
  const ops = ipcOps(app)
  const tools = agentTools(app, toolFiles)
  if (ops.length === 0 && tools.length === 0) continue
  const baseline = BASELINE[app] ?? {}
  const missing = ops.filter((op) => !(op in baseline))
  console.log(`\n${app}: ${ops.length} IPC operations, ${tools.length} agent tools`)
  if (missing.length === 0) {
    console.log('  every operation is either covered by a tool or listed as out of scope')
  } else {
    console.log(`  ${missing.length} operation(s) with no tool and no recorded reason:`)
    for (const op of missing) console.log(`    ${app}:${op}`)
    unexplained += missing.length
  }
  const gaps = Object.entries(baseline).filter(([, why]) => why.startsWith('GAP:'))
  if (gaps.length) {
    console.log(`  known gaps, kept deliberately for now:`)
    for (const [op, why] of gaps) console.log(`    ${app}:${op} — ${why.slice(5).trim()}`)
  }
}

if (check && unexplained > 0) {
  console.error(
    `\n${unexplained} IPC operation(s) are unreachable by the agent and unexplained.` +
      `\nAdd a tool, or record why the agent does not need it in tools/agent-reachability.mjs.`,
  )
  process.exit(1)
}
