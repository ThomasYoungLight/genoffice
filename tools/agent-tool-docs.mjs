#!/usr/bin/env node
/**
 * The agent's tool reference, generated from the tool definitions.
 *
 * What a tool accepts is declared once, in its `inputSchema`. What the agent is
 * *told* it accepts is prose in a system prompt, written by hand and never
 * re-checked. The two drift: a renamed tool, a dropped argument or a widened
 * enum changes the schema and leaves the prose describing the old shape, and
 * nothing fails — the model just gets bad instructions and we read its output
 * wondering why it keeps guessing.
 *
 * So the reference is generated, and the prompts are checked against the same
 * definitions:
 *
 *   node tools/agent-tool-docs.mjs           # write docs/agent-tools/*.md
 *   node tools/agent-tool-docs.mjs --check   # non-zero exit if stale or drifted
 *
 * --check fails on two things: a reference file that no longer matches the
 * definitions, and a tool-shaped name in a prompt that is not a tool. The
 * second is the one that catches renames. Argument names and enum values are
 * read out of the schemas, so only genuinely unrecognised words need the
 * allowlist below.
 */
import { build } from 'esbuild'
import ts from 'typescript'
import { format, resolveConfig } from 'prettier'
import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { tmpdir } from 'node:os'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const OUT_DIR = join(ROOT, 'docs', 'agent-tools')

/**
 * Where each app's tools live. `export` is either an array of tool definitions
 * or a skill factory, which is called with a stub dependency — the skills whose
 * tools are declared inside a factory take a single accessor.
 */
const APPS = [
  {
    app: 'slides',
    title: 'Slides',
    sources: [
      ['src/renderer/ai/slides-skill.ts', 'TOOLS'],
      ['src/renderer/ai/files-skill.ts', 'createFilesSkill'],
    ],
    prompts: ['src/renderer/ai/slides-skill.ts', 'src/renderer/ai/slide-qc.ts'],
  },
  {
    app: 'docs',
    title: 'Docs',
    sources: [
      ['src/renderer/ai/tools.ts', 'AGENT_TOOLS'],
      ['src/renderer/ai/files-skill.ts', 'createFilesSkill'],
    ],
    prompts: ['src/renderer/ai/docs-skill.ts'],
  },
  {
    app: 'sheets',
    title: 'Sheets',
    sources: [
      ['src/renderer/ai/tools.ts', 'WORKBOOK_TOOLS'],
      ['src/renderer/ai/files-skill.ts', 'createFilesSkill'],
      ['src/renderer/ai/search-skill.ts', 'createSearchSkill'],
    ],
    prompts: ['src/renderer/ai/workbook-skill.ts', 'src/renderer/ai/prompts/base.md'],
  },
  {
    app: 'pdf',
    title: 'PDF',
    sources: [['src/renderer/ai/tools.ts', 'AGENT_TOOLS']],
    prompts: ['src/renderer/ai/pdf-skill.ts'],
  },
]

/**
 * Tool-shaped words that are not tools, with the reason. Argument names and
 * enum values come from the schemas automatically and are not listed here; this
 * is for vocabulary the prompts invent — protocol tags, field names in examples,
 * things the model is told to write rather than call.
 */
const NOT_TOOLS = {
  slides: {
    add_: 'the add_* family, written as a stem',
    set_element_: 'the set_element_* family, written as a stem',
    execute_apps_script: "Google Slides' equivalent, named as an analogy",
    execute_layout_script:
      'a legacy alias the executor still accepts so old sessions keep working; no tool declares it',
    hero_big_number: 'a layout name offered as an example',
    insert_at: 'a page-insertion mode',
    left_text_right_image: 'a layout name offered as an example',
    max_tokens: 'a model API field, quoted in a note about generation limits',
    replace_at: 'a page-insertion mode',
    slide_generate: 'the cloud page-generation endpoint',
    three_column_cards: 'a layout name offered as an example',
    two_column: 'a layout name offered as an example',
  },
  docs: {},
  sheets: {},
  pdf: {},
}

/** A tool name: lower snake_case with at least one underscore. */
const TOOL_SHAPED = /\b[a-z][a-z0-9]*(?:_[a-z0-9]*)+\b/g

/**
 * Sheets reaches the workbook through one tool, `propose_operations`, whose
 * argument is a list of DSL operations. Those 52 operations are the surface the
 * model actually writes, and they are declared as a zod union — so the
 * reference for them is generated from that union rather than from the one
 * tool's schema, which only says "a list of operations".
 */
const DSL = {
  app: 'sheets',
  title: 'Sheets workbook operations',
  source: ['src/domain/workbook-dsl.ts', 'workbookOperationSchema'],
  file: 'sheets-operations.md',
}

/**
 * Load named exports out of an app's renderer TypeScript. These modules import
 * workspace packages and `?raw` markdown, so bundle first rather than asking
 * node to resolve any of it.
 */
async function loadExports(app, sources) {
  const entry = sources
    .map(
      ([rel, name], i) =>
        `export { ${name} as e${i} } from ${JSON.stringify(join(ROOT, 'apps', app, rel))}`,
    )
    .join('\n')
  const result = await build({
    stdin: { contents: entry, resolveDir: ROOT, sourcefile: 'entry.ts', loader: 'ts' },
    bundle: true,
    format: 'esm',
    platform: 'node',
    write: false,
    logLevel: 'silent',
    plugins: [
      {
        // markdown and stylesheets reach these modules through vite's ?raw and
        // css handling; neither carries a tool definition
        name: 'stub-assets',
        setup(b) {
          b.onLoad({ filter: /\.(md|css|svg|png)(\?raw)?$/ }, () => ({
            contents: '',
            loader: 'text',
          }))
        },
      },
    ],
  })
  const file = join(tmpdir(), `genoffice-tool-docs-${app}-${sources.length}-${process.pid}.mjs`)
  writeFileSync(file, result.outputFiles[0].text)
  const mod = await import(pathToFileURL(file).href)
  return sources.map((_, i) => mod[`e${i}`])
}

async function loadTools(app, sources) {
  const values = await loadExports(app, sources)
  const tools = []
  for (const value of values) {
    // a skill factory declares its tools inside; the stub accessor is enough to
    // build one, since nothing runs until the agent calls a tool
    const defs = typeof value === 'function' ? value(() => []).tools : value
    for (const def of defs) tools.push(def)
  }
  return tools.sort((a, b) => a.name.localeCompare(b.name))
}

// ── rendering ───────────────────────────────────────────

/** A schema's type, with enum values and array element type spelled out. */
function typeOf(schema) {
  if (!schema || typeof schema !== 'object') return 'any'
  if (Array.isArray(schema.enum)) return schema.enum.map((v) => `\`${v}\``).join(' \\| ')
  const type = Array.isArray(schema.type) ? schema.type.join('\\|') : (schema.type ?? 'any')
  if (type === 'array') return `${typeOf(schema.items)}[]`
  return type
}

/**
 * Flatten a schema into one row per argument, nesting through objects and array
 * items with dotted paths (`pages[].title`) so a reader sees the whole shape
 * without a second table.
 */
function rows(schema, prefix = '', required = new Set(), depth = 0) {
  const out = []
  const props = schema?.properties
  if (!props || depth > 3) return out
  for (const [key, value] of Object.entries(props)) {
    const path = prefix ? `${prefix}.${key}` : key
    out.push({
      path,
      type: typeOf(value),
      required: required.has(key),
      description: (value?.description ?? '').replace(/\s*\n\s*/g, ' ').replace(/\|/g, '\\|'),
    })
    const nested = value?.type === 'array' ? value.items : value
    if (nested?.type === 'object' && nested.properties) {
      const label = value?.type === 'array' ? `${path}[]` : path
      out.push(...rows(nested, label, new Set(nested.required ?? []), depth + 1))
    }
  }
  return out
}

function render(title, tools) {
  const lines = [
    `# ${title} agent tools`,
    '',
    '<!-- Generated by tools/agent-tool-docs.mjs from the tool definitions. Do not edit by hand. -->',
    '',
    `${tools.length} tools, as the model sees them: the name it calls, the description it is given, and every argument its schema accepts.`,
    '',
  ]
  for (const tool of tools) {
    lines.push(`## \`${tool.name}\``, '')
    lines.push(tool.description.replace(/\s*\n\s*/g, ' '), '')
    const args = rows(tool.inputSchema, '', new Set(tool.inputSchema?.required ?? []))
    if (args.length === 0) {
      lines.push('No arguments.', '')
      continue
    }
    lines.push('| argument | type | required | description |', '| --- | --- | --- | --- |')
    for (const a of args) {
      lines.push(`| \`${a.path}\` | ${a.type} | ${a.required ? 'yes' : ''} | ${a.description} |`)
    }
    lines.push('')
  }
  return lines.join('\n')
}

// ── the workbook DSL, read out of its zod union ─────────

const defOf = (schema) => schema?.def ?? schema?._def

/** Strip the wrappers that only say whether a value may be absent. */
function unwrap(schema) {
  const def = defOf(schema)
  if (!def) return schema
  if (def.type === 'optional' || def.type === 'default' || def.type === 'catch')
    return unwrap(def.innerType)
  return schema
}

/**
 * The discriminant key of a tagged union, or undefined. zod 4 types a
 * discriminated union as a plain `union` and carries the key alongside, so the
 * key's presence is what distinguishes the two.
 */
const discriminatorOf = (def) => (def?.type === 'union' ? def.discriminator : undefined)

const isOptional = (schema) => {
  const type = defOf(schema)?.type
  return type === 'optional' || type === 'default'
}

/** A field's type, spelled the way the model has to write it. */
function zodType(schema, depth = 0) {
  const def = defOf(unwrap(schema))
  if (!def) return 'any'
  switch (def.type) {
    case 'nullable':
      return `${zodType(def.innerType, depth)} \\| null`
    case 'array':
      return `${zodType(def.element, depth + 1)}[]`
    case 'literal':
      return def.values.map((v) => `\`${v}\``).join(' \\| ')
    case 'enum':
      return Object.values(def.entries ?? def.values ?? {})
        .map((v) => `\`${v}\``)
        .join(' \\| ')
    case 'object':
      return 'object'
    case 'record':
      return `object of ${zodType(def.valueType, depth + 1)}`
    case 'union': {
      const discriminator = discriminatorOf(def)
      if (discriminator) {
        // name the variants by their discriminant; the fields of each are
        // listed underneath, so "object | object | …" would say nothing
        const tags = def.options.map(
          (o) => defOf(unwrap(defOf(unwrap(o)).shape?.[discriminator]))?.values?.[0],
        )
        if (tags.every(Boolean)) return tags.map((t) => `\`${t}\``).join(' \\| ')
      }
      if (depth > 1) return 'union'
      const parts = def.options.map((o) => zodType(o, depth + 1))
      return [...new Set(parts)].join(' \\| ')
    }
    case 'pipe':
      return zodType(def.in, depth)
    default:
      return def.type
  }
}

/** One row per field, nesting through objects and arrays with dotted paths. */
function zodRows(objectSchema, prefix = '', depth = 0) {
  const shape = defOf(unwrap(objectSchema))?.shape
  if (!shape || depth > 2) return []
  const out = []
  for (const [key, value] of Object.entries(shape)) {
    if (key === 'op' && prefix === '') continue // the discriminant is the heading
    const path = prefix ? `${prefix}.${key}` : key
    out.push({ path, type: zodType(value), required: !isOptional(value) })
    const inner = unwrap(value)
    const innerDef = defOf(inner)
    const element = innerDef?.type === 'array' ? unwrap(innerDef.element) : inner
    const elementDef = defOf(element)
    const label = innerDef?.type === 'array' ? `${path}[]` : path
    const discriminator = discriminatorOf(elementDef)
    if (elementDef?.type === 'object') {
      out.push(...zodRows(element, label, depth + 1))
    } else if (discriminator) {
      // the variants are the whole content of a field like `rule` or `validation`;
      // collapsing them to "object" would hide the part the model has to get right
      for (const option of elementDef.options) {
        const shape = defOf(unwrap(option))?.shape ?? {}
        const tag = defOf(unwrap(shape[discriminator]))?.values?.[0]
        const prefix = `${label}[${discriminator}=${tag}]`
        // the discriminant is already named in the prefix
        out.push(
          ...zodRows(option, prefix, depth + 1).filter(
            (r) => r.path !== `${prefix}.${discriminator}`,
          ),
        )
      }
    }
  }
  return out
}

function dslOperations(union) {
  return defOf(union).options.map((option) => ({
    name: defOf(unwrap(option).shape.op).values[0],
    fields: zodRows(option),
  }))
}

function renderDsl(title, operations) {
  const lines = [
    `# ${title}`,
    '',
    '<!-- Generated by tools/agent-tool-docs.mjs from workbookOperationSchema. Do not edit by hand. -->',
    '',
    `The \`propose_operations\` tool takes a list of these. ${operations.length} operations; every field is validated, so an unknown one is rejected rather than ignored.`,
    '',
  ]
  for (const op of operations) {
    lines.push(`## \`${op.name}\``, '')
    if (op.fields.length === 0) {
      lines.push('No fields.', '')
      continue
    }
    lines.push('| field | type | required |', '| --- | --- | --- |')
    for (const f of op.fields)
      lines.push(`| \`${f.path}\` | ${f.type} | ${f.required ? 'yes' : ''} |`)
    lines.push('')
  }
  return lines.join('\n')
}

// ── the docs commands, read out of their TypeScript types ──

/**
 * Docs has the same shape of surface as sheets: one tool, `apply_commands`,
 * whose argument is a list of commands the model writes as JSON. Unlike the
 * workbook DSL these are declared as TypeScript interfaces, which are gone by
 * runtime — so they are read from the source with the compiler's own parser
 * rather than by introspecting a value.
 */
const COMMANDS = {
  app: 'docs',
  title: 'Docs commands',
  source: 'src/renderer/ai/commands.ts',
  union: 'Command',
  file: 'docs-commands.md',
}

function docsCommands() {
  const path = join(ROOT, 'apps', COMMANDS.app, COMMANDS.source)
  const source = ts.createSourceFile(path, readFileSync(path, 'utf8'), ts.ScriptTarget.Latest, true)
  const interfaces = new Map()
  let union = null
  for (const statement of source.statements) {
    if (ts.isInterfaceDeclaration(statement)) interfaces.set(statement.name.text, statement)
    if (ts.isTypeAliasDeclaration(statement) && statement.name.text === COMMANDS.union) {
      union = statement.type
    }
  }
  if (!union || !ts.isUnionTypeNode(union)) throw new Error(`${COMMANDS.union} is not a union type`)

  /** the one-line doc comment above a member, if it has one */
  const commentOf = (node) => {
    const text = source.getFullText().slice(node.pos, node.getStart())
    const doc = /\/\*\*([\s\S]*?)\*\//.exec(text)?.[1]
    return doc
      ? doc
          .replace(/^\s*\*\s?/gm, '')
          .replace(/\s*\n\s*/g, ' ')
          .trim()
          .replace(/\|/g, '\\|')
      : ''
  }

  /** array element type node, or the node itself */
  const elementOf = (node) => (node && ts.isArrayTypeNode(node) ? node.elementType : node)

  const fieldsOf = (node, prefix = '', depth = 0) => {
    const decl = typeof node === 'string' ? interfaces.get(node) : node
    if (!decl || !decl.members || depth > 2) return []
    const out = []
    for (const member of decl.members) {
      if (!ts.isPropertySignature(member) || !member.name) continue
      const key = member.name.getText(source)
      const path = prefix ? `${prefix}.${key}` : key
      const element = elementOf(member.type)
      const array = member.type !== element
      // a nested shape is expanded into its own rows; writing the literal type
      // into the cell would paste a multi-line declaration into a table
      const nested =
        element && (ts.isTypeLiteralNode(element) || interfaces.has(element.getText(source)))
      const type = nested
        ? array
          ? 'object[]'
          : 'object'
        : (member.type?.getText(source).replace(/\s*\n\s*/g, ' ') ?? 'unknown')
      out.push({
        path,
        type: type.replace(/\|/g, '\\|'),
        required: !member.questionToken,
        description: commentOf(member),
      })
      if (nested) {
        const target = ts.isTypeLiteralNode(element)
          ? element
          : interfaces.get(element.getText(source))
        out.push(...fieldsOf(target, array ? `${path}[]` : path, depth + 1))
      }
    }
    return out
  }

  return union.types.map((member) => {
    const property = ts.isTypeLiteralNode(member) ? member.members[0] : null
    const key = property?.name?.getText(source) ?? '?'
    const type = property?.type?.getText(source) ?? ''
    return {
      name: key,
      fields: fieldsOf(type),
      description: commentOf(interfaces.get(type) ?? member),
    }
  })
}

function renderCommands(title, commands) {
  const lines = [
    `# ${title}`,
    '',
    `<!-- Generated by tools/agent-tool-docs.mjs from the ${COMMANDS.union} type. Do not edit by hand. -->`,
    '',
    `The \`apply_commands\` tool takes a list of these, each a single-key object, applied in order. ${commands.length} commands.`,
    '',
  ]
  for (const command of commands) {
    lines.push(`## \`${command.name}\``, '')
    if (command.description) lines.push(command.description, '')
    if (command.fields.length === 0) {
      lines.push('No fields.', '')
      continue
    }
    lines.push('| field | type | required | description |', '| --- | --- | --- | --- |')
    for (const f of command.fields) {
      lines.push(
        `| \`${f.path}\` | \`${f.type}\` | ${f.required ? 'yes' : ''} | ${f.description} |`,
      )
    }
    lines.push('')
  }
  return lines.join('\n')
}

// ── prompt drift ────────────────────────────────────────

/** Every argument name and enum value anywhere in a set of schemas. */
function schemaVocabulary(tools) {
  const words = new Set()
  const walk = (node) => {
    if (!node || typeof node !== 'object') return
    if (Array.isArray(node)) return node.forEach(walk)
    for (const [key, value] of Object.entries(node)) {
      if (key === 'properties' && value && typeof value === 'object') {
        for (const name of Object.keys(value)) words.add(name)
      }
      if (key === 'enum' && Array.isArray(value)) {
        for (const v of value) if (typeof v === 'string') words.add(v)
      }
      walk(value)
    }
  }
  for (const tool of tools) walk(tool.inputSchema)
  return words
}

/**
 * Tool-shaped words in the prompts that name nothing. A word that survives the
 * tool names, the schema vocabulary and the allowlist is usually a tool that was
 * renamed or removed while the prose kept referring to it.
 */
/** Every field name, literal and enum value anywhere in a zod schema. */
function zodVocabulary(schema, words = new Set(), depth = 0) {
  const def = defOf(schema)
  if (!def || depth > 8) return words
  if (def.shape) for (const key of Object.keys(def.shape)) words.add(key)
  for (const v of def.values ?? []) if (typeof v === 'string') words.add(v)
  for (const v of Object.values(def.entries ?? {})) if (typeof v === 'string') words.add(v)
  for (const child of [
    def.innerType,
    def.element,
    def.in,
    def.out,
    ...(def.options ?? []),
    ...Object.values(def.shape ?? {}),
  ]) {
    if (child) zodVocabulary(child, words, depth + 1)
  }
  return words
}

function promptDrift(app, tools, promptFiles, extra = new Set()) {
  const known = new Set(tools.map((t) => t.name))
  const vocabulary = schemaVocabulary(tools)
  for (const word of extra) vocabulary.add(word)
  const allowed = NOT_TOOLS[app] ?? {}
  const found = new Map()
  for (const rel of promptFiles) {
    const path = join(ROOT, 'apps', app, rel)
    if (!existsSync(path)) continue
    const src = readFileSync(path, 'utf8')
    for (const [word] of src.matchAll(TOOL_SHAPED)) {
      if (known.has(word) || vocabulary.has(word) || word in allowed) continue
      if (!found.has(word)) found.set(word, rel)
    }
  }
  return found
}

// ── run ─────────────────────────────────────────────────

const check = process.argv.includes('--check')
let stale = 0
let drifted = 0

if (!check) mkdirSync(OUT_DIR, { recursive: true })

const prettierConfig = await resolveConfig(join(OUT_DIR, 'x.md'))

/**
 * Write, or in --check mode report, one generated reference file. The output is
 * run through prettier so the committed files also pass `format:check` —
 * otherwise formatting and generation would each keep undoing the other.
 */
async function emit(file, source, note) {
  const path = join(OUT_DIR, file)
  const doc = await format(source, { ...prettierConfig, parser: 'markdown' })
  const current = existsSync(path) ? readFileSync(path, 'utf8') : null
  if (current === doc) {
    if (!check) console.log(`  docs/agent-tools/${file} up to date (${note})`)
    return
  }
  if (check) {
    console.log(`  docs/agent-tools/${file} is out of date`)
    stale++
    return
  }
  writeFileSync(path, doc)
  console.log(`  wrote docs/agent-tools/${file} (${note})`)
}

const [dslUnion] = await loadExports(DSL.app, [DSL.source])
const operations = dslOperations(dslUnion)

for (const { app, title, sources, prompts } of APPS) {
  const tools = await loadTools(app, sources)
  console.log(`\n${app}:`)
  await emit(`${app}.md`, render(title, tools), `${tools.length} tools`)

  const extra = new Set()
  if (app === DSL.app) {
    await emit(DSL.file, renderDsl(DSL.title, operations), `${operations.length} operations`)
    zodVocabulary(dslUnion, extra)
  }
  if (app === COMMANDS.app) {
    const commands = docsCommands()
    await emit(
      COMMANDS.file,
      renderCommands(COMMANDS.title, commands),
      `${commands.length} commands`,
    )
    for (const command of commands) {
      extra.add(command.name)
      for (const field of command.fields) extra.add(field.path.split('.').pop())
    }
  }

  const drift = promptDrift(app, tools, prompts, extra)
  if (drift.size) {
    console.log(`  ${drift.size} tool-shaped name(s) in the prompts that are not tools:`)
    for (const [word, rel] of drift) console.log(`    ${word}  (${rel})`)
    drifted += drift.size
  }
}

if (check && (stale || drifted)) {
  if (stale)
    console.error(`\nRun \`node tools/agent-tool-docs.mjs\` to regenerate the tool reference.`)
  if (drifted)
    console.error(
      `\nA prompt names something tool-shaped that no tool provides.` +
        `\nFix the prompt, or record what the word is in NOT_TOOLS in tools/agent-tool-docs.mjs.`,
    )
  process.exit(1)
}
