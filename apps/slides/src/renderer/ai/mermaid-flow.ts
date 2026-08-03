/**
 * Mermaid flowcharts → positioned shapes and connectors.
 *
 * A flowchart is the one diagram a model writes reliably and our SmartArt
 * cannot express: SmartArt offers seven fixed layouts, and an arbitrary graph
 * is not one of them. Mermaid is the notation models already know, so the
 * agent writes `flowchart TD; A[Start] --> B{OK?}` and this turns it into
 * native shapes the user can then move and restyle.
 *
 * Pure: source in, geometry out. The tool layer does the IPC. Only flowcharts
 * are handled — sequence and the rest are a different layout problem and are
 * rejected by name rather than mis-drawn.
 */

/** node outline; maps onto OOXML preset geometry at emit time */
export type NodeShape = 'box' | 'round' | 'stadium' | 'diamond' | 'circle'

export interface FlowNode {
  id: string
  text: string
  shape: NodeShape
}

export interface FlowEdge {
  from: string
  to: string
  label?: string | undefined
  /** dotted in the source (-.->): drawn as a lighter line */
  dashed?: boolean | undefined
}

export interface FlowGraph {
  /** layer direction: top-down or left-right */
  dir: 'TD' | 'LR'
  nodes: FlowNode[]
  edges: FlowEdge[]
}

const HEADER = /^\s*(?:flowchart|graph)\s+(TB|TD|BT|LR|RL)\b/i
/** `A[Text]`, `A(Text)`, `A([Text])`, `A((Text))`, `A{Text}`, or bare `A` */
const NODE = /^([A-Za-z0-9_-]+)\s*(\(\(|\(\[|\[|\(|\{)?\s*([^\]})]*?)\s*(\)\)|\]\)|\]|\)|\})?$/
/** `-->`, `---`, `-.->`, `==>`, each with an optional `|label|` */
const LINK = /\s*(-{2,3}>|-{3}|-\.-+>|={2,}>)\s*(?:\|([^|]*)\|)?\s*/

function shapeOf(open: string | undefined): NodeShape {
  switch (open) {
    case '((':
      return 'circle'
    case '([':
      return 'stadium'
    case '(':
      return 'round'
    case '{':
      return 'diamond'
    default:
      return 'box'
  }
}

/**
 * Parse the flowchart subset. Returns an error string rather than throwing:
 * the model wrote this text and the message goes back to it as a correction.
 */
export function parseMermaidFlow(source: string): { graph: FlowGraph } | { error: string } {
  const text = source.replace(/\r/g, '').trim()
  if (!text) return { error: 'The diagram source is empty' }
  const firstLine = text.split('\n')[0] ?? ''
  const head = HEADER.exec(firstLine)
  if (!head) {
    const kind = /^\s*([a-zA-Z]+)/.exec(firstLine)?.[1] ?? 'that'
    return {
      error: `Only flowcharts are supported; "${kind}" is not one. Start the source with "flowchart TD" or "flowchart LR".`,
    }
  }
  const raw = head[1]!.toUpperCase()
  // BT and RL are the same layout reversed; draw them in the natural
  // direction rather than silently ignoring the reversal
  const dir: 'TD' | 'LR' = raw === 'LR' || raw === 'RL' ? 'LR' : 'TD'
  const reversed = raw === 'BT' || raw === 'RL'

  const nodes = new Map<string, FlowNode>()
  const edges: FlowEdge[] = []
  const declare = (token: string): string | null => {
    const m = NODE.exec(token.trim())
    if (!m) return null
    const id = m[1]!
    const label = (m[3] ?? '').trim().replace(/^["']|["']$/g, '')
    const existing = nodes.get(id)
    if (!existing) nodes.set(id, { id, text: label || id, shape: shapeOf(m[2]) })
    else if (label && existing.text === id) {
      existing.text = label
      existing.shape = shapeOf(m[2])
    }
    return id
  }

  // the header may be followed by statements on the same line
  // ("flowchart TD; A --> B"), which is how a diagram written inline arrives
  const lines = text.split('\n')
  const statements = [firstLine.slice(head[0].length), ...lines.slice(1)]
    .flatMap((line) => line.split(';'))
    .map((s) => s.trim())
    .filter((s) => s && !s.startsWith('%%'))

  for (const statement of statements) {
    if (/^(subgraph|end|classDef|class|style|click|linkStyle)\b/i.test(statement)) continue
    /**
     * Split into alternating node tokens and links, then join them up:
     * `A --> B --> C` is three tokens and two links, and every link connects
     * the token before it to the token after.
     */
    const tokens: string[] = []
    const links: Array<{ label?: string | undefined; dashed: boolean }> = []
    let rest = statement
    for (let guard = 0; guard < 50; guard++) {
      const link = LINK.exec(rest)
      if (!link) {
        tokens.push(rest)
        break
      }
      tokens.push(rest.slice(0, link.index))
      links.push({ label: link[2]?.trim(), dashed: link[1]!.includes('.') })
      rest = rest.slice(link.index + link[0].length)
    }
    const ids: string[] = []
    for (const token of tokens) {
      const id = declare(token)
      if (id === null) return { error: `Could not read "${token.trim()}" as a node` }
      ids.push(id)
    }
    links.forEach((link, i) => {
      const from = ids[i]
      const to = ids[i + 1]
      if (!from || !to) return
      edges.push({
        from,
        to,
        ...(link.label ? { label: link.label } : {}),
        ...(link.dashed ? { dashed: true } : {}),
      })
    })
  }

  if (nodes.size === 0) return { error: 'The diagram declares no nodes' }
  const graph: FlowGraph = {
    dir,
    nodes: [...nodes.values()],
    edges: reversed ? edges.map((e) => ({ ...e, from: e.to, to: e.from })) : edges,
  }
  return { graph }
}

// ── layout ──────────────────────────────────────────────

export interface FlowShape {
  id: string
  text: string
  shape: NodeShape
  x: number
  y: number
  w: number
  h: number
}

export interface FlowConnector {
  from: string
  to: string
  label?: string | undefined
  dashed?: boolean | undefined
  /** bounding box of the line */
  x: number
  y: number
  w: number
  h: number
  /** the line runs right-to-left or bottom-to-top inside its box */
  flipH: boolean
  flipV: boolean
}

export interface FlowLayout {
  shapes: FlowShape[]
  connectors: FlowConnector[]
}

/**
 * Longest-path layering: a node sits one layer after the deepest thing that
 * reaches it, which is what makes a flowchart read as a sequence. Cycles are
 * broken by the visit guard rather than hanging.
 */
function rankNodes(graph: FlowGraph): Map<string, number> {
  const rank = new Map<string, number>()
  for (const n of graph.nodes) rank.set(n.id, 0)
  const incoming = new Map<string, string[]>()
  for (const e of graph.edges) incoming.set(e.to, [...(incoming.get(e.to) ?? []), e.from])
  // relax repeatedly; node count bounds the longest possible path
  for (let pass = 0; pass < graph.nodes.length; pass++) {
    let moved = false
    for (const n of graph.nodes) {
      const parents = incoming.get(n.id) ?? []
      if (!parents.length) continue
      const want = Math.max(...parents.map((p) => (rank.get(p) ?? 0) + 1))
      if (want > (rank.get(n.id) ?? 0)) {
        rank.set(n.id, want)
        moved = true
      }
    }
    if (!moved) break
  }
  return rank
}

/** Place the graph inside `box`, in the direction the source asked for. */
export function layoutFlow(
  graph: FlowGraph,
  box: { x: number; y: number; w: number; h: number },
): FlowLayout {
  const rank = rankNodes(graph)
  const layers = new Map<number, FlowNode[]>()
  for (const n of graph.nodes) {
    const r = rank.get(n.id) ?? 0
    layers.set(r, [...(layers.get(r) ?? []), n])
  }
  const depth = Math.max(...layers.keys()) + 1
  const widest = Math.max(...[...layers.values()].map((l) => l.length))
  const alongGap = 28
  const acrossGap = 24

  // "along" runs in the flow direction, "across" is the spread within a layer
  const alongTotal = graph.dir === 'TD' ? box.h : box.w
  const acrossTotal = graph.dir === 'TD' ? box.w : box.h
  const alongSize = Math.max(48, (alongTotal - alongGap * (depth - 1)) / depth)
  const acrossSize = Math.max(80, (acrossTotal - acrossGap * (widest - 1)) / widest)

  const shapes: FlowShape[] = []
  const at = new Map<string, FlowShape>()
  for (const [r, layer] of [...layers.entries()].sort((a, b) => a[0] - b[0])) {
    const span = layer.length * acrossSize + (layer.length - 1) * acrossGap
    const acrossStart = (acrossTotal - span) / 2
    layer.forEach((node, i) => {
      const along = r * (alongSize + alongGap)
      const across = acrossStart + i * (acrossSize + acrossGap)
      const shape: FlowShape =
        graph.dir === 'TD'
          ? {
              ...node,
              x: Math.round(box.x + across),
              y: Math.round(box.y + along),
              w: Math.round(acrossSize),
              h: Math.round(alongSize),
            }
          : {
              ...node,
              x: Math.round(box.x + along),
              y: Math.round(box.y + across),
              w: Math.round(alongSize),
              h: Math.round(acrossSize),
            }
      shapes.push(shape)
      at.set(node.id, shape)
    })
  }

  const connectors: FlowConnector[] = []
  for (const e of graph.edges) {
    const a = at.get(e.from)
    const b = at.get(e.to)
    if (!a || !b || a === b) continue
    // leave from the face pointing at the target and arrive on the opposite one
    const [x1, y1, x2, y2] =
      graph.dir === 'TD'
        ? [a.x + a.w / 2, a.y + a.h, b.x + b.w / 2, b.y]
        : [a.x + a.w, a.y + a.h / 2, b.x, b.y + b.h / 2]
    connectors.push({
      from: e.from,
      to: e.to,
      ...(e.label ? { label: e.label } : {}),
      ...(e.dashed ? { dashed: true } : {}),
      x: Math.round(Math.min(x1, x2)),
      y: Math.round(Math.min(y1, y2)),
      // a zero-width box would collapse; keep a hairline so the line exists
      w: Math.max(1, Math.round(Math.abs(x2 - x1))),
      h: Math.max(1, Math.round(Math.abs(y2 - y1))),
      // the box loses direction, so record it: the arrow must point at the target
      flipH: x2 < x1,
      flipV: y2 < y1,
    })
  }
  return { shapes, connectors }
}

/** OOXML preset geometry for a node outline. */
export function presetFor(shape: NodeShape): string {
  switch (shape) {
    case 'diamond':
      return 'diamond'
    case 'circle':
      return 'ellipse'
    case 'stadium':
    case 'round':
      return 'roundRect'
    default:
      return 'rect'
  }
}
