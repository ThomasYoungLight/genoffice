/**
 * Mermaid flowchart parsing and layout.
 *
 * The model writes this text, so the parser's job is to be forgiving about
 * the ways it is written and precise about what it produces. The layout's job
 * is that a reader can follow the arrows: layers in flow order, nothing
 * overlapping, and every arrow pointing at its target — the last one matters
 * because a connector's bounding box does not record direction, so a
 * right-to-left edge drawn without the flip points backwards.
 */
import { describe, expect, it } from 'vitest'
import {
  layoutFlow,
  parseMermaidFlow,
  presetFor,
  type FlowGraph,
} from '../src/renderer/ai/mermaid-flow'

const parse = (src: string) => {
  const r = parseMermaidFlow(src)
  if ('error' in r) throw new Error(r.error)
  return r.graph
}

describe('parseMermaidFlow', () => {
  it('reads nodes, labels and shapes', () => {
    const g = parse('flowchart TD\n A[Start] --> B{Ready?}\n B --> C((Done))')
    expect(g.dir).toBe('TD')
    expect(g.nodes.map((n) => [n.id, n.text, n.shape])).toEqual([
      ['A', 'Start', 'box'],
      ['B', 'Ready?', 'diamond'],
      ['C', 'Done', 'circle'],
    ])
    expect(g.edges).toEqual([
      { from: 'A', to: 'B' },
      { from: 'B', to: 'C' },
    ])
  })

  it('handles a chain written on one line with semicolons', () => {
    const g = parse('flowchart LR; A[One] --> B[Two] --> C[Three]')
    expect(g.dir).toBe('LR')
    expect(g.edges).toEqual([
      { from: 'A', to: 'B' },
      { from: 'B', to: 'C' },
    ])
    expect(g.nodes).toHaveLength(3)
  })

  it('keeps edge labels and dotted links', () => {
    const g = parse('flowchart TD\n A -->|yes| B\n A -.->|no| C')
    expect(g.edges[0]).toEqual({ from: 'A', to: 'B', label: 'yes' })
    expect(g.edges[1]).toEqual({ from: 'A', to: 'C', label: 'no', dashed: true })
  })

  it('does not lose a label declared on a later mention of the same node', () => {
    const g = parse('flowchart TD\n A --> B\n B[Named later]')
    expect(g.nodes.find((n) => n.id === 'B')?.text).toBe('Named later')
  })

  it('draws BT and RL in the natural direction with the edges reversed', () => {
    const bt = parse('flowchart BT\n A --> B')
    expect(bt.dir).toBe('TD')
    expect(bt.edges).toEqual([{ from: 'B', to: 'A' }])
  })

  it('ignores styling and subgraph statements rather than choking on them', () => {
    const g = parse(
      'flowchart TD\n %% a comment\n classDef big fill:#f00\n A --> B\n style A fill:#0f0',
    )
    expect(g.nodes.map((n) => n.id)).toEqual(['A', 'B'])
  })

  it('names the diagram type it cannot draw, instead of drawing it wrong', () => {
    const r = parseMermaidFlow('sequenceDiagram\n A->>B: hi')
    expect('error' in r && r.error).toContain('sequenceDiagram')
    expect('error' in r && r.error).toContain('flowchart')
  })

  it('rejects empty and node-less sources', () => {
    expect('error' in parseMermaidFlow('')).toBe(true)
    expect('error' in parseMermaidFlow('flowchart TD')).toBe(true)
  })
})

describe('layoutFlow', () => {
  const BOX = { x: 84, y: 200, w: 1112, h: 420 }
  const shapesOf = (g: FlowGraph) => layoutFlow(g, BOX)

  it('puts each node one layer past the deepest thing that reaches it', () => {
    const { shapes } = shapesOf(parse('flowchart TD\n A --> B\n A --> C\n B --> D\n C --> D'))
    const y = Object.fromEntries(shapes.map((s) => [s.id, s.y]))
    expect(y.A).toBeLessThan(y.B!)
    expect(y.B).toBe(y.C) // same layer
    expect(y.D).toBeGreaterThan(y.B!) // after the deepest parent, not the first
  })

  it('lays LR out along x instead of y', () => {
    const { shapes } = shapesOf(parse('flowchart LR\n A --> B'))
    const [a, b] = shapes
    expect(b!.x).toBeGreaterThan(a!.x)
    expect(b!.y).toBe(a!.y)
  })

  it('keeps every shape inside the box it was given', () => {
    const { shapes } = shapesOf(
      parse('flowchart TD\n A --> B\n A --> C\n A --> D\n B --> E\n C --> E\n D --> E'),
    )
    for (const s of shapes) {
      expect(s.x).toBeGreaterThanOrEqual(BOX.x)
      expect(s.y).toBeGreaterThanOrEqual(BOX.y)
      expect(s.x + s.w).toBeLessThanOrEqual(BOX.x + BOX.w + 1)
      expect(s.y + s.h).toBeLessThanOrEqual(BOX.y + BOX.h + 1)
    }
  })

  it('does not overlap shapes', () => {
    const { shapes } = shapesOf(parse('flowchart TD\n A --> B\n A --> C\n A --> D\n A --> E'))
    for (let i = 0; i < shapes.length; i++) {
      for (let j = i + 1; j < shapes.length; j++) {
        const a = shapes[i]!
        const b = shapes[j]!
        const hit = a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h
        expect(hit, `${a.id} overlaps ${b.id}`).toBe(false)
      }
    }
  })

  it('records the direction the box cannot carry, so arrows point at the target', () => {
    // B is the left child, C the right; A --> B runs leftwards
    const { shapes, connectors } = shapesOf(parse('flowchart TD\n A --> B\n A --> C'))
    const left = shapes.find((s) => s.id === 'B')!
    const right = shapes.find((s) => s.id === 'C')!
    expect(left.x).toBeLessThan(right.x)
    const toLeft = connectors.find((c) => c.to === 'B')!
    const toRight = connectors.find((c) => c.to === 'C')!
    expect(toLeft.flipH).toBe(true)
    expect(toRight.flipH).toBe(false)
    expect(toLeft.flipV).toBe(false) // both run downwards
  })

  it('leaves a drawable box for a perfectly straight connector', () => {
    const { connectors } = shapesOf(parse('flowchart TD\n A --> B'))
    expect(connectors[0]!.w).toBeGreaterThan(0)
    expect(connectors[0]!.h).toBeGreaterThan(0)
  })

  it('connects from the face pointing at the target', () => {
    const { shapes, connectors } = shapesOf(parse('flowchart TD\n A --> B'))
    const a = shapes.find((s) => s.id === 'A')!
    const c = connectors[0]!
    expect(c.y).toBe(a.y + a.h) // leaves the bottom edge
  })

  it('skips an edge to a node that was never declared', () => {
    const graph = { ...parse('flowchart TD\n A --> B'), edges: [{ from: 'A', to: 'ghost' }] }
    expect(layoutFlow(graph, BOX).connectors).toHaveLength(0)
  })

  it('terminates on a cycle', () => {
    const { shapes } = shapesOf(parse('flowchart TD\n A --> B\n B --> C\n C --> A'))
    expect(shapes).toHaveLength(3)
  })
})

describe('presetFor', () => {
  it('maps mermaid outlines onto preset geometry the engine knows', () => {
    expect(presetFor('diamond')).toBe('diamond')
    expect(presetFor('circle')).toBe('ellipse')
    expect(presetFor('round')).toBe('roundRect')
    expect(presetFor('stadium')).toBe('roundRect')
    expect(presetFor('box')).toBe('rect')
  })
})
