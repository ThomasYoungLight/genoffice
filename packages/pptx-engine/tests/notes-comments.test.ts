import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  addSlideComment,
  createBlankPptx,
  deleteSlideComment,
  getSlideComments,
  getSlideNotes,
  openPptx,
  savePptx,
  setSlideNotes,
} from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

describe('speaker notes', () => {
  it('creates notesSlide (and notesMaster) on a blank deck and survives save → reopen', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(getSlideNotes(opened.archive, opened.deck.slides[0]!.path)).toBe('')

    expect(setSlideNotes(opened, 0, 'note line one\nline two <special & chars>')).toBe(true)
    expect(getSlideNotes(opened.archive, opened.deck.slides[0]!.path)).toBe(
      'note line one\nline two <special & chars>',
    )

    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideNotes(reopened.archive, reopened.deck.slides[0]!.path)).toBe(
      'note line one\nline two <special & chars>',
    )
    // notesMaster is registered
    expect(reopened.archive.has('ppt/notesMasters/notesMaster1.xml')).toBe(true)
    expect(reopened.archive.readText('ppt/presentation.xml')).toContain('notesMasterIdLst')
    expect(reopened.archive.readText('[Content_Types].xml')).toContain('notesSlide+xml')
  })

  it('overwrites existing notes and clears them', async () => {
    const opened = await openPptx(await createBlankPptx())
    setSlideNotes(opened, 0, 'old note')
    setSlideNotes(opened, 0, 'new note')
    expect(getSlideNotes(opened.archive, opened.deck.slides[0]!.path)).toBe('new note')
    setSlideNotes(opened, 0, '')
    expect(getSlideNotes(opened.archive, opened.deck.slides[0]!.path)).toBe('')
  })

  it('works on a real pptx fixture', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    setSlideNotes(opened, 0, 'Introduce yourself first')
    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideNotes(reopened.archive, reopened.deck.slides[0]!.path)).toBe(
      'Introduce yourself first',
    )
  })
})

describe('slide comments', () => {
  it('adds comments (authors part + slide part) and survives save → reopen', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slidePath = opened.deck.slides[0]!.path
    expect(getSlideComments(opened.archive, slidePath)).toEqual([])

    const c1 = addSlideComment(opened, 0, { author: 'Carol', text: 'Title is too long' })!
    const c2 = addSlideComment(opened, 0, { author: 'Carol', text: 'Change palette & fonts' })!
    expect(c1.idx).toBe(1)
    expect(c2.idx).toBe(2)
    expect(c1.authorId).toBe(c2.authorId)

    const reopened = await openPptx(await savePptx(opened))
    const list = getSlideComments(reopened.archive, reopened.deck.slides[0]!.path)
    expect(list.map((c) => c.text)).toEqual(['Title is too long', 'Change palette & fonts'])
    expect(list[0]!.author).toBe('Carol')
    expect(reopened.archive.readText('ppt/presentation.xml')).toBeTruthy()
    expect(reopened.archive.readText('[Content_Types].xml')).toContain('commentAuthors+xml')
  })

  it('assigns distinct author ids and deletes by (authorId, idx)', async () => {
    const opened = await openPptx(await createBlankPptx())
    const a = addSlideComment(opened, 0, { author: 'Alice', text: 'A' })!
    const b = addSlideComment(opened, 0, { author: 'Bob', text: 'B' })!
    expect(a.authorId).not.toBe(b.authorId)

    expect(deleteSlideComment(opened, 0, { authorId: a.authorId, idx: a.idx })).toBe(true)
    const slidePath = opened.deck.slides[0]!.path
    expect(getSlideComments(opened.archive, slidePath).map((c) => c.text)).toEqual(['B'])
    // Deleting a non-existent one again → false
    expect(deleteSlideComment(opened, 0, { authorId: a.authorId, idx: a.idx })).toBe(false)
  })

  /**
   * Both of these were verified against PowerPoint by adding comments through
   * its own UI and reading back what it wrote: dt="2026-08-04T13:29:58.737"
   * (local, no Z) and pos 106 → 202 → 298.
   */
  it('timestamps a comment in local time, which is how PowerPoint reads dt back', async () => {
    const opened = await openPptx(await createBlankPptx())
    const before = new Date()
    const c = addSlideComment(opened, 0, { author: 'Carol', text: 'when?' })!

    // a trailing Z is the bug: PowerPoint takes the digits as local time either
    // way, so a UTC timestamp shows up offset by the author's UTC offset
    expect(c.dt).not.toMatch(/Z$/)
    expect(c.dt).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}$/)

    // the digits are local wall-clock, so reading them as local recovers now
    const parsed = new Date(c.dt).getTime()
    expect(parsed).toBeGreaterThanOrEqual(before.getTime() - 1000)
    expect(parsed).toBeLessThanOrEqual(Date.now() + 1000)
  })

  it('separates comment badges far enough apart to be distinguishable', async () => {
    const opened = await openPptx(await createBlankPptx())
    for (const text of ['one', 'two', 'three']) addSlideComment(opened, 0, { author: 'Carol', text })

    const part = opened.archive.readText('ppt/comments/comment1.xml') ?? ''
    const xs = [...part.matchAll(/<p:pos x="(\d+)"/g)].map((m) => Number(m[1]))
    expect(xs).toEqual([106, 202, 298])
  })

  it('comments on different slides go to different parts', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    addSlideComment(opened, 0, { author: 'Carol', text: 'first slide' })
    addSlideComment(opened, 1, { author: 'Carol', text: 'second slide' })
    const reopened = await openPptx(await savePptx(opened))
    expect(
      getSlideComments(reopened.archive, reopened.deck.slides[0]!.path).map((c) => c.text),
    ).toEqual(['first slide'])
    expect(
      getSlideComments(reopened.archive, reopened.deck.slides[1]!.path).map((c) => c.text),
    ).toEqual(['second slide'])
  })
})
