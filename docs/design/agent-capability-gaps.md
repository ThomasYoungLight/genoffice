# Remaining agent capability gaps — design

Status: proposed. Written 2026-08-04, after closing the reachability gaps in
slides, sheets and docs.

This records the work that is still outstanding after the OfficeCLI comparison,
why each piece is harder than it looks, and how to build it. It exists because
the remaining items lived only in conversation: `tools/agent-reachability.mjs`
records which operations the agent may not reach and why, but nothing recorded
what we had decided _not to build yet_.

## Where things stand

Every IPC operation in every app is now either reachable by that app's agent or
carries a recorded reason it is not (`npm run agent-check`). The reference in
[`docs/agent-tools/`](../agent-tools/README.md) is generated from the tool
definitions and checked in CI, so it cannot drift from them.

What is left is not reachability. It is four things the products genuinely
cannot do.

| #   | Item                                   | Blocked on                              |
| --- | -------------------------------------- | --------------------------------------- |
| 1   | Raw OOXML for docs and sheets          | save-path design, per format            |
| 2   | Sheets slicers, page breaks, rich text | new engine + sidecar support            |
| 3   | Docs caption and index fields          | nothing; small                          |
| 4   | OLE, Zoom links, threaded comments, …  | nothing; low value (recommend: decline) |

## Background: three different package architectures

Everything below turns on one fact, so it is worth stating plainly. The three
engines hold an open file in three different ways, and "write raw XML into a
part" means something different in each.

**pptx — a live part map.** `PackageArchive` keeps the original bytes of every
zip entry in a mutable `Map`, and `savePptx` writes unmodified entries back
byte-for-byte. Editing a part is `entries.set(path, bytes)`, and the save picks
it up for free. This is why `edit_raw_xml` already exists for slides.

**docx — an immutable base plus a rebuilt body.** `saveDocx(parsed, finalBlocks,
options)` starts from `parsed.internal.originalBytes` and `documentXml`, then
**regenerates the body of document.xml from the ProseMirror blocks**. There is
no part-override hook in `SaveOptions`. So a raw write to document.xml's body
would be silently discarded at save — the worst possible failure mode for this
feature — and a raw write to any other part has nowhere to live.

**xlsx — replacement-based save through the sidecar.** The Rust sidecar already
speaks part surgery: `ArchiveManifest`, `ReadEntries`, `ScanEntries` and
`SaveArchive { replacements, removals, additions }`. `xlsx-package-io.ts` drives
the save through exactly that replacement set. A raw part edit is one more entry
in it.

An earlier read of this got sheets and docs backwards — sheets was assumed to be
the hard one because of the sidecar, when the sidecar is precisely what makes it
easy, and docs was assumed to be a mechanical port of the slides work when its
save path makes it the awkward one. The table above reflects the corrected view.

---

## 1. Raw OOXML for docs and sheets

### Goal

The same escape hatch slides has: reach a property no tool models, without
opening the file in another program. Same shape as `read_raw_xml` /
`edit_raw_xml` — a unique-match find/replace, not "here is a new part", because
a small edit is one we can check.

### 1a. Sheets

The mechanism exists; this is wiring.

1. **`sheets:raw-parts`** — `ArchiveManifest` on the workbook path, filtered to
   `.xml` and `.rels`, with the short names `/workbook`, `/sheet[N]` (resolved
   through the workbook's sheet order, not file numbering), `/styles`,
   `/sharedStrings`, `/theme`.
2. **`sheets:raw-get`** — `ReadEntries` into a temp dir, read the file back,
   apply the same 400 KB cap slides uses.
3. **`sheets:raw-set`** — resolve, validate (unique match, well-formed XML),
   then add the part to the pending replacement set that `xlsx-package-io`
   already maintains.

The validation chain from `packages/pptx-engine/src/raw.ts` ports directly.
What does **not** port is the fourth gate, "the affected sheets still parse":
sheets has no in-memory re-parse to run. The equivalent is to round-trip the
edited part through the sidecar's own reader (`ReadRange` on an affected sheet)
and treat an error as a rollback signal. Without that gate a raw edit to
styles.xml can produce a workbook Excel refuses, and we would not find out until
the user did.

**Univer divergence.** Sheets renders through Univer from a model the sidecar
produced. A raw edit changes the file but not the model, so the grid will show
stale values until reload. Either re-read the affected sheet after a successful
edit (preferred) or say so plainly in the tool result. Do not leave it silent —
that is the "session renders one document and saves another" failure the slides
version was designed to avoid.

### 1b. Docs

Needs a save-path change first.

1. **`SaveOptions.partOverrides?: ReadonlyMap<string, Uint8Array>`** — parts
   written verbatim into the output zip, applied after the body rebuild so the
   two can never fight.
2. **Refuse writes to `word/document.xml`.** Its body is regenerated from the
   editor's blocks, so a raw write there is a lie. The refusal message should
   name the tool that does own the change (`apply_commands`, `replace_blocks`).
   Reads of document.xml stay allowed and are genuinely useful.
3. **Reads: everything. Writes: styles, numbering, theme, settings, header and
   footer parts, `[Content_Types].xml`.** That set is exactly the part of a
   document our model does not represent, which is what the escape hatch is for.
4. Re-parse gate: run `parseDocx` over the patched bytes and roll back on
   failure, mirroring slides.

### Risk

Unchanged from the slides version and stated in `raw.ts`: well-formed,
parseable XML can still be schema-invalid OOXML that Word or Excel refuses to
open. Validation cannot close that, which is why the edit shape stays narrow and
the tool tells the model to say what it changed.

---

## 2. Sheets slicers, page breaks, rich text

The only item here that is not plumbing. Each needs a DSL operation, gateway
support, Univer rendering and sidecar round-tripping, and each is independent of
the others.

### Rich text within a cell

The highest value of the three: a bold word inside a sentence in a cell is
ordinary formatting that we cannot express, because `set_cell` takes a scalar.

- DSL: `set_cell_rich { sheetId, address, runs: [{ text, bold?, italic?,
color?, ... }] }`.
- Storage: `sharedStrings.xml` `<si>` with `<r>` runs rather than a bare `<t>`.
  `packages/file-parse/src/xlsx.ts` already reads runs, so the reader half
  exists.
- Univer: rich text in a cell is supported by the presets; map runs to its
  inline style model.
- Watch: shared-string deduplication. A rich string must not collide with the
  plain string of the same text.

### Row and column page breaks

- DSL: `set_page_breaks { sheetId, rows: number[], cols: number[] }` — declare
  the full set rather than add/remove one at a time, which keeps it idempotent
  and matches how `set_freeze` already works.
- Storage: `<rowBreaks>`/`<colBreaks>` in the worksheet part.
- Univer: renders print-area breaks; verify before promising it.

### Slicers

The largest and the least valuable of the three. A slicer is four coupled parts
(`slicer`, `slicerCache`, a drawing anchor, and a table or pivot relationship),
and Univer has no slicer UI, so we would be writing something we cannot render.

**Recommendation: do rich text and page breaks; leave slicers until a user asks.**
Writing a control the app cannot display is how you get a file that survives one
round trip and breaks on the second.

---

## 3. Docs caption and index fields

`generateCaptionXml` and `generateIndexFieldXml` exist in the engine and are
driven by the References ribbon. Exposing them is the same `DocExtras` pattern
used for footnotes, watermark, sources, page setup and comments. Small, and
worth doing when someone next touches that file.

---

## 4. The tail: OLE, Zoom links, threaded comments, form fields, permStart

All real absences against OfficeCLI. All expensive relative to what they buy:

- **OLE embedding** — needs binary part handling and has no in-app rendering
  story in any of the three apps.
- **PowerPoint Zoom links** — a presentation-navigation feature we do not
  render.
- **Modern threaded comments** — we write the legacy comment part, which
  PowerPoint still reads. The gain is thread fidelity, not capability.
- **Legacy form fields, permStart** — superseded by content controls and
  document protection respectively.

**Recommendation: decline these explicitly** and record them in
`tools/agent-reachability.mjs`'s baseline vocabulary so the decision is visible
rather than looking like an oversight — the same discipline the rest of the
surface follows.

---

## Verification

Each item follows the pattern the recent work established:

1. Engine-level tests against a real fixture file, covering the refusals, not
   just the happy path. The raw work's most useful test was the one proving a
   rejected edit leaves the bytes untouched.
2. Tool-level tests that the engine's refusal reaches the model verbatim — a
   tool that summarises "raw edit failed" leaves it guessing.
3. Round trip: save, reopen, assert the feature survived.
4. E2E in the running app for anything that crosses the IPC seam, which is where
   the last several real bugs lived.

**Office is installed on the development machine** (Word, Excel and PowerPoint
under `/Applications`), so generated OOXML can and should be opened in the real
application before a format feature is called done. This is not optional
polish — the first time it was run, on the slide equations shipped in `849281c`,
it found a bug that every layer of our own testing had passed: `\\sum_{i=1}^{n} i`
left the n-ary operator's `m:e` slot empty and emitted the summand as a sibling,
which Word and PowerPoint both draw as a dotted placeholder box followed by a
stray symbol. Our tests asserted the OMML contained `<m:nary>`, which it did.

Two lessons worth keeping:

- **Check the exported PDF, not just the editing canvas.** PowerPoint draws
  empty slots as dotted boxes while editing; a square root's `<m:deg/>` is
  legitimately empty (with `degHide`) and shows a marker on canvas that is
  absent from the rendered output. Reading the canvas alone would have produced
  a second "bug" that does not exist.
- A file opening without a repair prompt is necessary, not sufficient. Both
  defects above opened cleanly.

---

## Plan

Revised after the first Office check found a defect our whole test suite had
passed. That changes the ordering: verification is no longer a footnote at the
end, it is the cheapest bug-finding tool available and it goes first.

### Step 0 — Verify what has already shipped (highest value, days not weeks)

Eleven commits of format-writing work landed before anyone opened the output in
Office. One of them had a bug. The rest are unexamined, and the cost of looking
is minutes each.

Open a file exercising each feature in the real application, check the exported
PDF as well as the canvas, and fix what turns up. Ordered by how novel or
fragile the markup is, not by how recently it shipped:

| Priority | Feature                                            | Commit               | Why it is first                                                                                 |
| -------- | -------------------------------------------------- | -------------------- | ----------------------------------------------------------------------------------------------- |
| ~~1~~    | ~~Slide transitions and animations~~               | `4575901`            | **Done.** Animations pass; morph was broken by a wrong namespace URI, now fixed                 |
| ~~2~~    | ~~Sheets icon-set / aboveAverage / timePeriod CF~~ | `cea0bd3`            | **Done.** Icons and aboveAverage correct; timePeriod could not be saved at all, now implemented |
| ~~3~~    | ~~Docs formulas in Word~~                          | `9b4ec40`            | **Done.** Formulas correct; turned up an unrelated Compatibility Mode issue                     |
| ~~4~~    | ~~The five added chart types~~                     | `e54265a`            | **Done.** All five correct; one cosmetic axis issue left open                                   |
| ~~5~~    | ~~pptx sections and comments~~                     | `cea0bd3`            | **Done.** Sections correct; comment badges all stacked and timestamps were 8 hours out          |
| ~~6~~    | ~~Docs comments, text boxes, shapes~~              | `e113810`            | **Done.** All three correct; `w:date` was true UTC where Word wants local digits                |
| ~~7~~    | ~~Header/footer, both formats~~                    | `9e0dd02`, `31f9b34` | **Done.** Output right in both; a partial update silently deleted the untouched fields          |
| ~~8~~    | ~~Mermaid flowcharts as native shapes~~            | `76b7678`            | **Done.** Graph, shapes and arrow directions all correct; node captions were pinned to the top  |

### Step 0 results

**Row 1 — animations and transitions: one bug, fixed.**

Passing: all four animation classes (entrance, emphasis, exit, motion path),
their triggers grouped correctly in PowerPoint's Animation Pane, the motion path
drawn on the canvas, and the `push` and `dissolve` transitions.

Failing: `morph`. We wrote the 2015 extension namespace as
`.../powerpoint/2015/main`; PowerPoint's is `.../powerpoint/2015/09/main`. One
missing path segment, and the symptom is not a broken file — PowerPoint matches
no `mc:Choice`, declines the `mc:Fallback` as well, and reports the slide as
having no transition at all. The existing test asserted `<p159:morph`,
`Requires="p159"` and the fallback markup, every one of which was correct. It
never asserted the URI. It does now.

Two methodology notes for the remaining rows, both learned the hard way:

- **The accessibility tree's "Has Transition" description does not report
  morph** — not for our file, and not for one PowerPoint wrote itself. Read the
  Transitions ribbon's highlighted button, after confirming from the status bar
  which slide is actually selected.
- **Saving via AppleScript dropped the morph child element**, producing a
  reference sample that looked authoritative and was not. Use ⌘S. A good while
  went into diagnosing a file PowerPoint had quietly degraded on the way out.

**Row 2 — conditional formats: one bug, fixed.**

Passing in Excel: icon sets in natural order, icon sets inverted via
`reverse="1"`, the worst-first `4Rating` set, and `aboveAverage` (highlighting
exactly the values above the mean). The icon-order inversion this row was
prioritised for turns out to be correct.

Failing: `timePeriod` could not be written at all. `xlsx-cf.ts` threw
"Date-occurring rules cannot be saved yet" — deliberately, with a test asserting
it. The bug was not in the writer; it was that `cea0bd3` added `timePeriod` to
the workbook DSL and to the Univer mapping anyway. The agent could create a rule
that rendered in the grid and then made the whole workbook fail to save, because
`applyCfRules` throws and nothing on the save path catches it. Nothing checked
the two halves against each other.

Now implemented for all ten periods the DSL offers, each with the formula Excel
evaluates — the `timePeriod` attribute alone highlights nothing. A test walks
the DSL's own list so the two cannot drift apart again.

**Row 3 — docs formulas in Word: no bug in the formulas.**

Word typesets all five block equations and all three inline ones, and the
summand sits inside the operator, so the fix made for slides in `9b4ec40`
reaches docs through the shared converter as intended. Nothing to change.

It did turn up something unrelated: **Word opens our documents in Compatibility
Mode.** `buildBlankDocx` writes `document.xml`, `styles.xml` and
`numbering.xml`, and no `word/settings.xml`; without a `compatibilityMode`
compat setting Word assumes an older format, disables newer features and says so
in the title bar. It affects new documents only — a document opened from an
existing file keeps its original settings part through the patch-based save.

**Fixed.** `buildBlankDocx` now writes a `settings.xml` declaring compatibility
mode 15, registered both in `[Content_Types].xml` and in the document rels — a
part that is written but not declared is invisible to Word, and one declared but
missing makes the package invalid. Verified in Word: the title bar reads
"blank-new" where it previously read "formulas - Compatibility Mode".

**Row 4 — the five added chart types: all correct, one cosmetic issue open.**

Verified in PowerPoint: `scatter` plots its points, `radar` draws four spokes
with both series as polygons, `comboBarLine` puts the last series on a line
against a secondary right axis, `barH` is genuinely horizontal, and
`barPercentStacked` stacks to 100% with the right proportions. No repair prompt,
all data correct.

Open: on axes with many ticks — the radar's spokes and the percent-stacked
value axis — PowerPoint draws dense tick marks that read as hatching across the
plot. Our chart parts emit no `majorTickMark`/`minorTickMark` at all, so this is
PowerPoint's default rather than something we set. Two hypotheses were tried and
**both were wrong**: defaulting `majorGridlines` on for radar, and suppressing
tick marks on the radar axes. Neither changed the rendering, and both were
reverted rather than committed as unverified "fixes". Whatever causes it is not
those two things, and it is cosmetic — the data is right in every chart.

**Row 5 — pptx sections and comments: sections correct, two comment bugs, both
fixed.**

Sections are right. PowerPoint's thumbnail pane shows "Overview" (slide 1),
"Findings" (2–3) and "Next steps" (4–5), and its accessibility tree agrees
("Findings Expanded section 2 of 3 (2 slides)"). The auto-created "Default
Section" for the leading unsectioned slide is what PowerPoint would have done
itself. No repair prompt.

Comments carried two defects, neither of which produced an error:

- **Every badge on a slide stacked into one.** The stagger was
  `10 + (count % 8) * 6`, and `p:pos` units are small enough that six of them
  are a fraction of a pixel. Three comments drew one badge.
- **Timestamps were wrong by the author's UTC offset.** `dt` was written with
  `new Date().toISOString()`, so it ended in `Z`. PowerPoint ignores the
  designator and reads the digits as local time: comments created seconds
  earlier displayed as "8 hours ago" in UTC+8. Worse, the error is permanent —
  PowerPoint's own re-save strips the `Z` and keeps the shifted digits.

Both fixes were measured rather than guessed. Adding comments through
PowerPoint's own UI and reading back what it wrote gave
`dt="2026-08-04T13:29:58.737"` (local, no designator) and positions 106 → 202 →
298, i.e. a step of 96 from a start of 106. The fix reproduces both exactly, and
re-verification shows three separate badges and three "A few seconds ago".

PowerPoint also writes `<p15:threadingInfo timeZoneBias="-480"/>` on its own
comments. That was deliberately **not** copied: the local-time fix alone makes
the timestamps display correctly, and adding markup that is not needed to fix
the observed defect is how unverified guesses get shipped.

Worth carrying into row 6: the docs side writes `w:date` with a `Z`
(`review-actions.ts`, `revisions.ts`, `protocol.ts`). That may well be correct —
Word and PowerPoint need not agree — so it is a question for row 6 to answer in
Word, not a bug to fix here on the strength of PowerPoint's behaviour.

**Row 6 — docs comments, text boxes and shapes: one bug, fixed; it was the
question row 5 left open.**

Comments, text boxes and shapes all reach Word correctly. The comment anchors to
its range and shows in the pane with the right author, initials and text; the
ellipse renders as an ellipse with the right fill; the text box renders. No
repair prompt, and the title bar reads "row6" rather than "Compatibility Mode",
which incidentally confirms the settings-part fix on the docs save path too.

The open question from row 5 turned out to be a bug after all, and the answer is
stranger than either alternative I had in mind. **Word writes local wall-clock
digits with a `Z` suffix** and reads them back as local, keeping the true
instant in a separate `w16du:dateUtc`. Letting Word author both a comment and a
tracked insertion at 13:49 local in UTC+8 gave:

```xml
<w:ins w:date="2026-08-04T13:49:00Z" w16du:dateUtc="2026-08-04T05:49:00Z">
```

So `w:date` is not UTC despite the designator, and our four writers — two in
`review-actions.ts`, one in `protocol.ts`, one in `revisions.ts` — all put a
genuine UTC timestamp there. Every comment and every tracked change displayed
shifted by the author's UTC offset: a comment made at 1:51 PM read "5:42 AM"
next to Word's own "1:46 PM" in the same pane. All four now share one
`wordTimestamp()` helper in the engine, next to the wire format it belongs to.
Re-verified in Word: "4/8/26 1:51 PM".

`w16du:dateUtc` is deliberately not written. It needs namespace plumbing on
every part that carries a date, and the display is already right without it.

Note that this is the *opposite* convention from PowerPoint, which writes local
digits with **no** designator. The rule that holds across both is the one worth
remembering: **the digits are local wall-clock time**; only the suffix differs.

An observation, not a bug: inserting a text box and then a shape leaves the two
floats overlapping, because each anchors at the end of the document with the
same offset. Word behaves comparably when two shapes are inserted without being
moved, and nothing about the file is wrong, so this was left alone.

A harness note: the first attempt opened in Compatibility Mode and briefly
looked like a regression of `0322ac2`. It was the *test helper* `buildDocx`,
which writes no settings part — not the product path. Rebasing the fixture on
`buildBlankDocx` removed the confound. Check what the harness builds before
believing what the harness shows.

**Row 7 — header/footer in both formats: the output is right in both, one
partial-update bug on the slides side.**

Both formats reach their application correctly. PowerPoint places the date left,
the footer centred and the slide number right on every slide, with
`<a:fld type="slidenum">` resolving per slide — and, the stronger check,
**PowerPoint's own Header and Footer dialog reads all three back**: "Date and
time / Fixed / 4 August 2026", "Slide number" ticked, "Footer" with our text. A
round trip through Office's own UI, not just its renderer. Word puts the header
at the top and the footer at the bottom with the `PAGE` field resolved
("Confidential 1"), verified in an exported PDF.

The bug is on the way in, not on the way out. `applyHeaderFooter` clears the
whole `dt`/`ftr`/`sldNum` family before writing the enabled parts back, so an
omitted field reads as a deletion. The dialog always sends all three and never
noticed. The agent sends only the field it was asked to change, so **"add slide
numbers" silently deleted the footer** — applying `{date}` to a deck with a
footer and slide numbers left only the date.

`mergeHeaderFooter` now fills the gaps from the current state before the engine
sees them; `null` still means remove, only an absent field inherits.

The instructive part is why a test suite with a test for exactly this missed it.
`arrangement-tools.test.ts` asserts the tool omits untouched fields —
`expect('date' in op).toBe(false)` — and then stops at the IPC mock. It pinned
the promise the skill makes and never checked the other side keeps it. A mock
boundary is where an invariant goes to die: both halves were self-consistent and
the contract between them was wrong. The new tests live in the engine, where the
clearing behaviour actually is, and one of them asserts that behaviour directly
so the reason `mergeHeaderFooter` exists cannot quietly stop being true.

**Row 8 — Mermaid flowcharts: one cosmetic bug, fixed. Sweep complete.**

The hard parts are right. `flowchart TD` with a decision branch and a join
parses to the expected graph, the decision node is a diamond and the rest are
rectangles, longest-path ranking puts the join below both branches, and — the
part the original commit flagged as needing care — **all five arrowheads point
at their target**, including the two right-to-left edges that depend on the
`flipH` the layout records. No repair prompt.

The bug: every node caption sat against the top edge of its box.
`buildSpXml` emitted `<a:bodyPr wrap="square" rtlCol="0"/>` with no `anchor`,
and the OOXML default is top. Correct for a text box, wrong for a labelled
shape. `NewElementOptions.anchor` now carries it through `AddElementOp`, and
`insert_diagram` asks for `middle`; re-rendered, the captions sit in the middle
of their boxes.

Two things deliberately not done. The default for `addElement` is unchanged —
centring every added autoshape may well be right, but I could not get Word or
PowerPoint to author a comparable shape for me to measure, and this sweep has
already produced three markup guesses that a screenshot refuted. Better a narrow
fix the diagram owns than a broad one resting on an assumption. And the existing
tool already reports undrawn arrow labels rather than dropping them, which is
the honest behaviour; drawing them is a layout problem, not a bug.

An observation: the diagram lands in a fixed content area and does not avoid
existing content, so inserting one onto a slide that already has body text
overlaps it. The tool takes explicit `x`/`y`/`w`/`h`, so the caller can place it;
nothing about the file is wrong.

A methodology note to go with row 1's: **rebuild before you conclude.** A "no
fill" screenshot sent me chasing the dxf markup, and a controlled pair differing
only in that markup rendered identically. The first file was simply stale.

Row 2 was first written down on a guess — that icon sets need an `x14`
extension block. They do not, here: `xlsx-cf.ts` writes plain OOXML and refuses
outright the mixed sets and custom orderings that would need the extended
format. The correction is the point, and the habit it stands for: check the code
before writing down a risk, and check the application before believing the code.

**Done when:** each row has been opened in its application, the PDF checked, and
either fixed or recorded as verified. Expect this to find more than one bug.

### Step 1 — Sheets raw (small)

Wire `raw-parts` / `raw-get` / `raw-set` onto the sidecar's existing archive
commands, port the validation chain from `pptx-engine/src/raw.ts`, add the
sidecar-reader round-trip gate and the Univer re-read.

**Done when:** the four gates are tested, a raw edit survives save/reopen, the
grid does not show stale values, and the result opens in Excel.

### Step 2 — Sheets rich text (medium)

`set_cell_rich`, shared-string runs, Univer mapping. The one item here a user
would notice unprompted.

**Done when:** a rich cell round-trips with its runs, does not collide with the
plain string of the same text, and Excel shows the runs.

### Step 3 — Docs raw (medium)

`SaveOptions.partOverrides` first, then the tools, with document.xml writes
refused and the refusal naming the tool that owns the change.

**Done when:** a styles.xml edit survives save/reopen and opens in Word, a
document.xml write is refused with a useful message, and a corrupting edit rolls
back.

### Step 4 — Sheets page breaks, docs caption/index (small)

Both are single-session pieces. Fold them into whatever touches those files
next rather than scheduling them.

### Not scheduled

Slicers, and the whole of section 4. Record the decisions; revisit on request.

### Standing rule

No format-writing feature is done until it has been opened in the application
that has to read it. This is a build step, not a release gate — the equations
bug would have been a five-minute fix on the day and instead shipped.
