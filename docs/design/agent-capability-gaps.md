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
| 5        | pptx sections and comments                         | `cea0bd3`            | Separate parts with their own relationships                                                     |
| 6        | Docs comments, text boxes, shapes                  | `e113810`            | Floating anchors and a comment part                                                             |
| 7        | Header/footer, both formats                        | `9e0dd02`, `31f9b34` | Placeholder fields repeated across every page/slide                                             |
| 8        | Mermaid flowcharts as native shapes                | `76b7678`            | Plain preset geometry; lowest risk of the set                                                   |

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
