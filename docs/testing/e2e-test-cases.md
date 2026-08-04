# Sheets end-to-end test cases

## What "end to end" means here, and what it cannot do

The standing rule for this project is that **no format-writing feature is done
until it has been opened in the application that has to read it**. These cases
encode that: each one ends in a real artefact opened by a real reader — Excel,
Preview, or the running app — not in an assertion about our own output.

Two consequences worth stating up front, because they shape everything below.

**E2E cases contribute nothing to code coverage.** They drive a separate
Electron process (and often Excel), which the vitest v8 provider cannot
instrument. Coverage comes from the integration tests in `tests/`, which
exercise the same code paths in-process. Where a case can be automated that
way, it is, and the case says so.

**Manual cases are cheap to run and expensive to trust.** Each manual case
records the exact commands and the fixture that produces it, so a re-run is
mechanical rather than a fresh investigation.

## Fixture builders

Fixtures are generated, never committed as binaries, so what they contain is
readable in the diff. `tests/fixtures/build-fixtures.mjs` writes them to a temp
directory; every case names the one it uses.

| fixture               | contains                                                        |
| --------------------- | --------------------------------------------------------------- |
| `twocell.xlsx`        | data + chart and image on **twoCellAnchor** (Excel's style)      |
| `onecell.xlsx`        | the same, on **oneCellAnchor** (openpyxl/pandas default)         |
| `multipage.xlsx`      | 120 rows + visuals at rows 2, 45, 75, 100                        |
| `cf.xlsx`             | colour scale, data bar, icon set, `cellIs` fill                  |
| `filtered.xlsx`       | an AutoFilter over a small table                                 |

---

## E1 — A workbook opens, renders, and saves without repair

**Why**: the whole product rests on Excel accepting what we write. A repair
prompt is a total failure however good the feature is.

1. Open `twocell.xlsx` in the app; wait for "Workbook fully loaded".
2. Edit one cell; save with ⌘S.
3. Open the saved file in Excel.

**Expect**: no "Repaired" or "Removed Records" dialog; the edit is present; the
chart and image are unchanged and in their original cells.

**Automated**: partly — `xlsx-gateway.test.ts` and `xlsx-package-io.test.ts`
cover the save plan and archive write. Excel's acceptance cannot be automated.

---

## E2 — Charts, images, shapes and sparklines reach print and PDF

**Why**: they are float DOM overlaid on the grid, outside Univer's model, so
nothing about the on-screen result implies they will print. They did not, for
the whole life of the feature, and nothing failed.

1. Open `twocell.xlsx`. Insert a shape via Insert → Shapes → Diamond.
2. File → Export PDF.
3. Open the PDF in Preview.

**Expect**: chart, image and shape all appear, each over the cells it occupies
on screen; no resize handles or delete buttons in the output.

**Automated**: yes for placement and chrome-stripping —
`print-visuals.test.ts` (layout) and `print-visuals-capture.test.ts` (DOM
capture). The PDF itself is manual.

---

## E3 — oneCellAnchor drawings print at the size the file asked for

**Why**: openpyxl and pandas write `oneCellAnchor` by default, so this is the
common shape for any script- or agent-generated workbook, not an edge case. It
used to collapse into a sliver of one cell.

1. Open `onecell.xlsx`.
2. Compare on-screen cell span against the same file open in Excel.
3. Export to PDF.

**Expect**: chart spans D2:H12 and image D16:E19 in both apps; the PDF matches.

**Automated**: yes for the arithmetic — `drawing-anchor.test.ts`. The
cross-application comparison is manual and is the part that actually proves it.

---

## E4 — Moving and resizing a visual survives the round trip

**Why**: the edit rewrites XML in place. A malformed anchor is not caught by
anything we own; Excel is the only judge.

1. Open `onecell.xlsx`. Drag the picture to a new cell. Drag its SE handle.
2. Save. Open in Excel.

**Expect**: no repair prompt; the picture is where it was dropped, at the size
it was given; it is still a `oneCellAnchor` with no `<to>` marker; the file
contains no `xdr:` prefix it did not already have.

**Automated**: yes for the XML — `xlsx-drawing-edit.test.ts` covers prefixed
and bare namespaces, move, resize and delete. Excel's verdict is manual.

---

## E5 — Conditional formatting fidelity in print

**Why**: a known, deliberate partial gap. Recording it stops it being
rediscovered as a bug.

1. Open `cf.xlsx`. 2. Export to PDF. 3. Compare against the screen.

**Expect**: colour scales and `cellIs` fills print (they are cell styles);
**data bars and icon sets do not** (Univer paints them on the canvas). The
`render_preview` tool description states this so the agent is not misled.

**Automated**: no. This case exists to pin a known limitation.

---

## E6 — The agent can see what it built

**Why**: without it the agent edits blind, and every visual defect has to be
caught by a human.

1. Configure a provider with a key (see the note on `userData` below).
2. Open `twocell.xlsx`. Ask: "Call render_preview and describe what you see."

**Expect**: the agent calls `render_preview` once and describes the chart title,
the bar labels and the coloured rectangle — content not derivable from cells.

**Automated**: partly — `loop.test.ts` pins that images ride on a user turn
after the tool results and never inside a tool result. The model actually
looking is manual.

**Precondition that bites**: `userData` is per `productName`, so
`npm run dev -w @genoffice/sheets` ("GenOffice Sheets") does **not** see the
settings configured in the shell ("GenOffice" / "GenOffice Dev"). Run the
shell, or the default provider resolves to genspark with no key.

---

## E7 — The raw OOXML escape hatch cannot corrupt a file

**Why**: it bypasses every model we have. Its only safety is its gates.

1. Read a part, apply an edit that produces malformed XML, and one that makes
   the workbook unopenable. 2. Save.

**Expect**: both refused with a specific error; the file on disk is untouched;
a well-formed edit applies and the workbook still opens in Excel.

**Automated**: yes — `xlsx-raw.test.ts` covers the gates.

---

## E8 — Page setup reaches the printed page

1. Set print area, repeat row 1, gridlines on, landscape, fit-to-width.
2. Export to PDF.

**Expect**: each setting is visible in the output; repeated title rows appear on
every page and do not push visuals out of register with their cells.

**Automated**: partly — `print-html.test.ts` and `xlsx-page-setup.test.ts`
cover the payload and the XML. Pagination is manual.

---

## E9 — The preload boundary rejects malformed traffic

**Why**: preload is the only validation between an untrusted renderer and the
main process, and it is an allow-list — anything it does not name is silently
dropped, which is how `extWidthEmu` went missing with no error anywhere.

Exercised as integration tests rather than E2E: every bridge method is called
with malformed input and must throw rather than forward, and with valid input
must forward exactly and validate the response.

**Automated**: yes — `preload-bridge.test.ts`.

---

## E10 — Autosave recovery

1. Edit a workbook. Kill the app without saving. Reopen the same file.

**Expect**: the recovery prompt appears; Discard leaves the file as it was on
disk; Restore reinstates the edits.

**Automated**: no.

---

## Coverage

Coverage is measured on `apps/sheets/src/**` with
`npx vitest run --coverage.enabled --coverage.provider=v8 --coverage.include='src/**'`.

E2E cases do not appear in it, by construction. The number is a statement about
the integration tests in `tests/`, and is tracked per area because the areas
differ by an order of magnitude and a single figure hides that.

---

## Execution log

Run on macOS 15 (Darwin 25.5.0), Excel for Mac, Preview, Electron dev build.
Automated results are from `npx vitest run --root apps/sheets`.

| case | result | evidence |
| ---- | ------ | -------- |
| E1 workbook opens, renders, saves without repair | **partial** | verified with a *visual* edit, not a cell edit: `onecell.xlsx` had its picture moved, saved from the app, reopened in Excel with no repair prompt and the chart untouched. The cell-edit path is covered only by the automated save tests; the Excel half of it has not been run |
| E2 visuals reach print/PDF | **pass** | chart, image and a diamond shape inserted in-app all present in the exported PDF, over the right cells, no handles or delete buttons |
| E3 oneCellAnchor sizes correctly | **pass** | same file in GenOffice and Excel both span chart D2:H12, image D16:E19; PDF matches |
| E4 move/resize round trip | **pass** | picture dragged and resized, saved, opened in Excel: no repair, position and size held, still a `oneCellAnchor`, no `<to>`, no stray `xdr:` |
| E5 conditional formatting fidelity | **pass (records a known gap)** | colour scale and `cellIs` fill print; **data bars and icon sets do not** |
| E6 agent sees what it built | **blocked** | no provider key in the `GenOffice Sheets` userData; `loop.test.ts` covers the image-delivery half |
| E7 raw OOXML gates | **pass** | automated, `xlsx-raw.test.ts` |
| E8 page setup reaches the page | **pass** | 3-page export with Repeat Row 1: visuals stayed in register with their cells across page breaks (measured: identical 482px offset with and without the repeated header) |
| E9 preload boundary | **pass** | automated, `preload-bridge.test.ts`, 53 cases |
| E10 autosave recovery | **partial** | the prompt appeared on reopen after the app was killed, and Discard correctly loaded the on-disk file. **Restore was not exercised** — the half that actually recovers work is unverified |

## Coverage target: the application layer

The meaningful target is the **application layer** — every `.ts` under
`apps/sheets/src`, i.e. domain, gateway, main, preload, shared, and the
renderer's non-component logic. It excludes `.tsx` React components, which sit
at 2.3% and are the wrong place to spend testing effort in this codebase: every
defect found in this project to date has been in format logic or a boundary
layer, never in a component.

| | statements | |
| --- | --- | --- |
| application layer (`.ts`) | 7726/14845 | **52.0%** |
| UI components (`.tsx`) | 88/3775 | 2.3% |

Reaching 80% of the application layer needs **4,150 more covered statements**,
and they are concentrated in ten files:

| uncovered | cumulative | file |
| --------: | ---------: | ---- |
| 1274 | 1274 | `renderer/univer-sync.ts` |
| 908 | 2182 | `main/sheets-main.ts` |
| 615 | 2797 | `renderer/ribbon-actions.ts` |
| 595 | 3392 | `renderer/workbook-ops.ts` |
| 523 | 3915 | `preload/index.ts` |
| 342 | 4257 | `renderer/pivot-actions.ts` |
| 276 | 4533 | `renderer/data-tools-actions.ts` |
| 232 | 4765 | `renderer/visual-actions.ts` |
| 221 | 4986 | `renderer/visual-edit-sync.ts` |
| 200 | 5186 | `renderer/plan-operations.ts` |

Covering ~80% of those ten reaches the target almost exactly. They are all
tractable for the same reason: each was extracted from App.tsx so that "every
function receives its runtime and state explicitly", so a test double supplies
the runtime. `tests/helpers/fake-univer.ts` is that double and is already in
place — building it was the bulk of the cost for the first file, and it is
reusable across the remaining nine.

Two lessons from doing the first one, worth having before starting the rest:

- **Derive the double's shape from the consumer, not the schema.** Grepping
  `fileMeta.` and `state.editJournal.` in the module under test gives the exact
  field list in one step; guessing from the Zod schema took several rounds,
  because the readers dereference fields the schema marks optional.
- **A failing assertion is as likely to be a wrong expectation as a bug.**
  `readFormats` returning nothing for an unstyled cell looked like a defect and
  is the correct, deliberate contract — it keeps the agent's payload small.

## Coverage result

Measured on `apps/sheets/src/**`, 1,019 tests:

| area     | before | after | statements |
| -------- | ------ | ----- | ---------- |
| gateway  | 87.5%  | 87.5% | 4025/4601  |
| domain   | 82.1%  | 82.1% | 1450/1767  |
| ai       | 79.1%  | 79.1% | 34/43      |
| shared   | 69.2%  | 70.1% | 75/107     |
| preload  | 0.0%   | 14.3% | 92/643     |
| renderer | 18.4%  | 18.4% | 1919/10429 |
| main     | 6.8%   | 6.8%  | 70/1030    |
| **total**| 40.66% | **41.16%** | 7665/18620 |

**The 80% target is not met, and no amount of E2E work would meet it.** E2E
cases drive a separate Electron process that the v8 provider cannot instrument,
so they contribute zero measured coverage by construction. The number moves
only with in-process tests.

What reaching 80% actually requires, from the table: 7,231 more covered
statements, of which **8,510 of the 10,948 uncovered ones are in `renderer`**.
That is React components, dialogs and the Univer integration layer — it needs
jsdom plus a substantial Univer test double, and it is measured in days.

The order that buys the most per unit of effort:

1. **`renderer/*.ts` non-component logic first** — `print-html`, `edit-journal`,
   `univer-sync`, `visual-actions`, `workbook-ops` are plain functions already
   partly covered; finishing them is ordinary unit testing.
2. **`main` (6.8%)** — mostly IPC handlers; testable with an `electron` mock in
   the same style as `preload-bridge.test.ts`.
3. **Finish `preload` (14.3%)** — same file, same pattern, ~35 methods left.
4. **React components last** — highest cost, lowest defect density in this
   codebase's history. Every defect found in this project so far has been in
   format logic or the boundary layers, not in components.
