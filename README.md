# GenOffice

An AI-native office suite for macOS and Windows: word processor, spreadsheet,
presentations, and PDF — five Electron apps sharing one engine layer, built
around AI editing as a first-class workflow rather than a bolted-on chat box.

## Download

Signed installers built from `main`:

- **macOS** (Apple Silicon): [GenOffice-0.4.110-arm64.dmg](https://github.com/genspark-ai/genoffice/releases/download/v0.4.110/GenOffice-0.4.110-arm64.dmg)
- **Windows** (x64): [GenOfficeSetup-v0.4.110.exe](https://github.com/genspark-ai/genoffice/releases/download/v0.4.110/GenOfficeSetup-v0.4.110.exe)

Other versions are on the [Releases](https://github.com/genspark-ai/genoffice/releases) page.

## Apps

| App           | Product              | What it is                                                                                                                                                                                                                                                                                                                                                 |
| ------------- | -------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `apps/docs`   | **GenOffice Docs**   | `.docx` word processor. Byte-preserving round trip: only dirty paragraphs are regenerated (paragraph patch), everything else in the original file is kept byte-for-byte, so opening and saving never breaks layout in Word. Paginated view whose line metrics reproduce the original document's layout, tracked changes, comments, styles, equations, ink. |
| `apps/sheets` | **GenOffice Sheets** | `.xlsx` spreadsheet. UI built on the open-source [Univer](https://github.com/dream-num/univer) core (Apache-2.0) with a large layer of in-house extensions; xlsx import/export runs through an in-house Rust sidecar (calamine + IronCalc), charts are rendered in-house (Konva), plus pivot tables, slicers, conditional formatting, and formula tracing. |
| `apps/slides` | **GenOffice Slides** | `.pptx` presentations. In-house pptx parse/render/edit engine with masters, charts, cropping, ink, and text shaping (HarfBuzz metrics).                                                                                                                                                                                                                    |
| `apps/pdf`    | **GenOffice PDF**    | PDF viewer/editor on pdf.js + pdf-lib: annotations, forms, outlines, stamps, signatures, page operations, print.                                                                                                                                                                                                                                           |
| `apps/shell`  | **GenOffice**        | The suite shell: home screen, tabbed hosting of the four editors, auto-update.                                                                                                                                                                                                                                                                             |

Every app embeds the same AI panel: block-granular AI editing with version
snapshots and diffs in docs, a tool-calling agent over workbook/slide/PDF
state in the others. Docs and slides can also **generate images** with your own
key (`gpt-image-2` by default; set `providers.openai.imageModel` in the
settings file to use another) and insert them straight into the document.

**AI providers.** Three ways to power the AI panel: sign in to a Genspark
account and route calls through the Genspark service with no key of your own;
bring your own model API key (Claude, Gemini, DeepSeek, OpenAI — over its
Responses API, so the reasoning models can use tools with reasoning on — or any
OpenAI-compatible endpoint); or drive a coding-agent CLI you already have
installed and signed in — **Claude Code** or **Codex** — which needs no API key
and bills through that subscription rather than per token here.

Set one up from the gear in any AI panel (or the home screen's Settings menu);
the choice lives in `userData/ai-settings.json` and applies across all four
editors. Once more than one is configured, the picker in the composer footer
switches between them in a click — and a provider+model pairing can be saved
as a named **preset** ("fast", "deep") so two models from the same provider are
one click apart too. **Test** checks the key, model and endpoint with a
one-token request before you rely on them, and the refresh button next to the
model field re-reads the catalogue instead of a list baked into the build —
from the provider's API, or from the CLI itself for a CLI backend, so it lists
what that key or subscription can actually run today. Typing in the model field
filters the list. Keys are encrypted with the OS keychain where one is available and
are never handed to a renderer — the main process attaches the key when it
makes the request. See [SECURITY.md](SECURITY.md).

## Engine packages

All pure TypeScript, no Electron dependency, unit-tested (except the UI kit):

- `packages/docx-engine` — docx parsing → block tree (with `docxIndex`
  anchors and passthrough), OOXML fragment generation, byte-level paragraph
  patching.
- `packages/pptx-engine` / `packages/pptx-render` — pptx model and rendering.
- `packages/file-parse` — text extraction for AI attachments (office formats,
  text formats).
- `packages/agent-core` — the AI agent loop and skill composition shared by
  every app.
- `packages/ai-provider` — provider abstraction and streaming for the model
  backends.
- `packages/ai-cli` — locally installed agent CLIs (Claude Code, Codex) as AI
  backends: detection, model discovery, subprocess streaming, and the
  prompt-level tool-call protocol that lets an agent CLI drive the app's tools.
- `packages/ai-search` — Genspark auth + web/image search tools.
- `packages/i18n`, `packages/ui`, `packages/project-store`,
  `packages/electron-utils` — shared i18n core, React UI kit, recent-files
  store, and Electron main-process helpers.

## Development

```bash
npm install
npm run fixtures     # generate test .docx fixtures
npm test             # engine + app unit tests (docs/sheets/slides need no display)
npm run typecheck    # tsc --noEmit across every workspace
npm run dev          # all four editors + shell against Vite dev servers
npm run dev:docs     # a single app (same pattern works per workspace)
npm run dist:mac     # package macOS dmg (regenerates third-party notices)
npm run dist:win     # package Windows nsis installer
```

The sheets app additionally needs a Rust toolchain for its xlsx sidecar
(`cargo` on PATH); `npm run build -w @genoffice/sheets` compiles it
automatically.

Local UI/e2e driver scripts (Playwright + Electron, for local acceptance, not
committed by default) live in [`scripts/drivers/`](scripts/drivers/README.md).

## Architecture notes (docx round trip)

```
open docx ─► archive original by hash (never touched)
          ─► docx-engine parses word/document.xml top-level elements (w:p / w:tbl / …)
          ─► Block tree, each block anchored by docxIndex + original XML slice
          ─► TipTap streaming editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into original document.xml (untouched blocks keep original bytes)
          ─► repack zip; all other entries copied byte-for-byte
```

The same philosophy holds in sheets and slides: the original file is the
source of truth, edits are applied as narrow patches, and everything the
editor didn't touch survives the round trip untouched.

## Security

See [SECURITY.md](SECURITY.md) for the process security posture (renderer
sandboxing, IPC validation, external-link gating) and the threat models for
AI-generated content.

## Third-party notices

`npm run notices` regenerates the bundled third-party license summary
(`tools/gen-third-party-notices.mjs`); all runtime dependencies are
MIT/Apache-2.0/OFL, and the bundled fonts (Liberation, Carlito, Caladea, Noto
CJK subsets) are OFL/Apache.

## License

GenOffice is licensed under the [Apache License 2.0](LICENSE), with one
exception: the `ee/` directory is reserved for future enterprise modules and
is covered by the [GenOffice Enterprise License](ee/LICENSE).

The GenOffice and Genspark names and logos are trademarks of Mainfunc, Inc.
The Apache-2.0 license does not grant permission to use them (see section 6);
forks should use their own branding.
