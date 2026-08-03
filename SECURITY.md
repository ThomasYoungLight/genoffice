# Security Policy

## Reporting a Vulnerability

Please report suspected vulnerabilities privately via GitHub's
[private vulnerability reporting](https://github.com/genspark-ai/genoffice/security/advisories/new)
on this repository. Do not open public issues for security reports. We aim to
acknowledge reports within 72 hours.

## Process Security Posture

All application windows run with the full Electron renderer lockdown:

- `contextIsolation: true`, `nodeIntegration: false`, `sandbox: true` for every
  document window and tab view (docs, sheets, slides, pdf, shell, updater).
- Renderers reach the main process only through typed, validated IPC channels
  (payloads are schema-checked in the main process; sheets uses zod end to end).
- Every `shell.openExternal` call goes through a single shared gate
  (`@genoffice/electron-utils` → `safeExternalUrl`) that parses the URL and
  enforces a protocol allowlist (http/https; pdf link annotations additionally
  allow mailto). `file:`, `javascript:`, and custom schemes are always rejected.
- No API keys are hardcoded. AI requests are proxied through the signed-in
  Genspark account by default; a user-supplied model API key is stored in
  `userData/ai-settings.json`, encrypted through Electron `safeStorage` (the
  macOS Keychain / Windows DPAPI) wherever the OS offers it. On a system with
  no keyring the value is written readable and marked as such in the file.
- A model API key never enters a renderer. `ai:get-settings` blanks every key
  and reports only whether one is stored; the main process fills the real key
  in when it issues the request. Renderers host document content, which is
  prompt-injectable, so they are kept out of the key's blast radius.
- A renderer never chooses the endpoint a stored key is sent to. It selects the
  provider and model; the base URL always comes from the settings file
  (`AiSettingsStore.configFor`). The one exception is a key typed into the
  settings dialog and not yet saved, which is probed against the endpoint typed
  alongside it — there nothing stored is at risk. Without that rule, pointing
  the "custom" provider at an attacker-controlled base URL would post the saved
  key to it.

Generated images follow the same custody rule: the prompt goes to the provider,
the bytes come back to the main process, and they are written to a temp file
behind an opaque marker that only the process which issued it can redeem. An
insert asked for anything else still goes through the SSRF-guarded download —
a model-supplied `file:` or private-network URL is refused as before.

## Local Agent CLI Backends

Selecting Claude Code or Codex as the backend makes the main process spawn that
CLI. Two deliberate constraints:

- **Their own tools are off.** Claude Code runs with `--tools ""` and Codex with
  `--sandbox read-only`, so neither can touch the filesystem or run shell
  commands on the user's behalf. They are used purely as text generators; the
  only edits that happen are the app's own tools, applied through the same
  command pipeline as a manual edit.
- **They run outside the project.** The subprocess starts in the home directory
  rather than the app's working directory, so it does not silently pick up a
  `CLAUDE.md` / `AGENTS.md` from wherever the app happens to be launched.
- **Model discovery is a read-only query.** Refreshing the model list starts the
  CLI in its machine-protocol mode (Claude Code's control channel, Codex's
  app-server) and asks for the catalogue its own picker shows. No turn is
  started, no prompt or document content is sent, and the subprocess is killed
  as soon as the answer arrives.

The conversation is passed on stdin, never as argv, and the CLI's own login is
the credential — this app neither reads nor stores one for these backends.

Note that document content is prompt-injectable and is sent to the CLI like any
other backend. The tool-call contract is carried in the prompt (the CLIs accept
no caller-supplied tool definitions), so a call is only executed when its name
matches a tool the app offered for that turn; anything else is discarded.

## Threat Model: AI-Generated Layout Scripts (slides)

The slides AI can adjust slide layouts by emitting a small script that is
parsed with Acorn and evaluated by a constrained AST interpreter
(`apps/slides/src/renderer/ai/layout-script-interpreter.ts`). The source looks
like a small, synchronous subset of JavaScript for model compatibility, but it
is not passed to `eval`, `Function`, a VM context, a worker, or the JavaScript
engine as executable source.

**What the script can do by design:** read prototype-free JSON copies of
`els`/`canvas`, perform bounded arithmetic/control flow, use explicitly
implemented string/array/regular-expression/Math helpers, and call
`setBox/moveBy/resizeBy/setText/setStyle/setFill/setStroke/log`. Every edit
primitive validates its arguments (element existence, read-only flags, finite
numbers, hex colors) and writes only into an op buffer that is applied through
the same command pipeline as manual edits.

**Interpreter boundary:**

1. Identifiers resolve only in interpreter-owned lexical scopes seeded with the
   documented data and callables. There are no ambient globals, module loader,
   DOM, network, IPC bridge, timers, process APIs, or dynamic code primitives.
2. Property reads are dispatched by value type. Data objects expose own JSON
   fields only; arrays, strings, and regexes expose a small method allowlist.
   Host prototypes and function properties are never traversed, including
   through computed property names.
3. Calls accept only interpreter-created functions or explicit builtins. A host
   function obtained through a constructor/prototype chain cannot be
   represented.
4. Inputs and values crossing into edit primitives are recursively copied as
   JSON-like, prototype-free data. Errors discard all buffered operations;
   logs are capped.
5. Execution has statement/expression and call-depth limits to bound runaway
   loops or recursion.

The Electron renderer sandbox remains defense in depth, but it is not the
layout-script security boundary. The interpreter is designed so a layout
script cannot obtain renderer capabilities in the first place.

If you find a way for a layout script to reach anything beyond the injected
primitives (network, storage, IPC channels not reachable by design, or the
main process), that is a vulnerability — please report it.

## Threat Model: Rendering AI-Generated HTML (slides export)

The HTML-to-pptx export pipeline renders AI-generated HTML in a hidden
`BrowserWindow`. That window is treated as hostile content: full renderer
lockdown (`sandbox: true`, `contextIsolation: true`, `nodeIntegration: false`),
no preload script, no IPC surface — the main process drives it exclusively
through `executeJavaScript` and destroys it under a watchdog timeout.

## Out of Scope

- The cloud AI services this client talks to are operated separately and are
  not part of this repository; issues with them should be reported through the
  service provider's channels.
- Vulnerabilities that require an already-compromised machine or a modified
  binary. This includes the deliberate environment-variable override points
  for local development (`GSK_CLI_PATH`, `XLSX_SIDECAR_PATH`): setting them
  requires control of the process environment, which is equivalent to code
  execution on the machine.
