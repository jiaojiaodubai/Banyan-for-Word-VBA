# Banyan for Word

Banyan for Word is a VBA add-in (`.dotm`) that brings Banyan citation, bibliography, refresh, conversion, finalize, chapter-break, and preference workflows into Microsoft Word. It is the **Word client** of the Banyan project; the companion [backend](https://github.com/jiaojiaodubai/Banyan) is a Zotero plugin that renders citations/bibliographies over a local HTTP API (Zotero port `23119`, debug `23124`). Install the Zotero plugin first - without it, live citation/bibliography workflows cannot complete.

## Table of Contents

- [Installation (end users)](#installation-end-users)
- [Using Banyan for Word](#using-banyan-for-word)
- [For developers](#for-developers)
  - [Repository layout](#repository-layout)
  - [VBA modules](#vba-modules)
  - [Development environment](#development-environment)
  - [Ribbon customization (icons and callbacks)](#ribbon-customization-icons-and-callbacks)
  - [Building the template (Import-BanyanDotm.ps1)](#building-the-template-import-banyandotmps1)
  - [Testing](#testing)
  - [Development rules](#development-rules)
  - [Encoding and template notes](#encoding-and-template-notes)

## Installation (end users)

### Requirements

- Microsoft Word 2016 or later on Windows or macOS, with macros enabled.
- Zotero with the Banyan backend plugin installed and enabled (see the [backend repository](https://github.com/jiaojiaodubai/Banyan)).

### Install Banyan for Word

1. Close Word.
2. Download the `Banyan.dotm` file (the checked-in `Banyan.dotm` in this repository is the release build).
3. Copy `Banyan.dotm` into Word's startup folder for your platform:

| Platform | Startup folder |
| --- | --- |
| Windows | `%APPDATA%\Microsoft\Word\STARTUP\` |
| macOS | `~/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word/` |

4. Start Word. If prompted, enable macros / trust this document.
5. The **Banyan** tab appears on the ribbon.

End-user notes:

- Because the file lives in the STARTUP folder, Word loads the add-in automatically on every start.
- The add-in only edits the open document (Banyan fields and preferences); it installs nothing else on your machine.
- Everything else in this repository (PowerShell tooling, VBA source, tests) is for developers - end users only need the `.dotm`.

### Update and uninstall

- **Update:** close Word, replace the `.dotm` in the startup folder, restart Word.
- **Uninstall:** close Word, delete `Banyan.dotm` from the startup folder, restart Word.

## Using Banyan for Word

| Button | Action |
| --- | --- |
| Insert/Edit Citation | Insert a new citation or edit an existing Banyan citation field. |
| Insert Break | Insert a ChapterBreak field and manage per-chapter overrides. |
| Insert/Edit Bibliography | Insert or edit a bibliography block. |
| Refresh | Re-render citations and bibliography from the backend. |
| Convert Zotero Fields | Convert Zotero ADDIN fields to Banyan fields. |
| Finalize | Backup, refresh, unlink citations/bibliography, and remove ChapterBreak fields. |
| Preferences | Open the settings form (citation style, fonts, and other options). |

Live features require Zotero with the Banyan plugin running.

## For developers

The rest of this file is for people building, extending, or testing the add-in from source.

### Repository layout

| Path | Purpose |
| --- | --- |
| `Banyan.dotm` | Built Word template/add-in artifact (release build; what end users install). |
| `Ribbon.xml` | Custom Ribbon UI definition; callbacks dispatch to `modRibbonCallbacks`. |
| `src/*.bas`, `src/*.cls`, `src/*.frm` (+ `.frx`) | Exported VBA source modules (see table below). |
| `test/test*.bas` | Per-module test suites (`testI10n`, `testJson`, `testDict`, `testHttp`, `testPreference`, `testField`). |
| `test/Run-Tests.ps1` | Full-suite dev gate: COM-injects all src + test modules and requires every test to pass. |
| `Import-BanyanDotm.ps1` | **Build script**: imports all VBA source into `Banyan.dotm` (release by default, `-DevMode` for a debug build). |
| `Verify-VbaImport.ps1` | Imports modules into a temporary Word document as a compile/import smoke test (src + all test modules). |
| `GenerateSettingsForm.ps1` | Rebuilds the settings UserForm through the VBIDE designer. |
| `Convert-FrmEncoding.ps1` | Converts UserForm text encoding for reliable VBE import. |
| `VbeEncodingHelpers.ps1` | Shared text/encoding helpers for the PowerShell tooling (including `Set-VbeDevModeConstant`). |

### VBA modules

| Module | Responsibility |
| --- | --- |
| `Dictionary.cls` | Cross-platform drop-in replacement for `Scripting.Dictionary` from VBA-tools/VBA-Dictionary. |
| `JsonConverter.bas` | Third-party VBA-JSON parser/serializer. Keep vendor code isolated. |
| `modJson.bas` | Banyan JSON adapter with error-safe parse/stringify helpers. |
| `modI10n.bas` | Lightweight localization registry and lookup. |
| `modShell.bas` | Shared macOS shell execution through `popen`. |
| `modHttp.bas` | Synchronous HTTP client for Windows and macOS. |
| `modProgress.bas` | Best-effort backend progress UI calls. |
| `modDiagnostics.bas` | Shared diagnostic message helpers (incl. the dev-mode re-raise policy). |
| `modField.bas` | Banyan ADDIN field creation, rendering, migration, and cleanup helpers. |
| `modChapterBreak.bas` | ChapterBreak ADDIN fields and current-chapter range lookup. |
| `modPreference.bas` | Global preferences in `BANYAN_PREF` CustomDocumentProperty plus chapter overrides in ChapterBreak fields. |
| `modCitation.bas` | Insert/edit citation workflow. |
| `modBibliography.bas` | Insert/edit bibliography workflow. |
| `modRefresh.bas` | Refresh citations and bibliography blocks. |
| `modConvert.bas` | Convert Zotero ADDIN fields to Banyan fields. |
| `modFinalize.bas` | Backup, refresh, unlink citations/bibliography, and remove ChapterBreak fields. |
| `modBusinessLogic.bas` | Settings command and shared command helpers. |
| `modRibbonCallbacks.bas` | Ribbon XML callback entry points. |
| `modTest.bas` | Concise in-editor test aggregator (pulls a smoke set from `test/*`) + the `DEV_MODE` switch used by debug builds. |

### Development environment

Prerequisites for building and testing this repository:

- **Microsoft Word 2016+ with VBA**, on Windows or macOS.
- **Trust access to the VBA project object model** enabled (Word Trust Center > Macro Settings). The build/test scripts drive the VBIDE through COM and fail without it.
- **PowerShell** (Windows PowerShell 5.1 or PowerShell 7) to run the scripts in this repo.
- **Zotero with the Banyan backend plugin** only if you need to exercise live routes manually - the automated tests never require it.
- **Office RibbonX Editor** only when editing the ribbon XML or its icons (see below).

### Ribbon customization (icons and callbacks)

- `Ribbon.xml` declares the `Banyan` tab. Callbacks (`onLoad="OnRibbonLoad"`, `getLabel`, `getVisible`, `onAction`) are handled in `modRibbonCallbacks.bas`.
- VBA never supplies images. Each `<button>` references an embedded image by name (`image="btnTest"` for the custom test icon, or built-in glyph names such as `format-quote-open`). **During development, inject and replace these image resources in `Banyan.dotm` with Office RibbonX Editor** (Custom UI Editor for Microsoft Office) - they live inside the `.dotm` as binary resources, not in the source tree.
- Typical icon workflow: edit `Ribbon.xml` -> open `Banyan.dotm` in Office RibbonX Editor -> replace the `customUI` part and (re)add icon parts whose names match the `image=` attributes -> save -> reload Word (re-run `Import-BanyanDotm.ps1` afterwards only to refresh the VBA side). There is no VBA image callback to modify.

### Building the template (`Import-BanyanDotm.ps1`)

`Import-BanyanDotm.ps1` is the **build script** for the add-in. It is used both for the everyday dev loop (re-import after a code change into the already-open template) and to produce the distributable `Banyan.dotm`. It:

1. Locates and opens `Banyan.dotm` (already-open template, explicit `-TemplatePath`, or common locations).
2. Removes stale components and the Microsoft Scripting Runtime reference.
3. Imports every `src` module plus the shipped concise test modules (`testI10n`, `testJson`, `testDict`) through the GBK temp-copy pipeline.
4. Saves the template.

Release build (default; what gets distributed):

```powershell
.\Import-BanyanDotm.ps1 -TemplatePath ".\Banyan.dotm" -CloseWord
```

Debug build (for development/testing):

```powershell
.\Import-BanyanDotm.ps1 -TemplatePath ".\Banyan.dotm" -DevMode -CloseWord
```

**Release vs Debug (DEV_MODE).** `src/modTest.bas` holds the `DEV_MODE` constant that (a) shows the `btnTest` ribbon button and (b) enables the `DiagnosticsReraiseIfDev` dev-mode error policy. It **defaults to `False` (release)** in the checked-in source, so a plain build is a release build. To avoid hand-editing a debug switch into committed code, the build script rewrites the constant at import time **only in the temporary copy it imports** (see the shared `Set-VbeDevModeConstant` helper in `VbeEncodingHelpers.ps1`): `-DevMode` produces `DEV_MODE = True`, a plain build keeps `False`. `src/modTest.bas` on disk is never modified. `test\Run-Tests.ps1` likewise always imports `modTest.bas` in Debug mode so the dev gate keeps the dev error policy active.

Other dev commands:

```powershell
# Compile/import smoke test (src + all test modules) without running tests
.\Verify-VbaImport.ps1

# Rebuild the settings UserForm layout from the designer
.\GenerateSettingsForm.ps1

# Manually normalize a UserForm's text encoding before a manual .frm import
.\Convert-FrmEncoding.ps1
```

### Testing

- Each Banyan module has a matching test module under `test/` (named `test*`, exposing `RunTests() -> String`).
- `modTest.RunAllTests()` (the dev test button, visible only in debug builds) aggregates a **concise smoke set** from `test/` (`testI10n`, `testJson`, `testDict`). It runs fast and does not mutate the document.
- **Dev gate:** after code changes, run the full suite:

```powershell
.\test\Run-Tests.ps1
```

  It COM-injects all `src` + `test` modules into a throwaway document, runs every `test*` module's `RunTests()`, and exits non-zero on any failure (the VBA equivalent of "lint + mocha" in a TS project). `testField` mutates the document, so it runs last. The gate imports `modTest.bas` with `DEV_MODE = True`, so dev-mode re-raises stay active and swallowed bugs fail loudly.
- **Import `.cls` class modules through the shared encoding pipeline - never import a raw source file
  directly.** A raw `VBComponents.Import` of `src/Dictionary.cls` can create a STANDARD module, so its
  `VERSION 1.0 CLASS` / `BEGIN` / `END` header lines become ordinary code and the whole project fails to
  compile ("Expected: statement"). Always: `Get-VbeImportText` (fallback UTF-8/ANSI decode) -> prepend
  `Attribute VB_Name = "..."` for `.bas` -> `Write-VbeImportText` a GBK temp copy -> `VBComponents.Import`
  the temp copy (this is exactly what `Run-Tests.ps1` and `Import-BanyanDotm.ps1` do). `.cls` files stay
  class modules; `.frm` files also need their sibling `.frx` next to the temp copy. Do not re-import
  `.cls`/`.frm` into an already-working template just to refresh a `.bas` change.
- `Import-BanyanDotm.ps1` ships the concise test modules (`testI10n`, `testJson`, `testDict`) with the
  packaged `.dotm`; the rest of `test/` is dev-only.
- **Test modules must never show dialogs** (`MsgBox`/`InputBox`): they block the COM automation flow in
  `Run-Tests.ps1` and can hang the dev gate. Report everything through the `RunTests()` string. The only
  `MsgBox` in the test path is `modTest.RunAllTests()`, which is for the manual Ribbon debug button.
- **Word instance cleanup (mandatory).** Every script that creates a Word instance through COM (the test
  runner, import/build scripts, ad-hoc probes) MUST fully quit it and verify no stray `WINWORD` process
  remains (`(Get-Process WINWORD).Count` must be 0). A `Quit()` can silently fail when a modal dialog
  blocks the instance, so scripts must also force-kill only the automation instances they spawned
  (compare WINWORD PIDs before/after creation; kill the `/Automation` leftovers). Lingering Word
  processes hold template/document locks and pollute the environment.
- **Automated tests must never call the GUI-triggering backend routes** `/banyan/style` and
  `/banyan/citation`: they open Zotero GUI dialogs (style picker / item selector) that require manual
  interaction and hang an unattended test run. Test the HTTP/request code with fabricated/mocked payloads
  instead - `testHttp`, `testPreference`, `testField` already build fake data locally and never touch the
  live backend. Live-route probing is manual-only and must require an explicit opt-in flag.

### Development rules

- Keep `Dictionary.cls` and `JsonConverter.bas` as vendor code. Put Banyan-specific behavior outside these files.
- **All Banyan dictionaries are `Dictionary.cls` instances and must be read through `modDict`.** See the "Dictionary safety" section below.
- Avoid public pass-through wrappers. A public procedure should either be a real external callback/entry point or own meaningful behavior.
- Keep compatibility wrappers only when an external contract requires them. Document the reason near the wrapper.
- Prefer direct Ribbon dispatch to the owning module's command entry point.
- Keep module-level comments useful: the module header explains ownership and public API; internal sections use `' --- Section ---`.
- Use ASCII punctuation in code comments unless Chinese UI strings or external text require otherwise.
- Do not hand-edit `frmSettings.frx`. Rebuild the UserForm through `GenerateSettingsForm.ps1`.

#### Dictionary safety (`modDict`)

`Dictionary.cls` (vba-tools/vba-dictionary) is used for macOS compatibility. It is declared
`VB_Exposed = True`, and when an instance is held in a **Variant**, feeding its default `Item` result
(`dict("key")`) into VBA intrinsics (`IsNull`, `IsEmpty`, `IsObject`, `VarType`) raises runtime error 458
("Automation type not supported") or misbehaves.

Rules:

- Use `modDict` helpers (`DictHasKey`, `DictKeyString`, `DictKeyLong`, `DictKeyBool`, `DictKeyValue`,
  `DictKeyObject`, `DictKeyIs*`, `DictIs*`) for all dictionary reads and type checks. They coerce to an
  `Object` local first, which is the proven-safe pattern.
- Do NOT write `IsNull(x("key"))`, `IsEmpty(x("key"))`, `IsObject(x("key"))`, or `VarType(x("key"))`
  where `x` may be a Variant holding a dictionary. Conversion reads (`CStr`/`CLng`/`CBool` on
  `x("key")`) and object self-checks (`TypeName(x)`, `x.Exists(k)`) are safe, but prefer the `modDict`
  helpers anyway for uniformity.
- **NEVER read `param("key")` through a `ByVal ... As Variant` parameter.** In this Word build that
  raises runtime error 5/91 (invalid call / object variable not set) even for plain scalar string
  values - the classic `CStr(x("key"))` conversion is only safe through `As Object` variables, NOT
  through a `ByVal Variant` parameter. Under `On Error Resume Next` the error is silently swallowed,
  so the value becomes `""` and gets written (this is the root cause of the empty style-id bug in
  `CloneStyle`/`CloneStyleForRefresh`). Always route through `DictKeyString(dict, key)` etc., which
  coerce internally - or change the parameter to `As Object`. Functions with `ByVal ... As Variant`
  params must never touch `param("key")` directly.
- **Object-valued members must be read with `Set`, never `Let`.** `v = dict("objKey")` (where the
  member holds a Collection/Dictionary) raises runtime error 450 ("Wrong number of arguments") through
  an `As Object` variable. Use `Set x = dict("objKey")` or `Set x = DictKeyObject(dict, key)`; guard
  with `IsObject(dict("objKey"))`. This bit `modRefresh.CopyDictionaryValue` and silently killed the
  citation refresh (field stayed at the `{ INTEXT_CITATION }` placeholder).
- `DictKeyValue` is a **scalar read** — do not `Let`-assign its result for object-valued keys. Its
  result is safe when passed by value into `DictIs*`/`DictKeyIs*` (which is how those helpers work).
- **Copy members between dictionaries with `DictCopyKey(target, targetKey, source, sourceKey)`** — the
  single canonical safe copy (Set for objects, Let for scalars). `modRefresh` uses it to build refresh
  contexts; do not hand-roll `v = source("key")` loops.
- **`modDict` helpers accept `Dictionary.cls` instances ONLY - never a VBA `Collection`.** A `Collection`
  is a different type with no `.Exists`/`.Keys`/`.Item(key)`; passing one to `DictKeyString`/`DictHasKey`/
  `DictKeyIs*`/etc. succeeds the `AsDict` coercion (it is an object) and then fails on `.Exists` with
  runtime error 438 ("object doesn't support this property or method") on every platform. Read VBA
  Collections with Collection syntax: `col("key")` or `CStr(col("key"))`. This exact bug hit
  `modBibliography.BibliographyAction`, which read its `mode` Collection (built by `MakeAddMode`/
  `MakeEditMode`) via `DictKeyString(mode, "type")` and 438'd on every bibliography insert (introduced
  by the 2026-08-23 modDict sweep) - fixed by `CStr(mode("type"))`. Watch for it whenever a helper that
  is *named* like a dictionary helper is applied to a variable declared `As Collection`.
- `Dictionary.cls` and `JsonConverter.bas` remain vendor code; `modDict.bas` is the Banyan-owned layer
  that contains all trap-avoiding access logic.

#### Dev-mode error policy (`DiagnosticsReraiseIfDev`)

VBA error handlers commonly swallow errors (empty `ErrHandler` or a fallback return), which makes
business bugs surface as "nothing happened" with no diagnostic. The `DEV_MODE` constant in `modTest.bas`
(defaults to `False`; flipped to `True` only when building with `Import-BanyanDotm.ps1 -DevMode` or when
the dev gate `Run-Tests.ps1` imports the modules) also drives a dev-mode error policy:

- `modDiagnostics.DiagnosticsReraiseIfDev` is the **first statement of every `ErrHandler`** in the
  business modules (`modField`, `modChapterBreak`, `modPreference`, `modCitation`, `modBibliography`,
  `modRefresh`, `modConvert`, `modFinalize`, `modBusinessLogic`, `modProgress`, `modShell`, `modDict`).
  Every call site passes its identity: `DiagnosticsReraiseIfDev "modX.FuncName"`.
- When `IsDevMode()` is `True`, it re-raises the current error (`Err.Raise Err.Number, ...`), so the
  error propagates to the caller / the VBE breaks instead of being swallowed. Handlers run deepest-
  first, so the FIRST re-raiser's identity is appended to the re-raised `Err.Description` as
  `[deepest: modX.FuncName]`; shallower re-raises see the marker already present and pass it through
  unchanged. The origin therefore survives to the VBE break / COM caller even though the top-level
  handler re-raises before it can show its dialog. Stateless - no reset needed.
- When `IsDevMode()` is `False` (release), it is a no-op and every fallback keeps its production
  behavior exactly as before.
- `modHttp` (transport retry/offline fallback) and `modTest`/`test*` (tests must report `[FAIL]`, not
  throw) are intentionally **not** instrumented.

Rules:

- New `ErrHandler` blocks in business modules should start with
  `DiagnosticsReraiseIfDev "modX.FuncName"` (the enclosing module + procedure) before any fallback
  assignment.
- Do **not** use an error handler as normal control flow for an *expected* condition. Expected states
  must be probed explicitly (e.g. `GetPrefFromDocumentProperty` checks for the property before reading
  it); otherwise dev mode turns a valid state into a throw. When a dev-mode run flags an expected state,
  refactor that lookup to an explicit check instead of removing the re-raise.
- To locate a swallowed bug: build a Debug template (`Import-BanyanDotm.ps1 -DevMode`) or run the dev
  gate, reproduce, and the VBE breaks at the re-raise with the original `Err.Number`/`Err.Description`;
  the dev gate (`Run-Tests.ps1`) then also fails loudly on the offending path.

### Encoding and template notes

`Import-BanyanDotm.ps1` converts source to GBK temporary files before VBE import so Chinese UI text survives zh-CN Office installations. `.bas` files in `src` intentionally omit exported `Attribute VB_Name` headers; the import script injects them into temporary copies.

`Convert-FrmEncoding.ps1`, `Import-BanyanDotm.ps1`, and `Verify-VbaImport.ps1` share `VbeEncodingHelpers.ps1` for text decoding, CRLF normalization, and `.frx` copying. `Convert-FrmEncoding.ps1` is the manual in-place normalizer; the other scripts use the same helpers only for temporary import files.

If template import hangs, close all Word windows and lingering `WINWORD.EXE` automation instances, then rerun the import command.

## Distribution

The release build of `Import-BanyanDotm.ps1` (no `-DevMode`) produces the `Banyan.dotm` that end users install into Word's STARTUP folder - see [Installation (end users)](#installation-end-users). Ribbon icons live inside the `.dotm` (injected with Office RibbonX Editor) and never ship in the VBA source.
