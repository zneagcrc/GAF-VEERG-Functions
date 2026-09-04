# Full enterprise build pipeline (module .xlf -> final add-in-free Excel)

Repeatable order to go from sources of truth to the final clean, distributable
enterprise workbook.

## Stage 0 — Pre-flight (RUN BEFORE EVERY BUILD STAGE 2-5, not just once)
Repeat this check ahead of each Excel-touching stage below (build, build-enterprise,
clean-enterprise, expand-lambda). Any relevant workbook left open fails that stage's
pre-flight lock check.
- Close ALL workbooks in Excel that the build touches (`Excel/*.xlsx`,
  `Excel/Enterprises/*.xlsx`) or the pre-flight lock check (`scripts/file-access.ps1`)
  aborts the build. A locked EXCEL process with `MainWindowHandle != 0` = user has it open.
- Never kill EXCEL mid-build (throws `RPC_E_DISCONNECTED 0x80010108`). When NO build is
  running, clear leftover HEADLESS COM servers only:
  `Get-Process EXCEL | ? { $_.MainWindowHandle -eq 0 } | Stop-Process -Force`.

## Stage 1 — Edit sources of truth
- Equations / input functions / source data -> the `.xlf` files.
- Input sheet layout / templates -> module `.xlsx` under `Excel/` + enterprise
  `Excel/Enterprises/*_Template_*.xlsx`.
- Enterprise composition -> `Enterprises/Enterprise_<Id>.json` (+ `_ModuleRegistry.json`).

## Stage 2 — `npm run build` (build.ps1)
Syncs every `.xlf` into the Excel Labs (AFE) modules in all `Excel/*.xlsx`, propagates
shared common sheets, then refreshes `generated-sourcedata/*.json` + `InputFields/*.json`.
`-- -Menu` also rebuilds nav menus. Dry: `npm run build:dry`.

## Stage 3 — `npm run build-enterprise [-- <Id>]`
Copies template fresh, imports module sheets, merges/updates AFE modules, prunes
sheet-scoped shadow names, restores `VEERG_*_Result_Method*` names from source
(see enterprise-result-names-restore.md), sets `fullCalcOnLoad`. Omit `<Id>` =
build ALL `Enterprise_*.json`. Output: `Excel/Enterprises/Enterprise_<Id>_WIP_v01.xlsx`.
Dry: `:dry`. Shadow-only fix: `npm run prune-enterprise-names`.

### Resumable / incremental "build all" (2026-09)
Auto-discovery mode (no `<Id>`) now:
- runs **each enterprise in its own `powershell` process** — a crash / OOM in one
  no longer aborts the rest; prints a per-enterprise summary (built / skipped /
  FAILED + duration) and exits 1 with a `-- -Only <ids>` retry hint if any failed.
- **freshness-skips** an enterprise whose `_WIP_v01.xlsx` is newer than its config,
  `_ModuleRegistry.json`, template, every resolved module source workbook, AND the
  build scripts (`build-enterprise-excel.ps1` + everything it dot-sources). Any
  build-script change therefore invalidates ALL outputs (intentional; use `-Only`
  to test one).
- flags: `-Only Swine,Poultry` (subset), `-From Poultry` (resume from, discovery
  order), `-Skip Dairy,Feedlot`, `-Force` (ignore freshness). `-DryRun` and
  explicitly-named (`-Only`/`-From`) enterprises always build.
- npm passes `-Only Swine,Poultry` as ONE literal token (not a PS array) — the
  script splits on `[,;\s]`.

`build-scope3` got the same freshness gate (+ `-Force`); it opens no workbook when
the output is already up to date.

### Memory: prefer batches over "build all" (2026-09, confirmed by user)
The single-process `& $PSCommandPath` loop (pre-2026-09) built all 7 enterprises
in ONE PowerShell process, accumulating leaked Excel RCWs -> the process ran out
of memory and crashed near the last enterprise (observed: run completed all 7
but the terminal died at/after `Enterprise_Swine`). The 2026-09 per-enterprise
child-process model + a headless-EXCEL sweep between iterations removes the
accumulation, BUT on a memory-constrained machine the user still prefers NOT to
run all at once. Recommended: `npm run build-enterprise -- -Only <id>` (one or a
few), or `-From <id>` to resume, rather than a bare `build-enterprise`. The
freshness gate already means a bare run after a small change only rebuilds the
stale ones, so "all 7" is rare in practice anyway.

## Stage 4 — `npm run clean-enterprise [-- <Id>]`
Writes a SEPARATE `…_Clean_…​.xlsx` with every input cell/table field blanked (formulas
untouched). Source workbook never modified. WEB-APP DELIVERABLE: the conversation input
web app consumes the CLEANED enterprise file + the enterprise InputFields JSON (NOT the
WIP build). So clean-enterprise + build:input-fields are on the web-app delivery path, not
optional. clean-enterprise reads the OVERRIDE-MERGED InputFields JSON, so it must run
AFTER: build:input-fields -> define InputFields/_overrides/<Enterprise>.json ->
build:input-fields (re-merge) -> clean-enterprise.

## Stage 5 — `npm run expand-lambda-functions -- <path>`  (runs on the STAGE-4 _Clean file)
Inlines all `Module.Func()` / SourceData LAMBDA references -> add-in-free
`…_Clean_…​_expanded.xlsx`. Example:
`npm run expand-lambda-functions -- .\Excel\Enterprises\Enterprise_CroppingGrains_Clean_WIP_v01.xlsx`.
WEB-APP DELIVERABLE (confirmed): the web app loads the ADD-IN-FREE `_Clean_…_expanded.xlsx`
(NOT the plain `_Clean`), because it evaluates Excel WITHOUT the Excel Labs AFE add-in. So
the two final web-app artifacts are: `Enterprise_<Id>_Clean_WIP_v01_expanded.xlsx` + the
enterprise InputFields JSON. Dry: `:dry -- <path>`; debug failed writes: `:debug -- <path>`.

## Optional / supporting
- `npm run build-enterprise-testinput` -> `Test/Enterprises/TestInput_Enterprise_<Id>.json`.
- `npm run build-scope3`.
- `npm run find-errors` (scan built workbooks for #REF!/#VALUE! etc.).

## Ordering rule that bites
Two modules sharing an input sheet: list the CANONICAL module FIRST in config `modules[]`;
the duplicate AFTER with `renameSheets`. Wrong order floods "Duplicate input cell/table
name" warnings; needs a full build-enterprise (prune can't fix it).
