# build-enterprise-excel.ps1 — Excel Labs (AFE) module merge

`Merge-AfeModules` merges Excel Labs modules from the source/chapter workbooks
(`$resolvedModuleWorkbooks`) into the enterprise output (built from a template
in `Excel/Enterprises/*_Template_*.xlsx`). AFE file objects have `.path` and
`.text` (module code). The blob lives in `customXml/item*.xml` as base64
`<AFEJSONBlob>` (UTF-16/Unicode).

## Bug fixed (2026-07): existing modules were never refreshed
Original loop did `if ($existing.ContainsKey($mp)) { continue }` — it only ADDED
missing modules and skipped any already in the template. So updating a `.xlf`
(e.g. `Common_InputFunctions`), running `npm run build` (which syncs it into the
chapter workbooks), then `npm run build-enterprise` did NOT propagate the change
to the enterprise files, because the template already contained the module.

Fix: `$existing` now maps path -> file object; for a required module already
present, compare `.text` (LF-normalized) against the synced source and update it
in place when different. Write-back now triggers on additions OR updates.
`Merge-AfeModules` returns `[pscustomobject]@{ Added=@(); Updated=@() }`; summary
prints "Excel Labs modules add/upd". Verified: both Enterprise_PastureBeef and
Enterprise_EnvironmentalPlantings report `updated: /projects/Common_InputFunctions`.

Workflow to propagate an xlf change to enterprises: edit .xlf -> `npm run build`
(syncs chapters) -> `npm run build-enterprise`.

## Shadow-name pruning + standalone prune mode (2026-08)
Copying module sheets into an enterprise book makes Excel create SHEET-SCOPED
shadow copies of every workbook-scoped name a sheet references (`'Sheet'!X_Cell_*`,
`'Sheet'!Module.Func`, etc.). `Remove-RedundantSheetScopedNames` prunes them from
`xl/workbook.xml` (zip/XML, no COM — deleting via COM is pathologically slow). A
sheet-scoped name is removed iff it (a) is IDENTICAL to the workbook-scoped one, or
(b) the name is AUTHORITATIVE (in the template's workbook-scoped set from
`Get-WorkbookScopedNameSet`). Also sets `<calcPr fullCalcOnLoad="1">` so cached
#REF!/#VALUE! from removed shadows recompute on open.

These shadows surface as `Duplicate input cell/table name` warnings in
build-input-fields (it strips the `Sheet!` scope so shadow == workbook-scoped).
Fix without a full rebuild via `-PruneShadowsOnly` (npm `prune-enterprise-names`):
resolves the template + authoritative names as normal, then runs the prune on the
EXISTING output workbook and returns. Guards the template Copy-Item so the built
book isn't clobbered; supports `-DryRun`; flows through auto-discovery (no
`-EnterpriseId` = all enterprises). `-EnterpriseId Dairy|PastureBeef` maps to
`Enterprises/Enterprise_<Id>.json`.

## PastureBeef `templateWorkbook` misconfig -> #REF! never self-heals (2026-08)
SYMPTOM: after `npm run build-enterprise -- PastureBeef`, `'15 Scope 3'!D320:H320`
show `#REF! (in formula)`: `IFERROR(INDEX(#REF!, SEQUENCE(10)), "")` where the arg
used to be `Table_Input_Lime[Product]`. NOT present in the Scope3 source workbook
nor in Enterprise_CroppingGrains (same Scope3 source, same Fertiliser-before-Scope3
module order). Confirmed pre-existing (already in the last committed WIP file, not
a new regression) and deterministic (every rebuild reproduces it).

ROOT CAUSE (two layered bugs):
1. `Enterprises/Enterprise_PastureBeef.json`'s `templateWorkbook` pointed at
   `Enterprise_PastureBeef_Clean_WIP_v01.xlsx` (the Stage-4 clean-enterprise
   OUTPUT - a derivative of a past full build with input cells blanked) instead
   of the true bare `Enterprise_PastureBeef_Template_WIP_v01.xlsx`. Every OTHER
   enterprise (CroppingGrains/Dairy/EnvironmentalPlantings) correctly points at
   its bare `_Template_` file. FIXED: repointed to the Template file (2026-08-25).
   IMPACT: because the Clean file already has all 42 module sheets baked in, the
   copy-plan loop (`if ($targetSheetNames.Contains($entry.Name)) { ...skip... }`
   at build-enterprise-excel.ps1 ~line 810) always sees every sheet as "already
   present" and never re-imports anything from the true module source workbooks.
   Confirmed via a real run: "Sheets imported: 0 / Sheets already present: 42".
   So PastureBeef has been running in incremental-refresh-only mode (AFE
   module/name updates only) the whole time, NOT the "full deterministic
   rebuild" the pipeline promises - whatever got corrupted in the Clean file at
   any point in the past (this Lime #REF!) silently perpetuates forever, since
   the true source is never re-consulted. Re-running build-enterprise after the
   config fix will NOT retroactively repair the current WIP output - it repairs
   FUTURE builds. The current file still needs either a genuine full rebuild
   (copy fresh from the now-correct Template) or a direct cell-level patch.
2. Independent, compounding bug in the sheet-copy localisation pass itself
   (build-enterprise-excel.ps1 ~line 977-1004, "Localise externalised
   references"): when a genuine fresh copy DOES happen, cross-workbook
   STRUCTURED TABLE refs (`Table_X[Col]`) to a table that lives on a sheet not
   included in that particular Worksheet.Copy() get externalised by Excel while
   the source workbook is still open (sources aren't closed until after save).
   In that state Excel's live `.Formula2` text for a table ref uses an
   UNQUOTED qualifier when the filename has no spaces, e.g.
   `13_Scope3_WIP_v14.xlsx!Table_Input_Lime[Product]` - confirmed via an
   isolated repro (Workbooks.Add() + Worksheet.Copy(), no repo files touched).
   The old `$reTableRef`/`$reExternalCell` regexes required a LEADING QUOTE
   (`'...xlsx'!Name[`), so this unquoted form was silently never localised -
   even though the identically-named table already existed locally (imported
   earlier by the Fertiliser module, which is correctly ordered before Scope3
   in both enterprises' `modules[]`). Left unlocalised, the later unconditional
   `$target.BreakLink(...)` call (~line 1164, meant to make the workbook
   self-contained) can't cleanly convert this multi-cell CSE/array-filled
   `INDEX(Table[...], SEQUENCE(10))` block to a flat value, and bakes a literal
   `#REF!` into the surviving formula text in place of just the table token -
   exactly matching the observed bug. Verified with a real-files repro (actual
   5_Fertiliser_WIP_v10.xlsx + 13_Scope3_WIP_v14.xlsx, replicating the real
   copy+localise sequence) that the widened regex correctly resolves D320 to
   "Lime product 1". FIXED: `$reTableRef`/`$reExternalCell` now also match the
   unquoted `filename.xls*!Name[` form (anchored to plain filename chars only -
   `[A-Za-z0-9_-]+\.xls[a-z]*`, with a negative lookbehind so it can't swallow
   preceding formula text like `IFERROR(INDEX(13_Foo.xlsx`).
   Sibling tables on the same "Input - Fertiliser" sheet
   (InorganicFertiliser_Brands/_Custom) did NOT show this corruption - they
   apparently bind to the local table at copy time and never externalise in the
   first place. The asymmetry is Excel-COM-internals-dependent (plausibly
   related to PastureBeef having more modules/sheets ahead of Scope3 in its
   import plan than CroppingGrains - 8 vs 6 - affecting COM binding timing) and
   wasn't fully pinned down; the regex fix is a safe superset (handles the new
   unquoted case in addition to the already-working quoted case) regardless of
   the exact trigger.
DIAGNOSIS METHOD: static XML/zip inspection only (no COM) to compare D320
across the Scope3 source, CroppingGrains, and PastureBeef outputs; then
isolated Excel-COM repros in scratch files (never touching repo files) to
reproduce the exact externalisation + BreakLink corruption mechanism before
patching the real script.

## IMPORT-ORDER RULE for shared/duplicate input sheets (2026-08)
When two modules ship the SAME input sheet + named ranges (Enteric vs
ManureManagement both have 'Input - Dairy' / 'Input - Pasture Beef'), the
FIRST module listed in config `modules[]` claims the workbook-scoped
X_Cell_*/X_Table_* names and keeps its sheet as the unrenamed master. So the
CANONICAL module (the one with the full/correct input set = ManureManagement)
MUST be listed FIRST; the duplicate (Enteric) is listed AFTER with
`renameSheets` (e.g. 'Input - Dairy' -> 'Input - Dairy (Enteric)').
- Symptom of getting it wrong: build-input-fields floods `Duplicate input
  cell/table name` warnings AND the workbook-scoped name binds to the wrong
  (subset/renamed) sheet. Probing the pruned Dairy book showed [WB] name ->
  'Input - Dairy (Enteric)' while 10-14 sheet-scoped shadows -> master
  'Input - Dairy'. Fix = reorder so Manure is first (done for Dairy +
  PastureBeef). Requires a full `npm run build-enterprise` to take effect;
  prune-enterprise-names cannot fix it (shadows differ from the wb-scoped copy
  and aren't in the template's authoritative set, so the prune correctly leaves
  them).
- Open follow-up: shadows also disagree with each other (e.g.
  X_Cell_Dairy_MilkProductionUnit -> $E$49 vs $E$50), i.e. the Enteric and
  Manure 'Input - Dairy' sheets have a row-offset drift. Verify the module
  workbooks' input rows line up after the reorder rebuild.
