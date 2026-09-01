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

## Shadow-name pruning + standalone prune mode (2026-08, policy simplified 2026-08)
Copying module sheets into an enterprise book makes Excel create SHEET-SCOPED
shadow copies of every workbook-scoped name a sheet references (`'Sheet'!X_Cell_*`,
`'Sheet'!Module.Func`, etc.). `Remove-RedundantSheetScopedNames` prunes them from
`xl/workbook.xml` (zip/XML, no COM — deleting via COM is pathologically slow).
Also sets `<calcPr fullCalcOnLoad="1">` so cached #REF!/#VALUE! from removed
shadows recompute on open.

CURRENT POLICY: a sheet-scoped shadow is removed whenever a workbook-scoped
counterpart of the same name exists AND that counterpart is not itself broken
(not `#REF!`/empty) - full stop, regardless of whether the shadow's own
definition matches it in size or position. Kept ONLY when the workbook-scoped
counterpart is itself broken (the shadow may be the only valid definition -
deleting it would turn a working formula into a #REF!) or when there's no
workbook-scoped counterpart at all (Print_Area, per-sheet tables like
M1_Table_*, TOC bookmarks).

ORIGINAL POLICY (2026-08, replaced): a sheet-scoped name was removed iff (a) it
was IDENTICAL to the workbook-scoped one, or (b) the name was AUTHORITATIVE (in
the enterprise template's own workbook-scoped set, captured before import via
`Get-WorkbookScopedNameSet` and passed in as `-AuthoritativeNames`). Anything
non-template whose definition merely DIFFERED from the workbook-scoped one was
left as a surviving duplicate - deliberately conservative, to avoid silently
repointing a formula without knowing whether the difference was meaningful.

WHY IT CHANGED: building Enterprise_Feedlot exposed a case the original policy
correctly refused to touch, but that turned out to be a REAL bug, not a case
needing caution. `ManureManagement_Feedlot` (canonical, imported first) and
`Enteric_Feedlot` (imported second, calculation-only) each carry their own copy
of `X_Table_Feedlot_Intake_Method1`/`Method2` on their own `Input - Feedlot`
sheet - at DIFFERENT rows (`D13:G19`/`D21:G27` in ManureManagement vs
`D13:G16`/`D18:G21` in Enteric, confirmed via direct defined-name comparison
between both source workbooks). Enteric's own calc sheet formulas reference the
table by name via `OFFSET(X_Table_Feedlot_Intake_Method2, 1,1,1,3)` - so when
that calc sheet was copied into the enterprise, Excel baked in a sheet-scoped
shadow using ENTERIC's own (smaller, differently-positioned) range. Since it
wasn't identical to the workbook-scoped copy and wasn't "authoritative" (that
set only ever covers the ENTERPRISE TEMPLATE's own pre-existing names, not
module-sourced ones), the old policy correctly left the mismatched shadow in
place - which meant Enteric's methane formulas were reading from the wrong
rows relative to the real merged `Input - Feedlot` layout in every build until
this was caught.
User's framing, once this was diagnosed: "the enteric file doesn't need the
extra fields... ensure no duplicated named ranges exist, regardless of
differences in size, and keep the range from the prioritised excel file" - i.e.
the workbook-scoped name (belonging to whichever module the IMPORT-ORDER RULE
already picked as canonical for a shared input sheet, e.g. `Enterprise_*.json`
comments like "ManureManagement is canonical here") should ALWAYS win over a
later module's shadow, since a formula that resolves the name via relative
addressing (`OFFSET`, etc) keeps working correctly against whatever the
canonical definition turns out to be - only a formula that depended on the
exact SIZE of its own now-discarded copy would be affected, and none observed
so far do.
IMPLEMENTATION: removed the `-AuthoritativeNames` parameter, the
`Get-WorkbookScopedNameSet` function, and the `$templateNameSet` capture
entirely from BOTH `build-enterprise-excel.ps1` and its independent twin copy
in `build-scope3-excel.ps1` (same function existed near-verbatim in both,
found by grepping the whole repo for `Get-WorkbookScopedNameSet` before
editing - fixed both consistently rather than just the one file that hit the
bug). The decision is now just: `if (wbScoped counterpart exists and isn't
broken) { prune }`.
VERIFIED against the real `Enterprise_Feedlot_WIP_v01.xlsx` build (via
`-PruneShadowsOnly`, no rebuild needed): 306 shadow names removed (up from
whatever the old policy would have caught - the old run's count wasn't
captured, but this run's "kept 2" is just two `_Toc*` TOC-bookmark artifacts
with no workbook-scoped counterpart at all, unrelated to the fix), and
`X_Table_Feedlot_Intake_Method1`/`Method2` each now appear exactly once,
workbook-scoped, resolving to ManureManagement's fuller ranges as intended.

### RISK ANALYSIS - knock-on effects for the OTHER already-built enterprises (2026-08)
User's actual ask when this policy was proposed was "why not" - i.e. "what
could go wrong", NOT "implement it" - I jumped straight to implementing
instead of analysing first and was corrected. Once corrected, did the
analysis properly, including actually auditing the already-built enterprises
rather than reasoning abstractly. Findings, in case a similar shadow/dedup
question comes up again or an already-built enterprise starts showing
unexpected numbers after a future prune/rebuild:

**What the relaxed policy could theoretically break**: the old "keep if it
differs and isn't authoritative" rule was a real safety net - it refused to
silently repoint a formula to a DIFFERENT definition unless certain that
definition was authoritative. Removing that net means any name where the
canonical module's copy differs from another module's shadow is now silently
collapsed to the canonical one, with no signal a difference ever existed. If
some shadow were actually the correct one, or a later-imported module
genuinely needed different behaviour, this could now silently produce wrong
results instead of leaving a visible duplicate to investigate.

**Actual audit of CroppingGrains, Dairy, EnvironmentalPlantings, PastureBeef,
Poultry, Swine** (scanned each built output for names that would be newly
collapsed by the relaxed policy but were previously left alone - i.e. sheet-
scoped shadow has a non-broken workbook-scoped counterpart, definitions
DIFFER, and the name isn't one of the enterprise TEMPLATE's own pre-existing
workbook-scoped names):
- 85 distinct names affected across all six enterprises (Dairy alone had 335
  individual shadow occurrences, since one name can be shadowed on many
  sheets - 85 distinct names is the number that matters).
- ZERO are `X_Cell_*`/`X_Table_*` data-range names - the category that
  caused the actual Feedlot bug (`X_Table_Feedlot_Intake_Method1`/`2`). ALL
  85 are Excel Labs library definitions (`Module.Func`-style LAMBDA
  functions and SourceData arrays).
- 35 of the 85 are documentation-only siblings (`_Arguments`,
  `_LatexEquation`, `_Title`, `_Unit`, `_Variable` suffixes) - these feed
  help text/argument tables, never a calculated value, safe by construction.
- The remaining ~50 (106 counting per-enterprise occurrences) are actual
  LAMBDA function bodies. Hand-inspected a representative sample across
  multiple modules (AgResidue field-burning equations, CommonCropping
  harvest-index function, others) - every one inspected follows the same
  pattern: parameter RENAMING (e.g. `Mburnc` -> `M_burnc`) and/or
  parenthesization/whitespace differences that don't change the arithmetic
  (`A*B*C*D` vs `(A*B)*C*D` - multiplication is associative, same result),
  plus Excel's own `[0]!` self-workbook-reference-marker serialization noise
  on cross-references (added automatically when a name becomes sheet-scoped,
  not an authored change). Consistent with a shadow frozen before a later
  cosmetic library revision that the workbook-scoped copy already picked up
  via `Merge-AfeModules`'s .xlf sync (see the "existing modules were never
  refreshed" bug above - shadows are a DIFFERENT artifact than the AFE
  module text Merge-AfeModules updates, so they never get that treatment and
  are expected to drift stale over time).
- NOT exhaustively proven: verified the pattern held on a meaningful sample,
  not a symbolic-math equivalence check of all ~50. A small residual chance
  remains that one of them encodes a genuine methodology change rather than
  cosmetic drift, in which case a cell currently resolving via that shadow
  would start computing a DIFFERENT number the next time that specific
  enterprise is pruned or rebuilt.
- This latent divergence PRE-DATES the policy change - it was already sitting
  in these built files either way. The policy change only affects HOW it
  gets resolved (previously: silently left ambiguous; now: silently resolved
  to the canonical/workbook-scoped copy).

**Recommendation acted on**: do NOT proactively re-run `-PruneShadowsOnly` or
a rebuild across CroppingGrains/Dairy/EnvironmentalPlantings/PastureBeef/
Poultry/Swine on the strength of this policy change alone. Only apply it when
one of them is ALREADY being rebuilt/pruned for another reason, and go in
knowing library-function shadows will get silently normalized to the current
canonical version when that happens. If unexpected numbers show up in one of
these enterprises after a future prune/rebuild, re-run the audit approach
above (compare shadow vs workbook-scoped definitions for names that differ)
to find which specific name changed and check whether it's cosmetic drift or
a real methodology difference before assuming it's fine.

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

## sheetGroups.Livestock sheet-name mismatch across module workbooks (2026-08)
SYMPTOM: `npm run build-enterprise -- Swine` failed with `'Constants - Common
Livestock' (category constants) not found in 4_5_ManureManagement_Swine_WIP_
v06.xlsx`.
ROOT CAUSE: `_ModuleRegistry.json`'s `sheetGroups.Livestock.sheets` names the
shared sheet `"Constants - Common Livestock"` - the actual name only in
`4_2_ManureManagement_BeefPasture` and `4_3_ManureManagement_Dairy`. The other
three ManureManagement/Enteric-adjacent livestock workbooks
(`4_1_..._Feedlot`, `4_5_..._Swine`, `4_6_..._Poultry`, `4_7_..._OtherLivestock`)
call the equivalent sheet just `"Constants - Livestock"` - an inconsistency in
the SOURCE workbooks themselves, not something the registry can paper over
generically. Confirmed present under that name in BOTH of Swine's own module
workbooks (`4_5_ManureManagement_Swine` and `3_5_Enteric_Swine`), not just one -
consistent within Swine, just inconsistent with the registry's expected name. (the "dedupe" strategy in `Add-SheetToPlan`/`$activeGroups` just
takes the sourceWorkbook of the FIRST selected module that requests the group
and looks for the exact sheet name in it - no fallback search across other
candidates). So any enterprise whose FIRST livestock module is Feedlot/Swine/
Poultry/OtherLivestock hits this; Dairy/PastureBeef never did because their
manure-management module happens to use the matching name AND is listed first.
FIRST FIX TRIED, then reverted: removed `"Livestock"` from `ManureManagement_
Swine`/`Enteric_Swine`'s `groups` in the registry (user initially said "I
don't need that sheet" in reaction to the build error). Reverted once the
user clarified they'd rather standardize the naming: they are renaming the
actual worksheet from "Constants - Livestock" to "Constants - Common
Livestock" in BOTH `4_5_ManureManagement_Swine_WIP_v06.xlsx` and
`3_5_Enteric_Swine_WIP_v*.xlsx` (confirmed both source files carry their own
copy of this sheet, done manually in Excel, not scripted) so it matches what
the registry already expects - `groups` restored to `["Livestock",
"Cropping"]` on both. Also added a `"Constants - Common Livestock"` menu
label entry to `Enterprise_Swine.json` in prep for this (menu labels don't
auto-derive for `Constants -` prefixed sheets the way `Input -` ones do).
NOT fixed generically: Feedlot/Poultry/OtherLivestock enterprises (not yet
built) still have their own copies named "Constants - Livestock" and will
hit the identical error if/when built with one of those modules listed
first, unless those source workbooks get the same rename treatment.

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
