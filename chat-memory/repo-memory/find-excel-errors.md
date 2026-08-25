# find-excel-errors.ps1 (npm run find-errors)

Single read-only checker consolidating error detection across the pipeline.
Scans Excel/*.xlsx + Excel/Enterprises/*.xlsx (excludes ~$, _expanded,
_template, _clean, .bak). Reads .xlsx package XML directly — NO COM, NO Excel, NO
add-in (deliberate: a headless recalc would falsely report #NAME? for every VEERG.*
LAMBDA since the add-in isn't loaded). Reports CACHED saved values (see the
fullCalcOnLoad handling below for how provisional cached errors are treated).

Categories: [cell] (cached t="e" + formula-text error tokens), [name] (defined
names with error/empty RefersTo), [link] (external links), [sheet] (formula /
defined-name RefersTo referencing a sheet the workbook lacks — static, catches
#REF! a cached scan misses), [sum] (SUM literal-range size vs actual data block).

Categories:
- [cell]  Two forms: (1) CACHED error = <c t="e"> cells (currently evaluate to an
          error); (2) FORMULA error = a cell whose stored <f> text contains an
          error token (e.g. #REF! left in args) even though the cached value is
          valid because a function swallowed the bad ref — shown as
          "#REF! (in formula)". Both give sheet, addr, token, stored formula.
          SKIPS cells with a `vm` (value-metadata) attribute (t="e" only): those
          are RICH VALUES (modern "Place in Cell" images backed by xl/richData/*,
          linked data types, etc.) whose t="e"/#VALUE! is only fallback text, not
          a real error. The VEERG workbooks have a pasted logo placed in A1 of
          most sheets this way (<c r="A1" t="e" vm="1"><v>#VALUE!</v></c>).
          NOTE: the formula-text scan is what catches latent #REF! that a wrapper
          function masks, e.g. CommonCropping_GetFracWET(I16,I18,#REF!,#REF!,...)
          returning 1 (found on the Poultry leach/runoff sheets, 204 across 11
          workbooks). Fast-skip gate now also checks ErrorRegex over sheet text,
          not just t="e".
- [name]  defined names in xl/workbook.xml whose RefersTo has an error token
          (dangling = plain ref, broken-fn = LAMBDA with internal #REF!) or is
          empty. Sheet vs workbook scope resolved via localSheetId → <sheets>.
- [link]  external links from xl/externalLinks/_rels/*.rels targets.

Params: -RepoRoot, -WorkbookPath (single file), -Max (per-category cap, default
50, 0=unlimited), -IncludeLibraryFunctions (also show broken-fn .xlf library
LAMBDAs; HIDDEN by default as known noise), -SkipSumCheck, -IncludeCachedErrors
(list provisional cached errors hidden under fullCalcOnLoad — see below),
-FailOnError (exit 1 if any shown [cell]/[name] found).

Baseline run 2026-08-03: 493 cell errors, 746 broken names, 0 external links
across all 28 workbooks. Most broken names are the .xlf library LAMBDAs with
internal #REF! (broken-fn — known WIP, audit-names never auto-deletes them) plus
dangling #REF! cruft (audit-names:commit removes). Cell #NAME?/#VALUE! are mostly
stale cached values that clear on a real recalc in Excel-with-add-in.

FULLCALCONLOAD CACHED-ERROR HANDLING - ITERATION HISTORY (2026-08): tried (1)
suppress ALL cached under fullCalcOnLoad -> missed genuine J584; (2) narrow to
#REF!/#VALUE! -> still missed J584; (3) report ALL -> false positives on transient
Input-Site #REF!. SUPERSEDED by the HIDDEN-BY-DEFAULT + OPT-IN design below.

FULLCALCONLOAD CACHED ERRORS = HIDDEN-BY-DEFAULT + OPT-IN (2026-08, FINAL): cached
(t="e") cell errors in a `fullCalcOnLoad="1"` workbook are PROVISIONAL - some are
permanent (genuine #REF! from a missing sheet, e.g. Enterprise `'15 Scope 3'!J584
=VEERG_5_2_1_1__1_...Result_Method1`, `G590 =E588`), some are TRANSIENT and clear on
Excel's on-open recalc (e.g. 13_Scope3 `'Input - Site'!E13
=DATE(YEAR(X_Cell_Site_StartDate)...` - the names resolve fine, cached #REF! is stale).
find-errors reads PRE-recalc cache so it CANNOT tell them apart from the token alone.
Iterated: (1) suppress all under fullCalcOnLoad -> missed J584; (2) report all ->
false positives on Input-Site. FINAL DESIGN: under fullCalcOnLoad HIDE cached errors
by DEFAULT (no false positives) but COUNT them and print a note ("N provisional
cached cell error(s) hidden ... pass -IncludeCachedErrors to list"); new switch
`-IncludeCachedErrors` lists them for a deep hunt. Structural errors ALWAYS show:
#REF! in formula TEXT, broken [name]s, and the [sheet] missing-sheet check. Impl:
Get-CellErrors tags each object `IsCached`; main loop `$shownCellErrors = @(if
($fullCalcOnLoad -and -not $IncludeCachedErrors) { $cellErrors | ? { -not $_.IsCached } }
else { $cellErrors })`. GOTCHA: `$x = if(..){@()}else{..}` yields $NULL when the
branch is an EMPTY array (empty array enumerates to nothing) -> `.Count` throws under
StrictMode; wrap the WHOLE if in @() (not the inner branch). Non-fullCalcOnLoad
workbooks (raw module books) still report cached errors normally.

NEW [sheet] CATEGORY - references to a missing sheet (2026-08): `Get-MissingSheetReferences`
STATICALLY flags formulas / defined-name RefersTo that reference a `'Sheet'!` the
workbook doesn't contain (guaranteed #REF! on recalc even when the cached value is a
stale-good number a cached scan can't see). Strips double-quoted strings first
(INDIRECT("Sheet!..")), skips external `[N]` book refs, handles 3D `'S1:S2'` and
`''` escapes. `$script:SheetQualifierRegex` g1=quoted g2=unquoted, `#` in the
unquoted lookbehind stops error tokens matching. GOTCHA: the `$findMissing`
scriptblock returns a List -> PowerShell UNROLLS it; wrap call sites in @() or
`.Count` throws under StrictMode ("property Count cannot be found"). NOTE: did NOT
catch J584 (that ref is a defined NAME/same-sheet cell, not a `'Sheet'!` qualifier -
caught instead by the cached-#REF! report above); the [sheet] check is complementary
coverage for DIRECT sheet-qualifier refs.

[sum] FALSE-POSITIVE FIX (2026-08): the data-map builder (Build-SheetCellMaps)
used to exclude ANY cell whose formula merely contained SUM/SUBTOTAL/AGGREGATE
(old $script:AggregateFnRegex) as a "total" cell. This wrongly dropped VALUE
cells that sum a STRUCTURED TABLE REF, e.g. Input - Enterprise E67:E71 =
IF(...,SUM(Table_Input_Enterprise_PlantingRemovals[...]),"Enter value"). Those
cells fell out of the Numeric map, so =SUM(E67:E71) in E72 reported "data spans
E68" (only E68 survived, as an "Enter value" placeholder). FIX: replaced with
$script:AggregateOfPlainRangeRegex which only excludes a cell when it aggregates
a PLAIN A1 range (SUM(D10:D25) / SUBTOTAL(9,D10:D25) - optional leading \d+, for
the fn-number arg), NOT SUM(Table[Col]). Preserves block-total protection while
counting table-sum value cells as data.

[sum] FALSE POSITIVE FROM STALE CACHE - shared-formula text column miscounted as
data (2026-08): reported as `'3.2.1.1-2 Enteric Methane'!V162  SUM range
V149:AE160 but data spans U149:AE160` in a freshly-built (not-yet-recalculated)
Enterprise_PastureBeef. Column U (immediately left of the summed V:AE block) is
a SHARED FORMULA `<f t="shared" ref="U149:U160" si="3">U133</f>` that renders
Jan-Dec as TEXT. Under fullCalcOnLoad the anchor cell (U149) still cached valid
text, but its shared-formula siblings (U150:U160) cached stale `t="e"` #VALUE!
errors from before the last build's reference fixes landed. Build-SheetCellMaps
counts ANY non-text cell with a value as Numeric, including cached errors, so
the erroring "month name" cells got counted as data and the walk swallowed the
label column. Diagnosed but NOT special-cased in the checker: blanket-excluding
`t="e"` cells breaks far more than it fixes, because the genuine numeric block
(V:AE) was ALSO all cached #VALUE! at the time (same stale-cache root cause,
different cells) via individual per-cell `t="array"` formulas - excluding all
error cells would leave Resolve-AxisRange with zero Numeric evidence anywhere
and produce a degenerate "actual" range, worse than the current false positive.
A shared-formula-group heuristic (trust the group's non-error member) was
designed but NOT implemented - REJECTED in favour of fixing the cache instead
of teaching the checker to guess around it: [cell] already discounts provisional
cached errors under fullCalcOnLoad; [sum] has no equivalent safeguard and reads
raw pre-recalc XML, so its results are only trustworthy AFTER a real recalc.
FIX = workflow, not code: run find-errors only after opening the target
workbook(s) in Excel WITH THE EXCEL LABS ADD-IN LOADED, forcing a full recalc
(Ctrl+Alt+F9), and saving - confirmed this clears the false positive (PastureBeef
"full of errors" on open, F9 recalc -> "all is well"). Do NOT headless-recalc as
a substitute (no add-in -> #NAME? baked into every Module.Func() cell, see
build-scope3-excel.md "Critical lesson"). find-excel-errors.ps1 now
Write-Warnings when a fullCalcOnLoad workbook has [sum] mismatches, reminding to
recalc+save and re-run before trusting the results.

PS5.1 gotchas hit while building it:
- ZipArchiveMode lives in System.IO.Compression; ZipFile in
  System.IO.Compression.FileSystem — Add-Type BOTH assemblies.
- Returning an XmlDocument or XmlNamespaceManager from a function → PowerShell
  auto-enumerates it (doc→child nodes, nsmgr→prefix strings) so the caller gets
  an Object[]. Wrap the return with the comma operator: `return ,$doc` /
  `return ,$ns`.
- Dot-sourcing the script sets Set-StrictMode -Version Latest in the interactive
  session, which breaks the npm.ps1 wrapper ($MyInvocation.Statement). Run it as
  a child process (powershell -File ...) or Set-StrictMode -Off after.
