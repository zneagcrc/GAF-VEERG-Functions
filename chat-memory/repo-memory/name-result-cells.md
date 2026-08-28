# scripts/name-result-cells.ps1 (npm `name-result-cells[:commit]`)

Names the top-level Method 1/Method 2 result cell for every calculation
sheet that has one but isn't published as a `VEERG_..._Result_MethodN` name
yet - built after a gap was found by hand: `4_5_ManureManagement_Swine_WIP_
v06.xlsx` had a Result name for only 1 of its 5 calc sheets (see
enterprise-template-authoring.md / list-enterprise-results.ps1's unit-lookup
history for how that surfaced).

## The mechanism (source of the design, all from the user, not derived)
Every module workbook has a 'Results' sheet with a table headed "GHG
emissions method 1" / "GHG emissions method 2", one row per calc sheet. When
a sheet's result IS already named, that column holds a bare name reference
(`=VEERG_..._Result_Method1`); when it ISN'T, Excel falls back to a direct
single-cell cross-sheet reference (`='4.5.1.3-4 Direct N2O'!E297`) instead.
That fallback IS the "missing name" signal - no scanning for repeated
function calls, no guessing from nearby "(Method 1)" text labels (both were
seriously considered and abandoned - see history below). The equation name to
publish under is the calc sheet's own "__1"-suffixed equation (e.g. sheet
'4.5.1.3-4 Direct N2O' -> VEERG number prefix '4_5_1_3' -> search the
workbook's defined/AFE names for `VEERG_4_5_1_3__1_<Name>` - the "__1" is
always the section's top-level/final equation, "__2"/"__3" etc. are always
intermediate helpers).

SCOPE, deliberately narrow (explicit user instruction): only the "GHG
emissions method 1/2" table is touched. A Results sheet often has OTHER
breakdown tables further down (e.g. a "Manure nitrogen use" allocation table
with columns "Amount of N in manure"/"N unit") - those have different
headers and are never matched, so a SUM-of-breakdown or multi-component
bottom line is correctly left alone; naming that is a separate, harder
problem this script does not attempt.

## History that led here (earlier theories tried and abandoned)
1. First idea: scan every sheet for a VEERG function called exactly once
   (vs. many times in a per-row table) and name that cell. Investigated two
   real Result cells to validate - BOTH turned out to be `SUM(...)` over a
   per-row/monthly breakdown, not a bare single call, so "called once" would
   have missed the common case rather than caught it. Abandoned before
   writing any code.
2. Second idea: detect Method 1 vs Method 2 by searching for a nearby
   "(Method 1)"/"(Method 2)" text label. Found real examples (Dairy: label
   one column left of the value; PastureBeef: a section header ~27 rows
   above the actual total, not adjacent at all) - a workable but fragile
   heuristic with no fixed offset. Abandoned once the user pointed at the
   Results sheet instead, which needs no heuristic at all: the table's own
   column headers ARE "method 1"/"method 2", unambiguously, every time.

## BUG shipped in the first commit, then fixed (2026-08)
The equation-name regex captured ONLY the part after `VEERG_<prefix>__1_`
(the capture group), and the constructed name used just that captured
suffix - e.g. `AnnualAtmosphericDepositionEmissionsFromMMS_Result_Method1`
instead of the correct `VEERG_4_5_1_5__1_AnnualAtmosphericDepositionEmissio
nsFromMMS_Result_Method1`. Silent, no error: the name was created and the
Results-sheet formula was correctly repointed to it, so `[name]`/`[ref]`
lines in the dry-run output all looked right. Only surfaced when
`list-enterprise-results.ps1` (which matches names via `^VEERG_`) kept
showing just 1 of the expected 4 sheets after a real commit - the missing-
prefix names are invisible to anything that filters on that convention, not
just this one script. FIX: reconstruct the full `'VEERG_' + prefix + '__1_'
+ capturedSuffix` string as the base name, not just the captured suffix.
13 already-committed names across 4 workbooks (3_1_Enteric_Feedlot,
3_6_Enteric_OtherLivestock, 4_5_ManureManagement_Swine,
4_7_ManureManagement_OtherLivestock) had to be repaired after the fact via a
COM rename pass (`Name.Name = newName` - Excel auto-updates every formula
already referencing the old name, including the Results-sheet cell, so no
separate formula-rewrite was needed for the repair). LESSON: dry-run output
that "looks right" (name created, reference rewritten, no errors) is not the
same as verifying against the actual downstream consumer's matching rule -
should have re-run list-enterprise-results.ps1 (or grepped for `^VEERG_` on
every new name) as part of validating the dry-run BEFORE recommending a
commit, not after the user caught it in normal use.

## COM gotcha hit twice in one session (already documented generally in
   excel-com-errors.md, re-confirmed here)
A shared Excel COM Application instance across a batch of sequential
workbook opens degrades and throws "You cannot call a method on a null-
valued expression" on an unpredictable subset of files (3 of 23 here) - the
SAME files ran clean when processed standalone. Fix: a FRESH `New-Object
-ComObject Excel.Application` per workbook, quit+release immediately after,
not one instance shared for the whole batch. Cost is slower (Excel startup
overhead per file) but eliminates the flakiness entirely - confirmed 0
FATAL errors across all 23 workbooks after switching. Hit this a SECOND time
independently in the one-off repair script for the bug above (first 2 of 4
files fine sequentially, 3rd crashed) before applying the same fix there
too - this is not a one-off, treat "one shared Excel instance across many
workbook opens in a loop" as unsafe by default in this codebase, not just a
one-time gotcha to remember for name-sourcedata-constants.ps1 specifically.
