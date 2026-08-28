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

[shift] NEW CATEGORY - numeric-operand type mismatch (2026-08): hunts the
"formula was copied/pasted and a relative reference landed on the wrong cell"
bug. `Get-NumericOperandIssues` flags a cell used as a direct operand of an
arithmetic operator (+ - * / ^) whose contents is TEXT or BLANK instead of a
number/formula. Design decisions (confirmed with user via AskUserQuestion):
- Scope = direct arithmetic-operator operands ONLY, not function-call
  arguments (SUM(A1), ROUND(A1,2) etc. are out of scope - too many functions
  are mixed-type to do reliably, e.g. IF/XLOOKUP/CONCATENATE).
- Blank IS flagged (not just text) - user's explicit choice despite blank
  being common/intentional (blank = 0 in Excel arithmetic), accepting the
  noise tradeoff.
- Only PLAIN A1-style refs are examined - never named ranges (X_Cell_*,
  Result_*, VEERG_*, Table_*). This isn't a shortcut, it's the actual point:
  a named reference is immune to copy-paste reference drift since it doesn't
  move when a formula is pasted elsewhere, so it's assumed correct by design,
  not merely unchecked.
- A referenced cell that itself has a `<f>` (any formula, regardless of
  current cached type/value/error) is always trusted, never flagged - "not a
  number or another formula" per the original ask. A shared-formula follower
  with an empty `<f t="shared" si="N"/>` still counts as "has a formula" here.

FALSE POSITIVE caught + fixed before shipping: `Table_GWPData[CO2-e factor]`
- the regex matched "CO2" inside the structured-table column name as column
"CO" row 2, and the following literal "-" (from "CO2-e factor") looked like
an adjacent minus operator. FIX: strip ALL `[...]` bracketed spans (not just
`"..."` string literals) before scanning, repeated to peel nested spans
(`Table[[#Headers],[Col]]`) from the inside out - `$script:BracketSpanRegex`.
Column/argument names inside brackets are free text and can contain almost
anything, so nothing bracketed should ever be scanned for cell refs.

Verified full-suite run: 32 hits across 3 of 26 workbooks (13_Scope3_WIP_v14,
Enterprise_CroppingGrains, Enterprise_PastureBeef), two flavours -
(a) genuine bug: CroppingGrains `'5.1.1.1-2 Inorganic fert N2O'!Q12
=O12*P12` where P12 holds text "Grains" (a category label multiplied
directly - the kind of thing this check exists for); (b) lower-signal noise:
`'15 Scope 3'!K507..K514 =...(E507,IF(J507="kg / amount unit",
I507*10^-3, I507))` where I507:I514 are blank optional-table rows in a
fixed-length "OtherPurchasedGoodsAndServices" input table (same template
formula copied down further than actual data rows) - recurs identically in
the raw Scope3 module and both enterprises that import it. Blank-flagging
was a deliberate user choice knowing this tradeoff; if this specific
recurring pattern turns out to dominate the noise floor in practice, revisit
rather than adding a per-cell marker (user has separately said they dislike
that workaround for [sum] and would rather fix detection).

[shift] BLANK-CHECKING DROPPED (2026-08): user reported "a lt of false
positives where I have user input values that the user has not entered a
value" - formulas referencing a not-yet-filled-in input cell by plain address
tripped the same flag as a genuine drifted reference. Considered scoping the
suppression to "Input -" sheets only (never flag blank there) but the user
chose the simpler option instead: drop blank-checking from [shift] entirely,
text-only from now on. `Get-NumericOperandIssues`/`Build-CellClassificationMap`
simplified accordingly (no more `Kind = 'Blank'` case - an operand absent from
the sheet's classification map is now just skipped, `continue`). [shift] only
ever flags TEXT held by a plain-ref operand now.

[series] NEW CATEGORY - range-follow consistency check (2026-08): catches the
OTHER half of the "copied formula" bug family [shift] doesn't reach - a
function repeated down/across a run of adjacent cells where a parallel-range
operand should shift in lockstep with the run but either (a) DRIFTS (step
size inconsistent partway through - the "pasted one too many/few times"
mistake) or (b) is FROZEN at one cell across the whole run when its neighbors
in the same slot do shift. User's own framing: "a function...takes values
from other ranges of the same size. It might also take a single common value
from another cell...sometimes these cell references drift, or sometimes only
one cell is being referenced across all equations when it should be
following the pattern."

Design decisions (confirmed with user via AskUserQuestion before building):
- Series = adjacent cells (3+ in a row/column, MinSeriesRunLength=4 in code)
  whose formula is structurally IDENTICAL once cell refs are stripped to a
  placeholder (`Get-FormulaSkeleton`) - the same shape Excel's own
  "Inconsistent Formula" warning uses. NOT limited to XLSX shared-formula
  groups (`t="shared"`) - those can't drift by construction (Excel computes
  every follower cell with a guaranteed-consistent relative offset), so this
  check only looks at cells with their OWN independently-stored `<f>` text
  (`Get-SheetFormulaCells` skips empty shared-formula followers) - that's
  where a human copy-paste mistake can actually happen.
- Frozen operands are flagged regardless of $ anchoring (user's explicit
  choice: "flag any operand frozen across the whole series, absolute or
  not") - both a deliberate common value (GWP lookup, usually $-anchored)
  and an accidentally-left-behind reference (usually not) surface, for a
  human to tell apart via the Detail text rather than the checker guessing.
- Only plain A1-style refs are examined, same regex/reasoning as [shift] -
  named ranges are immune to this bug and are simply not part of any run's
  ref-slot comparison (a formula with zero plain refs returns a null
  skeleton from `Get-FormulaSkeleton` and can't join a run at all).

AXIS-AWARE ANCHOR BUG caught + fixed before shipping: the first cut judged
"deliberately anchored" by requiring BOTH ColAbs AND RowAbs (full $A$1-style
absolute). A MIXED reference like `E$60` referenced from every cell of a
COLUMN-run (varies by row) is a completely normal, correct idiom - only the
ROW needs `$` to explain why it stays fixed while the formula's own row
changes; the column's `$` status is irrelevant since nothing would move it
in a column-run regardless. Requiring both flagged this legitimate pattern
as "relative, no $ - confirm..." noise on every single occurrence. FIX:
`Test-SeriesRun` now takes an `-AxisIsColumn` flag (same one `Find-SeriesRuns`
already uses to build the run) and checks only the axis-relevant flag
(`ColAbs` for a row-run, `RowAbs` for a column-run) to decide the note text.
Verified on Enterprise_PastureBeef: before the fix, `Input - Pasture Beef`
lookup-table rows using `E$60`-style row-anchored refs in a column-run all
said "confirm..."; after, they correctly say "$-anchored on the axis that
varies - likely a deliberate common value". Real-world volume note: 496 [series] hits on PastureBeef alone on first
pass (mostly Frozen lookups in `Input - Pasture Beef` plus a batch of Drift
hits in Fuel's stationary-fuel constants lookup that may be a legitimate
per-row table lookup rather than a bug) - advisory-only by design, but
genuinely noisy.

FULLY-$-ANCHORED FROZEN SUPPRESSED (2026-08, follow-up): user's own read of
the volume - "hide the fully frozen (eg $E$6) and keep partially frozen (eg:
E$6 or $E6)". A run's dominant PastureBeef noise source turned out to be
`InputFunctions_PastureBeef.LookupLiveweightFromSeason($E$7, ...)` where
operand #1 (`$E$7`) is fully absolute in EVERY call - unambiguous, nothing to
verify. `Test-SeriesRun`'s Frozen branch now skips entirely when
`ColAbs -and RowAbs` are BOTH true; a partial anchor (only one axis, e.g. the
same function's operand #3 `E$60`/`F$66`/`G$82`... each row-run-varying
column but row-anchored) or no anchor at all still gets reported, since
either is still ambiguous enough to warrant a look. Verified: PastureBeef
496 -> 219 hits, with the (correctly-still-shown) `E$60`-style partial
anchors reworded to "$-anchored on the axis that varies, but not the other -
confirm the un-anchored axis was meant to stay fixed too" instead of the
prior wording that treated any anchoring as sufficient. STILL TOO NOISY, CORRECTED FURTHER (2026-08, second follow-up): user found a
concrete false positive in the 219-hit version and pointed at the actual
mechanism - `Input - Pasture Beef` row 61's
`InputFunctions_PastureBeef.LookupLiveweightFromSeason($E$7, ...,$D61,E$60,
...)` repeats down E61:E64/F61:F64/... (a column-run, varies by row).
Operand #4 (`E$60`) is row-anchored, column not - and that's ALREADY
sufficient: a relative fill only auto-shifts a reference along the run's OWN
axis (row, for a column-run), so anchoring just that one axis is the
complete, correct idiom for a fixed lookup-header reference; the other
axis's $ status was never meaningful and requiring BOTH (the "hide only
fully $E$6" rule from the first follow-up) was still too strict - "This will
happen a lot across all worksheets" per the user, and it did. FIX: dropped
the "both axes" requirement entirely - Frozen is now suppressed whenever the
AXIS-RELEVANT flag (ColAbs for a row-run, RowAbs for a column-run) is
anchored, full stop, regardless of the other axis. Only reported when the
axis that varies has NO $ at all. Verified: PastureBeef 219 -> 144 hits, and
a spot-check of the survivors looks like real signal (e.g. `3.2.1.1-2
Enteric Methane` E26:E29 repeatedly comparing against `D21` with no $ inside
an IFS season-rotation formula, and several `4.2.1.1-2 Methane` Drift hits
referencing `Input - Pasture Beef` cells with inconsistent steps). Lesson:
"partially anchored" was never really the right frame - what matters is
ONLY whether the axis that actually moves during a fill is anchored; the
orthogonal axis's $ status carries no information either way and should
never have been part of the suppression condition.

THIRD FALSE-POSITIVE PATTERN - "*_Arguments()" metadata spill (2026-08):
user pointed at a DIFFERENT un-anchored Frozen shape: a data range (e.g.
Dairy `3.3.1.1-2 Enteric Methane` rows) has a UNIT cell just outside it to
the right showing "kg/head"/"tonnes"/etc, produced by
`=CHOOSECOLS(EntericMethane_Dairy_Equations.VEERG_3_3_1_1__1_TotalAnnual
MethaneFromEntericFermentation_Arguments(),3,4)` - a VEERG `*_Arguments()`
equation returns a metadata array (name/unit/description per argument);
CHOOSECOLS picks specific columns and the result SPILLS as a dynamic array.
"Because that arguments call returns an array, the cell being referenced
might be one of the array values, and the _Arguments function is called
above it" - i.e. the formula lives on ONE anchor cell but the actual cell a
data-run row references (the unit text) is a SPILL MEMBER with no `<f>` of
its own (OOXML dynamic-array spill cells store only a cached `<v>`, same
"no formula to trust" shape as a shared-formula follower) - and every row of
the data block legitimately references that SAME spilled unit cell without
any `$`, since it's a single value for the whole block. User: "This will
happen a lot across all worksheets."
FIX: new `Get-ArgumentsSpillCells` (per sheet, cached lazily in
`Get-SeriesConsistencyIssues` via `$argsSpillCache`, same lazy-per-sheet-name
pattern as [shift]'s classCache) scans every formula cell whose text
contains `_Arguments(`, and unions in (a) that cell's own address and (b)
every address in its `<f ref="...">` dynamic-array spill footprint (walked
with a new tiny `Expand-CellRange` helper - safe to fully enumerate since a
spill footprint is a handful of cells, not a data range). `Test-SeriesRun`'s
Frozen branch now ALSO skips reporting when the un-anchored target address
falls in this set, alongside the existing axis-anchored check. Verified: no
regression on PastureBeef (144, unchanged - the pattern doesn't occur
there); ran clean (no crash, 145 hits, 0 residual "Arguments" mentions) on
Enterprise_Dairy_WIP_v01, where the pattern DOES exist, though the workbook
containing the user's literal example (`3_3_Enteric_Dairy_WIP_v04.xlsx`) was
open in their own Excel session at the time and couldn't be scanned directly
(the checker's ZipFile.Open needs the file free) - worth a follow-up direct
check once it's closed, but the fix logic mirrors the already-proven [shift]
classCache pattern closely enough to trust without it.
Three Frozen-suppression rules now, in order: (1) axis-relevant $-anchor
present, (2) target is an `_Arguments()` spill cell, (3) neither -> report.
If a FOURTH recurring false-positive shape turns up, add another rule to
this list rather than loosening the axis or $ logic further - the pattern
so far is "each shape has a specific, checkable reason it's legitimately
frozen," not "frozen is just inherently too noisy to check."

FOURTH ROUND - STYLE-BASED EXCLUSION SUPERSEDES per-pattern chasing
(2026-08): user reported the Arguments false positives were "still being
found" (the formula-text-substring fix from the prior round wasn't fully
covering it), AND gave a NEW distinct false positive:
`=HYPERLINK(Common_InputFunctions.Utility_SourceHyperlink($D$77),$D$77)`
repeated down a column of source-table links - "not part of a range series"
per the user. Crucially, the user described BOTH the unit cells and the
hyperlink cells by their CELL FORMATTING/STYLE ("Unit no indent",
"Arguments hyperlink") rather than by formula shape - confirmed these are
real, workbook-wide named cell styles (`xl/styles.xml`
`<cellStyle name="Arguments hyperlink" xfId="64">` etc, present consistently
across nearly every module workbook checked). Rather than keep chasing
individual formula-text patterns as new shapes surface, switched to a
GENERAL style-based exclusion that is authoritative (the author's own
explicit intent-marking) instead of inferred: `Get-CellStyleNameMap` (once
per workbook) resolves `xl/styles.xml`'s standard `cellStyle name/xfId` ->
`cellStyleXfs` position <- `cellXfs` `xf/@xfId` linkage into a `cellXfs
index -> style name` map; `Get-ExcludedStyleCells` (per sheet, lazily
cached) uses it to find every cell address styled with a name in
`$script:SeriesExcludedStyleNames = @('Unit no indent', 'Arguments
hyperlink')`. Applied at BOTH ends: `Get-SeriesConsistencyIssues` excludes
styled cells from `$parsed` entirely (never a run SOURCE member - fixes
HYPERLINK rows being treated as a series of their own), and `Test-SeriesRun`
checks the Frozen operand's TARGET address against the same per-sheet cache
(fixes a styled-but-not-`_Arguments()`-shaped unit cell being missed).
KEPT the existing `_Arguments(` text-based `Get-ArgumentsSpillCells` check
and added a bare `HYPERLINK(` text-based exclusion in `Get-FormulaSkeleton`
as belt-and-suspenders alongside the style check, not instead of it - in
case a cell's style is ever missing/wrong but the formula shape is still
identifiable. Extend `$script:SeriesExcludedStyleNames` first if a new
documentation-style false positive turns up; only fall back to formula-text
matching for shapes that AREN'T reliably styled.

BUG SHIPPED THEN CAUGHT BEFORE the fix was usable: `Get-ArgumentsSpillCells`
and `Get-ExcludedStyleCells` both `return`ed their HashSet bare (`return
$set`), and BOTH are captured via SIMPLE ASSIGNMENT at their call sites
(`$cache[$sheetName] = Get-Whatever ...`) - this is the exact
"empty/enumerable collection auto-unrolls on return" gotcha already
documented at the bottom of this file for `$findMissing`'s List<string>
("PowerShell UNROLLS it... wrap call sites in @() or .Count throws"), except
here the fix can't be "wrap the call site in @()" (that would flatten the
HashSet into a plain array, losing `.Contains()`) - the correct fix for a
collection you need to KEEP AS A COLLECTION on return is the COMMA-OPERATOR
PREFIX: `return ,$set`, forcing PowerShell to treat the whole HashSet as ONE
output object instead of enumerating it. Without this, an EMPTY HashSet
(the common case - most sheets have zero styled/Arguments cells) collapses
to `$null` on assignment, and the next `.Contains()` call throws "You cannot
call a method on a null-valued expression" - caught immediately on the very
first all-empty sheet once `Get-ExcludedStyleCells` started running
unconditionally for every sheet (unlike the narrower `_Arguments(`-gated
`Get-ArgumentsSpillCells`, which had been getting lucky - most sheets
happen to contain `_Arguments(` SOMEWHERE, so its result was rarely
actually empty in practice, masking the same latent bug in the prior
round). LESSON: any new HashSet/List-returning helper in this script needs
the `,$x` treatment by default whenever its call site is a simple
assignment rather than `@()`-wrapped or `foreach`-consumed - don't rely on
"it happened not to crash in testing" as proof a collection-returning
function is safe.

FIFTH ROUND - Drift false positives from CONSOLIDATED multi-range targets
(2026-08): user's exact example - `4_1_ManureManagement_Feedlot_WIP_v08.xlsx`,
where named range `M1_Table_M_j_m_T` is built by concatenating three
separate source tables (`M1_Table_M_j_m5_T1`, `M1_Table_M_j_m_T2`,
`M1_Table_M_j_m1_T3`) into one target list. The OLD Drift algorithm compared
every step in a run to ONE global "majority" delta and flagged every
deviation - so a legitimate 3-way consolidation (each source segment
internally consistent, but discontinuous AT the boundary between segments,
and possibly with a genuinely different internal step per segment) got
flagged extensively, not just at the boundaries. User's own suggested fix
(check whether the leftmost column TEXT/label matches between source and
target rows) was NOT implemented - reading and semantically comparing cell
DISPLAY VALUES across sheets is heavier and more fragile than a structural
fix. Instead, two complementary structural fixes:
1. `Test-SeriesRun`'s Drift branch, new constant `$script:MinDriftSegmentLength = 2`:
step-comparison now RUN-LENGTH-ENCODES the step sequence into
segments of consecutive equal deltas. A segment repeated 2+ times in a row
is trusted as its OWN legitimate sub-pattern (not compared against the
overall majority at all) - a real one-cell mistake essentially never
reproduces the exact same wrong offset twice in a row by chance, so a
sustained alternate step is much more likely to be a second/third deliberate
source segment. This alone fixed the BULK of the false positives (every cell
WITHIN each of the 3 consolidated segments), leaving only the 1-cell
transitions AT each boundary still ambiguous.
2. `Get-NamedRangeAreas`/`Find-EnclosingRangeName`: for the remaining
boundary-transition ambiguity (a lone anomalous step sandwiched between two
trusted segments is STRUCTURALLY IDENTICAL whether it's a genuine
single-cell bug or a deliberate cross-table boundary - address deltas alone
truly cannot disambiguate this, confirmed by direct reasoning/synthetic
test), added a STRUCTURAL tie-breaker: parse every workbook/sheet-scoped
defined name whose RefersTo is a simple single-rectangle A1 reference (skips
multi-area unions, structured tables, error tokens) into a list of
{Name, Sheet, C1,C2,R1,R2} areas. When flagging a short/isolated Drift
segment, check whether the FROM and TO cells at that exact step fall inside
two DIFFERENT named-range areas (e.g. last row of `M1_Table_M_j_m5_T1`,
first row of `M1_Table_M_j_m_T2`) - if so, it's a real cross-table boundary,
suppressed; if both are in the SAME named range (or neither is in any named
range), it's still flagged, preserving the check's ability to catch a
genuine single-cell mistake INSIDE one source table. This is a STRUCTURAL
signal (crossing a real, named boundary) rather than inferring intent from
cell content - directly usable here because the user's own example named
three SPECIFIC named ranges, confirming named ranges are how this codebase
already demarcates "this is table N of a consolidation."
VERIFIED against the user's exact file (closed by them mid-session so it
could be scanned): 0 "M1_Table" mentions survive in the output (down from
an unmeasured but clearly large false-positive count per the user's report);
only 2 unrelated Drift hits remain elsewhere in the workbook, both plausible
genuine outliers (`$G$139` referenced at two isolated spots in an otherwise
column-stepping row). No regression on PastureBeef (140, unchanged - it
doesn't have this consolidation pattern).
KNOWN REMAINING GAP: a consolidation between UNNAMED plain cell blocks (no
defined name backing either side of the boundary) still isn't
disambiguated by the named-range check, and could still be flagged if the
unnamed sub-range is only 1-2 rows long (too short to earn
MinDriftSegmentLength's trust on its own). Revisit with the user's original
content-matching idea only if this specific unnamed-block case turns out to
matter in practice - the named-range fix already covers the concrete
example given and is a much smaller, safer piece of machinery than full
cross-sheet value comparison.

SIXTH ROUND - the "known remaining gap" above hit almost immediately
(2026-08): user's next example, `4_2_ManureManagement_BeefPasture_WIP_v10.xlsx`,
named range `M2_Table_E_CH4` - TWO MMS-type row groups ("Pasture range and
paddock" then "Anaerobic lagoon") consolidated into ONE table, one name, no
second named range to cross-check against, so the fifth-round fix couldn't
help: `'4.2.1.1-2 Methane'!AC263  operand #3 ... expected step (0,1) from
AC262, found step (0,-11)`. Asked the user directly (AskUserQuestion) how to
treat an isolated single-step transition sandwiched between two trusted
segments when there's no second name available - three options: suppress
unconditionally, suppress only when both flanking segments are long enough,
or leave it flagged. User picked the MIDDLE option - flagging still nothing
by default is too aggressive (risks masking a genuine one-cell mistake that
happens to sit between two long uniform runs), leaving it flagged is what
prompted this whole line of fixes in the first place, so a LENGTH-GATED
trust threshold is the compromise.
FIX: new constant `$script:MinBoundaryFlankLength = 4` (deliberately higher
than `MinDriftSegmentLength`'s bar of 2 - a segment merely reaching 2 is
enough to be "not a mistake ITSELF", but a much longer, well-established
run on BOTH sides is a materially stronger signal before trusting an
adjacent isolated transition as a sub-group boundary rather than a stray
error). `Test-SeriesRun`'s Drift loop switched from `foreach ($seg in
$segments)` to an indexed `for` so each short segment can look at its
immediate neighbours (`$segments[$si-1]`/`$segments[$si+1]`); an isolated
segment with BOTH neighbours >= `MinBoundaryFlankLength` is skipped
entirely (not even per-cell-checked), same as a trusted segment. The
existing named-range check still runs afterward for segments that don't
clear this bar, so a short-flanked isolated step can still be rescued if it
DOES happen to cross two different named ranges.
VERIFIED against the user's exact file: 0 "AC263" mentions survive; total
[series] hits dropped 80 -> 72 with no crash. Also re-ran
Enterprise_PastureBeef_WIP_v01 (which embeds this same ManureManagement_
BeefPasture module) and saw a further drop, 140 -> 88, confirming the fix
generalizes beyond the raw module file into the built enterprise output
too. Spot-checked a sample of the remaining Drift hits (several repeating
`step (0,-3)` patterns at regular column intervals, e.g. `4.2.1.1-2 Methane`
H181/J181/... columns) - a DIFFERENT, not-yet-reported shape, left alone
rather than guessed at pre-emptively; flag if the user reports it as noise
too.

SEVENTH ROUND - the "different shape" above WAS the same false positive, an
OFF-BY-ONE in the sixth round's threshold (2026-08): user confirmed with
the exact `H181` example flagged above and noted "each section is made up
of four rows." Root cause: `$seg.Length` (from the run-length encoding) is
a DELTA count, i.e. one LESS than the segment's cell/row count (N cells
produce N-1 pairwise deltas) - a genuine 4-row MMS-type block therefore has
`Length == 3`, which never cleared the sixth round's `-ge 4` check against
`MinBoundaryFlankLength` (itself named/reasoned about in row terms despite
comparing a delta count). FIX: renamed the constant to
`$script:MinBoundaryFlankRows` (same value, 4) and compare `(seg.Length + 1)`
- the actual cell count - against it, everywhere it's used. LESSON: when a
threshold is meant to represent "N rows/cells," compare it against a
CELL/ROW count, never directly against a segment's `.Length` field (a delta
count) without the `+1` - the same units mismatch could recur anywhere else
`.Length` is compared to a "how many rows" intuition rather than "how many
steps between rows" one.
VERIFIED against the same BeefPasture file: 0 "H181" mentions survive;
total [series] hits dropped 72 -> 19 (a much bigger cut than the sixth
round's own already-large improvement, confirming most of the recurring
noise really was this exact 4-row-section shape, not a separate issue).
Remaining hits include a cluster of `step (0,-3)` drifts in
`EntericMethane_PastureBeef_Equations`/`IFS(...Calving season...)`-style
rotating-lookup formulas - structurally similar to the very first Frozen
false positive found (a small fixed reference block referenced in rotating
order) but manifesting as Drift instead; NOT yet reported by the user as a
problem, left alone rather than guessed at pre-emptively per the pattern
established throughout this whole investigation - only fix what's actually
been flagged as noise.

EIGHTH ROUND - a group can be as short as ONE row (2026-08): user's next
example, `4_5_ManureManagement_Swine_WIP_v06.xlsx` `'4.5.1.1-2 Methane'!
E108`: a "Solid storage" MMS-type group only 1 row long, sandwiched between
a long established segment and (presumably) another. A 1-row group has NO
internal repetition of its own - `MinDriftSegmentLength`/
`MinBoundaryFlankRows` structurally cannot validate it, since there's
nothing to measure; this is a genuine dead end for address-only heuristics,
not a tunable threshold problem like the sixth/seventh rounds were. Asked
the user directly (AskUserQuestion) among three paths: implement the
original label-matching idea (real, substantial new work - infer which
column holds a row label per sheet, compare text robustly), relax the
flank rule to require only ONE well-established neighbour instead of both,
or accept this as residual noise. User picked the middle option, with an
explicit qualifier: "apply only to rows."
FIX: `Test-SeriesRun`'s `$isTrustedBoundary` now branches on `$AxisIsColumn`
- for a COLUMN-run (varies by ROW - every example in this whole thread has
been this shape), only ONE of `$prevSeg`/`$nextSeg` needs to clear
`MinBoundaryFlankRows`, not both; a ROW-run (varies by COLUMN) still
requires BOTH, since there's no evidence yet that consolidation happens
along that axis and the user explicitly scoped the relaxation to rows.
Accepted tradeoff (stated in-code and here): a genuine one-cell mistake
sitting at the very EDGE of a long uniform column-run (only one long flank,
nothing established on the other side) can now also go unflagged - judged
less likely in practice than a real short/1-row group, per the user's
choice.
VERIFIED against the user's exact file: 0 "E108" mentions survive; total
[series] hits on that workbook dropped further (26 remaining, all a
different not-yet-reported shape - VLOOKUP-based Stage2 lookups and a
cross-sheet `$G$13` reference, left alone). Re-checked BeefPasture (still
19, unchanged - the one-sided relaxation only ADDS suppression, never
removes any) and Enterprise_PastureBeef (88 -> 36) to confirm no
regressions from loosening the rule.
CUMULATIVE RESULT across all eight rounds of this investigation: PastureBeef
enterprise [series] hits went from an initial peak of 496 down to 36 (~93%
reduction), through a sequence of increasingly precise structural
suppression rules (axis-aware $-anchoring, `_Arguments()` spill detection,
named-style exclusion, HYPERLINK exclusion, Drift segmentation, named-range
boundary crossing, flank-length gating, and finally axis-specific one-sided
flank trust) - NONE of which required reading or comparing actual cell
VALUES/labels, only formula structure, cell styles, and workbook defined
names. The original user-proposed label-matching approach was never needed
to reach this point; keep it in reserve only if a future false-positive
shape turns out to be unreachable by structural means alone.

NINTH ROUND - two more named styles added directly (2026-08): user asked
to also ignore cells styled "Unit" and "Arguments" (the base styles, not
just their "no indent"/"hyperlink" variants already handled). Confirmed
both exist verbatim in `xl/styles.xml`'s `<cellStyle name="...">` list
(alongside a separate "Unit total" style NOT requested, left out).
`$script:SeriesExcludedStyleNames` extended from `@('Unit no indent',
'Arguments hyperlink')` to `@('Unit', 'Unit no indent', 'Arguments',
'Arguments hyperlink')` - no other code changes needed, since the whole
style-exclusion mechanism (`Get-CellStyleNameMap`/`Get-ExcludedStyleCells`)
already applies every name in this list identically. Verified on
Enterprise_PastureBeef_WIP_v01: 36 -> 35 (a small further drop - most
"Unit"/"Arguments"-styled cells were apparently already being caught by
earlier structural checks, e.g. axis-anchoring or the `_Arguments(`
formula-text heuristic, so this mainly closes remaining gaps rather than
being a big new source of noise on its own).

TENTH ROUND - a REPEATING group pattern, not a single boundary (2026-08):
user's next example, `4_5_ManureManagement_Swine_WIP_v06.xlsx`
`'4.5.1.1-2 Methane'!G175`/`G178` etc: a VLOOKUP-based Stage2Methane
formula where operand #1 stays frozen for ~3 rows (one swine-class row
group) then steps forward once, repeating throughout the WHOLE table -
structurally different from every prior round: not one boundary between
two/three named consolidated ranges, but MANY repeating transitions of the
identical shape. Every prior fix failed on this because the "step" delta
NEVER repeats consecutively (each group boundary is a lone isolated step,
forever) - `MinDriftSegmentLength`'s consecutive-repeat rule can't see it.
FIRST FIX (periodicity): new `Test-PeriodicDeltaPattern` - for candidate
period P=2..floor(n/2), check whether the WHOLE delta sequence is fully
explained by `keys[i]` depending only on `(i mod P)` (every phase position
identical across all complete cycles). If found, skip Drift entirely for
that slot (checked BEFORE the segment-based logic, since it's a strictly
stronger/cleaner explanation when it applies). Confirmed with the user
up-front via AskUserQuestion which of three approaches to take (trust any
3+ total repeat regardless of spacing / require true even-spacing (period)
/ leave flagged) - user picked the STRICTER periodic option specifically to
avoid two coincidentally-matching independent bugs getting masked by a
looser "3+ times anywhere" rule.
DIDN'T FULLY WORK - traced the REAL data directly (`4_5_...v06.xlsx`,
unlocked at the time) and found the actual row-group sizes are 3, 4, 3, 3,
3, 4 (column G, rows 165-184, source rows F103-F108) - NOT a fixed period.
The strict periodicity check correctly declined to fire (as designed - it
requires a PERFECT fit), so `G175`/`G178` stayed flagged. This is a
concrete, real instance of the exact "irregular groupings not caught"
caveat the user had explicitly accepted - not a bug, but the chosen
tradeoff visibly costing something in practice. Reported this diagnosis
honestly and asked a SECOND AskUserQuestion: extend to the previously-
declined "N+ total occurrences, no fixed spacing required" rule after all
(now that a real irregular case is in hand), or accept the gap. User this
time chose to extend it, WITH A STRICTER BAR than originally offered (3
occurrences, not 2 - the same reasoning as before, minimizing risk of
masking a small delta like (0,1) that happens to be both a common
deliberate step AND a common accidental mistake).
SECOND FIX (total-occurrence fallback): new constant
`$script:MinDriftRepeatCount = 3`. Computed once per slot,
`$totalCounts[key]` = raw frequency of each delta value across the WHOLE
`$keys` array regardless of position/adjacency. A segment is now trusted
(excluded from `$trustedCoverage`'s majority-selection AND from being
flagged) if EITHER its own consecutive length >= `MinDriftSegmentLength`
(existing rule) OR `$totalCounts[$seg.Key] -ge $script:MinDriftRepeatCount`
(new rule) - this runs AFTER the periodicity check fails, as a fallback
specifically for irregular-but-still-repeating structures.
VERIFIED against the user's exact file: 0 mentions of any of the 12 flagged
cells (G175/G178/F175/F178/H175/H178/P175/P178/Q175/Q178/R175/R178)
survive; total [series] hits on that workbook dropped 18 -> 6, with the
remaining 6 being a DIFFERENT, plausible-looking pattern (`Constants -
Swine` rows referencing arbitrary non-sequential `Constants - Common
livestock` rows, e.g. jumps of 5/2/6 - similar in shape to the very first
Fuel example seen early in this whole investigation) not yet reported by
the user. No regression on Enterprise_PastureBeef (still 35).
Drift's trust logic is now FOUR layers deep: (1) periodicity (strict, no
exceptions), (2) consecutive-segment length, (3) total-occurrence count
(non-consecutive), (4) named-range-crossing / flank-length for whatever's
still isolated after 1-3. Each layer was added only after a CONCRETE
real-file counter-example demonstrated the previous layers' combined
insufficiency - resist the urge to pre-emptively add a fifth layer without
one.

ELEVENTH ROUND - root-caused instead of adding a fifth trust layer
(2026-08): user's next two examples - `8_Fuel_WIP_v05.xlsx` `'8.2.1.1-2
Stationary fuel'!E26..E39` and `4_5_ManureManagement_Swine_WIP_v06.xlsx`
`'Constants - Swine'!D42..D60` - both `=SomeSheet!$D$nn`-style operands
jumping by essentially ARBITRARY, non-repeating amounts (5, 4, 6, 2, -19,
26 in the Fuel case). User's own explanation: "there are more items in the
source data than needed to be displayed in the target cells" - i.e. the
target is a CURATED, hand-picked subset of a larger source list, not a
fill-down at all. This didn't fit the tenth round's fix (occurrences here
repeat only 2x, not the 3+ `MinDriftRepeatCount` bar) and wouldn't have
been fixed by lowering that threshold either, since some jumps (Fuel's
-19 and 26) never repeat even once.
ROOT CAUSE, not another trust-layer patch: EVERY operand in both examples
is FULLY $-anchored (`$D$63`, `$D$67`, `$E$15`, etc - both column AND row
absolute). Excel NEVER shifts a fully-absolute reference when a formula is
filled or copied - so if the address differs between cells anyway, each
cell's formula was necessarily typed or edited INDIVIDUALLY, which is
categorically not the "relative reference computed wrong during a fill"
bug class this whole check exists to catch, no matter how irregular the
jumps look. This is the exact same logic the Frozen check already applies
when a fully-anchored operand stays CONSTANT (suppressed, un-conditionally,
since round one of this whole investigation) - the missing piece was
extending it to the case where a fully-anchored operand VARIES.
FIX: `Test-SeriesRun` now checks, immediately after building `$slotRefs`
for a slot (before ANY delta/step/pattern analysis, Frozen or Drift alike):
if every ref in the slot has BOTH `ColAbs` and `RowAbs` true, `continue` to
the next slot - skips the whole slot outright. Simpler and more general
than every prior layer (no threshold to tune, no periodicity search) -
because it addresses the actual mechanism (how Excel's fill/copy handles
$-anchoring) rather than pattern-matching the SYMPTOM (irregular jump
sizes) with progressively cleverer statistics. No behavior change for the
already-handled fully-anchored-and-CONSTANT case (it was already always
suppressed via the axis-anchoring check) - this only newly covers
fully-anchored-and-VARYING.
VERIFIED: `4_5_ManureManagement_Swine_WIP_v06.xlsx` went from 6 [series]
hits to 0 (completely clean) - all 6 were exactly this shape. Fuel's exact
file was open in the user's own Excel session at verification time and
couldn't be scanned directly, but it's the identical mechanism, already
proven correct on Swine; re-check when convenient. Enterprise_PastureBeef
dropped further, 35 -> 24, with no crash - confirms the rule generalizes
without needing per-file tuning.
LESSON for future false-positive rounds in this check: before adding
another trust-layer/threshold to the Drift heuristics, check whether every
operand in the run is fully $-anchored FIRST - if so, the whole
step-pattern question is moot and no amount of periodicity/repeat-count
cleverness was ever going to be the right fix.

TWELFTH ROUND - [series] flipped to OFF BY DEFAULT (2026-08): after eleven
rounds of false-positive tuning, user judged the check to have reached
diminishing returns - "the remaining range series errors are tricky cases
that are still valid" (i.e. genuinely ambiguous, needing a human judgment
call each time, not more heuristics) - and asked for it out of the default
`npm run find-errors` scan, with an explicit opt-in flag to run it when
wanted. Renamed the switch from `-SkipSeriesCheck` (opt-out, default-on) to
`-IncludeSeriesCheck` (opt-in, default-off) - inverted logic
(`if ($IncludeSeriesCheck)` instead of `if (-not $SkipSeriesCheck)`) at the
single call site in the main loop. Every other category ([cell]/[name]/
[link]/[sheet]/[sum]/[shift]) stays default-on as before - this is scoped
to [series] only, per what was asked. Added a discoverability note to the
summary output (`([series] range-follow check skipped by default; pass
-IncludeSeriesCheck to run it)`), mirroring the existing hiddenCached-note
pattern, so a default run still tells you the check exists rather than
silently omitting it. `npm run find-errors -- -IncludeSeriesCheck` works
via the existing pass-through `--` npm convention - no package.json change
needed. Verified both directions: default run on Enterprise_PastureBeef
shows CLEAN + the new note with zero [series] noise; `-IncludeSeriesCheck`
still finds all 24 hits exactly as before.

MAJOR ENGINE GOTCHA found while building [series] (2026-08, generalizes well
beyond this script): on this machine's PowerShell 5.1 build (5.1.26100.8875),
wrapping a MATERIALIZED `System.Collections.Generic.List[T]` VARIABLE
directly with the `@()` array-subexpression operator throws
`System.ArgumentException: Argument types do not match` - reproduced for
List[object]/List[string], populated OR empty, with or without
Set-StrictMode, in a completely fresh `powershell.exe` process (not session
contamination). `[List[object]]::new()` built the exact same runtime type
but did NOT trigger it when wrapped the same way - the bug is specifically
about `New-Object`-constructed generic collections hitting `@()`, not the
type itself. Cost real debugging time: the exception's reported position was
the literal `return @($issues)` line, but earlier isolated tests of that
same line's `-f`/hashtable-literal shape in isolation passed fine, so the
list-wrapping nature of the bug wasn't obvious until directly bisecting with
`Measure`-style debug prints down to that one line.
CRITICAL CONTEXT: every OTHER check in this script accumulates issues via a
plain array (`$issues = @()` + `+=`), NOT `Generic.List[T]` - [series] was
the only new code using a List for the top-level accumulator, which is why
this had never surfaced here before. SAFE PATTERNS (both confirmed working):
(1) don't use Generic.List for a value you intend to return - use `$x = @()`
+ `+=` like every other check in this file; or (2) if you do use a List
internally, `return` it BARE (`return $issues`, no `@()`) and let it unroll
into the pipeline/output stream one item at a time (exactly like a bare
`return` of any enumerable already does throughout this script) - the
CALLER'S `@(Get-Whatever ...)` wrap (capturing a function call's output
stream) is unaffected, only wrapping the List VARIABLE directly is broken.
`.ToArray()` also produces a real `object[]` that's safe to `@()`-wrap
afterward (used inside `Find-SeriesRuns` for exactly this reason). If a
future check on this codebase mysteriously throws "Argument types do not
match" pointing at an otherwise-innocuous `return` or assignment line, check
for a `@(<GenericListVariable>)` first before doubting the surrounding logic.

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
- A multi-arg `-f` call (`"..." -f $a, $b, $c`) used directly as a hashtable-
  literal VALUE parses fine on its own, but breaks with "Unexpected token
  'Formula'" / "hash literal was incomplete" errors if that `@{...}` literal
  is itself an ARGUMENT to a method call, e.g. `$list.Add([pscustomobject]@{
  X = "..." -f $a, $b })` - the parser loses track of where the hash literal
  ends once it's nested inside the method call's own parens and the comma-
  list leaks out as extra method arguments. Fix: wrap the whole `-f`
  expression in its own parens: `X = ("..." -f $a, $b)`. Same family as the
  documented "inline if(){}else{} in a hashtable value" gotcha - both are
  really "multi-token/comma-bearing expressions need explicit parens when
  used as a hashtable value nested inside another call's argument list".
- Wrapping a `New-Object System.Collections.Generic.List[T]`-constructed
  variable directly with `@()` can throw `System.ArgumentException: Argument
  types do not match` on some PS5.1 builds (see the [series] MAJOR ENGINE
  GOTCHA entry above) - avoid Generic.List for a function's return-value
  accumulator; use a plain array + `+=` like every other check in this file.
