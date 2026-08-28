# Enterprise template authoring (manual process) + list-enterprise-results.ps1

## list-enterprise-results.ps1 does NOT respect `include` subsetting (2026-08)
Found while drafting `Enterprise_Poultry.json`: its `modules[]` entry for
Fertiliser deliberately imports ONLY `Constants - Fertiliser` (a dependency
of the Scope3 "manure sent off site" calc), not any of Fertiliser's own
calculation sheets - yet `list-enterprise-results.ps1 -EnterpriseId Poultry`
still printed a full "Fertiliser and lime" section with all of Fertiliser's
Method1/Method2 results (5.1.1.1-2, 5.2.1.1-2, 5.3.1.1-2, etc). The script
has no awareness of a module selection's `include` filter at all (grepped
the script - no reference to "Include" anywhere) - it lists every result
name found in a SELECTED module's source workbook unconditionally, as if
the whole module were imported. Not a bug exactly (documented nowhere as
respecting `include`, and fixing it would need duplicating
`build-enterprise-excel.ps1`'s subsetting logic) but a real trap when
copying its output into a hand-authored Results sheet: a constants-only
module's calc-sheet results will be listed even though they will NOT exist
in the actual built enterprise workbook. Manually skip any section for a
module whose `include` doesn't cover `calculation` when using this
script's output - same caveat applies to any future enterprise with a
constants-only (or narrowly-subsetted) module dependency, not just Poultry.

## Result names aren't always Method1/Method2 (2026-08)
`list-enterprise-results.ps1` originally matched only `_Result_Method[12]`
suffixes. The Fuel module (`8_Fuel_WIP_v05.xlsx`, also embedded inside
`13_Scope3_WIP_v14.xlsx` for section 15.11) instead publishes ONE result PER
GAS from a single calculation - `_Result_CO2`/`_Result_CH4`/`_Result_N2O`/
`_Result_Scope3`, no Method1/Method2 concept at all (fuel combustion doesn't
have "two alternate methodologies" the way most VEERG equations do). The
names were always correctly published; the script's regex just couldn't see
them, so Fuel silently showed NOTHING for every enterprise that includes it
(Dairy, PastureBeef, Swine, ...) - not a Swine-specific or missing-name bug,
just never noticed until someone asked "where did Fuel go" for Swine
specifically. FIX: `$resultNameRegex` now captures the suffix generically
(`_Result_(?<suffix>[A-Za-z0-9]+)`); output renders the familiar fixed
"Method 1: .../Method 2: (none)" two-slot form when a (sheet, base) group's
suffixes are a subset of {Method1, Method2} (preserves the "(none)" signal
that's genuinely useful there), and falls back to listing whatever suffixes
ARE present with no "(none)" padding otherwise (a fixed expected set doesn't
exist for the per-gas shape). Confirmed no other module uses a non-Method
suffix (grepped every `Excel/*.xlsx` for `_Result_<suffix>` where suffix
isn't Method1/Method2 - only Fuel's two equations, 4 suffixes each, appeared).
If a THIRD naming shape ever turns up, extend the same way rather than
special-casing Fuel by name.

Making a new enterprise's `Excel/Enterprises/*_Template_*.xlsx` is the ONE stage
in the pipeline (see enterprise-build-pipeline.md Stage 1) with zero tooling -
everything from Stage 2 (`build-enterprise`) onward is scripted, the template's
CONTENT is hand-authored. Automating the full build was explicitly ruled out by
the user (2026-08): "varying positioning of cells in excel from enterprise type
to enterprise type" makes it not worth the effort. What IS worth automating is
the tedious LOOKUP work a human does while hand-authoring - see
list-enterprise-results.ps1 below.

## The process (as dictated 2026-08, in progress - more sheets to follow)

### 1. 'Input - Enterprise' sheet
- Common fields/table/results, identical every enterprise (mechanical, same six
  names + one table + four results every time):
  X_Cell_Site_FarmName, X_Cell_Site_ReportingEntityName, X_Cell_Site_StartDate,
  X_Cell_Site_EndDate, X_Cell_Enterprise_ProductionTimeframeStart/End,
  Table_Input_Enterprise_PlantingRemovals,
  Result_EmissionsForWholeProductionCycle_Gross_Method1/2,
  Result_CarbonRemovalsFromPlantings,
  Result_EmissionsForWholeProductionCycle_Net_Method1/2.
- Unique inputs (yield/income/proportion-sold fields, e.g.
  X_Cell_Enterprise_GrainYield, X_Cell_Enterprise_PercentGrainSold) CANNOT be
  derived - human judgement call per enterprise, by design.

### 2. 'Results - Activity period' sheet
- List Scope 1 emissions sources = the imported chapter modules (as-is or
  edited for readability), possibly a sub-section of a chapter (e.g.
  "5.1.1.1-2 Inorganic fert N2O"). Varies per enterprise (driven by that
  enterprise's `modules[]`).
- Reference the Method 1 and Method 2 result named cell for each, e.g.
  `=VEERG_5_1_1_1__1_InorganicFertiliserApplicationN2OEmissions_Result_Method1`.
- Display the GHG unit per source (t N2O, t CO2, t CH4, t CO2e, ...).
- Convert to CO2-e via a common function + GWP lookup, e.g.
  `Common_Equations.Common_ConvertN2OToCO2e(E9, XLOOKUP("N2O",
  Table_GWPData[Gas], Table_GWPData[CO2-e factor]))`.
- Same treatment for Scope 2 - IDENTICAL across every enterprise (always the
  Electricity module), copy-paste-able as-is, not enterprise-specific.
- Same treatment for Scope 3 - varies per enterprise like Scope 1.
- Charts for Scope 1 and Scope 3 (detail not yet worked out).
- 'Carbon removals for activity period' = `=SUM(Table_Input_Enterprise_
  PlantingRemovals[Removal allocated to this enterprise])`.
- Gross/net summary table, Method 1 and 2, formulas FIXED regardless of
  enterprise (only the Scope 1/3 rows feeding them vary):
  `Result_Enterprise_Scope1_Method1+Result_Enterprise_Scope2_Method1+
  Result_Enterprise_Scope3_Method1` (gross), `Result_Enterprise_Gross_Method2-
  Result_Enterprise_Removals` (net).

### 3. 'Product emissions intensity' sheet
Not yet described by the user ("There's a lot going on") - come back to this
when they walk through it.

## scripts/list-enterprise-results.ps1 (npm `list-enterprise-results -- <Id>`)

Built to eliminate the tedious part of step 2 above: hunting down, per chapter
module, the exact `VEERG_..._Result_Method1/2` name + cell + display unit to
paste into 'Results - Activity period'. Read-only, XML-only (no COM/Excel -
safe to run while other workbooks are open, confirmed working while
Enterprise_Dairy_Template was open+locked in Excel).

- For each module in `Enterprises/Enterprise_<Id>.json`'s `modules[]` (bare
  string = whole module, object with `include.calculation` = subset -
  resolved the same way build-enterprise-excel.ps1 does), resolves the source
  workbook via `_ModuleRegistry.json` (same version-agnostic prefix match as
  `Resolve-SourceWorkbook`), reads its workbook-scoped `<definedName>`s.
- Finds every name matching `^VEERG_.+_Result_Method[12]$`, groups Method1/
  Method2 pairs by shared base name (stripped of the `_Result_MethodN`
  suffix), prints grouped by module then calc sheet as ready-to-paste
  `=VEERG_..._Result_MethodN` text with the source cell noted.
- Deliberately does NOT classify Scope 1/2/3 - that's the author's
  VEERG-methodology call when pasting, not something to guess at.

### Unit lookup: read the ADJACENT CELL, not the `.xlf`/`_Unit` metadata (2026-08)
First implementation resolved the unit from the equation's own `_Unit` LAMBDA
sibling name (module-prefixed, e.g. `Fertiliser_Equations_InorganicN2O.
VEERG_..._Unit`), with a fallback that stripped trailing `_Scope1`/`_Scope3`/
`_BeefCattle`-style qualifiers from the result name to find a shared base
equation's unit (needed because e.g. `..._EmissionsFromWasteManagement_Scope1_
Result_Method1` and `..._Scope3_Result_Method1` share ONE underlying equation/
unit). Still left 2 results genuinely unresolved (`?`) - equations with no
`_Unit` sibling at all (e.g. `MarketBasedScopeElectricityEmissionsVariation`,
confirmed no `_Arguments`/`_Source`/`_Title`/`_Unit`/`_Variable` names exist
for it at all in the source).

User's fix (better, adopted): read the cell immediately to the RIGHT of the
result cell instead - that's what the sheet itself displays as the unit.
Confirmed via direct XML inspection: F39 next to `$E$39` on Electricity's
'14.1.1-2 Purchased electricity' = "t CO2e" (formula `=E29`, cached text used
directly - no need to actually follow the formula, the cached `<v>` already
has the resolved display text). Resolved BOTH previously-unresolved cases,
0 `[?]` across PastureBeef and Dairy.

CAUGHT A REAL SOURCE BUG this way: `VEERG_5_2_1_1__2_MassOfNitrogenInOrganic
Fertiliser_Result_Method1` (Fertiliser module) points at `'5.2.1.1-2 Organic
Fert N2O'!$E$50`, but E50's actual formula calls a DIFFERENT equation
(`..._5_2_1_1__1_OrganicFertiliserApplicationN2OEmissions`, not the `__2_`
equation the result's own name claims) - a mismatched equation number baked
into the result NAME itself in `5_Fertiliser_WIP_v10.xlsx`. The old name-based
`_Unit` lookup returned the WRONG-equation's unit ("kg N", an intermediate
mass calc) because it trusted the (mislabelled) name; the adjacent-cell read
returned the CORRECT unit ("t N2O", matching what's actually in the cell)
because it reads ground truth instead of a label. Not fixed (out of scope) -
just noted here in case this mismatch bites something else later. Implemented
adjacent-cell version now the ONLY unit-resolution path (old `_Unit` lookup
removed entirely, not kept as a fallback, since it can silently return a
wrong-but-plausible-looking unit).

Implementation: `Open-WorkbookContext` opens the source workbook zip once per
module (FileShare.ReadWrite so it works even if the file is open elsewhere),
returns sheet-name -> part-path map + sharedStrings + definedNames; sheet XML
docs are lazily loaded and cached per sheet since multiple results often share
one calc sheet. `Get-AdjacentCellAddr` does the column-letter+1 arithmetic on
a `$`-stripped A1 ref. `Get-CellUnitText` resolves `t="s"` (shared string),
`t="str"` (formula-cached string), and `t="inlineStr"`; anything else
(numeric/empty adjacent cell) returns null -> reported as `[?]`.

PS5.1 gotcha hit: inline `if(){}else{}` inside a `-f` format-string argument
position throws "term 'if' is not recognized" (same class as the
test-scenarios.md note) - compute into a variable first.
