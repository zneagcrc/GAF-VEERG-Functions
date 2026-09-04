# Enterprise test-input builder

`scripts/build-enterprise-testinput-json.ps1` (npm `build-enterprise-testinput`[`:dry`]).
Merges per-module `Test/**/TestInput_*.json` into one enterprise test input.

- Module selection = enterprise config `modules` array (Enterprises/Enterprise_PastureBeef.json).
- Mapping chain: module id -> registry `_ModuleRegistry.json` `sourceWorkbook` -> version-stripped
  base -> Test/Test.json entry by exact or token-boundary prefix match (longest wins).
  This prefix match handles `14_Electricity_Scope2` (source) -> `14_Electricity` (TestExcelFile).
- Conflict winner = config-designated master: module WITHOUT a `renameSheets` entry beats one
  WITH it (rank = (hasRename?1000:0)+configIndex, lower wins; ties = config order). Enteric renames
  'Input - Pasture Beef' so Manure wins shared X_Cell_PastureBeef_* fields AND shared X_Table_*
  tables. Applies to both scalars and InputTables (duplicate TableName keeps winner's copy).
- X_Cell_Site_* overridden by options.commonSheetProviders["Input - Site"] (=ManureManagement_BeefPasture);
  mid-merge site conflicts are suppressed (provider override is authoritative).
- Output: Test/Enterprises/TestInput_Enterprise_<enterprise.id>.json, TAB-indented via custom
  ConvertTo-TabJson (NOT ConvertTo-Json — PS5.1 space quirks), UTF-8 no BOM.
- PastureBeef result: 42 cells, 19 tables.
- Per-module `include` subsetting in the enterprise config is NOT applied to test input (full
  module TestInput merged; extra cells/tables with no enterprise named range are ignored downstream).

## Enterprise Test/ entries (2026-09)
`Test/Test.json` now has a `Tests.Enterprises` map (PastureBeef, EnvironmentalPlantings,
CroppingCotton, CroppingGrains, Dairy, Feedlot, Poultry, Swine), each ->
`Test/Enterprises/<Name>/TestInput_Enterprise_<Name>.json` + `TestResults_...`.
How they were seeded (starting point only — values need a manual pass):
- TestInput = `build-enterprise-testinput` module-merge, then gap-filled from
  `InputFields/Enterprise_<Id>_InputFields.json`: number->1, percent->"10%",
  select->first option, text/date->"", `formula` cells skipped, `X_Table_Result_*`
  skipped; table skeletons = Row1 populated by CellType, rest null, N rows =
  `NumberOfRows` or 3.
- TestResults = every `VEERG_*_Result_*` defined-name in the built enterprise
  workbook that also appears in a merged module `TestResults_*.json` (with an
  `_CropOnly`/`_PastureOnly` infix-strip fallback), value from the module file;
  plus every enterprise-level `Result_*` roll-up name as `0`. Results with no
  module oracle are omitted. NOT computed via Excel.
- Feedlot had no `Enterprise_Feedlot_InputFields.json` -> its TestInput is the raw
  module-merge, no gap-fill.
