# create-test-excel-from-json.ps1 — scenarios, batch modes, follow-on tests

`npm run create-test-excel -- -TestID <id> [...]`. Copies the source workbook,
writes the TestInput cells/tables into named ranges, recalculates, compares the
computed result named-cells against the TestResults file (tolerance
`-DifferenceTolerance`, default 1e-5). Output: `Excel/TestExcel/<name>_test_<ts>.xlsx`.

## Named scenarios

TestInput / TestResults JSON may hold a top-level `Scenarios` array. Each scenario:
`ScenarioName` + flat fields (`X_Cell_*` for inputs / result-cell names for
results) + optional `InputTables` + optional `ScenarioDescription`.

- `-ScenarioName "Scenario 2"` selects by name (case-insensitive); omitted => first.
- Output filename gets a sanitized `_<ScenarioName>` tag before the timestamp.
- Files with NO `Scenarios` key keep the old flat behaviour.
- Results scenario is matched by explicit `-ScenarioName`, else the input's
  resolved scenario name — so TestResults scenario names must line up with
  TestInput's.

### Inheritance — `ScenarioExtends` (2026-09)
Per-scenario, replaces the old implicit "everything inherits scenario[0]":

| `ScenarioExtends` | behaviour |
|---|---|
| `"<name>"` | deep-merge this scenario onto that named one (resolved first, so chains `C→B→A` work) |
| `null` / `""` | standalone — no inheritance |
| key absent | **legacy**: first scenario is the standalone default, every other deep-merges onto it |

A file may mix all three. Unknown parent name or a circular chain = hard error
(`Resolve-Scenario` tracks a visiting set). Deep-merge is unchanged
(`Merge-ScenarioOnDefault` → `Merge-InputTables` by `TableName` → `Merge-ColumnMap`
by column/field; scenario[0]/parent deep-cloned first, never mutated).

## Batch modes (2026-09)

Instead of one `-TestID`, run many entries, each in its own child `powershell`
process (a crash/OOM in one doesn't abort the rest); per-entry PASS/FAIL summary,
exit 1 if any failed. None may combine with `-TestID`.

| switch | scope | npm |
|---|---|---|
| `-ModulesOnly` | every TestID **not** under `Tests.Enterprises` | `create-test-excel:modules` |
| `-EnterprisesOnly` | every TestID under `Tests.Enterprises` | `create-test-excel:enterprises` |
| `-All` | both | `create-test-excel:all` |

- Partition: `Get-TestIdsUnderNode` recursively collects `TestID`s under
  `$config.Tests.Enterprises` (enterprise set) vs the rest (module set) — enteric/
  manure per-flavour entries land in the module set.
- Every other passed parameter is forwarded to each child; `-TestID` and the batch
  switches (and `-TestInputPath`/`-TestResultsPath`) are not.
- `-ShowFailuresOnly`: single `-TestID` => the "Result differences" list shows only
  the FAIL rows (summary line still prints); batch => passing entries print
  nothing, only failing entries show their output, summary lists only failures.

## Follow-on tests (`-RunFollowOns`, 2026-09)

Chain a computed result of one test into another. In a resolved TestInput (flat
file or the selected scenario) add `FollowOnTests: [ { ... } ]`:

- `TestID` (downstream Test.json entry), optional `ScenarioName` (downstream base
  scenario, default = its first/flat), `PassOnResults` (array of THIS test's
  result named-cells to harvest), `Inputs` (downstream flat overrides),
  `InputTables` (downstream table overrides, same shape as a scenario's),
  `ExpectedResults` (`{name: number}` the follow-on asserts).
- With `-RunFollowOns`: after this test runs, its `PassOnResults` cells are read
  from the still-open workbook (`Convert-ToNullableDouble` of `RefersToRange.Value2`);
  then each entry runs the downstream test against **downstream default TestInput
  (or `ScenarioName`) + `Inputs` + `InputTables`** layered on top, with any value
  string — flat OR inside a table cell — that equals a `PassOnResults` name
  replaced by the harvested number (`Convert-FollowOnPlaceholders`, recursive).
- It's a normal child run against synthesized temp input/results files, via new
  params `-TestInputPath` / `-TestResultsPath` (also usable directly: "run this
  TestID with my file"). Temp files land in `Excel/TestExcel/_followon_*` and are
  deleted after.
- **One level deep** — the child is NOT given `-RunFollowOns`.
- Parent exits non-zero if its own results OR any follow-on fails. Without
  `-RunFollowOns` the follow-ons are noted and skipped.
- `FollowOnTests` / `ScenarioDescription` / `ScenarioExtends` are in the input
  skip-list and the results skip-check so they aren't treated as workbook cells.

## InputTables `"Replace": true` (2026-09)

On an `InputTables` entry (sibling of `TableName`), swap the base table's
`Rows`/`Cols` out **wholesale** instead of merging by row key. Needed when your
row keys differ from the base's — otherwise `Merge-ColumnMap` APPENDS your rows
and they overflow the Excel table's fixed `DataBodyRange` height →
`Table_X: row position N is outside range height M`. Works in a follow-on's
`InputTables` and in a `ScenarioExtends` override. The marker is stripped before
the file is written (`Remove-ReplaceMarker` / a `continue` in `Merge-InputTables`).
Example that bit us: `5_Fertiliser`'s default `Table_Input_OrganicFertiliser` has
3 rows `Lime1/2/3`; a follow-on row keyed `OrganicFertiliser1` needs `Replace:true`.

## RowsToCols table writer — column mapping (IMPORTANT for authoring rows)

Under `-Context Test` (the DEFAULT, and what follow-ons use) `$overwriteAllFormulas`
is true, so the writer maps each row field to a column **positionally, 1:1 in
table order** — it does NOT match by field name.

- A calculated/formula column you don't want to write must be present as `null`
  (its slot is consumed, the write skipped). **Omitting a non-trailing column
  shifts every later field one column left.** Trailing calc columns may be omitted.
- Under `-Context Emissions` the writer instead probes for non-formula columns and
  maps the Nth non-null field to the Nth non-formula column (the app omits
  formula fields there).
- Bug fixed 2026-09 (`Table_Input_RECConsumedByEntity`, cols `Source`/`MWh`/
  `Elegibility` are Excel calculated columns; test row had them `null`): the old
  code dropped null fields but still used a `1..N` positions list, so
  `GenerationDate` landed in the `Source` column. Fix: in the formula-overwrite
  path KEEP null fields for positional accounting and skip only the write
  (`if ($null -eq $prop.Value) { continue }`).
- `Table_Input_OrganicFertiliser` (5_Fertiliser, `D53:J56`): cols 4-7
  (`AmountNAppliedPerHectare`, `ApplicationArea` = `X_Cell_Fertiliser_AreaUnderCropping`,
  `TotalOrganicFertiliserApplied`, `TotalNApplied`) are calc — set only cols 1-3
  and feed area via the `X_Cell_Fertiliser_AreaUnderCropping` input.

## Implementation (scripts/create-test-excel-from-json.ps1)
- Scenario helpers near the top: `Copy-JsonObject` (JSON round-trip clone),
  `Set-ObjProp`, `Merge-ColumnMap`, `Merge-InputTables`, `Remove-ReplaceMarker`,
  `Merge-ScenarioOnDefault`, `Get-ScenarioByName`, `Resolve-Scenario`,
  `Select-ScenarioObject`, then `Convert-FollowOnPlaceholders`, `Invoke-FollowOnTest`.
- Batch dispatcher runs right after config resolution, before the TestID-required
  check; `exit`s so the single-entry body is never entered.
- Follow-on block runs after the summary Write-Hosts, before the final `throw`.
- Docs: `SCRIPTS.md` "## npm run create-test-excel".

## PowerShell 5.1 gotchas hit here
- **Native call output pollutes a function's return**: `& powershell @args` inside
  a function emits the child's stdout lines into the function's output stream, so
  they become part of the return value — a caller's `if ($returned)` then sees a
  non-empty array and is always truthy. Route child streams to the host:
  `& powershell @args 2>&1 | ForEach-Object { Write-Host $_ }` (or `| Out-Host`).
- **`$x = if (c) { @() } else {...}`**: the empty-array branch collapses to `$null`
  under StrictMode Latest, then `$x.Count` throws. Assign `@(...)` directly, then
  conditionally reassign.
- **`$Obj.PSObject.Properties.Name -contains 'X'`** throws "property 'Name' cannot
  be found" under StrictMode Latest when `$Obj` has ZERO properties. Use the
  indexer `$Obj.PSObject.Properties['X']` (`$null` if absent) — see `Set-ObjProp`.
- **`2>&1` capture of a native command under `$ErrorActionPreference='Stop'`**:
  child stderr is wrapped as a terminating `NativeCommandError` that aborts the
  enclosing loop. A dispatcher that checks `$LASTEXITCODE` itself must set
  `$ErrorActionPreference='Continue'` around the child invocations. Plain
  un-redirected `& powershell` does NOT abort — only `2>&1` / `*>` capture does.
- npm passes `-Foo a,b` as ONE literal token — split on `[,;\s]+` for a list.
- `if(){}else{}` works as an assignment expression but NOT inside a string / `-f`
  argument position ("term 'if' is not recognized") — compute into a var first.
