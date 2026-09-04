- In scripts/expand-veerg-lambda-references.ps1, SourceData array results must be written as a matrix via cell offsets; stringifying spill/array values collapses multiple rows into one cell.
- Zero-arg SourceData lambdas can contain either string literals or Excel array constants; preserve both when replacing formulas with values.
- A bare `return $twoDArray` (or a bare 1-row jagged `return`) in a PS function is flattened to a 1-D array by the caller, losing row/col shape: a horizontal 1xN spill (e.g. Common_SourceData_FracLEACH_Data = MAKEARRAY(1,2,...)) then gets written down a column. Get-ExcelCellResolvedValue and Parse-ExcelArrayConstant return `, $value` to keep the shape.

## Batch modes (2026-09)
`expand-veerg-lambda-references.ps1` took `-ModulesOnly` / `-EnterprisesOnly` / `-All`
(`-SourceWorkbook` now optional; guard rejects combining them). Each discovered
workbook runs in its own child process; per-file `[OK]`/`[FAILED]` summary, exit 1
if any failed; `$ErrorActionPreference='Continue'` in the dispatcher so a failing
child doesn't abort the loop.
- `-ModulesOnly` -> `Excel\*.xlsx` top level only (~22, incl. `Common_v03`, excl.
  `13_Scope3_Template`). `-EnterprisesOnly` -> `Excel\Enterprises\**\*.xlsx` (~9:
  the `_WIP_v01` + `PastureBeef_Clean_WIP_v01`). `-All` = union.
- filter for all scopes: exclude `*_expanded*`/`*_tmp*`, `*template*`, `~$*`, `Old/`.
- npm: `expand-lambda-functions:modules` / `:enterprises` (+ `:dry`). The older
  `:auto` (whole `Excel/` tree, recursive) is unchanged.
- `-Include *.xlsx` on `Get-ChildItem` needs `-Recurse` OR a `\*` path suffix,
  else it returns nothing (used a splat that switches between the two).
