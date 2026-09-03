# restyle-fonts.ps1 — one-off font + row-height pass (2026-09)

`npm run restyle-fonts` / `:commit`. Zip + XML directly (no Excel/COM). Full docs
in SCRIPTS.md. Summary of what mattered:

- Renames "Times New Roman" -> "Arial" and shrinks font sizes (`>=12pt: -2`,
  `10-11pt: -1`, `<=9pt: keep`) EVERYWHERE a face/size lives, not just named
  styles: `xl/styles.xml` `<fonts>` + `<dxfs>`, `sharedStrings.xml` runs,
  `theme/themeN.xml` `<fontScheme>`, `drawings/drawingN.xml` + `charts/chartN.xml`
  DrawingML runs (`@sz` is 1/100 pt), `worksheets/sheetN.xml` `&"font"`/`&size`
  header-footer codes + inline-string runs, `commentsN.xml`.
- Row heights: forces an explicit `ht` + `customHeight="1"` on EVERY `<row>`
  (row 1 -> 30, others -> 20) + `sheetFormatPr/@defaultRowHeight`. `defaultRowHeight`
  alone is NOT authoritative per-row - Excel auto-fits rows that only inherit it
  (verified via COM: inheritance gave 18 / 20.25, not 20).
- The size rule is **cumulative** - a `-Commit` that resized stamps a
  `RestyleFontsApplied` custom doc property (creating `docProps/custom.xml` +
  content-type + rel if absent); later resizing runs skip a marked workbook
  unless `-Force`. Rename + row-height passes are idempotent and never marker-blocked.

## PowerShell gotchas (both cost real time)
- **`@()` around a re-surfaced `List[object]`**: `@($r.Changes)` where `$r.Changes`
  is a `System.Collections.Generic.List` returned from a function threw
  `ArgumentException: Argument types do not match` in Windows PowerShell 5.1.
  Return `.ToArray()` from the function and index/pipe the array directly.
- **`return $twoDArray` flattens**: a bare `object[,]` returned from a PS function
  is flattened to a 1-D `object[]` (row-major) by the caller - shape lost. Use
  `return , $value`. This is the same bug that made
  `expand-veerg-lambda-references.ps1` write a horizontal 1xN spill down a column
  (see expand-script-notes.md) - `Get-ExcelCellResolvedValue` and
  `Parse-ExcelArrayConstant` now `return , $value`.
