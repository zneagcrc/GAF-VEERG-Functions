# nav-menu.ps1 — shared column-A navigation menu generator

Dot-sourced by `build-enterprise-excel.ps1`, `build-scope3-excel.ps1`,
`sync-xlf-to-excel-labs.ps1` (the `build.ps1 -Menu` path). Rebuilds the left-hand
menu (`=HYPERLINK("#'Sheet'!A1","Label")`) on every worksheet from the final sheet
set, freezes column A.

## "Back to top" links (2026-09)
Column A is frozen, but the MENU ITEMS (rows 3..N) still scroll out of view
vertically on long sheets, leaving a blank frozen gutter. So the generator also
drops `=HYPERLINK("#A1","▲ Back to top")` (style `Back to top link`, a named style
already in the module workbooks; `▲` = U+25B2) down column A:
- only if the last value in **columns B onward** is > 10 rows below the bottom
  menu item;
- first link on the 11th row below the menu bottom;
- then one every 30 rows down to that last value.
Constants: `$bttGapBelowMenu = 10`, `$bttInterval = 30`. `#A1` (bare, no sheet
qualifier) = A1 of the sheet the link sits on - matches the pre-existing
hand-authored cells. Column A below the logo is cleared to its last used cell
first, so old / hand-placed back-to-top cells don't survive as orphans.
Perf: ~+5% on `Set-NavMenu` (measured on the 49-sheet PastureBeef enterprise:
104s -> 109s baseline vs modified).

## PowerShell gotchas
- **`switch` + `continue` inside a `foreach`**: `continue` in a scalar `switch`
  does NOT skip the code after the `switch` in the enclosing loop body. The menu
  render loop's `'blank' { $r++ ; continue }` branch therefore advances `$r` by
  **2** (branch `$r++` + the post-switch `$r++`), while `title`/`link` advance by
  1. Never derive the menu-bottom row from `$rows.Count` - track the last row an
  actual title/link is written to.
- **Last-value-row scan**: reading `Range.Value2` for a big block to find the
  bottom-most non-empty cell in columns B+ is slow. `Range.Find("*", after,
  xlValues, xlPart, xlByRows, xlPrevious)` is one native COM call and much
  faster (a formula that currently shows `""` correctly does not count).
