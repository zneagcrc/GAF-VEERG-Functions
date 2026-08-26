# Shared build modules: nav-menu.ps1 + worksheet-view.ps1

Dot-sourced helpers under `scripts/`, used by build-enterprise-excel.ps1,
build-scope3-excel.ps1, and the generic chapter build sync-xlf-to-excel-labs.ps1.
Pattern: `. (Join-Path $PSScriptRoot 'name.ps1')` near the other dot-sources.

## scripts/nav-menu.ps1 (COM-based; caller passes an OPEN workbook)
- `Set-NavMenu -Target <wb COM> -CategoryMap <hashtable sheet->category> -Labels <hashtable>`
  rebuilds the column-A menu on EVERY sheet. Returns count of sheets updated.
  - Groups (render order): untitled (Home,Overview,Results) -> INPUTS (Input*) ->
    CALCULATIONS (category 'calculation') -> APPENDICES (else).
  - Untitled order forced: Home, Overview, Results.
  - Styles used: `Menu section title`, `Menu link default`,
    `Menu link selected results|input|equations`, `Menu link selected`,
    `Input page heading` (A1 logo, copied from a donor sheet).
  - Clears A2:A160, writes `=HYPERLINK("#'Sheet'!A1","Label")` then applies style
    (HYPERLINK auto-applies built-in Hyperlink style, so set menu style AFTER).
  - Style assignment wrapped in try/catch -> missing styles just skip styling.
- `Get-InferredCategoryMap -Workbook <wb COM>` -> hashtable by tab-name convention:
  `Input*`->input; `Constants*`->constants; `^\s*\d` OR `Calc*`->calculation.
  (Home/Overview/Results handled by name inside Set-NavMenu.)

## scripts/worksheet-view.ps1 (XML-only; headless-safe)
- `Set-WorkbookZoom -Path <xlsx> [-Zoom 100]` returns count of changed sheet parts.
  - Loads System.IO.Compression.FileSystem itself (self-contained).
  - Opens zip 'Update'; for each `^xl/worksheets/sheet\d+\.xml$` rewrites every
    `<sheetView ...>` start tag (regex `'<sheetView\b[^>]*?/?>'`): strips existing
    `zoomScale(Normal|Sheet|PageLayoutView)?="..."`, appends
    ` zoomScale="N" zoomScaleNormal="N"`. Two-phase (read all, then delete+recreate).
    UTF8 no BOM.
  - Wired into all 3 builders, guarded `if (-not $DryRun)`, after the final XML
    passes / near summary. COM Save preserves the XML-set zoom.

## Chapter build nav-menu integration (sync-xlf-to-excel-labs.ps1)
- OPT-IN via `-Menu` switch (default OFF; normal `npm run build` does not rewrite
  menus). `build.ps1` threads `-Menu` through; npm script `build:menu`.
- `Update-ChapterNavMenus -Workbooks <FileInfo[]> -DryRun` opens each workbook via
  COM, `Get-InferredCategoryMap` -> `Set-NavMenu`, Save, Close. Same COM cleanup
  (`Invoke-ComObjectCleanup`) + transient-COM handling as the propagation pass.
  Called LAST (after Sync-CommonSheetsAcrossWorkbooks) so menu reflects final sheets.
- Verified all chapter workbooks carry the 7 menu styles and tab names match the
  inference conventions (fast XML probe of styles.xml + workbook.xml, no COM).
- Scope3 seeds `Get-InferredCategoryMap` then overlays its `$ImportMap` categories
  before calling Set-NavMenu (so template calc sheets not in ImportMap classify right).

## BUG FIXED: Set-NavMenu's FreezePanes ignores the current selection on some sheets (2026-08)
SYMPTOM: Dairy enterprise sheets `'4.3.1.3-4 Direct N20'`, `'4.3.1.8 Manure
applied to soils'`, `'4.3.1.9 Soil direct N2O'` (and, spotted along the way,
`'Constants - Ag Residue'`) "not rendering the column A menu correctly" -
menu CELL CONTENT was always correct (verified identical to a working sheet:
same links, labels, 'selected'-style row), the bug was entirely in the VIEW:
`<pane xSplit="6" ySplit="17" topLeftCell="H18" state="frozen"/>` (xSplit="8"
for Ag Residue) instead of the standard `<pane xSplit="1" topLeftCell="B1"
state="frozen"/>` every other sheet got.

TWO WRONG THEORIES tried and disproven before the real cause, both worth
recording because the disproof method is the reusable lesson:
1. "Inherited from source, and FreezePanes=$true is a no-op if already
   $true." DISPROVEN by reading every `4.3.1.x` sheet's raw `sheetView` XML
   in the source workbook directly - none carry a `<pane>` at all. (Initially
   asserted this about the whole file after checking only ONE sheet in it;
   the user's pushback that no sheets in that file had frozen panes was
   correct and caught an insufficiently-verified claim.)
2. "`$ws.Activate()` doesn't land synchronously in a 50+ sheet loop, so
   FreezePanes fires while a different sheet is still active (a race)."
   Plausible-sounding and consistent with the scatter pattern, but DISPROVEN
   by running the SAME build twice with instrumentation: the exact same
   values (6/17/H18, 8/17/J18) reproduced bit-for-bit both times. A genuine
   timing race would not be perfectly deterministic across separate runs.

ACTUAL ROOT CAUSE (confirmed via inline `Write-Host` instrumentation added
directly to the freeze block, then running the real build and reading its
console output - not guessing from the saved file after the fact): activation
was fine (`landed=True attempt=0`), the pre-freeze state was clean
(`FreezePanes=False Split=False SplitColumn=0 SplitRow=0`), B1 was correctly
selected (`ActiveCell=$B$1`, correct `ActiveSheet`) - and STILL, the single
call `$win.FreezePanes = $true` immediately produced `SplitColumn=6
SplitRow=17` on these specific sheets. So `FreezePanes = $true` is not purely
"freeze at the current selection" in Excel's COM model - on some sheets
(large tables, most likely ones containing a genuine Excel Table/ListObject)
it snaps to a per-sheet "smart" position instead, deterministically, every
time, regardless of what's actually selected.

FIX, part 1 (the split itself): stop relying on ActiveCell position at all -
set `$win.SplitColumn = 1; $win.SplitRow = 0` EXPLICITLY before
`FreezePanes = $true`. This reliably bypassed the "smart" positioning.
Verified: xSplit/ySplit correct on all 54 sheets after rebuild.

FIX, part 2 (a smaller side effect the same root cause also produced): those
same sheets still had a wrong horizontal SCROLL position after the split was
otherwise fixed - `topLeftCell="C1"` instead of "B1" (a narrow ~4.7-wide
spacer column landing just out of view - cosmetic, but inconsistent with
every other sheet). Column B was NOT hidden on these sheets (ruled out).
THREE post-freeze reset attempts all failed to move it: `Window.ScrollColumn`
/ `ScrollRow`, re-`Select('B1')` after freezing, and `Window.ActivePane.
ScrollColumn` - rebuilt and re-verified after each, all still showed C1.
What actually worked: resetting `$win.ScrollColumn = 1; $win.ScrollRow = 1`
BEFORE establishing the split at all (while the sheet is still fully
unfrozen), not after. Whatever "smart" memory Excel has for these sheets
apparently seeds the new pane's initial scroll from pre-split window state,
which no post-split reset could override - only clearing it before the split
existed worked. Verified: `topLeftCell="B1"` identical on all 54 sheets.

Final freeze sequence in Set-NavMenu, in order: Activate -> clear
FreezePanes/Split if set -> reset ScrollColumn/ScrollRow to 1/1 -> Select B1
-> set SplitColumn=1/SplitRow=0 -> FreezePanes=$true. Each step earlier in
this order was tried and found insufficient alone; all are needed together.
Side effect (harmless): `state` now serializes as `"frozenSplit"` instead of
`"frozen"` on every sheet, universally, not just the previously-broken ones -
a valid Excel view state, functionally identical to the user, not worth
chasing further.

Since this is Excel's per-sheet "smart" freeze heuristic (not a per-workbook
data issue or a timing race), it can in principle affect any enterprise's
build, on whichever sheets happen to trigger it - but the fix is now
unconditional (applies the same explicit reset to every sheet regardless),
so no per-enterprise or per-sheet follow-up should be needed.

## Gotchas
- Verify zoom/menu via XML (read sheetN.xml for zoomScale / count `HYPERLINK("#`),
  NOT via headless recalc.
- File-lock IOException on rebuild = workbook open in Excel; ask user to close.
- Kill ONLY headless Excel: `Get-Process EXCEL | ? { $_.MainWindowHandle -eq 0 } | Stop-Process -Force`.
