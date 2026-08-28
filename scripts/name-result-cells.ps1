<#
.SYNOPSIS
  Name the top-level Method 1/Method 2 result cell for every calculation
  sheet that has one, by reading it straight off the module's own 'Results'
  sheet - then those cells' formulas are repointed to use the new name.

.DESCRIPTION
  Every module workbook has a 'Results' sheet with a table whose columns are
  headed "GHG emissions method 1" / "GHG emissions method 2" (one row per
  calculation sheet in that module). Where a sheet's result is ALREADY named,
  that column holds a bare name reference, e.g.
      =VEERG_4_5_1_1__1_TotalAnnualMethaneEmissions_Result_Method1
  Where it is NOT yet named, Excel falls back to a direct single-cell
  cross-sheet reference instead, e.g.
      ='4.5.1.3-4 Direct N2O'!E297
  That fallback IS the "missing name" signal - no scanning for repeated
  function calls, no guessing from nearby "(Method 1)" labels needed. This
  script finds every such direct reference in that specific table (and ONLY
  that table - see SCOPE below), adds a
  '<Equation>_Result_Method1'/'_Result_Method2' workbook-scoped name for the
  referenced cell, and repoints the Results-sheet formula to use it.

  <Equation> is resolved from the calc sheet's OWN VEERG number prefix (e.g.
  sheet '4.5.1.3-4 Direct N2O' -> prefix '4.5.1.3' -> 'VEERG_4_5_1_3'): the
  workbook's defined names (bare, or module-prefixed AFE names) are searched
  for one whose local name matches 'VEERG_4_5_1_3__1_<Name>' - the "__1"
  suffix specifically, since that is always the section's top-level/final
  equation (an "__2"/"__3" etc. suffix is an intermediate helper equation,
  never the one to publish as a Result). If zero or more than one such name
  is found, that row is skipped and reported rather than guessed at.

  SCOPE (deliberately narrow): only the table whose header cells literally
  read "GHG emissions method 1" / "GHG emissions method 2" is touched. A
  module's Results sheet may also have OTHER breakdown tables lower down
  (e.g. a "Manure nitrogen use" allocation table) - those have different
  column headers ("Amount of N in manure", "N unit", ...) and are never
  matched by this scan, so they are left alone entirely; naming a
  SUM-of-breakdown or multi-component bottom line is a separate, harder
  problem this script does not attempt.

  DRY-RUN by default: names + rewrites are performed IN MEMORY and reported,
  the file is NOT written. Pass -Commit to save. Add -Backup to also write a
  one-time backup under Excel/Backups/Backup_PreResultName/ before saving.

  After -Commit, consider also running `npm run apply-names:commit` - it
  sweeps the WHOLE workbook (not just the Results sheet) for any other raw
  reference to a now-named cell and converts it too.

.PARAMETER WorkbookPath
  Full path to a single .xlsx to process. If omitted, every top-level
  Excel/*.xlsx (excluding lock files, *_expanded* outputs and *.bak) is
  processed - Result names belong in the source MODULE workbooks (where the
  'Results' sheet and calc sheets live), not in the assembled enterprise
  workbooks under Excel/Enterprises/, which inherit them automatically on
  their next build-enterprise run.

.PARAMETER Commit
  Actually write the names / rewrites and save. Without it, reports only.

.PARAMETER Backup
  When saving (-Commit), first write a one-time backup copy of the workbook
  under Excel/Backups/Backup_PreResultName/. Off by default.

.EXAMPLE
  npm run name-result-cells
  npm run name-result-cells:commit
  npm run name-result-cells -- -WorkbookPath .\Excel\4_5_ManureManagement_Swine_WIP_v06.xlsx
#>
param(
  [string] $RepoRoot = (Split-Path $PSScriptRoot -Parent),
  [string] $WorkbookPath,
  [switch] $Commit,
  [switch] $Backup
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Workbook discovery.
# ---------------------------------------------------------------------------
function Get-TargetWorkbooks {
  param([string] $RepoRoot, [string] $WorkbookPath)
  if (-not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
    return @((Resolve-Path -LiteralPath $WorkbookPath).Path)
  }
  $excelDir = Join-Path $RepoRoot 'Excel'
  if (-not (Test-Path -LiteralPath $excelDir)) { throw "Excel folder not found: $excelDir" }
  Get-ChildItem -LiteralPath $excelDir -Filter '*.xlsx' -File |
    Where-Object {
      $_.Name -notlike '~$*' -and
      $_.BaseName -notmatch '(?i)_expanded' -and
      $_.Name -notmatch '(?i)\.bak'
    } |
    Sort-Object FullName |
    ForEach-Object { $_.FullName }
}

# ---------------------------------------------------------------------------
# Formula-shape helpers. A candidate cell's ENTIRE formula must be one of:
#   - a bare identifier that is an EXISTING defined name (already named - skip)
#   - a single-cell cross-sheet reference (the "missing name" case)
# Anything else (SUM, IF, multi-term, range, ...) is out of scope on purpose.
# ---------------------------------------------------------------------------
$script:BareNameRegex = [regex] '^[A-Za-z_][A-Za-z0-9_.]*$'
$script:SingleCellRefRegex = [regex] "^(?:'(?<qsheet>[^']+)'|(?<sheet>[A-Za-z_][A-Za-z0-9_.]*))!\`$?(?<col>[A-Za-z]{1,3})\`$?(?<row>[0-9]+)$"
# The section's top-level equation: "__1" specifically (an "__2"/"__3" etc.
# suffix is always an intermediate helper, never the one to publish).
function Get-TopLevelEquationRegex {
  param([string] $NumPrefix)   # e.g. "4_5_1_3"
  return [regex]::new('(?:^|\.)VEERG_' + [regex]::Escape($NumPrefix) + '__1_([A-Za-z0-9]+)$')
}

function Get-VeergNumberPrefix {
  # Leading digit.digit... run from a calc sheet's display name, dots -> underscores.
  # '4.5.1.3-4 Direct N2O' -> '4_5_1_3'   '4.5.1.9 Manure applied to soils' -> '4_5_1_9'
  param([string] $SheetName)
  $m = [regex]::Match($SheetName, '^(?<num>[0-9]+(?:\.[0-9]+)*)')
  if (-not $m.Success) { return $null }
  return ($m.Groups['num'].Value -replace '\.', '_')
}

# ---------------------------------------------------------------------------
# Process a single workbook.
# ---------------------------------------------------------------------------
function Invoke-Workbook {
  param($Excel, [string] $Path, [switch] $Commit)

  Write-Host ''
  Write-Host ('=' * 78)
  Write-Host ("Workbook : {0}" -f (Split-Path $Path -Leaf))

  $wb = $Excel.Workbooks.Open($Path, 0, $false)   # read-write; we control saving
  $namedCount = 0
  $rewriteCount = 0
  $skippedAmbiguous = 0

  try {
    $resultsWs = $null
    foreach ($ws in $wb.Worksheets) { if ([string] $ws.Name -eq 'Results') { $resultsWs = $ws; break } }
    if ($null -eq $resultsWs) {
      Write-Host "  No 'Results' sheet - skipped." -ForegroundColor DarkGray
      return [pscustomobject]@{ Named = 0; Rewrites = 0; Skipped = 0 }
    }

    # Existing workbook-scoped names, both for "already named" checks and for
    # resolving each sheet's __1 equation name. Key = local (post-'.') name.
    $namesByLocal = @{}          # local name -> Name COM object (first wins)
    $namesByRefersTo = @{}       # normalised RefersTo -> Name COM object
    foreach ($n in $wb.Names) {
      $nm = $null; $rt = $null
      try { $nm = [string] $n.Name; $rt = [string] $n.RefersTo } catch { continue }
      if ([string]::IsNullOrWhiteSpace($nm)) { continue }
      $local = $nm; $dot = $nm.LastIndexOf('.'); if ($dot -ge 0) { $local = $nm.Substring($dot + 1) }
      if (-not $namesByLocal.ContainsKey($local)) { $namesByLocal[$local] = $n }
      if (-not [string]::IsNullOrWhiteSpace($rt)) {
        $key = ($rt -replace '\s', '').ToUpperInvariant()
        if (-not $namesByRefersTo.ContainsKey($key)) { $namesByRefersTo[$key] = $n }
      }
    }

    # ---- Pass 1: find the "GHG emissions method 1/2" header cell pairs. ---
    # A Results sheet can have more than one such table (e.g. repeated per
    # scope); scan for every occurrence rather than assuming just one.
    $ur = $resultsWs.UsedRange
    $rows = [int] $ur.Rows.Count
    $cols = [int] $ur.Columns.Count
    $baseRow = [int] $ur.Row
    $baseCol = [int] $ur.Column
    $vals = $ur.Value2

    $headerPairs = New-Object System.Collections.Generic.List[object]
    for ($r = 1; $r -le $rows; $r++) {
      $m1Col = -1; $m2Col = -1
      for ($c = 1; $c -le $cols; $c++) {
        $v = if ($rows -eq 1 -and $cols -eq 1) { $vals } else { $vals.GetValue($r, $c) }
        if ($null -eq $v -or $v -isnot [string]) { continue }
        $t = $v.Trim()
        if ($t -imatch '^GHG emissions method 1$') { $m1Col = $c }
        elseif ($t -imatch '^GHG emissions method 2$') { $m2Col = $c }
      }
      if ($m1Col -gt 0 -and $m2Col -gt 0) {
        $headerPairs.Add([pscustomobject]@{ Row = $baseRow + $r - 1; M1Col = $baseCol + $m1Col - 1; M2Col = $baseCol + $m2Col - 1 })
      }
    }

    if ($headerPairs.Count -eq 0) {
      Write-Host "  No 'GHG emissions method 1/2' table found on Results - skipped." -ForegroundColor DarkGray
      return [pscustomobject]@{ Named = 0; Rewrites = 0; Skipped = 0 }
    }

    # ---- Pass 2: walk each table downward from its header, one row per calc
    #      sheet, until a row with nothing in either method column. --------
    $candidates = New-Object System.Collections.Generic.List[object]
    foreach ($hp in $headerPairs) {
      $r = $hp.Row + 1
      while ($true) {
        $c1 = $resultsWs.Cells.Item($r, $hp.M1Col)
        $c2 = $resultsWs.Cells.Item($r, $hp.M2Col)
        $f1 = $null; $f2 = $null
        try { $f1 = [string] $c1.Formula2 } catch { }
        try { $f2 = [string] $c2.Formula2 } catch { }
        if ([string]::IsNullOrWhiteSpace($f1) -and [string]::IsNullOrWhiteSpace($f2)) { break }   # end of table
        foreach ($pair in @(@{ Cell = $c1; F = $f1; Method = 1 }, @{ Cell = $c2; F = $f2; Method = 2 })) {
          if ([string]::IsNullOrWhiteSpace($pair.F) -or $pair.F[0] -ne '=') { continue }
          $body = $pair.F.Substring(1).Trim()
          if ($script:BareNameRegex.IsMatch($body)) { continue }   # already a name reference
          $rm = $script:SingleCellRefRegex.Match($body)
          if (-not $rm.Success) { continue }   # not a plain single-cell ref (SUM/IF/etc.) - out of scope
          $targetSheet = if ($rm.Groups['qsheet'].Success) { $rm.Groups['qsheet'].Value } else { $rm.Groups['sheet'].Value }
          $resultsCol = if ($pair.Method -eq 1) { $hp.M1Col } else { $hp.M2Col }
          $candidates.Add([pscustomobject]@{
            ResultsRow = $r; ResultsCol = $resultsCol
            Method = $pair.Method; TargetSheet = $targetSheet
            TargetCol = $rm.Groups['col'].Value.ToUpperInvariant(); TargetRow = $rm.Groups['row'].Value
          })
        }
        $r++
      }
    }

    if ($candidates.Count -eq 0) {
      Write-Host '  Nothing to name (every row already named, or no plain single-cell references found).' -ForegroundColor DarkGray
      return [pscustomobject]@{ Named = 0; Rewrites = 0; Skipped = 0 }
    }

    # ---- Pass 3: resolve each candidate's __1 equation name, add the name,
    #      queue the Results-sheet rewrite. ---------------------------------
    $rewrites = New-Object System.Collections.Generic.List[object]
    $equationCache = @{}   # VEERG number prefix -> equation name ($null = ambiguous/not found)

    foreach ($cand in $candidates) {
      $numPrefix = Get-VeergNumberPrefix -SheetName $cand.TargetSheet
      if ($null -eq $numPrefix) {
        Write-Host ("  [skip]  {0}!{1}{2} -> could not derive a VEERG number from sheet name '{3}'" -f $resultsWs.Name, (ConvertTo-ColLettersLocal $cand.ResultsCol), $cand.ResultsRow, $cand.TargetSheet) -ForegroundColor Yellow
        $skippedAmbiguous++
        continue
      }
      if (-not $equationCache.ContainsKey($numPrefix)) {
        $re = Get-TopLevelEquationRegex -NumPrefix $numPrefix
        $found = @()
        foreach ($local in $namesByLocal.Keys) {
          if ($local -match '_Result_Method[12]$') { continue }   # never treat an existing Result name as the source equation
          $m = $re.Match($local)
          # Keep the FULL "VEERG_<prefix>__1_<name>" as the base - matching
          # the existing naming convention (e.g. VEERG_4_5_1_1__1_Total...) is
          # not cosmetic: list-enterprise-results.ps1 (and anything else that
          # looks for a result cell) matches names via `^VEERG_`, so a name
          # missing that prefix is silently invisible to every consumer.
          if ($m.Success) { $found += ('VEERG_' + $numPrefix + '__1_' + $m.Groups[1].Value) }
        }
        $found = @($found | Select-Object -Unique)
        $equationCache[$numPrefix] = if ($found.Count -eq 1) { $found[0] } else { $null }
        if ($found.Count -eq 0) {
          Write-Host ("  [skip]  '{0}': no VEERG_{1}__1_* equation name found in this workbook" -f $cand.TargetSheet, $numPrefix) -ForegroundColor Yellow
        } elseif ($found.Count -gt 1) {
          Write-Host ("  [skip]  '{0}': ambiguous - {1} different VEERG_{2}__1_* equation names found ({3})" -f $cand.TargetSheet, $found.Count, $numPrefix, ($found -join ', ')) -ForegroundColor Yellow
        }
      }
      $equationName = $equationCache[$numPrefix]
      if ($null -eq $equationName) { $skippedAmbiguous++; continue }

      $refersTo = "='" + ($cand.TargetSheet -replace "'", "''") + "'!$" + $cand.TargetCol + '$' + $cand.TargetRow
      $wantedName = "{0}_Result_Method{1}" -f $equationName, $cand.Method

      # If the target cell is ALREADY named under some other name, reuse it
      # rather than creating a duplicate name for the same cell.
      $refKey = ($refersTo -replace '\s', '').ToUpperInvariant()
      $finalName = $wantedName
      if ($namesByRefersTo.ContainsKey($refKey)) {
        $finalName = [string] $namesByRefersTo[$refKey].Name
        $dot = $finalName.LastIndexOf('.'); if ($dot -ge 0) { $finalName = $finalName.Substring($dot + 1) }
      } elseif ($namesByLocal.ContainsKey($wantedName)) {
        $existingRt = ''
        try { $existingRt = [string] $namesByLocal[$wantedName].RefersTo } catch { }
        if (($existingRt -replace '\s', '').ToUpperInvariant() -ne $refKey) {
          Write-Host ("  [SKIP]  '{0}' already exists pointing elsewhere ({1}); {2}!{3}{4} left as-is" -f $wantedName, $existingRt, $resultsWs.Name, $cand.TargetCol, $cand.TargetRow) -ForegroundColor Yellow
          $skippedAmbiguous++
          continue
        }
      } else {
        try {
          [void] $wb.Names.Add($wantedName, $refersTo)
          $namesByLocal[$wantedName] = $wb.Names.Item($wantedName)
          $namesByRefersTo[$refKey] = $namesByLocal[$wantedName]
          $namedCount++
          Write-Host ("  [name]  '{0}'!{1}{2}  ->  {3}" -f $cand.TargetSheet, $cand.TargetCol, $cand.TargetRow, $wantedName) -ForegroundColor Green
        } catch {
          Write-Host ("  [ERR]   could not add name '{0}' for '{1}'!{2}{3}: {4}" -f $wantedName, $cand.TargetSheet, $cand.TargetCol, $cand.TargetRow, $_.Exception.Message) -ForegroundColor Red
          $skippedAmbiguous++
          continue
        }
      }

      $rewrites.Add([pscustomobject]@{ Row = $cand.ResultsRow; Col = $cand.ResultsCol; Name = $finalName })
    }

    # ---- Pass 4: repoint the Results-sheet cells that led us here. --------
    foreach ($rw in $rewrites) {
      $cell = $resultsWs.Cells.Item($rw.Row, $rw.Col)
      $old = ''
      try { $old = [string] $cell.Formula2 } catch { }
      $new = '=' + $rw.Name
      if ($old -eq $new) { continue }
      $rewriteCount++
      Write-Host ("  [ref]   {0}!{1}{2}  {3}  ->  {4}" -f $resultsWs.Name, (ConvertTo-ColLettersLocal $rw.Col), $rw.Row, $old, $new) -ForegroundColor Cyan
      if ($Commit) { try { $cell.Formula2 = $new } catch { Write-Host ("  [ERR]   could not rewrite {0}!{1}{2}: {3}" -f $resultsWs.Name, $rw.Col, $rw.Row, $_.Exception.Message) -ForegroundColor Red } }
    }

    Write-Host ("  Summary: {0} named, {1} Results-sheet references rewritten, {2} skipped (ambiguous/unresolvable)" -f $namedCount, $rewriteCount, $skippedAmbiguous)

    if ($Commit -and ($namedCount -gt 0 -or $rewriteCount -gt 0)) {
      if ($Backup) {
        $backupDir = [IO.Path]::Combine((Split-Path $Path -Parent), 'Backups', 'Backup_PreResultName')
        if (-not (Test-Path -LiteralPath $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
        $bak = [IO.Path]::Combine($backupDir, [IO.Path]::GetFileName($Path))
        if (-not (Test-Path -LiteralPath $bak)) {
          Copy-Item -LiteralPath $Path -Destination $bak -Force
          Write-Host ("  Backup : {0}" -f (Join-Path 'Backups\Backup_PreResultName' (Split-Path $bak -Leaf))) -ForegroundColor DarkGray
        }
      }
      $wb.Save()
      Write-Host '  Saved.' -ForegroundColor Green
    }

    return [pscustomobject]@{ Named = $namedCount; Rewrites = $rewriteCount; Skipped = $skippedAmbiguous }
  }
  finally {
    try { $wb.Close($false) } catch { }
    if ($null -ne $wb) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
  }
}

function ConvertTo-ColLettersLocal {
  param([int] $Col)
  $s = ''
  while ($Col -gt 0) {
    $rem = ($Col - 1) % 26
    $s = [char](65 + $rem) + $s
    $Col = [int](($Col - $rem - 1) / 26)
  }
  return $s
}

# ---------------------------------------------------------------------------
# Main.
# ---------------------------------------------------------------------------
$workbooks = @(Get-TargetWorkbooks -RepoRoot $RepoRoot -WorkbookPath $WorkbookPath)
$modeLabel = if ($Commit) { 'COMMIT' } else { 'DRY-RUN (no files written)' }
Write-Host ("Mode      : {0}" -f $modeLabel)
Write-Host ("Workbooks : {0}" -f $workbooks.Count)

function New-FreshExcelApp {
  $app = New-Object -ComObject Excel.Application
  $app.Visible = $false
  $app.DisplayAlerts = $false
  $app.AskToUpdateLinks = $false
  $app.ScreenUpdating = $false
  return $app
}

$excel = $null
$totNamed = 0; $totRewrites = 0; $totSkipped = 0
try {
  # A FRESH Excel COM instance per workbook, not one shared across the whole
  # batch: this repo has previously hit COM state degrading across many
  # sequential workbook opens in one session (see excel-com-errors.md,
  # build-input-fields-json.ps1's "SILENT DATA LOSS" note) - confirmed here
  # too: 3 workbooks that threw "cannot call a method on a null-valued
  # expression" when processed in a shared-session batch ran clean in
  # isolation. Slower (Excel startup overhead per file) but reliable.
  foreach ($path in $workbooks) {
    try {
      if ($null -eq $excel) { $excel = New-FreshExcelApp }
      $res = Invoke-Workbook -Excel $excel -Path $path -Commit:$Commit
      $totNamed += $res.Named
      $totRewrites += $res.Rewrites
      $totSkipped += $res.Skipped
    } catch {
      Write-Host ("  [FATAL] {0}: {1} (line {2}: {3})" -f (Split-Path $path -Leaf), $_.Exception.Message, $_.InvocationInfo.ScriptLineNumber, $_.InvocationInfo.Line.Trim()) -ForegroundColor Red
    } finally {
      if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch { }
        $excel = $null
      }
    }
  }
}
finally {
  if ($null -ne $excel) {
    try { $excel.Quit() } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch { }
  }
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}

Write-Host ''
Write-Host ('=' * 78)
Write-Host ("TOTAL: {0} cell(s) named, {1} Results-sheet reference(s) rewritten, {2} skipped across {3} workbook(s)." -f `
    $totNamed, $totRewrites, $totSkipped, $workbooks.Count)
if ($Commit -and $totNamed -gt 0) {
  Write-Host ''
  Write-Host 'Consider running `npm run apply-names:commit` next - it sweeps every' -ForegroundColor DarkGray
  Write-Host 'workbook (not just each Results sheet) for other raw references to the' -ForegroundColor DarkGray
  Write-Host 'cells just named and converts those too.' -ForegroundColor DarkGray
}
