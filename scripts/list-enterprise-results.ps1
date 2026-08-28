<#
.SYNOPSIS
  Lists every emissions-result named cell (…_Result_Method1/Method2) available
  from an enterprise's selected chapter modules, with its display unit, as
  ready-to-paste formula text for building 'Results - Activity period' by hand.

.DESCRIPTION
  Read-only, XML-only (no COM, no Excel needed, safe to run while workbooks are
  open elsewhere). For each module selected in Enterprises/Enterprise_<Id>.json:
    - resolves its source workbook via _ModuleRegistry.json (version-agnostic,
      same resolution as build-enterprise-excel.ps1's Resolve-SourceWorkbook)
    - reads that workbook's workbook-scoped defined names directly from
      xl/workbook.xml
    - finds every name matching VEERG_..._Result_Method(1|2), groups Method1/
      Method2 pairs by their shared base name
    - reads the unit from the cell immediately to the right of the result cell
      (the sheet's own displayed unit, not derived from the equation name -
      more reliable than guessing from the .xlf metadata, since it's exactly
      what the sheet itself shows next to the value)
    - prints them grouped by module then calc sheet.

  Does NOT attempt to classify Scope 1/2/3 - that's a VEERG-methodology call
  the author makes when pasting; this only saves the tedious part (finding the
  exact result name + cell + unit for every result cell in the module set).

.PARAMETER EnterpriseId
  Enterprise id, matching Enterprises/Enterprise_<Id>.json (e.g. PastureBeef).

.PARAMETER RepoRoot
  Repo root. Defaults to the parent of this script's directory.

.EXAMPLE
  npm run list-enterprise-results -- Dairy
#>
param(
  [Parameter(Mandatory = $true)] [string] $EnterpriseId,
  [string] $RepoRoot
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
  $RepoRoot = Split-Path -Parent $PSScriptRoot
}
$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path
$excelDir = Join-Path $RepoRoot 'Excel'
$enterprisesDir = Join-Path $RepoRoot 'Enterprises'

# ---------------------------------------------------------------------------
# Small XML/zip helpers (read-only; same approach as find-excel-errors.ps1 -
# never open the target workbooks via COM/Excel).
# ---------------------------------------------------------------------------
$script:NsMain = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:NsRel  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:NsPkgRel = 'http://schemas.openxmlformats.org/package/2006/relationships'

function ConvertFrom-ColumnLetters {
  param([string] $Letters)
  $n = 0
  foreach ($ch in $Letters.ToUpperInvariant().ToCharArray()) { $n = $n * 26 + ([int]$ch - [int][char]'A' + 1) }
  return $n
}

function ConvertTo-ColumnLetters {
  param([int] $Number)
  $n = $Number; $letters = ''
  while ($n -gt 0) { $rem = ($n - 1) % 26; $letters = [string][char](65 + $rem) + $letters; $n = [int](($n - $rem - 1) / 26) }
  return $letters
}

function Get-AdjacentCellAddr {
  # One column to the right of a $-qualified or bare A1 ref (e.g. "$E$80" -> "F80").
  param([string] $CellRef)
  $clean = $CellRef -replace '\$', ''
  if ($clean -notmatch '^([A-Za-z]+)([0-9]+)$') { return $null }
  $col = ConvertFrom-ColumnLetters $matches[1]
  return (ConvertTo-ColumnLetters ($col + 1)) + $matches[2]
}

function Read-ZipEntryText {
  param([System.IO.Compression.ZipArchive] $Zip, [string] $EntryName)
  $entry = $Zip.GetEntry($EntryName)
  if ($null -eq $entry) { return $null }
  $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
  try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Open-WorkbookContext {
  # Opens the xlsx read-only (shared, safe even if open elsewhere in Excel) and
  # returns everything needed to resolve names + adjacent-cell unit text,
  # without going back to disk per sheet. Caller must call Close-WorkbookContext.
  param([string] $XlsxPath)

  $fs = New-Object System.IO.FileStream($XlsxPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
  $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)

  $wbText = Read-ZipEntryText -Zip $zip -EntryName 'xl/workbook.xml'
  $wbDoc = New-Object System.Xml.XmlDocument
  $wbDoc.LoadXml($wbText)
  $wbNs = New-Object System.Xml.XmlNamespaceManager($wbDoc.NameTable)
  $wbNs.AddNamespace('x', $script:NsMain); $wbNs.AddNamespace('r', $script:NsRel)

  $relsText = Read-ZipEntryText -Zip $zip -EntryName 'xl/_rels/workbook.xml.rels'
  $relsDoc = New-Object System.Xml.XmlDocument
  $relsDoc.LoadXml($relsText)
  $relsNs = New-Object System.Xml.XmlNamespaceManager($relsDoc.NameTable)
  $relsNs.AddNamespace('p', $script:NsPkgRel)

  # sheet display name -> part path (e.g. "xl/worksheets/sheet6.xml")
  $sheetParts = @{}
  foreach ($sn in @($wbDoc.SelectNodes('//x:sheets/x:sheet', $wbNs))) {
    $rid = $sn.GetAttribute('id', $script:NsRel)
    $rel = $relsDoc.SelectSingleNode("//p:Relationship[@Id='$rid']", $relsNs)
    if ($null -eq $rel) { continue }
    $target = $rel.GetAttribute('Target') -replace '^/', ''
    if ($target -notmatch '^xl/') { $target = "xl/$target" }
    $sheetParts[$sn.GetAttribute('name')] = $target
  }

  # sharedStrings.xml -> index-ordered display strings.
  $sharedStrings = New-Object System.Collections.Generic.List[string]
  $sstText = Read-ZipEntryText -Zip $zip -EntryName 'xl/sharedStrings.xml'
  if ($null -ne $sstText) {
    $sstDoc = New-Object System.Xml.XmlDocument
    $sstDoc.LoadXml($sstText)
    $sstNs = New-Object System.Xml.XmlNamespaceManager($sstDoc.NameTable)
    $sstNs.AddNamespace('x', $script:NsMain)
    foreach ($si in @($sstDoc.SelectNodes('//x:sst/x:si', $sstNs))) {
      $sb = New-Object System.Text.StringBuilder
      foreach ($t in @($si.SelectNodes('.//x:t', $sstNs))) { [void] $sb.Append($t.InnerText) }
      $sharedStrings.Add($sb.ToString())
    }
  }

  $definedNames = New-Object System.Collections.Generic.List[object]
  foreach ($n in @($wbDoc.SelectNodes('//x:definedNames/x:definedName', $wbNs))) {
    if ($n.GetAttribute('localSheetId')) { continue }   # workbook-scoped only
    $definedNames.Add([pscustomobject]@{ Name = $n.GetAttribute('name'); RefersTo = $n.InnerText })
  }

  return [pscustomobject]@{
    Zip = $zip; FileStream = $fs
    SheetParts = $sheetParts; SharedStrings = $sharedStrings; DefinedNames = $definedNames
    SheetDocCache = @{}
  }
}

function Close-WorkbookContext {
  param($Ctx)
  $Ctx.Zip.Dispose(); $Ctx.FileStream.Dispose()
}

function Get-CellUnitText {
  # Returns the display text of the cell immediately right of $CellRef on
  # $SheetName, or $null if there's no sheet part / no cell / it's not text
  # (numeric adjacent cell means there's no unit label there to read).
  param($Ctx, [string] $SheetName, [string] $CellRef)
  if (-not $Ctx.SheetParts.ContainsKey($SheetName)) { return $null }
  $addr = Get-AdjacentCellAddr $CellRef
  if ($null -eq $addr) { return $null }

  if (-not $Ctx.SheetDocCache.ContainsKey($SheetName)) {
    $text = Read-ZipEntryText -Zip $Ctx.Zip -EntryName $Ctx.SheetParts[$SheetName]
    $doc = $null
    if ($null -ne $text) { $doc = New-Object System.Xml.XmlDocument; $doc.LoadXml($text) }
    $Ctx.SheetDocCache[$SheetName] = $doc
  }
  $doc = $Ctx.SheetDocCache[$SheetName]
  if ($null -eq $doc) { return $null }
  $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
  $ns.AddNamespace('x', $script:NsMain)

  $c = $doc.SelectSingleNode("//x:sheetData/x:row/x:c[@r='$addr']", $ns)
  if ($null -eq $c) { return $null }
  $t = $c.GetAttribute('t')
  switch ($t) {
    's' {
      $v = $c.SelectSingleNode('x:v', $ns)
      if ($null -eq $v) { return $null }
      $idx = 0
      if (-not [int]::TryParse($v.InnerText, [ref] $idx)) { return $null }
      if ($idx -lt 0 -or $idx -ge $Ctx.SharedStrings.Count) { return $null }
      return $Ctx.SharedStrings[$idx]
    }
    'str' {
      $v = $c.SelectSingleNode('x:v', $ns)
      if ($null -eq $v) { return $null }
      return $v.InnerText
    }
    'inlineStr' {
      $isNode = $c.SelectSingleNode('x:is', $ns)
      if ($null -eq $isNode) { return $null }
      $sb = New-Object System.Text.StringBuilder
      foreach ($tn in @($isNode.SelectNodes('.//x:t', $ns))) { [void] $sb.Append($tn.InnerText) }
      return $sb.ToString()
    }
    default { return $null }   # numeric/empty adjacent cell - no unit label there
  }
}

function Resolve-SourceWorkbook {
  # Version-agnostic prefix resolution, mirrors build-enterprise-excel.ps1's
  # Resolve-SourceWorkbook so this script and the real build agree on which
  # workbook "PastureBeef" etc. actually means today.
  param([string] $ExcelDir, [string] $HintName)

  $exact = Join-Path $ExcelDir $HintName
  if (Test-Path -LiteralPath $exact) { return (Resolve-Path -LiteralPath $exact).Path }

  $base = [System.IO.Path]::GetFileNameWithoutExtension($HintName)
  $stem = [regex]::Replace($base, '_v\d+$', '')
  $stemRegex = '(?i)^' + [regex]::Escape($stem) + '(_.*)?_v\d+$'

  $best = Get-ChildItem -Path $ExcelDir -File |
    Where-Object {
      ($_.Extension -eq '.xlsx' -or $_.Extension -eq '.xlsm') -and
      $_.Name -notlike '~$*' -and
      $_.BaseName -notmatch '(?i)_expanded' -and
      $_.BaseName -notmatch '(?i)\.bak$' -and
      ($_.BaseName -match $stemRegex -or
       [regex]::Replace($_.BaseName, '_v\d+$', '') -eq $stem)
    } |
    Sort-Object @{ Expression = { if ($_.BaseName -match '_v(\d+)$') { [int] $matches[1] } else { -1 } } }, Name -Descending |
    Select-Object -First 1

  if ($null -ne $best) { return $best.FullName }
  throw "Could not resolve source workbook for hint '$HintName' under $ExcelDir"
}

# ---------------------------------------------------------------------------
# Load enterprise config + module registry, resolve selected modules the same
# way build-enterprise-excel.ps1 does (bare string = whole module; object with
# 'include' = subset of that module's input/calculation/constants sheets).
# ---------------------------------------------------------------------------
$configPath = Join-Path $enterprisesDir "Enterprise_$EnterpriseId.json"
if (-not (Test-Path -LiteralPath $configPath)) { throw "Enterprise config not found: $configPath" }
$config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json

$registryPath = Join-Path $enterprisesDir '_ModuleRegistry.json'
$registry = Get-Content -LiteralPath $registryPath -Raw | ConvertFrom-Json

$selectedModules = New-Object System.Collections.Generic.List[object]
foreach ($m in @($config.modules)) {
  if ($m -is [string]) {
    $selectedModules.Add([pscustomobject]@{ Id = $m; CalcSheets = $null })
  } else {
    $props = $m.PSObject.Properties.Name
    $calcSheets = $null
    if (($props -contains 'include') -and $null -ne $m.include -and
        ($m.include.PSObject.Properties.Name -contains 'calculation') -and $null -ne $m.include.calculation) {
      $calcSheets = @($m.include.calculation)
    }
    $selectedModules.Add([pscustomobject]@{ Id = [string] $m.id; CalcSheets = $calcSheets })
  }
}

# ---------------------------------------------------------------------------
# Per module: resolve workbook, read result names, group by (sheet,base),
# read the unit from the sheet cell immediately right of the result cell.
# ---------------------------------------------------------------------------
# Suffix is usually Method1/Method2 (two alternate calculation methodologies
# for the same result), but at least one module (Fuel) instead publishes one
# result PER GAS from a single calculation (_Result_CO2/_Result_CH4/
# _Result_N2O/_Result_Scope3, no Method1/Method2 concept at all) - so the
# suffix is captured generically and the two shapes are rendered differently
# below rather than assuming every result is a Method1/Method2 pair.
$resultNameRegex = [regex] '^(?<base>VEERG_.+)_Result_(?<suffix>[A-Za-z0-9]+)$'
$refersToRegex   = [regex] "^'(?<sheet>[^']+)'!(?<cell>.+)$"
$methodOnlySet = @{ 'Method1' = $true; 'Method2' = $true }

Write-Host ("Enterprise : {0}" -f $EnterpriseId)
Write-Host ("Config     : {0}" -f $configPath)
Write-Host ''

foreach ($sel in $selectedModules) {
  $mod = $registry.modules.($sel.Id)
  if ($null -eq $mod) { Write-Warning "Module '$($sel.Id)' not found in registry - skipped."; continue }

  $calcSheets = if ($null -ne $sel.CalcSheets) { $sel.CalcSheets } else { @($mod.sheets.calculation) }
  if (@($calcSheets).Count -eq 0) { continue }

  $wbPath = Resolve-SourceWorkbook -ExcelDir $excelDir -HintName ([string] $mod.sourceWorkbook)
  $ctx = Open-WorkbookContext -XlsxPath $wbPath
  try {
    # sheet -> base -> ordered{ suffix -> {Name; Cell} }, all in first-seen order.
    $bySheet = [ordered]@{}
    foreach ($n in $ctx.DefinedNames) {
      $m = $resultNameRegex.Match($n.Name)
      if (-not $m.Success) { continue }
      $rt = $refersToRegex.Match($n.RefersTo)
      if (-not $rt.Success) { continue }   # unresolved/#REF! or unexpected form - skip rather than guess
      $sheet = $rt.Groups['sheet'].Value
      if ($calcSheets -notcontains $sheet) { continue }   # not one of this module's selected calc sheets

      $base = $m.Groups['base'].Value
      $suffix = $m.Groups['suffix'].Value
      if (-not $bySheet.Contains($sheet)) { $bySheet[$sheet] = [ordered]@{} }
      if (-not $bySheet[$sheet].Contains($base)) { $bySheet[$sheet][$base] = [ordered]@{} }
      $bySheet[$sheet][$base][$suffix] = [pscustomobject]@{ Name = $n.Name; Cell = $rt.Groups['cell'].Value }
    }

    if ($bySheet.Count -eq 0) { continue }

    $moduleLabel = if ($mod.title) { $mod.title } else { $sel.Id }
    Write-Host ("=== {0}  ({1}) ===" -f $moduleLabel, (Split-Path $wbPath -Leaf))
    foreach ($sheet in $calcSheets) {
      if (-not $bySheet.Contains($sheet)) { continue }
      foreach ($base in $bySheet[$sheet].Keys) {
        $suffixes = $bySheet[$sheet][$base]
        $firstCell = ($suffixes.Values | Select-Object -First 1).Cell
        $unit = Get-CellUnitText -Ctx $ctx -SheetName $sheet -CellRef $firstCell
        if ([string]::IsNullOrWhiteSpace($unit)) { $unit = '?' }
        Write-Host ("  {0}   [{1}]" -f $sheet, $unit)

        $isMethodPair = $true
        foreach ($k in $suffixes.Keys) { if (-not $methodOnlySet.ContainsKey($k)) { $isMethodPair = $false; break } }
        if ($isMethodPair) {
          # Familiar fixed two-slot rendering, "(none)" when truly absent.
          foreach ($label in @('Method1', 'Method2')) {
            $shown = if ($label -eq 'Method1') { 'Method 1' } else { 'Method 2' }
            if ($suffixes.Contains($label)) { Write-Host ("    {0}: ={1}   ({2})" -f $shown, $suffixes[$label].Name, $suffixes[$label].Cell) }
            else { Write-Host ("    {0}: (none)" -f $shown) }
          }
        } else {
          # Per-gas (or other non-Method) shape: list whatever is present,
          # no "(none)" - there's no fixed expected set to compare against.
          foreach ($k in $suffixes.Keys) {
            Write-Host ("    {0}: ={1}   ({2})" -f $k, $suffixes[$k].Name, $suffixes[$k].Cell)
          }
        }
      }
    }
    Write-Host ''
  } finally {
    Close-WorkbookContext $ctx
  }
}
