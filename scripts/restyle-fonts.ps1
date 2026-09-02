<#
.SYNOPSIS
  One-off: restyle Excel workbook XML - rename "Times New Roman" -> "Arial", shrink
  font sizes, and normalise row heights (row 1 taller, every other row uniform).
  Bold / italic / underline / colour / sub-super script are left untouched.

.DESCRIPTION
  Edits the OOXML parts directly (zip + XML, no Excel, no COM), so it covers every
  place a font face or size is stored - not just the named cell styles:

    xl/styles.xml            <fonts> table (EVERY direct-formatted or styled cell
                             font resolves to an entry here) + <dxfs> (conditional-
                             format and table-style fonts)
    xl/sharedStrings.xml     rich-text run props  (<rPr><rFont/><sz/></rPr>)
    xl/theme/themeN.xml      <fontScheme> major/minor typefaces (the workbook's
                             default face, used by any font tagged scheme="minor")
    xl/drawings/drawingN.xml  shape / text-box runs (a:latin/a:cs/a:ea @typeface,
                             a:rPr/a:defRPr/a:endParaRPr @sz)
    xl/charts/chartN.xml     chart text runs (same DrawingML props)
    xl/worksheets/sheetN.xml  &"font,style" / &<size> header & footer codes, plus
                             inline-string runs
    xl/commentsN.xml         legacy cell-comment run props

  DrawingML sizes are stored in 1/100 pt (sz="1400" = 14 pt); the rule is applied
  to the point value and written back in hundredths.

  Size rule (identical whether the source value is nominally px or pt):
      size >= 12        ->  size - 2      (12->10, 14->12, 16->14, 20->18 ...)
      10 <= size < 12   ->  size - 1      (10->9, 11->10)
      size <  10        ->  unchanged     (9, 8, ... kept as-is)

  When a font's face is renamed to Arial, any <scheme val="minor|major"/> child is
  dropped so the literal "Arial" takes effect regardless of the theme (matters when
  -StylesOnly is used and the theme is left pointing at Times New Roman).

  Row heights (every worksheet): sheetFormatPr/@defaultRowHeight is set to
  -RowHeight (default 20) with customHeight="1"; every existing <row> below row 1
  has its explicit ht/customHeight stripped so it inherits that default; row 1 is
  forced to ht=-Row1Height (default 30, customHeight="1"), created if the sheet has
  no <row r="1"> yet. Rows that were deliberately made taller (wrapped headings,
  spacers) are flattened to the uniform height too.

  Read-only unless -Commit is passed (mirrors remove-phantom-external-links.ps1 /
  apply-names.ps1 / audit-names.ps1).

.PARAMETER RepoRoot
  Repository root. Defaults to the parent of the scripts folder.

.PARAMETER WorkbookPath
  Full path to a single .xlsx to process. If omitted, every top-level Excel/*.xlsx
  and Excel/Enterprises/*.xlsx is processed - template workbooks INCLUDED (skips
  ~$*, *_expanded*, *.bak).

.PARAMETER Commit
  Write the changes. Without it, only reports what would change.

.PARAMETER Backup
  With -Commit, copy each workbook once to Excel/Backups/Backup_PreFont/ before
  editing it (never overwrites an existing backup).

.PARAMETER StylesOnly
  Restrict to xl/styles.xml (named styles + direct cell fonts + conditional-format
  fonts). Skips theme, shared strings, drawings, charts, headers/footers, comments
  AND row heights.

.PARAMETER RowHeightsOnly
  Do only the worksheet row-height pass; skip all font work.

.PARAMETER SkipRename
  Do not touch font faces (only resize + row heights).

.PARAMETER SkipResize
  Do not touch font sizes (only rename + row heights).

.PARAMETER SkipRowHeights
  Do not touch row heights (only the font passes).

.PARAMETER Row1Height
  Height for row 1 of every sheet (default 30).

.PARAMETER RowHeight
  Height for every row below row 1, and the sheet default (default 20).

.PARAMETER Force
  Re-process a workbook that already carries the "restyled" marker. The font-size
  rule is CUMULATIVE (every run shrinks eligible sizes again), so a successful
  -Commit that resized fonts stamps a RestyleFontsApplied custom document property
  and later resizing -Commit runs skip that workbook unless -Force is given. Rename
  and row-height passes are idempotent and are never marker-blocked.

.PARAMETER ShowMax
  Max distinct sample change lines to print per workbook (default 30; 0 = none,
  -1 = all).

.EXAMPLE
  powershell -File .\scripts\restyle-fonts.ps1 -WorkbookPath .\Excel\5_Fertiliser_WIP_v10.xlsx
  powershell -File .\scripts\restyle-fonts.ps1 -WorkbookPath .\Excel\5_Fertiliser_WIP_v10.xlsx -Commit -Backup
  powershell -File .\scripts\restyle-fonts.ps1 -RowHeightsOnly -Commit
  npm run restyle-fonts
  npm run restyle-fonts:commit
#>
param(
  [string] $RepoRoot = (Split-Path $PSScriptRoot -Parent),
  [string] $WorkbookPath,
  [switch] $Commit,
  [switch] $Backup,
  [switch] $StylesOnly,
  [switch] $RowHeightsOnly,
  [switch] $SkipRename,
  [switch] $SkipResize,
  [switch] $SkipRowHeights,
  [double] $Row1Height = 30,
  [double] $RowHeight = 20,
  [switch] $Force,
  [int]    $ShowMax = 30
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

if ($StylesOnly -and $RowHeightsOnly) { throw 'Conflicting flags: -StylesOnly and -RowHeightsOnly.' }

$script:DoFonts      = -not $RowHeightsOnly
$script:DoRename     = $script:DoFonts -and -not $SkipRename
$script:DoResize     = $script:DoFonts -and -not $SkipResize
$script:DoRowHeights = (-not $SkipRowHeights) -and (-not $StylesOnly)

if (-not $script:DoRename -and -not $script:DoResize -and -not $script:DoRowHeights) {
  throw 'Nothing to do: the skip flags disable every pass.'
}

# Per-workbook effective toggles the transform functions read. ApplyResize is
# lowered for a workbook that already carries the marker (unless -Force) so a run
# still applies the idempotent rename + row-height passes without re-shrinking.
$script:ApplyRename = $script:DoRename
$script:ApplyResize = $script:DoResize

$script:NsMain = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:NsDml  = 'http://schemas.openxmlformats.org/drawingml/2006/main'
$script:Inv    = [System.Globalization.CultureInfo]::InvariantCulture

# ---------------------------------------------------------------------------
# Workbook discovery (matches remove-phantom-external-links.ps1 / audit-names.ps1).
# ---------------------------------------------------------------------------
function Get-TargetWorkbooks {
  param([string] $RepoRoot, [string] $WorkbookPath)
  if (-not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
    return @((Resolve-Path -LiteralPath $WorkbookPath).Path)
  }
  $excelDir = Join-Path $RepoRoot 'Excel'
  if (-not (Test-Path -LiteralPath $excelDir)) { throw "Excel folder not found: $excelDir" }
  $dirs = @($excelDir, (Join-Path $excelDir 'Enterprises')) | Where-Object { Test-Path -LiteralPath $_ }
  $files = foreach ($d in $dirs) {
    Get-ChildItem -LiteralPath $d -Filter '*.xlsx' -File |
      Where-Object {
        $_.Name -notlike '~$*' -and
        $_.BaseName -notmatch '(?i)_expanded' -and
        $_.Name -notmatch '(?i)\.bak'
      }
  }
  $files | Sort-Object FullName | ForEach-Object { $_.FullName }
}

# ---------------------------------------------------------------------------
# Zip / XML helpers.
# ---------------------------------------------------------------------------
function Read-ZipEntryText {
  param($Zip, [string] $EntryName)
  $entry = $Zip.GetEntry($EntryName)
  if ($null -eq $entry) { return $null }
  $reader = [System.IO.StreamReader]::new($entry.Open(), [System.Text.Encoding]::UTF8)
  try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Save-XmlDocEntry {
  param($Zip, [string] $EntryName, $Doc)
  $existing = $Zip.GetEntry($EntryName)
  if ($null -ne $existing) { $existing.Delete() }
  $entry = $Zip.CreateEntry($EntryName)
  $writer = [System.IO.StreamWriter]::new($entry.Open(), [System.Text.UTF8Encoding]::new($false))
  try { $Doc.Save($writer) } finally { $writer.Dispose() }
}

function Write-ZipEntryText {
  param($Zip, [string] $EntryName, [string] $Text)
  $existing = $Zip.GetEntry($EntryName)
  if ($null -ne $existing) { $existing.Delete() }
  $entry = $Zip.CreateEntry($EntryName)
  $writer = [System.IO.StreamWriter]::new($entry.Open(), [System.Text.UTF8Encoding]::new($false))
  try { $writer.Write($Text) } finally { $writer.Dispose() }
}

function New-XmlDoc {
  param([string] $Text)
  $doc = New-Object System.Xml.XmlDocument
  $doc.PreserveWhitespace = $true
  $doc.LoadXml($Text)
  return , $doc   # comma stops PowerShell enumerating the doc's child nodes
}

function New-NsManager {
  param($Doc, [hashtable] $Namespaces)
  $ns = New-Object System.Xml.XmlNamespaceManager($Doc.NameTable)
  foreach ($k in $Namespaces.Keys) { $ns.AddNamespace($k, $Namespaces[$k]) }
  return , $ns
}

# ---------------------------------------------------------------------------
# Size rule.
# ---------------------------------------------------------------------------
function Get-NewPointSize {
  param([double] $Pt)
  if ($Pt -ge 12) { return $Pt - 2 }
  if ($Pt -ge 10) { return $Pt - 1 }
  return $Pt
}

function Format-PointSize {
  # Preserve integer-vs-decimal formatting of the original literal.
  param([string] $Original, [double] $New)
  if ($Original.Contains('.')) { return $New.ToString($script:Inv) }
  return ([int][math]::Round($New)).ToString($script:Inv)
}

function Convert-ToDoubleOrNull {
  param([string] $Text)
  $d = 0.0
  if ([double]::TryParse($Text, [System.Globalization.NumberStyles]::Float, $script:Inv, [ref] $d)) { return $d }
  return $null
}

# ---------------------------------------------------------------------------
# Change log (reset per workbook).
# ---------------------------------------------------------------------------
$script:Recs = $null
function Add-Rec {
  param([string] $Part, [string] $Kind, [string] $Old, [string] $New)
  [void] $script:Recs.Add([pscustomobject]@{ Part = $Part; Kind = $Kind; Old = $Old; New = $New })
}

function Rename-FaceValue {
  # Returns 'Arial' if $Value is Times New Roman (case / whitespace tolerant), else $null.
  param([string] $Value)
  if ($null -eq $Value) { return $null }
  if ($Value.Trim() -ieq 'Times New Roman') { return 'Arial' }
  return $null
}

# ---------------------------------------------------------------------------
# Transform: SpreadsheetML <font> nodes  (xl/styles.xml <fonts> AND <dxfs>).
# ---------------------------------------------------------------------------
function Update-SmlFontNodes {
  param($Doc, $Ns, [string] $Part)
  $count = 0
  foreach ($font in @($Doc.SelectNodes('//x:font', $Ns))) {
    if ($script:ApplyRename) {
      $nameEl = $font.SelectSingleNode('x:name', $Ns)
      if ($null -eq $nameEl) { $nameEl = $font.SelectSingleNode('x:rFont', $Ns) }
      if ($null -ne $nameEl) {
        $old = [string] $nameEl.GetAttribute('val')
        $new = Rename-FaceValue $old
        if ($null -ne $new) {
          $nameEl.SetAttribute('val', $new)
          $sch = $font.SelectSingleNode('x:scheme', $Ns)
          if ($null -ne $sch) { [void] $font.RemoveChild($sch) }
          Add-Rec $Part 'face' $old $new
          $count++
        }
      }
    }
    if ($script:ApplyResize) {
      $szEl = $font.SelectSingleNode('x:sz', $Ns)
      if ($null -ne $szEl) {
        $old = [string] $szEl.GetAttribute('val')
        $pt = Convert-ToDoubleOrNull $old
        if ($null -ne $pt) {
          $np = Get-NewPointSize $pt
          if ($np -ne $pt) {
            $nv = Format-PointSize $old $np
            $szEl.SetAttribute('val', $nv)
            Add-Rec $Part 'size' ("{0}pt" -f $old) ("{0}pt" -f $nv)
            $count++
          }
        }
      }
    }
  }
  return $count
}

# ---------------------------------------------------------------------------
# Transform: SpreadsheetML rich-run props  (<rPr><rFont/><sz/></rPr>) - used by
# sharedStrings, legacy comments, and worksheet inline strings.
# ---------------------------------------------------------------------------
function Update-SmlRunProps {
  param($Doc, $Ns, [string] $Part)
  $count = 0
  foreach ($rpr in @($Doc.SelectNodes('//x:rPr', $Ns))) {
    if ($script:ApplyRename) {
      $rf = $rpr.SelectSingleNode('x:rFont', $Ns)
      if ($null -ne $rf) {
        $old = [string] $rf.GetAttribute('val')
        $new = Rename-FaceValue $old
        if ($null -ne $new) {
          $rf.SetAttribute('val', $new)
          $sch = $rpr.SelectSingleNode('x:scheme', $Ns)
          if ($null -ne $sch) { [void] $rpr.RemoveChild($sch) }
          Add-Rec $Part 'face' $old $new
          $count++
        }
      }
    }
    if ($script:ApplyResize) {
      $sz = $rpr.SelectSingleNode('x:sz', $Ns)
      if ($null -ne $sz) {
        $old = [string] $sz.GetAttribute('val')
        $pt = Convert-ToDoubleOrNull $old
        if ($null -ne $pt) {
          $np = Get-NewPointSize $pt
          if ($np -ne $pt) {
            $nv = Format-PointSize $old $np
            $sz.SetAttribute('val', $nv)
            Add-Rec $Part 'size' ("{0}pt" -f $old) ("{0}pt" -f $nv)
            $count++
          }
        }
      }
    }
  }
  return $count
}

# ---------------------------------------------------------------------------
# Transform: DrawingML run props  (drawings + charts). @typeface faces and @sz
# sizes (hundredths of a point).
# ---------------------------------------------------------------------------
function Update-DrawingmlProps {
  param($Doc, $Ns, [string] $Part)
  $count = 0
  if ($script:ApplyRename) {
    foreach ($tf in @($Doc.SelectNodes('//a:latin[@typeface] | //a:cs[@typeface] | //a:ea[@typeface]', $Ns))) {
      $old = [string] $tf.GetAttribute('typeface')
      $new = Rename-FaceValue $old
      if ($null -ne $new) { $tf.SetAttribute('typeface', $new); Add-Rec $Part 'face' $old $new; $count++ }
    }
  }
  if ($script:ApplyResize) {
    foreach ($rp in @($Doc.SelectNodes('//a:rPr[@sz] | //a:defRPr[@sz] | //a:endParaRPr[@sz]', $Ns))) {
      $old = [string] $rp.GetAttribute('sz')
      $h = 0
      if ([int]::TryParse($old, [System.Globalization.NumberStyles]::Integer, $script:Inv, [ref] $h)) {
        $pt = $h / 100.0
        $np = Get-NewPointSize $pt
        if ($np -ne $pt) {
          $nh = [int][math]::Round($np * 100)
          $rp.SetAttribute('sz', $nh.ToString($script:Inv))
          Add-Rec $Part 'size' ("{0}pt" -f $pt.ToString($script:Inv)) ("{0}pt" -f $np.ToString($script:Inv))
          $count++
        }
      }
    }
  }
  return $count
}

# ---------------------------------------------------------------------------
# Transform: theme <fontScheme> typefaces (workbook default face).
# ---------------------------------------------------------------------------
function Update-ThemeFonts {
  param($Doc, $Ns, [string] $Part)
  if (-not $script:ApplyRename) { return 0 }
  $count = 0
  foreach ($tf in @($Doc.SelectNodes('//a:fontScheme//*[@typeface]', $Ns))) {
    $old = [string] $tf.GetAttribute('typeface')
    $new = Rename-FaceValue $old
    if ($null -ne $new) { $tf.SetAttribute('typeface', $new); Add-Rec $Part 'face' $old $new; $count++ }
  }
  return $count
}

# ---------------------------------------------------------------------------
# Transform: worksheet header / footer codes -  &"Font,Style"  and  &<size>.
# ---------------------------------------------------------------------------
function Update-HeaderFooter {
  param($Doc, $Ns, [string] $Part)
  $hf = $Doc.SelectSingleNode('//x:headerFooter', $Ns)
  if ($null -eq $hf) { return 0 }
  $count = 0
  $els = @($hf.SelectNodes('x:oddHeader | x:evenHeader | x:firstHeader | x:oddFooter | x:evenFooter | x:firstFooter', $Ns))
  foreach ($el in $els) {
    $t = [string] $el.InnerText
    if ([string]::IsNullOrEmpty($t)) { continue }
    $orig = $t

    if ($script:ApplyRename) {
      # &"Font Name,Style"  (not preceded by another & - that is a literal ampersand)
      $t = [regex]::Replace($t, '(?<!&)&"([^"]*)"', {
        param($m)
        $spec = $m.Groups[1].Value
        $parts = $spec.Split(',')
        if ($parts.Length -ge 1 -and (Rename-FaceValue $parts[0])) {
          $parts[0] = 'Arial'
          return '&"' + ($parts -join ',') + '"'
        }
        return $m.Value
      })
    }
    if ($script:ApplyResize) {
      # &<digits>  size code  (not &&<digits>, which is a literal '&' then text)
      $t = [regex]::Replace($t, '(?<!&)&(\d+)', {
        param($m)
        $pt = [double]::Parse($m.Groups[1].Value, $script:Inv)
        $np = Get-NewPointSize $pt
        if ($np -ne $pt) { return '&' + ([int][math]::Round($np)).ToString($script:Inv) }
        return $m.Value
      })
    }

    if ($t -ne $orig) { $el.InnerText = $t; Add-Rec $Part 'header/footer' $orig $t; $count++ }
  }
  return $count
}

# ---------------------------------------------------------------------------
# Transform: worksheet row heights. Row 1 -> $Row1Height, every other row ->
# $RowHeight. An explicit ht + customHeight="1" is forced on every <row> element
# (Excel auto-fits rows that only inherit sheetFormatPr/@defaultRowHeight), and
# defaultRowHeight is set to $RowHeight to cover rows that have no <row> element.
# ---------------------------------------------------------------------------
function Update-RowHeights {
  param($Doc, $Ns, [string] $Part, [double] $Row1, [double] $Rest)

  $wsEl = $Doc.SelectSingleNode('/x:worksheet', $Ns)
  if ($null -eq $wsEl) { return 0 }
  $count = 0
  $restStr = $Rest.ToString($script:Inv)
  $row1Str = $Row1.ToString($script:Inv)

  # 1. sheetFormatPr/@defaultRowHeight  (create the element if absent; schema
  #    position is after sheetViews, before cols / sheetData).
  $sfp = $Doc.SelectSingleNode('/x:worksheet/x:sheetFormatPr', $Ns)
  if ($null -eq $sfp) {
    $sfp = $Doc.CreateElement('sheetFormatPr', $script:NsMain)
    $anchor = $Doc.SelectSingleNode('/x:worksheet/x:cols', $Ns)
    if ($null -eq $anchor) { $anchor = $Doc.SelectSingleNode('/x:worksheet/x:sheetData', $Ns) }
    if ($null -ne $anchor) { [void] $wsEl.InsertBefore($sfp, $anchor) } else { [void] $wsEl.AppendChild($sfp) }
  }
  if ($sfp.GetAttribute('defaultRowHeight') -ne $restStr -or $sfp.GetAttribute('customHeight') -ne '1') {
    $wasDef = [string] $sfp.GetAttribute('defaultRowHeight')
    $sfp.SetAttribute('defaultRowHeight', $restStr)
    $sfp.SetAttribute('customHeight', '1')
    Add-Rec $Part 'rowheight' ("default=" + $(if ($wasDef) { $wasDef } else { '(unset)' })) ("default=" + $restStr)
    $count++
  }

  # 2. Force an explicit height on every <row>.
  $sheetData = $Doc.SelectSingleNode('/x:worksheet/x:sheetData', $Ns)
  if ($null -ne $sheetData) {
    $haveRow1 = $false
    $set1 = 0
    $setRest = 0
    foreach ($row in @($sheetData.SelectNodes('x:row', $Ns))) {
      $isRow1 = ([string] $row.GetAttribute('r') -eq '1')
      if ($isRow1) { $haveRow1 = $true }
      $want = if ($isRow1) { $row1Str } else { $restStr }
      if ($row.GetAttribute('ht') -ne $want -or $row.GetAttribute('customHeight') -ne '1') {
        $row.SetAttribute('ht', $want)
        $row.SetAttribute('customHeight', '1')
        $count++
        if ($isRow1) { $set1++ } else { $setRest++ }
      }
    }
    if ($set1 -gt 0)    { Add-Rec $Part 'rowheight' 'row 1' ("ht=" + $row1Str) }
    if ($setRest -gt 0) { Add-Rec $Part 'rowheight' ("{0} row(s)" -f $setRest) ("ht=" + $restStr) }
    if (-not $haveRow1) {
      $r1 = $Doc.CreateElement('row', $script:NsMain)
      $r1.SetAttribute('r', '1')
      $r1.SetAttribute('ht', $row1Str)
      $r1.SetAttribute('customHeight', '1')
      $firstRow = $sheetData.SelectSingleNode('x:row', $Ns)
      if ($null -ne $firstRow) { [void] $sheetData.InsertBefore($r1, $firstRow) } else { [void] $sheetData.AppendChild($r1) }
      Add-Rec $Part 'rowheight' 'row 1 (added)' ("ht=" + $row1Str)
      $count++
    }
  }
  return $count
}

# ---------------------------------------------------------------------------
# "Already restyled" marker - a custom document property. The size rule is
# cumulative, so -Commit stamps this and refuses to re-shrink on a later run
# unless -Force is passed. XML-only; creates docProps/custom.xml (+ its content-
# type Override and package relationship) if the workbook has none.
# ---------------------------------------------------------------------------
$script:MarkerName     = 'RestyleFontsApplied'
$script:NsCustomProps  = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
$script:NsVt           = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'

function Get-RestyleMarker {
  param($Zip)
  $text = Read-ZipEntryText -Zip $Zip -EntryName 'docProps/custom.xml'
  if ($null -eq $text) { return $null }
  $m = [regex]::Match($text, ('<property[^>]*\bname="' + [regex]::Escape($script:MarkerName) + '"[^>]*>\s*<vt:lpwstr>(?<v>[^<]*)</vt:lpwstr>'))
  if ($m.Success) { return $m.Groups['v'].Value }
  return $null
}

function Set-RestyleMarker {
  param($Zip, [string] $Value)

  $text = Read-ZipEntryText -Zip $Zip -EntryName 'docProps/custom.xml'
  if ($null -eq $text) {
    $text = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            ('<Properties xmlns="{0}" xmlns:vt="{1}"></Properties>' -f $script:NsCustomProps, $script:NsVt)
  }
  $doc = New-XmlDoc -Text $text
  $ns = New-NsManager -Doc $doc -Namespaces @{ p = $script:NsCustomProps; vt = $script:NsVt }
  $root = $doc.SelectSingleNode('/p:Properties', $ns)

  $maxPid = 1
  $target = $null
  foreach ($p in @($root.SelectNodes('p:property', $ns))) {
    $pidVal = 0
    if ([int]::TryParse([string] $p.GetAttribute('pid'), [ref] $pidVal) -and $pidVal -gt $maxPid) { $maxPid = $pidVal }
    if ([string] $p.GetAttribute('name') -eq $script:MarkerName) { $target = $p }
  }
  if ($null -eq $target) {
    $target = $doc.CreateElement('property', $script:NsCustomProps)
    $target.SetAttribute('fmtid', '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}')
    $target.SetAttribute('pid', ($maxPid + 1).ToString($script:Inv))
    $target.SetAttribute('name', $script:MarkerName)
    [void] $root.AppendChild($target)
  }
  while ($target.HasChildNodes) { [void] $target.RemoveChild($target.FirstChild) }
  $vt = $doc.CreateElement('vt', 'lpwstr', $script:NsVt)
  $vt.InnerText = $Value
  [void] $target.AppendChild($vt)
  Save-XmlDocEntry -Zip $Zip -EntryName 'docProps/custom.xml' -Doc $doc

  $ctText = Read-ZipEntryText -Zip $Zip -EntryName '[Content_Types].xml'
  if ($null -ne $ctText -and $ctText -notmatch '/docProps/custom\.xml') {
    $ov = '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>'
    Write-ZipEntryText -Zip $Zip -EntryName '[Content_Types].xml' -Text ($ctText -replace '</Types>', ($ov + '</Types>'))
  }

  $relsText = Read-ZipEntryText -Zip $Zip -EntryName '_rels/.rels'
  if ($null -ne $relsText -and $relsText -notmatch 'docProps/custom\.xml') {
    $nums = @([regex]::Matches($relsText, 'Id="rId(\d+)"') | ForEach-Object { [int] $_.Groups[1].Value })
    $nextId = if ($nums.Count -gt 0) { ($nums | Measure-Object -Maximum).Maximum + 1 } else { 1 }
    $rel = ('<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>' -f $nextId)
    Write-ZipEntryText -Zip $Zip -EntryName '_rels/.rels' -Text ($relsText -replace '</Relationships>', ($rel + '</Relationships>'))
  }
}

# ---------------------------------------------------------------------------
# Process one workbook. No writes unless $Commit.
# ---------------------------------------------------------------------------
function Invoke-Workbook {
  param([string] $Path)

  $script:Recs = New-Object System.Collections.Generic.List[object]

  # --- Analysis pass: always READ-only. The transformed DOMs are detached
  #     in-memory copies, so they stay valid after this archive is disposed.
  #     A Update-mode archive rewrites the whole .xlsx on Dispose even when
  #     nothing changed, so only the write pass below opens Update mode.
  $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Read)
  try {
    $marker = Get-RestyleMarker -Zip $zip

    # A workbook that already carries the marker has had its (cumulative) font
    # resize done once - suppress just that pass here unless -Force, but still
    # run the idempotent rename + row-height passes.
    $resizeBlocked = $script:DoResize -and $null -ne $marker -and -not $Force
    $script:ApplyResize = $script:DoResize -and -not $resizeBlocked
    # (entryName, transformKind)
    $plan = New-Object System.Collections.Generic.List[object]
    $fontParts = $script:DoFonts -and -not $StylesOnly
    if ($script:DoFonts -and $null -ne $zip.GetEntry('xl/styles.xml')) { $plan.Add(@('xl/styles.xml', 'sml-fonts')) }
    if ($fontParts -and $null -ne $zip.GetEntry('xl/sharedStrings.xml')) { $plan.Add(@('xl/sharedStrings.xml', 'sml-runs')) }
    foreach ($e in $zip.Entries) {
      $fn = $e.FullName
      if ($fontParts) {
        if     ($fn -match '^xl/theme/theme\d+\.xml$')      { $plan.Add(@($fn, 'theme')); continue }
        elseif ($fn -match '^xl/drawings/drawing\d+\.xml$') { $plan.Add(@($fn, 'dml'));   continue }
        elseif ($fn -match '^xl/charts/chart\d+\.xml$')     { $plan.Add(@($fn, 'dml'));   continue }
        elseif ($fn -match '^xl/comments\d+\.xml$')         { $plan.Add(@($fn, 'sml-runs')); continue }
      }
      if ($fn -match '^xl/worksheets/sheet\d+\.xml$' -and ($fontParts -or $script:DoRowHeights)) {
        $plan.Add(@($fn, 'sheet'))
      }
    }

    $changedDocs = @{}
    foreach ($item in $plan) {
      $entryName = [string] $item[0]
      $kind      = [string] $item[1]
      $text = Read-ZipEntryText -Zip $zip -EntryName $entryName
      if ($null -eq $text) { continue }
      $doc = New-XmlDoc -Text $text
      $c = 0
      switch ($kind) {
        'sml-fonts' { $ns = New-NsManager -Doc $doc -Namespaces @{ x = $script:NsMain }; $c = Update-SmlFontNodes -Doc $doc -Ns $ns -Part $entryName }
        'sml-runs'  { $ns = New-NsManager -Doc $doc -Namespaces @{ x = $script:NsMain }; $c = Update-SmlRunProps  -Doc $doc -Ns $ns -Part $entryName }
        'theme'     { $ns = New-NsManager -Doc $doc -Namespaces @{ a = $script:NsDml  }; $c = Update-ThemeFonts    -Doc $doc -Ns $ns -Part $entryName }
        'dml'       { $ns = New-NsManager -Doc $doc -Namespaces @{ a = $script:NsDml  }; $c = Update-DrawingmlProps -Doc $doc -Ns $ns -Part $entryName }
        'sheet' {
          $ns = New-NsManager -Doc $doc -Namespaces @{ x = $script:NsMain }
          $c = 0
          if ($script:DoFonts -and -not $StylesOnly) {
            $c += Update-SmlRunProps  -Doc $doc -Ns $ns -Part $entryName    # inline-string runs
            $c += Update-HeaderFooter -Doc $doc -Ns $ns -Part $entryName    # &"font" / &size codes
          }
          if ($script:DoRowHeights) {
            $c += Update-RowHeights -Doc $doc -Ns $ns -Part $entryName -Row1 $Row1Height -Rest $RowHeight
          }
        }
      }
      if ($c -gt 0) { $changedDocs[$entryName] = $doc }
    }

  } finally {
    $zip.Dispose()
  }

  # --- Write pass: only when there is something to write and we are committing.
  $wroteFiles = $false
  if ($Commit -and $changedDocs.Count -gt 0) {
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
      foreach ($k in @($changedDocs.Keys)) { Save-XmlDocEntry -Zip $zip -EntryName $k -Doc $changedDocs[$k] }
      # Stamp the marker only when this run actually performed the cumulative
      # resize pass (so a rename-only / row-height-only run never blocks a later
      # full run).
      if ($script:ApplyResize) {
        $stamp = 'applied {0}; rename={1} resize={2} rowheights={3} scope={4}' -f `
          (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss'),
          $script:DoRename, $script:ApplyResize, $script:DoRowHeights,
          $(if ($StylesOnly) { 'styles' } else { 'all' })
        Set-RestyleMarker -Zip $zip -Value $stamp
      }
    } finally {
      $zip.Dispose()
    }
    $wroteFiles = $true
  }

  # Return a materialised array (not the live List) - the array subexpression
  # operator @() chokes on a re-surfaced generic List in Windows PowerShell 5.1.
  return [pscustomobject]@{
    Changes       = $script:Recs.ToArray()
    PartCount     = $changedDocs.Count
    Marker        = $marker
    Wrote         = $wroteFiles
    ResizeBlocked = $resizeBlocked
  }
}

# ===========================================================================
# Main.
# ===========================================================================
$workbooks = Get-TargetWorkbooks -RepoRoot $RepoRoot -WorkbookPath $WorkbookPath
if (@($workbooks).Count -eq 0) { Write-Host 'No workbooks found to process.'; return }

$modeLabel = if ($Commit) { 'COMMIT' } else { 'DRY RUN' }
$ruleBits = @()
if ($script:DoRename)     { $ruleBits += '"Times New Roman" -> "Arial"' }
if ($script:DoResize)     { $ruleBits += '>=12pt: -2   10-11pt: -1   <=9pt: keep' }
if ($script:DoRowHeights) { $ruleBits += ('row 1 = {0}, other rows = {1}' -f $Row1Height, $RowHeight) }
$scopeLabel = if ($StylesOnly) { 'xl/styles.xml only' } elseif ($RowHeightsOnly) { 'worksheet row heights only' } else { 'all parts' }

Write-Host ("restyle-fonts [{0}]  -  {1} workbook(s)  -  {2}" -f $modeLabel, @($workbooks).Count, $scopeLabel)
Write-Host ("Rules: {0}" -f ($ruleBits -join '    |    '))
Write-Host ''

$grandChanges = 0
$grandWorkbooks = 0
$grandBlocked = 0

foreach ($path in $workbooks) {
  $leaf = Split-Path $path -Leaf

  if ($Commit -and $Backup) {
    $bdir = Join-Path $RepoRoot 'Excel\Backups\Backup_PreFont'
    if (-not (Test-Path -LiteralPath $bdir)) { New-Item -ItemType Directory -Path $bdir -Force | Out-Null }
    $bpath = Join-Path $bdir $leaf
    if (-not (Test-Path -LiteralPath $bpath)) { Copy-Item -LiteralPath $path -Destination $bpath }
  }

  try {
    $r = Invoke-Workbook -Path $path
  } catch {
    Write-Host ''
    Write-Warning ("{0}: could not process - {1}" -f $leaf, $_.Exception.Message)
    continue
  }

  if ($r.ResizeBlocked) { $grandBlocked++ }

  $changeList = $r.Changes
  $n = $changeList.Count
  if ($n -eq 0) {
    $note = if ($r.ResizeBlocked) { ' (already restyled; font resize suppressed - nothing else to do)' } else { '' }
    Write-Host ("{0}: no changes{1}" -f $leaf, $note)
    continue
  }

  $grandChanges += $n
  $grandWorkbooks++

  $suffix = if ($r.Wrote) { 'written' } elseif ($Commit) { 'commit skipped' } else { 'dry run' }
  $markNote = if ($r.ResizeBlocked) { '  (marker present: font resize suppressed; pass -Force to re-shrink)' } else { '' }
  Write-Host ("{0}: {1} change(s) across {2} part(s)  [{3}]{4}" -f $leaf, $n, $r.PartCount, $suffix, $markNote) -ForegroundColor Cyan

  foreach ($g in ($changeList | Group-Object Part | Sort-Object Name)) {
    $kinds = ($g.Group | Group-Object Kind | Sort-Object Name | ForEach-Object { "{0} {1}" -f $_.Count, $_.Name }) -join ', '
    Write-Host ("    {0,-26} {1}" -f (Split-Path $g.Name -Leaf), $kinds)
  }

  if ($ShowMax -ne 0) {
    # Collapse to distinct transforms with an occurrence count.
    $distinct = $changeList | ForEach-Object {
      if ($_.Kind -eq 'header/footer') { "[{0}] header/footer code rewritten" -f (Split-Path $_.Part -Leaf) }
      else { "[{0}] {1}: {2} -> {3}" -f (Split-Path $_.Part -Leaf), $_.Kind, $_.Old, $_.New }
    } | Group-Object | Sort-Object Name
    $show = if ($ShowMax -lt 0) { $distinct } else { $distinct | Select-Object -First $ShowMax }
    foreach ($d in $show) {
      Write-Host ("      {0,6} x  {1}" -f $d.Count, $d.Name) -ForegroundColor DarkGray
    }
    if ($ShowMax -ge 0 -and @($distinct).Count -gt $ShowMax) {
      Write-Host ("      ... {0} more distinct transform(s)" -f (@($distinct).Count - $ShowMax)) -ForegroundColor DarkGray
    }
  }
  Write-Host ''
}

Write-Host ''
if ($grandWorkbooks -eq 0) {
  Write-Host ("CLEAN: nothing to restyle in {0} workbook(s)." -f @($workbooks).Count) -ForegroundColor Green
} else {
  Write-Host ("TOTAL: {0} change(s) across {1} of {2} workbook(s)." -f $grandChanges, $grandWorkbooks, @($workbooks).Count)
}
if ($grandBlocked -gt 0) {
  Write-Host ("{0} workbook(s) already carry the marker - font resize was suppressed there (rename + row heights still applied); pass -Force to re-shrink." -f $grandBlocked) -ForegroundColor Yellow
}
if (-not $Commit) {
  Write-Host 'Dry run only - add -Commit (and -Backup) to write the changes.' -ForegroundColor Yellow
}
if ($StylesOnly -and $script:DoRename) {
  Write-Host 'Note: -StylesOnly leaves xl/theme untouched; renamed fonts had their <scheme> dropped so they still render Arial.' -ForegroundColor DarkYellow
}
if ($script:DoRowHeights) {
  Write-Host ('Note: every row below row 1 is forced to {0} - rows deliberately made taller (wrapped headings, spacers) are flattened too.' -f $RowHeight) -ForegroundColor DarkYellow
}
