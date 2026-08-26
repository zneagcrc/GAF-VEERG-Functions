<#
.SYNOPSIS
  Repoints defined names whose 'Refers to' was externalised by an Excel
  copy-between-workbooks operation back to the LOCAL sheet, when that sheet
  already exists in the target workbook - without deleting the names - then
  strips the now-orphaned external-link artifacts.

.DESCRIPTION
  Copying a sheet (or pasting cells) from one workbook into another makes
  Excel externalise any reference the copied content carries to a sheet that
  wasn't copied along, e.g.:
    'C:\...\Source.xlsx'!$E$192          (plain cell/range ref)
    '[1]Sheet Name'!$E$192               (already-registered external index)
    '[SheetNameThatLooksLikeAFile]junk'! (Excel's own mangled form when the
                                           visible RefersTo text was hand-
                                           edited in Name Manager against a
                                           still-open external link - the
                                           period-in-sheet-name symptom)
  Hand-editing the Name Manager 'Refers to' box cannot fix this: Excel keeps
  re-validating against the workbook's registered xl/externalLinks/* parts no
  matter what text is typed, so it re-derives a broken reference from its own
  link table. This script rewrites the RefersTo text directly in xl/workbook.xml
  (never opens the file in Excel/COM, so nothing re-parses it) whenever the
  externalised sheet name can be matched - by any of the forms above - to a
  sheet that is ACTUALLY PRESENT in this workbook, then calls
  Remove-ExternalLinkArtifacts (external-links.ps1) to delete the now-unused
  xl/externalLinks/* parts, <externalReferences>, rels and content-type
  overrides so Excel has nothing left to fall back on.

  Names that still can't be resolved to a local sheet (genuinely external, or
  the target sheet really isn't in this workbook) are left untouched and
  reported - Remove-ExternalLinkArtifacts only removes names it finds still
  carrying a live '[N]' (N>=1) reference after this pass.

.PARAMETER WorkbookPath
  The .xlsx to repair in place. Must be closed (not open in Excel) - the
  script fails fast with a clear message if it's locked, rather than
  corrupting a half-written save.

.PARAMETER DryRun
  Report what would change without writing anything.

.EXAMPLE
  npm run repoint-local-names -- ".\Excel\Enterprises\Enterprise_Dairy_Template_WIP_v01.xlsx"
.EXAMPLE
  npm run repoint-local-names:dry -- ".\Excel\Enterprises\Enterprise_Dairy_Template_WIP_v01.xlsx"
#>
param(
  [Parameter(Mandatory = $true)] [string] $WorkbookPath,
  [switch] $DryRun
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

. (Join-Path $PSScriptRoot 'file-access.ps1')
. (Join-Path $PSScriptRoot 'external-links.ps1')

if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
$WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
Assert-FilesAccessible -RequiredReadPaths @($WorkbookPath) -WritePaths @(if (-not $DryRun) { $WorkbookPath })

$mainNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

# ---------------------------------------------------------------------------
# Read the local sheet name set (case-insensitive) and every workbook-scoped
# definedName's current RefersTo.
# ---------------------------------------------------------------------------
$zipMode = if ($DryRun) { [System.IO.Compression.ZipArchiveMode]::Read } else { [System.IO.Compression.ZipArchiveMode]::Update }
$zip = [System.IO.Compression.ZipFile]::Open($WorkbookPath, $zipMode)
try {
  $wbEntry = $zip.GetEntry('xl/workbook.xml')
  if ($null -eq $wbEntry) { throw "xl/workbook.xml missing from $WorkbookPath - not a valid xlsx?" }
  $reader = [System.IO.StreamReader]::new($wbEntry.Open(), [System.Text.Encoding]::UTF8)
  try { $wbText = $reader.ReadToEnd() } finally { $reader.Dispose() }

  $doc = New-Object System.Xml.XmlDocument
  $doc.PreserveWhitespace = $true
  $doc.LoadXml($wbText)
  $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
  $ns.AddNamespace('x', $mainNs)

  $localSheets = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($sn in @($doc.SelectNodes('//x:sheets/x:sheet', $ns))) { [void] $localSheets.Add($sn.GetAttribute('name')) }

  # Matches a quoted qualifier that carries a bracketed token, in ANY of:
  #   '<path>[<file-or-index>]<sheet-or-junk>'!
  #   '[<file-or-index>]<sheet-or-junk>'!
  # <pre> is any leading path text (or empty), <mid> is the bracket content
  # (a real filename, a numeric external-link index, OR - Excel's mangled form
  # - a sheet name that leaked into the bracket slot), <post> is whatever
  # follows the ']' up to the closing quote (a real sheet name, OR - in the
  # mangled form - truncated junk).
  $reBracketed = [regex] "'(?<pre>[^'\[\]]*)\[(?<mid>[^\[\]]*)\](?<post>[^']*)'!"

  # A quoted path/table qualifier with no brackets at all (the form seen for
  # structured-table refs once the source workbook is closed):
  #   'C:\...\Source.xlsx'!Table_X[Col]   or   'C:\...\Source.xlsx'!$A$1
  $rePathOnly = [regex] "'(?<path>[^']*\.xls[a-z]*)'!"

  $repointed = New-Object System.Collections.Generic.List[string]
  $stillExternal = New-Object System.Collections.Generic.List[string]

  $evalBracketed = [System.Text.RegularExpressions.MatchEvaluator] {
    param($m)
    $post = $m.Groups['post'].Value
    $mid  = $m.Groups['mid'].Value
    if ($localSheets.Contains($post)) { return "'" + $post + "'!" }   # standard form: sheet is after ']'
    if ($localSheets.Contains($mid))  { return "'" + $mid  + "'!" }   # mangled form: sheet leaked into the brackets
    return $m.Value   # can't resolve - leave as-is, Remove-ExternalLinkArtifacts will report/clean it later
  }
  $evalPathOnly = [System.Text.RegularExpressions.MatchEvaluator] {
    param($m)
    return $m.Value   # no sheet name available in this form to match against - never resolvable here
  }

  $definedNamesNode = $doc.SelectSingleNode('//x:definedNames', $ns)
  $changed = $false
  if ($null -ne $definedNamesNode) {
    foreach ($n in @($definedNamesNode.SelectNodes('x:definedName', $ns))) {
      $before = $n.InnerText
      if ([string]::IsNullOrEmpty($before) -or $before.IndexOf('[') -lt 0) { continue }   # nothing bracketed to fix
      $after = $reBracketed.Replace($before, $evalBracketed)
      [void] $rePathOnly.Replace($after, $evalPathOnly)   # no-op today; keeps the path-only form visible in $stillExternal below
      if ($after -ne $before) {
        $n.InnerText = $after
        $changed = $true
        $repointed.Add(("{0}:`n    {1}`n -> {2}" -f $n.GetAttribute('name'), $before, $after))
      } elseif ($after.IndexOf('[') -ge 0 -or $rePathOnly.IsMatch($after)) {
        $stillExternal.Add(("{0}: {1}" -f $n.GetAttribute('name'), $after))
      }
    }
  }

  if ($repointed.Count -gt 0) {
    Write-Host ("Repointed {0} name(s) to a local sheet:" -f $repointed.Count) -ForegroundColor Green
    foreach ($r in $repointed) { Write-Host ("  {0}" -f $r) }
  } else {
    Write-Host 'No externalised names could be repointed (none found, or none matched a local sheet).'
  }
  if ($stillExternal.Count -gt 0) {
    Write-Host ("{0} name(s) remain external (no matching local sheet found) - leaving as-is:" -f $stillExternal.Count) -ForegroundColor Yellow
    foreach ($s in $stillExternal) { Write-Host ("  {0}" -f $s) }
  }

  if (-not $DryRun -and $changed) {
    $wbEntry.Delete()
    $newWb = $zip.CreateEntry('xl/workbook.xml')
    $writer = [System.IO.StreamWriter]::new($newWb.Open(), [System.Text.UTF8Encoding]::new($false))
    try { $doc.Save($writer) } finally { $writer.Dispose() }
  }
} finally {
  $zip.Dispose()
}

if ($DryRun) {
  Write-Host ''
  Write-Host 'Dry run - no changes written, external-link artifacts not touched.'
  return
}

# ---------------------------------------------------------------------------
# Purge whatever external-link plumbing is now orphaned (names this pass
# couldn't repoint are removed here too, per Remove-ExternalLinkArtifacts'
# own [N]>=1 rule - see external-links.ps1).
# ---------------------------------------------------------------------------
$linkResult = Remove-ExternalLinkArtifacts -TargetPath $WorkbookPath
Write-Host ''
Write-Host ("Removed {0} still-external defined name(s), {1} <externalReference> entr(y/ies), {2} externalLinks part(s)." -f `
  $linkResult.NamesRemoved, $linkResult.ReferencesRemoved, $linkResult.PartsRemoved)
