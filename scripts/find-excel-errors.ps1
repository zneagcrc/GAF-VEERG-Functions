<#
.SYNOPSIS
  Scan VEERG Excel workbooks for remaining errors and report them in one place.

.DESCRIPTION
  A single, read-only checker that consolidates the error categories the build
  pipeline can leave behind. It reads the .xlsx package XML directly (no Excel,
  no COM, no add-in required) so it is fast and never produces the false #NAME?
  results a headless recalc would (VEERG.* LAMBDAs only resolve with the Excel
  Labs add-in loaded). It reports whatever was last computed and saved:

    [cell]      Cells whose cached value is an Excel error (t="e"), e.g.
                #REF!, #NAME?, #VALUE!, #DIV/0!, #N/A, #NULL!, #NUM!, #SPILL!,
                #CALC!. In a workbook flagged fullCalcOnLoad="1" (Excel recomputes
                every formula on open) these cached errors are PROVISIONAL - a
                stale one clears on recalc while a genuine one (missing sheet /
                name) persists - so they are HIDDEN by default and only counted;
                pass -IncludeCachedErrors to list them. A #REF! token in the
                formula TEXT is structural and is always reported.

    [name]      Defined names whose RefersTo contains an Excel error token
                (dangling / broken function) or is empty. Workbook- and
                sheet-scoped names are both checked. The .xlf-maintained library
                LAMBDAs that carry an internal #REF! (category broken-fn) are
                NOISE here and are hidden by default; pass -IncludeLibraryFunctions
                to list them too.

    [link]      Leftover external links to other workbook files. A
                self-contained VEERG workbook should have none.

    [sheet]     Formulas (and defined names) that reference a worksheet the
                workbook does not contain - a guaranteed #REF! on recalc that a
                cached-value scan can miss because the saved value is a stale-good
                number carried over from a source workbook where the sheet existed.

    [sum]       SUM(...) calls whose literal range argument (e.g. SUM(D10:D25))
                is a different size than the contiguous block of data actually
                sitting in those rows/columns. Catches the classic "rows were
                added/removed but the SUM range wasn't updated" mistake.
                Structured table references (Table[Column]) and plain named
                ranges are skipped on purpose - those resize automatically.
                Text cells reading "N/A", "Not used", "Summed above" or "Enter
                value" count as valid data, not a label, since they're a
                deliberate stand-in for a number. To document a SUM that is
                intentionally partial, add a no-op comment to the formula, e.g.
                =SUM(G128:G131)+N("Partial: subtotal 1 of 3, see also I138")
                - N() turns the text into 0 so the result is unaffected, and
                the checker skips any formula containing N("Partial...").
                Advisory only (heuristic, not included in -FailOnError).

    [shift]     A cell used as a direct operand of an arithmetic operator
                (+ - * / ^) holds text instead of a number or a formula.
                Catches the classic "formula was copied/pasted and a relative
                reference landed on the wrong cell" mistake. Only plain
                A1-style references are examined - a named reference
                (X_Cell_*, Result_*, VEERG_*, Table_*) is assumed correct and
                is never flagged, since it doesn't move when a formula is
                pasted elsewhere and so isn't at risk of this bug. Blank
                operands are NOT flagged (dropped 2026-08 - too many false
                positives from formulas that legitimately reference a
                not-yet-filled-in user input cell by plain address). Advisory
                only (heuristic, not included in -FailOnError).

    [series]    OFF BY DEFAULT - pass -IncludeSeriesCheck to run it. A
                function repeated across a run of 4+ adjacent cells in the
                same row/column (identical formula shape once cell refs are
                stripped out - the same shape Excel groups as "copied
                formula") should have each plain-reference operand shift in
                lockstep with the run, one cell per step, UNLESS it's a
                genuine shared single value. An operand that is FULLY
                $-anchored (both column AND row absolute, e.g. $D$63) in
                EVERY cell of the run is checked for NEITHER sub-case below
                - Excel never shifts a fully-absolute reference when a
                formula is filled/copied, so if its address varies between
                cells anyway, each cell's reference was necessarily typed
                or edited individually (e.g. manually picking specific rows
                out of a larger source table, in whatever order they're
                needed) - categorically not a fill/copy drift bug, however
                irregular the jumps look. Two sub-cases, both advisory:
                  Frozen - an operand lands on the exact same cell in every
                  formula in the run. A relative fill only auto-shifts a
                  reference along the run's OWN axis (column for a row-run,
                  row for a column-run), so $-anchoring JUST that one axis
                  (e.g. E$60 in a column-run - row-anchored, column not) is
                  already the complete, correct idiom for a fixed lookup
                  reference; the other axis's $ status is irrelevant and not
                  reported on. NOT reported when the axis that varies IS
                  $-anchored (regardless of the other axis). IS reported
                  when the axis that varies has NO $ at all - staying fixed
                  despite nothing anchoring the axis that should have moved
                  it is the genuinely ambiguous case worth a human look.
                  ALSO not reported when the un-anchored target cell falls
                  inside a "*_Arguments(" formula's dynamic-array spill (a
                  VEERG equation's argument-metadata call, e.g.
                  CHOOSECOLS(Module.Func_Arguments(),3,4) showing a unit
                  like "kg/head" next to a data range) - the spill's
                  non-anchor cells carry no formula of their own to trust,
                  but every row of a data block legitimately references the
                  same shared unit/label cell without $, so this is not a
                  copy-paste bug either. ALSO not reported (as either a run
                  member OR a referenced target) when a cell carries one of
                  a small set of named CELL STYLES VEERG authors apply to
                  mark non-data documentation cells - currently "Unit" /
                  "Unit no indent" (a unit label beside a data range) and
                  "Arguments" / "Arguments hyperlink" (equation
                  argument-metadata cells, including a HYPERLINK(...) link
                  to a source table, e.g. =HYPERLINK(Common_InputFunctions.
                  Utility_SourceHyperlink($D$77),$D$77) - repeated down a
                  column, these are documentation links, not a data
                  series). HYPERLINK(...) formulas are additionally
                  excluded by name regardless of style, as a second line of
                  defense.
                  Drift - an operand's step is an ISOLATED anomaly: it
                  differs from both neighbours and does not repeat for at
                  least $script:MinDriftSegmentLength consecutive steps of
                  its own - the classic "pasted one too many/too few times"
                  mistake, where exactly one cell has a wrong offset. A step
                  that DOES repeat for 2+ consecutive cells is treated as its
                  own legitimate sub-pattern rather than flagged - this is
                  what lets a genuine multi-range CONSOLIDATION (several
                  source ranges concatenated into one target range, each
                  internally consistent but discontinuous at the boundary)
                  pass cleanly instead of every boundary looking like drift.
                  A single-step (isolated) boundary is ALSO not reported
                  when the two cells straddling it fall inside two
                  DIFFERENT named ranges (e.g. crossing from one source
                  table like M1_Table_M_j_m5_T1 into another,
                  M1_Table_M_j_m_T2) - a real, deliberate cross-table
                  boundary, structurally indistinguishable from a one-cell
                  mistake by address alone, but recognizable via workbook
                  defined names. ALSO not reported when the isolated step
                  sits next to a well-established segment (>= $script:
                  MinBoundaryFlankRows rows/cols) - a within-table sub-group
                  boundary (e.g. two MMS-type row groups, like "Pasture
                  range and paddock" then "Anaerobic lagoon", or a single-
                  row group like "Solid storage", consolidated into ONE
                  named table) where no second name exists to check
                  against. For a COLUMN-run (varying by row) only ONE side
                  needs to be well-established, since a genuine sub-group
                  can be as short as one row with no internal pattern of
                  its own; a ROW-run (varying by column) still requires
                  BOTH sides. FINALLY, an operand is not checked for drift
                  AT ALL when its WHOLE step sequence is fully explained by
                  a repeating PERIOD (e.g. frozen for 2 rows, steps forward
                  once, repeating throughout the table - a fixed-size
                  row-group structure, like 3 rows per swine class sharing
                  one lookup row) - every single step must fit the cycle
                  exactly, so one real mistake anywhere still falls through
                  to the other checks. For an IRREGULARLY-sized grouping
                  (e.g. row-groups of 3, 4, 3, 3, 3, 4 - no fixed period),
                  a delta that recurs $script:MinDriftRepeatCount+ times
                  ANYWHERE in the run (even scattered, never twice
                  consecutively) is ALSO trusted - 3+ independent
                  occurrences of the identical wrong offset is not
                  something a real copy-paste mistake produces by chance.
                  Only the isolated cell(s) that remain after all five
                  checks are reported, not the whole run.
                Only plain A1-style references are examined, same as [shift]
                - named references are immune to this bug by construction.
                Cells whose formula text is inherited from an XLSX
                shared-formula master (no text of its own) are skipped -
                Excel computes those with a guaranteed-consistent relative
                offset, so they cannot exhibit this bug; only
                independently-stored formula text is at risk. Advisory only
                (heuristic, not included in -FailOnError).

  DRY / read-only ALWAYS: nothing is ever written. Use it after a build (e.g.
  `npm run build-enterprise`) to confirm the workbooks are clean.

.PARAMETER RepoRoot
  Repository root. Defaults to the parent of the scripts folder.

.PARAMETER WorkbookPath
  Full path to a single .xlsx to scan. If omitted, every Excel/*.xlsx and
  Excel/Enterprises/*.xlsx (excluding lock files, *_expanded*, *_template*,
  *_clean* generated copies and *.bak backups) is scanned.

.PARAMETER Max
  Maximum rows to list per category per workbook before summarising the rest.
  Default 50. Pass 0 for unlimited.

.PARAMETER FailOnError
  Exit with code 1 if any [cell] or [name] error is found (useful for CI /
  pre-commit gating). External links are always reported but never fail on their
  own. Off by default (always exits 0).

.PARAMETER IncludeLibraryFunctions
  Also report broken-fn names: the .xlf-maintained Excel Labs library LAMBDAs
  (Module.Func) whose body contains an internal #REF!. Hidden by default because
  they are known, tracked in the .xlf source, and never auto-deleted.

.PARAMETER SkipSumCheck
  Skip the [sum] SUM-range-size heuristic (see DESCRIPTION). Use this to speed
  up a scan when you only care about hard errors.

.PARAMETER SkipNumericOperandCheck
  Skip the [shift] numeric-operand-type heuristic (see DESCRIPTION).

.PARAMETER IncludeSeriesCheck
  Also run the [series] range-follow consistency heuristic (see DESCRIPTION).
  OFF by default - even after extensive false-positive tuning, genuine
  remaining hits are rare and the ones that do turn up tend to need a human
  judgment call, so it's opt-in rather than part of the default scan.

.PARAMETER IncludeCachedErrors
  List cached cell errors (t="e") even in a fullCalcOnLoad workbook, where they
  are provisional and hidden by default. Use it to hunt a genuine persistent
  error (e.g. a #REF! from a sheet that was never imported) among the stale ones
  that clear on Excel's on-open recalc.

.EXAMPLE
  npm run find-errors
  npm run find-errors -- -WorkbookPath .\Excel\Enterprises\Enterprise_Dairy_WIP_v01.xlsx
  npm run find-errors -- -Max 0 -FailOnError
  npm run find-errors -- -IncludeLibraryFunctions
  npm run find-errors -- -WorkbookPath .\Excel\Enterprises\Enterprise_CroppingGrains_WIP_v01.xlsx -IncludeCachedErrors
  npm run find-errors -- -IncludeSeriesCheck
#>
param(
  [string] $RepoRoot = (Split-Path $PSScriptRoot -Parent),
  [string] $WorkbookPath,
  [int]    $Max = 50,
  [switch] $IncludeLibraryFunctions,
  [switch] $SkipSumCheck,
  [switch] $SkipNumericOperandCheck,
  [switch] $IncludeSeriesCheck,
  [switch] $IncludeCachedErrors,
  [switch] $FailOnError
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

# Namespaces used across the OOXML parts.
$script:NsMain = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:NsRel  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:NsPkg  = 'http://schemas.openxmlformats.org/package/2006/relationships'

# Excel error tokens that can appear in a defined name's RefersTo.
$script:ErrorRegex = [regex]::new('#(REF!|NAME\?|DIV/0!|VALUE!|N/A|NULL!|NUM!|SPILL!|CALC!|FIELD!|GETTING_DATA|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!)', 'IgnoreCase')

# Sheet qualifier in a formula / RefersTo: 'Sheet Name'! or Sheet! (group 1 =
# quoted name with '' escapes, group 2 = unquoted). Used to flag a reference to a
# sheet the workbook does NOT contain - a guaranteed #REF! on recalc even when the
# cached value is a stale-good number (so a cached-only scan can't see it). The
# `#` in the unquoted lookbehind keeps error tokens (#REF! etc.) from matching.
$script:SheetQualifierRegex = [regex]::new("(?:'((?:[^']|'')*)'|(?<![A-Za-z0-9_.\[#])([A-Za-z_][A-Za-z0-9_.]*))!", 'IgnoreCase')
$script:StringLiteralRegex = [regex]::new('"(?:[^"]|"")*"')

# --- [sum] SUM-range-size check regexes -------------------------------------
# A call to SUM( - lookbehind blocks SUMIF/SUMIFS/SUMPRODUCT/SUMSQ/CUSTOM_SUM etc,
# since those have something other than '(' right after "SUM".
$script:SumCallRegex = [regex]::new('(?<![A-Za-z0-9_.])SUM\(', 'IgnoreCase')
# A single plain A1-style range argument, optionally sheet-qualified. Deliberately
# does NOT match structured refs (Table[Col]) or bare defined names - both are
# skipped on purpose (tables/names resize themselves).
$script:RangeArgRegex = [regex]::new("^(?:(?<sheet>'[^']+'|[A-Za-z_][A-Za-z0-9_]*)!)?\`$?(?<c1>[A-Za-z]{1,3})\`$?(?<r1>[0-9]+):\`$?(?<c2>[A-Za-z]{1,3})\`$?(?<r2>[0-9]+)$")
$script:CellRefRegex = [regex]::new('^([A-Za-z]+)([0-9]+)$')
# A cell is treated as a summary/total (not list data) ONLY when its formula
# aggregates a PLAIN contiguous range - the classic total pattern SUM(D10:D25)
# or SUBTOTAL(9,D10:D25). A value cell that merely uses SUM/SUBTOTAL/AGGREGATE
# over a structured table ref (e.g. SUM(Table[Col])) is genuine data and must
# still count, so it is NOT excluded. The optional leading `\d+,` handles the
# function-number first argument of SUBTOTAL/AGGREGATE.
$script:AggregateOfPlainRangeRegex = [regex]::new('(?<![A-Za-z0-9_])(?:SUM|SUBTOTAL|AGGREGATE)\s*\(\s*(?:\d+\s*,\s*)*(?:(?:''[^'']+''|[A-Za-z_][A-Za-z0-9_]*)!)?\$?[A-Za-z]{1,3}\$?\d+\s*:', 'IgnoreCase')
# Text cells holding one of these placeholder phrases are a deliberate stand-in
# for a number (SUM ignores text anyway) - treat the cell as valid data, not a
# label/title, so it doesn't look like the row/column is missing from the SUM.
$script:ValidPlaceholderTextRegex = [regex]::new('^\s*(n/a|not used|summed above|enter value|-|No data)\s*$', 'IgnoreCase')
# Deliberate-partial-sum marker: a formula carrying `+N("Partial...")` (or
# "Partial:"/"Partial -"/etc) documents in the formula bar, in plain human
# language, that the SUM range is intentionally smaller/larger than the
# surrounding data block. N() coerces text to 0, so it never changes the
# result. Any SUM range mismatch in a formula that carries this marker is
# skipped rather than reported.
$script:PartialSumMarkerRegex = [regex]::new('N\(\s*"\s*partial\b', 'IgnoreCase')

# --- [shift] numeric-operand type check regexes -----------------------------
# A plain A1-style cell reference, optionally sheet-qualified, guarded on both
# sides so it can't match a fragment of a longer identifier, a range endpoint
# (A1:B1), a structured-table ref, or a function name (LOG10(...), including
# with a space before the paren). Deliberately does NOT match defined names
# (X_Cell_*, Result_*, VEERG_*, ...) - those never have this col-then-digits
# shape, which is exactly the point: a named reference is immune to the
# copy-paste "reference shifted to the wrong cell" bug this check hunts for,
# since it doesn't move when the formula is pasted elsewhere, only a plain
# relative/absolute A1 reference can land on the wrong cell that way.
$script:ArithOperandRegex = [regex]::new(@'
(?<![A-Za-z0-9_:])
(?:(?<sheet>'(?:[^']|'')*'|[A-Za-z_][A-Za-z0-9_.]*)!)?
\$?(?<col>[A-Za-z]{1,3})\$?(?<row>[0-9]+)
(?![A-Za-z0-9_\[:])(?!\s*\()
'@, 'IgnorePatternWhitespace')
# A bracketed span - structured-table column name (Table[CO2-e factor]) or
# LAMBDA array literal. Column names are free text and can contain almost
# anything (hyphens, "CO2-e factor" reads as column "CO" row 2 to the cell-ref
# regex, with the following "-" then looking like an adjacent minus operator),
# so these must be stripped before scanning, same as string literals. Applied
# repeatedly to peel nested spans (Table[[#Headers],[Col]]) from the inside out.
$script:BracketSpanRegex = [regex]::new('\[[^\[\]]*\]')

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
  $dirs = @($excelDir, (Join-Path $excelDir 'Enterprises')) | Where-Object { Test-Path -LiteralPath $_ }
  $files = foreach ($d in $dirs) {
    Get-ChildItem -LiteralPath $d -Filter '*.xlsx' -File |
      Where-Object {
        $_.Name -notlike '~$*' -and
        $_.BaseName -notmatch '(?i)_expanded' -and
        $_.BaseName -notmatch '(?i)_template' -and
        $_.BaseName -notmatch '(?i)_clean' -and
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

function New-XmlDoc {
  param([string] $Text)
  $doc = New-Object System.Xml.XmlDocument
  $doc.PreserveWhitespace = $true
  $doc.LoadXml($Text)
  return ,$doc   # comma prevents PowerShell from enumerating the doc's child nodes
}

function New-NsManager {
  param($Doc, [hashtable] $Namespaces)
  $ns = New-Object System.Xml.XmlNamespaceManager($Doc.NameTable)
  foreach ($k in $Namespaces.Keys) { $ns.AddNamespace($k, $Namespaces[$k]) }
  return ,$ns   # comma prevents PowerShell from enumerating the manager's prefixes
}

# ---------------------------------------------------------------------------
# Map each worksheet to its part name, in workbook order.
#   Returns an ordered array of @{ Name; Part } plus a 0-based index lookup used
#   to resolve a defined name's localSheetId to a sheet name.
# ---------------------------------------------------------------------------
function Get-SheetMap {
  param($Zip)
  $result = @()
  $wbText = Read-ZipEntryText -Zip $Zip -EntryName 'xl/workbook.xml'
  if ($null -eq $wbText) { return $result }
  $relText = Read-ZipEntryText -Zip $Zip -EntryName 'xl/_rels/workbook.xml.rels'

  # rId -> target part
  $relMap = @{}
  if ($null -ne $relText) {
    $relDoc = New-XmlDoc -Text $relText
    $relNs = New-NsManager -Doc $relDoc -Namespaces @{ p = $script:NsPkg }
    foreach ($r in @($relDoc.SelectNodes('//p:Relationship', $relNs))) {
      $relMap[$r.GetAttribute('Id')] = $r.GetAttribute('Target')
    }
  }

  $wbDoc = New-XmlDoc -Text $wbText
  $wbNs = New-NsManager -Doc $wbDoc -Namespaces @{ x = $script:NsMain; r = $script:NsRel }
  foreach ($s in @($wbDoc.SelectNodes('//x:sheets/x:sheet', $wbNs))) {
    $name = $s.GetAttribute('name')
    $rid  = $s.GetAttribute('id', $script:NsRel)
    $part = $null
    if ($relMap.ContainsKey($rid)) {
      $tgt = $relMap[$rid]
      if ($tgt -match '^/') { $part = $tgt.TrimStart('/') } else { $part = 'xl/' + $tgt }
    }
    $result += [pscustomobject]@{ Name = $name; Part = $part }
  }
  return $result
}

# ---------------------------------------------------------------------------
# Whether the workbook is flagged for a full recalc the next time it's opened
# in Excel (xl/workbook.xml's <calcPr fullCalcOnLoad="1">). Excel sets this
# after structural edits it isn't sure the calc chain / cached values still
# cover (sheet copy/delete, name changes, etc.) - when set, EVERY cached cell
# value (t="e" included) is provisional and gets silently overwritten by a
# real recalculation on next open, so a cached #REF! here is not evidence of
# an actual error, just a stale snapshot from before the flag was set.
# ---------------------------------------------------------------------------
function Get-FullCalcOnLoad {
  param($Zip)
  $wbText = Read-ZipEntryText -Zip $Zip -EntryName 'xl/workbook.xml'
  if ($null -eq $wbText) { return $false }
  return [regex]::IsMatch($wbText, '<calcPr\b[^>]*\bfullCalcOnLoad="1"')
}

# ---------------------------------------------------------------------------
# Cell errors. Two kinds are reported:
#   * cached error   - the cell's stored value is an Excel error (t="e"), i.e. it
#                      currently EVALUATES to an error. In a workbook flagged
#                      fullCalcOnLoad="1" these are PROVISIONAL (Excel recomputes
#                      every formula on open) so the caller hides them by default
#                      and only counts them; each returned object carries IsCached
#                      so the caller can filter (see -IncludeCachedErrors).
#   * formula error  - the cell's stored <f> formula TEXT contains an error token
#                      (e.g. a #REF! left in the arguments) even though the cell
#                      currently evaluates to a valid value because a function
#                      swallowed the bad reference. A #REF! baked into a formula
#                      is always a real problem, so it is flagged regardless of
#                      the cached result or fullCalcOnLoad.
# ---------------------------------------------------------------------------
function Get-CellErrors {
  param($Zip, [array] $SheetMap)
  $errors = @()
  foreach ($sheet in $SheetMap) {
    if ([string]::IsNullOrEmpty($sheet.Part)) { continue }
    $text = Read-ZipEntryText -Zip $Zip -EntryName $sheet.Part
    if ($null -eq $text) { continue }
    # Fast skip: nothing to find unless the sheet has a cached error cell or an
    # error token somewhere in its text (e.g. #REF! inside a formula).
    if ($text.IndexOf('t="e"') -lt 0 -and -not $script:ErrorRegex.IsMatch($text)) { continue }
    $doc = New-XmlDoc -Text $text
    $ns = New-NsManager -Doc $doc -Namespaces @{ x = $script:NsMain }
    foreach ($c in @($doc.SelectNodes("//x:c[@t='e' or x:f]", $ns))) {
      $ref = $c.GetAttribute('r')
      $isCachedError = ($c.GetAttribute('t') -eq 'e')
      # Cells with a value-metadata pointer (vm) are RICH VALUES - an in-cell
      # ("Place in Cell") image, a linked data type, etc. Their t="e"/#VALUE! is
      # just the fallback text for clients that can't render the rich value, not
      # a real error, so skip them.
      if ($isCachedError -and -not [string]::IsNullOrEmpty($c.GetAttribute('vm'))) { continue }
      $vNode = $c.SelectSingleNode('x:v', $ns)
      $fNode = $c.SelectSingleNode('x:f', $ns)
      $formulaText = if ($null -ne $fNode) { $fNode.InnerText } else { '' }
      $formulaMatch = $script:ErrorRegex.Match($formulaText)

      if ($isCachedError) {
        # Currently evaluates to an error - reported regardless of fullCalcOnLoad.
        # A genuine #REF!/#NAME? (missing sheet/name) does NOT recalc away; and now
        # that the build no longer headless-recalcs, a recalc-transient cell saves a
        # GOOD cached value (not an error), so a cached error on disk is real.
        $val = if ($null -ne $vNode) { $vNode.InnerText } else { '#(error)' }
      } elseif ($formulaMatch.Success) {
        # Latent error: the formula carries an error token but the cell currently
        # evaluates fine (a function masked the bad reference).
        $val = $formulaMatch.Value + ' (in formula)'
      } else {
        continue   # ordinary formula cell, no error
      }
      $formula = if ($null -ne $fNode) { '=' + $formulaText } else { '' }
      $errors += [pscustomobject]@{
        Sheet = $sheet.Name; Cell = $ref; Error = $val; Formula = $formula; IsCached = $isCachedError
      }
    }
  }
  return $errors
}

# ---------------------------------------------------------------------------
# Broken defined names: RefersTo contains an error token or is empty.
# ---------------------------------------------------------------------------
function Get-NameErrors {
  param($Zip, [array] $SheetMap, [switch] $IncludeLibraryFunctions)
  $issues = @()
  $wbText = Read-ZipEntryText -Zip $Zip -EntryName 'xl/workbook.xml'
  if ($null -eq $wbText) { return $issues }
  $doc = New-XmlDoc -Text $wbText
  $ns = New-NsManager -Doc $doc -Namespaces @{ x = $script:NsMain }
  foreach ($n in @($doc.SelectNodes('//x:definedNames/x:definedName', $ns))) {
    $nm = $n.GetAttribute('name')
    if ([string]::IsNullOrEmpty($nm)) { continue }
    $rt = $n.InnerText
    $localId = $n.GetAttribute('localSheetId')
    $scope = 'workbook'
    if (-not [string]::IsNullOrEmpty($localId)) {
      $idx = 0; [void][int]::TryParse($localId, [ref]$idx)
      $scope = if ($idx -ge 0 -and $idx -lt $SheetMap.Count) { "'" + $SheetMap[$idx].Name + "'" } else { "sheet#$localId" }
    }
    $category = $null
    if ([string]::IsNullOrWhiteSpace($rt)) { $category = 'empty' }
    elseif ($script:ErrorRegex.IsMatch($rt)) {
      # A plain (non-formula) ref that errors is dangling cruft; a formula that
      # contains an error token is a broken function.
      $category = if ($rt.IndexOf('(') -ge 0) { 'broken-fn' } else { 'dangling' }
    }
    if ($null -eq $category) { continue }
    # broken-fn = .xlf library LAMBDAs with an internal #REF!: known noise, hidden
    # unless explicitly requested.
    if ($category -eq 'broken-fn' -and -not $IncludeLibraryFunctions) { continue }
    $issues += [pscustomobject]@{
      Name = $nm; Scope = $scope; Category = $category; RefersTo = $rt
    }
  }
  return $issues
}

# ---------------------------------------------------------------------------
# Leftover external links to other workbook files.
# ---------------------------------------------------------------------------
function Get-ExternalLinks {
  param($Zip)
  $links = @()
  foreach ($entry in $Zip.Entries) {
    if ($entry.FullName -notmatch '(?i)^xl/externalLinks/_rels/externalLink\d+\.xml\.rels$') { continue }
    $text = Read-ZipEntryText -Zip $Zip -EntryName $entry.FullName
    if ($null -eq $text) { continue }
    $doc = New-XmlDoc -Text $text
    $ns = New-NsManager -Doc $doc -Namespaces @{ p = $script:NsPkg }
    foreach ($r in @($doc.SelectNodes('//p:Relationship', $ns))) {
      $tgt = $r.GetAttribute('Target')
      if (-not [string]::IsNullOrWhiteSpace($tgt)) { $links += $tgt }
    }
  }
  return $links
}

# ---------------------------------------------------------------------------
# [sheet] References to a sheet the workbook does not contain. Scans cell
# formulas and defined-name RefersTo for `'Sheet'!` / `Sheet!` qualifiers and
# flags any sheet name not present in the workbook. Such a reference ALWAYS
# evaluates to #REF! on recalc, but the saved cached value can be a stale-good
# number (copied from a source workbook where the sheet existed), so the cached
# [cell] scan misses it - this static check catches it regardless. External
# `[N]` book references are left to the [link] category.
# ---------------------------------------------------------------------------
function Get-MissingSheetReferences {
  param($Zip, [array] $SheetMap)
  $issues = @()
  $sheetSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($s in $SheetMap) { if (-not [string]::IsNullOrEmpty($s.Name)) { [void] $sheetSet.Add($s.Name) } }

  $findMissing = {
    param([string] $Text)
    $missing = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrEmpty($Text) -or $Text.IndexOf('!') -lt 0) { return $missing }
    $noStr = $script:StringLiteralRegex.Replace($Text, '')   # drop "..." so INDIRECT("Sheet!..") strings don't match
    foreach ($m in $script:SheetQualifierRegex.Matches($noStr)) {
      $name = if ($m.Groups[1].Success) { $m.Groups[1].Value -replace "''", "'" } else { $m.Groups[2].Value }
      if ([string]::IsNullOrWhiteSpace($name) -or $name.IndexOf('[') -ge 0) { continue }   # external book ref
      foreach ($part in ($name -split ':')) {                                                # 3D ref 'S1:S2'
        $p = $part.Trim()
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        if (-not $sheetSet.Contains($p) -and -not $missing.Contains($p)) { $missing.Add($p) }
      }
    }
    return $missing
  }

  foreach ($sheet in $SheetMap) {
    if ([string]::IsNullOrEmpty($sheet.Part)) { continue }
    $text = Read-ZipEntryText -Zip $Zip -EntryName $sheet.Part
    if ($null -eq $text -or $text.IndexOf('!') -lt 0) { continue }
    $doc = New-XmlDoc -Text $text
    $ns = New-NsManager -Doc $doc -Namespaces @{ x = $script:NsMain }
    foreach ($fNode in @($doc.SelectNodes('//x:sheetData/x:row/x:c/x:f', $ns))) {
      $missing = @(& $findMissing $fNode.InnerText)
      if ($missing.Count -gt 0) {
        $issues += [pscustomobject]@{
          Location = ("'{0}'!{1}" -f $sheet.Name, $fNode.ParentNode.GetAttribute('r'))
          Missing  = ($missing -join ', ')
          Formula  = '=' + $fNode.InnerText
        }
      }
    }
  }

  # Defined names: plain range refs only (skip LAMBDA / function bodies).
  $wbText = Read-ZipEntryText -Zip $Zip -EntryName 'xl/workbook.xml'
  if ($null -ne $wbText) {
    $doc = New-XmlDoc -Text $wbText
    $ns = New-NsManager -Doc $doc -Namespaces @{ x = $script:NsMain }
    foreach ($n in @($doc.SelectNodes('//x:definedNames/x:definedName', $ns))) {
      $rt = $n.InnerText
      if ([string]::IsNullOrEmpty($rt) -or $rt.IndexOf('(') -ge 0) { continue }
      $missing = @(& $findMissing $rt)
      if ($missing.Count -gt 0) {
        $issues += [pscustomobject]@{ Location = ("name {0}" -f $n.GetAttribute('name')); Missing = ($missing -join ', '); Formula = $rt }
      }
    }
  }
  return $issues
}

# ---------------------------------------------------------------------------
# [sum] SUM-range-size check.
#
# For every literal-range SUM(...) argument, walk outward from the stated
# range along its own axis (a single-row SUM only looks left/right in that
# row; a single-column SUM only looks up/down in that column; a genuine
# multi-row+multi-column block looks in both axes) and find the true extent
# of contiguous data. A text cell (label/title) or a block-total cell (an
# aggregate over a PLAIN contiguous range, e.g. SUM(D10:D25)) stops the walk
# without being counted as data, so row/column titles at the near edge and
# "Total" cells at the far edge are both ignored, per design - but a value cell
# that sums a structured table ref (SUM(Table[Col])) still counts. If the
# actual extent differs from what the SUM
# argument states, it's reported.
# ---------------------------------------------------------------------------
function ConvertFrom-ColumnLetters {
  param([string] $Letters)
  $n = 0
  foreach ($ch in $Letters.ToUpperInvariant().ToCharArray()) {
    $n = $n * 26 + ([int]$ch - [int][char]'A' + 1)
  }
  return $n
}

function ConvertTo-ColumnLetters {
  param([int] $Number)
  $n = $Number
  $letters = ''
  while ($n -gt 0) {
    $rem = ($n - 1) % 26
    $letters = [string][char](65 + $rem) + $letters
    $n = [int](($n - $rem - 1) / 26)
  }
  return $letters
}

function Format-RangeAddr {
  param([int] $R1, [int] $R2, [int] $C1, [int] $C2)
  $a1 = (ConvertTo-ColumnLetters $C1) + $R1
  if ($R1 -eq $R2 -and $C1 -eq $C2) { return $a1 }
  $a2 = (ConvertTo-ColumnLetters $C2) + $R2
  return "$a1`:$a2"
}

# Extracts the text between the '(' at $OpenParenIndex and its matching ')',
# respecting nested parens and quoted strings (sheet names / string literals).
function Get-BalancedArgsText {
  param([string] $Text, [int] $OpenParenIndex)
  $depth = 0; $inSingle = $false; $inDouble = $false
  for ($i = $OpenParenIndex; $i -lt $Text.Length; $i++) {
    $ch = $Text[$i]
    if ($inSingle) { if ($ch -eq "'") { $inSingle = $false }; continue }
    if ($inDouble) { if ($ch -eq '"') { $inDouble = $false }; continue }
    switch ($ch) {
      "'"  { $inSingle = $true }
      '"'  { $inDouble = $true }
      '('  { $depth++ }
      ')'  {
        $depth--
        if ($depth -eq 0) { return $Text.Substring($OpenParenIndex + 1, $i - $OpenParenIndex - 1) }
      }
    }
  }
  return $null   # unbalanced - malformed/unsupported, caller should skip
}

# Splits on top-level commas only (ignores commas nested in parens/quotes).
function Split-TopLevelCommaText {
  param([string] $Text)
  $parts = New-Object System.Collections.Generic.List[string]
  $depth = 0; $inSingle = $false; $inDouble = $false; $start = 0
  for ($i = 0; $i -lt $Text.Length; $i++) {
    $ch = $Text[$i]
    if ($inSingle) { if ($ch -eq "'") { $inSingle = $false }; continue }
    if ($inDouble) { if ($ch -eq '"') { $inDouble = $false }; continue }
    switch ($ch) {
      "'"  { $inSingle = $true }
      '"'  { $inDouble = $true }
      '('  { $depth++ }
      ')'  { $depth-- }
      ','  { if ($depth -eq 0) { $parts.Add($Text.Substring($start, $i - $start)); $start = $i + 1 } }
    }
  }
  $parts.Add($Text.Substring($start))
  return ,@($parts)
}

# Parses xl/sharedStrings.xml into an index-ordered list of display strings
# (each <si> may be a plain <t> or several rich-text <r><t> runs to concatenate).
function Get-SharedStrings {
  param($Zip)
  $result = New-Object System.Collections.Generic.List[string]
  $text = Read-ZipEntryText -Zip $Zip -EntryName 'xl/sharedStrings.xml'
  if ($null -eq $text) { return ,$result }
  $doc = New-XmlDoc -Text $text
  $ns = New-NsManager -Doc $doc -Namespaces @{ x = $script:NsMain }
  foreach ($si in @($doc.SelectNodes('/x:sst/x:si', $ns))) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in @($si.SelectNodes('.//x:t', $ns))) { [void] $sb.Append($t.InnerText) }
    $result.Add($sb.ToString())
  }
  return ,$result
}

# Resolves the display text of a text-typed cell (t="s"/"str"/"inlineStr").
# Returns $null if it can't be resolved (e.g. a shared-string index out of range).
function Get-CellDisplayText {
  param($Cell, [string] $Type, $VNode, $Ns, $SharedStrings)
  switch ($Type) {
    's' {
      if ($null -eq $VNode -or [string]::IsNullOrEmpty($VNode.InnerText)) { return $null }
      $idx = 0
      if (-not [int]::TryParse($VNode.InnerText, [ref] $idx)) { return $null }
      if ($idx -lt 0 -or $idx -ge $SharedStrings.Count) { return $null }
      return $SharedStrings[$idx]
    }
    'str' { if ($null -eq $VNode) { return $null } else { return $VNode.InnerText } }
    'inlineStr' {
      $isNode = $Cell.SelectSingleNode('x:is', $Ns)
      if ($null -eq $isNode) { return $null }
      $sb = New-Object System.Text.StringBuilder
      foreach ($t in @($isNode.SelectNodes('.//x:t', $Ns))) { [void] $sb.Append($t.InnerText) }
      return $sb.ToString()
    }
    default { return $null }
  }
}

# Builds three sparse "row,col" -> $true maps for one sheet:
#   Numeric      - unconditional data: a non-empty, non-text value that is not
#                  itself a block total (an aggregate over a plain contiguous
#                  range, e.g. SUM(D10:D25)). A cell that sums a structured
#                  table ref (SUM(Table[Col])) is real data and still counts.
#   Placeholder  - text cells matching a recognised placeholder (N/A, Not used,
#                  Summed above). These are CONDITIONAL data: they only count
#                  when they sit among numbers, not among real text labels (see
#                  Test-StripHasData) - a "N/A" mixed into a label column is a
#                  label, not a stand-in number.
#   Label        - text cells that do NOT match a placeholder: real labels/
#                  titles. Never data; also what disqualifies a Placeholder
#                  cell from counting in the same row/column.
function Build-SheetCellMaps {
  param($Zip, $Sheet, $SharedStrings)
  $numeric = @{}; $placeholder = @{}; $label = @{}
  $result = [pscustomobject]@{ Numeric = $numeric; Placeholder = $placeholder; Label = $label }
  if ([string]::IsNullOrEmpty($Sheet.Part)) { return $result }
  $text = Read-ZipEntryText -Zip $Zip -EntryName $Sheet.Part
  if ($null -eq $text) { return $result }
  $doc = New-XmlDoc -Text $text
  $ns = New-NsManager -Doc $doc -Namespaces @{ x = $script:NsMain }
  foreach ($c in @($doc.SelectNodes('//x:sheetData/x:row/x:c', $ns))) {
    $ref = $c.GetAttribute('r')
    if ([string]::IsNullOrEmpty($ref)) { continue }
    $m = $script:CellRefRegex.Match($ref)
    if (-not $m.Success) { continue }
    $t = $c.GetAttribute('t')
    $vNode = $c.SelectSingleNode('x:v', $ns)
    $col = ConvertFrom-ColumnLetters $m.Groups[1].Value
    $row = [int]$m.Groups[2].Value
    $key = "$row,$col"
    if ($t -eq 's' -or $t -eq 'str' -or $t -eq 'inlineStr') {
      $cellText = Get-CellDisplayText -Cell $c -Type $t -VNode $vNode -Ns $ns -SharedStrings $SharedStrings
      if ([string]::IsNullOrEmpty($cellText)) { continue }   # empty text cell
      if ($script:ValidPlaceholderTextRegex.IsMatch($cellText)) { $placeholder[$key] = $true }
      else { $label[$key] = $true }   # real label/title
    } else {
      $fNode = $c.SelectSingleNode('x:f', $ns)
      if ($null -ne $fNode -and $script:AggregateOfPlainRangeRegex.IsMatch($fNode.InnerText)) { continue }   # block total over a plain range
      if ($null -eq $vNode -or [string]::IsNullOrEmpty($vNode.InnerText)) { continue }   # empty cell
      $numeric[$key] = $true
    }
  }
  return $result
}

# A strip [R1..R2]x[C1..C2] "has data" if it contains a Numeric cell outright,
# or contains a Placeholder cell AND the reference range [RefR1..RefR2]x
# [RefC1..RefC2] - "the other values in the column/row being summed" - has no
# real Label cell. The reference range is usually the strip itself (block
# sums, where the cross-axis already spans several cells); for a single-row/
# single-column sum it's the SUM's own stated range on that same row/column,
# since the strip there is just one candidate cell with nothing else to judge
# it by (see Resolve-AxisRange).
function Test-StripHasData {
  param($Maps, [int] $R1, [int] $R2, [int] $C1, [int] $C2, [int] $RefR1, [int] $RefR2, [int] $RefC1, [int] $RefC2)
  for ($r = $R1; $r -le $R2; $r++) {
    for ($c = $C1; $c -le $C2; $c++) {
      if ($Maps.Numeric.ContainsKey("$r,$c")) { return $true }
    }
  }
  $hasPlaceholder = $false
  for ($r = $R1; $r -le $R2; $r++) {
    for ($c = $C1; $c -le $C2; $c++) {
      if ($Maps.Placeholder.ContainsKey("$r,$c")) { $hasPlaceholder = $true }
    }
  }
  if (-not $hasPlaceholder) { return $false }
  for ($r = $RefR1; $r -le $RefR2; $r++) {
    for ($c = $RefC1; $c -le $RefC2; $c++) {
      if ($Maps.Label.ContainsKey("$r,$c")) { return $false }   # a real label sits alongside - this is a label column/row
    }
  }
  return $true
}

# Resolves the true [From,To] extent along one axis, walking outward from the
# stated [StateFrom,StateTo] while the cross-strip (fixed at CrossFrom..CrossTo
# on the OTHER axis) has data, or walking inward while it does NOT (handles a
# SUM that over-includes now-empty rows/cols after a deletion). Axis 'row'
# means the moving position is a row number (cross = a column range); 'col'
# means it's a column number (cross = a row range).
function Resolve-AxisRange {
  param($Maps, [string] $Axis, [int] $CrossFrom, [int] $CrossTo, [int] $StateFrom, [int] $StateTo, [int] $MinBound = 1)
  # Multi-cell cross (a genuine block axis) judges a candidate line by its own
  # span; a single-cell cross (a plain row/column sum) has nothing else to
  # compare within the strip, so it falls back to the SUM's own stated range
  # on that same fixed line.
  $crossIsBlock = ($CrossFrom -ne $CrossTo)
  $hasData = {
    param($pos)
    if ($Axis -eq 'row') {
      if ($crossIsBlock) {
        return Test-StripHasData -Maps $Maps -R1 $pos -R2 $pos -C1 $CrossFrom -C2 $CrossTo -RefR1 $pos -RefR2 $pos -RefC1 $CrossFrom -RefC2 $CrossTo
      } else {
        return Test-StripHasData -Maps $Maps -R1 $pos -R2 $pos -C1 $CrossFrom -C2 $CrossTo -RefR1 $StateFrom -RefR2 $StateTo -RefC1 $CrossFrom -RefC2 $CrossTo
      }
    } else {
      if ($crossIsBlock) {
        return Test-StripHasData -Maps $Maps -R1 $CrossFrom -R2 $CrossTo -C1 $pos -C2 $pos -RefR1 $CrossFrom -RefR2 $CrossTo -RefC1 $pos -RefC2 $pos
      } else {
        return Test-StripHasData -Maps $Maps -R1 $CrossFrom -R2 $CrossTo -C1 $pos -C2 $pos -RefR1 $CrossFrom -RefR2 $CrossTo -RefC1 $StateFrom -RefC2 $StateTo
      }
    }
  }
  $from = $StateFrom
  if (& $hasData $from) {
    while (($from - 1) -ge $MinBound -and (& $hasData ($from - 1))) { $from-- }
  } else {
    while ($from -lt $StateTo -and -not (& $hasData $from)) { $from++ }
  }
  $to = $StateTo
  if (& $hasData $to) {
    while (& $hasData ($to + 1)) { $to++ }
  } else {
    while ($to -gt $from -and -not (& $hasData $to)) { $to-- }
  }
  return [pscustomobject]@{ From = $from; To = $to }
}

function Get-SumRangeMismatches {
  param($Zip, [array] $SheetMap)
  $issues = @()
  $markedCount = 0
  $dataMapCache = @{}
  $sharedStrings = Get-SharedStrings -Zip $Zip

  foreach ($sheet in $SheetMap) {
    if ([string]::IsNullOrEmpty($sheet.Part)) { continue }
    $text = Read-ZipEntryText -Zip $Zip -EntryName $sheet.Part
    if ($null -eq $text) { continue }
    if ($text.IndexOf('SUM', [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
    $doc = New-XmlDoc -Text $text
    $ns = New-NsManager -Doc $doc -Namespaces @{ x = $script:NsMain }
    foreach ($fNode in @($doc.SelectNodes('//x:sheetData/x:row/x:c/x:f', $ns))) {
      $formulaText = $fNode.InnerText
      if ([string]::IsNullOrEmpty($formulaText)) { continue }   # e.g. non-master shared formula
      $sumMatches = @($script:SumCallRegex.Matches($formulaText))
      if ($sumMatches.Count -eq 0) { continue }
      if ($script:PartialSumMarkerRegex.IsMatch($formulaText)) { $markedCount++; continue }   # deliberately partial, documented in-cell
      $cellRef = $fNode.ParentNode.GetAttribute('r')

      foreach ($sm in $sumMatches) {
        $openParenIdx = $sm.Index + $sm.Length - 1
        $argsText = Get-BalancedArgsText -Text $formulaText -OpenParenIndex $openParenIdx
        if ($null -eq $argsText) { continue }

        foreach ($arg in (Split-TopLevelCommaText -Text $argsText)) {
          $a = $arg.Trim()
          $rm = $script:RangeArgRegex.Match($a)
          if (-not $rm.Success) { continue }   # not a plain range (table ref, named range, literal, nested call, ...)

          $rangeSheetName = if ($rm.Groups['sheet'].Success) { $rm.Groups['sheet'].Value.Trim("'") } else { $sheet.Name }
          $c1 = ConvertFrom-ColumnLetters $rm.Groups['c1'].Value
          $c2 = ConvertFrom-ColumnLetters $rm.Groups['c2'].Value
          $r1 = [int] $rm.Groups['r1'].Value
          $r2 = [int] $rm.Groups['r2'].Value
          if ($c1 -gt $c2) { $tmp = $c1; $c1 = $c2; $c2 = $tmp }
          if ($r1 -gt $r2) { $tmp = $r1; $r1 = $r2; $r2 = $tmp }
          if ($r1 -eq $r2 -and $c1 -eq $c2) { continue }   # single cell, not a range

          if (-not $dataMapCache.ContainsKey($rangeSheetName)) {
            $rangeSheetInfo = $SheetMap | Where-Object { $_.Name -eq $rangeSheetName } | Select-Object -First 1
            if ($null -eq $rangeSheetInfo) { $dataMapCache[$rangeSheetName] = $null }
            else { $dataMapCache[$rangeSheetName] = Build-SheetCellMaps -Zip $Zip -Sheet $rangeSheetInfo -SharedStrings $sharedStrings }
          }
          $maps = $dataMapCache[$rangeSheetName]
          if ($null -eq $maps) { continue }   # sheet not found (e.g. unresolved external ref) - skip rather than guess

          $actR1 = $r1; $actR2 = $r2; $actC1 = $c1; $actC2 = $c2
          if ($r1 -eq $r2) {
            # Single row, multiple columns: only look left/right within that row.
            $res = Resolve-AxisRange -Maps $maps -Axis 'col' -CrossFrom $r1 -CrossTo $r1 -StateFrom $c1 -StateTo $c2
            $actC1 = $res.From; $actC2 = $res.To
          } elseif ($c1 -eq $c2) {
            # Single column, multiple rows: only look up/down within that column.
            $res = Resolve-AxisRange -Maps $maps -Axis 'row' -CrossFrom $c1 -CrossTo $c1 -StateFrom $r1 -StateTo $r2
            $actR1 = $res.From; $actR2 = $res.To
          } else {
            # A genuine block: check both axes, each against the OTHER axis's
            # original stated span (not the evolving one - keeps the walk
            # simple/predictable and avoids one axis dragging the other along).
            $resCols = Resolve-AxisRange -Maps $maps -Axis 'col' -CrossFrom $r1 -CrossTo $r2 -StateFrom $c1 -StateTo $c2
            $resRows = Resolve-AxisRange -Maps $maps -Axis 'row' -CrossFrom $c1 -CrossTo $c2 -StateFrom $r1 -StateTo $r2
            $actC1 = $resCols.From; $actC2 = $resCols.To
            $actR1 = $resRows.From; $actR2 = $resRows.To
          }

          $stated = Format-RangeAddr -R1 $r1 -R2 $r2 -C1 $c1 -C2 $c2
          $actual = Format-RangeAddr -R1 $actR1 -R2 $actR2 -C1 $actC1 -C2 $actC2
          if ($stated -eq $actual) { continue }

          $issues += [pscustomobject]@{
            Sheet = $sheet.Name; Cell = $cellRef; RangeSheet = $rangeSheetName
            Stated = $stated; Actual = $actual; Formula = '=' + $formulaText
          }
        }
      }
    }
  }
  return [pscustomobject]@{ Issues = $issues; Marked = $markedCount }
}

# ---------------------------------------------------------------------------
# [shift] Numeric-operand type check.
#
# Hunts for the classic "formula was copied/pasted and a relative reference
# landed on the wrong cell" mistake: a cell used as a direct operand of an
# arithmetic operator (+ - * / ^) should hold a number, or a formula (trusted
# to produce one) - not hard-coded text. Blank is NOT flagged (dropped 2026-08 -
# a plain-address reference to a not-yet-filled user input cell is legitimately
# blank, and that dwarfed the genuine drifted-reference signal in practice).
# Only plain A1-style references are examined (see $script:ArithOperandRegex) - a named reference
# (X_Cell_*, Result_*, VEERG_*, Table_*) never has that shape and is immune to
# this bug in the first place, since it doesn't move when a formula is pasted
# elsewhere; nothing here second-guesses a named reference.
# ---------------------------------------------------------------------------
function Test-ArithOperatorAdjacent {
  # True if the cell-ref match at [Start,End) in $Text is immediately next to
  # (ignoring spaces) a +, -, *, / or ^ on either side.
  param([string] $Text, [int] $Start, [int] $End)
  $ops = '+-*/^'
  $i = $Start - 1
  while ($i -ge 0 -and $Text[$i] -eq ' ') { $i-- }
  if ($i -ge 0 -and $ops.IndexOf($Text[$i]) -ge 0) { return $true }
  $j = $End
  while ($j -lt $Text.Length -and $Text[$j] -eq ' ') { $j++ }
  if ($j -lt $Text.Length -and $ops.IndexOf($Text[$j]) -ge 0) { return $true }
  return $false
}

function Build-CellClassificationMap {
  # One entry per non-empty cell on the sheet: Kind is 'Formula' (has <f>,
  # trusted regardless of cached type/value), 'Numeric', 'Boolean' (both
  # coerce fine in arithmetic), or 'Text' (Display holds the resolved text).
  # A bare error CONSTANT (t="e", no <f> - vanishingly rare; the shared-
  # formula-follower shape that looks like this always still carries an
  # empty <f t="shared".../> node, which already counts as 'Formula') is
  # treated as 'Formula' too rather than adding a third flag-worthy kind that
  # would just duplicate the [cell] category's job.
  param($Zip, $Sheet, $SharedStrings)
  $map = @{}
  if ([string]::IsNullOrEmpty($Sheet.Part)) { return $map }
  $text = Read-ZipEntryText -Zip $Zip -EntryName $Sheet.Part
  if ($null -eq $text) { return $map }
  $doc = New-XmlDoc -Text $text
  $ns = New-NsManager -Doc $doc -Namespaces @{ x = $script:NsMain }
  foreach ($c in @($doc.SelectNodes('//x:sheetData/x:row/x:c', $ns))) {
    $ref = $c.GetAttribute('r')
    if ([string]::IsNullOrEmpty($ref)) { continue }
    if ($null -ne $c.SelectSingleNode('x:f', $ns)) { $map[$ref] = [pscustomobject]@{ Kind = 'Formula'; Text = $null }; continue }
    $t = $c.GetAttribute('t')
    $vNode = $c.SelectSingleNode('x:v', $ns)
    if ($t -eq 'e') { $map[$ref] = [pscustomobject]@{ Kind = 'Formula'; Text = $null }; continue }
    if ($t -eq 'b') { $map[$ref] = [pscustomobject]@{ Kind = 'Boolean'; Text = $null }; continue }
    if ($t -eq 's' -or $t -eq 'str' -or $t -eq 'inlineStr') {
      $display = Get-CellDisplayText -Cell $c -Type $t -VNode $vNode -Ns $ns -SharedStrings $SharedStrings
      $map[$ref] = [pscustomobject]@{ Kind = 'Text'; Text = $display }
      continue
    }
    if ($null -eq $vNode -or [string]::IsNullOrEmpty($vNode.InnerText)) { continue }   # blank - simply absent from the map
    $map[$ref] = [pscustomobject]@{ Kind = 'Numeric'; Text = $null }
  }
  return $map
}

function Get-NumericOperandIssues {
  param($Zip, [array] $SheetMap)
  $issues = @()
  $sharedStrings = Get-SharedStrings -Zip $Zip
  $classCache = @{}

  foreach ($sheet in $SheetMap) {
    if ([string]::IsNullOrEmpty($sheet.Part)) { continue }
    $text = Read-ZipEntryText -Zip $Zip -EntryName $sheet.Part
    if ($null -eq $text) { continue }
    if ($text.IndexOf('<f') -lt 0 -or $text.IndexOfAny([char[]] '+-*/^') -lt 0) { continue }   # fast skip
    $doc = New-XmlDoc -Text $text
    $ns = New-NsManager -Doc $doc -Namespaces @{ x = $script:NsMain }
    foreach ($fNode in @($doc.SelectNodes('//x:sheetData/x:row/x:c/x:f', $ns))) {
      $formulaText = $fNode.InnerText
      if ([string]::IsNullOrEmpty($formulaText)) { continue }   # e.g. non-master shared formula
      $noStr = $script:StringLiteralRegex.Replace($formulaText, '')
      do { $prev = $noStr; $noStr = $script:BracketSpanRegex.Replace($noStr, '') } while ($noStr -ne $prev)
      if ($noStr.IndexOfAny([char[]] '+-*/^') -lt 0) { continue }

      $cellRef = $fNode.ParentNode.GetAttribute('r')
      $seen = New-Object 'System.Collections.Generic.HashSet[string]'
      foreach ($m in $script:ArithOperandRegex.Matches($noStr)) {
        if (-not (Test-ArithOperatorAdjacent -Text $noStr -Start $m.Index -End ($m.Index + $m.Length))) { continue }

        $refSheetName = if ($m.Groups['sheet'].Success) { $m.Groups['sheet'].Value.Trim("'") -replace "''", "'" } else { $sheet.Name }
        $addr = $m.Groups['col'].Value.ToUpperInvariant() + $m.Groups['row'].Value

        if (-not $classCache.ContainsKey($refSheetName)) {
          $refSheetInfo = $SheetMap | Where-Object { $_.Name -eq $refSheetName } | Select-Object -First 1
          if ($null -eq $refSheetInfo) { $classCache[$refSheetName] = $null }
          else { $classCache[$refSheetName] = Build-CellClassificationMap -Zip $Zip -Sheet $refSheetInfo -SharedStrings $sharedStrings }
        }
        $map = $classCache[$refSheetName]
        if ($null -eq $map) { continue }   # unresolved sheet (external ref) - not this check's job

        if (-not $map.ContainsKey($addr)) { continue }   # blank - not flagged, see comment above
        $cls = $map[$addr]
        if ($cls.Kind -ne 'Text') { continue }

        $refLabel = if ($refSheetName -eq $sheet.Name) { $addr } else { "'$refSheetName'!$addr" }
        if (-not $seen.Add("$cellRef|$refLabel")) { continue }   # same operand hit twice in one formula (e.g. A1*A1)

        $detail = 'holds text "{0}"' -f $cls.Text
        $issues += [pscustomobject]@{
          Sheet = $sheet.Name; Cell = $cellRef; RefCell = $refLabel; Detail = $detail; Formula = '=' + $formulaText
        }
      }
    }
  }
  return $issues
}

# ---------------------------------------------------------------------------
# [series] Range-follow consistency check.
#
# A function repeated across a run of adjacent cells in the same row/column
# (identical formula shape once cell refs are stripped out - the same shape
# Excel itself groups as "copied formula") should have each plain-reference
# operand shift in lockstep with the run, one cell per step, UNLESS it is a
# genuine shared single value. Two sub-cases:
#   Frozen - an operand lands on the exact SAME cell in every formula in the
#   run. Reported regardless of $ absolute marking, so a deliberate common
#   value (e.g. a GWP lookup cell, usually $-anchored) and a reference that
#   should have moved but got left behind by a copy/paste (usually NOT
#   $-anchored) both surface for a human to tell apart.
#   Drift - an operand's step size is NOT consistent across the run (jumps by
#   2 for one step, lands on the wrong column, etc) - the classic "pasted one
#   too many/too few times" mistake. Only the cell(s) that break from the
#   run's majority step are reported, not the whole run.
# Only plain A1-style references are examined (same regex as [shift]) - a
# named reference is immune to this bug by construction and is never part of
# a run's comparison. Cells whose formula text is inherited from an XLSX
# shared-formula master (t="shared" with no text of its own) are skipped -
# Excel computes those with a guaranteed-consistent relative offset, so they
# cannot exhibit this bug in the first place; only independently-stored
# formula text (typed or pasted into each cell individually) is at risk.
# ---------------------------------------------------------------------------
$script:MinSeriesRunLength = 4   # need >=3 pairwise steps for "consistent vs not" to mean anything
$script:MinDriftSegmentLength = 2   # a step repeated this many times in a row is its own trusted sub-pattern, not a mistake - see Test-SeriesRun's Drift segmentation
$script:MinBoundaryFlankRows = 4   # both segments (each measured in CELLS = seg.Length + 1, not raw delta-count) flanking an isolated single-step transition must span at least this many rows/cols before the transition itself is trusted as a same-table sub-group boundary (e.g. two MMS-type row groups sharing one named range) rather than a mistake
$script:MinDriftRepeatCount = 3   # a delta value occurring at least this many times ANYWHERE in a run (even non-consecutively, no fixed period required) is trusted as a deliberate repeating structure rather than a mistake - see Test-SeriesRun's Drift segmentation

function Get-SheetFormulaCells {
  # One entry per cell with its OWN formula text (skips shared-formula
  # follower cells - see comment above).
  param($Zip, $Sheet)
  $result = New-Object System.Collections.Generic.List[object]
  if ([string]::IsNullOrEmpty($Sheet.Part)) { return $result }
  $text = Read-ZipEntryText -Zip $Zip -EntryName $Sheet.Part
  if ($null -eq $text) { return $result }
  $doc = New-XmlDoc -Text $text
  $ns = New-NsManager -Doc $doc -Namespaces @{ x = $script:NsMain }
  foreach ($c in @($doc.SelectNodes('//x:sheetData/x:row/x:c', $ns))) {
    $fNode = $c.SelectSingleNode('x:f', $ns)
    if ($null -eq $fNode) { continue }
    $formulaText = $fNode.InnerText
    if ([string]::IsNullOrEmpty($formulaText)) { continue }
    $ref = $c.GetAttribute('r')
    if ([string]::IsNullOrEmpty($ref)) { continue }
    $m = $script:CellRefRegex.Match($ref)
    if (-not $m.Success) { continue }
    $result.Add([pscustomobject]@{
      Col = ConvertFrom-ColumnLetters $m.Groups[1].Value
      Row = [int] $m.Groups[2].Value
      Addr = $ref
      FormulaText = $formulaText
    })
  }
  return $result
}

function Get-FormulaSkeleton {
  # Strips string literals + bracketed spans (as [shift] does), then replaces
  # every plain cell-ref match with a fixed placeholder so two formulas that
  # differ ONLY in which cells they reference compare as textually equal.
  # Returns $null when the formula has no plain refs at all (named-range-only
  # formulas can't exhibit this bug and are never part of a run), OR when the
  # formula is a HYPERLINK(...) call - a source-table link cell (VEERG's
  # Common_InputFunctions.Utility_SourceHyperlink convention, styled
  # "Arguments hyperlink") is documentation, not a data value following a
  # fill pattern, even when several sit in the same row/column.
  param([string] $FormulaText)
  if ($FormulaText.IndexOf('HYPERLINK(', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $null }
  $noStr = $script:StringLiteralRegex.Replace($FormulaText, '')
  do { $prev = $noStr; $noStr = $script:BracketSpanRegex.Replace($noStr, '') } while ($noStr -ne $prev)
  $matches = @($script:ArithOperandRegex.Matches($noStr))
  if ($matches.Count -eq 0) { return $null }
  $sb = New-Object System.Text.StringBuilder
  $refs = New-Object System.Collections.Generic.List[object]
  $pos = 0
  foreach ($m in $matches) {
    [void] $sb.Append($noStr.Substring($pos, $m.Index - $pos))
    [void] $sb.Append('@REF@')
    $sheetExplicit = $null
    if ($m.Groups['sheet'].Success) { $sheetExplicit = $m.Groups['sheet'].Value.Trim("'") -replace "''", "'" }
    $colGroup = $m.Groups['col']; $rowGroup = $m.Groups['row']
    $colAbs = ($colGroup.Index -gt 0 -and $noStr[$colGroup.Index - 1] -eq '$')
    $rowAbs = ($rowGroup.Index -gt 0 -and $noStr[$rowGroup.Index - 1] -eq '$')
    $refs.Add([pscustomobject]@{
      SheetExplicit = $sheetExplicit
      Col = ConvertFrom-ColumnLetters $colGroup.Value
      Row = [int] $rowGroup.Value
      ColAbs = $colAbs
      RowAbs = $rowAbs
      Addr = $m.Value
    })
    $pos = $m.Index + $m.Length
  }
  [void] $sb.Append($noStr.Substring($pos))
  return [pscustomobject]@{ Skeleton = $sb.ToString(); Refs = $refs }
}

function Expand-CellRange {
  # "F60:G65" -> every bare A1 address inside it, inclusive. Only ever called
  # on a dynamic-array spill footprint (a handful of cells), never a data
  # range, so no size guard is needed.
  param([string] $RangeAddr)
  $addrs = New-Object System.Collections.Generic.List[string]
  $parts = $RangeAddr -split ':'
  if ($parts.Count -ne 2) { return $addrs }
  $m1 = $script:CellRefRegex.Match($parts[0]); $m2 = $script:CellRefRegex.Match($parts[1])
  if (-not $m1.Success -or -not $m2.Success) { return $addrs }
  $c1 = ConvertFrom-ColumnLetters $m1.Groups[1].Value; $r1 = [int] $m1.Groups[2].Value
  $c2 = ConvertFrom-ColumnLetters $m2.Groups[1].Value; $r2 = [int] $m2.Groups[2].Value
  for ($r = $r1; $r -le $r2; $r++) {
    for ($c = $c1; $c -le $c2; $c++) { [void] $addrs.Add((ConvertTo-ColumnLetters $c) + $r) }
  }
  return $addrs
}

function Get-ArgumentsSpillCells {
  # Addresses on this sheet populated by a formula whose text calls a
  # "*_Arguments(" function - either the formula's own cell, or (since a
  # VEERG *_Arguments() equation returns a metadata array - argument
  # name/unit/etc, e.g. CHOOSECOLS(Module.Func_Arguments(),3,4)) every cell
  # in its dynamic-array SPILL footprint. A spill's non-anchor cells carry
  # no <f> of their own (just a cached <v>), so a plain reference landing on
  # one of them has no formula to trust - but it's legitimately a shared
  # unit/label lookup referenced identically by every row of a data block,
  # not a copy-paste bug, so [series]'s Frozen check treats it like a
  # deliberate common value.
  param($Zip, $Sheet)
  $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  if ([string]::IsNullOrEmpty($Sheet.Part)) { return ,$set }
  $text = Read-ZipEntryText -Zip $Zip -EntryName $Sheet.Part
  if ($null -eq $text -or $text.IndexOf('_Arguments(') -lt 0) { return ,$set }
  $doc = New-XmlDoc -Text $text
  $ns = New-NsManager -Doc $doc -Namespaces @{ x = $script:NsMain }
  foreach ($c in @($doc.SelectNodes('//x:sheetData/x:row/x:c', $ns))) {
    $fNode = $c.SelectSingleNode('x:f', $ns)
    if ($null -eq $fNode) { continue }
    $formulaText = $fNode.InnerText
    if ([string]::IsNullOrEmpty($formulaText) -or $formulaText.IndexOf('_Arguments(') -lt 0) { continue }
    $ref = $c.GetAttribute('r')
    if (-not [string]::IsNullOrEmpty($ref)) { [void] $set.Add($ref) }
    $spillRef = $fNode.GetAttribute('ref')
    if (-not [string]::IsNullOrEmpty($spillRef) -and $spillRef.Contains(':')) {
      foreach ($addr in (Expand-CellRange -RangeAddr $spillRef)) { [void] $set.Add($addr) }
    }
  }
  # Comma-operator prefix: a bare `return $set` lets PowerShell auto-enumerate
  # the HashSet into the output stream; when the caller captures it via simple
  # assignment ($x = Get-ArgumentsSpillCells ...), an EMPTY set enumerates to
  # zero output items and $x becomes $null (not an empty HashSet) - the next
  # `.Contains()` call then throws "cannot call a method on a null-valued
  # expression". `,$set` forces the whole HashSet through as ONE object.
  return ,$set
}

# VEERG authors apply specific NAMED CELL STYLES to mark non-data
# documentation/metadata cells - "Unit"/"Unit no indent" for a unit label
# next to a data range, "Arguments"/"Arguments hyperlink" for equation
# argument-metadata cells (including a HYPERLINK(...) source-table link).
# [series] excludes any cell carrying one of these styles from range-series
# checking entirely (as a run member AND as a referenced target), on top of
# (not instead of) the formula-text heuristics above - the style is the
# author's own authoritative signal of intent, more robust than inferring it
# from a specific function name every time a new documentation-cell shape
# turns up. Extend this list if another such style is found.
$script:SeriesExcludedStyleNames = @('Unit', 'Unit no indent', 'Arguments', 'Arguments hyperlink')

function Get-CellStyleNameMap {
  # Workbook-level (not per-sheet): maps a cellXfs INDEX (the value every
  # cell's `s="N"` attribute references) to the NAMED cell style it's based
  # on, via xl/styles.xml's standard <cellStyle name=".." xfId="N"/> ->
  # <cellStyleXfs> position <- <cellXfs><xf xfId="N"/> linkage. A cellXfs
  # entry with no xfId, or an xfId with no matching named cellStyle (e.g. 0 =
  # built-in "Normal"), is simply absent from the returned map.
  param($Zip)
  $map = @{}
  $text = Read-ZipEntryText -Zip $Zip -EntryName 'xl/styles.xml'
  if ($null -eq $text) { return $map }
  $doc = New-XmlDoc -Text $text
  $ns = New-NsManager -Doc $doc -Namespaces @{ x = $script:NsMain }
  $namesByXfId = @{}
  foreach ($cs in @($doc.SelectNodes('//x:cellStyles/x:cellStyle', $ns))) {
    $name = $cs.GetAttribute('name'); $xfId = $cs.GetAttribute('xfId')
    if (-not [string]::IsNullOrEmpty($name) -and -not [string]::IsNullOrEmpty($xfId)) { $namesByXfId[$xfId] = $name }
  }
  if ($namesByXfId.Count -eq 0) { return $map }
  $i = 0
  foreach ($xf in @($doc.SelectNodes('//x:cellXfs/x:xf', $ns))) {
    $xfId = $xf.GetAttribute('xfId')
    if (-not [string]::IsNullOrEmpty($xfId) -and $namesByXfId.ContainsKey($xfId)) { $map[[string] $i] = $namesByXfId[$xfId] }
    $i++
  }
  return $map
}

function Get-ExcludedStyleCells {
  # Addresses on this sheet whose applied named style is in
  # $script:SeriesExcludedStyleNames. Every return uses the `,$set` comma-
  # operator prefix - see the comment on Get-ArgumentsSpillCells's return for
  # why a bare `return $set` is unsafe here (an empty HashSet collapses to
  # $null when the caller captures it via simple assignment).
  param($Zip, $Sheet, [hashtable] $StyleNameMap)
  $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  if ($StyleNameMap.Count -eq 0) { return ,$set }
  if ([string]::IsNullOrEmpty($Sheet.Part)) { return ,$set }
  $text = Read-ZipEntryText -Zip $Zip -EntryName $Sheet.Part
  if ($null -eq $text) { return ,$set }
  $doc = New-XmlDoc -Text $text
  $ns = New-NsManager -Doc $doc -Namespaces @{ x = $script:NsMain }
  foreach ($c in @($doc.SelectNodes('//x:sheetData/x:row/x:c', $ns))) {
    $s = $c.GetAttribute('s')
    if ([string]::IsNullOrEmpty($s) -or -not $StyleNameMap.ContainsKey($s)) { continue }
    if ($script:SeriesExcludedStyleNames -contains $StyleNameMap[$s]) {
      $ref = $c.GetAttribute('r')
      if (-not [string]::IsNullOrEmpty($ref)) { [void] $set.Add($ref) }
    }
  }
  return ,$set
}

# Regex for a single-area RefersTo: optional 'Sheet'!/Sheet! qualifier, then
# a plain cell or A1:B5-style range. Deliberately narrower than
# $script:RangeArgRegex (which is anchored for a SUM argument) - this one
# also accepts a BARE single cell (a 1x1 "range"), and is applied to the
# WHOLE trimmed RefersTo text, not a substring match.
$script:NamedRangeAreaRegex = [regex]::new(@'
^(?:(?<sheet>'(?:[^']|'')*'|[A-Za-z_][A-Za-z0-9_.]*)!)?
\$?(?<c1>[A-Za-z]{1,3})\$?(?<r1>[0-9]+)
(?::\$?(?<c2>[A-Za-z]{1,3})\$?(?<r2>[0-9]+))?$
'@, 'IgnorePatternWhitespace')

function Get-NamedRangeAreas {
  # Every defined name (workbook- or sheet-scoped) whose RefersTo is a
  # SIMPLE, single-rectangle A1 reference (no error token, no multi-area
  # union, no structured-table syntax). Used by [series]'s Drift check to
  # recognize when an anomalous step is actually the BOUNDARY between two
  # different named ranges - a deliberate multi-range CONSOLIDATION (e.g.
  # three separate manure-management source tables concatenated into one
  # target list, each with its own name like M1_Table_M_j_m5_T1) rather than
  # a genuine single-cell copy-paste mistake. This is a STRUCTURAL signal
  # (crossing a real, named boundary) rather than inferring intent from cell
  # VALUES/labels, which is both more robust and far cheaper to check.
  param($Zip)
  $areas = New-Object System.Collections.Generic.List[object]
  $text = Read-ZipEntryText -Zip $Zip -EntryName 'xl/workbook.xml'
  if ($null -eq $text) { return ,$areas }
  $doc = New-XmlDoc -Text $text
  $ns = New-NsManager -Doc $doc -Namespaces @{ x = $script:NsMain }
  $sheetsNode = @($doc.SelectNodes('//x:sheets/x:sheet', $ns))
  foreach ($n in @($doc.SelectNodes('//x:definedNames/x:definedName', $ns))) {
    $rt = $n.InnerText
    if ([string]::IsNullOrWhiteSpace($rt) -or $rt.IndexOf(',') -ge 0) { continue }   # empty or multi-area union
    if ($script:ErrorRegex.IsMatch($rt)) { continue }
    $m = $script:NamedRangeAreaRegex.Match($rt.Trim())
    if (-not $m.Success) { continue }
    $sheetName = $null
    if ($m.Groups['sheet'].Success) { $sheetName = $m.Groups['sheet'].Value.Trim("'") -replace "''", "'" }
    else {
      $localId = $n.GetAttribute('localSheetId')
      if (-not [string]::IsNullOrEmpty($localId)) {
        $idx = 0; [void] [int]::TryParse($localId, [ref] $idx)
        if ($idx -ge 0 -and $idx -lt $sheetsNode.Count) { $sheetName = $sheetsNode[$idx].GetAttribute('name') }
      }
    }
    if ([string]::IsNullOrEmpty($sheetName)) { continue }   # workbook-scoped name with no sheet qualifier - can't place it
    $c1 = ConvertFrom-ColumnLetters $m.Groups['c1'].Value; $r1 = [int] $m.Groups['r1'].Value
    $c2 = if ($m.Groups['c2'].Success) { ConvertFrom-ColumnLetters $m.Groups['c2'].Value } else { $c1 }
    $r2 = if ($m.Groups['r2'].Success) { [int] $m.Groups['r2'].Value } else { $r1 }
    [void] $areas.Add([pscustomobject]@{
      Name = $n.GetAttribute('name'); Sheet = $sheetName
      C1 = [Math]::Min($c1, $c2); C2 = [Math]::Max($c1, $c2)
      R1 = [Math]::Min($r1, $r2); R2 = [Math]::Max($r1, $r2)
    })
  }
  return ,$areas
}

function Find-EnclosingRangeName {
  # First named-range area (from Get-NamedRangeAreas) whose rectangle
  # contains (SheetName, Col, Row), or $null if none does. Overlapping named
  # ranges are rare and not disambiguated - the first match is used.
  param([System.Collections.Generic.List[object]] $Areas, [string] $SheetName, [int] $Col, [int] $Row)
  foreach ($a in $Areas) {
    if ($a.Sheet -ne $SheetName) { continue }
    if ($Col -ge $a.C1 -and $Col -le $a.C2 -and $Row -ge $a.R1 -and $Row -le $a.R2) { return $a.Name }
  }
  return $null
}

function Find-SeriesRuns {
  # Walks cells already sorted along one axis, splitting into maximal runs of
  # cells that are BOTH positionally contiguous (no gap) AND share an
  # identical formula skeleton. AxisIsColumn = $true means position is Col
  # (a row-run, reading across columns); $false means position is Row (a
  # column-run, reading down rows).
  param([array] $OrderedCells, [hashtable] $Parsed, [bool] $AxisIsColumn)
  $runs = New-Object System.Collections.Generic.List[object]
  $current = New-Object System.Collections.Generic.List[object]
  $prevPos = $null
  $prevSkeleton = $null
  foreach ($cell in $OrderedCells) {
    $pos = if ($AxisIsColumn) { $cell.Col } else { $cell.Row }
    if (-not $Parsed.ContainsKey($cell.Addr)) {
      if ($current.Count -ge $script:MinSeriesRunLength) { $runs.Add(@($current.ToArray())) }
      $current.Clear(); $prevPos = $null; $prevSkeleton = $null
      continue
    }
    $sk = $Parsed[$cell.Addr].Skeleton
    $contiguous = ($null -ne $prevPos) -and ($pos -eq $prevPos + 1) -and ($sk -eq $prevSkeleton)
    if (-not $contiguous -and $current.Count -gt 0) {
      if ($current.Count -ge $script:MinSeriesRunLength) { $runs.Add(@($current.ToArray())) }
      $current.Clear()
    }
    $current.Add($cell)
    $prevPos = $pos
    $prevSkeleton = $sk
  }
  if ($current.Count -ge $script:MinSeriesRunLength) { $runs.Add(@($current.ToArray())) }
  return $runs
}

function Test-PeriodicDeltaPattern {
  # Smallest period P (>=2) for which the WHOLE delta sequence is fully
  # explained - every index i has the same key as every other index sharing
  # (i mod P), across at least 2 complete cycles (P is capped at
  # floor(Count/2) so that's guaranteed). Returns $null if no such P exists
  # for P in [2, floor(Count/2)]. Deliberately requires a PERFECT fit (no
  # exceptions at all) - a single genuine mistake anywhere in the run breaks
  # every candidate period, so this only ever fires for a truly regular,
  # deliberate repeating structure (e.g. a fixed-size row-group table),
  # never as a way to partially excuse a mostly-consistent-but-buggy run.
  param([System.Collections.Generic.List[string]] $Keys)
  $n = $Keys.Count
  for ($p = 2; $p -le [Math]::Floor($n / 2); $p++) {
    $phaseValue = @{}
    $ok = $true
    for ($i = 0; $i -lt $n; $i++) {
      $r = $i % $p
      if (-not $phaseValue.ContainsKey($r)) { $phaseValue[$r] = $Keys[$i] }
      elseif ($phaseValue[$r] -ne $Keys[$i]) { $ok = $false; break }
    }
    if ($ok) { return $p }
  }
  return $null
}

function Test-SeriesRun {
  # Classifies every ref "slot" (Nth plain reference, in textual order) shared
  # by every formula in the run. A slot's step is the (sheet-changed?, dCol,
  # dRow) delta between consecutive cells' reference at that slot.
  param([array] $RunCells, [hashtable] $Parsed, [string] $HostSheet, [bool] $AxisIsColumn, $Zip, [array] $SheetMap, [hashtable] $ArgsSpillCache, [hashtable] $StyleNameMap, [hashtable] $StyleCellCache, [System.Collections.Generic.List[object]] $NamedRangeAreas)
  $issues = New-Object System.Collections.Generic.List[object]
  $refCount = $Parsed[$RunCells[0].Addr].Refs.Count
  for ($slot = 0; $slot -lt $refCount; $slot++) {
    $slotRefs = @($RunCells | ForEach-Object { $Parsed[$_.Addr].Refs[$slot] })

    # A FULLY $-anchored operand (both Col AND Row absolute, e.g. $D$63) in
    # EVERY cell of the run is never a copy-paste-fill drift bug, regardless
    # of whether its address is constant or varies between cells: Excel
    # never shifts a fully-absolute reference when a formula is filled/
    # copied, so if the address DOES differ between cells, each cell's
    # reference was necessarily typed/edited individually - a deliberate,
    # manually-curated lookup into a source table (e.g. a source list with
    # more rows than needed, picking specific ones out of order), not a
    # relative reference that "should have" followed a fill pattern. This
    # is the same reasoning the Frozen check already applies when an
    # operand is fully anchored and constant; here it's extended to the
    # case where it's fully anchored but VARIES (arbitrary, non-monotonic,
    # sometimes non-repeating jumps) - skip both Frozen and Drift for this
    # slot entirely, before any step/pattern analysis.
    $allFullyAnchored = $true
    foreach ($r in $slotRefs) { if (-not ($r.ColAbs -and $r.RowAbs)) { $allFullyAnchored = $false; break } }
    if ($allFullyAnchored) { continue }

    $keys = New-Object System.Collections.Generic.List[string]
    $dCols = New-Object System.Collections.Generic.List[int]
    $dRows = New-Object System.Collections.Generic.List[int]
    $sameSheets = New-Object System.Collections.Generic.List[bool]
    for ($i = 1; $i -lt $slotRefs.Count; $i++) {
      $a = $slotRefs[$i - 1]; $b = $slotRefs[$i]
      $sheetA = if ($a.SheetExplicit) { $a.SheetExplicit } else { $HostSheet }
      $sheetB = if ($b.SheetExplicit) { $b.SheetExplicit } else { $HostSheet }
      $same = ($sheetA -eq $sheetB)
      $dc = $b.Col - $a.Col
      $dr = $b.Row - $a.Row
      $key = if ($same) { "$dc,$dr" } else { 'X' }
      [void] $keys.Add($key); [void] $dCols.Add($dc); [void] $dRows.Add($dr); [void] $sameSheets.Add($same)
    }

    $distinctKeys = @($keys | Select-Object -Unique)
    if ($distinctKeys.Count -le 1) {
      if ($distinctKeys.Count -eq 1 -and $distinctKeys[0] -eq '0,0') {
        $firstRef = $slotRefs[0]
        # A relative fill only auto-shifts a reference along the run's OWN
        # axis (column for a row-run, row for a column-run) - the other axis
        # can never move regardless of $ status, so $-anchoring JUST the
        # axis that varies (e.g. E$60 in a column-run, row-anchored, column
        # not) is already the complete, correct, unambiguous idiom for
        # pinning a lookup-header reference during a fill - the other axis's
        # $ status is irrelevant noise, not evidence of anything. Only
        # suppress-worthy when the axis-relevant flag is absolute; only
        # flag when it's NOT, since that's the case where "frozen despite
        # being relative on the axis that should have moved it" is
        # genuinely suspicious.
        $axisAnchored = if ($AxisIsColumn) { $firstRef.ColAbs } else { $firstRef.RowAbs }
        $isArgsSpill = $false
        $isStyleExcluded = $false
        if (-not $axisAnchored) {
          $effSheet = if ($firstRef.SheetExplicit) { $firstRef.SheetExplicit } else { $HostSheet }
          if (-not $ArgsSpillCache.ContainsKey($effSheet)) {
            $sheetInfo = $SheetMap | Where-Object { $_.Name -eq $effSheet } | Select-Object -First 1
            $ArgsSpillCache[$effSheet] = if ($null -eq $sheetInfo) { New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase) } else { Get-ArgumentsSpillCells -Zip $Zip -Sheet $sheetInfo }
          }
          if (-not $StyleCellCache.ContainsKey($effSheet)) {
            $sheetInfo2 = $SheetMap | Where-Object { $_.Name -eq $effSheet } | Select-Object -First 1
            $StyleCellCache[$effSheet] = if ($null -eq $sheetInfo2) { New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase) } else { Get-ExcludedStyleCells -Zip $Zip -Sheet $sheetInfo2 -StyleNameMap $StyleNameMap }
          }
          $bareAddr = (ConvertTo-ColumnLetters $firstRef.Col) + $firstRef.Row
          $isArgsSpill = $ArgsSpillCache[$effSheet].Contains($bareAddr)
          $isStyleExcluded = $StyleCellCache[$effSheet].Contains($bareAddr)
        }
        if (-not $axisAnchored -and -not $isArgsSpill -and -not $isStyleExcluded) {
          $runLabel = "{0}{1}:{2}{3}" -f (ConvertTo-ColumnLetters $RunCells[0].Col), $RunCells[0].Row, (ConvertTo-ColumnLetters $RunCells[-1].Col), $RunCells[-1].Row
          [void] $issues.Add([pscustomobject]@{
            Sheet = $HostSheet; Kind = 'Frozen'; Cell = $runLabel
            Detail = ("operand #{0} stays fixed at {1} across the whole series {2} (not `$-anchored on the axis that varies - confirm it was meant to stay fixed rather than follow the series)" -f ($slot + 1), $firstRef.Addr, $runLabel)
            Formula = '=' + $RunCells[0].FormulaText
          })
        }
      }
      continue   # consistent non-zero step across the whole run - healthy parallel-range operand
    }

    # A REPEATING GROUP pattern (e.g. 3 target rows share one source lookup
    # row, then it steps forward once, repeating throughout the whole
    # table - "mostly frozen, steps once per group boundary") looks like
    # drift under the segment logic below, because the "step" delta never
    # repeats CONSECUTIVELY (each group boundary is a single isolated step,
    # forever) even though it recurs many times overall. Checked BEFORE
    # segmentation: if the WHOLE delta sequence for this slot is fully
    # explained by some period P (every position with the same i-mod-P has
    # the identical delta, for at least 2 complete cycles), it's a genuine
    # periodic/grouped structure, not a mistake - skip Drift entirely for
    # this slot. Deliberately strict (must explain EVERY step, not just
    # most) so a real bug breaking the cycle still falls through to the
    # segment-based check below; a table with IRREGULARLY sized groups
    # (not a fixed period) is not caught by this and may still be flagged.
    $periodicP = Test-PeriodicDeltaPattern -Keys $keys
    if ($null -ne $periodicP) { continue }

    # Drift: run-length-encode the step sequence into segments of consecutive
    # equal deltas, rather than comparing every step to one global majority.
    # A segment of length >= $script:MinDriftSegmentLength (repeated 2+ times
    # in a row) is treated as ITS OWN legitimate sub-pattern, not a mistake -
    # this is what lets a genuine multi-range CONSOLIDATION (e.g. three
    # source manure-management tables concatenated into one target range,
    # each internally consistent but discontinuous at the boundary) pass
    # cleanly, even though the overall run has multiple distinct steps. Only
    # an ISOLATED step (a segment shorter than the threshold - a lone cell
    # whose step differs from both neighbours) is flagged: a real
    # copy-paste-drift bug essentially never reproduces the SAME wrong
    # offset for two-plus consecutive cells by chance, so a sustained
    # alternate step is far more likely to be a second (or third...) deliberate
    # source range than a mistake. A remaining ambiguous case - an isolated
    # single-step BOUNDARY between two otherwise-consistent segments, which
    # looks identical whether it's a genuine one-cell mistake or a
    # deliberate multi-range consolidation boundary - is resolved below via
    # named-range membership (Get-NamedRangeAreas): crossing from one named
    # range into a different one is treated as a legitimate boundary, not
    # flagged. KNOWN GAP: a consolidation between UNNAMED plain cell blocks
    # (no defined name backing either side) still can't be disambiguated
    # this way and may still be flagged if its sub-range is only 1-2 rows
    # long (too short to earn MinDriftSegmentLength's trust on its own).
    $segments = New-Object System.Collections.Generic.List[object]
    $segStart = 0
    for ($i = 1; $i -le $keys.Count; $i++) {
      if ($i -eq $keys.Count -or $keys[$i] -ne $keys[$segStart]) {
        [void] $segments.Add([pscustomobject]@{ Key = $keys[$segStart]; Start = $segStart; Length = ($i - $segStart) })
        $segStart = $i
      }
    }

    # A key that recurs $script:MinDriftRepeatCount+ times ANYWHERE in the
    # run - even scattered, never twice consecutively - is ALSO trusted, on
    # top of the consecutive-segment rule above. This catches an
    # IRREGULARLY-sized repeating group table (e.g. row-groups of 3, 4, 3,
    # 3, 3, 4 rows - no fixed period for Test-PeriodicDeltaPattern to find,
    # but the same "step" delta still shows up several separate times) that
    # neither the segment-length rule nor the periodicity check above can
    # explain on their own. 3+ independent occurrences of the exact same
    # delta is strong evidence of a deliberate repeating structure - a
    # genuine copy-paste mistake essentially never reproduces the identical
    # wrong offset three separate times by chance.
    $totalCounts = @{}
    foreach ($k in $keys) { if (-not $totalCounts.ContainsKey($k)) { $totalCounts[$k] = 0 }; $totalCounts[$k]++ }

    $trustedCoverage = @{}
    foreach ($seg in $segments) {
      if ($seg.Length -ge $script:MinDriftSegmentLength -or $totalCounts[$seg.Key] -ge $script:MinDriftRepeatCount) {
        if (-not $trustedCoverage.ContainsKey($seg.Key)) { $trustedCoverage[$seg.Key] = 0 }
        $trustedCoverage[$seg.Key] += $seg.Length
      }
    }
    if ($trustedCoverage.Count -gt 0) {
      $majorityKey = ($trustedCoverage.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
    } else {
      # No segment reaches the trust threshold at all (every step differs
      # from its neighbours) - fall back to plain frequency, as before.
      $majorityKey = ($totalCounts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
    }

    for ($si = 0; $si -lt $segments.Count; $si++) {
      $seg = $segments[$si]
      if ($seg.Length -ge $script:MinDriftSegmentLength -or $totalCounts[$seg.Key] -ge $script:MinDriftRepeatCount) { continue }   # sustained OR non-consecutively-repeating pattern - a legitimate consolidated sub-range, not a mistake

      # An isolated (short) segment directly between two segments is trusted
      # as a within-table sub-group boundary rather than a mistake - e.g.
      # two MMS-type row groups (Pasture range and paddock / Anaerobic
      # lagoon) consolidated into ONE named table, where the named-range
      # check below can't help since both sides share the SAME name. A
      # segment's Length is a DELTA count, one less than its cell count
      # (e.g. a 4-row block is Length=3), so the check is against
      # Length+1 = $script:MinBoundaryFlankRows.
      # For a COLUMN-run (AxisIsColumn=$false, varying by ROW - the shape
      # every example so far has been), only ONE flank needs to be
      # well-established: a genuine sub-group can be as short as ONE row
      # (e.g. "Solid storage" - a single-row MMS-type group with no
      # internal repetition of its own to prove itself), so requiring BOTH
      # neighbours to be long was too strict for this axis. For a ROW-run
      # (AxisIsColumn=$true, varying by COLUMN), keep requiring BOTH flanks
      # - no evidence yet that consolidation happens along that axis, so
      # stay conservative there. Tradeoff (accepted): a genuine one-cell
      # mistake sitting at the EDGE of a long uniform column-run (next to
      # only one long flank, nothing on the other side) can now also be
      # missed - judged less likely in practice than a real short group.
      $prevSeg = if ($si -gt 0) { $segments[$si - 1] } else { $null }
      $nextSeg = if ($si -lt $segments.Count - 1) { $segments[$si + 1] } else { $null }
      $prevOk = ($null -ne $prevSeg -and ($prevSeg.Length + 1) -ge $script:MinBoundaryFlankRows)
      $nextOk = ($null -ne $nextSeg -and ($nextSeg.Length + 1) -ge $script:MinBoundaryFlankRows)
      $isTrustedBoundary = if (-not $AxisIsColumn) { $prevOk -or $nextOk } else { $prevOk -and $nextOk }
      if ($isTrustedBoundary) { continue }

      for ($j = 0; $j -lt $seg.Length; $j++) {
        $i = $seg.Start + $j
        if ($keys[$i] -eq $majorityKey) { continue }

        # A length-1 (or 2-, up to the threshold) anomaly sandwiched between
        # two otherwise-normal steps is structurally IDENTICAL whether it's
        # a genuine single-cell mistake or the boundary of a deliberate
        # multi-range consolidation - address deltas alone can't tell them
        # apart. But if the FROM and TO cells at this exact step land inside
        # two DIFFERENT known named ranges (e.g. the last row of
        # M1_Table_M_j_m5_T1 and the first row of M1_Table_M_j_m_T2), that's
        # a real, deliberate boundary between two named source tables, not a
        # mistake - suppress it. Still flagged when both sides are in the
        # SAME named range (or neither is in any named range at all), since
        # that's the genuine single-cell-drift shape this check exists for.
        $fromRef = $slotRefs[$i]; $toRef = $slotRefs[$i + 1]
        $fromSheet = if ($fromRef.SheetExplicit) { $fromRef.SheetExplicit } else { $HostSheet }
        $toSheet = if ($toRef.SheetExplicit) { $toRef.SheetExplicit } else { $HostSheet }
        $fromRangeName = Find-EnclosingRangeName -Areas $NamedRangeAreas -SheetName $fromSheet -Col $fromRef.Col -Row $fromRef.Row
        $toRangeName = Find-EnclosingRangeName -Areas $NamedRangeAreas -SheetName $toSheet -Col $toRef.Col -Row $toRef.Row
        if ($null -ne $fromRangeName -and $null -ne $toRangeName -and $fromRangeName -ne $toRangeName) { continue }

        $fromCell = $RunCells[$i]; $toCell = $RunCells[$i + 1]
        $expected = if ($majorityKey -eq 'X') { 'a same-sheet reference' } else { "step ($majorityKey)" }
        $actual = if ($sameSheets[$i]) { "step ($($dCols[$i]),$($dRows[$i]))" } else { 'a different sheet' }
        [void] $issues.Add([pscustomobject]@{
          Sheet = $HostSheet; Kind = 'Drift'; Cell = $toCell.Addr
          Detail = ("operand #{0} at {1} breaks the run's pattern: expected {2} from {3}, found {4}" -f ($slot + 1), $toCell.Addr, $expected, $fromCell.Addr, $actual)
          Formula = '=' + $toCell.FormulaText
        })
      }
    }
  }
  return $issues
}

function Get-SeriesConsistencyIssues {
  param($Zip, [array] $SheetMap)
  $issues = New-Object System.Collections.Generic.List[object]
  $argsSpillCache = @{}    # sheet name -> HashSet[string], built lazily, shared across the whole scan
  $styleCellCache = @{}    # sheet name -> HashSet[string] of excluded-style addresses, built lazily
  $styleNameMap = Get-CellStyleNameMap -Zip $Zip   # workbook-level, computed once
  $namedRangeAreas = Get-NamedRangeAreas -Zip $Zip   # workbook-level, computed once
  foreach ($sheet in $SheetMap) {
    $cells = @(Get-SheetFormulaCells -Zip $Zip -Sheet $sheet)
    if ($cells.Count -lt $script:MinSeriesRunLength) { continue }

    if (-not $styleCellCache.ContainsKey($sheet.Name)) {
      $styleCellCache[$sheet.Name] = Get-ExcludedStyleCells -Zip $Zip -Sheet $sheet -StyleNameMap $styleNameMap
    }
    $excludedHere = $styleCellCache[$sheet.Name]

    $parsed = @{}
    foreach ($cell in $cells) {
      if ($excludedHere.Contains($cell.Addr)) { continue }   # "Unit no indent" / "Arguments hyperlink" - never a run member
      $sk = Get-FormulaSkeleton -FormulaText $cell.FormulaText
      if ($null -ne $sk) { $parsed[$cell.Addr] = $sk }
    }
    if ($parsed.Count -eq 0) { continue }

    foreach ($g in ($cells | Group-Object Row)) {
      $ordered = @($g.Group | Sort-Object Col)
      foreach ($run in (Find-SeriesRuns -OrderedCells $ordered -Parsed $parsed -AxisIsColumn $true)) {
        foreach ($iss in (Test-SeriesRun -RunCells $run -Parsed $parsed -HostSheet $sheet.Name -AxisIsColumn $true -Zip $Zip -SheetMap $SheetMap -ArgsSpillCache $argsSpillCache -StyleNameMap $styleNameMap -StyleCellCache $styleCellCache -NamedRangeAreas $namedRangeAreas)) { [void] $issues.Add($iss) }
      }
    }
    foreach ($g in ($cells | Group-Object Col)) {
      $ordered = @($g.Group | Sort-Object Row)
      foreach ($run in (Find-SeriesRuns -OrderedCells $ordered -Parsed $parsed -AxisIsColumn $false)) {
        foreach ($iss in (Test-SeriesRun -RunCells $run -Parsed $parsed -HostSheet $sheet.Name -AxisIsColumn $false -Zip $Zip -SheetMap $SheetMap -ArgsSpillCache $argsSpillCache -StyleNameMap $styleNameMap -StyleCellCache $styleCellCache -NamedRangeAreas $namedRangeAreas)) { [void] $issues.Add($iss) }
      }
    }
  }
  # NOTE: return the List<object> itself, NOT @($issues) - on this PowerShell
  # 5.1 build, wrapping a materialized Generic.List<T> VARIABLE directly with
  # the @() array-subexpression operator throws "Argument types do not match"
  # (ArgumentException), reproduced even for an EMPTY list, with or without
  # Set-StrictMode, in a fresh process - a genuine engine quirk on this build,
  # not something specific to this data. @() on a REGULAR array, or on a
  # pipeline/function-call result (which unrolls a List's contents into the
  # output stream one item at a time, same as every other check in this
  # script already does), is unaffected - only @() applied straight to a
  # List<T>-typed variable breaks. The caller already wraps this call with
  # @(Get-SeriesConsistencyIssues ...), which is the safe pattern.
  return $issues
}

# ---------------------------------------------------------------------------
# Report a capped list.
# ---------------------------------------------------------------------------
function Write-Capped {
  param([array] $Items, [int] $Max, [scriptblock] $Format)
  $shown = if ($Max -le 0) { $Items.Count } else { [Math]::Min($Max, $Items.Count) }
  for ($i = 0; $i -lt $shown; $i++) { Write-Host ('    ' + (& $Format $Items[$i])) }
  if ($shown -lt $Items.Count) { Write-Host ("    ... and {0} more" -f ($Items.Count - $shown)) }
}

# ===========================================================================
# Main.
# ===========================================================================
$workbooks = Get-TargetWorkbooks -RepoRoot $RepoRoot -WorkbookPath $WorkbookPath
if (@($workbooks).Count -eq 0) { Write-Host 'No workbooks found to scan.'; return }

Write-Host ("Scanning {0} workbook(s) for errors (read-only)..." -f @($workbooks).Count)

$grand = @{ cell = 0; name = 0; link = 0; sheet = 0; sum = 0; marked = 0; shift = 0; series = 0; hiddenCached = 0; wbWithIssues = 0 }

foreach ($path in $workbooks) {
  $leaf = Split-Path $path -Leaf
  $cellErrors = @(); $nameErrors = @(); $extLinks = @(); $sumMismatches = @(); $sumMarked = 0; $missingSheetRefs = @(); $numericOperandIssues = @(); $seriesIssues = @(); $fullCalcOnLoad = $false
  try {
    $zip = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
      $sheetMap   = @(Get-SheetMap -Zip $zip)
      $fullCalcOnLoad = Get-FullCalcOnLoad -Zip $zip
      $cellErrors = @(Get-CellErrors  -Zip $zip -SheetMap $sheetMap)
      $nameErrors = @(Get-NameErrors  -Zip $zip -SheetMap $sheetMap -IncludeLibraryFunctions:$IncludeLibraryFunctions)
      $extLinks   = @(Get-ExternalLinks -Zip $zip)
      $missingSheetRefs = @(Get-MissingSheetReferences -Zip $zip -SheetMap $sheetMap)
      if (-not $SkipSumCheck) {
        $sumResult     = Get-SumRangeMismatches -Zip $zip -SheetMap $sheetMap
        $sumMismatches = @($sumResult.Issues)
        $sumMarked     = $sumResult.Marked
      }
      if (-not $SkipNumericOperandCheck) {
        $numericOperandIssues = @(Get-NumericOperandIssues -Zip $zip -SheetMap $sheetMap)
      }
      if ($IncludeSeriesCheck) {
        $seriesIssues = @(Get-SeriesConsistencyIssues -Zip $zip -SheetMap $sheetMap)
      }
    } finally { $zip.Dispose() }
  } catch {
    Write-Host ''
    Write-Host ('=' * 78)
    Write-Host ("Workbook : {0}" -f $leaf)
    Write-Warning ("  Could not read workbook: {0}" -f $_.Exception.Message)
    continue
  }

  $grand.marked += $sumMarked
  # In a fullCalcOnLoad workbook cached cell errors are provisional (Excel recomputes
  # on open); hide them by default, still count them. Structural formula-text errors
  # (IsCached=$false) are always shown.
  $shownCellErrors = @(if ($fullCalcOnLoad -and -not $IncludeCachedErrors) { $cellErrors | Where-Object { -not $_.IsCached } } else { $cellErrors })
  $hiddenCached = @($cellErrors).Count - $shownCellErrors.Count
  $grand.hiddenCached += $hiddenCached
  $total = $shownCellErrors.Count + $nameErrors.Count + $extLinks.Count + $missingSheetRefs.Count + $sumMismatches.Count + $numericOperandIssues.Count + $seriesIssues.Count
  if ($total -eq 0) { continue }
  $grand.wbWithIssues++

  Write-Host ''
  Write-Host ('=' * 78)
  Write-Host ("Workbook : {0}" -f $leaf)

  if ($shownCellErrors.Count -gt 0) {
    $grand.cell += $shownCellErrors.Count
    Write-Host ("  Cell errors ({0}):" -f $shownCellErrors.Count) -ForegroundColor Red
    Write-Capped -Items $shownCellErrors -Max $Max -Format {
      param($e) "[{0,-8}] '{1}'!{2}  {3}" -f $e.Error, $e.Sheet, $e.Cell, $e.Formula
    }
  }
  if ($hiddenCached -gt 0) {
    Write-Host ("  {0} cached cell error(s) hidden (workbook recalculates on open); pass -IncludeCachedErrors to list." -f $hiddenCached) -ForegroundColor DarkGray
  }

  if ($nameErrors.Count -gt 0) {
    $grand.name += $nameErrors.Count
    Write-Host ("  Broken defined names ({0}):" -f $nameErrors.Count) -ForegroundColor DarkYellow
    Write-Capped -Items $nameErrors -Max $Max -Format {
      param($n) "[{0,-9}] {1} ({2})  {3}" -f $n.Category, $n.Name, $n.Scope, $n.RefersTo
    }
  }

  if ($extLinks.Count -gt 0) {
    $grand.link += $extLinks.Count
    Write-Host ("  External links ({0}):" -f $extLinks.Count) -ForegroundColor Yellow
    Write-Capped -Items $extLinks -Max $Max -Format { param($l) $l }
  }

  if ($missingSheetRefs.Count -gt 0) {
    $grand.sheet += $missingSheetRefs.Count
    Write-Host ("  References to missing sheets ({0}):" -f $missingSheetRefs.Count) -ForegroundColor Magenta
    Write-Capped -Items $missingSheetRefs -Max $Max -Format {
      param($s) "{0}  -> missing sheet(s): {1}   {2}" -f $s.Location, $s.Missing, $s.Formula
    }
  }

  if ($sumMismatches.Count -gt 0) {
    $grand.sum += $sumMismatches.Count
    Write-Host ("  SUM range size mismatches ({0}):" -f $sumMismatches.Count) -ForegroundColor Cyan
    Write-Capped -Items $sumMismatches -Max $Max -Format {
      param($s)
      $sheetTag = if ($s.RangeSheet -eq $s.Sheet) { '' } else { " [range on '$($s.RangeSheet)']" }
      "'{0}'!{1}  SUM range {2} but data spans {3}{4}   {5}" -f $s.Sheet, $s.Cell, $s.Stated, $s.Actual, $sheetTag, $s.Formula
    }
    if ($sumMarked -gt 0) {
      Write-Host ("    ({0} other SUM formula(s) marked intentional via N(`"Partial...`") - skipped)" -f $sumMarked) -ForegroundColor DarkGray
    }
    if ($fullCalcOnLoad) {
      # Unlike [cell], the [sum] check has no provisional/stale-cache awareness -
      # it reads cached values straight from the pre-recalc XML. A cell that is
      # really a text label (e.g. a shared formula rendering month names) can
      # show a stale cached error under fullCalcOnLoad and get miscounted as
      # numeric data, widening the detected block past a label column that was
      # never meant to be summed. Recalculating (Excel + Excel Labs add-in,
      # Ctrl+Alt+F9) and saving replaces every provisional cached value with a
      # real one, which resolves this at the source - a headless recalc must
      # NOT be used instead (no add-in loaded -> bakes #NAME? into every
      # Module.Func() cell, see build-scope3-excel.md).
      Write-Warning ("  '{0}' has not been recalculated since building (fullCalcOnLoad=1). SUM range mismatches above may be false positives from stale cached values - open in Excel with the Excel Labs add-in, force a full recalc, and save, then re-run find-errors." -f $leaf)
    }
  }

  if ($numericOperandIssues.Count -gt 0) {
    $grand.shift += $numericOperandIssues.Count
    Write-Host ("  Numeric-operand type mismatches ({0}):" -f $numericOperandIssues.Count) -ForegroundColor Magenta
    Write-Capped -Items $numericOperandIssues -Max $Max -Format {
      param($s)
      "'{0}'!{1}  operand {2} {3}   {4}" -f $s.Sheet, $s.Cell, $s.RefCell, $s.Detail, $s.Formula
    }
  }

  if ($seriesIssues.Count -gt 0) {
    $grand.series += $seriesIssues.Count
    Write-Host ("  Range-series reference issues ({0}):" -f $seriesIssues.Count) -ForegroundColor Blue
    Write-Capped -Items $seriesIssues -Max $Max -Format {
      param($s) "[{0,-6}] '{1}'!{2}  {3}   {4}" -f $s.Kind, $s.Sheet, $s.Cell, $s.Detail, $s.Formula
    }
  }
}

Write-Host ''
Write-Host ('=' * 78)
if ($grand.wbWithIssues -eq 0) {
  Write-Host ("CLEAN: no errors found in {0} workbook(s)." -f @($workbooks).Count) -ForegroundColor Green
} else {
  Write-Host ("TOTAL: {0} cell error(s), {1} broken name(s), {2} external link(s), {3} missing-sheet ref(s), {4} SUM range mismatch(es), {5} numeric-operand mismatch(es), {6} range-series issue(s) across {7} of {8} workbook(s)." -f `
    $grand.cell, $grand.name, $grand.link, $grand.sheet, $grand.sum, $grand.shift, $grand.series, $grand.wbWithIssues, @($workbooks).Count)
}
if ($grand.marked -gt 0) {
  Write-Host ("       ({0} SUM formula(s) marked intentional via N(`"Partial...`") across all scanned workbooks)" -f $grand.marked) -ForegroundColor DarkGray
}
if ($grand.hiddenCached -gt 0 -and -not $IncludeCachedErrors) {
  Write-Host ("       ({0} provisional cached cell error(s) hidden under fullCalcOnLoad; pass -IncludeCachedErrors to list)" -f $grand.hiddenCached) -ForegroundColor DarkGray
}
if (-not $IncludeSeriesCheck) {
  Write-Host "       ([series] range-follow check skipped by default; pass -IncludeSeriesCheck to run it)" -ForegroundColor DarkGray
}

if ($FailOnError -and ($grand.cell -gt 0 -or $grand.name -gt 0)) { exit 1 }
