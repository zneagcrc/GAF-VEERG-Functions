# 'VEERG documents/2. Chapter_2_Implementation_Guides_and_Data_Checklists.docx'

Authoritative source for WHICH chapters/modules apply to WHICH enterprise
type when drafting a new enterprise's `modules[]`.

## USE THE MODULE MAP TABLE, NOT THE PROSE SENTENCE (corrected 2026-08)
Each enterprise section has BOTH a structured "Module Map" TABLE
(`Emission Activity/Source -> Relevant Module`, grouped by Scope 1/2/3) AND a
prose sentence later ("Additional information and guidance on specific
parameters can be found within Chapter 3 – Section; 3.<N> ... With additional
related parameters found within Chapter <X> – Sections; ..."). The FIRST pass
at this (drafting Enterprise_Swine.json) used the prose sentence and got it
WRONG in two ways, caught only because the user knows the domain and pushed
back: the sentence claimed Chapter 6 (Ag Residue) applied to swine (it
doesn't - confirmed absent from the actual table) and OMITTED Chapter 8
(Fuel) entirely even though the table clearly lists it (Transport Fuel 8.1,
Stationary Fuel 8.2). The prose sentence is a loose, hand-written summary and
is NOT reliable - always find and read the Module Map TABLE for the
enterprise in question; treat the prose sentence as, at best, a secondary
cross-check, not the source.

## How to read it (no Word/COM needed)
It's a zip of XML like xlsx - same read-only zip-only approach used
throughout this repo's Excel tooling, pointed at a docx instead. To get the
TABLE structure (not just paragraphs), preserve row/cell boundaries as well
as paragraph breaks before stripping tags, or the table collapses into
unreadable run-on text:
```powershell
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$fs = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
$za = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
$xml = (New-Object System.IO.StreamReader(($za.Entries | ? FullName -eq 'word/document.xml').Open())).ReadToEnd()
$xml = $xml -replace '</w:tr>', "`n"    # table row break
$xml = $xml -replace '</w:tc>', "`t"    # table cell break
$xml = $xml -replace '</w:p>', "`n"     # paragraph break
$text = [regex]::Replace($xml, '<[^>]+>', '') -replace '&amp;','&' -replace '&lt;','<' -replace '&gt;','>' -replace '&quot;','"' -replace '&apos;',"'"
```
Then find the enterprise's own "Module Map" heading: search for the
enterprise name as a standalone heading line (e.g. a line that's just
"Swine") followed shortly by a line "Module Map" - there's one per
enterprise section, plus other unrelated "Module Map" headings for every
OTHER enterprise in the same doc, so anchor on the enterprise name heading
first, not just "Module Map" alone (that string alone hits ~13+ times, one
per production system).

## Confirmed Module Map tables (from the actual table, not the prose)
- Swine (2026-08): Scope 1 = Enteric Methane 3.5, Manure Management 4.5,
  Transport Fuel 8.1, Stationary Fuel 8.2, Solid Waste Treatment 10.1.
  Scope 2 = Purchased Electricity 14.1. Scope 3 = Purchased Livestock 15.1,
  Purchased Feed 15.2, Purchased Services/Contractors 15.7, Other Purchased
  Goods and Services 15.10, Upstream Emission from Fuel 15.11, Upstream
  Emissions of Purchased Electricity 15.12, Management of Waste 15.13.
  NO Fertiliser (Ch 5), NO Agricultural Residue (Ch 6), NO Freight (15.14)
  anywhere in the table - all three were things earlier reasoning (mine, off
  the prose sentence) got wrong or left as open questions for Swine.
  `Enterprise_Swine.json`'s Scope3 include still lists `Input - Freight` as
  an OPEN QUESTION comment (not yet resolved either way) since the table
  gives no positive evidence for or against it beyond simply not listing
  15.14 as a distinct row.

- Beef pasture/range/paddock and Dairy: NOT yet re-verified against their
  own Module Map tables (only checked via the prose sentence, which is now
  known to be unreliable) - they happen to match their existing working
  configs when read via the prose sentence, which was the original basis
  for trusting the prose sentence at all, but given the Swine mismatches,
  worth a real table-based re-check before treating them as fully confirmed
  reference examples for a future enterprise.

## Gotchas found drafting Enterprise_Swine.json (2026-08, still valid)
User hand-authored a first draft by adapting PastureBeef's/Dairy's JSON and
introduced real bugs this way, worth checking for on any new enterprise
drafted the same way:
- `include: { "3.5.1.1-2 Enteric Methane": [...] }` - used the SHEET NAME as
  the include key instead of `"calculation"`. Structurally wrong; the
  builder's `Add-SheetToPlan` loop only recognises `input`/`calculation`/
  `constants` as include keys, so this key is silently ignored and the
  category falls back to the module's FULL registry default instead of the
  intended subset.
- Leftover copy-paste text: a comment mentioning "Pasture Beef herd
  movement" and a menu label `"Constants - Pasture Beef": "Constants -
  Pasture Beef"` that should have been `"Constants - Swine"`.
- Orphaned labels for chapters that aren't actually included (e.g.
  `"5.1.1.3-4 Inorganic fert CO2"` left over after Fertiliser was correctly
  excluded) - harmless (an unused label just never matches a sheet) but
  worth cleaning up for clarity.
None of these break silently in an obviously loud way (no error, no
warning) - only a careful read-through (or comparing `list-enterprise-
results.ps1` output against the Module Map TABLE) catches them. Worth a
deliberate check on every new enterprise draft.
