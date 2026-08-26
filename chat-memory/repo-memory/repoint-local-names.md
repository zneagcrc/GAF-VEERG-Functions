# scripts/repoint-local-names.ps1 (npm `repoint-local-names[:dry] -- <path>`)

Fixes a Name Manager symptom hit while hand-copying cells/sheets between
enterprise TEMPLATE files (2026-08, building a new enterprise's template -
see enterprise-template-authoring.md for why that step is manual): after
copying, a defined name's "Refers to" shows something like
`='[10.1.1-2 Solid waste treatment]10.1'!$E$192` even after the user deletes
the visible source-file path text and retypes just the sheet reference. A
plain-cell reference with no period in the sheet name (`='15 Scope 3'!$E$159`)
was NOT affected - the user initially suspected periods in the sheet name were
the trigger, but that's not actually it (see ROOT CAUSE).

## ROOT CAUSE
Copying a sheet/cells across workbooks makes Excel externalise any reference
the copied content carries to a sheet that wasn't copied along (same
mechanism documented in build-enterprise-afe.md's sheet-copy notes, just
triggered by hand instead of by `build-enterprise-excel.ps1`). Hand-editing
the Name Manager "Refers to" box CANNOT fix this: Excel keeps re-validating
against the workbook's registered `xl/externalLinks/*` parts no matter what
text is typed, so on commit it re-derives a broken reference from its own
link table instead of accepting the clean typed text - exactly the class of
problem `external-links.ps1`'s header comment already documents ("COM cannot
do this ... only in XML").

## FIX = script, not Name Manager
`scripts/repoint-local-names.ps1`:
1. Opens the target `.xlsx` directly as a zip (never via COM/Excel, so
   nothing re-parses/re-mangles the text), reads `xl/workbook.xml`'s local
   sheet name set.
2. For every `<definedName>` whose RefersTo contains a bracketed qualifier,
   tries to resolve it against a LOCAL sheet in TWO forms (regex
   `'(?<pre>[^'\[\]]*)\[(?<mid>[^\[\]]*)\](?<post>[^']*)'!`):
   - standard form: text AFTER `]` (`post`) matches a local sheet -> keep that.
   - MANGLED form (the one the user hit): text INSIDE the brackets (`mid`)
     matches a local sheet, `post` is garbage/truncated -> use `mid` instead.
   Whichever resolves, the whole qualifier + `!` is replaced with
   `'<sheet>'!`, keeping the name (does NOT delete names, unlike
   `Remove-ExternalLinkArtifacts`, since the point is to keep them pointed
   locally).
3. Names that don't resolve to any local sheet are left untouched and
   reported (genuinely external, or the target sheet really isn't here).
4. Calls the existing `Remove-ExternalLinkArtifacts` (external-links.ps1) to
   purge whatever `xl/externalLinks/*`/`<externalReferences>`/rels/
   content-type plumbing is now orphaned, and to remove any name STILL
   carrying a real `[N]` (N>=1) index after step 2 (its own, already-proven
   rule - untouched here).
- `-DryRun` reports without writing; real run only proceeds if the file isn't
  locked (dot-sources `file-access.ps1`'s `Assert-FilesAccessible`, same
  pre-flight pattern as the builders - fails fast with a clear "close it in
  Excel" message instead of a raw IOException).

## Verified (scratch file, not a real repo workbook)
Built a throwaway xlsx with 4 defined names covering: the exact mangled form
(`'[10.1.1-2 Solid waste treatment]10.1'!`), the standard externalised form
(`'[SomeOtherBook.xlsx]10.1.1-2 Solid waste treatment'!`), a genuinely
external one (target sheet not present locally), and an already-clean local
name. Dry-run and real run both correctly repointed the first two, correctly
left the genuinely-external one alone (and it survives the
Remove-ExternalLinkArtifacts pass too, since `[SomeOtherBook.xlsx]` isn't the
digits-only `[N]` form that function targets - matches its documented scope),
left the already-clean one untouched, and the saved `workbook.xml` re-parses
cleanly afterward.

GOTCHA hit building the scratch-file harness (not in the real script): don't
build injected XML text via PowerShell's `-replace` operator when the
replacement string itself contains literal `$digit`-shaped substrings (e.g.
`$A$1` inside a cell reference) - `-replace`'s replacement argument uses .NET
regex substitution syntax, so `$1` gets interpreted as "insert capture group
1" even when it's just literal text you're injecting, silently corrupting the
XML. Use plain `.Replace()` (or a MatchEvaluator) for literal injection
instead.
