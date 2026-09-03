# build-enterprise-excel.ps1 — VEERG_*_Result_Method* names silently dropped

## Symptom (2026-09)
A built enterprise is missing some workbook-scoped `VEERG_<eq>_Result_Method<N>`
defined names. The custom Results sheets read those names (`=VEERG_..._Result_Method1`),
so the affected results show **0** (or `#NAME?`). No error, no warning.

Observed on `Enterprise_CroppingCotton` AND `Enterprise_CroppingGrains` (identical
2 names missing), so it is **not enterprise-specific** — every enterprise built
before the fix is affected. In the cotton case the two lost names were:
- `VEERG_5_5_1_1__1_FertiliserLeachingAndRunoffN2OEmissions_Result_Method1` -> `'5.5.1.1-2 Leaching and runoff'!$E$94`
- `VEERG_5_4_1_3__1_OrganicFertiliserAtmosphericDepositionN2OEmissions_Result_Method1` -> `'5.4.1.1-4 Atmos deposition'!$E$146`

## Why the cell's value seemed to matter
A user found: making the source result cell evaluate to a **non-zero number**
kept the name; leaving it **zero** dropped it. Actual trigger is narrower:
a result cell that evaluates to **`#VALUE!`** in the add-in-less headless build
gets its result name externalised (see below). Cells that are `0` or `#N/A`
do not. (`#VALUE!` here is a "cached error hidden under fullCalcOnLoad" that
recomputes fine when the user opens the file with the Excel Labs add-in — but by
then the name is already gone.)

## Root cause (the deletion chain)
1. A custom sheet (`Results - Activity period` / `Results - Product EI`) has
   `=VEERG_X_Result_Method1` and is present in the template BEFORE the owning
   module calc sheet is imported. Excel resolves the name by **externalising**
   it: the workbook-scoped `<definedName>` becomes
   `[N]'5.5.1.1-2 Leaching and runoff'!$E$94` (or `[5_Fertiliser_WIP_v10.xlsx]'…'!…`).
2. COM `Name.RefersTo` **throws** on such an unresolvable `[N]` name, so the
   name-upsert's re-point guard (`$curRefers = try{ $existing.RefersTo }catch{''}`
   then `if ($curRefers -match '\[' …)`) never fires — `$curRefers` is `''`.
   Excel also *re-externalises* the name if you set `RefersTo` while the source
   workbook is still open, so a COM-phase repair does not hold.
3. `Remove-RedundantSheetScopedNames` deletes the sheet-scoped shadow
   (`'Sheet'!VEERG_…`) because a workbook-scoped counterpart exists.
4. `Remove-ExternalLinkArtifacts` deletes the workbook-scoped copy because its
   RefersTo matches `\[[1-9]\d*\]`.
   -> nothing left. `=Name` reads 0.

## Fix: Restore-SourceWorkbookScopedNames (final XML pass)
New function in `build-enterprise-excel.ps1`, called LAST (after
`Remove-ExternalLinkArtifacts`, before zoom/fullCalcOnLoad), XML-only:
- read each source module workbook's **workbook-scoped** `<definedName>`s matching
  `(?i)^VEERG_.*_Result_Method\d+$`, keep those whose RefersTo is a plain
  single-sheet ref (`'Sheet'!$cell`, no `[`, no `#REF`)
- keep only names whose referenced sheet is present in the enterprise
- in the enterprise `xl/workbook.xml`: add the name if absent; replace it if the
  current definition is externalised (`[..]`) / `#REF!` / wrong; leave an
  already-correct local definition alone (may be a deliberate template override)
- first source workbook to supply a locally-valid definition wins (import order —
  matches the "canonical module first" rule for duplicate input sheets). This is
  why Scope3's `'[6]Calcs - Manure sent off-site'!$E$90` copy of the same name is
  ignored: that sheet is not imported into a cropping enterprise.
- summary line: `Result names restored : N`

Verified on cotton: `Result names restored : 2`, both names back as clean
workbook-scoped local refs, `find-errors` clean. **Rebuild every enterprise** to
recover their dropped result names.

## PowerShell gotcha hit while writing it
`[regex] "…\$?[A-Za-z]…"` in a DOUBLE-quoted string: `$?` is the automatic
success variable, so the pattern becomes `…\True?[A-Za-z]…` -> `[regex]` throws
`Unrecognized escape sequence \T` at runtime -> the build aborts mid-way (~7 min
in) leaving a partial output. Use SINGLE-quoted regex strings, or `` `$ ``.
