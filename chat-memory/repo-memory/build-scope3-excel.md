# build-scope3-excel.ps1 (Scope 3 workbook builder)

## CITATION-[N] FALSE DELETE bug in Remove-ExternalLinkArtifacts (2026-08)
- SYMPTOM: generated Scope 3 shows genuine `#NAME?` on cells calling AFE metadata
  LAMBDAs, e.g. `Constants - Swine!D20 =SourceData_Swine.SourceData_Swine_MethaneVolatileSolidsEmissionsPotential_Source()`
  and `Constants - Dairy!D20 =SourceData_Dairy.SourceData_Dairy_LiveweightCowsAndHeifers_Variation()`.
  The AFE MODULE + function are present in the output (verified: 59 modules,
  hasFunction=True) - but the PUBLISHED workbook-scoped defined name is MISSING
  (template had 58 `SourceData_Swine.*` names, output had 57; the one dropped =
  the failing function). RefersTo was `_xlfn.LAMBDA("IPCC (2019), Chapter 10 [4]")`.
- ROOT CAUSE: `Remove-ExternalLinkArtifacts` deleted every `<definedName>` whose
  InnerText matched `\[\d+\]` (meant for external-book refs like `='[1]Sheet'!$A$1`).
  A `_Source`/`_Variation` citation LAMBDA whose STRING body contains a bracketed
  reference number (`[4]`, IPCC citations) matched the same pattern -> the legit AFE
  function name was wrongly deleted -> #NAME?. Explains multiple cell errors (cleanup
  removed ~208 names/run; some were legit citation LAMBDAs).
- FIX: strip double-quoted string literals BEFORE the `\[\d+\]` test
  (`$reStringLiteral = [regex] '"[^"]*"'`; test `$reStringLiteral.Replace($n.InnerText,'')`).
  External `[N]` book indexes live OUTSIDE double quotes (references / single-quoted
  sheet names), so stripping double-quoted constants keeps genuine external-ref
  deletion working while sparing citation LAMBDAs.
- DIAGNOSE with a read-only AFE probe (zip customXml AFEJSONBlob = base64 UTF-16 JSON,
  `.files[].path`/`.text`) + workbook.xml `<definedName>` compare across output vs
  template vs source; NO COM/recalc needed. Do NOT blanket-suppress AFE `Module.Func()`
  #NAME? in find-errors (tried + reverted; it masked THIS real bug).

- RUN: `npm run build-scope3` (added to package.json) = `powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-scope3-excel.ps1`. No DryRun variant.
- KEY FILES: output `Excel\13_Scope3_WIP_v13.xlsx`; template `Excel\13_Scope3_Template_WIP_v13.xlsx`; source module workbooks under `Excel\` (e.g. `10_SolidWaste_WIP_v02.xlsx`, `14_Electricity_Scope2_WIP_v02.xlsx`). Template is copied fresh at start, then module sheets imported on top (deterministic full rebuild).
- `$ImportMap` (~line 62): 17 sheets (16 replace + 1 add). Source file per sheet resolved by version-agnostic prefix (highest `_v<NN>`), same pattern as enterprise `Resolve-SourceWorkbook`.

## Verified working (do NOT re-investigate)
- IMPORT IS CORRECT. Sheets like `Input - Solid waste` ARE imported faithfully: formulas identical (COM diff 0), cell text identical (shared strings resolve to same text despite different SST indices), data-validation dropdowns identical. If a user reports "sheet not imported / F5 missing", FIRST check they are opening the freshly-built output (timestamp) not a stale copy, and that their edit is in the resolved source version. Confirmed a user F5 edit in source DID appear in output after rebuild.
- Comparison pitfall: naive regex `<c r="A6"[^>]*>.*?</c>` mis-captures because self-closing cells `<c r="A6" s="12"/>` make `.*?` run into the NEXT cell. Also SST indices differ between files (output merges the shared-string table). Compare RESOLVED TEXT, not raw `<v>` indices.

## Two fixes implemented this cycle
1. PHANTOM EXTERNAL LINKS: 190 orphan wb-scoped `M1_Table_*` defined names carried `[N]`-prefixed refs (e.g. `='[1]4.3.1'!$N$212:$R$213`) to non-imported calc sheets, keeping ~32 `xl/externalLinks/externalLinkNN.xml` parts alive. ZERO cell formulas use `[N]`. COM CANNOT remove: `BreakLink` throws on `xlPathMissing`; `.RefersTo` throws on unresolvable `[N]` names (only caught 18/190). FIX = `Remove-ExternalLinkArtifacts -TargetPath` (XML-only zip edit, runs after `Repair-BrokenWorkbookNames`): strip every workbook.xml `<definedName>` whose InnerText matches `\[\d+\]`, remove empty `<definedNames>`, remove `<externalReferences>`, remove external-link `<Relationship>` in `xl/_rels/workbook.xml.rels`, delete all `xl/externalLinks/*`, remove `[Content_Types].xml` overrides for those parts. Last run: 208 names + 64 parts removed. Verified: LinkSources=(none), 0 externalReference, 0 `[N]`.
2. `Input - Electricity!E16` needed manual F9 (stale `#REF!` cached because XML name-repair runs after last save). FIX = `Invoke-FinalRecalcAndLinkCleanup -TargetPath` sets `<calcPr fullCalcOnLoad="1"/>` in workbook.xml (XML-only). User's Excel (with Excel Labs add-in) recalcs on open.

## Critical lesson: DO NOT headless-recalc this workbook
- Headless COM (`New-Object -ComObject Excel.Application`) has NO Excel Labs (AFE) add-in loaded, so `CalculateFullRebuild()` evaluates module-namespaced funcs (`SourceData_Swine.<fn>()`, `SourceData_Dairy.<fn>()`) to `#NAME?` and BAKES that into cached values (broke `Constants - Swine!D20`, `Constants - Dairy!D20`) and can RE-CREATE an external link. Use `fullCalcOnLoad` flag instead — never a headless rebuild. Good D20 cached values are plain strings (t="str"), e.g. "IPCC (2019), Chapter 10 [4]".
- Therefore VERIFY with a NO-RECALC probe: open COM readonly WITHOUT recalc, or read cached `<v>` straight from sheet XML. A verify script that calls `CalculateFullRebuild` gives FALSE `#NAME?` on the two D20 cells.
- Setting `$x.Calculation = xlCalculationManual (-4135)` BEFORE any workbook is open throws 0x800A03EC; set it after Open (or just read cached values from XML).

## Known pre-existing (not part of the two fixes)
- `15 Scope 3!H596` freight `#CALC!` = incomplete LAMBDA template formula (called with no args). Scope-inherent, reported separately.

## Later additions (see build-shared-modules.md)
- Nav menu: dot-sources nav-menu.ps1; before `$target.Save()` seeds `Get-InferredCategoryMap` then overlays `$ImportMap` categories, calls `Set-NavMenu -Target $target -CategoryMap $menuCategory -Labels @{}`. `Overview` sits in untitled group (Home,Overview,Results); template calc sheets `15 Scope 3` + `Calcs - manure sent off-site` classify as calculation.
- Zoom: dot-sources worksheet-view.ps1; `Set-WorkbookZoom -Path $OutputPath -Zoom 100` after the final XML passes, guarded `if (-not $DryRun)`.
