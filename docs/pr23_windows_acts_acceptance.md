# PR23 Windows acceptance — 25.10.2026

## Result

PASS for first-passport VIK/UZK variable-part automation, native Office
generation, isolated transaction/recovery and separate RC3 test delivery.
Official passport11/2026 NOT issued. A fresh official GUI release remains
blocked by absent trustworthy current control facts and form approval; it is
not reported as a full GUI issuance PASS.

Delivery: `C:\VZMK-Passport\1.0.0-RC3-acts-test\VZMK-Passport-1.0.0.exe`.
The onedir bundle includes Python/runtime dependencies. Installed Word and
Excel remain required. No previous EXE or desktop shortcut was replaced.

## Integration and preserved behavior

Latest local RC3-updated source was clean at
`6cda293a55d102743499725f4a6116664f212be8`, matching its delivery source pin.
New worktree: `passport-rc3-acts\Pasport`, branch `codex/rc3-acts-test`.
Verified checkpoint: `C:\VZMKPassportChecks\RC3_ACTS_20261025\BEFORE_PR23.bundle`.
Applied the complete upstream range from `f9fb2ed` to exact
`f1753a362393e77c2b04c2b3d0b92e89c1da29b1` with three-way application.
Later local dates, PDF, mass precision, shared-UZK races, preview consent and
registry preservation/recovery behavior were retained.

The GUI has no layout/button changes. Its only source change is a separate
`rc3-acts-settings.json` filename, preventing previous versions' profile/form
paths from selecting the old templates. Production initial paths were migrated;
all training/UI tests used private LOCALAPPDATA directories.

## Source examination and geometry

The original first VIK/UZK were byte-copied for examination; source files were
read without modification. Word opened the separate exact copies ReadOnly.
All four original pages were exported/rendered and reviewed.

- VIK: two acts, portrait A4, original tables/fonts/technical wording retained.
- UZK: landscape A4, original eight data columns, company/requisite/equipment
  blocks and cell widths retained.
- Cleared historic marks, dates, welders, inspectors, defect assessments,
  conclusions and signatory identities. Source DOCX had no media/signature
  images. New signable spaces remain blank.
- Adaptation changes only document.xml; other ZIP parts and relationships are
  byte-identical before native Word saves generated output. Page geometry and
  source table/cell/paragraph/run properties are preserved. The detailed slot
  contract and reference inventory are in evidence `artifact.md` and
  `ВИК_STRUCTURE.json` / `УЗК_STRUCTURE.json`.
- Constant welding method/type/regulation and instrument/calibration requisites
  remain form data. Their current applicability still requires responsible
  approval; historic personnel and favorable outcomes are not constants.

The accepted Stage4 workbook was found at the retained source project's
`DEV\Справочники (2)\Профили.xlsx`. It was examined and explicitly promoted to
the new internal Templates directory, not used as a silent DEV fallback.
Workbook, user-state and catalog hashes match profiles.build.json. The original
user decision attachment hash also matches the retained approval record.
Internal workbook SHA:
`7b420f6dead96408680a64d220c83e8980aeec1c38635e35b2e4723a98e3831b`.

Row683:16П, ГОСТ8240-97, h160/b64/s5/t8.4; confirmed rule ШВЕЛЛЕР_16П_1.
Its dimensions have two recorded source provenances. For 8 selected items:

| Part | Per item mm | Selected length mm | Thickness mm | MHz | Angle |
|---|---:|---:|---:|---:|---:|
| Web | 160 | 1280 | 5 | 5 | 70 |
| Flanges | 2×64 | 1024 | 8.4 | 5 | 70 |

No different marks are summed. Repeated VOM rows do not multiply the already
calculated shared-UZK quantity twice. Geometry is not an actual inspection,
defect verdict, thickness-method qualification, welder identity or signature.

## Real four-mark example

Real VOM read directly: 4 marks,11 items,511.98kg,19.533m2.
VIK contains Бп2-1/Бп2-5/Бп2-18/Бп2-19 in current order.
Shared read-only UZK gives Бп2-1:8; other three:0. Only the positively selected
mark is printed in UZK, with its independent web/flange characteristics.

Read-only registry calculation on a disposable copy gives next11/2026,
start05.08.2026, proposed ending25.10.2026, VIK18.10.2026, UZK20.10.2026,
AKZ24.10.2026. These are registry-derived dates/proposals, not proof that the
current inspections were performed.

Editable real-data drafts: VIK.docx (2 pages), UZK.docx (1 page), combined draft
PDF (3 pages) in the delivery's `Реальный пример — ЧЕРНОВИК` folder.
Missing actual outcomes are explicitly unconfirmed; no historical positive
verdict, personnel or signature is inserted. All final pages were reviewed.

## Defects corrected

1. Clearing historical personnel shortened VIK and orphaned the second act's
   heading at the end of page1. Added an explicit independent page start and
   functional regression. Final: two complete acts on two pages.
2. The source second-number SDT is PASSPORT_NO but its value includes -1.
   The safe Word binder correctly rejected that mixed historical-field shape.
   Adaptation now tags the second number VIK_PROCESS_NUMBER and binds the
   derived current number-1, preserving the original SDT/run formatting.
   Historic-field checks were not weakened. Initial failed CHECK is retained.
3. The source UZK mixed metadata/data in one table and split a row over pages.
   A minimal separator keeps the two source-derived tables independent; only
   the eight-column header repeats and body rows cannot split. Widths/styles
   remain original.18 synthetic independent marks produced3 reviewed pages.
4. Added geometry SHA comparison between calculation and final input snapshot;
   concurrent changes cannot be silently adopted. Current frequency/angle
   discrepancies are rejected rather than overwritten by design values.
5. Namespace-preserving OOXML patching uses packaged lxml, retaining original
   namespace declarations, opaque parts and layout instead of rebuilding DOCX.

## Regression and native acceptance

Final single invocation: **409 PASS,0 FAIL,0 ERROR,0 SKIP**.
The same accepted production/stage6/stage9/office-unbound runner collects the
383 previous tests plus13 upstream and13 native-act tests. Historic stage5
evidence-dependent scripts are explicitly listed outside the runner manifest.

Earlier diagnostic results remain in evidence. The initial regression had
2 stale mass-test mocks lacking VOM rows; a later new SDT test fixture included
the heading inside a number-only control and failed its assertion. Fixtures
were corrected to the real contracts; mass/number/date checks remain active.
The final regression ran on the exact runtime code subsequently packaged.

Both source and installed EXE workers completed:
CHECK→PREPARE→VERIFY→CONFIRM→RELEASE→COMMIT→READBACK→COMPLETE.
Explicit synthetic facts/protocols and a program-created registry copy only.
Three DOCX and three XLSX were opened/verified by native Word/Excel.

| Document | Common PDF pages |
|---|---|
| Quality | 1–2 |
| Appendix No.1 | 3 |
| VIK, two acts | 4–5 |
| UZK, source eight-column layout | 6 |
| AKZ0, one coat | 7 |
| AKZ appendix0 | 8 |

All8 pages were visually reviewed. The EXE preview page images match the
reviewed source preview exactly; the EXE final PDF page images match its
preview exactly. PDF filename equals its passport folder and sits beside
editable files. Training climate remains25/21/55/dew15.3/dry110 per agreed
23..27/20..22/50..60/95..125 synthetic ranges. Signatures remain blank.
The earlier 104-mark production print engines were not replaced with old code.

Both declined-consent combinations refuse confirmation with unchanged copied
registry. Positive confirmation was supplied only for explicitly synthetic
training after visual review; not on behalf of a human official release.

Actual native EXE GUI: selected Garage and new11/2026; loaded the real desktop
VOM, displayed4/11/511.98/19.533; Facts showed the same four columns, [16П,
AKZ0,8/0/0/0 and Back. No new geometry button/table. Back with empty ranges
returned without registry write. Screenshots retained in `gui_evidence`.
Full training transaction used the existing EXE worker entry points; this
does not claim a fresh GUI-to-official-issue flow has current factual sources.

## Refusals, recovery and preservation

- ACTUAL-mode request with no invented verdict correctly refuses:
  `Нет подтверждённого результата ВИК и первичного протокола: Бп2-1`.
- Changing only a separate geometry copy after CHECK blocks PREPARE; installed
  geometry and copied registry remain unchanged.
- Controlled exception immediately after actual Windows ReplaceFile:
  recovery refuses an injected parallel edit without changing it; restoring
  that known private test edit permits exact-byte rollback. A second recovery
  is idempotent. Unconfirmed target documents are not left published.
- All Office supervisors report zero owned orphans/ownership errors and no
  forced termination. No foreign Word/Excel process was terminated.
- **11142 protected files unchanged**, including original registry, UZK,
  VOM/documents, accepted Stage4 inputs/approval and old deliveries.
-50 packaged Python modules match tested source files exactly.
- No main merge, push, previous-delivery replacement or new desktop shortcut.

## Remaining conditions for official release

Current VIK outcomes/protocols for the four marks, actual UZK results and
participants/protocol for the8 selected items, current AKZ facts and required
quality requisites must be trustworthy. Geometry/planned quantity cannot
confirm these. Applicability of the current control method/equipment to the
5mm web and8.4mm flanges, metrology and adapted form approval is not established
by a geometry calculation. The agreed simplified GUI still lacks a source for
part of these required primary facts; no extra controls were added against the
user's instruction. Official issuance remains NO_GO until those conditions
are met. The separate PR23 structural/test build is ready for inspection.

Evidence root: `C:\VZMKPassportChecks\RC3_ACTS_20261025`.
