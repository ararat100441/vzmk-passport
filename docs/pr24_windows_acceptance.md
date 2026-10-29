# PR24 Windows acceptance — 26.10.2026

## Result and boundaries

PASS: exact marked first-passport VIK/UZK adaptation, real-data draft
generation, full isolated Office transaction, recovery, regression and separate
Windows EXE delivery. Official passport 11/2026 was NOT issued.
Fresh official issuance through the simple GUI remains BLOCKED by missing
trustworthy current control facts; it is not counted as a full GUI issue PASS.

EXE: `C:\VZMK-Passport\1.0.0-RC3-marked-forms\VZMK-Passport-1.0.0.exe`.
Onedir distribution includes Python/runtime dependencies; installed Microsoft
Word and Excel remain necessary. No old EXE or shortcut was replaced.

## Integration and checkpoint

The latest local RC3-acts-test source was clean at
`0b7d8ff93e49e71f16c4978a6e0e701d43e7d13d`, matching its installed source pin.
Checkpoint: `C:\VZMKPassportChecks\RC3_MARKED_20261025\BEFORE_PR24.bundle`,
verified before changes. Evidence uses the start date 25 October; final
generation/acceptance occurred after midnight on 26 October, Europe/Samara.

Fetched exact upstream `9b490f3bfdc9d13e54c2032345997a611aeeab55` from
`feature/exact-highlighted-vik-uzk`. Its delta was applied with three-way
integration in new worktree `passport-rc3-marked\Pasport`, local branch
`codex/rc3-marked-forms`. Conflicts retained all later local legacy/native
renderers, namespace preservation and lxml dependency. Geometry SHA race
checks, shared UZK, VOM parser, dates, PDF ordering and transaction protections
remain. No main change, PR merge, push or new PR.

No GUI layout, button or field changes. Only its settings filename is isolated
as `rc3-marked-settings.json`; accepted production paths were migrated to the
new delivery. All UI and transaction tests used private profiles/copies.

## Current original forms and exact changes

The two original DOCX had already changed since PR23. Their fresh entry hashes
are recorded separately in PREEXISTING_CHANGES.json, not reported as changes
made by this task. Original paths are under
`Работа\Документы\Гараж казахстан\Паспорт`:

| Original | SHA-256 at entry and final verification |
|---|---|
| 1_2026 - ВИК.docx | 3b040590582324bacd3ccf565f61b7a7950f1afca313c8deb12addaa9bd87028 |
| 1_2026 - УЗК.docx | c0e254c837bc32a242083a0920f99d05ac8d8abc6e21059381181f0023a57af4 |

Originals were read only; exact separate copies were opened ReadOnly in native
Word. All four original pages were exported/rendered and visually reviewed.

- VIK retains the original two acts, tables, prose, fonts, names, certificates
  and signature spaces. Existing number/date controls are updated. The yellow
  act-2 product phrase becomes unique item name `Балка`, never a list of marks.
  Number pair: 11/2026 and 11/2026-1.
- UZK retains the original combined requisite/data table and its eight visible
  columns (nine logical cells, characteristic cells merged). No PR23 table
  reconstruction/splitting is applied in this exact-source path. One positively
  selected mark yields one row: Бп2-1. Its one characteristics cell contains
  `L=1280х5 мм` and `L=1024х8,4 мм`; frequency 5/5 MHz, angle 70/70 degrees.
  The selected quantity is 8, not eight duplicated rows. Old mark hyperlinks
  to unrelated drawings are removed from XML and unused relationships.
- Only replaced yellow slots are cleared. Unchanged neighboring highlighting
  is retained. Other personnel/requisite/ДНО/годен text remains VERBATIM as
  explicitly required by PR24, with unverified-current-control provenance.
  It is not promoted to evidence of a new inspection. Draft/training banners
  and existing current-fact/official approval guards remain.
- Before native Word saves new outputs, all untouched DOCX ZIP parts are
  byte-identical; only document.xml and, for UZK, its relationships are edited.

## Real-data drafts

Direct reading of the current real `ВОМ отгрузочный счет.xlsx`:
4 marks (Бп2-1, Бп2-5, Бп2-18, Бп2-19), 11 items, 511.98 kg, 19.533 m².
This current source is the four-mark example, not the earlier 104-mark file.
Read-only shared UZK: Бп2-1 = 8; the other three = 0.
Accepted Stage4 geometry `[16П`, row683: 1280×5 and 1024×8.4 mm for 8 items;
different marks are not summed. Geometry is not an actual control verdict.

Native Excel calculation on a disposable copy of the real registry gives
next 11/2026; proposed end 26.10.2026, VIK 19.10.2026, UZK 21.10.2026,
AKZ 25.10.2026, first-passport start 05.08.2026. No new row was written in the
real registry. These calculated dates do not prove an inspection occurred.

Editable real-data VIK.docx (2 pages) and UZK.docx (1 page) plus combined
3-page PDF are in the delivery's `Реальный пример — ЧЕРНОВИК` subfolder.
Every page was visually checked against the exact source: no cut text,
overlapping rows or empty pages, correct two act headings, Балка, one UZK row,
both L×S values, dates/frequency/angle and preserved unchanged requisites.
The real draft PDF contains only VIK followed by UZK; missing current facts
for a complete official package were not fabricated.

## Corrections and regression

Local compatibility/format fixes atop PR24:

1. Preserve opaque source ZIP parts instead of rewriting them through the
   python-docx package serializer; edit only explicit XML parts.
2. Retain a second-number -1 suffix whether inside its SDT or in the adjacent
   source run. Other numbering series opt out to the existing generator.
3. Clear yellow on touched runs only, including number/date fields. The real
   UZK also stored yellow in paragraph run properties: five inherited editing
   highlights remained after cell replacement. Scoped clearing now removes
   them only in changed cells. A regression covers this actual source shape.
4. Reject nonpositive geometry/frequency/angle and out-of-range frequency/angle.

Final full accepted Windows runner, one invocation on packaged runtime code:
**427 PASS, 0 FAIL, 0 ERROR, 0 SKIP** (409 previous + 18 marked-form tests).
Manifest covers production, stage6, stage9 and office-unbound suites. Four
historical Stage5 evidence-dependent scripts are listed outside this runner;
they are not silently claimed PASS. A bundled artifact Python attempt could
not import pythoncom; the draft Office run was rerun with the existing native
Python 3.12 COM environment. Diagnostic log is retained. No production guard
was weakened and no failing final test was omitted.

## Native source and installed EXE acceptance

Both source and actually installed EXE completed, on distinct program-created
training registry copies and explicitly synthetic evidence:
CHECK → PREPARE → VERIFY → CONFIRM → RELEASE → COMMIT → READBACK → COMPLETE.
Original registry remains unchanged through preparation/confirmation/release;
backup matches the pre-commit SHA. Exactly one new 11/2026 row is read back in
the training registry (row432), journal COMPLETE, first start retained.

Three DOCX and three XLSX were saved, closed and opened/verified by native
Microsoft Office. Synthetic dates intentionally derive from their own copied
registry, and are 26.10.2026; real drafts retain the distinct dates above.

| Common PDF pages | Document |
|---|---|
| 1–2 | Quality document |
| 3 | Appendix No.1 |
| 4–5 | Both VIK acts |
| 6 | UZK |
| 7 | One-layer AKZ variant0 |
| 8 | Corresponding AKZ appendix |

Every source preview page was rendered and visually reviewed. Every installed
EXE preview/final page is pixel-identical to that reviewed page, 8/8. No cut,
blank or overlapping pages; seven Excel columns visible; totals 11/511.98/
19.533 preserved numerically. Final PDF is beside editable files, filename
equals its passport folder. No physical printing.

EXE GUI was launched in a separate LOCALAPPDATA with an untouched registry
copy. Verified Garage selection, automatic first-source paths, new11/2026,
real VOM 4/11/511.98/19.533, four-column facts table, q8/0/0/0, auto project
catalog and AKZ0 one layer, existing Back behavior, no new fields/buttons.
Missing actual climate caused the existing message `Требуется число: air от`
before any issue. Window closed normally. No current facts were invented to
make the fresh official GUI flow succeed. Full EXE engine acceptance above
uses its normal task worker interface, not a claim of fresh GUI official issue.

## Refusal, recovery and preservation

- Installed EXE rejects an actual request lacking current VIK proof with
  `Нет подтверждённого результата ВИК и первичного протокола: Бп2-1`;
  private registry SHA unchanged.
- Missing PDF review and declined approval both refuse confirmation without
  writing the registry, in both source and installed EXE training cases.
- Private geometry edited after CHECK is rejected by PREPARE through SHA;
  installed geometry and registry remain untouched.
- Separate full training case injects a controlled failure immediately after
  actual Windows ReplaceFile and before journal completion. Only that fault
  injection step uses the source fault worker; normal stages/recovery use EXE.
  Recovery rejects an unknown parallel registry edit without changing it.
  After reversing only the known injected test edit, recovery restores the
  exact pre-transaction registry SHA, unpublishes target documents to recovery
  storage and is idempotent on a second call. Expected negative-case FAILs
  are retained explicitly as PASS_EXPECTED_REFUSAL, not hidden.
- 32 Office supervisor records: zero owned orphans/ownership errors, no forced
  termination, all owned instances closed. Final Windows process inventory
  showed no Word/Excel processes; no foreign Office process was killed.
- **12,353 protected files unchanged**, including real registry/VOM/shared UZK,
  first documents and earlier deliveries. Old worktree remains clean at its
  original SHA. 51 packaged source modules match tested source SHA values.

## Delivery and remaining official blockers

Delivery contains EXE/dependencies, migrated production settings, internal
templates/geometry, real draft example, full training example, short guide,
technical evidence, source archive/local commit pin and SHA-256 manifest.
No new shortcut, old-version replacement, main change, push or merge.

PR24 integration/testing: GO. Official 11/2026 issue: NO_GO until trustworthy
current VIK/UZK results and primary protocols, current AKZ climate/thickness,
necessary quality facts and approval of forms/personnel/equipment/requisites
are supplied/confirmed. Historical favorable text and calculated dates do
not satisfy this. No new fields were added to the agreed simple screen.
This is the requested separate marked-forms test delivery, not an assertion
of final production 1.0.0 acceptance.

Evidence root: `C:\VZMKPassportChecks\RC3_MARKED_20261025`.
