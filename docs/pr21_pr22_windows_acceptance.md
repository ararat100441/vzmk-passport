# PR21 + PR22 — Windows integration, 25.10.2026

## Result and scope

Integration and regression PASS. The separate onedir application is
`C:\VZMK-Passport\1.0.0-RC3-updated\VZMK-Passport-1.0.0.exe`.
This is an RC3 integration build, not a declaration of unrestricted production
readiness. Official passport 11/2026 was NOT issued.

The complete native Office transaction PASS below used the installed EXE's
existing worker entry points. Fresh GUI input/Back/settings acceptance was
performed separately. A complete fresh GUI-to-release transaction is BLOCKED:
the requested PR22 interface removes the input path for required primary control
facts and approved-form requisites. Such facts cannot be inferred from VOM,
planned UZK counts, first-passport acts, or climate ranges. No hidden GUI data
injection or manufactured positive verdict was used to claim GUI E2E PASS.

## Source preservation and integration

- Latest installed PR20 source: `c41bacd909803039771084080d022b7cd4f76ccb`,
  worktree `passport-rc3-screen\Pasport`, clean before work.
- Integrated upstream `f9fb2ed12495be0b2b1ed95cb65d40468934d9f9` relative to its
  PR20 base `79084d3cfe33ee91bab2b586a1daba3f8ecb1b6c`, using three-way application
  and explicit conflict resolution, preserving later local fixes.
- New worktree: `C:\Users\PUBLIC_USER\.codex\worktrees\passport-rc3-updated\Pasport`.
  Local branch: `codex/rc3-updated`. No push or PR/main merge.
- Recoverable pre-integration checkpoint: verified all-refs bundle
  `C:\VZMKPassportChecks\RC3_UPDATED_20261025\BEFORE_PR21_22.bundle`.
- The early GUI date helper now accepts the fixed first-passport start date
  while still requiring current control-date formulas. Added a functional test
  for fixed start acceptance and literal current-control-date refusal.
- Preserved local preview pins, shared-UZK source-change refusal, current-project
  output folder selection and legacy preview handling. Resolved the preview
  callback to a single explicit final-consent handler.
- New settings filename `rc3-updated-settings.json` isolates previous versions.
- Updated superseded screen assertions and a race-test mock to accept the new
  `defer_approvals` argument; production validation and race checks remain active.
- Final comparison: 36 pre-existing engine/parser Python modules unchanged;
  41 packaged source modules exactly match tested worktree files.

## Windows regression

Final single run: **383 PASS, 0 FAIL, 0 ERROR, 0 SKIP**.
Manifest and verbose output:
`C:\VZMKPassportChecks\RC3_UPDATED_20261025\REGRESSION_MANIFEST.json`,
`REGRESSION_RESULT.json`, `regression-final.log`.
The manifest collects current production, stage6, stage9 and office-unbound
unittests. Historic stage5 evidence-dependent acceptance scripts are explicitly
listed as outside this runner, not reported as current native acceptance.

The first diagnostic run had 382 tests, 0 failures and 1 error: an older shared
UZK race-test mock did not accept the new keyword argument. Its log and result
are retained as `regression-initial.log` and `REGRESSION_INITIAL.json`. After
fixing the mock and adding the fixed-start helper regression, the final suite
above passed. Runtime code did not change after build and final regression.

## Real source preflight and GUI

Real registry read-only: `Работа\ВЗМК (2).xlsx`.
H353 is a fixed date **05.08.2026**; first Garage passport **1/2026**, next
**11/2026**. Current VIK/UZK/AKZ source cells M353/N353/O353 are formulas.
The original date/formulas were not edited. Training copies alone have explicit
synthetic current-control formula dates; the fixed first start remains original.

| VOM | Marks | Items | Mass kg | Area m2 | Planned UZK |
|---|---:|---:|---:|---:|---:|
| Current real desktop file | 4 | 11 | 511.98 | 19.533 | 8 |
| Preserved full real VOM copy | 104 | 288 | 15822.14 | 1021.644 | 33 |

The full VOM was independently read from the previously preserved exact source
copy in `RC3_SOURCES_20261025\inputs\REAL_VOM_COPY.xlsx`; the current desktop
file contains four marks and is not misreported as containing 104.
Detailed paths, raw sums and source SHA: `REAL_PREFLIGHT_PR21_22.json`.

Native installed-EXE GUI checks:

- Selected Garage; real four-mark VOM loaded, correct quantities/totals and
  automatic first-passport source discovery. Shared UZK reads planned counts.
- Facts has four columns, only applicable AKZ0 (one coat), automatic Magnus dew
  point, no proof/detail/approval tables below climate, fixed Back footer.
- Back with empty climate returns without validation. Back with partial air
  value 23 and a still-active edited profile retains both on reopening.
- A separate synthetic keyboard-only GUI range check displayed the automatic
  dew point -17.3 C for air midpoint25/surface2/RH5. These unusual explicit
  test values stayed in a private training session and were never published.
- Escape returns to main without generating or committing documents. Reopen
  persistence was visually checked for Back; Escape/X bindings are additionally
  covered by source-contract regression, not claimed as separate repeated GUI
  persistence runs.
- Normal environment startup displays WORK mode and the real registry/archive/
  shared-UZK paths, without training profile contamination. Settings has only
  those three paths and Recovery, no Create Training button.
- Registry hash remained unchanged during draft input; no draft GUI transaction
  journal was created. Screenshots: `gui_evidence` in the evidence root.

The built EXE separately rejects a complete-range request with missing current
verdicts: **Нет подтверждённого результата УЗК: Бп2-1**.
No registry write; result `missing_facts_checked\VERDICT.json`.
The initial negative-test fixture omitted the training-copy marker and was
correctly refused by training_guard. Its failed harness assertion is retained
in `missing_facts`; the corrected fixture used the normal training_copy helper.
It was not a production defect or a bypass of the training guard.

## Native Office EXE transaction and visual acceptance

The actual installed EXE completed CHECK → PREPARE → VERIFY → CONFIRM → RELEASE
→ COMMIT → READBACK → COMPLETE using a program-created isolated registry copy,
104-mark VOM, explicit synthetic protocols and blank signatures.
`native_full\E2E_STATE.json` and `STAGE_*.json` retain each result.

| Document | Editable file | PDF pages |
|---|---|---|
| Quality document | DOCX | 1–2 |
| Appendix No.1 | XLSX | 3–5 |
| VIK, two acts | DOCX | 6–7 |
| Detailed UZK | DOCX | 8–9 |
| AKZ0, one coat | XLSX | 10 |
| AKZ appendix0 | XLSX | 11–13 |

All 13 preview pages were rendered and visually reviewed: repeated headers,
seven appendix columns, full mark rows, no empty/cropped sheets or overlaid
text. Word/Excel native reopen and checks passed for all six editable files.
Dates: project start 05.08.2026, training issue 25.10.2026. AKZ values in this
package use air25/surface21/RH55/dew15.3, dry midrange110 (95..125), per the
agreed method. All pages are visibly training material, not real control proof.

Final PDF is beside DOCX/XLSX, named exactly as its passport folder. Final page
renders matched all reviewed preview page image hashes. `FINAL_READBACK.json`:
selected Garage row432, passport11/2026, correct fixed start, journal COMPLETE.
The duplicate-number readback lookup was scoped to the selected project;
matching numbers in other objects are not duplicate Garage passports.

Both declined preview/consent combinations were refused with unchanged copied
registry and no confirmation file (`DECLINE_RESULT.json`). Positive approval
was passed only after the 13-page visual review, in the explicit training
worker request. This is not a claim that a human approved an official passport.

## Recovery, Office ownership and SHA

`native_full\G4_recovery\RESULT.json`: actual Windows ReplaceFile followed by
a controlled exception before journal update; parallel private-registry edit
correctly refused recovery without changing bytes; after removal of that
known test edit recovery restored exact original SHA and was idempotent.
The candidate output was not published. This test used disposable copies only.

All Office supervisors reported zero owned orphans, zero ownership errors and
graceful owned-Office exits. The user's pre-existing Excel PID45604 was left
running and was not closed; no newly orphaned Word/Excel processes remained.
Final SHA comparison: **9936 protected files unchanged** including real registry,
shared UZK, source documents, old deliveries and previous source checkpoints.
`FINAL_PROTECTED_SHA.json` contains counts and EXE SHA. No old EXE, shortcut or
original document was replaced.

## Remaining real blockers

Official fresh GUI issue is **NO_GO** until trustworthy current VIK/UZK/AKZ
results, detailed UZK connections/protocols and required approved-form requisites
are available from approved sources. PR22 intentionally removed their daily
input controls; the current simplified screen cannot supply these missing
facts. Old first-passport results/signatures are never reused. This limitation
is already explicit in `rc3_simple_facts_back_integration.md` and was preserved
instead of changing the agreed interface or weakening validation.

GO for opening and checking this separate integrated RC3 GUI. NO_GO for calling
it an unconditionally accepted final production release or reporting a complete
fresh GUI issuance as PASS. No official issue, GitHub push or main merge.
