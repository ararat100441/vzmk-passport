# PR19 — Windows acceptance, 25 October 2026

Delivery: `C:\VZMK-Passport\1.0.0-RC3-sources-test\VZMK-Passport-1.0.0.exe`.
Branch: `codex/rc3-sources-test` (local; no push, PR or merge).
Base: `7622459a90c219ff828e4ed8980a58739238c3f8`, the latest saved PR18 Windows integration.
Requested upstream: `19dcb1c4f6d91763c33160f5437b429e10513b39`.
Runtime commits: `e3d2732` and `e02d8c0`.
Evidence root: `C:\VZMKPassportChecks\RC3_SOURCES_20261025`.

## Integration

A separate worktree was created. The original PR18 worktree and older deliveries
were preserved; BEFORE_PR19.bundle is a recoverable checkpoint.
The PR18-to-PR19 patch was applied selectively using a three-way merge.
The settings-save conflict retained the current isolated settings path.
The Word, Excel, PDF engines and flat shipment VOM reader were not replaced:
INTEGRATION.json records their unchanged hashes.

Settings expose only registry, common Documents root and common UZK XLSM.
The first passport and selected project's catalogue are discovered automatically.
Prepared internal forms remain separate from historical filled acts. There is no
main Templates button or CSV creation/import workflow. Detailed current weld facts
can be entered in the existing results window.

Shared XLSM is read only, without VBA execution or saving. Per-mark quantities
are derived from the current VOM and selected project's remaining quantities;
they are immutable in the results grid. No inspection conclusion is auto-filled.
The source XLSM hash is pinned together with the remaining inputs.
The actual existing project folder's spelling is retained for new output.

## Fixes made during integration

1. A source-change race between reading UZK quantities and freezing CHECK input
   hashes could bind old quantities to a new XLSM hash. CHECK now rejects this
   interval. The regression changes a disposable valid XLSM between those steps;
   it failed before the fix and passes after it. See UZK_RACE_REPRO/FIXED.log.
2. A worker's cleanup exception could erase the primary operation exception.
   The worker now retains both error/traceback pairs and still fails closed.
   A reproducing regression failed before the fix and passes after it.

The first frozen candidate had a native Excel bind/teardown failure during CHECK.
Its original operation error was masked by the second defect, so the exact initial
bind cause is not claimed. The supervisor recorded normal native close-without-save
and Quit of its precisely owned Excel instance, no forced termination or orphan.
No journal or registry write occurred. That candidate and failure logs are retained.
The rebuilt candidate subsequently passed CHECK and both complete EXE runs below.
The diagnostic fix is not presented as proof of a cure for the unknown bind cause.

## Regression and built EXE

Final runtime regression, one invocation: **343 PASS, 0 FAIL, 0 ERROR, 0 SKIP**.
REGRESSION_MANIFEST.json identifies the current Windows/production and Office
safety suites and the archived date-bound scripts outside that suite.
REGRESSION_RESULT.json and regression_final.log contain the complete result.
PyInstaller build rc3-sources-test2 passed. BUILD_INPUTS.json pins its runtime
inputs. The executable and its dependencies were copied to the separate delivery;
older programs and shortcuts were not replaced. No desktop shortcut was created.

## Real inputs — read only

The actual shipment invoice was read directly and again through the final EXE:
104 marks, 288 items, mass 15822.14 kg, area 1021.644 square metres.
The current working registry proposes 11/2026; the first passport is 1/2026.
All six first-passport sources were discovered in the exact Garage Kazakhstan
project's Passport folder. Historical control results and signatures were not used
as current facts. The production output preview is:
`Работа\Документы\Гараж казахстан\Паспорт\11_2026 от ДД.ММ.ГГГГ`.

The real common UZK XLSM produces planned control of 33 items across four marks.
REAL_UZK_RESULT.json records every mark and source provenance, without conclusions.
The final GUI shows the three actual paths, working mode and real VOM totals.
FINAL_THREE_SETTINGS.png and FINAL_REAL_VOM.png are saved in gui_screens.

## Full native Office package on a copy

The installed executable's actual worker entry point ran, serially:
CHECK, PREPARE, VERIFY, CONFIRM, RELEASE, COMMIT, READBACK, COMPLETE: **PASS**.
This is executable/Office proof, not a source Python run. The external test driver
only supplies requests and checks responses. Training markers guard the registry.
The 104-mark VOM and shared XLSM were copied byte-for-byte. Current controls,
catalogue assignments, measurement ranges and protocol are explicitly synthetic.
Only the private training registry's date formulas were adjusted; no production
date or result was invented. Internal cleaned Word forms were used, not filled
historical acts. AKZ0 remains the explicitly identified one-layer form candidate.

Six editable files were reopened and checked by Microsoft Word/Excel. VIK contains
both acts in one two-page DOCX. The merged PDF has **13 pages**, all visually viewed:

| Pages | Document |
|---|---|
| 1–2 | Quality document |
| 3–5 | Appendix No. 1, all 104 VOM rows and totals |
| 6–7 | Two VIK acts |
| 8–9 | Detailed UZK, four actually controlled training marks |
| 10 | One-layer AKZ No. 0 |
| 11–13 | Corresponding AKZ appendix |

Seven Excel columns, table continuation headers, long object names and totals
remain visible. No clipped row or empty page was found. The second UZK page
contains its final inspector/qualification block; it is not an empty page.
Training ranges yielded air 21, surface 24, humidity 55, dew 10, dry thickness 110.
Final rendered pages are byte-identical to all 13 reviewed preview pages.
PDF filename equals its passport folder name. No physical printer was used.

## Ordinary GUI of the installed executable

A separate private LOCALAPPDATA and marked training registry were used.
No facts were injected into the GUI: keyboard/window input filled the existing
results screen, including a current synthetic detailed weld row without CSV.
The short VOM contains two real source marks, nine items, 443.92 kg, 16.956 square
metres. Its computed UZK quantities are 8 and 0; the initial conclusions were blank.
The training catalogue auto-filled profile, paint and AKZ0 assignments.
One-layer ranges were entered and their agreed midpoints displayed.

The one-button pipeline prepared Word/Excel and a six-page PDF on built-in generic
training forms. All six pages were visually reviewed. Clicking No left the registry
hash unchanged, without creating a transaction journal or publishing a passport.
Returning to the same preview and explicitly confirming training release completed
all eight stages through the normal GUI. Backup hash matches the pre-run registry;
the GUI then proposed 12/2026. Final pages match the reviewed preview exactly.
This short generic-form GUI test is distinct from the adapted 13-page package
tested through the same executable worker; it is not represented as a second
visual review of the historical/adapted production forms.

Windows displayed a separate persistent **Font Capture** exception, 0xc06d007e,
after Adobe Acrobat opened the preview. It also occluded captures of other windows.
The passport EXE completed normally, and PDFium independently rendered and verified
every preview and final page. The exact external fault cause is unproven; Windows,
fonts and viewer configuration were not changed. The external viewer warning is
not silently counted as a viewer PASS. Another installed PDF reader can open the
validated files. No foreign Office process was terminated.

## Recovery and source preservation

On another disposable training registry the fault worker deliberately raised after
Windows ReplaceFile and before journal update. Normal stages and recovery used
the installed EXE. A concurrent edit of that copy made recovery refuse without
overwriting it. After removing precisely the known injected edit, recovery restored
the exact original SHA, withdrew the unpublished target and was idempotent on retry.
G4_recovery/RESULT.json and recovery.log retain the expected FAILs and final PASS.

All normal final-run Office supervisors report zero owned orphans, no ownership
errors and graceful exit of their own instances. The original shared UZK and VOM
remain unchanged. PROTECTED_FINAL.json checks **7098 protected files**, including
original documents, registers and previous EXE deliveries: no changed or missing file. REAL_SOURCES_FINAL.json rechecks the
four actual sources after the last GUI read. No official passport was issued.

## Remaining real shipment blockers

This is a separately tested RC3 Sources delivery, not permission to issue 11/2026.
REAL_DATE_BLOCKER.json reproduces the current old literal date in H353 START_DATE.
The agreed formula-only new-date rule stops it; no alternative method was invented.
Current approved date sources are needed for this shipment's start/control dates.
The project's unchanged catalogue supplies profiles for 55 of 104 marks; 49
profiles and all paint/AKZ assignments still need confirmation. Current VIK/UZK/AKZ
protocols, actual measurements and required form requisites must be supplied and
confirmed. The new cleaned one-layer AKZ0 form needs responsible approval.
Old passport conclusions and signatures are never a substitute.

Integration/regression, native full package, normal GUI training and recovery:
**PASS**. Actual official 11/2026: **NOT ISSUED / NO_GO until its source data and
form approval are supplied**. External PDF viewer: warning described above.
