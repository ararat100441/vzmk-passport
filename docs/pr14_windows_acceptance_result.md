# PR14 Windows acceptance

Base: `774869eae1e15626eab8f70e0131e7306f83e9c1`.
Runtime fixes: `13980df6ea3f5abe212899e6721db3def279c128` and
`8fc458f7c7fa8ef8f65024316f4eb99c7a19bc66`.

## Changes verified on Windows

- Keep each UZK weld record on one page; verify the saved Word row property.
  This also keeps the closing information with data on the final diagnostic page.
- Print mass from the already verified registry row. Shipment calculations and
  registry writes are unchanged; no new rounding rule is introduced.
- Add tests for preserved mass precision, unchanged date refusal and intact rows.

## Evidence

| Gate | Result |
| --- | --- |
| G0 isolated worktree and source hashes | PASS |
| G1 native Word and visual review | PASS: 11 pages, 60 synthetic welds, 8 columns |
| G1 native Excel and visual review | PASS: 4 and 6 pages, 7 columns, automatic row heights |
| G2 current automated regression | PASS: 240 tests, zero failures/errors/skips in one run; historical ALL scope incomplete |
| G3 training Office CHECK through COMPLETE | PASS |
| G3 combined PDF | PASS: 13 pages, correct order and matching folder/file name |
| G4 actual ReplaceFile fault and rollback | PASS: parallel-edit refusal, exact SHA rollback, idempotence |
| G4 temporary PyInstaller EXE | PASS: smoke and full training cycle with System32-only PATH |
| G4 shipped GUI component integration | PASS: real data-dialog validation, CSV import, real frozen workers |
| G4 external mouse/keyboard GUI acceptance | BLOCKED: window activation/capture unavailable |

The current regression manifest includes every production test module and the
self-contained Stage6/Stage9/Office cleanup suites. Historical date-bound Stage5
review scripts require their old before/after evidence and were not rerun.

The complete package contains quality (2 pages), Appendix 1 (3), VIK (2), UZK (2),
AKZ act (1), AKZ appendix (3), in that order. All 13 final pages were reviewed.
The additional frozen-EXE and GUI-callback PDFs each have 13 pages and match the
reviewed package pixel for pixel. All control data are explicitly synthetic and
the program's isolated TRAINING mode is enforced.

The first visual clipping hypothesis was not confirmed by saved full-image
comparisons: the header content is present. Contact sheets and independent
PDFium rendering were retained; no original form was changed to hide clipping.

The temporary EXE still carries the existing `1.0.0-rc3` program version. It is a
separate PR14 smoke artifact, not an installed replacement or final 1.0.0 release.

## Remaining gate

Both supported external UI paths returned `failed to activate captured window`;
the capture showed desktop wallpaper. A responding program window was registered,
but external GUI selection/input is **not verified**. The shipped-source Tk
component test is separate evidence and does not replace this gate.

The literal requirement to rerun ALL historical tests is also not fully proved:
four Stage5 acceptance scripts remain NOT_RUN because their historical evidence
and execution context were not reconstructed. The 240-test result is not a claim
that those scripts passed.

Until external GUI acceptance and the historical regression scope are resolved,
the overall acceptance closeout is **NO_GO**. Draft only; no merge or production issue.
