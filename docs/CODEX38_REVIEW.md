# CODEX38: editable project quantity and historical accounting

Column B in the active project ledger is the current editable total. Increases
and decreases require no migration or manual checksum refresh. E uses B minus
the confirmed F+ shipments. Checks reject totals below already shipped stock,
noninteger totals, altered historical quantities and requests above availability.
Unfilled totals of unrelated marks remain unknown. Technical `_VZMK_SOURCE`
does not supply project quantities.

An absent historical source path is distinguished from changed XLSX bytes.
Relocation is limited to one file with the same exact basename in the same
project's Passport folder or one of its immediate subfolders. Ambiguity stops.
On either relocation or byte mismatch, exact number/date/object, confirmed mark
quantities, row masses against registry mass, and supported total formulas are
checked. Formula-based body quantities/masses without independently verified
values stop. Once recorded, per-mark mass must also remain identical.
Other byte changes are classified only as accounting-identical, with no claim
that every document cell is unchanged or that the exact editing operation is known.

Reconciliation never saves the historical XLSX. CHECK is read only. The new
source/hash and prior binding audit are written into the shipment candidate.
Normal coordinated commit verifies and retains the original ledger backup
before replacement; history bindings cannot bypass its locks, input SHA pins,
confirmation, journal, recovery or rollback. Until commit the original binding
is retained and subsequent checks repeat the reconciliation automatically.

Tests use isolated copies and training mode. No official passport is issued.
Independent review is required. PRODUCTION READY is not assigned.
