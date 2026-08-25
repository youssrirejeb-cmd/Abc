# DMIC Commissioning Control (v2)

Merges the v50 calculation engine with a flat TRACKER table and the visual
treatment of v1.

| File | Purpose |
|---|---|
| `DMIC_Commissioning_Control.xlsx` | Production file. One worked example per sheet, to overwrite. |
| `DMIC_Commissioning_Control_DEMO.xlsx` | 6 systems, 20 equipment items, 96 deliverables, 18 actions, so every list, colour rule and chart is populated. Use it as a test copy, not as the working file. |
| `build_v2.py` | Generator. Both files come from this one script, so their structure, formulas, validation and formatting are identical by construction. |

```bash
python3 build_v2.py        -o DMIC_Commissioning_Control.xlsx
python3 build_v2.py --demo -o DMIC_Commissioning_Control_DEMO.xlsx
```

## Sheets

`DASHBOARD` · `TRACKER` · `ACTIONS` · `SETUP` · `HOW IT WORKS` · `LISTS` (hidden)

## The one rule

Target approval date = the earliest commissioning start among what the document
covers (actual if known, otherwise planned) − the lead time of its type.
Post-commissioning types are measured back from their anchor milestone instead.

Days late: positive = late, negative = float. One convention everywhere.

The "actual if known" half is real: fill **Actual start** against an equipment
item on SETUP and every target behind that item is measured from the actual date
from then on. The hidden `Start used` column (SETUP, section 7) shows which of
the two dates is in force. Equipment with neither date leaves its documents with
a blank target, and the `Check` column says so in words.

## Document identity

Two columns on the TRACKER: **Teamcenter no.** and **Rev.** Both are optional
metadata, not a gate — a row is real, counted, and coloured the moment **Type**
is filled in, whatever else is still blank. (There was a third column,
**Document no.**, in an earlier revision. It was removed: every colour and
total formula gated on it being non-empty as well as Type, so a row typed in
without a document number never lit up. Teamcenter no. and Rev. gate nothing.)

## SETUP is protected against row deletion

The rest of the workbook reads SETUP by position, so deleting a row there
destroys a named reference and the dashboard lists go blank without saying why.
SETUP is therefore protected against inserting and deleting rows and columns —
every blue cell stays editable as normal. To drop a status code, clear the cells
on that line; an empty line is ignored everywhere. There is no password, so
Review → Unprotect Sheet lifts it for anyone who really means to restructure the
sheet — and if a deletion does break a reference, a red banner appears at the top
of the DASHBOARD rather than the lists silently emptying.

Because rows can't be inserted either, running out of capacity in a SETUP
section (more than 50 systems, more than 80 equipment, ...) isn't something to
fix by hand in Excel — it means asking for the workbook to be regenerated with
a bigger number. `build_v2.py` computes every SETUP section's row range from
the section above it (`_section()`, near the top of the file), so raising
`N_SYS`, `N_EQP`, `N_MS_SLOT`, or adding a `WF_CODES` / `TYPES` entry and
rebuilding is the whole change — no row number anywhere else in the file needs
touching. (An earlier revision hard-coded each section's starting row; raising
`N_SYS` alone made the SYSTEMS section overrun into EQUIPMENT's first rows and
broke both. That's fixed structurally now, not just for this one number.)

## Kept from v50

- Per-coverage target dates (`All` / `System` / `Equipment`).
- A named anchor milestone per document type, not a guessed day count.
- 24 workflow codes with a visible Family / "Counts as" mapping. S4A and the
  P-codes stop the delay counter; A-codes and W2 are Inactive.
- Core-gap detection: the expected matrix is built from the type library and the
  real absences are reported, ordered by when each would have been needed.
- The `Check` column — a per-row data-quality validator.
- Fully parameterised wording; SETUP as the single control surface.

## Changed from v50

- **Flat TRACKER instead of ~30 fixed blocks.** v50 gave 12 rows to each
  equipment-scoped type while its own SETUP allowed 80 equipment items, so it
  overflowed at 13 equipment while ~150 rows sat unused in other blocks. Capacity
  is now shared across 500 rows, and filter/sort works across every type at once.
- **Dashboard lists show "10 of N"**, so a truncated list is never mistaken for
  the whole picture.
- **No `MINIFS`.** v50 needed Excel 2019+; the equivalent `SUMPRODUCT(MIN(...))`
  runs on any Excel from 2007.
- **Target override + Source column**, so a manual date never silently destroys
  the formula.
- **REVIEW folded into the DASHBOARD** (by system, equipment to watch, milestones).

## Kept from v1

Data bars with the percentage shown, coloured states, section banners, four
charts, the assumptions documented in cell comments at the point of use, and the
recalculate-to-zero-errors discipline.

## Capacity

TRACKER 500 rows · ACTIONS 150 · systems 50 · equipment 80 · milestones 15 ·
core types tracked for gaps 12.

## Verification

Both files recalculate clean: 22 024 formulas, 0 errors (LibreOffice). Roll-ups
hand-checked against the source data: contract earliest start, target-date
arithmetic, share late, and gap totals reconcile across contract / system /
equipment levels.

Additionally swept, in both files, for the two failure modes that produce a date
cell no column width can fix: **0 cells holding a negative date serial** (what
Excel renders as `######`) and **0 cells sitting on the 1900 epoch**. Every date
that reaches a visible cell is guarded by an `N(...)<=0` test rather than a
comparison against `""` — `INDEX` on a blank cell returns `0`, not `""`, and in
Excel `0=""` is FALSE, so the old guard let `0 − lead` through as a negative
serial.
