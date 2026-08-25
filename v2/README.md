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

TRACKER 500 rows · ACTIONS 150 · systems 25 · equipment 80 · milestones 15 ·
core types tracked for gaps 12.

## Verification

Both files recalculate clean: 20 366 formulas, 0 errors (LibreOffice). All 40 734
generated formulas checked for balanced parentheses. Roll-ups hand-checked
against the source data: contract earliest start, target-date arithmetic, share
late, and gap totals reconcile across contract / system / equipment levels.
