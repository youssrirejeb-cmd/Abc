# DMIC Commissioning Contract Management Workbook

Excel workbook built to the specification
*DMIC Commissioning Contract Management Workbook v3*.

## Files

| File | Purpose |
|---|---|
| `DMIC_Commissioning_Contract_Management_Workbook_DEMO.xlsx` | Pre-filled with fictional data (3 scopes, 12 equipment items, 58 deliverables, 15 actions) so every KPI, colour rule and chart is alive on opening. |
| `DMIC_Commissioning_Contract_Management_Workbook_TEMPLATE.xlsx` | Production copy. Row 4 of each tracker holds one example record to overwrite. |
| `build_workbook.py` | Generator. Both workbooks come from this one script, so their structure, formulas, validation and formatting are identical by construction. |

Regenerate with:

```bash
python3 build_workbook.py --demo -o DMIC_Commissioning_Contract_Management_Workbook_DEMO.xlsx
python3 build_workbook.py        -o DMIC_Commissioning_Contract_Management_Workbook_TEMPLATE.xlsx
```

`--start YYYY-MM-DD` sets the seeded Planned Commissioning Start Date (default `2027-06-01`).

## Sheets

`Read Me` · `Control Panel` · `Scope Tracker` · `Equipment Tracker` ·
`Document Tracker` · `Action Tracker` · `Dashboard` · `Configuration` (hidden)

## How it works

Everything is driven by one date — **Planned Commissioning Start** on the Control
Panel. Target submission and approval dates for all 20 deliverable types are
generated from it via editable lead times.

Readiness rolls up: equipment → scope → contract.

- **Equipment Readiness** = 50% phase progress + 50% share of that item's linked
  documents approved. Equipment with no linked documents falls back to phase
  progress alone. Weights live on the `Configuration` sheet.
- **Scope Readiness** = average equipment readiness within the scope.
- **Contract Readiness** = average equipment readiness across the contract.
- **Contract Health Index** = 40% Documents + 20% Readiness + 15% Actions +
  10% Milestones + 10% Equipment Progress + 5% Handovers.
- **ATS Readiness** — a required type counts as met only when it exists in the
  tracker *and* none of that type is still outstanding.

## Assumptions

The specification left some things undefined. Each is documented in the
workbook itself (`Read Me` sheet, plus cell comments at the point of use):

- **Phase progress ramp** — the 11 phases carry no percentages in the spec;
  seeded 0% → 100% and editable.
- **Months → days** — converted at 30 days per month, which reproduces the
  spec's own worked example exactly (Planned Start 01-Jan-2028, Commissioning
  Procedure lead 90 days → Target Approval 03-Oct-2027).
- **Post-commissioning lead times** — "During Commissioning", "Before
  FTC/TFM/TFTO" and "Before Final Takeover" have no stated duration; seeded as
  T+30 to T+120 days.
- **"Missing Documents"** — counted as deliverables present as a row but still
  "Not Started". A row that was never created cannot be counted.
- **"Open Risks"** — the spec defines no Risk sheet; counted as open actions
  flagged as a Commissioning Blocker.
- **Authority To Start Evidence** — added as a 20th document type. The spec
  lists 19 mandatory types, but the ATS gate and the Control Panel lead-time
  list both require it.
- **Readiness Trend** — Excel keeps no history; record a monthly snapshot in the
  trend log on the Dashboard.

## Capacity

Scope 100 rows · Equipment 100 · Documents 150 · Actions 100 · Milestones 30.
To extend, copy the last data row down and widen the named ranges in
Formulas → Name Manager.

## Verification

Both files recalculate clean: 3053 formulas, 0 errors, checked with
LibreOffice. Roll-up values were verified by hand against the source data.
