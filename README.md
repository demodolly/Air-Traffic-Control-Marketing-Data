This repository will capture excel, word and powerpoint slides that capture the marketing data that is needed for the new Air Traffic Control project that will hae All marketing activities running through an orchestration tool and published to showflake. This repository will allow us to monitor all data elements needed with a valid reason for their use which can be refrenced when new projects are added that may impact Attribution and Operational Reporting, Lead Scoring Entry and Seller Enablement

## Current baseline (use this file)

**`Baseline Data Use Case Alignment - Day 20 V2.xlsx`** — prior baseline workbook.

**`Business Data Use Case Alignment - Day 21 V1.xlsx`** — current business alignment workbook (see the **README** tab in that file for use-case tab header definitions).

Supporting reference (field order within sections 2–6):

- **`Data Element Order.xlsx`**

## Source extracts (active)

| File | Purpose |
|------|---------|
| `Contact Us Forms.xlsx` | Contact Us form source data |
| `Offer Form Submissions.xlsx` | Offer form source data |

## Archived older versions

Superseded baselines live under **`archive/baseline-workbooks/`** (see `archive/README.md`). They are not maintained for day-to-day work.

Includes: Day 20 (original), Day 20 Reformatted, All Use Cases V1/V2, Offers V1/V2, and Section 6 reference workbook.

## Scripts

| Script | Status |
|--------|--------|
| `scripts/update_day21_v1_readme.py` | Appends Process / Lead Category / Use Case / Activity Type definitions to **Day 21 V1** README tab only |
| `scripts/revert_day20_v2_readme_headers.py` | Removes those definitions from Day 20 V2 README if present |
| `scripts/enhance_day20_v2_workbook.py` | Applies reviewer columns, navigation links, and README audit text to Day 20 V2 |
| `scripts/add_section6_fields.py` | Legacy — reads archived All Use Cases V2 |
| `scripts/create_section6_reference.py` | Legacy — reads archived All Use Cases V2 |
