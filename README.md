This repository will capture excel, word and powerpoint slides that capture the marketing data that is needed for the new Air Traffic Control project that will hae All marketing activities running through an orchestration tool and published to showflake. This repository will allow us to monitor all data elements needed with a valid reason for their use which can be refrenced when new projects are added that may impact Attribution and Operational Reporting, Lead Scoring Entry and Seller Enablement

## Current baseline (use this file)

**`Baseline Data Use Case Alignment - Day 20 V2.xlsx`** — authoritative workbook. Includes use-case tabs, dynamic Excel counts, **Data Element Matrix**, **Ref - Master Schema**, and **Use Case Index**.

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
| `scripts/reformat_day20_workbook.py` | **Deprecated** — edit **Day 20 V2** directly in Excel (do not run openpyxl scripts on V2; they break formulas) |
| `scripts/add_section6_fields.py` | Legacy — reads archived All Use Cases V2 |
| `scripts/create_section6_reference.py` | Legacy — reads archived All Use Cases V2 |
