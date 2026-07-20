This repository will capture excel, word and powerpoint slides that capture the marketing data that is needed for the new Air Traffic Control project that will hae All marketing activities running through an orchestration tool and published to showflake. This repository will allow us to monitor all data elements needed with a valid reason for their use which can be refrenced when new projects are added that may impact Attribution and Operational Reporting, Lead Scoring Entry and Seller Enablement

## Current baseline (use this file)

**`Baseline Data Use Case Alignment - Day 20 V2.xlsx`** — authoritative workbook. Includes use-case tabs, dynamic Excel counts, **Data Element Matrix**, **Ref - Master Schema**, and **Use Case Index**.

Supporting reference (field order within sections 2–6):

- **`Data Element Order.xlsx`**

## Superseded (removed from the repo root)

These older Day 20 files were deleted to avoid version confusion. They remain in **git history** if you need to recover them.

| Removed file | Replaced by |
|--------------|-------------|
| `Baseline Data Use Case Alignment - Day 20.xlsx` | Day 20 **V2** |
| `Baseline Data Use Case Alignment - Day 20 - Reformatted.xlsx` | Day 20 **V2** (same use-case layout, plus your formulas and matrix sheets) |

## Other workbooks in this repo

| File | Notes |
|------|--------|
| `Baseline Data Use Case Alignment - All Use Cases V1.xlsx` / `V2.xlsx` | Earlier baseline snapshots; not the active Day 20 workbook |
| `Baseline Data Use Case Alignment - Offers V1.xlsx` / `V2.xlsx` | Offers-specific alignment |
| `Contact Us Forms.xlsx` / `Offer Form Submissions.xlsx` | Source extracts for form parameters |
| `Section 6 Future Tracking Parameters.xlsx` | Section 6 field reference (optional) |

## Scripts

| Script | Status |
|--------|--------|
| `scripts/reformat_day20_workbook.py` | **Deprecated** — was used to build the old “Reformatted” file from the long-form Day 20 export. Edit **Day 20 V2** directly instead. |
