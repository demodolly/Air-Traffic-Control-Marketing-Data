This repository will capture excel, word and powerpoint slides that capture the marketing data that is needed for the new Air Traffic Control project that will hae All marketing activities running through an orchestration tool and published to showflake. This repository will allow us to monitor all data elements needed with a valid reason for their use which can be refrenced when new projects are added that may impact Attribution and Operational Reporting, Lead Scoring Entry and Seller Enablement

## Current baseline (use this file)

**`Baseline Data Use Case Alignment - Day 20 V2.xlsx`** — authoritative workbook. Includes use-case tabs, dynamic Excel counts, **Data Element Matrix**, **Ref - Master Schema**, **Use Case Index**, and a **README** sheet with workbook guidance.

### Use case tab headers (rows 3–6 on each UC sheet)

On every use case tab, column A is the field name and column B is the value for that use case.

| Header | Definition |
|--------|------------|
| **Process** | How the business refers to the overall action or channel family (for example Content Syndication, Pathfactory Webinars, Manual Uploads, or Online Forms). |
| **Lead Category** | What happens with the record after capture. **Hand Raiser** activities are processed for routing to the sales team. **Non-Hand Raiser** activities are captured as transactions for future nurturing and lead scoring. |
| **Use Case** | The specific type of action within the Process, aligned to the FY27 Vehicles and Activities documentation (the content vehicle). Examples: online forms → Offers - Contact Us or Offers - Demos; Pathfactory Webinars → Events - Webinars. |
| **Activity Type** | The most granular activity within the Use Case (also from FY27 Vehicles and Activities). This is the level of detail used in attribution—for example Events - Webinars: Registered, Attended Live, Attended Virtual. |

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
| `scripts/update_day20_v2_readme.py` | Updates only the **README** sheet header definitions in Day 20 V2 |
| `scripts/enhance_day20_v2_workbook.py` | Applies reviewer columns, navigation links, and README audit text to Day 20 V2 |
| `scripts/add_section6_fields.py` | Legacy — reads archived All Use Cases V2 |
| `scripts/create_section6_reference.py` | Legacy — reads archived All Use Cases V2 |
