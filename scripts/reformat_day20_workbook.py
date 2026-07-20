#!/usr/bin/env python3
"""Reformat Day 20 baseline alignment into a use-case-centric workbook."""

from __future__ import annotations

import re
from collections import defaultdict

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

SOURCE_PATH = "/workspace/Baseline Data Use Case Alignment - Day 20.xlsx"
OUTPUT_PATH = (
    "/workspace/Baseline Data Use Case Alignment - Day 20 - Reformatted.xlsx"
)

HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
SECTION_FILL = PatternFill(start_color="D6DCE4", end_color="D6DCE4", fill_type="solid")
SECTION_FONT = Font(bold=True)
UC_HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
UC_HEADER_FONT = Font(bold=True, color="FFFFFF")
ALT_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

INDEX_HEADERS = [
    "Use Case ID",
    "Process",
    "Lead Category",
    "Use Case",
    "Activity Type",
    "Data Element Count",
    "Open Questions Count",
]

REQUIREMENT_HEADERS = [
    "Use Case ID",
    "Process",
    "Lead Category",
    "Use Case",
    "Activity Type",
    "Section",
    "Data Element",
    "Display Status",
    "Why We Need This",
    "How It Is Captured (Value Source)",
    "Data Rules",
    "Business-Friendly Prompt",
    "Input Type",
    "Attribution",
    "Operational Reporting",
    "Lead Scoring",
    "Seller Enablement",
    "Outstanding Questions",
    "Amanda Reviewed",
    "Review Status",
]

PROCESS_VIEW_HEADERS = [
    "Section",
    "Data Element",
    "Display Status",
    "Why We Need This",
    "How It Is Captured (Value Source)",
    "Data Rules",
    "Outstanding Questions",
    "Attribution",
    "Operational Reporting",
    "Lead Scoring",
    "Seller Enablement",
    "Amanda Reviewed",
]


def style_header_row(ws, row_num: int = 1) -> None:
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="top")


def auto_width(ws, max_width: int = 55) -> None:
    for col in range(1, ws.max_column + 1):
        max_len = 0
        letter = get_column_letter(col)
        for row in range(1, min(ws.max_row, 200) + 1):
            value = ws.cell(row=row, column=col).value
            if value is not None:
                max_len = max(max_len, len(str(value)))
        ws.column_dimensions[letter].width = min(max_len + 2, max_width)


def merge_outstanding_questions(row: tuple) -> str | None:
    parts = []
    business = row[13]
    if business:
        parts.append(str(business).strip())
    for idx, label in ((19, "Kim"), (20, "Stephanie"), (21, "Briana")):
        value = row[idx]
        if value:
            parts.append(f"{label}: {str(value).strip()}")
    if not parts:
        return None
    return " | ".join(parts)


def load_source_rows():
    wb = openpyxl.load_workbook(SOURCE_PATH, data_only=True)
    ws = wb["Sheet1"]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        rows.append(row)
    wb.close()
    return rows


def build_use_case_index(rows):
    combos = []
    seen = set()
    combo_counts = defaultdict(int)
    combo_questions = defaultdict(int)
    for row in rows:
        combo = (row[1], row[2], row[3], row[4])
        combo_counts[combo] += 1
        if merge_outstanding_questions(row):
            combo_questions[combo] += 1
        if combo not in seen:
            seen.add(combo)
            combos.append(combo)

    index = []
    for idx, combo in enumerate(sorted(combos), start=1):
        uc_id = f"UC{idx:02d}"
        index.append(
            {
                "id": uc_id,
                "combo": combo,
                "field_count": combo_counts[combo],
                "question_count": combo_questions[combo],
            }
        )
    return index


def combo_to_id(index):
    return {entry["combo"]: entry["id"] for entry in index}


def requirement_record(row, uc_id):
    questions = merge_outstanding_questions(row)
    return [
        uc_id,
        row[1],
        row[2],
        row[3],
        row[4],
        row[5],
        row[6],
        row[7],
        row[12],
        row[9],
        row[10],
        row[8],
        row[11],
        row[14],
        row[15],
        row[16],
        row[17],
        questions,
        row[18],
        None,
    ]


def apply_review_status_validation(ws, column_letter: str, max_row: int) -> None:
    dv = DataValidation(
        type="list",
        formula1='"Confirmed,In Review,Blocked,Not Applicable"',
        allow_blank=True,
    )
    dv.error = "Choose a review status from the list."
    dv.errorTitle = "Invalid Review Status"
    ws.add_data_validation(dv)
    dv.add(f"{column_letter}2:{column_letter}{max_row}")


def write_readme_sheet(wb):
    ws = wb.active
    ws.title = "README"
    lines = [
        "Baseline Data Use Case Alignment — Reformatted Workbook",
        "",
        "Purpose",
        "This workbook reorganizes the Day 20 baseline so you can work one use case at a time,",
        "see why each data element is needed, and track outstanding questions in one place.",
        "",
        "Sheets",
        "1. Use Case Index — master list of all use case combinations (UC01, UC02, ...).",
        "2. Use Case Requirements — one row per data element per use case (filter by Use Case ID).",
        "3. Open Questions — only rows with Outstanding Questions filled in.",
        "4. Process Views — Content Syndication, Manual Uploads, Online Forms (grouped by use case).",
        "",
        "How to use",
        "- Start on Use Case Index and pick a Use Case ID.",
        "- Filter Use Case Requirements on that ID to see every data element, why it is needed,",
        "  and any open questions.",
        "- Use Review Status (blank by default) to track Confirmed / In Review / Blocked.",
        "- Open Questions sheet is your action list across all use cases.",
        "",
        "Source file: Baseline Data Use Case Alignment - Day 20.xlsx",
        "Regenerate with: python3 scripts/reformat_day20_workbook.py",
    ]
    for i, line in enumerate(lines, start=1):
        cell = ws.cell(row=i, column=1, value=line)
        if i == 1:
            cell.font = Font(bold=True, size=14)
        elif line in ("Purpose", "Sheets", "How to use"):
            cell.font = Font(bold=True)
    ws.column_dimensions["A"].width = 100


def write_index_sheet(wb, index):
    ws = wb.create_sheet("Use Case Index")
    ws.append(INDEX_HEADERS)
    style_header_row(ws)
    for entry in index:
        combo = entry["combo"]
        ws.append(
            [
                entry["id"],
                combo[0],
                combo[1],
                combo[2],
                combo[3],
                entry["field_count"],
                entry["question_count"],
            ]
        )
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    auto_width(ws)


def write_requirements_sheet(wb, rows, id_map):
    ws = wb.create_sheet("Use Case Requirements")
    ws.append(REQUIREMENT_HEADERS)
    style_header_row(ws)
    for row_num, row in enumerate(rows, start=2):
        combo = (row[1], row[2], row[3], row[4])
        ws.append(requirement_record(row, id_map[combo]))
        if row_num % 2 == 0:
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row_num, column=col).fill = ALT_FILL
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    apply_review_status_validation(ws, "T", ws.max_row)
    auto_width(ws)
    return ws


def write_open_questions_sheet(wb, requirements_ws):
    ws = wb.create_sheet("Open Questions")
    headers = [
        "Use Case ID",
        "Process",
        "Use Case",
        "Activity Type",
        "Section",
        "Data Element",
        "Outstanding Questions",
        "Why We Need This",
        "Review Status",
    ]
    ws.append(headers)
    style_header_row(ws)
    col = {h: i + 1 for i, h in enumerate(REQUIREMENT_HEADERS)}
    for r in range(2, requirements_ws.max_row + 1):
        question = requirements_ws.cell(row=r, column=col["Outstanding Questions"]).value
        if not question:
            continue
        ws.append(
            [
                requirements_ws.cell(row=r, column=col["Use Case ID"]).value,
                requirements_ws.cell(row=r, column=col["Process"]).value,
                requirements_ws.cell(row=r, column=col["Use Case"]).value,
                requirements_ws.cell(row=r, column=col["Activity Type"]).value,
                requirements_ws.cell(row=r, column=col["Section"]).value,
                requirements_ws.cell(row=r, column=col["Data Element"]).value,
                question,
                requirements_ws.cell(row=r, column=col["Why We Need This"]).value,
                requirements_ws.cell(row=r, column=col["Review Status"]).value,
            ]
        )
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    apply_review_status_validation(ws, "I", ws.max_row)
    auto_width(ws)


def safe_sheet_name(name: str) -> str:
    cleaned = re.sub(r"[\[\]\:\*\?\/\\]", " ", name)
    return cleaned[:31]


def write_process_view_sheet(wb, process_name, index, rows, id_map):
    ws = wb.create_sheet(safe_sheet_name(f"{process_name} View"))
    current_row = 1

    process_entries = [e for e in index if e["combo"][0] == process_name]
    for entry in process_entries:
        uc_id = entry["id"]
        combo = entry["combo"]
        title = (
            f"{uc_id} | {combo[0]} | {combo[1]} | {combo[2]} | {combo[3]}"
        )
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(PROCESS_VIEW_HEADERS))
        cell = ws.cell(row=current_row, column=1, value=title)
        cell.fill = UC_HEADER_FILL
        cell.font = UC_HEADER_FONT
        current_row += 1

        for col, header in enumerate(PROCESS_VIEW_HEADERS, start=1):
            c = ws.cell(row=current_row, column=col, value=header)
            c.fill = HEADER_FILL
            c.font = HEADER_FONT
            c.alignment = Alignment(wrap_text=True, vertical="top")
        current_row += 1

        combo_rows = [r for r in rows if (r[1], r[2], r[3], r[4]) == combo]
        section_order = []
        for r in combo_rows:
            if r[5] not in section_order:
                section_order.append(r[5])

        for section in section_order:
            section_rows = [r for r in combo_rows if r[5] == section]
            ws.merge_cells(
                start_row=current_row,
                start_column=1,
                end_row=current_row,
                end_column=len(PROCESS_VIEW_HEADERS),
            )
            sec_cell = ws.cell(row=current_row, column=1, value=section)
            sec_cell.fill = SECTION_FILL
            sec_cell.font = SECTION_FONT
            current_row += 1

            for r in section_rows:
                values = [
                    "",
                    r[6],
                    r[7],
                    r[12],
                    r[9],
                    r[10],
                    merge_outstanding_questions(r),
                    r[14],
                    r[15],
                    r[16],
                    r[17],
                    r[18],
                ]
                for col, value in enumerate(values, start=1):
                    ws.cell(row=current_row, column=col, value=value)
                current_row += 1

        current_row += 1

    auto_width(ws, max_width=50)


def main():
    rows = load_source_rows()
    index = build_use_case_index(rows)
    id_map = combo_to_id(index)

    wb = openpyxl.Workbook()
    write_readme_sheet(wb)
    write_index_sheet(wb, index)
    requirements_ws = write_requirements_sheet(wb, rows, id_map)
    write_open_questions_sheet(wb, requirements_ws)

    processes = sorted({entry["combo"][0] for entry in index})
    for process in processes:
        write_process_view_sheet(wb, process, index, rows, id_map)

    wb.save(OUTPUT_PATH)
    print(f"Wrote {OUTPUT_PATH}")
    print(f"  Use cases: {len(index)}")
    print(f"  Requirement rows: {len(rows)}")
    open_q = wb["Open Questions"].max_row - 1
    print(f"  Open question rows: {open_q}")


if __name__ == "__main__":
    main()
