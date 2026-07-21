#!/usr/bin/env python3
"""Append use-case tab header definitions to the README sheet only (Day 20 V2)."""

from __future__ import annotations

import openpyxl
from openpyxl.styles import Font

WORKBOOK_PATH = "/workspace/Baseline Data Use Case Alignment - Day 20 V2.xlsx"
README_SHEET = "README"
START_ROW = 21


def write_header_definitions(ws) -> None:
    lines: list[tuple[str, bool, int | None]] = [
        ("", False, None),
        ("Use case tab headers (rows 3–6 on each UC sheet)", True, None),
        (
            "Column A lists the classification field; column B holds the value for that use case tab. "
            "Together they describe how the business classifies the touchpoint before the data element sections.",
            False,
            None,
        ),
        ("", False, None),
        ("Process", True, None),
        (
            "How the business refers to the overall action or channel family—for example Content Syndication, "
            "Pathfactory Webinars, Manual Uploads, or Online Forms.",
            False,
            None,
        ),
        ("", False, None),
        ("Lead Category", True, None),
        (
            "Determines what happens with the record after capture. Hand Raiser activities are processed "
            "for routing to the sales team. Non-Hand Raiser activities are still captured as transactions "
            "for future nurturing and lead scoring.",
            False,
            None,
        ),
        ("", False, None),
        ("Use Case", True, None),
        (
            "Distinguishes the specific type of action within the Process. This aligns with the FY27 Vehicles "
            "and Activities documentation—the content vehicle. Examples: for online forms, Offers - Contact Us "
            "or Offers - Demos; for Pathfactory Webinars, Events - Webinars.",
            False,
            None,
        ),
        ("", False, None),
        ("Activity Type", True, None),
        (
            "The most granular level of activity captured within the Use Case, also aligned to FY27 Vehicles "
            "and Activities. This is the level of detail used in attribution. Example: for Events - Webinars, "
            "activities include Registered, Attended Live, and Attended Virtual.",
            False,
            None,
        ),
    ]

    for offset, (text, bold, size) in enumerate(lines):
        row = START_ROW + offset
        cell = ws.cell(row=row, column=1, value=text)
        if bold:
            cell.font = Font(bold=True, size=size or 11)


def main() -> None:
    wb = openpyxl.load_workbook(WORKBOOK_PATH)
    ws = wb[README_SHEET]
    write_header_definitions(ws)
    wb.save(WORKBOOK_PATH)
    print(f"Updated README sheet in {WORKBOOK_PATH}")


if __name__ == "__main__":
    main()
