#!/usr/bin/env python3
"""
spreadsheetml_to_csv.py

Converts Excel 2003 XML Spreadsheet (SpreadsheetML) to output.csv.

- Prompts for XML file path
- Uses the first Row as headers
- Writes subsequent Rows as data
- Handles ss:Index on Cell (skipped columns)
"""

from __future__ import annotations

import csv
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Optional


NS = {
    "ss": "urn:schemas-microsoft-com:office:spreadsheet",
}


def get_cell_text(cell: ET.Element) -> str:
    data = cell.find("ss:Data", NS)
    if data is None or data.text is None:
        return ""
    return data.text.strip()


def parse_rows(table: ET.Element) -> List[List[str]]:
    rows_out: List[List[str]] = []

    for row in table.findall("ss:Row", NS):
        values: List[str] = []
        current_col = 1  # SpreadsheetML columns are 1-based when using ss:Index

        for cell in row.findall("ss:Cell", NS):
            idx = cell.get(f"{{{NS['ss']}}}Index")  # ss:Index attribute
            if idx is not None:
                try:
                    target_col = int(idx)
                except ValueError:
                    target_col = current_col
                # Fill gaps with blanks
                while current_col < target_col:
                    values.append("")
                    current_col += 1

            values.append(get_cell_text(cell))
            current_col += 1

        rows_out.append(values)

    return rows_out


def main() -> int:
    xml_path_str = input("Enter the full path to the XML file: ").strip().strip('"').strip("'")
    xml_path = Path(xml_path_str)

    if not xml_path.exists() or not xml_path.is_file():
        print(f"Error: file not found: {xml_path}", file=sys.stderr)
        return 2

    try:
        tree = ET.parse(xml_path)
    except ET.ParseError as e:
        print(f"Error: failed to parse XML: {e}", file=sys.stderr)
        return 2

    root = tree.getroot()

    # Find first Worksheet/Table
    worksheet = root.find("ss:Worksheet", NS)
    if worksheet is None:
        print("Error: couldn't find <Worksheet> in this SpreadsheetML file.", file=sys.stderr)
        return 2

    table = worksheet.find("ss:Table", NS)
    if table is None:
        print("Error: couldn't find <Table> inside the worksheet.", file=sys.stderr)
        return 2

    all_rows = parse_rows(table)
    if not all_rows:
        print("Error: no rows found in the table.", file=sys.stderr)
        return 2

    headers = all_rows[0]
    data_rows = all_rows[1:]

    # Normalize row lengths to header length
    header_len = len(headers)
    normalized: List[List[str]] = []
    for r in data_rows:
        if len(r) < header_len:
            r = r + [""] * (header_len - len(r))
        elif len(r) > header_len:
            r = r[:header_len]
        normalized.append(r)

    out_csv = Path.cwd() / "output.csv"
    with out_csv.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(normalized)

    print(f"Done. Wrote {len(normalized)} rows to {out_csv}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
