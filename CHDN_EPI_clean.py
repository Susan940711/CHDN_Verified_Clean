from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook


def _is_blank(value) -> bool:
    return value is None or str(value).strip() == ""


def build_clean_sheet(
    input_file: Path | str,
    output_file: Path | str,
    source_sheet_name: str = "EPI-Child",
    target_sheet_name: str = "EPI-Child",
    pregnancy_source_sheet_name: str = "EPI-Pregnancy",
    pregnancy_target_sheet_name: str = "EPI-Pregnancy",
) -> int:
    input_path = Path(input_file).resolve()
    output_path = Path(output_file).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if not input_path.exists():
        raise FileNotFoundError(f"Input workbook not found: {input_path}")

    src_wb = load_workbook(input_path, data_only=False)
    out_wb = Workbook()
    out_wb.remove(out_wb.active)

    if source_sheet_name in src_wb.sheetnames:
        src_ws = src_wb[source_sheet_name]
        target_ws = out_wb.create_sheet(title=target_sheet_name)
        for row in src_ws.iter_rows(values_only=False):
            target_ws.append([cell.value for cell in row])

        headers = [cell.value for cell in src_ws[1]]
        if "Age at first visit" not in headers:
            target_ws.cell(row=1, column=len(headers) + 1, value="Age at first visit")
        if "completion_status" not in headers:
            target_ws.cell(row=1, column=target_ws.max_column + 1, value="completion_status")
        if "source" not in headers:
            target_ws.cell(row=1, column=target_ws.max_column + 1, value="source")

    if pregnancy_source_sheet_name in src_wb.sheetnames:
        src_pg = src_wb[pregnancy_source_sheet_name]
        target_pg = out_wb.create_sheet(title=pregnancy_target_sheet_name)
        for row in src_pg.iter_rows(values_only=False):
            target_pg.append([cell.value for cell in row])

    out_wb.save(output_path)
    return 0


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Fallback cleaner for CHDN")
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    sys.exit(build_clean_sheet(args.input, args.output))
