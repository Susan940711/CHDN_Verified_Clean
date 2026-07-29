from __future__ import annotations

import argparse
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

DEFAULT_INPUT = Path(__file__).resolve().with_name("EPI Database_CHDN.xlsx")
DEFAULT_OUTPUT = Path(__file__).resolve().with_name("CHDN_EPI_clean.xlsx")
DEFAULT_SOURCE_SHEET = "EPI-Child"
DEFAULT_TARGET_SHEET = "EPI-Child"
DEFAULT_PREGNANCY_SOURCE_SHEET = "EPI-Pregnancy"
DEFAULT_PREGNANCY_TARGET_SHEET = "EPI-Pregnancy"

COMPLETE_COLUMNS = [
    "CompleteInQ4 2024",
    "CompleteInQ1 2025",
    "CompleteInQ2 2025",
    "CompleteInQ3 2025",
    "CompleteInQ4 2025",
    "CompleteInQ1 2026",
    "CompleteInQ2 2026",
    "CompleteInQ3 2026",
    "CompleteInQ4 2026",
]

QUARTER_WINDOWS = {
    "CompleteInQ4 2024": ("Q4", date(2024, 9, 21), date(2024, 12, 20)),
    "CompleteInQ1 2025": ("Q1", date(2024, 12, 21), date(2025, 3, 20)),
    "CompleteInQ2 2025": ("Q2", date(2025, 3, 21), date(2025, 6, 20)),
    "CompleteInQ3 2025": ("Q3", date(2025, 6, 21), date(2025, 9, 20)),
    "CompleteInQ4 2025": ("Q4", date(2025, 9, 21), date(2025, 12, 20)),
    "CompleteInQ1 2026": ("Q1", date(2025, 12, 21), date(2026, 3, 20)),
    "CompleteInQ2 2026": ("Q2", date(2026, 3, 21), date(2026, 6, 20)),
    "CompleteInQ3 2026": ("Q3", date(2026, 6, 21), date(2026, 9, 20)),
    "CompleteInQ4 2026": ("Q4", date(2026, 9, 21), date(2026, 12, 20)),
}


def _normalize_col_name(name: str) -> str:
    return "".join(ch for ch in str(name).strip().lower() if ch.isalnum())


def _find_children_code_column(columns) -> int | None:
    aliases = {"childrencodetccode", "childrencode", "tccode", "tcode"}
    for idx, col in enumerate(columns, start=1):
        if _normalize_col_name(col) in aliases:
            return idx
    return None


def _find_pregnance_code_column(columns) -> int | None:
    aliases = {"pregnancecode", "pregnancycode", "pregnance_code", "pregnancy_code", "pwcode", "pw_code"}
    norm_aliases = {a.replace("_", "") for a in aliases}
    for idx, col in enumerate(columns, start=1):
        if _normalize_col_name(col) in norm_aliases:
            return idx
    return None


def _is_blank(value) -> bool:
    return value is None or str(value).strip() == ""


def _to_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        value = value.strip()
        if not value:
            return None
        parsed = pd.to_datetime(value, errors="coerce")
        if pd.isna(parsed):
            return None
        return parsed.date()
    if isinstance(value, (int, float)):
        try:
            parsed = pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
            if pd.isna(parsed):
                return None
            return parsed.date()
        except Exception:
            return None
    return None


def _find_header_index(header_index_map: dict[str, int], aliases: list[str]) -> int | None:
    for alias in aliases:
        if alias in header_index_map:
            return header_index_map[alias]
    normalized_aliases = {_normalize_col_name(alias) for alias in aliases}
    for header_name, index in header_index_map.items():
        if _normalize_col_name(header_name) in normalized_aliases:
            return index
    return None


def _row_value(row_values, header_index_map: dict[str, int], aliases: list[str]):
    index = _find_header_index(header_index_map, aliases)
    if index is None:
        return ""
    return row_values[index]


def _presence_ok(row_values, header_index_map: dict[str, int], dose_aliases: list[str], other_aliases: list[str]) -> bool:
    dose_value = _row_value(row_values, header_index_map, dose_aliases)
    other_value = _row_value(row_values, header_index_map, other_aliases)
    if not _is_blank(dose_value):
        return True
    if not _is_blank(other_value):
        return True
    return False


def _date_source_ok(row_values, header_index_map: dict[str, int], date_aliases: list[str], source_aliases: list[str], start: date, end: date) -> bool:
    row_date = _to_date(_row_value(row_values, header_index_map, date_aliases))
    if row_date is None or not (start <= row_date <= end):
        return False
    source_value = str(_row_value(row_values, header_index_map, source_aliases)).strip().upper()
    if source_value in {"CHDN", "CHDN/OTHER", "OTHER", ""}:
        return True
    return False


def _completion_value(row_values, column_name: str, header_index_map: dict[str, int]) -> str:
    quarter_label, start_date, end_date = QUARTER_WINDOWS[column_name]
    year_label = column_name.rsplit(" ", 1)[-1]

    u1_presence_rules = [
        (["BCG dose", "BCG"], ["BCG other (Y/N)", "BCG other", "BCG other yn"]),
        (["OPV first time dose", "OPV first timeadose", "OPV1 dose", "OPV1"], ["OPV first time dose other", "OPV1 other", "OPV1 other yn"]),
        (["OPV second time dose", "OPV2 dose", "OPV2"], ["OPV second time dose other", "OPV2 other", "OPV2 other yn"]),
        (["OPV third time dose", "OPV3 dose", "OPV3"], ["OPV third time dose other", "OPV3 other", "OPV3 other yn"]),
        (["Penta first time dose", "Penta1 dose", "Penta1"], ["Penta first time dose other", "Penta1 other", "Penta1 other yn"]),
        (["Penta second time dose", "Penta2 dose", "Penta2"], ["Penta second time dose other", "Penta2 other", "Penta2 other yn"]),
        (["Penta third time dose", "Penta3 dose", "Penta3"], ["Penta third time dose other", "Penta3 other", "Penta3 other yn"]),
        (["MMR first time dose", "MMR first timeadose", "MMR1 dose", "MMR1"], ["MMR first time dose other", "MMR1 other", "MMR1 other yn"]),
    ]
    u5_presence_rules = u1_presence_rules[1:]

    u1_date_source_rules = [
        (["BCG reporting month", "BCG report month"], ["BCG Source", "BCG source", "BCG src"]),
        (["OPV first time dose reporting month", "OPV1 reporting month"], ["OPV1 Source", "OPV1 source", "OPV1 src"]),
        (["OPV second time dose reporting month", "OPV2 reporting month"], ["OPV2 Source", "OPV2 source", "OPV2 src"]),
        (["OPV third time dose reporting month", "OPV3 reporting month"], ["OPV3 Source", "OPV3 source", "OPV3 src"]),
        (["Penta first time dose reporting month", "Penta1 reporting month"], ["Penta1 Source", "Penta1 source", "Penta1 src"]),
        (["Penta second time dose reporting month", "Penta2 reporting month"], ["Penta2 Source", "Penta2 source", "Penta2 src"]),
        (["Penta third time dose reporting month", "Penta3 reporting month"], ["Penta3 Source", "Penta3 source", "Penta3 src"]),
        (["MMR first time dose reporting month", "MMR1 reporting month"], ["MMR1 Source", "MMR1 source", "MMR1 src"]),
    ]
    u5_date_source_rules = u1_date_source_rules[1:]

    age_value = _row_value(row_values, header_index_map, ["Age at first visit", "Age", "Age in months", "Age at first visit "])
    try:
        age_num = float(age_value)
    except (TypeError, ValueError):
        age_num = None

    if age_num is not None and age_num <= 11:
        has_any_presence = any(_presence_ok(row_values, header_index_map, dose_aliases, other_aliases) for dose_aliases, other_aliases in u1_presence_rules)
        has_any_date_source = any(_date_source_ok(row_values, header_index_map, date_aliases, source_aliases, start_date, end_date) for date_aliases, source_aliases in u1_date_source_rules)
        if has_any_presence and has_any_date_source:
            return f"U1 complete in {quarter_label}_{year_label}"

    if age_num is not None and 11 < age_num <= 59:
        has_any_presence = any(_presence_ok(row_values, header_index_map, dose_aliases, other_aliases) for dose_aliases, other_aliases in u5_presence_rules)
        has_any_date_source = any(_date_source_ok(row_values, header_index_map, date_aliases, source_aliases, start_date, end_date) for date_aliases, source_aliases in u5_date_source_rules)
        if has_any_presence and has_any_date_source:
            return f"1-5 complete in {quarter_label}_{year_label}"

    return ""


def _warn_code_quality(sheet_or_values, code_col_idx_or_label, code_label: str | None = None, row_numbers: list[int] | None = None) -> None:
    if code_label is None and row_numbers is None:
        code_values = sheet_or_values
        label = code_col_idx_or_label
    else:
        sheet = sheet_or_values
        code_col_idx = code_col_idx_or_label
        code_values = [sheet.cell(row=row_num, column=code_col_idx).value for row_num in row_numbers or []]
        label = code_label

    normalized_codes = [str(value).strip() for value in code_values if not _is_blank(value)]
    missing_count = sum(1 for value in code_values if _is_blank(value))
    duplicate_codes = sorted({value for value in normalized_codes if normalized_codes.count(value) > 1})

    if missing_count > 0:
        print(f"WARNING: {label} missing. Missing count = {missing_count}")
    else:
        print(f"No missing {label}")
    if duplicate_codes:
        print(f"WARNING: duplication of {label} ({','.join(duplicate_codes)})")


def _get_kept_rows(values_sheet, data_start_row: int, max_col: int) -> tuple[list[int], int]:
    kept_rows: list[int] = []
    removed_count = 0
    max_row = values_sheet.max_row
    for row_num, row in enumerate(values_sheet.iter_rows(min_row=data_start_row, max_row=max_row, values_only=False), start=data_start_row):
        row_values = [cell.value for cell in row[:max_col]]
        if any(not _is_blank(value) for value in row_values):
            kept_rows.append(row_num)
        else:
            removed_count += 1
    return kept_rows, removed_count


def _write_sheet_values(src_ws, dst_ws) -> None:
    for row_num in range(1, src_ws.max_row + 1):
        for col_idx in range(1, src_ws.max_column + 1):
            dst_ws.cell(row=row_num, column=col_idx, value=src_ws.cell(row=row_num, column=col_idx).value)


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

    wb = load_workbook(input_path, data_only=True, read_only=True)
    if source_sheet_name not in wb.sheetnames:
        raise ValueError(f"Source sheet not found: {source_sheet_name}")

    child_ws = wb[source_sheet_name]
    max_col = child_ws.max_column
    source_headers = next(child_ws.iter_rows(min_row=1, max_row=1, values_only=True))
    source_headers = list(source_headers[:max_col])
    code_col_idx = _find_children_code_column(source_headers)
    if code_col_idx is None:
        raise ValueError("Could not find children_code(T/C Code) column in the source sheet")

    child_rows_to_keep: list[int] = []
    child_code_values: list[object] = []
    child_rows_data: list[list[object]] = []
    child_removed_count = 0

    for row_num, row in enumerate(child_ws.iter_rows(min_row=2, max_row=child_ws.max_row, values_only=True), start=2):
        row_values = list(row[:max_col])
        if any(not _is_blank(value) for value in row_values):
            child_rows_to_keep.append(row_num)
            child_code_values.append(row_values[code_col_idx - 1])
            child_rows_data.append(row_values)
        else:
            child_removed_count += 1

    if child_removed_count > 0:
        print(f"INFO: Removed {child_removed_count} formula-only rows from {source_sheet_name}")

    _warn_code_quality(child_code_values, "children_code")

    out_wb = Workbook(write_only=True)
    child_out_ws = out_wb.create_sheet(title=target_sheet_name)
    child_out_ws.append(list(source_headers) + list(COMPLETE_COLUMNS))

    header_index_map = {header_name: idx for idx, header_name in enumerate(source_headers)}

    for row_values in child_rows_data:
        completed_values = {col_name: _completion_value(row_values, col_name, header_index_map) for col_name in COMPLETE_COLUMNS}
        child_out_ws.append(row_values + [completed_values[col_name] for col_name in COMPLETE_COLUMNS])

    if pregnancy_source_sheet_name in wb.sheetnames:
        preg_ws = wb[pregnancy_source_sheet_name]
        preg_max_col = preg_ws.max_column
        preg_header_row = 2
        preg_headers = [preg_ws.cell(row=preg_header_row, column=idx).value for idx in range(1, preg_max_col + 1)]

        preg_rows_to_keep, preg_removed_count = _get_kept_rows(preg_ws, data_start_row=3, max_col=preg_max_col)
        if preg_removed_count > 0:
            print(f"INFO: Removed {preg_removed_count} formula-only rows from {pregnancy_source_sheet_name}")

        preg_code_col_idx = _find_pregnance_code_column(preg_headers)
        if preg_code_col_idx is None:
            print("WARNING: Could not find Pregnance_code column in EPI-Pregnancy sheet.")
        else:
            _warn_code_quality(preg_ws, preg_code_col_idx, "pw_code", preg_rows_to_keep)

        preg_out_ws = out_wb.create_sheet(title=pregnancy_target_sheet_name)
        preg_header_values = [preg_ws.cell(row=preg_header_row, column=col_idx).value for col_idx in range(1, preg_max_col + 1)]
        preg_out_ws.append(preg_header_values)

        for row_num in preg_rows_to_keep:
            row_values = [preg_ws.cell(row=row_num, column=col_idx).value for col_idx in range(1, preg_max_col + 1)]
            preg_out_ws.append(row_values)
    else:
        print(f"WARNING: Pregnancy source sheet not found: {pregnancy_source_sheet_name}")

    try:
        out_wb.save(output_path)
    except Exception as exc:
        raise RuntimeError(f"Failed to save workbook: {exc}") from exc

    print(f"Done: created {output_path.name}")
    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Create CHDN_EPI_clean.xlsx with completion columns")
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    sys.exit(build_clean_sheet(args.input, args.output))
