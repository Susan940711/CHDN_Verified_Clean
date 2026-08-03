from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path


def _find_file(start: Path, filename: str) -> Path | None:
    for base in [start, *start.parents]:
        if not base.exists():
            continue
        for path in base.rglob(filename):
            if path.is_file():
                return path.resolve()
    return None


def _load_cleaner_module():
    candidate_paths = [
        Path(__file__).resolve().parent / "CHDN_EPI_Report" / "CHDN_EPI_clean.py",
        Path(__file__).resolve().parent / "CHDN_EPI_clean.py",
    ]
    for cleaner_script in candidate_paths:
        if not cleaner_script.exists():
            continue
        spec = importlib.util.spec_from_file_location("chdn_clean_stage1", cleaner_script)
        if spec is None or spec.loader is None:
            continue
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    raise FileNotFoundError("Cleaner script not found. Searched the bundled app folder.")


def run_clean(
    input_path: Path | str,
    output_path: Path | str,
    source_sheet_name: str = "EPI-Child",
    target_sheet_name: str = "EPI-Child",
    pregnancy_source_sheet_name: str = "EPI-Pregnancy",
    pregnancy_target_sheet_name: str = "EPI-Pregnancy",
) -> int:
    input_path = Path(input_path).resolve()
    output_path = Path(output_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if not input_path.exists():
        raise FileNotFoundError(f"Input workbook not found: {input_path}")

    cleaner_module = _load_cleaner_module()
    return cleaner_module.build_clean_sheet(
        input_file=input_path,
        output_file=output_path,
        source_sheet_name=source_sheet_name,
        target_sheet_name=target_sheet_name,
        pregnancy_source_sheet_name=pregnancy_source_sheet_name,
        pregnancy_target_sheet_name=pregnancy_target_sheet_name,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Run only the CHDN cleaning step")
    parser.add_argument("--input", required=True, help="Path to the source Excel workbook")
    parser.add_argument("--output", required=True, help="Path to write the cleaned workbook")
    args = parser.parse_args()

    try:
        return run_clean(args.input, args.output)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
