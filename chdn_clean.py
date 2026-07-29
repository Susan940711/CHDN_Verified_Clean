from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[2]
CLEANER_SCRIPT = ROOT_DIR / "data" / "CHDN_clean" / "CHDN_EPI_clean.py"


def _load_cleaner_module():
    if not CLEANER_SCRIPT.exists():
        raise FileNotFoundError(f"Cleaner script not found: {CLEANER_SCRIPT}")

    spec = importlib.util.spec_from_file_location("chdn_clean_stage1", CLEANER_SCRIPT)
    if spec is None or spec.loader is None:
        raise ImportError(f"Unable to load cleaner module from {CLEANER_SCRIPT}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


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
