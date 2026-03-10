#!/usr/bin/env python3
"""Convert CSV files in a folder into a single Excel workbook.

Rules:
- Each CSV becomes a separate sheet.
- Sheet name is based on the CSV filename (without extension).
- Output workbook defaults to ./output_excel/workbook.xlsx
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Dict, Iterable

import pandas as pd


_INVALID_SHEET_RE = re.compile(r"[\\/*?:\[\]]")


def sanitize_sheet_name(name: str) -> str:
    """Return a safe Excel sheet name derived from a filename stem."""
    cleaned = _INVALID_SHEET_RE.sub("_", name.strip())
    cleaned = cleaned[:31]
    return cleaned or "sheet"


def _dedupe_name(stem: str, seen: Dict[str, int]) -> str:
    if stem not in seen:
        seen[stem] = 1
        return stem
    seen[stem] += 1
    return f"{stem}_{seen[stem]}"


def _iter_csv_files(input_dir: Path) -> Iterable[Path]:
    return sorted(p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() == ".csv")


def csvs_to_excel(input_dir: Path, output_file: Path) -> Dict[str, str]:
    """Convert all CSVs in input_dir to a single Excel workbook.

    Returns a mapping of csv filename to sheet name.
    """
    output_file.parent.mkdir(parents=True, exist_ok=True)

    csv_files = list(_iter_csv_files(input_dir))
    if not csv_files:
        raise ValueError(f"No CSV files found in: {input_dir}")

    written: Dict[str, str] = {}
    seen: Dict[str, int] = {}

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for csv_path in csv_files:
            df = pd.read_csv(csv_path)
            stem = sanitize_sheet_name(csv_path.stem)
            stem = _dedupe_name(stem, seen)
            df.to_excel(writer, sheet_name=stem, index=False)
            written[csv_path.name] = stem

    return written


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert CSVs in a folder to a single Excel workbook.")
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=Path.cwd() / "output",
        help="Directory containing CSV files (default: ./output)",
    )
    parser.add_argument(
        "--out-file",
        type=Path,
        default=Path.cwd() / "output_excel" / "workbook.xlsx",
        help="Output Excel file path (default: ./output_excel/workbook.xlsx)",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_dir: Path = args.input_dir
    if not input_dir.exists():
        raise SystemExit(f"Input directory not found: {input_dir}")

    written = csvs_to_excel(input_dir, args.out_file)
    for csv_name, sheet in written.items():
        print(f"{csv_name} -> {sheet}")
    print(f"Wrote: {args.out_file}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
