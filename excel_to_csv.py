#!/usr/bin/env python3
"""Convert an Excel workbook into CSV files, one per sheet.

Rules:
- Each sheet becomes a separate CSV.
- Output filename is the sheet name, lowercased, with spaces and special chars removed.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Dict

import pandas as pd


_SANITIZE_RE = re.compile(r"[^a-z0-9]")


def sanitize_sheet_name(name: str) -> str:
    """Return a lowercase, alphanumeric-only filename stem for a sheet name."""
    lowered = name.strip().lower()
    sanitized = _SANITIZE_RE.sub("", lowered)
    return sanitized or "sheet"


def _dedupe_name(stem: str, seen: Dict[str, int]) -> str:
    """Ensure unique filename stem by suffixing with _N if needed."""
    if stem not in seen:
        seen[stem] = 1
        return stem
    seen[stem] += 1
    return f"{stem}_{seen[stem]}"


def excel_to_csvs(excel_path: Path, output_dir: Path) -> Dict[str, Path]:
    """Convert all sheets in an Excel file to CSVs.

    Returns a mapping of sheet name to output CSV path.
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    sheets = pd.read_excel(excel_path, sheet_name=None, engine="openpyxl")
    written: Dict[str, Path] = {}
    seen: Dict[str, int] = {}

    for sheet_name, df in sheets.items():
        stem = sanitize_sheet_name(sheet_name)
        stem = _dedupe_name(stem, seen)
        out_path = output_dir / f"{stem}.csv"
        df.to_csv(out_path, index=False)
        written[sheet_name] = out_path

    return written


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert Excel sheets to CSV files.")
    parser.add_argument("excel", type=Path, help="Path to .xlsx file")
    parser.add_argument(
        "--out-dir",
        type=Path,
        default=Path.cwd() / "output",
        help="Output directory for CSV files (default: ./output)",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    excel_path: Path = args.excel
    if not excel_path.exists():
        raise SystemExit(f"Excel file not found: {excel_path}")

    written = excel_to_csvs(excel_path, args.out_dir)
    for sheet_name, path in written.items():
        print(f"{sheet_name} -> {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
