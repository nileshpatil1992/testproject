from __future__ import annotations

from pathlib import Path

import pandas as pd

from csv_to_excel import csvs_to_excel, sanitize_sheet_name


def test_sanitize_sheet_name():
    assert sanitize_sheet_name("salesq1") == "salesq1"
    assert sanitize_sheet_name("profit:loss") == "profit_loss"
    assert sanitize_sheet_name("a/b") == "a_b"
    assert sanitize_sheet_name("" * 5) == "sheet"


def test_csvs_to_excel(tmp_path: Path):
    input_dir = tmp_path / "output"
    input_dir.mkdir()

    pd.DataFrame({"a": [1, 2]}).to_csv(input_dir / "salesq1.csv", index=False)
    pd.DataFrame({"b": ["x", "y"]}).to_csv(input_dir / "profit-loss.csv", index=False)

    out_file = tmp_path / "output_excel" / "workbook.xlsx"
    written = csvs_to_excel(input_dir, out_file)

    assert set(written.keys()) == {"salesq1.csv", "profit-loss.csv"}
    assert out_file.exists()

    sheets = pd.read_excel(out_file, sheet_name=None)
    assert set(sheets.keys()) == {"salesq1", "profit-loss"}


def test_csvs_to_excel_deduped(tmp_path: Path):
    input_dir = tmp_path / "output"
    input_dir.mkdir()

    pd.DataFrame({"a": [1]}).to_csv(input_dir / "a?b.csv", index=False)
    pd.DataFrame({"a": [2]}).to_csv(input_dir / "a*b.csv", index=False)

    out_file = tmp_path / "output_excel" / "workbook.xlsx"
    written = csvs_to_excel(input_dir, out_file)

    assert set(written.keys()) == {"a?b.csv", "a*b.csv"}

    sheets = pd.read_excel(out_file, sheet_name=None)
    assert set(sheets.keys()) == {"a_b", "a_b_2"}
