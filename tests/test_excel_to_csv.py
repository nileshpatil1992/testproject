from __future__ import annotations

from pathlib import Path

import pandas as pd

from excel_to_csv import excel_to_csvs, sanitize_sheet_name


FIXTURE = Path(__file__).resolve().parents[1] / "sample_workbook.xlsx"


def test_sanitize_sheet_name():
    assert sanitize_sheet_name("Sales Q1") == "salesq1"
    assert sanitize_sheet_name(" Profit & Loss ") == "profitloss"
    assert sanitize_sheet_name("A/B") == "ab"
    assert sanitize_sheet_name("   ") == "sheet"


def test_excel_to_csvs_multiple_sheets(tmp_path: Path):
    excel_path = tmp_path / "workbook.xlsx"

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        pd.DataFrame({"a": [1, 2]}).to_excel(writer, sheet_name="Sales Q1", index=False)
        pd.DataFrame({"b": ["x", "y"]}).to_excel(writer, sheet_name="Profit & Loss", index=False)

    out_dir = tmp_path / "output"
    written = excel_to_csvs(excel_path, out_dir)

    assert set(written.keys()) == {"Sales Q1", "Profit & Loss"}
    assert (out_dir / "salesq1.csv").exists()
    assert (out_dir / "profitloss.csv").exists()


def test_excel_to_csvs_deduped_names(tmp_path: Path):
    excel_path = tmp_path / "workbook.xlsx"

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        pd.DataFrame({"a": [1]}).to_excel(writer, sheet_name="A B", index=False)
        pd.DataFrame({"a": [2]}).to_excel(writer, sheet_name="A-B", index=False)

    out_dir = tmp_path / "output"
    written = excel_to_csvs(excel_path, out_dir)

    assert set(written.keys()) == {"A B", "A-B"}
    assert (out_dir / "ab.csv").exists()
    assert (out_dir / "ab_2.csv").exists()


def test_excel_to_csvs_with_fixture(tmp_path: Path):
    out_dir = tmp_path / "output"
    written = excel_to_csvs(FIXTURE, out_dir)

    assert set(written.keys()) == {"Sales Q1", "Profit & Loss", "A B", "A-B"}
    assert (out_dir / "salesq1.csv").exists()
    assert (out_dir / "profitloss.csv").exists()
    assert (out_dir / "ab.csv").exists()
    assert (out_dir / "ab_2.csv").exists()
