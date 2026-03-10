# Excel to CSV Converter

This project provides a small, modular Python utility that converts an Excel workbook into multiple CSV files, one per sheet.

## Requirements

- Python 3.9+
- `pandas`
- `openpyxl`

Install dependencies:

```bash
pip install pandas openpyxl
```

## Usage

Basic usage (write CSVs to `./output`):

```bash
python3 excel_to_csv.py /path/to/workbook.xlsx
```

Specify an output directory:

```bash
python3 excel_to_csv.py /path/to/workbook.xlsx --out-dir /path/to/output
```

The script prints a mapping of original sheet names to the generated CSV files.

## How It Works

- Reads the entire workbook using `pandas.read_excel(..., sheet_name=None)` which returns a dictionary of `{sheet_name: DataFrame}`.
- For each sheet, creates a sanitized filename stem:
  - Lowercase
  - Remove spaces
  - Remove all non-alphanumeric characters
- Writes each sheet as a CSV (without index) into the chosen output directory. Default is `./output`.
- If two sheets sanitize to the same filename, it appends a numeric suffix (`_2`, `_3`, ...).

## Filename Rules

Given a sheet name, the output CSV filename is built from:

1. `lowercase(sheet_name)`
2. Remove all characters except `a-z` and `0-9`

Examples:

- `Sales Q1` -> `salesq1.csv`
- `Profit & Loss` -> `profitloss.csv`
- `  Summary  ` -> `summary.csv`
- `A/B` and `A B` would both become `ab.csv`; the second becomes `ab_2.csv`

## Project Structure

- `excel_to_csv.py` — main implementation
- `tests/` — unit tests (pytest)

## Running Tests

Install test dependencies:

```bash
pip install pytest
```

Run tests:

```bash
pytest
```
