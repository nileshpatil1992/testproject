# Excel to CSV Converter

This project provides small, modular Python utilities to convert:
- Excel workbooks into multiple CSVs (one per sheet)
- CSV files into a single Excel workbook (one sheet per CSV)

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

### CSV to Excel

Basic usage (read CSVs from `./output`, write to `./output_excel/workbook.xlsx`):

```bash
python3 csv_to_excel.py
```

Specify input directory and output file:

```bash
python3 csv_to_excel.py --input-dir /path/to/csvs --out-file /path/to/output.xlsx
```

## How It Works

- Reads the entire workbook using `pandas.read_excel(..., sheet_name=None)` which returns a dictionary of `{sheet_name: DataFrame}`.
- For each sheet, creates a sanitized filename stem:
  - Lowercase
  - Remove spaces
  - Remove all non-alphanumeric characters
- Writes each sheet as a CSV (without index) into the chosen output directory. Default is `./output`.
- If two sheets sanitize to the same filename, it appends a numeric suffix (`_2`, `_3`, ...).

### CSV to Excel

- Scans the input directory for `.csv` files and sorts them by filename.
- Uses each CSV filename (without extension) as the sheet name.
- Sanitizes sheet names to fit Excel limits and removes invalid characters.
- Writes all sheets to a single workbook at `./output_excel/workbook.xlsx` by default.

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
- `csv_to_excel.py` — CSV to Excel implementation
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
