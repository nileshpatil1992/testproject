#!/usr/bin/env bash
set -euo pipefail

python3 -m pip install --upgrade pip >/dev/null 2>&1 || true
python3 -m pip install pandas openpyxl pytest

python3 -m pytest
