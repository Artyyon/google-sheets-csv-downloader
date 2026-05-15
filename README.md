# Google Sheets CSV Downloader

![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Code Style](https://img.shields.io/badge/code%20style-black-000000)
![Lint](https://img.shields.io/badge/lint-ruff-red)
![Tests](https://img.shields.io/badge/tests-pytest-blueviolet)

Production-ready utility for downloading all worksheets from Google Sheets as CSV files using a robust and maintainable architecture.

---

# Features

- Automatic Google Sheets worksheet discovery
- Export all worksheets as CSV files
- Robust worksheet metadata extraction
- Cross-platform filename sanitization
- Structured logging
- Typed custom exceptions
- Production-oriented architecture
- Unit testing support
- Clean and maintainable codebase

---

# Why This Project Exists

Most Google Sheets export solutions found online rely on:

- fragile HTML scraping
- undocumented endpoints
- hardcoded worksheet names
- poor error handling

This project was designed to provide a more reliable and production-grade alternative using official Google Sheets export capabilities and clean software engineering practices.

---

# Installation

## Clone Repository

```bash
git clone https://github.com/Artyyon/google-sheets-csv-downloader.git
```

---

## Create Virtual Environment

### Windows

```bash
python -m venv venv
venv\Scripts\activate
```

### Linux / macOS

```bash
python3 -m venv venv
source venv/bin/activate
```

---

## Install Dependencies

```bash
pip install -r requirements.txt
```

---

# Usage

```python
from google_sheets_csv_downloader.downloader import (
    download_google_sheets
)

GOOGLE_SHEETS_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "YOUR_SHEET_ID/edit?usp=sharing"
)

download_google_sheets(
    sheet_url=GOOGLE_SHEETS_URL,
    output_dir="./data/"
)
```

---

# Example Output

```text
data/
├── Sales.csv
├── Customers.csv
├── Inventory.csv
└── Forecast.csv
```

---

# Project Structure

```text
google-sheets-csv-downloader/
│
├── google_sheets_csv_downloader/
│   ├── __init__.py
│   └── downloader.py
│
├── tests/
│
├── examples/
│
├── data/
│
├── README.md
├── LICENSE
├── requirements.txt
├── requirements-dev.txt
├── pyproject.toml
├── CONTRIBUTING.md
└── CHANGELOG.md
```

---

# Engineering Decisions

## Why XLSX Metadata Extraction?

Instead of relying on HTML parsing or undocumented APIs, the project uses the official XLSX export endpoint provided by Google Sheets.

This approach provides:

- better reliability
- long-term maintainability
- reduced breakage risk
- production-oriented stability

---

## Why `openpyxl`?

`openpyxl` allows reliable worksheet metadata extraction directly from workbook structure without depending on unstable frontend behavior.

---

# Development Workflow

This repository follows a professional Git workflow strategy:

- `main` → stable production branch
- `develop` → integration branch
- `feature/*` → feature development
- Pull Requests required for merges

---

# Running Tests

```bash
pytest
```

---

# Code Quality

## Ruff

```bash
ruff check .
```

## Black

```bash
black .
```

---

# Roadmap

Future improvements planned for this project:

- CLI interface
- Async downloads
- Retry strategies
- PyPI publishing
- GitHub Actions CI/CD
- Coverage reporting
- Docker support

---

# Contributing

Contributions, suggestions, and improvements are welcome.

Please open an issue or submit a pull request.

---

# License

This project is licensed under the MIT License.