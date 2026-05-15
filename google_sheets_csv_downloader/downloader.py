"""
============================================================
Google Sheets CSV Downloader
============================================================

A lightweight and robust utility for downloading all worksheets
from a public or shared Google Sheets document as individual
CSV files.

This module automatically:

- Extracts the spreadsheet ID from a Google Sheets URL
- Retrieves all worksheet names from the spreadsheet metadata
- Downloads each worksheet independently as CSV
- Sanitizes filenames for cross-platform compatibility
- Saves all worksheets locally

Architecture
------------
The implementation avoids fragile HTML scraping strategies and
instead relies on the official XLSX export endpoint provided
by Google Sheets.

Worksheet metadata is extracted using `openpyxl`, ensuring a
stable and production-grade approach for discovering worksheet
names.

Main Features
-------------
- Robust worksheet discovery
- Cross-platform filename sanitization
- Typed exceptions
- Clear separation of responsibilities
- Production-oriented error handling
- Minimal external dependencies

Requirements
------------
- pandas
- requests
- openpyxl

Example
-------
>>> GOOGLE_SHEETS_URL = (
...     "https://docs.google.com/spreadsheets/d/"
...     "YOUR_SHEET_ID/edit?usp=sharing"
... )
...
>>> download_google_sheets(
...     sheet_url=GOOGLE_SHEETS_URL,
...     output_dir="./data/"
... )

Author
------
Art

License
-------
MIT License
============================================================
"""

# =========================================================
# IMPORTS
# =========================================================

from io import BytesIO, StringIO
import logging
from pathlib import Path
import re
from urllib.parse import quote

import pandas as pd
import requests
from openpyxl import load_workbook
from requests.exceptions import RequestException


# =========================================================
# LOGGING CONFIGURATION
# =========================================================

logging.basicConfig(
    level=logging.INFO,
    format=(
        "%(asctime)s | "
        "%(levelname)s | "
        "%(message)s"
    )
)

logger = logging.getLogger(__name__)


# =========================================================
# CONFIGURATION
# =========================================================

REQUEST_TIMEOUT = 30

CSV_EXPORT_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
)

XLSX_EXPORT_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "{sheet_id}/export?format=xlsx"
)


# =========================================================
# CUSTOM EXCEPTIONS
# =========================================================

class GoogleSheetsError(Exception):
    """
    Base exception for all Google Sheets downloader errors.
    """


class InvalidGoogleSheetsURL(GoogleSheetsError):
    """
    Raised when the provided Google Sheets URL is invalid.
    """


class SheetMetadataError(GoogleSheetsError):
    """
    Raised when worksheet metadata cannot be retrieved.
    """


# =========================================================
# UTILITIES
# =========================================================

def extract_sheet_id(sheet_url: str) -> str:
    """
    Extract the Google Sheets document ID from a shared URL.

    Parameters
    ----------
    sheet_url : str
        Public or shared Google Sheets URL.

    Returns
    -------
    str
        Extracted spreadsheet identifier.

    Raises
    ------
    InvalidGoogleSheetsURL
        Raised when the provided URL does not contain a valid
        Google Sheets document identifier.

    Examples
    --------
    >>> extract_sheet_id(
    ...     "https://docs.google.com/spreadsheets/d/abc123/edit"
    ... )
    'abc123'
    """

    pattern = r"/spreadsheets/d/([a-zA-Z0-9-_]+)"

    match = re.search(pattern, sheet_url)

    if not match:
        raise InvalidGoogleSheetsURL(
            "Could not extract spreadsheet ID from URL."
        )

    return match.group(1)


def get_sheet_names(sheet_id: str) -> list[str]:
    """
    Retrieve all worksheet names from a Google Sheets document.

    Instead of relying on HTML scraping or undocumented Google
    endpoints, this implementation uses the official XLSX export
    functionality provided by Google Sheets.

    Parameters
    ----------
    sheet_id : str
        Unique Google Sheets document identifier.

    Returns
    -------
    list[str]
        Ordered list containing all worksheet names available
        in the spreadsheet.

    Raises
    ------
    SheetMetadataError
        Raised when worksheet metadata cannot be retrieved
        or interpreted.

    Notes
    -----
    Worksheet metadata is extracted directly from the XLSX
    workbook structure using `openpyxl`.
    """

    try:

        xlsx_url = XLSX_EXPORT_URL.format(
            sheet_id=sheet_id
        )

        logger.info(
            "Retrieving worksheet metadata..."
        )

        response = requests.get(
            xlsx_url,
            timeout=REQUEST_TIMEOUT
        )

        response.raise_for_status()

        workbook = load_workbook(
            filename=BytesIO(response.content),
            read_only=True
        )

        sheet_names = workbook.sheetnames

        if not sheet_names:
            raise SheetMetadataError(
                "No worksheets found in spreadsheet."
            )

        logger.info(
            "Found %d worksheet(s).",
            len(sheet_names)
        )

        return sheet_names

    except RequestException as error:
        raise SheetMetadataError(
            f"Connection error while retrieving "
            f"worksheet metadata: {error}"
        ) from error

    except Exception as error:
        raise SheetMetadataError(
            f"Unexpected error while retrieving "
            f"worksheet metadata: {error}"
        ) from error


def sanitize_filename(filename: str) -> str:
    """
    Sanitize filenames for safe cross-platform file storage.

    Invalid filesystem characters are replaced with underscores.

    Parameters
    ----------
    filename : str
        Original filename.

    Returns
    -------
    str
        Sanitized filename.
    """

    return re.sub(
        r'[\\/*?:"<>|]',
        "_",
        filename
    )


# =========================================================
# CSV DOWNLOAD
# =========================================================

def download_sheet_as_csv(
    sheet_id: str,
    sheet_name: str,
    output_dir: Path
) -> None:
    """
    Download a single worksheet as a CSV file.

    Parameters
    ----------
    sheet_id : str
        Unique Google Sheets document identifier.

    sheet_name : str
        Worksheet name.

    output_dir : Path
        Directory where the CSV file will be saved.

    Raises
    ------
    RequestException
        Raised when a network-related error occurs.

    pandas.errors.ParserError
        Raised when the CSV content cannot be parsed.
    """

    try:

        logger.info(
            "Downloading worksheet: %s",
            sheet_name
        )

        encoded_sheet_name = quote(sheet_name)

        csv_url = CSV_EXPORT_URL.format(
            sheet_id=sheet_id,
            sheet_name=encoded_sheet_name
        )

        response = requests.get(
            csv_url,
            timeout=REQUEST_TIMEOUT
        )

        response.raise_for_status()

        response.encoding = "utf-8"

        sheet_df = pd.read_csv(
            StringIO(response.text)
        )

        safe_filename = sanitize_filename(
            sheet_name
        )

        output_file = (
            output_dir /
            f"{safe_filename}.csv"
        )

        sheet_df.to_csv(
            output_file,
            index=False,
            encoding="utf-8"
        )

        logger.info(
            "Worksheet saved: %s",
            output_file
        )

    except pd.errors.ParserError as error:

        logger.error(
            "CSV parsing error in worksheet '%s': %s",
            sheet_name,
            error
        )

    except RequestException as error:

        logger.error(
            "Connection error while downloading "
            "worksheet '%s': %s",
            sheet_name,
            error
        )

    except Exception as error:

        logger.error(
            "Unexpected error while downloading "
            "worksheet '%s': %s",
            sheet_name,
            error
        )


# =========================================================
# MAIN WORKFLOW
# =========================================================

def download_google_sheets(
    sheet_url: str,
    output_dir: str = "./data/"
) -> None:
    """
    Download all worksheets from a Google Sheets document
    as CSV files.

    Parameters
    ----------
    sheet_url : str
        Public or shared Google Sheets URL.

    output_dir : str, optional
        Output directory where CSV files will be stored.

    Raises
    ------
    GoogleSheetsError
        Raised when any Google Sheets-specific operation fails.

    Notes
    -----
    Each worksheet is downloaded independently, allowing
    partial success even if one worksheet fails.
    """

    logger.info(
        "=" * 60
    )

    logger.info(
        "STARTING GOOGLE SHEETS DOWNLOAD"
    )

    logger.info(
        "=" * 60
    )

    try:

        # -------------------------------------------------
        # Output directory
        # -------------------------------------------------

        output_path = Path(output_dir)

        output_path.mkdir(
            parents=True,
            exist_ok=True
        )

        logger.info(
            "Output directory: %s",
            output_path.resolve()
        )

        # -------------------------------------------------
        # Spreadsheet ID
        # -------------------------------------------------

        sheet_id = extract_sheet_id(
            sheet_url
        )

        logger.info(
            "Spreadsheet ID: %s",
            sheet_id
        )

        # -------------------------------------------------
        # Worksheet discovery
        # -------------------------------------------------

        sheet_names = get_sheet_names(
            sheet_id
        )

        logger.info(
            "Worksheets discovered:"
        )

        for index, name in enumerate(
            sheet_names,
            start=1
        ):

            logger.info(
                "  %d. %s",
                index,
                name
            )

        # -------------------------------------------------
        # Worksheet downloads
        # -------------------------------------------------

        for sheet_name in sheet_names:

            download_sheet_as_csv(
                sheet_id=sheet_id,
                sheet_name=sheet_name,
                output_dir=output_path
            )

        logger.info(
            "Download process completed successfully."
        )

    except GoogleSheetsError as error:

        logger.error(
            "Google Sheets error: %s",
            error
        )

    except Exception as error:

        logger.exception(
            "Unexpected fatal error: %s",
            error
        )


# =========================================================
# EXAMPLE USAGE
# =========================================================

if __name__ == "__main__":

    GOOGLE_SHEETS_URL = (
        "https://docs.google.com/spreadsheets/d/"
        "1hOV1C_c0RujZqzl7FMsdY-uESj72Yd8bxBJVRAF1p5Q/"
        "edit?usp=sharing"
    )

    download_google_sheets(
        sheet_url=GOOGLE_SHEETS_URL,
        output_dir="./data/"
    )