from google_sheets_csv_downloader.downloader import extract_sheet_id


def test_extract_sheet_id():

    url = "https://docs.google.com/spreadsheets/d/" "abc123/edit"

    assert extract_sheet_id(url) == "abc123"
