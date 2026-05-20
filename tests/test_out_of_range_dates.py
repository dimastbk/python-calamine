"""
Tests verifying that date/datetime cells whose year falls outside Python's
datetime range are returned as ISO 8601 strings rather than raising an error.

Boundary: year <= 1000 or year > 9999 → str; otherwise → date/datetime.

Fixtures:
  out_of_range_dates.xlsx — two sheets with future date serials (year 10000)
  out_of_range_dates.ods  — two sheets with past ISO dates (year 500)
"""

from pathlib import Path

from python_calamine import CalamineWorkbook

PATH = Path(__file__).parent / "data"
XLSX = PATH / "out_of_range_dates.xlsx"
ODS = PATH / "out_of_range_dates.ods"


# ---------------------------------------------------------------------------
# xlsx — future dates (serial 2958466+ maps to year 10000)
# ---------------------------------------------------------------------------


def test_xlsx_future_date_returns_string():
    sheet = CalamineWorkbook.from_object(XLSX).get_sheet_by_name("future_date")
    rows = sheet.to_python()
    cell = rows[0][0]
    assert isinstance(cell, str), f"expected str for out-of-range date, got {type(cell)}: {cell!r}"


def test_xlsx_future_date_string_is_iso_format():
    sheet = CalamineWorkbook.from_object(XLSX).get_sheet_by_name("future_date")
    cell = sheet.to_python()[0][0]
    # chrono formats with %Y-%m-%d; year 10000 gives "10000-01-01" or similar
    parts = cell.split("-")
    assert len(parts) == 3, f"expected YYYY-MM-DD, got {cell!r}"
    assert int(parts[0]) > 9999, f"expected year > 9999, got {cell!r}"


def test_xlsx_future_datetime_returns_string():
    sheet = CalamineWorkbook.from_object(XLSX).get_sheet_by_name("future_datetime")
    rows = sheet.to_python()
    cell = rows[0][0]
    assert isinstance(cell, str), f"expected str for out-of-range datetime, got {type(cell)}: {cell!r}"


def test_xlsx_future_datetime_string_contains_time_component():
    sheet = CalamineWorkbook.from_object(XLSX).get_sheet_by_name("future_datetime")
    cell = sheet.to_python()[0][0]
    assert "T" in cell, f"expected ISO datetime with 'T' separator, got {cell!r}"
    date_part = cell.split("T")[0]
    year = int(date_part.split("-")[0])
    assert year > 9999, f"expected year > 9999 in {cell!r}"


# ---------------------------------------------------------------------------
# ODS — past dates (ISO value "0500-06-15", year 500)
# ---------------------------------------------------------------------------


def test_ods_past_date_returns_string():
    sheet = CalamineWorkbook.from_object(ODS).get_sheet_by_name("past_date")
    rows = sheet.to_python()
    cell = rows[0][0]
    assert isinstance(cell, str), f"expected str for out-of-range date, got {type(cell)}: {cell!r}"


def test_ods_past_date_string_is_iso_format():
    sheet = CalamineWorkbook.from_object(ODS).get_sheet_by_name("past_date")
    cell = sheet.to_python()[0][0]
    parts = cell.split("-")
    assert len(parts) == 3, f"expected YYYY-MM-DD, got {cell!r}"
    assert int(parts[0]) <= 1000, f"expected year <= 1000, got {cell!r}"


def test_ods_past_datetime_returns_string():
    sheet = CalamineWorkbook.from_object(ODS).get_sheet_by_name("past_datetime")
    rows = sheet.to_python()
    cell = rows[0][0]
    assert isinstance(cell, str), f"expected str for out-of-range datetime, got {type(cell)}: {cell!r}"


def test_ods_past_datetime_string_contains_time_component():
    sheet = CalamineWorkbook.from_object(ODS).get_sheet_by_name("past_datetime")
    cell = sheet.to_python()[0][0]
    assert "T" in cell, f"expected ISO datetime with 'T' separator, got {cell!r}"
    date_part = cell.split("T")[0]
    year = int(date_part.split("-")[0])
    assert year <= 1000, f"expected year <= 1000 in {cell!r}"
