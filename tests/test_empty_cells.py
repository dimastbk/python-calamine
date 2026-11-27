from pathlib import Path

import pytest
from python_calamine import CalamineWorkbook

DATA = Path(__file__).parent / "data"


@pytest.mark.parametrize("filename", ["empty-string.xls", "empty-string.xlsx"])
def test_empty_cell_and_empty_formula_result(filename):
    sheet = CalamineWorkbook.from_path(DATA / filename).get_sheet_by_index(0)
    expected = [["", "x"], [None, "x"]]
    default = [["", "x"], ["", "x"]]
    if filename.endswith(".xlsx"):
        # Error cells still use the legacy conversion, regardless of the option.
        expected = [["", "x", ""], [None, "x", None]]
        default = [["", "x", ""], ["", "x", ""]]

    assert sheet.to_python() == default
    assert list(sheet.iter_rows()) == default
    assert sheet.to_python(empty_as_none=True) == expected
    assert list(sheet.iter_rows(empty_as_none=True)) == expected


def test_leading_empty_rows_and_columns():
    sheet = CalamineWorkbook.from_path(DATA / "base.xlsx").get_sheet_by_index(0)
    assert sheet.to_python(skip_empty_area=False, nrows=1, empty_as_none=True) == [
        [None] * 10
    ]
    assert next(sheet.iter_rows(empty_as_none=True)) == [None] * 10
