from datetime import date, datetime, time, timedelta
from pathlib import Path

from python_calamine import CalamineWorkbook

PATH = Path(__file__).parent / "data"


def __old_convert_cell(value):
    if isinstance(value, float):
        # GH#54564 - is_integer() returns False for NaN/Inf,
        # so this safely avoids int() on non-finite values
        if value.is_integer():
            return int(value)
        return value
    elif isinstance(value, (datetime, timedelta)):
        # Return as-is to match openpyxl behavior (GH#59186)
        return value
    elif isinstance(value, date):
        # Convert date to datetime to match openpyxl behavior (GH#59186)
        return datetime(value.year, value.month, value.day)
    elif isinstance(value, time):
        return value
    return value


def test_old_pandas_ods_large_integer():
    sheet = CalamineWorkbook.from_object(
        PATH / "large_integer_pandas.ods"
    ).get_sheet_by_index(0)

    old_data = [[__old_convert_cell(y) for y in x] for x in sheet.to_python()]
    new_data = sheet.to_python_pandas()

    assert old_data == new_data


def test_old_pandas_xlsx_large_integer():
    sheet = CalamineWorkbook.from_object(
        PATH / "large_integer_pandas.xlsx"
    ).get_sheet_by_index(0)

    old_data = [[__old_convert_cell(y) for y in x] for x in sheet.to_python()]
    new_data = sheet.to_python_pandas()

    assert old_data == new_data


def test_old_pandas_xlsx():
    sheet_names = ["Sheet1", "Sheet2", "Merged Cells"]
    wb = CalamineWorkbook.from_object(PATH / "base.xlsx")

    for sheet_name in sheet_names:
        sheet = wb.get_sheet_by_name(sheet_name)
        old_data = [[__old_convert_cell(y) for y in x] for x in sheet.to_python()]
        new_data = sheet.to_python_pandas()

        assert old_data == new_data


def test_old_pandas_xls():
    sheet_names = ["Sheet1", "Sheet2", "Merged Cells"]
    wb = CalamineWorkbook.from_object(PATH / "base.xls")

    for sheet_name in sheet_names:
        sheet = wb.get_sheet_by_name(sheet_name)
        old_data = [[__old_convert_cell(y) for y in x] for x in sheet.to_python()]
        new_data = sheet.to_python_pandas()

        assert old_data == new_data


def test_old_pandas_xlsb():
    sheet_names = ["Sheet1", "Sheet2", "Merged Cells"]
    wb = CalamineWorkbook.from_object(PATH / "base.xlsb")

    for sheet_name in sheet_names:
        sheet = wb.get_sheet_by_name(sheet_name)
        old_data = [[__old_convert_cell(y) for y in x] for x in sheet.to_python()]
        new_data = sheet.to_python_pandas()

        assert old_data == new_data


def test_old_pandas_ods():
    sheet_names = ["Sheet1", "Sheet2", "Merged Cells"]
    wb = CalamineWorkbook.from_object(PATH / "base.ods")

    for sheet_name in sheet_names:
        sheet = wb.get_sheet_by_name(sheet_name)
        old_data = [[__old_convert_cell(y) for y in x] for x in sheet.to_python()]
        new_data = sheet.to_python_pandas()

        assert old_data == new_data
