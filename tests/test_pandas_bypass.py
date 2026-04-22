from datetime import date, datetime, time, timedelta
from pathlib import Path

from python_calamine import CalamineWorkbook

PATH = Path(__file__).parent / "data"


def __old_convert_cell(value):
    if isinstance(value, float):
        val = int(value)
        if val == value:
            return val
        else:
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


from pprint import pprint


def test_old_pandas_ods():
    sheet_names = ["Sheet1", "Sheet2", "Merged Cells"]
    wb = CalamineWorkbook.from_object(PATH / "base.ods")

    for sheet_name in sheet_names:
        sheet = wb.get_sheet_by_name(sheet_name)
        pprint(sheet.to_python())
        old_data = [[__old_convert_cell(y) for y in x] for x in sheet.to_python()]
        new_data = sheet.to_python_pandas()

        pprint(old_data)
        pprint(new_data)

        assert old_data == new_data
