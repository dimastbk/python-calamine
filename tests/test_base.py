from datetime import date
from pathlib import Path

from python_calamine import get_sheet_data, get_sheet_names

PATH = Path(__file__).parent / "data"


def test_ods_read():
    # calamine not supported ods date/datetime parse
    names = ["Sheet1", "Sheet2"]
    data = [
        ["", "", "", "", ""],
        ["String", 1, 1.1, True, False],
    ]

    assert names == get_sheet_names((PATH / "base.ods").as_posix())
    assert data == get_sheet_data(
        (PATH / "base.ods").as_posix(), 0, skip_empty_area=False
    )

    data_skipped = [
        ["String", 1, 1.1, True, False],
    ]
    assert data_skipped == get_sheet_data((PATH / "base.ods").as_posix(), 0)
    assert [] == get_sheet_data((PATH / "base.ods").as_posix(), 1)
    assert [] == get_sheet_data(
        (PATH / "base.ods").as_posix(), 1, skip_empty_area=False
    )


def test_xls_read():
    # calamine not supported xls date/datetime parse
    names = ["Sheet1", "Sheet2"]
    data = [
        ["", "", "", "", ""],
        ["String", 1, 1.1, True, False],
    ]

    assert names == get_sheet_names((PATH / "base.xls").as_posix())
    assert data == get_sheet_data(
        (PATH / "base.xls").as_posix(), 0, skip_empty_area=False
    )

    data_skipped = [
        ["String", 1, 1.1, True, False],
    ]
    assert data_skipped == get_sheet_data((PATH / "base.xls").as_posix(), 0)
    assert [] == get_sheet_data((PATH / "base.xls").as_posix(), 1)
    assert [] == get_sheet_data(
        (PATH / "base.xls").as_posix(), 1, skip_empty_area=False
    )


def test_xlsx_read():
    # calamine not supported xlsx datetime parse
    names = ["Sheet1", "Sheet2"]
    data = [
        ["", "", "", "", "", ""],
        ["String", 1, 1.1, True, False, date(2020, 1, 1)],
    ]

    assert names == get_sheet_names((PATH / "base.xlsx").as_posix())
    assert data == get_sheet_data(
        (PATH / "base.xlsx").as_posix(), 0, skip_empty_area=False
    )

    data_skipped = [
        ["String", 1, 1.1, True, False, date(2020, 1, 1)],
    ]
    assert data_skipped == get_sheet_data((PATH / "base.xlsx").as_posix(), 0)
    assert [] == get_sheet_data((PATH / "base.xlsx").as_posix(), 1)
    assert [] == get_sheet_data(
        (PATH / "base.xlsx").as_posix(), 1, skip_empty_area=False
    )
