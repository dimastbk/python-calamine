from datetime import date, datetime, time
from io import BytesIO
from pathlib import Path

import pytest
from python_calamine import CalamineWorkbook

PATH = Path(__file__).parent / "data"


def test_ods_read():
    names = ["Sheet1", "Sheet2"]
    data = [
        ["", "", "", "", "", "", "", "", "", ""],
        [
            "String",
            1.0,
            1.1,
            True,
            False,
            date(2010, 10, 10),
            datetime(2010, 10, 10, 10, 10, 10),
            time(10, 10, 10),
            time(10, 10, 10, 100000),
            # duration (255:10:10) isn't supported
            # see https://github.com/tafia/calamine/pull/288 and https://github.com/chronotope/chrono/issues/579
            "PT255H10M10S",
        ],
    ]

    reader = CalamineWorkbook.from_object(PATH / "base.ods")

    assert names == reader.sheet_names
    assert data == reader.get_sheet_by_name("Sheet1").to_python(skip_empty_area=False)

    data_skipped = [
        [
            "String",
            1.0,
            1.1,
            True,
            False,
            date(2010, 10, 10),
            datetime(2010, 10, 10, 10, 10, 10),
            time(10, 10, 10),
            time(10, 10, 10, 100000),
            # duration (255:10:10) isn't supported
            # see https://github.com/tafia/calamine/pull/288 and https://github.com/chronotope/chrono/issues/579
            "PT255H10M10S",
        ],
    ]
    assert data_skipped == reader.get_sheet_by_index(0).to_python()
    assert [] == reader.get_sheet_by_index(1).to_python()
    assert [] == reader.get_sheet_by_index(1).to_python(skip_empty_area=False)


def test_xls_read():
    # calamine not supported xls date/datetime parse
    names = ["Sheet1", "Sheet2"]
    data = [
        ["", "", "", "", ""],
        ["String", 1, 1.1, True, False],
    ]

    reader = CalamineWorkbook.from_object(PATH / "base.xls")

    assert names == reader.sheet_names
    assert data == reader.get_sheet_by_index(0).to_python(skip_empty_area=False)

    data_skipped = [
        ["String", 1, 1.1, True, False],
    ]
    assert data_skipped == reader.get_sheet_by_index(0).to_python()
    assert [] == reader.get_sheet_by_index(1).to_python()
    assert [] == reader.get_sheet_by_index(1).to_python(skip_empty_area=False)


def test_xlsx_read():
    # calamine not supported xlsx date parse
    names = ["Sheet1", "Sheet2"]
    data = [
        ["", "", "", "", "", ""],
        ["String", 1, 1.1, True, False, date(2020, 1, 1)],
    ]

    reader = CalamineWorkbook.from_object(PATH / "base.xlsx")

    assert names == reader.sheet_names
    assert data == reader.get_sheet_by_index(0).to_python(skip_empty_area=False)

    data_skipped = [
        ["String", 1, 1.1, True, False, date(2020, 1, 1)],
    ]
    assert data_skipped == reader.get_sheet_by_index(0).to_python()
    assert [] == reader.get_sheet_by_index(1).to_python()
    assert [] == reader.get_sheet_by_index(1).to_python(skip_empty_area=False)


@pytest.mark.parametrize(
    "obj",
    [
        (PATH / "base.xlsx").as_posix(),
    ],
)
def test_path(obj):
    CalamineWorkbook.from_path(obj)


@pytest.mark.parametrize(
    "obj",
    [
        open(PATH / "base.xlsx", "rb"),
        BytesIO(open(PATH / "base.xlsx", "rb").read()),
    ],
)
def test_filelike(obj):
    CalamineWorkbook.from_filelike(obj)


@pytest.mark.parametrize(
    "obj",
    [
        PATH / "base.xlsx",
        (PATH / "base.xlsx").as_posix(),
        open(PATH / "base.xlsx", "rb"),
        BytesIO(open(PATH / "base.xlsx", "rb").read()),
    ],
)
def test_path_or_filelike(obj):
    CalamineWorkbook.from_object(obj)


def test_path_or_filelike_error():
    with pytest.raises(TypeError):
        CalamineWorkbook.from_object(object())
