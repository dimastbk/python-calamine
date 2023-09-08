from datetime import date, datetime, time, timedelta
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
    names = ["Sheet1", "Sheet2"]
    data = [
        ["", "", "", "", "", "", "", "", "", ""],
        [
            "String",
            1,
            1.1,
            True,
            False,
            date(2010, 10, 10),
            datetime(2010, 10, 10, 10, 10, 10),
            time(10, 10, 10),
            timedelta(hours=10, minutes=10, seconds=10, microseconds=100000),
            timedelta(hours=255, minutes=10, seconds=10),
        ],
    ]

    reader = CalamineWorkbook.from_object(PATH / "base.xls")

    assert names == reader.sheet_names
    assert data == reader.get_sheet_by_index(0).to_python(skip_empty_area=False)

    data_skipped = [
        [
            "String",
            1,
            1.1,
            True,
            False,
            date(2010, 10, 10),
            datetime(2010, 10, 10, 10, 10, 10),
            time(10, 10, 10),
            timedelta(hours=10, minutes=10, seconds=10, microseconds=100000),
            timedelta(hours=255, minutes=10, seconds=10),
        ],
    ]
    assert data_skipped == reader.get_sheet_by_index(0).to_python()
    assert [] == reader.get_sheet_by_index(1).to_python()
    assert [] == reader.get_sheet_by_index(1).to_python(skip_empty_area=False)


def test_xlsb_read():
    names = ["Sheet1", "Sheet2", "Sheet3"]
    data = [
        ["", "", "", "", "", "", "", "", "", ""],
        [
            "String",
            1,
            1.1,
            True,
            False,
            date(2010, 10, 10),
            datetime(2010, 10, 10, 10, 10, 10),
            time(10, 10, 10),
            timedelta(hours=10, minutes=10, seconds=10, microseconds=100000),
            timedelta(hours=255, minutes=10, seconds=10),
        ],
    ]

    reader = CalamineWorkbook.from_object(PATH / "base.xlsb")

    assert names == reader.sheet_names
    assert data == reader.get_sheet_by_index(0).to_python(skip_empty_area=False)

    data_skipped = [
        [
            "String",
            1,
            1.1,
            True,
            False,
            date(2010, 10, 10),
            datetime(2010, 10, 10, 10, 10, 10),
            time(10, 10, 10),
            timedelta(hours=10, minutes=10, seconds=10, microseconds=100000),
            timedelta(hours=255, minutes=10, seconds=10),
        ],
    ]
    assert data_skipped == reader.get_sheet_by_index(0).to_python()
    assert [] == reader.get_sheet_by_index(1).to_python()
    assert [] == reader.get_sheet_by_index(1).to_python(skip_empty_area=False)


def test_xlsx_read():
    names = ["Sheet1", "Sheet2", "Sheet3"]
    data = [
        ["", "", "", "", "", "", "", "", "", ""],
        [
            "String",
            1,
            1.1,
            True,
            False,
            date(2010, 10, 10),
            datetime(2010, 10, 10, 10, 10, 10),
            time(10, 10, 10),
            timedelta(hours=10, minutes=10, seconds=10, microseconds=100000),
            timedelta(hours=255, minutes=10, seconds=10),
        ],
    ]

    reader = CalamineWorkbook.from_object(PATH / "base.xlsx")

    assert names == reader.sheet_names
    assert data == reader.get_sheet_by_index(0).to_python(skip_empty_area=False)

    data_skipped = [
        [
            "String",
            1,
            1.1,
            True,
            False,
            date(2010, 10, 10),
            datetime(2010, 10, 10, 10, 10, 10),
            time(10, 10, 10),
            timedelta(hours=10, minutes=10, seconds=10, microseconds=100000),
            timedelta(hours=255, minutes=10, seconds=10),
        ],
    ]
    assert data_skipped == reader.get_sheet_by_index(0).to_python()
    assert [] == reader.get_sheet_by_index(1).to_python()
    assert [] == reader.get_sheet_by_index(1).to_python(skip_empty_area=False)


def test_nrows():
    reader = CalamineWorkbook.from_object(PATH / "base.xlsx")
    sheet = reader.get_sheet_by_name("Sheet3")

    assert sheet.to_python(nrows=1) == [
        ["line1", "line1", "line1"],
    ]

    assert sheet.to_python(nrows=2) == [
        ["line1", "line1", "line1"],
        ["line2", "line2", "line2"],
    ]

    assert sheet.to_python(nrows=4) == [
        ["line1", "line1", "line1"],
        ["line2", "line2", "line2"],
        ["line3", "line3", "line3"],
    ]

    assert sheet.to_python(skip_empty_area=False, nrows=2) == [
        ["", "", "", ""],
        ["", "line1", "line1", "line1"],
    ]

    assert sheet.to_python(skip_empty_area=False, nrows=5) == [
        ["", "", "", ""],
        ["", "line1", "line1", "line1"],
        ["", "line2", "line2", "line2"],
        ["", "line3", "line3", "line3"],
    ]

    assert sheet.to_python() == [
        ["line1", "line1", "line1"],
        ["line2", "line2", "line2"],
        ["line3", "line3", "line3"],
    ]


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
