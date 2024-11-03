from datetime import date, datetime, time, timedelta
from io import BytesIO
from pathlib import Path

import pytest
from python_calamine import (
    CalamineWorkbook,
    PasswordError,
    WorkbookClosed,
    WorksheetNotFound,
    ZipError,
)

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


def test_xlsx_iter_rows():
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
    assert data == list(reader.get_sheet_by_index(0).iter_rows())


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
    "path",
    [
        PATH / "base.xlsx",
        PATH / "base.xls",
        PATH / "base.xlsb",
        PATH / "base.ods",
    ],
)
def test_worksheet_errors(path):
    reader = CalamineWorkbook.from_object(path)
    with pytest.raises(WorksheetNotFound):
        reader.get_sheet_by_name("Sheet4")


@pytest.mark.parametrize(
    "path",
    [
        PATH / "password.xlsx",
        PATH / "password.xls",
        PATH / "password.xlsb",
        PATH / "password.ods",
    ],
)
def test_password_errors(path):
    with pytest.raises(PasswordError):
        CalamineWorkbook.from_object(path)


@pytest.mark.parametrize(
    "path",
    [
        PATH / "empty_file.xlsx",
        PATH / "empty_file.xlsb",
        PATH / "empty_file.ods",
    ],
)
def test_zip_errors(path):
    with pytest.raises(ZipError):
        CalamineWorkbook.from_path(path)


@pytest.mark.parametrize(
    "path",
    [
        PATH / "non_existent_file.xlsx",
        PATH / "non_existent_file.xls",
        PATH / "non_existent_file.xlsb",
        PATH / "non_existent_file.ods",
    ],
)
def test_io_errors(path):
    with pytest.raises(IOError):
        CalamineWorkbook.from_path(path)


@pytest.mark.parametrize(
    "path",
    [
        PATH / "base.xlsx",
        (PATH / "base.xlsx").as_posix(),
    ],
)
def test_path(path):
    CalamineWorkbook.from_path(path)


def test_path_error():
    with pytest.raises(TypeError):
        CalamineWorkbook.from_path(object())


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


def test_close_workbook():
    reader = CalamineWorkbook.from_path(PATH / "base.xlsx")
    reader.close()

    with pytest.raises(WorkbookClosed):
        reader.get_sheet_by_index(1)


def test_close_workbook_double():
    reader = CalamineWorkbook.from_path(PATH / "base.xlsx")
    reader.close()

    with pytest.raises(WorkbookClosed):
        reader.close()
