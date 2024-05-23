from datetime import date, datetime, time, timedelta
from io import BytesIO
from pathlib import Path

import pytest
from python_calamine import CalamineWorkbook, PasswordError, WorksheetNotFound, ZipError

PATH = Path(__file__).parent / "data"

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
    assert data == list(reader.get_lazy_sheet_by_index(0))
