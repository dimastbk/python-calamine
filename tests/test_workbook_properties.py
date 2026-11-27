from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest
from python_calamine import CalamineWorkbook, WorkbookClosed, WorkbookProperties

PATH = Path(__file__).parent / "data"


def test_xlsx_workbook_properties():
    props = CalamineWorkbook.from_path(
        PATH / "workbook_properties.xlsx"
    ).workbook_properties

    assert isinstance(props, WorkbookProperties)
    assert props.creator == "Test Creator"
    assert props.last_modified_by == "Last Modifier"
    assert props.created == datetime(2024, 1, 15, 8, 30, tzinfo=timezone.utc)
    assert props.modified == datetime(2024, 6, 20, 14, 22, tzinfo=timezone.utc)
    assert props.title == "Workbook Title"
    assert props.subject == "Test Subject"
    assert props.description == "A test workbook"
    assert props.keywords == "test, calamine"
    assert props.category == "Testing"
    assert props.content_status == "Draft"
    assert props.revision == "3"
    assert props.version == "1.2"
    assert props.application == "Microsoft Excel"
    assert props.app_version == "16.0300"
    assert props.company == "Contoso & Co"
    assert props.template == "Book.xltx"
    assert props.manager == "Jane Doe"
    assert props.custom_properties == {"Department": "Finance", "Reviewed": "true"}


def test_ods_workbook_properties():
    props = CalamineWorkbook.from_path(
        PATH / "workbook_properties.ods"
    ).workbook_properties

    assert props.creator == "Test Creator"
    assert props.last_modified_by == "Last Modifier"
    assert props.created == datetime(2024, 1, 15, 8, 30)
    assert props.modified == datetime(2024, 6, 20, 14, 22, 0, 123456)
    assert props.title == "Workbook Title"
    assert props.subject == "Test Subject"
    assert props.description == "A test workbook"
    assert props.keywords == "test, calamine"
    assert props.revision == "3"
    assert props.application is not None
    assert props.application.startswith("LibreOffice/24.2")
    assert props.company is None
    assert props.manager is None
    assert props.custom_properties == {
        "Department": "Finance & Ops",
        "Reviewed": "true",
    }


@pytest.mark.parametrize(
    ("file", "modified", "application"),
    [
        (
            "any_sheets.xlsx",
            datetime(2023, 7, 28, 16, 31, 17, tzinfo=timezone.utc),
            "Microsoft Excel",
        ),
        (
            "any_sheets.xlsb",
            datetime(2023, 7, 28, 16, 31, 46, tzinfo=timezone.utc),
            "Microsoft Excel",
        ),
        (
            "any_sheets.xls",
            datetime(2023, 7, 28, 16, 31, 33, tzinfo=timezone.utc),
            "Microsoft Excel",
        ),
        (
            "any_sheets.ods",
            datetime(2023, 7, 28, 16, 32, 8, tzinfo=timezone.utc),
            "MicrosoftOffice/16.0 MicrosoftExcel/CalculationVersion-14420",
        ),
    ],
)
def test_workbook_properties_all_formats(file, modified, application):
    props = CalamineWorkbook.from_path(PATH / file).workbook_properties

    assert props.creator == "Алена"
    assert props.last_modified_by == "Алена"
    assert props.created == datetime(2023, 7, 28, 16, 28, 6, tzinfo=timezone.utc)
    assert props.modified == modified
    assert props.application == application
    assert props.title is None
    assert props.custom_properties == {}


@pytest.mark.parametrize(
    "file",
    ["any_sheets.xlsx", "any_sheets.xlsb", "any_sheets.xls", "any_sheets.ods"],
)
def test_workbook_properties_from_filelike(file):
    with open(PATH / file, "rb") as f:
        props = CalamineWorkbook.from_filelike(BytesIO(f.read())).workbook_properties

    assert props == CalamineWorkbook.from_path(PATH / file).workbook_properties
    assert props.creator == "Алена"


def test_workbook_properties_missing_fields():
    props = CalamineWorkbook.from_path(PATH / "base.xlsx").workbook_properties

    assert props.creator is None
    assert props.last_modified_by is None
    assert props.created == datetime(2021, 9, 17, 15, 27, 18, tzinfo=timezone.utc)
    assert props.app_version == "16.0300"


def test_workbook_properties_unrepresentable_dates():
    # Python's datetime has no year 0, so these dates are None.
    buffer = BytesIO()
    with ZipFile(PATH / "base.xlsx") as src, ZipFile(buffer, "w") as dst:
        for item in src.infolist():
            data = src.read(item)
            if item.filename == "docProps/core.xml":
                data = data.replace(b"2021-09-17T15:27:18Z", b"0000-01-01T00:00:00Z")
                data = data.replace(b"2025-04-13T17:29:27Z", b"0000-01-01T00:00:00")
            dst.writestr(item, data)
    buffer.seek(0)

    props = CalamineWorkbook.from_filelike(buffer).workbook_properties

    assert props.created is None
    assert props.modified is None
    assert repr(props) == "WorkbookProperties(app_version='16.0300')"


def test_workbook_properties_repr():
    props = CalamineWorkbook.from_path(PATH / "base.xls").workbook_properties

    assert repr(props) == (
        "WorkbookProperties("
        "created=datetime.datetime(2021, 9, 17, 5, 27, 18, tzinfo=datetime.timezone.utc), "
        "modified=datetime.datetime(2025, 4, 13, 18, 43, 16, tzinfo=datetime.timezone.utc), "
        "revision='18', app_version='16.0000', "
        "custom_properties={'AppVersion': '15.0000'})"
    )


def test_workbook_properties_closed():
    reader = CalamineWorkbook.from_path(PATH / "base.xlsx")
    reader.close()

    with pytest.raises(WorkbookClosed):
        reader.workbook_properties
