from io import BytesIO
from pathlib import Path

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
    assert props.created == "2024-01-15T08:30:00Z"
    assert props.modified == "2024-06-20T14:22:00Z"
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
    assert props.created == "2024-01-15T08:30:00"
    assert props.modified == "2024-06-20T14:22:00.123456789"
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
            "2023-07-28T16:31:17Z",
            "Microsoft Excel",
        ),
        (
            "any_sheets.xlsb",
            "2023-07-28T16:31:46Z",
            "Microsoft Excel",
        ),
        (
            "any_sheets.xls",
            "2023-07-28T16:31:33Z",
            "Microsoft Excel",
        ),
        (
            "any_sheets.ods",
            "2023-07-28T16:32:08Z",
            "MicrosoftOffice/16.0 MicrosoftExcel/CalculationVersion-14420",
        ),
    ],
)
def test_workbook_properties_all_formats(file, modified, application):
    props = CalamineWorkbook.from_path(PATH / file).workbook_properties

    assert props.creator == "Алена"
    assert props.last_modified_by == "Алена"
    assert props.created == "2023-07-28T16:28:06Z"
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
    assert props.created == "2021-09-17T15:27:18Z"
    assert props.app_version == "16.0300"


def test_workbook_properties_repr():
    props = CalamineWorkbook.from_path(PATH / "base.xls").workbook_properties

    assert repr(props) == (
        "WorkbookProperties(created='2021-09-17T05:27:18Z', "
        "modified='2025-04-13T18:43:16Z', revision='18', app_version='16.0000', "
        "custom_properties={'AppVersion': '15.0000'})"
    )


def test_workbook_properties_closed():
    reader = CalamineWorkbook.from_path(PATH / "base.xlsx")
    reader.close()

    with pytest.raises(WorkbookClosed):
        reader.workbook_properties
