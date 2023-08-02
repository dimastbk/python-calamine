from pathlib import Path

from python_calamine import (
    CalamineWorkbook,
    SheetMetadata,
    SheetTypeEnum,
    SheetVisibleEnum,
)

PATH = Path(__file__).parent / "data"


def test_sheet_type_ods():
    reader = CalamineWorkbook.from_object(PATH / "any_sheets.ods")

    assert reader.sheets_metadata == [
        SheetMetadata(
            name="Visible",
            typ=SheetTypeEnum.WorkSheet,
            visible=SheetVisibleEnum.Visible,
        ),
        SheetMetadata(
            name="Hidden",
            typ=SheetTypeEnum.WorkSheet,
            visible=SheetVisibleEnum.Hidden,
        ),
        # ODS doesn't support Very Hidden
        SheetMetadata(
            name="VeryHidden",
            typ=SheetTypeEnum.WorkSheet,
            visible=SheetVisibleEnum.Hidden,
        ),
        # ODS doesn't support chartsheet
        SheetMetadata(
            name="Chart",
            typ=SheetTypeEnum.WorkSheet,
            visible=SheetVisibleEnum.Visible,
        ),
    ]


def test_sheet_type_xls():
    reader = CalamineWorkbook.from_object(PATH / "any_sheets.xls")

    assert reader.sheets_metadata == [
        SheetMetadata(
            name="Visible",
            typ=SheetTypeEnum.WorkSheet,
            visible=SheetVisibleEnum.Visible,
        ),
        SheetMetadata(
            name="Hidden",
            typ=SheetTypeEnum.WorkSheet,
            visible=SheetVisibleEnum.Hidden,
        ),
        SheetMetadata(
            name="VeryHidden",
            typ=SheetTypeEnum.WorkSheet,
            visible=SheetVisibleEnum.VeryHidden,
        ),
        SheetMetadata(
            name="Chart",
            typ=SheetTypeEnum.ChartSheet,
            visible=SheetVisibleEnum.Visible,
        ),
    ]


def test_sheet_type_xlsx():
    reader = CalamineWorkbook.from_object(PATH / "any_sheets.xlsx")

    assert reader.sheets_metadata == [
        SheetMetadata(
            name="Visible",
            typ=SheetTypeEnum.WorkSheet,
            visible=SheetVisibleEnum.Visible,
        ),
        SheetMetadata(
            name="Hidden",
            typ=SheetTypeEnum.WorkSheet,
            visible=SheetVisibleEnum.Hidden,
        ),
        SheetMetadata(
            name="VeryHidden",
            typ=SheetTypeEnum.WorkSheet,
            visible=SheetVisibleEnum.VeryHidden,
        ),
        SheetMetadata(
            name="Chart",
            typ=SheetTypeEnum.ChartSheet,
            visible=SheetVisibleEnum.Visible,
        ),
    ]


def test_sheet_type_xlsb():
    reader = CalamineWorkbook.from_object(PATH / "any_sheets.xlsb")

    assert reader.sheets_metadata == [
        SheetMetadata(
            name="Visible",
            typ=SheetTypeEnum.WorkSheet,
            visible=SheetVisibleEnum.Visible,
        ),
        SheetMetadata(
            name="Hidden",
            typ=SheetTypeEnum.WorkSheet,
            visible=SheetVisibleEnum.Hidden,
        ),
        SheetMetadata(
            name="VeryHidden",
            typ=SheetTypeEnum.WorkSheet,
            visible=SheetVisibleEnum.VeryHidden,
        ),
        SheetMetadata(
            name="Chart",
            typ=SheetTypeEnum.ChartSheet,
            visible=SheetVisibleEnum.Visible,
        ),
    ]
