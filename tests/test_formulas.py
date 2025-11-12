from pathlib import Path

import pytest
from python_calamine import CalamineWorkbook

PATH = Path(__file__).parent / "data"


def test_formulas_read_enabled_xlsx():
    """Test reading formulas when read_formulas=True for xlsx files"""
    reader = CalamineWorkbook.from_object(PATH / "formula.xlsx", read_formulas=True)
    sheet = reader.get_sheet_by_name("Sheet1")

    # Test that iter_formulas works when enabled
    formulas = list(sheet.iter_formulas())

    # Should have same dimensions as data
    data = list(sheet.iter_rows())
    assert len(formulas) == len(data)
    assert len(formulas[0]) == len(data[0])

    # The formula should be "SUM(A1:B1)" at position C1 (column index 2)
    assert formulas[0][2] == "SUM(A1:B1)"
    # A1 and B1 should be empty (no formulas)
    assert formulas[0][0] == ""
    assert formulas[0][1] == ""


def test_formulas_read_enabled_xls():
    """Test reading formulas when read_formulas=True for xls files"""
    reader = CalamineWorkbook.from_object(PATH / "formula.xls", read_formulas=True)
    sheet = reader.get_sheet_by_name("Sheet1")

    # Test that iter_formulas works when enabled
    formulas = list(sheet.iter_formulas())

    # Should have same dimensions as data
    data = list(sheet.iter_rows())
    assert len(formulas) == len(data)
    assert len(formulas[0]) == len(data[0])

    # The formula should be at position C1 (column index 2) and contain SUM
    assert "SUM" in formulas[0][2]
    # A1 and B1 should be empty (no formulas)
    assert formulas[0][0] == ""
    assert formulas[0][1] == ""


def test_formulas_read_enabled_ods():
    """Test reading formulas when read_formulas=True for ods files"""
    reader = CalamineWorkbook.from_object(PATH / "formula.ods", read_formulas=True)
    sheet = reader.get_sheet_by_name("Sheet1")

    # Test that iter_formulas works when enabled
    formulas = list(sheet.iter_formulas())

    # Should have same dimensions as data
    data = list(sheet.iter_rows())
    assert len(formulas) == len(data)
    assert len(formulas[0]) == len(data[0])

    # ODS uses different formula syntax at position C1: "of:=SUM([.A1:.B1])"
    assert "SUM" in formulas[0][2]
    assert "of:=" in formulas[0][2]
    # A1 and B1 should be empty (no formulas)
    assert formulas[0][0] == ""
    assert formulas[0][1] == ""


@pytest.mark.parametrize(
    "file_ext",
    ["xlsx", "xls", "ods"],
)
def test_formulas_read_disabled_default(file_ext):
    """Test that formulas are disabled by default (read_formulas=False)"""
    reader = CalamineWorkbook.from_object(PATH / f"formula.{file_ext}")
    sheet = reader.get_sheet_by_name("Sheet1")

    # Should raise ValueError when trying to access formulas without enabling them
    with pytest.raises(ValueError, match="Formula iteration is disabled"):
        list(sheet.iter_formulas())


@pytest.mark.parametrize(
    "file_ext",
    ["xlsx", "xls", "ods"],
)
def test_formulas_read_disabled_explicit(file_ext):
    """Test that formulas are disabled when explicitly set to False"""
    reader = CalamineWorkbook.from_object(
        PATH / f"formula.{file_ext}", read_formulas=False
    )
    sheet = reader.get_sheet_by_name("Sheet1")

    # Should raise ValueError when trying to access formulas
    with pytest.raises(ValueError, match="Formula iteration is disabled"):
        list(sheet.iter_formulas())


def test_formulas_from_path_enabled():
    """Test reading formulas using from_path with read_formulas=True"""
    reader = CalamineWorkbook.from_path(PATH / "formula.xlsx", read_formulas=True)
    sheet = reader.get_sheet_by_name("Sheet1")

    formulas = list(sheet.iter_formulas())
    data = list(sheet.iter_rows())
    assert len(formulas) == len(data)
    assert formulas[0][2] == "SUM(A1:B1)"


def test_formulas_from_path_disabled():
    """Test that formulas are disabled by default with from_path"""
    reader = CalamineWorkbook.from_path(PATH / "formula.xlsx")
    sheet = reader.get_sheet_by_name("Sheet1")

    with pytest.raises(ValueError, match="Formula iteration is disabled"):
        list(sheet.iter_formulas())


def test_formulas_from_filelike_enabled():
    """Test reading formulas using from_filelike with read_formulas=True"""
    with open(PATH / "formula.xlsx", "rb") as f:
        reader = CalamineWorkbook.from_filelike(f, read_formulas=True)
        sheet = reader.get_sheet_by_name("Sheet1")

        formulas = list(sheet.iter_formulas())
        data = list(sheet.iter_rows())
        assert len(formulas) == len(data)
        assert formulas[0][2] == "SUM(A1:B1)"


def test_formulas_from_filelike_disabled():
    """Test that formulas are disabled by default with from_filelike"""
    with open(PATH / "formula.xlsx", "rb") as f:
        reader = CalamineWorkbook.from_filelike(f)
        sheet = reader.get_sheet_by_name("Sheet1")

        with pytest.raises(ValueError, match="Formula iteration is disabled"):
            list(sheet.iter_formulas())


def test_formulas_with_load_workbook_enabled():
    """Test reading formulas using load_workbook function with read_formulas=True"""
    from python_calamine import load_workbook

    reader = load_workbook(PATH / "formula.xlsx", read_formulas=True)
    sheet = reader.get_sheet_by_name("Sheet1")

    formulas = list(sheet.iter_formulas())
    data = list(sheet.iter_rows())
    assert len(formulas) == len(data)
    assert formulas[0][2] == "SUM(A1:B1)"


def test_formulas_with_load_workbook_disabled():
    """Test that formulas are disabled by default with load_workbook function"""
    from python_calamine import load_workbook

    reader = load_workbook(PATH / "formula.xlsx")
    sheet = reader.get_sheet_by_name("Sheet1")

    with pytest.raises(ValueError, match="Formula iteration is disabled"):
        list(sheet.iter_formulas())


def test_formulas_read_flag_exposed():
    """Test that read_formulas flag is accessible on the workbook"""
    reader_enabled = CalamineWorkbook.from_object(
        PATH / "formula.xlsx", read_formulas=True
    )
    reader_disabled = CalamineWorkbook.from_object(
        PATH / "formula.xlsx", read_formulas=False
    )

    assert reader_enabled.read_formulas is True
    assert reader_disabled.read_formulas is False


def test_formulas_iterator_properties():
    """Test that formula iterator exposes position, start, width, and height properties"""
    reader = CalamineWorkbook.from_object(PATH / "formula.xlsx", read_formulas=True)
    sheet = reader.get_sheet_by_name("Sheet1")

    formula_iter = sheet.iter_formulas()

    # Test that all properties are accessible
    assert hasattr(formula_iter, "position")
    assert hasattr(formula_iter, "start")
    assert hasattr(formula_iter, "width")
    assert hasattr(formula_iter, "height")

    # Initial position should be 0
    assert formula_iter.position == 0

    # Start should be a tuple of coordinates
    assert isinstance(formula_iter.start, tuple)
    assert len(formula_iter.start) == 2

    # Width and height should match data dimensions
    data = list(sheet.iter_rows())
    assert formula_iter.width == len(data[0])
    assert formula_iter.height == len(data)


def test_data_vs_formulas():
    """Test that data and formulas are different"""
    reader_data = CalamineWorkbook.from_object(
        PATH / "formula.xlsx", read_formulas=False
    )
    reader_formulas = CalamineWorkbook.from_object(
        PATH / "formula.xlsx", read_formulas=True
    )

    sheet_data = reader_data.get_sheet_by_name("Sheet1")
    sheet_formulas = reader_formulas.get_sheet_by_name("Sheet1")

    # Get the data values
    data = list(sheet_data.iter_rows())

    # Get the formulas
    formulas = list(sheet_formulas.iter_formulas())

    # Should have same dimensions
    assert len(data) == len(formulas)
    assert len(data[0]) == len(formulas[0])

    # The data should show the calculated result: [10.0, 15.0, 25.0]
    assert data[0] == [10.0, 15.0, 25.0]

    # The formula should show: ["", "", "SUM(A1:B1)"]
    assert formulas[0] == ["", "", "SUM(A1:B1)"]

    # C1 values should be different (result vs formula)
    assert data[0][2] != formulas[0][2]
    assert data[0][2] == 25.0
    assert formulas[0][2] == "SUM(A1:B1)"
