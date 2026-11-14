from pathlib import Path

import pytest
from python_calamine import CalamineWorkbook

DATA = Path(__file__).parent / "data"
COMMON_ERRORS = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"]


@pytest.mark.parametrize(
    ("filename", "errors"),
    [("error-cells.xlsx", COMMON_ERRORS), ("error-cells.xls", [*COMMON_ERRORS, "#DATA!"])],
)
def test_error_cells_keep_codes_in_both_sheet_apis(filename, errors):
    with CalamineWorkbook.from_path(DATA / filename) as workbook:
        sheet = workbook.get_sheet_by_name("Errors")
        expected = [errors + ["", "#N/A"]]
        assert list(sheet.iter_rows()) == expected
        assert sheet.to_python() == expected
