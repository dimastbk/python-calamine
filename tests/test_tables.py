from pathlib import Path

import pytest
from python_calamine import (
    CalamineWorkbook,
    TableNotFound,
    TablesNotLoaded,
    TablesNotSupported,
)

PATH = Path(__file__).parent / "data"


# NOTE: "table-multiple.xlsx" from tafia/calamine
def test_table_names_xlsx():
    reader = CalamineWorkbook.from_object(
        PATH / "table-multiple.xlsx", load_tables=True
    )

    assert reader.table_names == ["Inventory", "Pricing", "Sales_Bob", "Sales_Alice"]


def test_table_names_not_loaded():
    reader = CalamineWorkbook.from_object(PATH / "table-multiple.xlsx")

    with pytest.raises(TablesNotLoaded):
        reader.table_names


def test_table_names_not_supported():
    with pytest.raises(TablesNotSupported):
        CalamineWorkbook.from_object(PATH / "base.xlsb", load_tables=True)


def test_table_get_by_name():
    reader = CalamineWorkbook.from_object(
        PATH / "table-multiple.xlsx", load_tables=True
    )

    table = reader.get_table_by_name("Inventory")

    assert table.sheet == "Sheet1"
    assert table.name == "Inventory"
    assert table.columns == ["Item", "Type", "Quantity"]

    assert table.height == 4
    assert table.width == 3

    assert table.start == (1, 0)
    assert table.end == (4, 2)

    assert table.to_python() == [
        [1.0, "Apple", 50.0],
        [2.0, "Banana", 200.0],
        [3.0, "Orange", 60.0],
        [4.0, "Pear", 100.0],
    ]


def test_table_get_by_name_not_found():
    reader = CalamineWorkbook.from_object(
        PATH / "table-multiple.xlsx", load_tables=True
    )

    with pytest.raises(TableNotFound):
        reader.get_table_by_name("not found table")


def test_table_get_by_name_not_supported():
    reader = CalamineWorkbook.from_object(PATH / "base.xlsb")

    with pytest.raises(TablesNotSupported):
        reader.get_table_by_name("not found table")
