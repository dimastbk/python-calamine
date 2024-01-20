from importlib.metadata import version
from pathlib import Path

import pytest
from packaging.version import Version, parse

try:
    import pandas as pd
    import pandas._testing as tm
except ImportError:
    pd = None
    tm = None

PATH = Path(__file__).parent / "data"


@pytest.fixture(scope="module", autouse=True)
def pandas_monkeypatch():
    if parse(version("pandas")) >= Version("2.2.0"):
        yield
    else:
        from python_calamine.pandas import pandas_monkeypatch

        pandas_monkeypatch()
        yield


@pytest.mark.skipif(not pd, reason="pandas is required")
def test_ods_pandas(expected_df_ods):
    result = pd.read_excel(PATH / "base.ods", sheet_name="Sheet1", engine="calamine")

    tm.assert_frame_equal(result, expected_df_ods)


@pytest.mark.skipif(not pd, reason="pandas is required")
@pytest.mark.xfail(reason="OdfReader can't parse timedelta")
def test_ods_odfpy_pandas():
    result_calamine = pd.read_excel(
        PATH / "base.ods", sheet_name="Sheet1", engine="calamine"
    )
    result_default = pd.read_excel(PATH / "base.ods", sheet_name="Sheet1")

    result_calamine = result_calamine.drop(
        result_calamine.columns[[8, 9]], axis=1, inplace=False
    )
    result_default = result_default.drop(
        result_default.columns[[8, 9]], axis=1, inplace=False
    )

    tm.assert_frame_equal(result_calamine, result_default)


@pytest.mark.skipif(not pd, reason="pandas is required")
def test_xls_pandas(expected_df_excel):
    result = pd.read_excel(PATH / "base.xls", sheet_name="Sheet1", engine="calamine")

    tm.assert_frame_equal(result, expected_df_excel)


@pytest.mark.skipif(not pd, reason="pandas is required")
def test_xls_xlrd_pandas():
    result_calamine = pd.read_excel(
        PATH / "base.xls", sheet_name="Sheet1", engine="calamine"
    )
    result_default = pd.read_excel(PATH / "base.xls", sheet_name="Sheet1")

    # pyxlsb doesn't support timdelta
    result_calamine = result_calamine.drop(
        result_calamine.columns[[8, 9]], axis=1, inplace=False
    )
    result_default = result_default.drop(
        result_default.columns[[8, 9]], axis=1, inplace=False
    )

    tm.assert_frame_equal(result_calamine, result_default)


@pytest.mark.skipif(not pd, reason="pandas is required")
def test_xlsb_pandas(expected_df_excel):
    result = pd.read_excel(PATH / "base.xlsb", sheet_name="Sheet1", engine="calamine")

    tm.assert_frame_equal(result, expected_df_excel)


@pytest.mark.skipif(not pd, reason="pandas is required")
def test_xlsb_pyxlsb_pandas():
    result_calamine = pd.read_excel(
        PATH / "base.xlsb", sheet_name="Sheet1", engine="calamine"
    )
    result_default = pd.read_excel(PATH / "base.xlsb", sheet_name="Sheet1")

    # pyxlsb doesn't support datetime
    result_calamine = result_calamine.drop(
        result_calamine.columns[[5, 6, 7, 8, 9]], axis=1, inplace=False
    )
    result_default = result_default.drop(
        result_default.columns[[5, 6, 7, 8, 9]], axis=1, inplace=False
    )

    tm.assert_frame_equal(result_calamine, result_default)


@pytest.mark.skipif(not pd, reason="pandas is required")
def test_xlsx_pandas(expected_df_excel):
    result = pd.read_excel(PATH / "base.xlsx", sheet_name="Sheet1", engine="calamine")

    tm.assert_frame_equal(result, expected_df_excel)


@pytest.mark.skipif(not pd, reason="pandas is required")
def test_xlsb_openpyxl_pandas():
    result_calamine = pd.read_excel(
        PATH / "base.xlsx", sheet_name="Sheet1", engine="calamine"
    )
    result_default = pd.read_excel(PATH / "base.xlsx", sheet_name="Sheet1")

    # openpyxl doesn't support timdelta
    result_calamine = result_calamine.drop(
        result_calamine.columns[[8, 9]], axis=1, inplace=False
    )
    result_default = result_default.drop(
        result_default.columns[[8, 9]], axis=1, inplace=False
    )

    tm.assert_frame_equal(result_calamine, result_default)
