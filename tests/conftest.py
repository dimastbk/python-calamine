from datetime import datetime, time

import pytest


@pytest.fixture
def expected_df_ods():
    import pandas as pd

    return pd.DataFrame(
        [
            [
                "String",
                1,
                1.1,
                True,
                False,
                pd.Timestamp("2010-10-10"),
                datetime(2010, 10, 10, 10, 10, 10),
                time(10, 10, 10),
                time(10, 10, 10, 100000),
                # duration (255:10:10) isn't supported
                # see https://github.com/tafia/calamine/pull/288 and https://github.com/chronotope/chrono/issues/579
                "PT255H10M10S",
            ],
        ],
        columns=[
            "Unnamed: 0",
            "Unnamed: 1",
            "Unnamed: 2",
            "Unnamed: 3",
            "Unnamed: 4",
            "Unnamed: 5",
            "Unnamed: 6",
            "Unnamed: 7",
            "Unnamed: 8",
            "Unnamed: 9",
        ],
    )


@pytest.fixture
def expected_df_excel():
    import pandas as pd

    return pd.DataFrame(
        [
            [
                "String",
                1,
                1.1,
                True,
                False,
                pd.Timestamp("2010-10-10"),
                datetime(2010, 10, 10, 10, 10, 10),
                time(10, 10, 10),
                pd.Timedelta(hours=10, minutes=10, seconds=10, microseconds=100000),
                pd.Timedelta(hours=255, minutes=10, seconds=10),
            ],
        ],
        columns=[
            "Unnamed: 0",
            "Unnamed: 1",
            "Unnamed: 2",
            "Unnamed: 3",
            "Unnamed: 4",
            "Unnamed: 5",
            "Unnamed: 6",
            "Unnamed: 7",
            "Unnamed: 8",
            "Unnamed: 9",
        ],
    )
