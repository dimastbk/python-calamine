from __future__ import annotations

from datetime import date, datetime, time

def get_sheet_data(
    path: str, sheet: int
) -> list[list[int | float | str | bool | time | date | datetime]]: ...
def get_sheet_names(path: str) -> list[str]: ...

class CalamineError(Exception): ...
