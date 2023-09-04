from __future__ import annotations

import enum
from datetime import date, datetime, time, timedelta
from os import PathLike
from typing import Protocol

ValueT = int | float | str | bool | time | date | datetime | timedelta

class ReadBuffer(Protocol):
    def seek(self) -> int: ...
    def read(self) -> bytes: ...

class SheetTypeEnum(enum.Enum):
    WorkSheet = ...
    DialogSheet = ...
    MacroSheet = ...
    ChartSheet = ...
    VBA = ...

class SheetVisibleEnum(enum.Enum):
    Visible = ...
    Hidden = ...
    VeryHidden = ...

class SheetMetadata:
    name: str
    typ: SheetTypeEnum
    visible: SheetVisibleEnum

    def __init__(
        self, name: str, typ: SheetTypeEnum, visible: SheetVisibleEnum
    ) -> None: ...

class CalamineSheet:
    name: str
    @property
    def height(self) -> int: ...
    @property
    def width(self) -> int: ...
    @property
    def total_height(self) -> int: ...
    @property
    def total_width(self) -> int: ...
    @property
    def start(self) -> tuple[int, int] | None: ...
    @property
    def end(self) -> tuple[int, int] | None: ...
    def to_python(
        self, skip_empty_area: bool = True, nrows: int | None = None
    ) -> list[list[ValueT]]:
        """Retunrning data from sheet as list of lists.

        Parameters
        ----------
        skip_empty_area : bool
            By default, calamine skips empty rows/cols before data.
            For suppress this behaviour, set `skip_empty_area` to `False`.
        """

class CalamineWorkbook:
    path: str | None
    sheet_names: list[str]
    sheets_metadata: list[SheetMetadata]
    @classmethod
    def from_object(
        cls, path_or_filelike: str | PathLike | ReadBuffer
    ) -> "CalamineWorkbook":
        """Determining type of pyobject and reading from it.

        Parameters
        ----------
        path_or_filelike :
            - path (string),
            - pathlike (pathlib.Path),
            - IO (must imlpement read/seek methods).
        """
    @classmethod
    def from_path(cls, path: str) -> "CalamineWorkbook":
        """Reading file from path.

        Parameters
        ----------
        path : path (string)."""
    @classmethod
    def from_filelike(cls, filelike: ReadBuffer) -> "CalamineWorkbook":
        """Reading file from IO.

        Parameters
        ----------
        filelike : IO (must imlpement read/seek methods).
        """
    def get_sheet_by_name(self, name: str) -> CalamineSheet: ...
    def get_sheet_by_index(self, index: int) -> CalamineSheet: ...

class CalamineError(Exception): ...

def load_workbook(
    path_or_filelike: str | PathLike | ReadBuffer,
) -> CalamineWorkbook:
    """Determining type of pyobject and reading from it.

    Parameters
    ----------
    path_or_filelike :
        - path (string),
        - pathlike (pathlib.Path),
        - IO (must imlpement read/seek methods).
    """
