# python-calamine

Python binding for beautiful Rust's library for reading excel and odf file - [calamine](https://github.com/tafia/calamine).

### Is used
* [calamine](https://github.com/tafia/calamine)
* [pyo3](https://github.com/PyO3/pyo3)
* [maturin](https://github.com/PyO3/maturin)

### Installation
```
pip install python-calamine
```

### Example
```python
from python_calamine import get_sheet_data, get_sheet_names


get_sheet_names("file.xlsx")
# ['Sheet1', 'Sheet2']

get_sheet_data("file.xlsx", "Sheet1")
# [
# ['1',  '2',  '3',  '4',  '5',  '6',  '7'],
# ['1',  '2',  '3',  '4',  '5',  '6',  '7'],
# ['1',  '2',  '3',  '4',  '5',  '6',  '7'],
# ]
```

By default, calamine skips empty rows/cols before data. For suppress this behaviour, set `skip_empty_area` to `False`.
```python
from python_calamine import get_sheet_data

get_sheet_data("file.xlsx", "Sheet1", skip_empty_area=False)
# [
# ['',  '',  '',  '',  '',  '',  ''],
# ['1',  '2',  '3',  '4',  '5',  '6',  '7'],
# ['1',  '2',  '3',  '4',  '5',  '6',  '7'],
# ['1',  '2',  '3',  '4',  '5',  '6',  '7'],
# ]
```

Also, you can use monkeypatch for pandas for use this library as engine in `read_excel()`.
```python
from pandas import read_excel
from python_calamine.pandas import pandas_monkeypatch


pandas_monkeypatch()

read_excel("file.xlsx", engine="calamine")
#            1   2   3   4   5   6   7
# 0          1   2   3   4   5   6   7
# 1          1   2   3   4   5   6   7
```

Also, you can find additional examples in [tests](https://github.com/dimastbk/python-calamine/blob/master/tests/test_base.py).
