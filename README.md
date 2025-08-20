# python-calamine
[![PyPI - Version](https://img.shields.io/pypi/v/python-calamine)](https://pypi.org/project/python-calamine/)
[![Conda Version](https://img.shields.io/conda/vn/conda-forge/python-calamine.svg)](https://anaconda.org/conda-forge/python-calamine)
![Python Version from PEP 621 TOML](https://img.shields.io/python/required-version-toml?tomlFilePath=https%3A%2F%2Fraw.githubusercontent.com%2Fdimastbk%2Fpython-calamine%2Fmaster%2Fpyproject.toml)


Python binding for beautiful Rust's library for reading excel and odf file - [calamine](https://github.com/tafia/calamine).

### Is used
* [calamine](https://github.com/tafia/calamine)
* [pyo3](https://github.com/PyO3/pyo3)
* [maturin](https://github.com/PyO3/maturin)

### Installation
Pypi:
```
pip install python-calamine
```
Conda:
```
conda install -c conda-forge python-calamine
```

### Example
```python
from python_calamine import CalamineWorkbook

workbook = CalamineWorkbook.from_path("file.xlsx")
workbook.sheet_names
# ["Sheet1", "Sheet2"]

workbook.get_sheet_by_name("Sheet1").to_python()
# [
# ["1",  "2",  "3",  "4",  "5",  "6",  "7"],
# ["1",  "2",  "3",  "4",  "5",  "6",  "7"],
# ["1",  "2",  "3",  "4",  "5",  "6",  "7"],
# ]
```

By default, calamine skips empty rows/cols before data. For suppress this behaviour, set `skip_empty_area` to `False`.
```python
from python_calamine import CalamineWorkbook

workbook = CalamineWorkbook.from_path("file.xlsx").get_sheet_by_name("Sheet1").to_python(skip_empty_area=False)
# [
# [",  ",  ",  ",  ",  ",  "],
# ["1",  "2",  "3",  "4",  "5",  "6",  "7"],
# ["1",  "2",  "3",  "4",  "5",  "6",  "7"],
# ["1",  "2",  "3",  "4",  "5",  "6",  "7"],
# ]
```

Pandas 2.2 and above have built-in support of python-calamine.

Also, you can find additional examples in [tests](https://github.com/dimastbk/python-calamine/blob/master/tests/test_base.py).

### Development

You'll need rust [installed](https://rustup.rs/).

```python
# clone this repo or your fork
git clone git@github.com:dimastbk/python-calamine.git
cd python-calamine
# create a new virtual env
python3 -m venv env
source env/bin/activate
# install dev dependencies and install python-calamine
pip install -e .[dev]
# lint code
pre-commit run --all-files
# test code
pytest
```
