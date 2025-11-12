use std::sync::Arc;

use calamine::{Data, Range};
use pyo3::prelude::*;
use pyo3::types::PyList;

use crate::CellValue;

#[pyclass]
#[derive(Clone, PartialEq)]
pub struct CalamineTable {
    #[pyo3(get)]
    name: String,
    #[pyo3(get)]
    sheet: String,
    #[pyo3(get)]
    columns: Vec<String>,
    range: Arc<Range<Data>>,
}

impl CalamineTable {
    pub fn new(name: String, sheet_name: String, columns: Vec<String>, range: Range<Data>) -> Self {
        CalamineTable {
            name,
            sheet: sheet_name,
            columns,
            range: Arc::new(range),
        }
    }
}

#[pymethods]
impl CalamineTable {
    fn __repr__(&self) -> PyResult<String> {
        Ok(format!("CalamineTable(name='{}')", self.name))
    }

    #[getter]
    fn height(&self) -> usize {
        self.range.height()
    }

    #[getter]
    fn width(&self) -> usize {
        self.range.width()
    }

    #[getter]
    fn start(&self) -> Option<(u32, u32)> {
        self.range.start()
    }

    #[getter]
    fn end(&self) -> Option<(u32, u32)> {
        self.range.end()
    }

    fn to_python(slf: PyRef<'_, Self>) -> PyResult<Bound<'_, PyList>> {
        let range = Arc::clone(&slf.range);

        let py_list = PyList::empty(slf.py());

        for row in range.rows() {
            let py_row = PyList::new(slf.py(), row.iter().map(<&Data as Into<CellValue>>::into))?;

            py_list.append(py_row)?;
        }

        Ok(py_list)
    }
}
