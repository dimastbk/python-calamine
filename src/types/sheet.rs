use calamine::{DataType, Range};
use pyo3::prelude::*;

use crate::CellValue;

#[pyclass]
pub struct CalamineSheet {
    #[pyo3(get)]
    name: String,
    range: Range<DataType>,
}

impl CalamineSheet {
    pub fn new(name: String, range: Range<DataType>) -> Self {
        CalamineSheet { name, range }
    }
}

#[pymethods]
impl CalamineSheet {
    #[getter]
    fn height(&self) -> usize {
        self.range.height()
    }

    #[getter]
    fn width(&self) -> usize {
        self.range.height()
    }

    #[getter]
    fn total_height(&self) -> u32 {
        self.range.end().unwrap_or_default().0
    }

    #[getter]
    fn total_width(&self) -> u32 {
        self.range.end().unwrap_or_default().1
    }

    #[getter]
    fn start(&self) -> Option<(u32, u32)> {
        self.range.start()
    }

    #[getter]
    fn end(&self) -> Option<(u32, u32)> {
        self.range.end()
    }

    #[pyo3(signature = (skip_empty_area=true, nrows=None))]
    fn to_python(
        &self,
        skip_empty_area: bool,
        nrows: Option<u32>,
    ) -> PyResult<Vec<Vec<CellValue>>> {
        let mut range = self.range.to_owned();

        if !skip_empty_area {
            if let Some(end) = range.end() {
                range = range.range((0, 0), end)
            }
        }

        if let Some(nrows) = nrows {
            if self.range.end().is_some() && self.range.start().is_some() {
                range = range.range(
                    self.range.start().unwrap(),
                    (
                        self.range.start().unwrap().0 + nrows,
                        self.range.start().unwrap().1,
                    ),
                )
            }
        }

        Ok(range
            .rows()
            .map(|row| row.iter().map(|x| x.into()).collect())
            .collect())
    }
}
