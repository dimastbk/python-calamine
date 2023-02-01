use std::fs::File;
use std::io::BufReader;

use calamine::{open_workbook_auto, Error, Reader, Sheets};
use pyo3::prelude::*;
use pyo3::types::PyType;

use crate::utils::convert_err_to_py;
use crate::{CalamineError, CalamineSheet};

#[pyclass]
pub struct CalamineReader {
    sheets: Sheets<BufReader<File>>,
    #[pyo3(get)]
    sheet_names: Vec<String>,
}

#[pymethods]
impl CalamineReader {
    #[classmethod]
    fn from_path(_cls: &PyType, path: &str) -> PyResult<Self> {
        let sheets: Sheets<_> = open_workbook_auto(path).map_err(convert_err_to_py)?;
        let sheet_names = sheets.sheet_names().to_owned();
        Ok(Self {
            sheets,
            sheet_names,
        })
    }
    fn get_sheet_by_name(&mut self, name: &str) -> PyResult<CalamineSheet> {
        let range = self
            .sheets
            .worksheet_range(name)
            .unwrap_or_else(|| Err(Error::Msg("Workbook is empty")))
            .map_err(convert_err_to_py)?;
        Ok(CalamineSheet::new(name.to_owned(), range))
    }
    fn get_sheet_by_index(&mut self, index: usize) -> PyResult<CalamineSheet> {
        let name = self
            .sheet_names
            .get(index)
            .ok_or_else(|| CalamineError::new_err("Workbook is empty"))?
            .to_string();
        let range = self
            .sheets
            .worksheet_range_at(index)
            .unwrap_or_else(|| Err(Error::Msg("Workbook is empty")))
            .map_err(convert_err_to_py)?;
        Ok(CalamineSheet::new(name, range))
    }
}
