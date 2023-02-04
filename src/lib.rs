use calamine::{open_workbook_auto, Error, Reader, Sheets};
use pyo3::prelude::*;
use pyo3::wrap_pyfunction;

mod types;
mod utils;
use crate::types::{CalamineError, CalamineReader, CalamineSheet, CellValue};
use crate::utils::convert_err_to_py;

#[pyfunction]
#[pyo3(signature = (path, sheet, skip_empty_area=true))]
fn get_sheet_data(
    path: &str,
    sheet: usize,
    skip_empty_area: bool,
) -> PyResult<Vec<Vec<CellValue>>> {
    let mut excel: Sheets<_> = open_workbook_auto(path).map_err(convert_err_to_py)?;
    let readed_range = excel.worksheet_range_at(sheet);
    let mut range = readed_range
        .unwrap_or_else(|| Err(Error::Msg("Workbook is empty")))
        .map_err(convert_err_to_py)?;
    if !skip_empty_area {
        if let Some(end) = range.end() {
            range = range.range((0, 0), end)
        }
    }
    Ok(range
        .rows()
        .map(|row| row.iter().map(|x| x.into()).collect())
        .collect())
}

#[pyfunction]
fn get_sheet_names(path: &str) -> PyResult<Vec<String>> {
    let excel: Sheets<_> = open_workbook_auto(path).map_err(convert_err_to_py)?;
    Ok(excel.sheet_names().to_vec())
}

#[pymodule]
fn _python_calamine(py: Python, m: &PyModule) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(get_sheet_data, m)?)?;
    m.add_function(wrap_pyfunction!(get_sheet_names, m)?)?;
    m.add_class::<CalamineReader>()?;
    m.add_class::<CalamineSheet>()?;
    m.add("CalamineError", py.get_type::<CalamineError>())?;
    Ok(())
}
