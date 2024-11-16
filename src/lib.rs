use pyo3::prelude::*;

mod types;
mod utils;
use crate::types::{
    CalamineError, CalamineSheet, CalamineWorkbook, CellValue, Error, PasswordError, SheetMetadata,
    SheetTypeEnum, SheetVisibleEnum, WorkbookClosed, WorksheetNotFound, XmlError, ZipError,
};

#[pyfunction]
fn load_workbook(py: Python, path_or_filelike: PyObject) -> PyResult<CalamineWorkbook> {
    CalamineWorkbook::from_object(py, path_or_filelike)
}

#[pymodule(gil_used = false)]
fn _python_calamine(py: Python, m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(load_workbook, m)?)?;
    m.add_class::<CalamineWorkbook>()?;
    m.add_class::<CalamineSheet>()?;
    m.add_class::<SheetMetadata>()?;
    m.add_class::<SheetTypeEnum>()?;
    m.add_class::<SheetVisibleEnum>()?;
    m.add("CalamineError", py.get_type::<CalamineError>())?;
    m.add("PasswordError", py.get_type::<PasswordError>())?;
    m.add("WorksheetNotFound", py.get_type::<WorksheetNotFound>())?;
    m.add("XmlError", py.get_type::<XmlError>())?;
    m.add("ZipError", py.get_type::<ZipError>())?;
    m.add("WorkbookClosed", py.get_type::<WorkbookClosed>())?;
    Ok(())
}
