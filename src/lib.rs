use pyo3::prelude::*;

mod types;
use crate::types::{
    CalamineError, CalamineSheet, CalamineTable, CalamineWorkbook, CellValue, Error, PasswordError,
    SheetMetadata, SheetTypeEnum, SheetVisibleEnum, TableNotFound, TablesNotLoaded,
    TablesNotSupported, WorkbookClosed, WorksheetNotFound, XmlError, ZipError,
};

#[pyfunction]
#[pyo3(signature = (path_or_filelike, load_tables=false))]
fn load_workbook(
    py: Python,
    path_or_filelike: Py<PyAny>,
    load_tables: bool,
) -> PyResult<CalamineWorkbook> {
    CalamineWorkbook::from_object(py, path_or_filelike, load_tables)
}

#[pymodule]
fn _python_calamine(py: Python, m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(load_workbook, m)?)?;
    m.add_class::<CalamineWorkbook>()?;
    m.add_class::<CalamineSheet>()?;
    m.add_class::<SheetMetadata>()?;
    m.add_class::<SheetTypeEnum>()?;
    m.add_class::<SheetVisibleEnum>()?;
    m.add_class::<CalamineTable>()?;
    m.add("CalamineError", py.get_type::<CalamineError>())?;
    m.add("PasswordError", py.get_type::<PasswordError>())?;
    m.add("WorksheetNotFound", py.get_type::<WorksheetNotFound>())?;
    m.add("XmlError", py.get_type::<XmlError>())?;
    m.add("ZipError", py.get_type::<ZipError>())?;
    m.add("TablesNotSupported", py.get_type::<TablesNotSupported>())?;
    m.add("TablesNotLoaded", py.get_type::<TablesNotLoaded>())?;
    m.add("TableNotFound", py.get_type::<TableNotFound>())?;
    m.add("WorkbookClosed", py.get_type::<WorkbookClosed>())?;
    Ok(())
}
