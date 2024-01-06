use pyo3::create_exception;
use pyo3::exceptions::PyException;

mod cell;
mod sheet;
mod workbook;
pub use cell::CellValue;
pub use sheet::{CalamineSheet, SheetMetadata, SheetTypeEnum, SheetVisibleEnum};
pub use workbook::CalamineWorkbook;

create_exception!(python_calamine, CalamineError, PyException);
create_exception!(python_calamine, PasswordError, CalamineError);
create_exception!(python_calamine, WorksheetNotFound, CalamineError);
create_exception!(python_calamine, XmlError, CalamineError);
create_exception!(python_calamine, ZipError, CalamineError);
