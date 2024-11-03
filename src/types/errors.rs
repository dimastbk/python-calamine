use calamine::Error as CalamineCrateError;
use pyo3::create_exception;
use pyo3::exceptions::PyException;

#[derive(Debug)]
pub enum Error {
    Calamine(CalamineCrateError),
    WorkbookClosed,
}

create_exception!(python_calamine, CalamineError, PyException);
create_exception!(python_calamine, PasswordError, CalamineError);
create_exception!(python_calamine, WorksheetNotFound, CalamineError);
create_exception!(python_calamine, XmlError, CalamineError);
create_exception!(python_calamine, ZipError, CalamineError);
create_exception!(python_calamine, WorkbookClosed, CalamineError);
