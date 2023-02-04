use calamine::Error;
use pyo3::exceptions::PyIOError;
use pyo3::PyErr;

use crate::types::CalamineError;

pub fn convert_err_to_py(e: Error) -> PyErr {
    match e {
        Error::Io(err) => PyIOError::new_err(err.to_string()),
        _ => CalamineError::new_err(e.to_string()),
    }
}
