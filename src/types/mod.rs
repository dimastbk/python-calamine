use pyo3::create_exception;
use pyo3::exceptions::PyException;

mod cell;
mod reader;
mod sheet;
pub use cell::CellValue;
pub use reader::CalamineReader;
pub use sheet::CalamineSheet;

create_exception!(python_calamine, CalamineError, PyException);
