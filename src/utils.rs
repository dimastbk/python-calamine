use calamine::{Error as CalamineCrateError, OdsError, XlsError, XlsbError, XlsxError};
use pyo3::exceptions::PyIOError;
use pyo3::PyErr;

use crate::{
    CalamineError, Error, PasswordError, WorkbookClosed, WorksheetNotFound, XmlError, ZipError,
};

pub fn err_to_py(e: Error) -> PyErr {
    match e {
        Error::Calamine(calamine_error) => match calamine_error {
            CalamineCrateError::Io(err) => PyIOError::new_err(err.to_string()),
            CalamineCrateError::Ods(ref err) => match err {
                OdsError::Io(error) => PyIOError::new_err(error.to_string()),
                OdsError::Zip(error) => ZipError::new_err(error.to_string()),
                OdsError::Xml(error) => XmlError::new_err(error.to_string()),
                OdsError::XmlAttr(error) => XmlError::new_err(error.to_string()),
                OdsError::Password => PasswordError::new_err(err.to_string()),
                OdsError::WorksheetNotFound(error) => WorksheetNotFound::new_err(error.to_string()),
                _ => CalamineError::new_err(err.to_string()),
            },
            CalamineCrateError::Xls(ref err) => match err {
                XlsError::Io(error) => PyIOError::new_err(error.to_string()),
                XlsError::Password => PasswordError::new_err(err.to_string()),
                XlsError::WorksheetNotFound(error) => WorksheetNotFound::new_err(error.to_string()),
                _ => CalamineError::new_err(err.to_string()),
            },
            CalamineCrateError::Xlsx(ref err) => match err {
                XlsxError::Io(error) => PyIOError::new_err(error.to_string()),
                XlsxError::Zip(error) => ZipError::new_err(error.to_string()),
                XlsxError::Xml(error) => XmlError::new_err(error.to_string()),
                XlsxError::XmlAttr(error) => XmlError::new_err(error.to_string()),
                XlsxError::XmlEof(error) => XmlError::new_err(error.to_string()),
                XlsxError::Password => PasswordError::new_err(err.to_string()),
                XlsxError::WorksheetNotFound(error) => {
                    WorksheetNotFound::new_err(error.to_string())
                }
                _ => CalamineError::new_err(err.to_string()),
            },
            CalamineCrateError::Xlsb(ref err) => match err {
                XlsbError::Io(error) => PyIOError::new_err(error.to_string()),
                XlsbError::Zip(error) => ZipError::new_err(error.to_string()),
                XlsbError::Xml(error) => XmlError::new_err(error.to_string()),
                XlsbError::XmlAttr(error) => XmlError::new_err(error.to_string()),
                XlsbError::Password => PasswordError::new_err(err.to_string()),
                XlsbError::WorksheetNotFound(error) => {
                    WorksheetNotFound::new_err(error.to_string())
                }
                _ => CalamineError::new_err(err.to_string()),
            },
            _ => CalamineError::new_err(calamine_error.to_string()),
        },
        Error::WorkbookClosed => WorkbookClosed::new_err("".to_string()),
    }
}
