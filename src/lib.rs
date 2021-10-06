mod utils;

use calamine::{open_workbook_auto, DataType, Error, Reader, Sheets};
use pyo3::create_exception;
use pyo3::exceptions::*;
use pyo3::prelude::*;
use pyo3::wrap_pyfunction;

use utils::CellValue;

create_exception!(python_calamine, CalamineError, PyException);

fn _get_sheet_data(path: &str, sheet: usize) -> Result<Vec<Vec<CellValue>>, Error> {
    let mut excel: Sheets = open_workbook_auto(path)?;
    let range = excel.worksheet_range_at(sheet).unwrap()?;
    let mut result: Vec<Vec<CellValue>> = Vec::new();
    for row in range.rows() {
        let mut result_row: Vec<CellValue> = Vec::new();
        for value in row.iter() {
            match value {
                DataType::Int(v) => result_row.push(CellValue::Int(*v)),
                DataType::Float(v) => result_row.push(CellValue::Float(*v)),
                DataType::String(v) => result_row.push(CellValue::String(String::from(v))),
                DataType::DateTime(v) => {
                    if *v < 1.0 {
                        result_row.push(CellValue::Time(value.as_time().unwrap()))
                    } else if *v == (*v as u64) as f64 {
                        result_row.push(CellValue::Date(value.as_date().unwrap()))
                    } else {
                        result_row.push(CellValue::DateTime(value.as_datetime().unwrap()))
                    }
                }
                DataType::Bool(v) => result_row.push(CellValue::Bool(*v)),
                DataType::Error(_) => result_row.push(CellValue::Empty),
                DataType::Empty => result_row.push(CellValue::Empty),
            };
        }
        result.push(result_row);
    }
    Ok(result)
}

#[pyfunction]
#[pyo3(text_signature = "path: str, sheet: int")]
fn get_sheet_data(path: &str, sheet: usize) -> PyResult<Vec<Vec<CellValue>>> {
    match _get_sheet_data(path, sheet) {
        Ok(r) => Ok(r),
        Err(e) => match e {
            Error::Io(err) => Err(PyIOError::new_err(err.to_string())),
            _ => Err(CalamineError::new_err(e.to_string())),
        },
    }
}

fn _get_sheet_names(path: &str) -> Result<Vec<String>, Error> {
    let excel: Sheets = open_workbook_auto(path)?;
    Ok(excel.sheet_names().to_vec())
}

#[pyfunction]
#[pyo3(text_signature = "path: str")]
fn get_sheet_names(path: &str) -> PyResult<Vec<String>> {
    match _get_sheet_names(path) {
        Ok(r) => Ok(r),
        Err(e) => match e {
            Error::Io(err) => Err(PyIOError::new_err(err.to_string())),
            _ => Err(CalamineError::new_err(e.to_string())),
        },
    }
}

#[pymodule]
fn python_calamine(py: Python, m: &PyModule) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(get_sheet_data, m)?)?;
    m.add_function(wrap_pyfunction!(get_sheet_names, m)?)?;
    m.add("CalamineError", py.get_type::<CalamineError>())?;
    Ok(())
}
