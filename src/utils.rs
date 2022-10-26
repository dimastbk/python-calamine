use chrono;
use pyo3::prelude::*;

#[derive(Debug)]
pub enum CellValue {
    Int(i64),
    Float(f64),
    String(String),
    Time(chrono::NaiveTime),
    Date(chrono::NaiveDate),
    DateTime(chrono::NaiveDateTime),
    Bool(bool),
    Empty,
}

impl IntoPy<PyObject> for CellValue {
    fn into_py(self, py: Python) -> PyObject {
        match self {
            CellValue::Int(v) => v.to_object(py),
            CellValue::Float(v) => v.to_object(py),
            CellValue::String(v) => v.to_object(py),
            CellValue::Bool(v) => v.to_object(py),
            CellValue::Time(v) => v.to_object(py),
            CellValue::Date(v) => v.to_object(py),
            CellValue::DateTime(v) => v.to_object(py),
            CellValue::Empty => "".to_object(py),
        }
    }
}
