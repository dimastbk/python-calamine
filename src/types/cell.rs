use std::convert::From;

use calamine::DataType;
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

impl From<&DataType> for CellValue {
    fn from(value: &DataType) -> Self {
        match value {
            DataType::Int(v) => CellValue::Int(v.to_owned()),
            DataType::Float(v) => CellValue::Float(v.to_owned()),
            DataType::String(v) => CellValue::String(String::from(v)),
            DataType::DateTime(v) => {
                if v < &1.0 {
                    value.as_time().map(CellValue::Time)
                } else if *v == (*v as u64) as f64 {
                    value.as_date().map(CellValue::Date)
                } else {
                    value.as_datetime().map(CellValue::DateTime)
                }
            }
            .unwrap_or(CellValue::Empty),
            DataType::Bool(v) => CellValue::Bool(v.to_owned()),
            DataType::Error(_) => CellValue::Empty,
            DataType::Empty => CellValue::Empty,
        }
    }
}
