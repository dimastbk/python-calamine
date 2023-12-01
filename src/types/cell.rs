use std::convert::From;

use calamine::DataType;
use pyo3::prelude::*;

/// https://learn.microsoft.com/en-us/office/troubleshoot/excel/1900-and-1904-date-system
static EXCEL_1900_1904_DIFF: f64 = 1462.0;

#[derive(Debug)]
pub enum CellValue {
    Int(i64),
    Float(f64),
    String(String),
    Time(chrono::NaiveTime),
    Date(chrono::NaiveDate),
    DateTime(chrono::NaiveDateTime),
    Timedelta(chrono::Duration),
    Bool(bool),
    Empty,
}

impl IntoPy<PyObject> for CellValue {
    fn into_py(self, py: Python) -> PyObject {
        self.to_object(py)
    }
}

impl ToPyObject for CellValue {
    fn to_object(&self, py: Python<'_>) -> PyObject {
        match self {
            CellValue::Int(v) => v.to_object(py),
            CellValue::Float(v) => v.to_object(py),
            CellValue::String(v) => v.to_object(py),
            CellValue::Bool(v) => v.to_object(py),
            CellValue::Time(v) => v.to_object(py),
            CellValue::Date(v) => v.to_object(py),
            CellValue::DateTime(v) => v.to_object(py),
            CellValue::Timedelta(v) => v.to_object(py),
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
                // FIXME: need to fix after fixing in calamine
                if v < &1.0 || (*v - EXCEL_1900_1904_DIFF < 1.0 && *v - EXCEL_1900_1904_DIFF > 0.0)
                {
                    value.as_time().map(CellValue::Time)
                } else if *v == (*v as u64) as f64 {
                    value.as_date().map(CellValue::Date)
                } else {
                    value.as_datetime().map(CellValue::DateTime)
                }
            }
            .unwrap_or(CellValue::Float(v.to_owned())),
            DataType::DateTimeIso(v) => {
                if v.contains('T') {
                    value.as_datetime().map(CellValue::DateTime)
                } else if v.contains(':') {
                    value.as_time().map(CellValue::Time)
                } else {
                    value.as_date().map(CellValue::Date)
                }
            }
            .unwrap_or(CellValue::String(v.to_owned())),
            DataType::Duration(v) => value
                .as_duration()
                .map(CellValue::Timedelta)
                .unwrap_or(CellValue::Float(v.to_owned())),
            DataType::DurationIso(v) => value
                .as_time()
                .map(CellValue::Time)
                .unwrap_or(CellValue::String(v.to_owned())),
            DataType::Bool(v) => CellValue::Bool(v.to_owned()),
            DataType::Error(_) => CellValue::Empty,
            DataType::Empty => CellValue::Empty,
        }
    }
}
