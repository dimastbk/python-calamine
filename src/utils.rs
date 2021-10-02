use chrono;
use chrono::{Datelike, Timelike};
use pyo3::prelude::*;
use pyo3::types::{PyDate, PyDateTime, PyTime};

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
            CellValue::Time(v) => PyTime::new(
                py,
                v.hour() as u8,
                v.minute() as u8,
                v.second() as u8,
                (v.nanosecond() / 1000) as u32,
                None,
            )
            .unwrap()
            .to_object(py),
            CellValue::Date(v) => PyDate::new(py, v.year(), v.month() as u8, v.day() as u8)
                .unwrap()
                .to_object(py),
            CellValue::DateTime(v) => PyDateTime::new(
                py,
                v.year(),
                v.month() as u8,
                v.day() as u8,
                v.hour() as u8,
                v.minute() as u8,
                v.second() as u8,
                v.timestamp_subsec_micros(),
                None,
            )
            .unwrap()
            .to_object(py),
            CellValue::Empty => "".to_object(py),
        }
    }
}
