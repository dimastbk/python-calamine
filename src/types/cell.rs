use std::convert::From;

use calamine::{Data, DataType};
use chrono::Datelike;
use num_bigint::BigInt;
use num_traits::FromPrimitive;
use pyo3::prelude::*;

/// https://docs.python.org/3/library/datetime.html#constants
/// The smallest year number allowed in a date or datetime object. MINYEAR is 1.
const MINYEAR: i32 = 1;

/// The largest year number allowed in a date or datetime object. MAXYEAR is 9999.
const MAXYEAR: i32 = 9999;

#[derive(Debug, Clone)]
pub enum CellValue {
    BigInt(BigInt),
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

fn check_year_range<DT: Datelike>(value: DT) -> Option<DT> {
    if value.year() < MINYEAR || value.year() > MAXYEAR {
        None
    } else {
        Some(value)
    }
}

pub fn convert_to_pandas_cell(data: &Data) -> CellValue {
    match data {
        // # GH#54564
        // # pandas casts x.0 floats to x int
        Data::Float(f) => {
            if f.is_finite() && !f.is_nan() && f.fract() == 0. {
                if *f >= i64::MIN as f64 && *f < i64::MAX as f64 {
                    CellValue::Int(*f as i64)
                } else {
                    CellValue::BigInt(BigInt::from_f64(*f).unwrap())
                }
            } else {
                data.into()
            }
        }
        // Return timedeltas and datetimes as-is to match openpyxl behavior (GH#59186)
        Data::DateTime(dt) => {
            let v = dt.as_f64();
            if dt.is_duration() {
                data.as_duration().map(CellValue::Timedelta)
            } else if v < 1.0 {
                data.as_time().map(CellValue::Time)
            } else {
                data.as_datetime()
                    .and_then(check_year_range)
                    .map(CellValue::DateTime)
            }
            .unwrap_or(CellValue::Float(v))
        }
        Data::DateTimeIso(v) => {
            if v.contains('T') {
                data.as_datetime()
                    .and_then(check_year_range)
                    .map(CellValue::DateTime)
            } else if v.contains(':') {
                data.as_time().map(CellValue::Time)
            } else {
                data.as_date()
                    .and_then(check_year_range)
                    .and_then(|date| date.and_hms_opt(0, 0, 0))
                    .map(CellValue::DateTime)
            }
        }
        .unwrap_or(CellValue::String(v.to_owned())),
        _ => data.into(),
    }
}

impl<'py> IntoPyObject<'py> for CellValue {
    type Target = PyAny;
    type Output = Bound<'py, Self::Target>;
    type Error = PyErr;

    fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
        match self {
            CellValue::BigInt(v) => Ok(v.into_pyobject(py)?.into_any()),
            CellValue::Int(v) => Ok(v.into_pyobject(py)?.into_any()),
            CellValue::Float(v) => Ok(v.into_pyobject(py)?.into_any()),
            CellValue::String(v) => Ok(v.into_pyobject(py)?.into_any()),
            CellValue::Bool(v) => Ok(v.into_pyobject(py)?.to_owned().into_any()),
            CellValue::Time(v) => Ok(v.into_pyobject(py)?.into_any()),
            CellValue::Date(v) => Ok(v.into_pyobject(py)?.into_any()),
            CellValue::DateTime(v) => Ok(v.into_pyobject(py)?.into_any()),
            CellValue::Timedelta(v) => Ok(v.into_pyobject(py)?.into_any()),
            CellValue::Empty => Ok("".into_pyobject(py)?.into_any()),
        }
    }
}

impl<DT> From<&DT> for CellValue
where
    DT: DataType,
{
    fn from(value: &DT) -> Self {
        if value.is_int() {
            value
                .get_int()
                .map(CellValue::Int)
                .unwrap_or(CellValue::Empty)
        } else if value.is_float() {
            value
                .get_float()
                .map(CellValue::Float)
                .unwrap_or(CellValue::Empty)
        } else if value.is_string() {
            value
                .get_string()
                .map(|s| CellValue::String(s.to_owned()))
                .unwrap_or(CellValue::Empty)
        } else if value.is_datetime() {
            let dt = value.get_datetime().unwrap();
            let v = dt.as_f64();
            if dt.is_duration() {
                value.as_duration().map(CellValue::Timedelta)
            } else if v < 1.0 {
                value.as_time().map(CellValue::Time)
            } else if v == (v as u64) as f64 {
                value
                    .as_date()
                    .and_then(check_year_range)
                    .map(CellValue::Date)
            } else {
                value
                    .as_datetime()
                    .and_then(check_year_range)
                    .map(CellValue::DateTime)
            }
            .unwrap_or(CellValue::Float(v))
        } else if value.is_datetime_iso() {
            let v = value.get_datetime_iso().unwrap();
            if v.contains('T') {
                value
                    .as_datetime()
                    .and_then(check_year_range)
                    .map(CellValue::DateTime)
            } else if v.contains(':') {
                value.as_time().map(CellValue::Time)
            } else {
                value
                    .as_date()
                    .and_then(check_year_range)
                    .map(CellValue::Date)
            }
            .unwrap_or(CellValue::String(v.to_owned()))
        } else if value.is_duration_iso() {
            value.as_time().map(CellValue::Time).unwrap_or(
                value
                    .get_duration_iso()
                    .map(|s| CellValue::String(s.to_owned()))
                    .unwrap_or(CellValue::Empty),
            )
        } else if value.is_bool() {
            value
                .get_bool()
                .map(CellValue::Bool)
                .unwrap_or(CellValue::Empty)
        } else {
            CellValue::Empty
        }
    }
}
