use pyo3::prelude::*;

#[derive(Debug)]
pub enum CellValue {
    Int(i64),
    Float(f64),
    String(String),
    DateTime(f64),
    Bool(bool),
    None,
}

impl IntoPy<PyObject> for CellValue {
    fn into_py(self, py: Python) -> PyObject {
        match self {
            CellValue::Int(v) => v.to_object(py),
            CellValue::Float(v) => v.to_object(py),
            CellValue::String(v) => v.to_object(py),
            CellValue::Bool(v) => v.to_object(py),
            CellValue::DateTime(v) => v.to_object(py),
            CellValue::None => "".to_object(py),
        }
    }
}
