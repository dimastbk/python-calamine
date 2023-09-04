use std::fmt::Display;

use calamine::{DataType, Range, SheetType, SheetVisible};
use pyo3::class::basic::CompareOp;
use pyo3::prelude::*;

use crate::CellValue;

#[pyclass]
#[derive(Clone, Debug, PartialEq)]
pub enum SheetTypeEnum {
    /// WorkSheet
    WorkSheet,
    /// DialogSheet
    DialogSheet,
    /// MacroSheet
    MacroSheet,
    /// ChartSheet
    ChartSheet,
    /// VBA module
    Vba,
}

impl Display for SheetTypeEnum {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "SheetTypeEnum.{:?}", self)
    }
}

impl From<SheetType> for SheetTypeEnum {
    fn from(value: SheetType) -> Self {
        match value {
            SheetType::WorkSheet => Self::WorkSheet,
            SheetType::DialogSheet => Self::DialogSheet,
            SheetType::MacroSheet => Self::MacroSheet,
            SheetType::ChartSheet => Self::ChartSheet,
            SheetType::Vba => Self::Vba,
        }
    }
}

#[pyclass]
#[derive(Clone, Debug, PartialEq)]
pub enum SheetVisibleEnum {
    /// Visible
    Visible,
    /// Hidden
    Hidden,
    /// The sheet is hidden and cannot be displayed using the user interface. It is supported only by Excel formats.
    VeryHidden,
}

impl Display for SheetVisibleEnum {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "SheetVisibleEnum.{:?}", self)
    }
}

impl From<SheetVisible> for SheetVisibleEnum {
    fn from(value: SheetVisible) -> Self {
        match value {
            SheetVisible::Visible => Self::Visible,
            SheetVisible::Hidden => Self::Hidden,
            SheetVisible::VeryHidden => Self::VeryHidden,
        }
    }
}

#[pyclass]
#[derive(Clone, PartialEq)]
pub struct SheetMetadata {
    #[pyo3(get)]
    name: String,
    #[pyo3(get)]
    typ: SheetTypeEnum,
    #[pyo3(get)]
    visible: SheetVisibleEnum,
}

#[pymethods]
impl SheetMetadata {
    // implementation of some methods for testing
    #[new]
    fn py_new(name: &str, typ: SheetTypeEnum, visible: SheetVisibleEnum) -> Self {
        SheetMetadata {
            name: name.to_string(),
            typ,
            visible,
        }
    }

    fn __repr__(&self) -> PyResult<String> {
        Ok(format!(
            "SheetMetadata(name='{}', typ={}, visible={})",
            self.name, self.typ, self.visible
        ))
    }

    fn __richcmp__(&self, other: &Self, op: CompareOp, py: Python<'_>) -> PyObject {
        match op {
            CompareOp::Eq => self.eq(other).into_py(py),
            CompareOp::Ne => self.ne(other).into_py(py),
            _ => py.NotImplemented(),
        }
    }
}

impl SheetMetadata {
    pub fn new(name: String, typ: SheetType, visible: SheetVisible) -> Self {
        let typ = SheetTypeEnum::from(typ);
        let visible = SheetVisibleEnum::from(visible);
        SheetMetadata { name, typ, visible }
    }
}

#[pyclass]
pub struct CalamineSheet {
    #[pyo3(get)]
    name: String,
    range: Range<DataType>,
}

impl CalamineSheet {
    pub fn new(name: String, range: Range<DataType>) -> Self {
        CalamineSheet { name, range }
    }
}

#[pymethods]
impl CalamineSheet {
    fn __repr__(&self) -> PyResult<String> {
        Ok(format!("CalamineSheet(name='{}')", self.name))
    }

    #[getter]
    fn height(&self) -> usize {
        self.range.height()
    }

    #[getter]
    fn width(&self) -> usize {
        self.range.width()
    }

    #[getter]
    fn total_height(&self) -> u32 {
        self.range.end().unwrap_or_default().0
    }

    #[getter]
    fn total_width(&self) -> u32 {
        self.range.end().unwrap_or_default().1
    }

    #[getter]
    fn start(&self) -> Option<(u32, u32)> {
        self.range.start()
    }

    #[getter]
    fn end(&self) -> Option<(u32, u32)> {
        self.range.end()
    }

    #[pyo3(signature = (skip_empty_area=true, nrows=None))]
    fn to_python(
        &self,
        skip_empty_area: bool,
        nrows: Option<u32>,
    ) -> PyResult<Vec<Vec<CellValue>>> {
        let mut range = self.range.to_owned();

        if !skip_empty_area {
            if let Some(end) = range.end() {
                range = range.range((0, 0), end)
            }
        }

        if let Some(nrows) = nrows {
            if range.end().is_some() && range.start().is_some() {
                range = range.range(
                    range.start().unwrap(),
                    (
                        range.start().unwrap().0 + (nrows - 1),
                        range.end().unwrap().1,
                    ),
                )
            }
        }

        Ok(range
            .rows()
            .map(|row| row.iter().map(|x| x.into()).collect())
            .collect())
    }
}
