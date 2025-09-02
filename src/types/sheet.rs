use std::fmt::Display;
use std::sync::Arc;

use calamine::{Data, Dimensions, Range, Rows, SheetType, SheetVisible};
use pyo3::class::basic::CompareOp;
use pyo3::prelude::*;
use pyo3::types::PyList;

use crate::CellValue;

#[pyclass(eq, eq_int)]
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
        write!(f, "SheetTypeEnum.{self:?}")
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

#[pyclass(eq, eq_int)]
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
        write!(f, "SheetVisibleEnum.{self:?}")
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

type MergedCellRange = ((u32, u32), (u32, u32));

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

    fn __richcmp__(&self, other: &Self, op: CompareOp, py: Python<'_>) -> PyResult<Py<PyAny>> {
        match op {
            CompareOp::Eq => Ok(self
                .eq(other)
                .into_pyobject(py)
                .map_err(Into::<PyErr>::into)?
                .to_owned()
                .into_any()
                .unbind()),
            CompareOp::Ne => Ok(self
                .ne(other)
                .into_pyobject(py)
                .map_err(Into::<PyErr>::into)?
                .to_owned()
                .into_any()
                .unbind()),
            _ => Ok(py.NotImplemented()),
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
    range: Arc<Range<Data>>,
    formula_range: Option<Arc<Range<String>>>,
    merged_cell_ranges: Option<Vec<Dimensions>>,
}

impl CalamineSheet {
    pub fn new(
        name: String,
        range: Range<Data>,
        formula_range: Option<Range<String>>,
        merged_cell_ranges: Option<Vec<Dimensions>>,
    ) -> Self {
        CalamineSheet {
            name,
            range: Arc::new(range),
            formula_range: formula_range.map(Arc::new),
            merged_cell_ranges,
        }
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
        slf: PyRef<'_, Self>,
        skip_empty_area: bool,
        nrows: Option<u32>,
    ) -> PyResult<Bound<'_, PyList>> {
        let nrows = match nrows {
            Some(nrows) => nrows,
            None => slf.range.end().map_or(0, |end| end.0 + 1),
        };

        let range = if skip_empty_area || Some((0, 0)) == slf.range.start() {
            Arc::clone(&slf.range)
        } else if let Some(end) = slf.range.end() {
            Arc::new(slf.range.range(
                (0, 0),
                (if nrows > end.0 { end.0 } else { nrows - 1 }, end.1),
            ))
        } else {
            Arc::clone(&slf.range)
        };

        PyList::new(
            slf.py(),
            range.rows().take(nrows as usize).map(|row| {
                PyList::new(slf.py(), row.iter().map(<&Data as Into<CellValue>>::into)).unwrap()
            }),
        )
    }

    fn iter_rows(&self) -> CalamineCellIterator {
        CalamineCellIterator::from_range(Arc::clone(&self.range))
    }

    fn iter_formulas(&self) -> PyResult<CalamineFormulaIterator> {
        match &self.formula_range {
            Some(formula_range) => {
                let data_start = self.range.start().unwrap_or((0, 0));
                let data_end = self.range.end().unwrap_or((0, 0));
                Ok(CalamineFormulaIterator::from_range_with_data_bounds(
                    Arc::clone(formula_range),
                    data_start,
                    data_end
                ))
            },
            None => Err(pyo3::exceptions::PyValueError::new_err(
                "Formula iteration is disabled. Use read_formulas=True when creating the workbook to enable formula access."
            )),
        }
    }

    #[getter]
    fn merged_cell_ranges(slf: PyRef<'_, Self>) -> Option<Vec<MergedCellRange>> {
        slf.merged_cell_ranges
            .as_ref()
            .map(|r| r.iter().map(|d| (d.start, d.end)).collect())
    }
}

#[pyclass]
pub struct CalamineCellIterator {
    #[pyo3(get)]
    position: u32,
    #[pyo3(get)]
    start: (u32, u32),
    empty_row: Vec<CellValue>,
    iter: Rows<'static, Data>,
    #[allow(dead_code)]
    range: Arc<Range<Data>>,
}

impl CalamineCellIterator {
    fn from_range(range: Arc<Range<Data>>) -> CalamineCellIterator {
        let empty_row = (0..range.width())
            .map(|_| CellValue::String("".to_string()))
            .collect();
        CalamineCellIterator {
            empty_row,
            position: 0,
            start: range.start().unwrap_or((0, 0)),
            iter: unsafe {
                std::mem::transmute::<
                    calamine::Rows<'_, calamine::Data>,
                    calamine::Rows<'static, calamine::Data>,
                >(range.rows())
            },
            range,
        }
    }
}

#[pymethods]
impl CalamineCellIterator {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__(mut slf: PyRefMut<'_, Self>) -> PyResult<Option<Bound<'_, PyList>>> {
        slf.position += 1;
        if slf.position > slf.start.0 {
            slf.iter
                .next()
                .map(|row| PyList::new(slf.py(), row.iter().map(<&Data as Into<CellValue>>::into)))
                .transpose()
        } else {
            Some(PyList::new(slf.py(), slf.empty_row.clone())).transpose()
        }
    }

    #[getter]
    fn height(&self) -> usize {
        self.range.height()
    }
    #[getter]
    fn width(&self) -> usize {
        self.range.width()
    }
}

#[pyclass]
pub struct CalamineFormulaIterator {
    #[pyo3(get)]
    position: u32,
    #[pyo3(get)]
    start: (u32, u32),
    #[allow(dead_code)]
    range: Arc<Range<String>>,
    // Data range dimensions for coordinate mapping
    #[pyo3(get)]
    width: usize,
    #[pyo3(get)]
    height: usize,
}

impl CalamineFormulaIterator {
    fn from_range_with_data_bounds(
        range: Arc<Range<String>>,
        data_start: (u32, u32),
        data_end: (u32, u32),
    ) -> CalamineFormulaIterator {
        let width = if data_start <= data_end {
            (data_end.1 - data_start.1 + 1) as usize
        } else {
            0
        };

        let height = if data_start <= data_end {
            (data_end.0 - data_start.0 + 1) as usize
        } else {
            0
        };

        CalamineFormulaIterator {
            position: 0,
            start: data_start,
            range,
            width,
            height,
        }
    }
}

#[pymethods]
impl CalamineFormulaIterator {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__(mut slf: PyRefMut<'_, Self>) -> PyResult<Option<Bound<'_, PyList>>> {
        // Check if we've exceeded the data range height
        if slf.position >= slf.height as u32 {
            return Ok(None);
        }

        // Calculate the current absolute row position
        let current_row = slf.start.0 + slf.position;
        slf.position += 1;

        // Create the result row with proper width, filled with empty strings
        let mut result_row = vec!["".to_string(); slf.width];

        // Fill in formulas for this row by checking each column position
        for (col_idx, result_cell) in result_row.iter_mut().enumerate() {
            let current_col = slf.start.1 + col_idx as u32;
            if let Some(formula) = slf.range.get_value((current_row, current_col)) {
                if !formula.is_empty() {
                    *result_cell = formula.clone();
                }
            }
        }

        Some(PyList::new(slf.py(), result_row)).transpose()
    }
}
