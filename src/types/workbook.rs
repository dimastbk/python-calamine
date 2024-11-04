use std::fs::File;
use std::io::{BufReader, Cursor, Read};
use std::path::PathBuf;

use calamine::{open_workbook_auto, open_workbook_auto_from_rs, Reader, Sheets};
use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{PyString, PyType};
use pyo3_file::PyFileLikeObject;

use crate::utils::err_to_py;
use crate::{CalamineSheet, Error, SheetMetadata, WorksheetNotFound};

enum SheetsEnum {
    File(Sheets<BufReader<File>>),
    FileLike(Sheets<Cursor<Vec<u8>>>),
    None,
}

impl SheetsEnum {
    fn sheets_metadata(&self) -> Vec<SheetMetadata> {
        match self {
            SheetsEnum::File(f) => f.sheets_metadata(),
            SheetsEnum::FileLike(f) => f.sheets_metadata(),
            SheetsEnum::None => unreachable!(),
        }
        .iter()
        .map(|s| SheetMetadata::new(s.name.clone(), s.typ, s.visible))
        .collect()
    }

    fn sheet_names(&self) -> Vec<String> {
        match self {
            SheetsEnum::File(f) => f.sheet_names(),
            SheetsEnum::FileLike(f) => f.sheet_names(),
            SheetsEnum::None => unreachable!(),
        }
    }

    fn worksheet_range(&mut self, name: &str) -> Result<calamine::Range<calamine::Data>, Error> {
        match self {
            SheetsEnum::File(f) => f.worksheet_range(name).map_err(Error::Calamine),
            SheetsEnum::FileLike(f) => f.worksheet_range(name).map_err(Error::Calamine),
            SheetsEnum::None => Err(Error::WorkbookClosed),
        }
    }
}

#[pyclass]
pub struct CalamineWorkbook {
    #[pyo3(get)]
    path: Option<String>,
    sheets: SheetsEnum,
    #[pyo3(get)]
    sheets_metadata: Vec<SheetMetadata>,
    #[pyo3(get)]
    sheet_names: Vec<String>,
}

#[pymethods]
impl CalamineWorkbook {
    fn __repr__(&self) -> PyResult<String> {
        match &self.path {
            Some(path) => Ok(format!("CalamineWorkbook(path='{}')", path)),
            None => Ok("CalamineWorkbook(path='bytes')".to_string()),
        }
    }

    #[classmethod]
    #[pyo3(name = "from_object")]
    fn py_from_object(
        _cls: &Bound<'_, PyType>,
        py: Python<'_>,
        path_or_filelike: PyObject,
    ) -> PyResult<Self> {
        Self::from_object(py, path_or_filelike)
    }

    #[classmethod]
    #[pyo3(name = "from_filelike")]
    fn py_from_filelike(
        _cls: &Bound<'_, PyType>,
        py: Python<'_>,
        filelike: PyObject,
    ) -> PyResult<Self> {
        py.allow_threads(|| Self::from_filelike(filelike))
    }

    #[classmethod]
    #[pyo3(name = "from_path")]
    fn py_from_path(_cls: &Bound<'_, PyType>, py: Python<'_>, path: PyObject) -> PyResult<Self> {
        if let Ok(string_ref) = path.downcast_bound::<PyString>(py) {
            let path = string_ref.to_string_lossy().to_string();
            return py.allow_threads(|| Self::from_path(&path));
        }

        if let Ok(string_ref) = path.extract::<PathBuf>(py) {
            let path = string_ref.to_string_lossy().to_string();
            return py.allow_threads(|| Self::from_path(&path));
        }

        Err(PyTypeError::new_err(""))
    }

    #[pyo3(name = "get_sheet_by_name")]
    fn py_get_sheet_by_name(&mut self, py: Python<'_>, name: &str) -> PyResult<CalamineSheet> {
        py.allow_threads(|| self.get_sheet_by_name(name))
    }

    #[pyo3(name = "get_sheet_by_index")]
    fn py_get_sheet_by_index(&mut self, py: Python<'_>, index: usize) -> PyResult<CalamineSheet> {
        py.allow_threads(|| self.get_sheet_by_index(index))
    }

    fn close(&mut self) -> PyResult<()> {
        match self.sheets {
            SheetsEnum::None => Err(Error::WorkbookClosed),
            _ => {
                self.sheets = SheetsEnum::None;
                Ok(())
            }
        }
        .map_err(err_to_py)
    }

    fn __enter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __exit__(
        &mut self,
        _exc_type: PyObject,
        _exc_value: PyObject,
        _traceback: PyObject,
    ) -> PyResult<()> {
        self.close()
    }
}

impl CalamineWorkbook {
    pub fn from_object(py: Python<'_>, path_or_filelike: PyObject) -> PyResult<Self> {
        if let Ok(string_ref) = path_or_filelike.downcast_bound::<PyString>(py) {
            let path = string_ref.to_string_lossy().to_string();
            return py.allow_threads(|| Self::from_path(&path));
        }

        if let Ok(string_ref) = path_or_filelike.extract::<PathBuf>(py) {
            let path = string_ref.to_string_lossy().to_string();
            return py.allow_threads(|| Self::from_path(&path));
        }

        py.allow_threads(|| Self::from_filelike(path_or_filelike))
    }

    pub fn from_filelike(filelike: PyObject) -> PyResult<Self> {
        let mut buf = vec![];
        PyFileLikeObject::with_requirements(filelike, true, false, true, false)?
            .read_to_end(&mut buf)?;
        let reader = Cursor::new(buf);
        let sheets = SheetsEnum::FileLike(
            open_workbook_auto_from_rs(reader)
                .map_err(Error::Calamine)
                .map_err(err_to_py)?,
        );
        let sheet_names = sheets.sheet_names().to_owned();
        let sheets_metadata = sheets.sheets_metadata().to_owned();

        Ok(Self {
            path: None,
            sheets,
            sheets_metadata,
            sheet_names,
        })
    }

    pub fn from_path(path: &str) -> PyResult<Self> {
        let sheets = SheetsEnum::File(
            open_workbook_auto(path)
                .map_err(Error::Calamine)
                .map_err(err_to_py)?,
        );
        let sheet_names = sheets.sheet_names().to_owned();
        let sheets_metadata = sheets.sheets_metadata().to_owned();

        Ok(Self {
            path: Some(path.to_string()),
            sheets,
            sheets_metadata,
            sheet_names,
        })
    }

    fn get_sheet_by_name(&mut self, name: &str) -> PyResult<CalamineSheet> {
        let range = self.sheets.worksheet_range(name).map_err(err_to_py)?;
        Ok(CalamineSheet::new(name.to_owned(), range))
    }

    fn get_sheet_by_index(&mut self, index: usize) -> PyResult<CalamineSheet> {
        let name = self
            .sheet_names
            .get(index)
            .ok_or_else(|| WorksheetNotFound::new_err(format!("Worksheet '{}' not found", index)))?
            .to_string();
        self.get_sheet_by_name(&name)
    }
}
