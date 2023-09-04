use std::fs::File;
use std::io::{BufReader, Cursor, Read};
use std::path::PathBuf;

use calamine::{open_workbook_auto, open_workbook_auto_from_rs, Error, Reader, Sheets};
use pyo3::prelude::*;
use pyo3::types::{PyString, PyType};
use pyo3_file::PyFileLikeObject;

use crate::utils::err_to_py;
use crate::{CalamineError, CalamineSheet, SheetMetadata};

enum SheetsEnum {
    File(Sheets<BufReader<File>>),
    FileLike(Sheets<Cursor<Vec<u8>>>),
}

impl SheetsEnum {
    fn sheets_metadata(&self) -> Vec<SheetMetadata> {
        match self {
            SheetsEnum::File(f) => f.sheets_metadata(),
            SheetsEnum::FileLike(f) => f.sheets_metadata(),
        }
        .iter()
        .map(|s| SheetMetadata::new(s.name.clone(), s.typ, s.visible))
        .collect()
    }

    fn sheet_names(&self) -> Vec<String> {
        match self {
            SheetsEnum::File(f) => f.sheet_names(),
            SheetsEnum::FileLike(f) => f.sheet_names(),
        }
    }

    fn worksheet_range(
        &mut self,
        name: &str,
    ) -> Option<Result<calamine::Range<calamine::DataType>, Error>> {
        match self {
            SheetsEnum::File(f) => f.worksheet_range(name),
            SheetsEnum::FileLike(f) => f.worksheet_range(name),
        }
    }

    fn worksheet_range_at(
        &mut self,
        index: usize,
    ) -> Option<Result<calamine::Range<calamine::DataType>, Error>> {
        match self {
            SheetsEnum::File(f) => f.worksheet_range_at(index),
            SheetsEnum::FileLike(f) => f.worksheet_range_at(index),
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
    fn py_from_object(_cls: &PyType, path_or_filelike: PyObject) -> PyResult<Self> {
        Self::from_object(path_or_filelike)
    }

    #[classmethod]
    #[pyo3(name = "from_filelike")]
    fn py_from_filelike(_cls: &PyType, filelike: PyObject) -> PyResult<Self> {
        Self::from_filelike(filelike)
    }

    #[classmethod]
    #[pyo3(name = "from_path")]
    fn py_from_path(_cls: &PyType, path: &str) -> PyResult<Self> {
        Self::from_path(path)
    }

    fn get_sheet_by_name(&mut self, name: &str) -> PyResult<CalamineSheet> {
        let range = self
            .sheets
            .worksheet_range(name)
            .unwrap_or_else(|| Err(Error::Msg("Workbook is empty")))
            .map_err(err_to_py)?;
        Ok(CalamineSheet::new(name.to_owned(), range))
    }

    fn get_sheet_by_index(&mut self, index: usize) -> PyResult<CalamineSheet> {
        let name = self
            .sheet_names
            .get(index)
            .ok_or_else(|| CalamineError::new_err("Workbook is empty"))?
            .to_string();
        let range = self
            .sheets
            .worksheet_range_at(index)
            .unwrap_or_else(|| Err(Error::Msg("Workbook is empty")))
            .map_err(err_to_py)?;
        Ok(CalamineSheet::new(name, range))
    }
}

impl CalamineWorkbook {
    pub fn from_object(path_or_filelike: PyObject) -> PyResult<Self> {
        Python::with_gil(|py| {
            if let Ok(string_ref) = path_or_filelike.downcast::<PyString>(py) {
                return Self::from_path(string_ref.to_string_lossy().to_string().as_str());
            }

            if let Ok(string_ref) = path_or_filelike.extract::<PathBuf>(py) {
                return Self::from_path(string_ref.to_string_lossy().to_string().as_str());
            }

            Self::from_filelike(path_or_filelike)
        })
    }

    pub fn from_filelike(filelike: PyObject) -> PyResult<Self> {
        let mut buf = vec![];
        PyFileLikeObject::with_requirements(filelike, true, false, true)?.read_to_end(&mut buf)?;
        let reader = Cursor::new(buf);
        let sheets = SheetsEnum::FileLike(open_workbook_auto_from_rs(reader).map_err(err_to_py)?);
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
        let sheets = SheetsEnum::File(open_workbook_auto(path).map_err(err_to_py)?);
        let sheet_names = sheets.sheet_names().to_owned();
        let sheets_metadata = sheets.sheets_metadata().to_owned();

        Ok(Self {
            path: Some(path.to_string()),
            sheets,
            sheets_metadata,
            sheet_names,
        })
    }
}
