use std::fs::File;
use std::io::{BufReader, Cursor, Read};
use std::path::PathBuf;

use calamine::{
    open_workbook_auto, open_workbook_auto_from_rs, Error as CalamineCrateError, Reader, Sheets,
};
use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::PyType;
use pyo3_file::PyFileLikeObject;

use crate::{CalamineSheet, CalamineTable, Error, SheetMetadata, WorksheetNotFound};

enum SheetsEnum {
    File(Sheets<BufReader<File>>),
    FileLike(Sheets<Cursor<Vec<u8>>>),
    None,
}

enum WorkbookType {
    Xls,
    Xlsx,
    Xlsb,
    Ods,
}

impl From<&SheetsEnum> for WorkbookType {
    fn from(sheets: &SheetsEnum) -> Self {
        match sheets {
            SheetsEnum::File(f) => match f {
                Sheets::Xls(_) => WorkbookType::Xls,
                Sheets::Xlsx(_) => WorkbookType::Xlsx,
                Sheets::Xlsb(_) => WorkbookType::Xlsb,
                Sheets::Ods(_) => WorkbookType::Ods,
            },
            SheetsEnum::FileLike(f) => match f {
                Sheets::Xls(_) => WorkbookType::Xls,
                Sheets::Xlsx(_) => WorkbookType::Xlsx,
                Sheets::Xlsb(_) => WorkbookType::Xlsb,
                Sheets::Ods(_) => WorkbookType::Ods,
            },
            SheetsEnum::None => unreachable!(),
        }
    }
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

    fn worksheet_merge_cells(
        &mut self,
        name: &str,
    ) -> Result<Option<Vec<calamine::Dimensions>>, Error> {
        match self {
            SheetsEnum::File(f) => match f {
                Sheets::Xls(xls_f) => Ok(xls_f.worksheet_merge_cells(name)),
                Sheets::Xlsx(xlsx_f) => xlsx_f
                    .worksheet_merge_cells(name)
                    .transpose()
                    .map_err(CalamineCrateError::Xlsx)
                    .map_err(Error::Calamine)
                    .map(|inner| inner.or(Some(Vec::new()))),
                _ => Ok(None),
            },
            SheetsEnum::FileLike(f) => match f {
                Sheets::Xls(xls_f) => Ok(xls_f.worksheet_merge_cells(name)),
                Sheets::Xlsx(xlsx_f) => xlsx_f
                    .worksheet_merge_cells(name)
                    .transpose()
                    .map_err(CalamineCrateError::Xlsx)
                    .map_err(Error::Calamine)
                    .map(|inner| inner.or(Some(Vec::new()))),
                _ => Ok(None),
            },
            SheetsEnum::None => Err(Error::WorkbookClosed),
        }
    }

    fn load_tables(&mut self) -> Result<(), Error> {
        match self {
            SheetsEnum::File(f) => match f {
                Sheets::Xlsx(xlsx_f) => xlsx_f
                    .load_tables()
                    .map_err(CalamineCrateError::Xlsx)
                    .map_err(Error::Calamine),
                _ => Err(Error::TablesNotSupported),
            },
            SheetsEnum::FileLike(f) => match f {
                Sheets::Xlsx(xlsx_f) => xlsx_f
                    .load_tables()
                    .map_err(CalamineCrateError::Xlsx)
                    .map_err(Error::Calamine),
                _ => Err(Error::TablesNotSupported),
            },
            SheetsEnum::None => Err(Error::WorkbookClosed),
        }
    }

    fn table_names(&self) -> Result<Vec<String>, Error> {
        match self {
            SheetsEnum::File(f) => match f {
                Sheets::Xlsx(xlsx_f) => Ok(xlsx_f.table_names()),
                _ => Err(Error::TablesNotSupported),
            },
            SheetsEnum::FileLike(f) => match f {
                Sheets::Xlsx(xlsx_f) => Ok(xlsx_f.table_names()),
                _ => Err(Error::TablesNotSupported),
            },
            SheetsEnum::None => Err(Error::WorkbookClosed),
        }
        .map(|v| {
            v.iter()
                .map(|s| s.to_owned().to_owned())
                .collect::<Vec<String>>()
        })
    }

    fn get_table_by_name(&mut self, name: &str) -> Result<CalamineTable, Error> {
        match self {
            SheetsEnum::File(f) => match f {
                Sheets::Xlsx(xlsx_f) => xlsx_f
                    .table_by_name(name)
                    .map_err(CalamineCrateError::Xlsx)
                    .map_err(Error::Calamine)
                    .map(|t| {
                        CalamineTable::new(
                            t.name().to_owned(),
                            t.sheet_name().to_owned(),
                            t.columns().iter().map(|s| s.to_owned()).collect(),
                            t.data().to_owned(),
                        )
                    }),
                _ => Err(Error::TablesNotSupported),
            },
            SheetsEnum::FileLike(f) => match f {
                Sheets::Xlsx(xlsx_f) => xlsx_f
                    .table_by_name(name)
                    .map_err(CalamineCrateError::Xlsx)
                    .map_err(Error::Calamine)
                    .map(|t| {
                        CalamineTable::new(
                            t.name().to_owned(),
                            t.sheet_name().to_owned(),
                            t.columns().iter().map(|s| s.to_owned()).collect(),
                            t.data().to_owned(),
                        )
                    }),
                _ => Err(Error::TablesNotSupported),
            },
            SheetsEnum::None => Err(Error::WorkbookClosed),
        }
    }
}

#[pyclass]
pub struct CalamineWorkbook {
    #[pyo3(get)]
    path: Option<String>,
    workbook_type: WorkbookType,
    sheets: SheetsEnum,
    #[pyo3(get)]
    sheets_metadata: Vec<SheetMetadata>,
    #[pyo3(get)]
    sheet_names: Vec<String>,
    table_names: Option<Vec<String>>,
}

#[pymethods]
impl CalamineWorkbook {
    fn __repr__(&self) -> PyResult<String> {
        match &self.path {
            Some(path) => Ok(format!("CalamineWorkbook(path='{path}')")),
            None => Ok("CalamineWorkbook(path='bytes')".to_string()),
        }
    }

    #[classmethod]
    #[pyo3(name = "from_object", signature = (path_or_filelike, load_tables=false))]
    fn py_from_object(
        _cls: &Bound<'_, PyType>,
        py: Python<'_>,
        path_or_filelike: Py<PyAny>,
        load_tables: bool,
    ) -> PyResult<Self> {
        Self::from_object(py, path_or_filelike, load_tables)
    }

    #[classmethod]
    #[pyo3(name = "from_filelike", signature = (filelike, load_tables=false))]
    fn py_from_filelike(
        _cls: &Bound<'_, PyType>,
        py: Python<'_>,
        filelike: Py<PyAny>,
        load_tables: bool,
    ) -> PyResult<Self> {
        py.detach(|| Self::from_filelike(filelike, load_tables))
    }

    #[classmethod]
    #[pyo3(name = "from_path", signature = (path, load_tables=false))]
    fn py_from_path(
        _cls: &Bound<'_, PyType>,
        py: Python<'_>,
        path: Py<PyAny>,
        load_tables: bool,
    ) -> PyResult<Self> {
        if let Ok(string_ref) = path.extract::<PathBuf>(py) {
            let path = string_ref.to_string_lossy().to_string();
            return py.detach(|| Self::from_path(&path, load_tables));
        }

        Err(PyTypeError::new_err(""))
    }

    #[pyo3(name = "get_sheet_by_name")]
    fn py_get_sheet_by_name(&mut self, py: Python<'_>, name: &str) -> PyResult<CalamineSheet> {
        py.detach(|| self.get_sheet_by_name(name))
    }

    #[pyo3(name = "get_sheet_by_index")]
    fn py_get_sheet_by_index(&mut self, py: Python<'_>, index: usize) -> PyResult<CalamineSheet> {
        py.detach(|| self.get_sheet_by_index(index))
    }

    #[getter]
    fn table_names(&self) -> PyResult<Vec<String>> {
        match &self.workbook_type {
            WorkbookType::Xlsx => match &self.table_names {
                Some(v) => Ok(v.clone()),
                None => Err(Error::TablesNotLoaded.into()),
            },
            _ => Err(Error::TablesNotSupported.into()),
        }
    }

    #[pyo3(name = "get_table_by_name")]
    fn py_get_table_by_name(&mut self, py: Python<'_>, name: &str) -> PyResult<CalamineTable> {
        py.detach(|| self.get_table_by_name(name))
    }

    fn close(&mut self) -> PyResult<()> {
        match self.sheets {
            SheetsEnum::None => Err(Error::WorkbookClosed.into()),
            _ => {
                self.sheets = SheetsEnum::None;
                Ok(())
            }
        }
    }

    fn __enter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __exit__(
        &mut self,
        _exc_type: Py<PyAny>,
        _exc_value: Py<PyAny>,
        _traceback: Py<PyAny>,
    ) -> PyResult<()> {
        self.close()
    }
}

impl CalamineWorkbook {
    pub fn from_object(
        py: Python<'_>,
        path_or_filelike: Py<PyAny>,
        load_tables: bool,
    ) -> PyResult<Self> {
        if let Ok(string_ref) = path_or_filelike.extract::<PathBuf>(py) {
            let path = string_ref.to_string_lossy().to_string();
            return py.detach(|| Self::from_path(&path, load_tables));
        }

        py.detach(|| Self::from_filelike(path_or_filelike, load_tables))
    }

    pub fn from_filelike(filelike: Py<PyAny>, load_tables: bool) -> PyResult<Self> {
        let mut buf = vec![];
        PyFileLikeObject::with_requirements(filelike, true, false, true, false)?
            .read_to_end(&mut buf)?;
        let reader = Cursor::new(buf);
        let mut sheets =
            SheetsEnum::FileLike(open_workbook_auto_from_rs(reader).map_err(Error::Calamine)?);
        let sheet_names = sheets.sheet_names().to_owned();
        let sheets_metadata = sheets.sheets_metadata().to_owned();

        let mut table_names: Option<Vec<String>> = None;
        if load_tables {
            sheets.load_tables()?;
            table_names = Some(sheets.table_names()?);
        }

        Ok(Self {
            path: None,
            workbook_type: WorkbookType::from(&sheets),
            sheets,
            sheets_metadata,
            sheet_names,
            table_names,
        })
    }

    pub fn from_path(path: &str, load_tables: bool) -> PyResult<Self> {
        let mut sheets = SheetsEnum::File(open_workbook_auto(path).map_err(Error::Calamine)?);
        let sheet_names = sheets.sheet_names().to_owned();
        let sheets_metadata = sheets.sheets_metadata().to_owned();

        let mut table_names: Option<Vec<String>> = None;
        if load_tables {
            sheets.load_tables()?;
            table_names = Some(sheets.table_names()?);
        }
        Ok(Self {
            path: Some(path.to_string()),
            workbook_type: WorkbookType::from(&sheets),
            sheets,
            sheets_metadata,
            sheet_names,
            table_names,
        })
    }

    fn get_sheet_by_name(&mut self, name: &str) -> PyResult<CalamineSheet> {
        let range = self.sheets.worksheet_range(name)?;
        let merge_cells_range = self.sheets.worksheet_merge_cells(name)?;
        Ok(CalamineSheet::new(
            name.to_owned(),
            range,
            merge_cells_range,
        ))
    }

    fn get_sheet_by_index(&mut self, index: usize) -> PyResult<CalamineSheet> {
        let name = self
            .sheet_names
            .get(index)
            .ok_or_else(|| WorksheetNotFound::new_err(format!("Worksheet '{index}' not found")))?
            .to_string();
        self.get_sheet_by_name(&name)
    }

    fn get_table_by_name(&mut self, name: &str) -> PyResult<CalamineTable> {
        match &self.workbook_type {
            WorkbookType::Xlsx => match &self.table_names {
                Some(_) => Ok(self.sheets.get_table_by_name(name)?),
                None => Err(Error::TablesNotLoaded.into()),
            },
            _ => Err(Error::TablesNotSupported.into()),
        }
    }
}
