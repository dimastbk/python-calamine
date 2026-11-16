use std::collections::BTreeMap;

use chrono::{DateTime, Datelike, FixedOffset, NaiveDateTime};
use pyo3::prelude::*;

/// A document property date, with a UTC offset if the file stores one.
#[derive(Clone, Debug, PartialEq, IntoPyObject)]
enum PropertyDate {
    Aware(DateTime<FixedOffset>),
    Naive(NaiveDateTime),
}

impl PropertyDate {
    /// Parse a date such as "2024-01-15T08:30:00Z" (xlsx, xlsb, xls) or
    /// "2024-06-20T14:22:00.123456789" (ods). Other formats, and years that
    /// Python's `datetime` can't represent, return `None`.
    fn parse(value: &str) -> Option<Self> {
        let value = value.trim();
        let date = match DateTime::parse_from_rfc3339(value) {
            Ok(date) => PropertyDate::Aware(date),
            Err(_) => NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f")
                .ok()
                .map(PropertyDate::Naive)?,
        };
        let year = match &date {
            PropertyDate::Aware(date) => date.year(),
            PropertyDate::Naive(date) => date.year(),
        };
        (1..=9999).contains(&year).then_some(date)
    }
}

/// Document properties of a workbook, such as the creator, title and company.
#[pyclass(eq, from_py_object)]
#[derive(Clone, Debug, Default, PartialEq)]
pub struct WorkbookProperties {
    #[pyo3(get)]
    creator: Option<String>,
    #[pyo3(get)]
    last_modified_by: Option<String>,
    #[pyo3(get)]
    created: Option<PropertyDate>,
    #[pyo3(get)]
    modified: Option<PropertyDate>,
    #[pyo3(get)]
    title: Option<String>,
    #[pyo3(get)]
    subject: Option<String>,
    #[pyo3(get)]
    description: Option<String>,
    #[pyo3(get)]
    keywords: Option<String>,
    #[pyo3(get)]
    category: Option<String>,
    #[pyo3(get)]
    content_status: Option<String>,
    #[pyo3(get)]
    revision: Option<String>,
    #[pyo3(get)]
    version: Option<String>,
    #[pyo3(get)]
    application: Option<String>,
    #[pyo3(get)]
    app_version: Option<String>,
    #[pyo3(get)]
    company: Option<String>,
    #[pyo3(get)]
    template: Option<String>,
    #[pyo3(get)]
    manager: Option<String>,
    #[pyo3(get)]
    custom_properties: BTreeMap<String, String>,
}

#[pymethods]
impl WorkbookProperties {
    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        let fields = [
            ("creator", self.creator.clone().into_pyobject(py)?),
            (
                "last_modified_by",
                self.last_modified_by.clone().into_pyobject(py)?,
            ),
            ("created", self.created.clone().into_pyobject(py)?),
            ("modified", self.modified.clone().into_pyobject(py)?),
            ("title", self.title.clone().into_pyobject(py)?),
            ("subject", self.subject.clone().into_pyobject(py)?),
            ("description", self.description.clone().into_pyobject(py)?),
            ("keywords", self.keywords.clone().into_pyobject(py)?),
            ("category", self.category.clone().into_pyobject(py)?),
            (
                "content_status",
                self.content_status.clone().into_pyobject(py)?,
            ),
            ("revision", self.revision.clone().into_pyobject(py)?),
            ("version", self.version.clone().into_pyobject(py)?),
            ("application", self.application.clone().into_pyobject(py)?),
            ("app_version", self.app_version.clone().into_pyobject(py)?),
            ("company", self.company.clone().into_pyobject(py)?),
            ("template", self.template.clone().into_pyobject(py)?),
            ("manager", self.manager.clone().into_pyobject(py)?),
        ];
        let mut parts = Vec::new();
        for (name, value) in fields {
            if !value.is_none() {
                parts.push(format!("{name}={}", value.repr()?));
            }
        }
        if !self.custom_properties.is_empty() {
            let custom = (&self.custom_properties).into_pyobject(py)?;
            parts.push(format!("custom_properties={}", custom.repr()?));
        }
        Ok(format!("WorkbookProperties({})", parts.join(", ")))
    }
}

impl From<calamine::WorkbookProperties> for WorkbookProperties {
    fn from(props: calamine::WorkbookProperties) -> Self {
        WorkbookProperties {
            creator: props.creator,
            last_modified_by: props.last_modified_by,
            created: props.created.as_deref().and_then(PropertyDate::parse),
            modified: props.modified.as_deref().and_then(PropertyDate::parse),
            title: props.title,
            subject: props.subject,
            description: props.description,
            keywords: props.keywords,
            category: props.category,
            content_status: props.content_status,
            revision: props.revision,
            version: props.version,
            application: props.application,
            app_version: props.app_version,
            company: props.company,
            template: props.template,
            manager: props.manager,
            custom_properties: props.custom_properties.into_iter().collect(),
        }
    }
}
