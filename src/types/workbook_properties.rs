use std::collections::BTreeMap;

use pyo3::prelude::*;

/// Document properties of a workbook, such as the creator, title and company.
#[pyclass(eq, from_py_object)]
#[derive(Clone, Debug, Default, PartialEq)]
pub struct WorkbookProperties {
    #[pyo3(get)]
    creator: Option<String>,
    #[pyo3(get)]
    last_modified_by: Option<String>,
    #[pyo3(get)]
    created: Option<String>,
    #[pyo3(get)]
    modified: Option<String>,
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
            ("creator", &self.creator),
            ("last_modified_by", &self.last_modified_by),
            ("created", &self.created),
            ("modified", &self.modified),
            ("title", &self.title),
            ("subject", &self.subject),
            ("description", &self.description),
            ("keywords", &self.keywords),
            ("category", &self.category),
            ("content_status", &self.content_status),
            ("revision", &self.revision),
            ("version", &self.version),
            ("application", &self.application),
            ("app_version", &self.app_version),
            ("company", &self.company),
            ("template", &self.template),
            ("manager", &self.manager),
        ];
        let mut parts = Vec::new();
        for (name, value) in fields {
            if let Some(value) = value {
                parts.push(format!("{name}={}", value.into_pyobject(py)?.repr()?));
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
            created: props.created,
            modified: props.modified,
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
