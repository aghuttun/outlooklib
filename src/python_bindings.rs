use std::path::PathBuf;

use pyo3::prelude::*;
use pyo3::types::PyModule;

use crate::{ListFolders, ListMessages, MoveMessage, Outlook};

/// Python-facing response object mirroring the legacy `Response` shape.
#[pyclass(name = "Response")]
pub struct PyResponse {
    /// HTTP status code returned by Graph.
    #[pyo3(get)]
    pub status_code: u16,
    /// JSON payload encoded as string and lazily decoded when requested.
    content_json: Option<String>,
}

#[pymethods]
impl PyResponse {
    /// Return deserialized Python content (`dict`, `list`, or `None`).
    #[getter]
    fn content(&self, py: Python<'_>) -> PyResult<PyObject> {
        match &self.content_json {
            Some(raw_json) => {
                let json_module = PyModule::import(py, "json")?;
                let value = json_module.call_method1("loads", (raw_json,))?;
                Ok(value.into())
            }
            None => Ok(py.None()),
        }
    }
}

/// Python-facing Outlook client backed by the Rust core implementation.
#[pyclass(name = "Outlook")]
pub struct PyOutlook {
    /// Internal Rust Outlook client.
    inner: Outlook,
}

#[pymethods]
impl PyOutlook {
    /// Initialise a new Outlook client using Graph credentials.
    ///
    /// The `custom_logger` argument is accepted for API compatibility with the
    /// original Python implementation and is currently ignored.
    #[new]
    #[pyo3(signature = (client_id, tenant_id, client_secret, client_email, client_folder="Inbox", custom_logger=None))]
    fn new(
        client_id: String,
        tenant_id: String,
        client_secret: String,
        client_email: String,
        client_folder: &str,
        custom_logger: Option<PyObject>,
    ) -> PyResult<Self> {
        let _ = custom_logger;
        let inner = Outlook::new(
            client_id,
            tenant_id,
            client_secret,
            client_email,
            client_folder,
        )
        .map_err(into_pyerr)?;

        Ok(Self { inner })
    }

    /// Renew the current bearer token.
    fn renew_token(&mut self) -> PyResult<()> {
        self.inner.renew_token().map_err(into_pyerr)
    }

    /// Change the mailbox email for subsequent operations.
    fn change_client_email(&mut self, email: String) {
        self.inner.change_client_email(email);
    }

    /// Change the current folder used by Graph queries.
    fn change_folder(&mut self, id: String) {
        self.inner.change_folder(id);
    }

    /// List folders for the current mailbox context.
    #[pyo3(signature = (save_as=None))]
    fn list_folders(&self, save_as: Option<String>) -> PyResult<PyResponse> {
        let response = self
            .inner
            .list_folders(save_as.as_ref().map(PathBuf::from))
            .map_err(into_pyerr)?;

        py_response_from(response.status_code, response.content)
    }

    /// List messages from the selected folder using a Graph filter.
    #[pyo3(signature = (filter, save_as=None))]
    fn list_messages(&self, filter: String, save_as: Option<String>) -> PyResult<PyResponse> {
        let response = self
            .inner
            .list_messages(&filter, save_as.as_ref().map(PathBuf::from))
            .map_err(into_pyerr)?;

        py_response_from(response.status_code, response.content)
    }

    /// Move a message from the current folder to destination folder `to`.
    #[pyo3(signature = (id, to, save_as=None))]
    fn move_message(&self, id: String, to: String, save_as: Option<String>) -> PyResult<PyResponse> {
        let response = self
            .inner
            .move_message(&id, &to, save_as.as_ref().map(PathBuf::from))
            .map_err(into_pyerr)?;

        py_response_from(response.status_code, response.content)
    }

    /// Delete a message from the current folder.
    fn delete_message(&self, id: String) -> PyResult<PyResponse> {
        let response = self.inner.delete_message(&id).map_err(into_pyerr)?;
        py_response_from(response.status_code, response.content)
    }

    /// Download all attachments from a message into `path`.
    #[pyo3(signature = (id, path, index=false))]
    fn download_message_attachment(&self, id: String, path: String, index: bool) -> PyResult<PyResponse> {
        let response = self
            .inner
            .download_message_attachment(&id, PathBuf::from(path), index)
            .map_err(into_pyerr)?;

        py_response_from(response.status_code, response.content)
    }

    /// Send an HTML email with optional file attachments.
    #[pyo3(signature = (recipients, subject, message, attachments=None))]
    fn send_message(
        &self,
        recipients: Vec<String>,
        subject: String,
        message: String,
        attachments: Option<Vec<String>>,
    ) -> PyResult<PyResponse> {
        let response = self
            .inner
            .send_message(
                &recipients,
                &subject,
                &message,
                attachments.as_deref(),
            )
            .map_err(into_pyerr)?;

        py_response_from(response.status_code, response.content)
    }
}

/// Python module initialization hook.
#[pymodule]
fn outlooklib(_py: Python<'_>, module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<PyOutlook>()?;
    module.add_class::<PyResponse>()?;
    Ok(())
}

/// Convert internal Rust response payloads into Python `Response` objects.
fn py_response_from<T>(status_code: u16, content: Option<T>) -> PyResult<PyResponse>
where
    T: serde::Serialize,
{
    let content_json = match content {
        Some(value) => Some(serde_json::to_string(&value).map_err(into_pyerr)?),
        None => None,
    };

    Ok(PyResponse {
        status_code,
        content_json,
    })
}

/// Convert Rust errors into Python runtime errors.
fn into_pyerr<E>(error: E) -> PyErr
where
    E: std::fmt::Display,
{
    pyo3::exceptions::PyRuntimeError::new_err(error.to_string())
}

#[allow(dead_code)]
/// Keep model types referenced in this module to help maintenance.
fn _type_markers(_a: Option<ListFolders>, _b: Option<ListMessages>, _c: Option<MoveMessage>) {}
