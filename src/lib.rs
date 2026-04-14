mod models;
#[cfg(feature = "python")]
mod python_bindings;

use std::fs;
use std::path::{Path, PathBuf};

use base64::{Engine as _, engine::general_purpose};
pub use models::{ListFolders, ListMessages, MoveMessage};
use models::{
    AttachmentItem, GraphListResponse, SendAttachment, SendBody, SendEmailAddress, SendMailBody,
    SendMessage, SendRecipient,
};
use reqwest::StatusCode;
use reqwest::blocking::{Client, Response as HttpResponse};
use serde::Deserialize;
use serde::Serialize;
use thiserror::Error;

/// Store runtime configuration for the Outlook Graph client.
///
/// This struct holds Microsoft Graph endpoint metadata, OAuth credentials,
/// the current bearer token, and the selected mailbox context.
#[derive(Debug, Clone)]
pub struct Configuration {
    /// Graph API host name, usually `graph.microsoft.com`.
    pub api_domain: String,
    /// Graph API version segment, usually `v1.0`.
    pub api_version: String,
    /// Azure application client identifier.
    pub client_id: String,
    /// Azure tenant identifier.
    pub tenant_id: String,
    /// Azure application client secret.
    pub client_secret: String,
    /// Current OAuth bearer token returned by Microsoft Identity Platform.
    pub token: Option<String>,
    /// Mailbox owner address used in Graph `/users/{email}` routes.
    pub client_email: String,
    /// Currently selected folder name or REST id (for example `Inbox` or `root`).
    pub client_folder: String,
}

/// Generic response container used by public Outlook operations.
///
/// The payload shape depends on the called method. For non-successful API
/// statuses, `content` is typically `None`.
#[derive(Debug, Clone)]
pub struct Response<T> {
    /// HTTP status code returned by Microsoft Graph.
    pub status_code: u16,
    /// Optional operation payload for successful requests.
    pub content: Option<T>,
}

/// Client for interacting with Microsoft Outlook through Microsoft Graph API.
///
/// This implementation mirrors the previous Python library surface while using
/// strongly typed Rust models and error handling.
#[derive(Debug)]
pub struct Outlook {
    configuration: Configuration,
    client: Client,
}

/// Error type produced by Outlook operations.
#[derive(Debug, Error)]
pub enum OutlookError {
    /// HTTP transport or protocol error.
    #[error("HTTP error: {0}")]
    Http(#[from] reqwest::Error),
    /// Local file I/O error.
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
    /// JSON serialization/deserialization error.
    #[error("JSON error: {0}")]
    Json(#[from] serde_json::Error),
    /// Base64 decode error while handling attachments.
    #[error("Base64 error: {0}")]
    Base64(#[from] base64::DecodeError),
}

#[derive(Debug, Deserialize)]
/// Internal OAuth token response payload from Microsoft Identity Platform.
struct AuthTokenResponse {
    /// Bearer token used in `Authorization: Bearer ...` headers.
    access_token: String,
}

impl Outlook {
    /// Build a new client using `Inbox` as the default folder.
    ///
    /// This helper exists to mirror the Python API default folder behaviour.
    ///
    /// # Arguments
    /// - `client_id`: Azure application client identifier.
    /// - `tenant_id`: Azure tenant identifier.
    /// - `client_secret`: Azure application client secret.
    /// - `client_email`: Mailbox owner email used in Graph `/users/{email}` routes.
    ///
    /// # Returns
    /// A fully initialised and authenticated [`Outlook`] client.
    ///
    /// # Errors
    /// Returns [`OutlookError`] when client creation or authentication fails.
    pub fn new_default_folder(
        client_id: impl Into<String>,
        tenant_id: impl Into<String>,
        client_secret: impl Into<String>,
        client_email: impl Into<String>,
    ) -> Result<Self, OutlookError> {
        Self::new(
            client_id,
            tenant_id,
            client_secret,
            client_email,
            "Inbox",
        )
    }

    /// Build and authenticate a new Outlook Graph client.
    ///
    /// The constructor stores credentials and user context, applies the
    /// selected folder, and requests an OAuth token using client credentials.
    ///
    /// # Arguments
    /// - `client_id`: Azure application client identifier.
    /// - `tenant_id`: Azure tenant identifier.
    /// - `client_secret`: Azure application client secret.
    /// - `client_email`: Mailbox owner email used in Graph `/users/{email}` routes.
    /// - `client_folder`: Initial folder name or REST id.
    ///
    /// # Returns
    /// A fully initialised and authenticated [`Outlook`] client.
    ///
    /// # Errors
    /// Returns [`OutlookError`] when HTTP client setup or authentication fails.
    pub fn new(
        client_id: impl Into<String>,
        tenant_id: impl Into<String>,
        client_secret: impl Into<String>,
        client_email: impl Into<String>,
        client_folder: impl Into<String>,
    ) -> Result<Self, OutlookError> {
        let client = Client::builder().build()?;
        let mut this = Self {
            configuration: Configuration {
                api_domain: "graph.microsoft.com".to_string(),
                api_version: "v1.0".to_string(),
                client_id: client_id.into(),
                tenant_id: tenant_id.into(),
                client_secret: client_secret.into(),
                token: None,
                client_email: client_email.into(),
                client_folder: client_folder.into(),
            },
            client,
        };

        this.change_folder(this.configuration.client_folder.clone());
        this._authenticate()?;
        Ok(this)
    }

    /// Authenticate with Microsoft Identity Platform using client credentials.
    ///
    /// On HTTP 200, updates the in-memory bearer token used by subsequent
    /// Graph API requests.
    ///
    /// # Errors
    /// Returns [`OutlookError`] for HTTP failures or JSON deserialization errors.
    fn _authenticate(&mut self) -> Result<(), OutlookError> {
        let url_auth = format!(
            "https://login.microsoftonline.com/{}/oauth2/v2.0/token",
            self.configuration.tenant_id
        );

        let response = self
            .client
            .post(url_auth)
            .header("Content-Type", "application/x-www-form-urlencoded")
            .form(&[
                ("grant_type", "client_credentials"),
                ("client_id", self.configuration.client_id.as_str()),
                ("client_secret", self.configuration.client_secret.as_str()),
                ("scope", "https://graph.microsoft.com/.default"),
            ])
            .send()?;

        if response.status() == StatusCode::OK {
            let token: AuthTokenResponse = response.json()?;
            self.configuration.token = Some(token.access_token);
        }

        Ok(())
    }

    /// Force token renewal by re-running the authentication flow.
    ///
    /// # Returns
    /// `Ok(())` when the renewal attempt completes.
    ///
    /// # Errors
    /// Returns [`OutlookError`] when authentication fails.
    pub fn renew_token(&mut self) -> Result<(), OutlookError> {
        self._authenticate()
    }

    /// Change the target mailbox email for subsequent operations.
    ///
    /// # Arguments
    /// - `email`: Mailbox owner email used in Graph routes.
    pub fn change_client_email(&mut self, email: impl Into<String>) {
        self.configuration.client_email = email.into();
    }

    /// Change the current folder used by message and folder operations.
    ///
    /// Typical values include `Inbox`, `SentItems`, `Drafts`, `root`, or a
    /// folder REST id.
    ///
    /// # Arguments
    /// - `id`: Folder name or REST id.
    pub fn change_folder(&mut self, id: impl Into<String>) {
        self.configuration.client_folder = id.into();
    }

    /// Retrieve mail folders from the current mailbox context.
    ///
    /// When the selected folder is `root`, this lists root folders; otherwise
    /// it lists child folders under the currently selected folder.
    ///
    /// If `save_as` is provided, raw JSON response bytes are also written to
    /// disk.
    ///
    /// # Arguments
    /// - `save_as`: Optional output path for saving raw Graph JSON bytes.
    ///
    /// # Returns
    /// A [`Response`] containing status code and optional list of [`ListFolders`].
    ///
    /// # Errors
    /// Returns [`OutlookError`] for HTTP, I/O, or JSON processing failures.
    pub fn list_folders(
        &self,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response<Vec<ListFolders>>, OutlookError> {
        let url_query = if self.configuration.client_folder.eq_ignore_ascii_case("root") {
            format!(
                "https://{}/{}/users/{}/mailFolders/",
                self.configuration.api_domain,
                self.configuration.api_version,
                self.configuration.client_email
            )
        } else {
            format!(
                "https://{}/{}/users/{}/mailFolders/{}/childFolders",
                self.configuration.api_domain,
                self.configuration.api_version,
                self.configuration.client_email,
                self.configuration.client_folder
            )
        };

        let response = self
            .client
            .get(url_query)
            .bearer_auth(self.configuration.token.as_deref().unwrap_or_default())
            .query(&[
                ("$select", "id,displayName"),
                ("$top", "100"),
                ("includeHiddenFolders", "true"),
            ])
            .send()?;

        self._handle_list_response::<ListFolders>(response, save_as)
    }

    /// Retrieve up to 100 messages from the selected folder using a Graph filter.
    ///
    /// The `filter` string must follow Microsoft Graph OData filter syntax.
    /// If `save_as` is provided, raw JSON response bytes are written to disk.
    ///
    /// # Arguments
    /// - `filter`: OData filter expression used by Microsoft Graph.
    /// - `save_as`: Optional output path for saving raw Graph JSON bytes.
    ///
    /// # Returns
    /// A [`Response`] containing status code and optional list of [`ListMessages`].
    ///
    /// # Errors
    /// Returns [`OutlookError`] for HTTP, I/O, or JSON processing failures.
    pub fn list_messages(
        &self,
        filter: &str,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response<Vec<ListMessages>>, OutlookError> {
        let url_query = format!(
            "https://{}/{}/users/{}/mailFolders/{}/messages",
            self.configuration.api_domain,
            self.configuration.api_version,
            self.configuration.client_email,
            self.configuration.client_folder
        );

        let response = self
            .client
            .get(url_query)
            .bearer_auth(self.configuration.token.as_deref().unwrap_or_default())
            .query(&[
                (
                    "$select",
                    "id,sender,receivedDateTime,subject,isRead,hasAttachments,importance,flag,webLink",
                ),
                ("$filter", filter),
                ("$top", "100"),
            ])
            .send()?;

        self._handle_list_response::<ListMessages>(response, save_as)
    }

    /// Move a message from the current folder to another folder.
    ///
    /// Returns Graph move metadata (`id`, `changeKey`) on HTTP 201.
    /// If `save_as` is provided, raw JSON response bytes are written to disk.
    ///
    /// # Arguments
    /// - `id`: Message id to move.
    /// - `to`: Destination folder REST id.
    /// - `save_as`: Optional output path for saving raw Graph JSON bytes.
    ///
    /// # Returns
    /// A [`Response`] containing status code and optional [`MoveMessage`] payload.
    ///
    /// # Errors
    /// Returns [`OutlookError`] for HTTP, I/O, or JSON processing failures.
    pub fn move_message(
        &self,
        id: &str,
        to: &str,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response<MoveMessage>, OutlookError> {
        let url_query = format!(
            "https://{}/{}/users/{}/mailFolders/{}/messages/{}/move",
            self.configuration.api_domain,
            self.configuration.api_version,
            self.configuration.client_email,
            self.configuration.client_folder,
            id
        );

        let response = self
            .client
            .post(url_query)
            .bearer_auth(self.configuration.token.as_deref().unwrap_or_default())
            .query(&[("$select", "id,changeKey")])
            .json(&serde_json::json!({ "DestinationId": to }))
            .send()?;

        self._handle_scalar_response::<MoveMessage>(response, StatusCode::CREATED, save_as)
    }

    /// Delete a message from the current folder.
    ///
    /// Returns status 204 on successful deletion.
    ///
    /// # Arguments
    /// - `id`: Message id to delete.
    ///
    /// # Returns
    /// A [`Response`] containing the HTTP status code.
    ///
    /// # Errors
    /// Returns [`OutlookError`] for HTTP transport failures.
    pub fn delete_message(&self, id: &str) -> Result<Response<()>, OutlookError> {
        let url_query = format!(
            "https://{}/{}/users/{}/mailFolders/{}/messages/{}",
            self.configuration.api_domain,
            self.configuration.api_version,
            self.configuration.client_email,
            self.configuration.client_folder,
            id
        );

        let response = self
            .client
            .delete(url_query)
            .bearer_auth(self.configuration.token.as_deref().unwrap_or_default())
            .send()?;

        let status = response.status().as_u16();
        let content = if response.status() == StatusCode::NO_CONTENT {
            Some(())
        } else {
            None
        };

        Ok(Response {
            status_code: status,
            content,
        })
    }

    /// Download all file attachments from a given message id.
    ///
    /// Attachment names are sanitized to keep only ASCII alphanumeric
    /// characters and dots. When `index` is true, a numeric suffix is appended
    /// to each file name to reduce overwrite collisions.
    ///
    /// # Arguments
    /// - `id`: Message id containing attachments.
    /// - `path`: Target directory where files are written.
    /// - `index`: Whether to append numeric suffixes to file names.
    ///
    /// # Returns
    /// A [`Response`] containing the HTTP status code.
    ///
    /// # Errors
    /// Returns [`OutlookError`] for HTTP, filesystem, base64, or JSON failures.
    pub fn download_message_attachment(
        &self,
        id: &str,
        path: impl AsRef<Path>,
        index: bool,
    ) -> Result<Response<()>, OutlookError> {
        let url_query = format!(
            "https://{}/{}/users/{}/mailFolders/{}/messages/{}/attachments",
            self.configuration.api_domain,
            self.configuration.api_version,
            self.configuration.client_email,
            self.configuration.client_folder,
            id
        );

        let response = self
            .client
            .get(url_query)
            .bearer_auth(self.configuration.token.as_deref().unwrap_or_default())
            .send()?;

        let status = response.status().as_u16();
        if response.status() != StatusCode::OK {
            return Ok(Response {
                status_code: status,
                content: None,
            });
        }

        let payload: GraphListResponse<AttachmentItem> = response.json()?;
        let mut counter: usize = 0;

        for row in payload.value {
            if let Some(content_b64) = row.content_bytes {
                let file_content = general_purpose::STANDARD.decode(content_b64)?;
                let filename = if index {
                    counter += 1;
                    append_index_to_filename(&row.name, counter)
                } else {
                    row.name
                };

                let sanitized = sanitize_filename(&filename);
                let full_path = PathBuf::from(path.as_ref()).join(sanitized);
                fs::write(full_path, file_content)?;
            }
        }

        Ok(Response {
            status_code: status,
            content: Some(()),
        })
    }

    /// Send an HTML email and optionally include file attachments.
    ///
    /// `recipients` maps to Graph `toRecipients`. Optional file paths in
    /// `attachments` are read and sent as base64-encoded file attachments.
    ///
    /// # Arguments
    /// - `recipients`: Destination email addresses.
    /// - `subject`: Subject line.
    /// - `message`: HTML body content.
    /// - `attachments`: Optional list of file paths to include.
    ///
    /// # Returns
    /// A [`Response`] containing the HTTP status code.
    ///
    /// # Errors
    /// Returns [`OutlookError`] for HTTP or filesystem failures.
    pub fn send_message(
        &self,
        recipients: &[impl AsRef<str>],
        subject: &str,
        message: &str,
        attachments: Option<&[impl AsRef<str>]>,
    ) -> Result<Response<()>, OutlookError> {
        let url = format!(
            "https://{}/{}/users/{}/sendMail",
            self.configuration.api_domain, self.configuration.api_version, self.configuration.client_email
        );

        let to_recipients = recipients
            .iter()
            .map(|r| SendRecipient {
                email_address: SendEmailAddress {
                    address: r.as_ref(),
                },
            })
            .collect::<Vec<_>>();

        let encoded_attachments = if let Some(paths) = attachments {
            let mut out = Vec::with_capacity(paths.len());
            for file_path in paths {
                let file_path = file_path.as_ref();
                let file_content = fs::read(file_path)?;
                let encoded = general_purpose::STANDARD.encode(file_content);
                let name = Path::new(file_path)
                    .file_name()
                    .and_then(|s| s.to_str())
                    .unwrap_or("attachment.bin")
                    .to_string();

                out.push(SendAttachment {
                    odata_type: "#microsoft.graph.fileAttachment",
                    name,
                    content_bytes: encoded,
                });
            }
            Some(out)
        } else {
            None
        };

        let payload = SendMailBody {
            message: SendMessage {
                subject,
                body: SendBody {
                    content_type: "HTML",
                    content: message,
                },
                to_recipients,
                attachments: encoded_attachments,
            },
            save_to_sent_items: "true",
        };

        let response = self
            .client
            .post(url)
            .bearer_auth(self.configuration.token.as_deref().unwrap_or_default())
            .json(&payload)
            .send()?;

        Ok(Response {
            status_code: response.status().as_u16(),
            content: None,
        })
    }

    /// Persist raw Graph response bytes to disk when an output path is provided.
    fn _export_to_json(
        &self,
        body: &[u8],
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<(), OutlookError> {
        if let Some(path) = save_as {
            fs::write(path.as_ref(), body)?;
        }
        Ok(())
    }

    /// Handle Graph responses shaped as lists (`{ "value": [...] }`).
    ///
    /// Returns `content=None` when status is not 200, preserving the HTTP code.
    fn _handle_list_response<T>(
        &self,
        response: HttpResponse,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response<Vec<T>>, OutlookError>
    where
        T: for<'de> Deserialize<'de>,
    {
        let status = response.status().as_u16();
        if status != 200 {
            return Ok(Response {
                status_code: status,
                content: None,
            });
        }

        let body = response.bytes()?;
        self._export_to_json(&body, save_as)?;
        let parsed: GraphListResponse<T> = serde_json::from_slice(&body)?;

        Ok(Response {
            status_code: status,
            content: Some(parsed.value),
        })
    }

    /// Handle Graph responses containing a single JSON object.
    ///
    /// Returns `content=None` when status differs from `success_status`.
    fn _handle_scalar_response<T>(
        &self,
        response: HttpResponse,
        success_status: StatusCode,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response<T>, OutlookError>
    where
        T: for<'de> Deserialize<'de>,
    {
        let status = response.status().as_u16();
        if response.status() != success_status {
            return Ok(Response {
                status_code: status,
                content: None,
            });
        }

        let body = response.bytes()?;
        self._export_to_json(&body, save_as)?;
        let parsed: T = serde_json::from_slice(&body)?;

        Ok(Response {
            status_code: status,
            content: Some(parsed),
        })
    }
}

/// Append a numeric suffix to a file name before its extension.
fn append_index_to_filename(filename: &str, index: usize) -> String {
    let path = Path::new(filename);
    let stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or("file");
    let ext = path.extension().and_then(|s| s.to_str());

    match ext {
        Some(ext) if !ext.is_empty() => format!("{}_{}.{}", stem, index, ext),
        _ => format!("{}_{}", stem, index),
    }
}

/// Replace unsupported file name characters with underscores.
///
/// Allowed characters are ASCII alphanumeric characters and dots.
fn sanitize_filename(input: &str) -> String {
    input
        .chars()
        .map(|ch| {
            if ch.is_ascii_alphanumeric() || ch == '.' {
                ch
            } else {
                '_'
            }
        })
        .collect()
}

/// Serialize any value to JSON for debugging output.
fn _serialize_for_debug<T: Serialize>(value: &T) -> Result<String, OutlookError> {
    Ok(serde_json::to_string(value)?)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn filename_index_is_inserted_before_extension() {
        let out = append_index_to_filename("report.pdf", 2);
        assert_eq!(out, "report_2.pdf");
    }

    #[test]
    fn filename_index_handles_no_extension() {
        let out = append_index_to_filename("report", 3);
        assert_eq!(out, "report_3");
    }

    #[test]
    fn sanitize_filename_replaces_invalid_characters() {
        let out = sanitize_filename("my file(1).pdf");
        assert_eq!(out, "my_file_1_.pdf");
    }

    #[test]
    fn list_folders_alias_deserialization_works() {
        let raw = r#"{"id":"abc","displayName":"Inbox"}"#;
        let folder: ListFolders = serde_json::from_str(raw).expect("folder json should parse");
        assert_eq!(folder.id, "abc");
        assert_eq!(folder.display_name, "Inbox");
    }

    #[test]
    fn move_message_alias_deserialization_works() {
        let raw = r#"{"id":"msg-1","changeKey":"ck-2"}"#;
        let moved: MoveMessage = serde_json::from_str(raw).expect("move json should parse");
        assert_eq!(moved.id, "msg-1");
        assert_eq!(moved.change_key, "ck-2");
    }
}