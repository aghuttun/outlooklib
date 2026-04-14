use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};
use serde_json::Value;

/// Represent an Outlook folder with its identifier and display name.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ListFolders {
    /// Unique Graph identifier for the folder.
    pub id: String,
    /// Human readable folder name returned by Graph.
    #[serde(rename = "displayName")]
    pub display_name: String,
}

/// Represent an Outlook message with sender, subject, status, and metadata.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ListMessages {
    /// Unique Graph identifier for the message.
    pub id: String,
    /// Sender payload as returned by Graph.
    pub sender: Value,
    /// Message receive timestamp in UTC.
    #[serde(rename = "receivedDateTime")]
    pub received_date_time: DateTime<Utc>,
    /// Message subject line.
    pub subject: String,
    /// Read status of the message.
    #[serde(rename = "isRead")]
    pub is_read: bool,
    /// Indicates if the message has any attachment.
    #[serde(rename = "hasAttachments")]
    pub has_attachments: bool,
    /// Message importance label (for example low/normal/high).
    pub importance: String,
    /// Flag payload as returned by Graph.
    pub flag: Value,
    /// Browser URL for opening the message in Outlook Web.
    #[serde(rename = "webLink")]
    pub web_link: String,
}

/// Represent metadata returned by Graph after moving a message.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MoveMessage {
    /// Identifier of the moved message in its destination folder.
    pub id: String,
    /// Message version key after move operation.
    #[serde(rename = "changeKey")]
    pub change_key: String,
}

#[derive(Debug, Deserialize)]
/// Generic Graph list wrapper (`{ "value": [...] }`).
pub(crate) struct GraphListResponse<T> {
    /// List payload returned by Graph endpoints.
    pub value: Vec<T>,
}

#[derive(Debug, Deserialize)]
/// Internal representation for message attachment metadata.
pub(crate) struct AttachmentItem {
    /// Attachment file name.
    pub name: String,
    /// Base64-encoded file content when available.
    #[serde(rename = "contentBytes")]
    pub content_bytes: Option<String>,
}

#[derive(Debug, Serialize)]
/// Outbound body wrapper for Graph `sendMail` endpoint.
pub(crate) struct SendMailBody<'a> {
    /// Email message payload.
    pub message: SendMessage<'a>,
    /// Graph option for saving email in sent items.
    #[serde(rename = "saveToSentItems")]
    pub save_to_sent_items: &'a str,
}

#[derive(Debug, Serialize)]
/// Internal Graph message payload used by `send_message`.
pub(crate) struct SendMessage<'a> {
    /// Subject line.
    pub subject: &'a str,
    /// Body metadata and content.
    pub body: SendBody<'a>,
    /// Recipient list.
    #[serde(rename = "toRecipients")]
    pub to_recipients: Vec<SendRecipient<'a>>,
    /// Optional list of file attachments.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub attachments: Option<Vec<SendAttachment>>,
}

#[derive(Debug, Serialize)]
/// Internal representation of Graph message body.
pub(crate) struct SendBody<'a> {
    /// Graph content type (for example `HTML`).
    #[serde(rename = "contentType")]
    pub content_type: &'a str,
    /// Body content.
    pub content: &'a str,
}

#[derive(Debug, Serialize)]
/// Internal wrapper for a recipient email address.
pub(crate) struct SendRecipient<'a> {
    /// Recipient address object.
    #[serde(rename = "emailAddress")]
    pub email_address: SendEmailAddress<'a>,
}

#[derive(Debug, Serialize)]
/// Internal email address payload used in recipient lists.
pub(crate) struct SendEmailAddress<'a> {
    /// Recipient email address.
    pub address: &'a str,
}

#[derive(Debug, Serialize)]
/// Internal file attachment payload for Graph `sendMail`.
pub(crate) struct SendAttachment {
    /// Graph OData attachment type descriptor.
    #[serde(rename = "@odata.type")]
    pub odata_type: &'static str,
    /// Attachment file name.
    pub name: String,
    /// Base64-encoded content bytes.
    #[serde(rename = "contentBytes")]
    pub content_bytes: String,
}