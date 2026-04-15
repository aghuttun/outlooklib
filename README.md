# outlooklib

Microsoft Outlook client package for Python using Microsoft Graph.

- [Installation](#installation)
- [Usage](#usage)
- [Version](#version)
- [Licence](#licence)

## Installation

```bash
pip install outlooklib
```

## Usage

```python
import outlooklib

outlook = outlooklib.Outlook(
    client_id="<client-id>",
    tenant_id="<tenant-id>",
    client_secret="<client-secret>",
    client_email="team.mailbox@company.com",
    client_folder="Inbox",
)
```

```python
# List folders
response = outlook.list_folders()
if response.status_code == 200:
    print(response.content)
```

```python
# List unread messages
response = outlook.list_messages(filter="isRead ne true")
if response.status_code == 200:
    print(response.content)
```

```python
# Move message to another folder
response = outlook.move_message(id="<message-id>", to="<folder-id>")
if response.status_code == 201:
    print("Message moved")
```

```python
# Delete message
response = outlook.delete_message(id="<message-id>")
if response.status_code == 204:
    print("Message deleted")
```

```python
# Download attachments from a message
response = outlook.download_message_attachment(
    id="<message-id>",
    path=r"C:\\Temp\\attachments",
    index=True,
)
if response.status_code == 200:
    print("Attachments downloaded")
```

```python
# Send HTML email with optional attachments
response = outlook.send_message(
    recipients=["person@example.com"],
    subject="Status update",
    message="<p>Hello team,</p><p>All good.</p>",
    attachments=None,
)
if response.status_code in (200, 202):
    print("Email sent")
```

```python
# Change working context
outlook.change_folder("Inbox")
outlook.change_client_email("another.mailbox@company.com")
```

For technical and contributor documentation, see [DEV.md](DEV.md).

## Version

Recommended way to read the installed package version:

```python
from importlib.metadata import version

print(version("outlooklib"))
```

Convenience attribute (also available):

```python
import outlooklib

print(outlooklib.__version__)
```

## Licence

BSD-3-Clause Licence (see [LICENSE](LICENSE))
