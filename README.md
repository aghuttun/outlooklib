# outlooklib

- [Description](#package-description)
- [Usage](#usage)
- [Installation](#installation)
- [License](#license)

## Package Description

Microsoft Outlook client Python package that uses the [requests](https://pypi.org/project/requests/) library.

> [!IMPORTANT]  
> This packages uses pydantic v1!

## Usage

From a script:

```python
import outlooklib
import pandas as pd

client_id = "123"
client_secret = "123"
tenant_id = "123"

client_email = "team.email@company.com"

# Initialize Outlook client
outlook = outlooklib.Outlook(client_id=client_id,
                             tenant_id=tenant_id,
                             client_secret=client_secret,
                             client_email=client_email,
                             client_folder="Inbox")
```

```python
# Retrieve a list of mail folders
response = outlook.list_folders()
if response.status_code == 200:
    df = pd.DataFrame(response.content)
    print(df)
```

```python
# Retrieve the top 100 unread messages from the specified folder
response = outlook.list_messages(filter="isRead ne true")
if response.status_code == 200:
    df = pd.DataFrame(response.content)
    print(df)
```

```python
# Retrieve the top 100 messages from the specified folder, with more than 2 days
import datetime

n_days_ago = (datetime.datetime.now(datetime.UTC) - datetime.timedelta(days=2)).strftime("%Y-%m-%dT%H:%M:%SZ")

response = outlook.list_messages(filter=f"receivedDateTime le {n_days_ago}")
if response.status_code == 200:
    df = pd.DataFrame(response.content)
    print(df)
```

```python
# Download message attachments
message_id = "A...A=="
response = outlook.download_message_attachment(id=message_id, path=r"C:\Users\admin", index=True)
if response.status_code == 200:
    print("Attachment(s) downloaded successfully")
```

```python
# Delete a message from the current folder
message_id = "A...A=="
response = outlook.delete_message(id=message_id)
if response.status_code == 204:
    print("Message deleted successfully")
```

```python
# Delete messages from the current folder, one by one, with more than 3 days
import datetime

x_days_ago = (datetime.datetime.now(datetime.UTC) - datetime.timedelta(days=3)).strftime("%Y-%m-%dT%H:%M:%SZ")
response = outlook.list_messages(filter=f"receivedDateTime le {x_days_ago}")
if response.status_code == 200:
    df = pd.DataFrame(response.content)
    display(df)

    for msg_id in df["id"]:
        response = outlook.delete_message(id=msg_id)
        if response.status_code == 204:
            print(f"Message {msg_id} deleted successfully")
```

```python
# Send an email with HTML body and optional attachments
response = outlook.send_message(recipients=["peter.parker@example.com"],
                                subject="Web tests",
                                message="Something<br>to talk about...",
                                attachments=None)
if response.status_code == 200:
    print("Email sent")
```

```python
# Change current folder
outlook.change_folder(id="root")
```

```python
# Close
del(outlook)
```

## Installation

Install python and pip if you have not already.

Then run:

```bash
pip install pip --upgrade
```

For production:

```bash
pip install outlooklib
```

This will install the package and all of it's python dependencies.

If you want to install the project for development:

```bash
git clone https://github.com/aghuttun/outlooklib.git
cd outlooklib
pip install -e ".[dev]"
```

## Docstring

The script's docstrings follow the numpydoc style.

## License

BSD License (see license file)

[top](#outlooklib)
