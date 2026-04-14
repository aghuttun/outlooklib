# outlooklib

- [Description](#package-description)
- [Usage](#usage)
- [Installation](#installation)
- [Development](#development)
- [License](#license)

## Package Description

Microsoft Outlook client package for Python, powered by a Rust core and PyO3 bindings.

The public Python API keeps the same names used in previous versions:

- `Outlook`
- `Response`
- `renew_token`, `change_client_email`, `change_folder`
- `list_folders`, `list_messages`, `move_message`, `delete_message`
- `download_message_attachment`, `send_message`

## Usage

From a script:

```python
import outlooklib
import pandas as pd

client_id = "123"
client_secret = "123"
tenant_id = "123"

client_email = "team.email@company.com"

# 0. Initialise Outlook client
outlook = outlooklib.Outlook(
    client_id=client_id,
    tenant_id=tenant_id,
    client_secret=client_secret,
    client_email=client_email,
    client_folder="Inbox"
)
```

```python
# 1. Retrieve a list of mail folders
response = outlook.list_folders()

if response.status_code == 200:
    df = pd.DataFrame(response.content)
    display(df)
```

```python
# 2.1. Retrieve the top 100 unread messages from the specified folder
response = outlook.list_messages(filter="isRead ne true")

if response.status_code == 200:
    df = pd.DataFrame(response.content)
    display(df)
```

```python
# 2.2. Retrieve the top 100 messages from the specified folder, with more than 2 days
import datetime

n_days_ago = (datetime.datetime.now(datetime.UTC) - datetime.timedelta(days=2)).strftime("%Y-%m-%dT%H:%M:%SZ")

response = outlook.list_messages(filter=f"receivedDateTime le {n_days_ago}")

if response.status_code == 200:
    df = pd.DataFrame(response.content)
    print(df)
```

```python
# 3. Download message attachments
message_id = "A...A=="

response = outlook.download_message_attachment(
    id=message_id,
    path=r"C:\Users\admin",
    index=True,
)

if response.status_code == 200:
    print("Attachment(s) downloaded successfully")
```

```python
# 4.1. Delete a message from the current folder
message_id = "A...A=="

response = outlook.delete_message(id=message_id)

if response.status_code == 204:
    print("Message deleted successfully")
```

```python
# 4.2. Delete messages from the current folder, one by one, with more than 3 days
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
# 5. Send an email with HTML body and optional attachments
response = outlook.send_message(
    recipients=["peter.parker@example.com"],
    subject="Web tests",
    message="Something<br>to talk about...",
    attachments=None,
)

# In Microsoft Graph this operation commonly returns 202 Accepted.
if response.status_code in (200, 202):
    print("Email sent")
```

```python
# 6. Change current folder
outlook.change_folder(id="root")
```

```python
# Close
del outlook
```

## Installation

For production:

```bash
pip install outlooklib
```

For local development (editable):

```bash
git clone https://github.com/aghuttun/outlooklib.git
cd outlooklib
pip install -e ".[dev]"
```

## Development

### Build Python wheel from Rust

```bash
maturin build --release --features python --out dist
```

The wheel will be generated under `dist/`.

### Install local wheel

```bash
pip install --force-reinstall dist/outlooklib-*.whl
```

### Run Rust tests

```bash
cargo test
```

### Run Python smoke tests

```bash
python -m pytest tests/test_smoke.py -q
```

### Optional live Microsoft Graph test

The live test is skipped unless all variables are set:

- `OUTLOOKLIB_CLIENT_ID`
- `OUTLOOKLIB_TENANT_ID`
- `OUTLOOKLIB_CLIENT_SECRET`
- `OUTLOOKLIB_CLIENT_EMAIL`

You can copy values from `.env.example` and export them in your shell.

Run:

```bash
python -m pytest tests/test_live_graph.py -q
```

### Publish to PyPI

The GitHub Actions workflow builds wheels for every push and pull request.

For release publishing, create and push a tag in the `v*` format:

```bash
git tag v0.0.14
git push origin v0.0.14
```

The `publish-pypi` job uses trusted publishing via `pypa/gh-action-pypi-publish`.
Configure the repository as a trusted publisher in PyPI for the project.

Suggested trusted publisher settings in PyPI:

- Owner: `aghuttun`
- Repository: `outlooklib`
- Workflow file: `.github/workflows/python-wheel.yml`
- Environment name: leave empty (not required by this workflow)

Optional CI live test support:

If you add these repository secrets, the workflow also runs `tests/test_live_graph.py`:

- `OUTLOOKLIB_CLIENT_ID`
- `OUTLOOKLIB_TENANT_ID`
- `OUTLOOKLIB_CLIENT_SECRET`
- `OUTLOOKLIB_CLIENT_EMAIL`

## Notes

Python files use descriptive docstrings and Rust files use rustdoc comments.

## License

BSD Licence (see LICENSE file)

[top](#outlooklib)
