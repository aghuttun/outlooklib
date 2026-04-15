# outlooklib — Developer Guide

- [Overview](#overview)
- [Requirements](#requirements)
- [Version Access](#version-access)
- [Project Structure](#project-structure)
- [Building](#building)
- [API Reference](#api-reference)
- [Tests](#tests)
- [CI/CD](#cicd)
- [Publishing](#publishing)

## Overview

outlooklib is implemented in Rust using [PyO3](https://pyo3.rs/) for Python bindings and
[reqwest](https://crates.io/crates/reqwest) for Microsoft Graph HTTP operations.
[Maturin](https://www.maturin.rs/) is used as the build and packaging tool.

## Requirements

- Python 3.10 or higher
- Rust stable toolchain (edition 2024 crate)
- Maturin 1.7 or higher

```bash
pip install maturin
```

Install local development dependencies:

```bash
python -m pip install -e .[dev]
```

## Version Access

The recommended market standard is using installed package metadata:

```python
from importlib.metadata import version

print(version("outlooklib"))
```

The package also exposes `outlooklib.__version__` as a convenience alias.

## Project Structure

```
outlooklib/
├── src/
│   ├── lib.rs              # Rust Outlook/Graph core implementation
│   ├── models.rs           # Rust data models (Graph payloads)
│   └── python_bindings.rs  # PyO3 module and Python-facing wrappers
├── tests/
│   ├── test_smoke.py       # Public API smoke tests (no live Graph required)
│   └── test_live_graph.py  # Optional integration test against live Graph
├── old_python_version/     # Archived pure-Python implementation
├── .github/
│   └── workflows/
│       └── CI.yml          # GitHub Actions CI/CD workflow
├── Cargo.toml              # Rust crate configuration
├── pyproject.toml          # Python packaging configuration (maturin backend)
├── README.md               # User-facing documentation
└── DEV.md                  # This file
```

## Building

### Development build (fast iteration)

```bash
maturin develop --features python
```

### Release build (optimised)

```bash
maturin develop --release --features python
```

### Build distributable wheel

```bash
maturin build --release --features python --out dist
```

### Build source distribution

```bash
maturin sdist --out sdisthouse
```

## API Reference

### `Outlook(client_id, tenant_id, client_secret, client_email, client_folder="Inbox", custom_logger=None)`

Initialise a client and authenticate with Microsoft Graph.

| Parameter     | Type | Default | Description |
|---------------|------|---------|-------------|
| client_id     | str  | —       | Azure application client id |
| tenant_id     | str  | —       | Azure tenant id |
| client_secret | str  | —       | Azure application client secret |
| client_email  | str  | —       | Mailbox owner email used in Graph routes |
| client_folder | str  | Inbox   | Initial folder context |
| custom_logger | any  | None    | Ignored (kept for API compatibility) |

### Methods

| Method | Returns | Description |
|--------|---------|-------------|
| `renew_token()` | `None` | Re-authenticate and refresh bearer token |
| `change_client_email(email)` | `None` | Change mailbox context |
| `change_folder(id)` | `None` | Change current folder context |
| `list_folders(save_as=None)` | `Response` | List folders for current mailbox/folder context |
| `list_messages(filter, save_as=None)` | `Response` | List up to 100 messages matching OData filter |
| `move_message(id, to, save_as=None)` | `Response` | Move a message to another folder |
| `delete_message(id)` | `Response` | Delete a message |
| `download_message_attachment(id, path, index=False)` | `Response` | Download attachments from a message |
| `send_message(recipients, subject, message, attachments=None)` | `Response` | Send an HTML email with optional attachments |

### Response Shape

All operations that return `Response` expose:

- `status_code`: HTTP status from Microsoft Graph.
- `content`: decoded payload (`list`, `dict`, or `None`).

### Errors

Python-facing methods raise a Python `RuntimeError` when Rust operations fail.
The message is propagated from the internal `OutlookError` value.

## Tests

### Smoke tests (no live Graph required)

```bash
python -m pytest tests/test_smoke.py -q
```

### Integration test (live Graph required)

Set these environment variables:

```bash
export OUTLOOKLIB_CLIENT_ID=<client_id>
export OUTLOOKLIB_TENANT_ID=<tenant_id>
export OUTLOOKLIB_CLIENT_SECRET=<client_secret>
export OUTLOOKLIB_CLIENT_EMAIL=<mailbox_email>
```

Run:

```bash
python -m pytest tests/test_live_graph.py -q
```

Tests are skipped automatically if required variables are missing.

### Running all Python tests

```bash
python -m pytest tests/ -q
```

### Rust tests

```bash
cargo test -q
```

## CI/CD

The GitHub Actions pipeline in [.github/workflows/CI.yml](.github/workflows/CI.yml) runs on:

- pushes to `main` and `master`.
- pull requests.
- version tags (`v*`).
- manual workflow dispatch.

| Job | Trigger | Platforms |
|-----|---------|-----------|
| `test` | push / PR / tag / manual | ubuntu, windows, macos |
| `build-wheel` | after `test` | ubuntu, windows, macos |
| `build-sdist` | independent | ubuntu |
| `publish-pypi` | version tag (`v*`) only | ubuntu |

## Publishing

Publishing is automated via CI when a version tag is pushed:

```bash
git tag v0.1.0
git push origin v0.1.0
```

The pipeline uses the `PYPI_API_TOKEN` secret configured in repository settings.

To publish manually:

```bash
maturin publish --release --username __token__ --password "$PYPI_API_TOKEN"
```
