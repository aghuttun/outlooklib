# outlooklib

* [Description](#package-description)
* [Usage](#usage)
* [Installation](#installation)
* [License](#license)

## Package Description

Microsoft Outlook client Python package that uses the [requests](https://pypi.org/project/requests/) library.

> [!IMPORTANT]  
> This packages uses pydantic~=1.0!

## Usage

* [outlooklib](#outlooklib)

from a script:

```python
import outlooklib
import logging
```

## Installation

* [outlooklib](#outlooklib)

Install python and pip if you have not already.

Then run:

```bash
pip3 install pip --upgrade
pip3 install wheel
```

For production:

```bash
pip3 install outlooklib
```

This will install the package and all of it's python dependencies.

If you want to install the project for development:

```bash
git clone https://github.com/aghuttun/outlooklib.git
cd outlooklib
pip3 install -e .[dev]
```

To test the development package: [Testing](#testing)

## License

* [outlooklib](#outlooklib)

BSD License (see license file)
