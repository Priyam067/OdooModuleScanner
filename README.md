# Odoo Modules Data Exporter

This Python script scans a given directory for Odoo modules and generates an Excel file containing detailed information about each module, including:

- Module Name
- Path to Module
- Dependent Modules
- Other relevant module information extracted from the manifest files.

## Features

- **Automatic Scanning**: Recursively scans a directory for all Odoo modules.
- **Dependencies Check**: Lists dependent modules for each module.
- **Excel Export**: Outputs all module data into a structured Excel file for easy review using `xlwt`.

## Requirements

- Python 3.x
- `os` standard library for directory scanning.
- `ast` for parsing Python files.
- `xlwt` for generating Excel files.

To install the required package, run:
```bash
pip install xlwt
```

## Usage

1. Clone the repository or download the script to your local machine.
2. Place the script inside the directory where your Odoo modules are located or provide the path to the directory in the script.
3. Run the script:
```bash
python3 main.py
```
4.The script will generate an Excel file named odoo_modules_data.xls in the current directory.
