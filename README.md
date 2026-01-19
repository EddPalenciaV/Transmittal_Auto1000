# TransmittalFlow

Automates the creation and maintenance of engineering project transmittal Excel workbooks by discovering PDF drawings, extracting metadata from filenames, updating transmittal sheets with revisions, and exporting submission-ready PDFs with high fidelity.

## Overview

**TransmittalFlow** eliminates manual, repetitive work in construction and engineering workflows. It scans project directories for PDF drawings, parses their metadata (project number, drawing number, revision, and title), automatically updates or inserts rows in a transmittal workbook, and exports the final sheet to a professional PDF for submission.

### Why This Matters

- **Saves Hours**: Automates what would otherwise be manual copying, pasting, and formatting.
- **Reduces Errors**: Programmatic extraction and validation ensure consistency.
- **Preserves Fidelity**: Uses Excel COM automation to maintain formulas, formatting, and page layout in exported PDFs.
- **Scales Efficiently**: Handles hundreds of drawings and reorganizes them intelligently by group and sequence.

---

## Features

- **Recursive PDF Discovery**: Scans current folder and all subfolders for drawing PDFs matching a specific naming convention.
- **Intelligent Metadata Extraction**: Uses robust regex patterns to parse:
  - Project number (before department letter)
  - Drawing number (group and sequence)
  - Revision code (inside square brackets)
  - Drawing title (between bracket and file extension)
- **Smart Row Management**:
  - Matches drawings by title; updates revisions in existing rows.
  - Inserts new drawings in the correct position within their group (by sequence number).
  - Preserves formulas and cell formatting via xlwings COM automation.
- **Date-Aware Workflow**: Prompts for transmittal date; reuses existing date columns or creates new ones.
- **Professional PDF Export**: Exports the selected worksheet to A4 PDF with default margins using Excel's native export.
- **Defensive Error Handling**: Validates file existence, handles corrupted workbooks, detects missing columns, and manages process cleanup.

---

## Requirements

### System

- **Windows OS** (required for Excel COM automation via pywin32)
- **Microsoft Excel** installed and available

### Python

- Python 3.8 or later (tested on 3.9+)
- Virtual environment recommended

### Python Packages

- `openpyxl` — Read and write Excel workbooks
- `xlwings` — Excel automation (high-level API)
- `pywin32` — Excel COM automation (low-level control)
- `pathlib`, `datetime`, `re`, `glob`, `shutil` — Standard library

---

## Installation

### 1. Clone or Download the Repository

```bash
git clone https://github.com/EddPalenciaV/Transmittal_Auto1000.git
cd Transmittal_Auto1000
```

### 2. Create and Activate a Virtual Environment

```powershell
# PowerShell
python -m venv .venv
.venv\Scripts\Activate.ps1

# Command Prompt
python -m venv .venv
.venv\Scripts\activate.bat
```

### 3. Install Dependencies

```bash
pip install openpyxl xlwings pywin32
```

### 4. Configure Paths (Optional)

Edit `Transmit_Auto1000.py` and update the template file path in the `find_excel_transmittal()` function:

```python
template_file = r"C:\path\to\your\Transmittal_TEMPLATE.xlsx"
```

---

## Usage

### Running the Script

Activate your virtual environment, then run:

```powershell
python Transmit_Auto1000.py
```

### Workflow Steps

1. **Select Transmittal Date**: Choose today's date or enter a custom date (DD/MM/YY format).
2. **Locate/Create Transmittal**: The script finds the latest transmittal workbook or copies a template.
3. **Choose Sheet**: Select which sheet to update (CIVIL, ARCHITECT, STRUCTURE).
4. **Scan Drawings**: The script discovers PDF files in the current folder and subfolders.
5. **Parse & Match**: Drawing metadata is extracted and compared against existing transmittal entries.
6. **Update/Insert**: Matching drawings have revisions updated; new drawings are inserted in the correct position.
7. **Save & Export**: The workbook is saved and the selected sheet is exported to PDF.

### Expected Output

```
Transmit_Auto1000 Start
Searching for Transmittal Excel file...
The latest transmittal excel is: C:\Users\...\Transmittal 250119.xlsx
Choose sheet to process:
1. CIVIL
2. STRUCTURE
3. ARCHITECT
4. Exit Program
Enter your choice (1-3): 1
CIVIL sheet found.
Found revision column at index: 7
Searching for drawings in current folder and subfolders...
Found 5 drawings in the current folder.

Processing drawing file: C:\...\10009-C-00-01 [A] COVER SHEET.pdf
Matched 'COVER SHEET'. Updated revision to 'A' at row 24.
...
Changes saved in Transmittal excel 'Transmittal 250119.xlsx'.
Workbook closed successfully.

Converting 'CIVIL' from '...\Transmittal 250119.xlsx' to PDF using MS Excel...
Opening file: C:\...\Transmittal 250119.xlsx
Successfully saved PDF to 'C:\...\Transmittal 250119.pdf'
Excel application closed.

Created by Edd Palencia-Vanegas - June 2024. All rights reserved.
Version 5.1 - 13/01/2026
```

---

## PDF Filename Convention

The script expects PDF filenames to follow this pattern:

```
<ProjectNumber>-<Department>-<Group>-<Sequence> [<Revision>] <Title>.pdf
```

### Examples

```
10009-C-00-01 [A] COVER SHEET, LOCALITY PLAN AND DRAWING INDEX.pdf
10009-C-00-02 [F] GENERAL NOTES AND LEGEND.pdf
10009-C-10-03 [I] WATER PLAN SHEET 2.pdf
10014-C-00-11 [13] GENERAL ARRANGEMENT PLAN - SITE LAYOUT.pdf
```

### Naming Breakdown

- **ProjectNumber**: Numeric identifier (e.g., `10009`)
- **Department**: Single letter (C, A, S, etc.)
- **Group**: Two digits (e.g., `00`, `10`)
- **Sequence**: Two digits (e.g., `01`, `03`, `11`)
- **Revision**: Single character or digit inside brackets (e.g., `[A]`, `[13]`)
- **Title**: Description of the drawing content

---

## Project Structure

```
Transmittal_Auto1000/
├── Transmit_Auto1000.py        # Main script
├── Transmittal_TEMPLATE.xlsx   # Excel template (if applicable)
├── .venv/                       # Virtual environment
├── README.md                    # This file
├── .gitignore                   # Git ignore rules
└── [PDF drawings]               # Input: drawing files
    └── [subfolders/]            # Script searches recursively
```

---

## Key Functions

| Function                   | Purpose                                                              |
| -------------------------- | -------------------------------------------------------------------- |
| `Update_Transmittal()`     | Main logic: loads workbook, matches/inserts drawings, saves workbook |
| `Catch_Drawings()`         | Discovers PDF files matching naming convention recursively           |
| `find_excel_transmittal()` | Locates latest transmittal or copies template                        |
| `Request_Get_Date()`       | Prompts user for date; creates/updates transmittal with date headers |
| `Save_as_PDF()`            | Exports selected worksheet to PDF via Excel COM                      |

---

## Configuration

### Transmittal Template Path

Located in `find_excel_transmittal()`:

```python
template_file = r"C:\Users\eddpa\Desktop\GoalFolder\Transmittal_TEMPLATE.xlsx"
```

Update this to point to your company's transmittal template.

### Drawing Search Range

In `Update_Transmittal()`, the script searches rows 24–150 for drawing entries:

```python
for row_index in range(24, 151):
```

Adjust if your template uses different row ranges.

### Revision Column Detection

The script finds the revision column by locating the first empty cell in row 1, columns 5–30:

```python
for col_index in range(5, 31):
```

Customize if your date/revision layout differs.

---

## Troubleshooting

### "Revision column not found in worksheet"

- Ensure your transmittal template has a clear date row (row 1) with columns 5+ containing dates or being empty.
- The script expects the first empty cell to mark the start of revision columns.

### "No PDF drawings found to process"

- Verify PDF files are in the current working directory or subfolders.
- Check that filenames match the expected pattern (must contain `[X]` where X is alphanumeric).

### COM Error: "Open method of Workbooks class failed"

- Ensure no other Excel instance is using the transmittal file.
- Check file path contains only backslashes (script converts them automatically).
- Verify the file is not corrupted by opening it manually in Excel first.

### "Sheet 'CIVIL' not found"

- Confirm the transmittal template contains sheets named CIVIL, ARCHITECT, and/or STRUCTURE.
- Use Excel to verify sheet names match exactly (case-sensitive when accessed via COM).

### File Lock Issues

- The script includes a 5-second delay (`time.sleep(5)`) after closing workbooks to allow the OS to release locks.
- If issues persist, increase the delay or ensure no background Excel processes are running.

---

## Advanced Usage

### Batch Processing Multiple Sheets

Modify the script to loop through sheet choices or pass sheet name as a command-line argument:

```python
for sheet_choice in ['1', '2', '3']:
    # Run Update_Transmittal() for each sheet
```

### Custom Logging

Add logging to track all updates:

```python
import logging
logging.basicConfig(filename='transmittal.log', level=logging.INFO)
logging.info(f"Updated {drawing_Name} at row {row}")
```

### Integration with CI/CD

Wrap the main logic to accept command-line arguments (date, sheet, template path) for automated workflows.

---

## Contributing

Contributions are welcome! Please:

1. Fork the repository.
2. Create a feature branch: `git checkout -b feature/my-feature`
3. Commit changes: `git commit -m "Add my feature"`
4. Push to branch: `git push origin feature/my-feature`
5. Open a pull request.

### Areas for Enhancement

- Add support for macro-enabled workbooks (.xlsm).
- Implement fuzzy name matching to tolerate minor title variations.
- Add Excel file validation and repair on load.
- Extend to support other sheet types (MECHANICAL, ELECTRICAL, etc.).

---

## Known Limitations

- **Windows Only**: Requires Windows OS and Excel COM support.
- **Excel Dependency**: Must have Microsoft Excel installed.
- **Sparklines**: Excel sparkline objects are not preserved when modified by openpyxl.
- **Large Workbooks**: Performance may degrade with transmittals containing 500+ rows.

---

## License

This project is provided as-is. All rights reserved by the author.

---

## Author

**Edd Palencia-Vanegas**  
Created: June 2024  
Latest Update: January 13, 2026  
Version: 5.1

For questions, suggestions, or bug reports, please open an issue on the GitHub repository.

---

## Changelog

### Version 5.1 (13/01/2026)

- Stabilized Excel COM automation and process cleanup.
- Improved error messages for missing sheets and columns.
- Refactored date discovery logic to be more robust.

### Version 5.0 (Earlier)

- Initial public release with xlwings integration.
- Smart row insertion based on drawing group and sequence.
- Automated PDF export with A4 formatting.

---

## Acknowledgments

Built with:

- [openpyxl](https://openpyxl.readthedocs.io/) — Python Excel library
- [xlwings](https://www.xlwings.org/) — Excel automation
- [pywin32](https://github.com/pywin32/pywin32) — Windows COM support
