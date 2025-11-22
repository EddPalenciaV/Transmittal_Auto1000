Automates creation and upkeep of project transmittals by scanning PDF drawing files, extracting drawing metadata, updating a Transmittal Excel workbook, and exporting the final sheet to PDF. Built for Windows with MS Excel integration for high‑fidelity PDF output.

Why this is useful

Saves hours of repetitive manual work keeping transmittals up to date. Parses drawing filenames to extract project number, drawing number, revision, and title. Updates an existing Transmittal workbook (CIVIL sheet) or creates one from a template. Preserves Excel formulas and page layout; exports to PDF using Excel COM for exact rendering. Designed for real office workflows — recruiters/employers can see applied regex parsing, Excel automation, and production-quality considerations (error handling, templating, and venv usage).

Core features

Discover a Transmittal workbook (or copy template if missing). Scan current folder for PDF drawings with bracketed revision tokens (e.g., "10009-C-00-01 [0] COVER SHEET.pdf"). Extract values using robust regex patterns: project number before -C (e.g., 10009) drawing number between five-digit prefix and [ revision inside [...] drawing title between ] and .pdf Update matching rows (by drawing title) and revision column; add new drawings into the sheet when not found. Let user select transmittal date or use today’s date. Export the transmittal sheet to PDF (A4, default margins) using MS Excel via COM (pywin32) for accurate rendering.

Requirements

Windows with Microsoft Excel installed (required for COM export). Python 3.8+ (project tested with 3.10+). Recommended to run inside a virtual environment.

Python packages

openpyxl pywin32 (optional) spire.xls — commented in code (not required for COM method)

Quick install (PowerShell) python -m venv .venv .venv\Scripts\Activate.ps1 pip install openpyxl pywin32

Example usage

Activate your virtual environment (see above). Place your PDF drawings in the project folder (naming format examples in “Filename expectations”). Run the script: python Transmit_1000_v.5.py

What the script does when run (default flow)

Prompts for transmittal date (today or custom). Ensures a Transmittal workbook is available (copies template if needed). Scans for PDFs and parses metadata. Updates the CIVIL sheet: updates revisions for existing drawings or inserts new rows for missing ones. Saves the transmittal workbook and exports the CIVIL sheet to PDF (A4).

Filename expectations (examples)

10009-C-00-01 [0] COVER SHEET, LOCALITY PLAN AND DRAWING INDEX.pdf 10009-C-00-02 [F] GENERAL NOTES AND LEGEND.pdf 10009-C-10-03 [I] WATER PLAN SHEET 2.pdf

Notes and limitations

Excel COM export requires MS Excel installed and is Windows only. The script expects specific naming patterns; adjust the regex patterns in the code if your naming convention differs. The script expects a sheet named "CIVIL" in the Transmittal workbook. Keep a backup of your Transmittal workbook before bulk operations.

Extending the project

Add command-line flags (argparse) to skip interactive prompts or to select other sheets. Improve fuzzy matching for drawing titles to tolerate small naming differences. Move patterns and column indexes to a config file for easy adaptation across projects. Add unit tests for regex extraction and row-update logic.

Contributing

Fork the repository, create a feature branch, and open a pull request with a clear description. Keep commits small and focused; include examples when changing parsing behavior.

License

MIT.

Contact

Inspect the code in Transmit_1000_v.5.py for implementation details (regex patterns, Excel interaction, and workflow).