# Create process.py File Loader

This implementation plan outlines the creation of a new Python script `process.py` in the root of the project directory. The script's sole responsibility at this stage is to target and load the Excel workbook from the `input/` directory (`input/01_Roofing_Cold_Outreach_Leads_June2026.xlsx`) using `openpyxl` or `pandas`.

## Proposed Changes

### Lead Generation Pipeline

#### [NEW] [process.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/process.py)
Create a new Python script that:
1. Resolves the path to the Excel file inside the `input/` folder.
2. Loads the workbook using `openpyxl`.
3. Verifies success by printing basic metadata (e.g., sheet names, active sheet name, or dimension properties).

## Verification Plan

### Automated Tests
- Run the process script to confirm it successfully loads the workbook:
  ```bash
  .venv/bin/python process.py
  ```
