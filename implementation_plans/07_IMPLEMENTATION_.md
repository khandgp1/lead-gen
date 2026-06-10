# Populate Output Spreadsheet with Extracted Info in process.py

This implementation plan outlines the steps to update `process.py` to extract the lead data from the input spreadsheet and write it into a formatted 6-column output spreadsheet template matching the exact styling guidelines defined in `create_duplicate.py`.

## Proposed Changes

### Lead Generation Pipeline

#### [MODIFY] [process.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/process.py)
Update `process.py` to:
1. Extract the raw data from `input/01_Roofing_Cold_Outreach_Leads_June2026.xlsx` into a structured list of dictionaries. Ensure we capture:
   - Values for Business Name, Location, Website, Phone, Email, and Sources Verified.
   - Hyperlink targets for Website (Column C) and Email (Column E) if they exist.
2. Initialize a new openpyxl Workbook and apply the layout/style definitions from `create_duplicate.py`:
   - Set column dimensions and row heights.
   - Configure SheetView (zoom scale 110%, gridlines).
   - Merge `A1:F1` and `A2:F2`.
   - Setup Row 1 and Row 2 title cells (navy and light gray fills).
   - Setup Row 3 column headers (Business Name, Location, Website / Online Presence, Phone, Email, Sources Verified) with blue fill, bold white font, medium top border, and text wrap.
3. For each extracted lead record, write the values to rows 4-18:
   - Columns A to F.
   - Recreate actual hyperlink objects for Website (Column C) and Email (Column E) using their extracted targets.
   - Apply the correct styles (Arial 9pt left/top aligned, wrap text, thin gray borders).
   - Apply alternating background fills: Even rows (4, 6, 8...) get light green (`FFE2EFDA`), and Odd rows (5, 7, 9...) get white (`FFFFFFFF`).
   - Style cells with hyperlink targets using the Calibri 11pt blue/underlined hyperlink font style.
4. Save the final populated workbook to `output/Roofing_Outreach_Leads_Template.xlsx` (or a configured output name).

## Verification Plan

### Automated Tests
- Run the updated pipeline:
  ```bash
  .venv/bin/python process.py
  ```
- Run the validator `verify_columns.py` but updated to assert that cell values in rows 4-18 are now populated with the correct raw data and hyperlink objects instead of being empty.
