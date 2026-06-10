# Extract Specific Columns in process.py

This implementation plan outlines the updates to `process.py` to extract values for only the 6 columns configured in the duplicate template script (`create_duplicate.py`):
- Business Name (Original Column A)
- Location (Original Column B)
- Website / Online Presence (Original Column C)
- Phone (Original Column D)
- Email (Original Column E)
- Sources Verified (Original Column G)

The script will read rows 4 through 18, retrieve the values for these columns, and output the extracted records.

## Proposed Changes

### Lead Generation Pipeline

#### [MODIFY] [process.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/process.py)
Update `process.py` to:
1. Define the 6 columns to extract based on their original columns (A, B, C, D, E, G).
2. Loop through the rows (from row 4 to the last row).
3. Extract the cell value for each of the 6 columns in each row.
4. Extract the hyperlink targets for the Website (Column C) and Email (Column E) cells if they exist.
5. Print or display the extracted structured data.

## Verification Plan

### Automated Tests
- Run `process.py` and verify that only the 6 requested fields are displayed for all data rows (e.g., Wilkins Roofing, Mooney's Roofing, etc.):
  ```bash
  .venv/bin/python process.py
  ```
