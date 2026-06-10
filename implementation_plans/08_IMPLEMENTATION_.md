# Filter Out Leads with No Emails in process.py

This implementation plan outlines the updates to `process.py` to filter out lead records that do not contain a valid email address before writing the data rows to the output spreadsheet template.

A lead record will be filtered out (excluded) if its Email column value (Column E) is:
- `None` or an empty string `""`
- The literal string `'No email found'` (case-insensitive)

## Proposed Changes

### Lead Generation Pipeline

#### [MODIFY] [process.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/process.py)
Update the extraction/population pipeline in `process.py` to:
1. Load data from the input file.
2. During extraction, check the value of the Email column (original column 5 / Column E).
3. Filter out any rows where the email is empty, `None`, or equal to `'No email found'`.
4. Only populate the remaining lead records into the output template starting at row 4.
5. Set row heights dynamically based on the number of filtered leads to avoid hardcoding heights for empty rows.

## Verification Plan

### Automated Tests
- Run `process.py`:
  ```bash
  .venv/bin/python process.py
  ```
- Write a validation script `verify_emails.py` to verify that:
  - The generated duplicate `output/Roofing_Outreach_Leads_Template.xlsx` contains only rows where the email is populated with a valid email.
  - The number of data rows in the output matches the number of filtered rows.
  - Grid properties and formatting are applied properly to the new number of data rows.
