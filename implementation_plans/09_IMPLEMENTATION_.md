# Update Spreadsheet Output to Standard Row Heights

This implementation plan outlines the updates to `process.py` to remove the custom row height settings for data rows (rows 4 and onwards), resetting them to use Excel's standard/default row heights.

## Proposed Changes

### Lead Generation Pipeline

#### [MODIFY] [process.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/process.py)
Update `process.py` to:
1. Remove the custom row height assignment loop that sets data rows (rows 4 to 12) to `72.0`.
2. Keep the custom row heights for Rows 1 (30.0), 2 (16.0), and 3 (27.75) as they contain headers/metadata.

## Verification Plan

### Automated Tests
- Run `process.py`:
  ```bash
  .venv/bin/python process.py
  ```
- Update the validation script `verify_emails.py` to assert that row heights for rows 4 to 12 are `None` (standard default).
