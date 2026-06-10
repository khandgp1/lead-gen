# Remove Wrap Text from Column C in process.py

This implementation plan outlines the steps to remove the wrap text alignment feature from Column C (Website / Online Presence) cells (both header C3 and data cells in column C).

## Proposed Changes

### Lead Generation Pipeline

#### [MODIFY] [process.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/process.py)
Update `process.py` to:
1. When generating header row 3, configure alignment for column C (`C3`) to use `wrap_text=False`.
2. When writing data rows (rows 4 and onwards), configure the alignment for cells in column C to use `wrap_text=False` (or omit/disable `wrap_text`), while keeping `wrap_text=True` for other columns.

## Verification Plan

### Automated Tests
- Run `process.py`:
  ```bash
  .venv/bin/python process.py
  ```
- Update the validation script `verify_emails.py` to assert that all cells in Column C (from row 3 to 12) have `wrap_text` set to `False` or `None`.
