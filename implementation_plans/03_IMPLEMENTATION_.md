# Adjust Sheet Columns and Merges

This implementation plan outlines the steps to adjust the spreadsheet template so it only includes the 6 requested columns:
- Business Name
- Location
- Website / Online Presence
- Phone
- Email
- Sources Verified

Columns `Opportunity Reason` and `Validated` will be removed. The merges for rows 1 and 2 will be adjusted from `A1:H1` and `A2:H2` to `A1:F1` and `A2:F2` соответственно.

## Proposed Changes

### Lead Generation Pipeline

#### [MODIFY] [create_duplicate.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/create_duplicate.py)
Modify the script to:
1. Update column widths to reflect only the 6 required columns:
   - A (Business Name): 28.0
   - B (Location): 16.0
   - C (Website / Online Presence): 35.0
   - D (Phone): 16.0
   - E (Email): 29.6640625
   - F (Sources Verified): 38.0
2. Update row merges to `A1:F1` and `A2:F2`.
3. Clear references/styles to columns G and H.
4. Set column headers in Row 3 for only the 6 columns.
5. Apply styling and alternating fills to only columns A through F.

## Verification Plan

### Automated Tests
- Run the modified duplication script to generate the spreadsheet:
  ```bash
  .venv/bin/python create_duplicate.py
  ```
- Write a validation script `verify_columns.py` to assert that:
  - Total columns in sheet are 6.
  - Column names in row 3 are exactly the 6 requested columns.
  - Title merges are `A1:F1` and `A2:F2`.
  - Columns A to F have the correct formatting, styles, and heights/widths preserved.
  - Columns G and H do not exist in the output sheet views/dimensions.
