# Rename Output Spreadsheet Template

This implementation plan outlines the steps to rename the generated Excel spreadsheet template to a more descriptive and generic name suitable for template reuse.

We propose renaming the output file from:
`output/01_Roofing_Cold_Outreach_Leads_June2026.xlsx`
to:
`output/Roofing_Outreach_Leads_Template.xlsx`

## Proposed Changes

### Lead Generation Pipeline

#### [MODIFY] [create_duplicate.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/create_duplicate.py)
Update the output saving path in `create_duplicate.py` to:
`output/Roofing_Outreach_Leads_Template.xlsx`

#### [DELETE] [01_Roofing_Cold_Outreach_Leads_June2026.xlsx](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/output/01_Roofing_Cold_Outreach_Leads_June2026.xlsx)
Remove the old output template file.

## Verification Plan

### Automated Tests
- Run `create_duplicate.py` to generate the new file.
- Verify `output/Roofing_Outreach_Leads_Template.xlsx` is created.
- Run `verify_columns.py` updated to verify the new file name.
