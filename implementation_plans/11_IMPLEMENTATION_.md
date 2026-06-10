# Format Leads Spreadsheet to Match Column Order

This implementation plan outlines the creation of a new Python script `format_leads.py` that formats the workbook `input/02_Roofing_Cold_Outreach_Leads_June2026.xlsx` to use the same columns, in the same order, and with the same styles/widths as `input/01_Roofing_Cold_Outreach_Leads_June2026.xlsx`.

Specifically:
- Drop columns `Rank` (Column A) and `Website Status` (Column D).
- Reorder/shift the remaining 8 columns to match `01` (`Business Name`, `Location`, `Website / Online Presence`, `Phone`, `Email`, `Opportunity Reason`, `Sources Verified`, `Validated`).
- Adjust Column widths to match `01`.
- Update merges for Rows 1 and 2 to span `A1:H1` and `A2:H2` respectively.
- Retain all cell-level values, formatting, colors, font styling, wrap-text alignment, borders, and hyperlinks.

## Proposed Changes

### Lead Generation Pipeline

#### [NEW] [format_leads.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/format_leads.py)
Create a new python script that:
1. Loads `input/02_Roofing_Cold_Outreach_Leads_June2026.xlsx`.
2. Map target columns A-H to their respective original columns (B, C, E, F, G, H, I, J).
3. Reads the data and cell style configurations from the original sheet.
4. Generates a new workbook with the 8 columns, copying over styles (font properties, solid fills, alignments, borders, hyperlinks) for each shifted cell.
5. Updates column widths to `01`'s settings.
6. Re-merges `A1:H1` and `A2:H2`.
7. Saves the modified spreadsheet back to `input/02_Roofing_Cold_Outreach_Leads_June2026.xlsx`.

## Verification Plan

### Automated Tests
- Run `format_leads.py`:
  ```bash
  .venv/bin/python format_leads.py
  ```
- Write a validation script `verify_leads_format.py` that compares `input/02_Roofing_Cold_Outreach_Leads_June2026.xlsx` (after formatting) against `input/01_Roofing_Cold_Outreach_Leads_June2026.xlsx` schema to verify that dimensions are `A1:H17`, column headers match exactly, column widths match exactly, and merges are `A1:H1` / `A2:H2`.
