# Duplicate Excel Spreadsheet from Scratch

This implementation plan outlines the steps to create a Python script that builds an exact duplicate of the Excel file `input/01_Roofing_Cold_Outreach_Leads_June2026.xlsx` from scratch using `openpyxl`.

## User Review Required

> [!NOTE]
> The target file `input/01_Roofing_Cold_Outreach_Leads_June2026.xlsx` has a single sheet named `"Ranked Leads"`.
> Rows 1 and 2 are merged ranges (`A1:H1` and `A2:H2`) without text values, but with styling (Navy blue and Light gray solid fills respectively).
> Column `C` and `E` contain URLs/emails styled as standard Excel Hyperlinks (Calibri 11pt, single underline, theme color 10), and they have hyperlink targets attached (e.g. `mailto:` for emails), except cell `C18` which has the Hyperlink style but *no* hyperlink object target.
> We will replicate all of these visual and behavioral details exactly.

## Proposed Changes

### Lead Generation Pipeline

#### [NEW] [create_duplicate.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/create_duplicate.py)
A Python script that creates the exact duplicate of `input/01_Roofing_Cold_Outreach_Leads_June2026.xlsx` and saves it to a specified destination (e.g., `output/01_Roofing_Cold_Outreach_Leads_June2026_duplicate.xlsx`).

Key functionalities of the script:
1. Define the exact raw data structure for rows 3 to 18.
2. Initialize an openpyxl Workbook, rename the default sheet to `"Ranked Leads"`.
3. Set the column dimensions (widths) and row dimensions (heights) to match the original.
4. Merge cells `A1:H1` and `A2:H2`.
5. Apply styling:
   - Fills: Navy blue (`FF1F3864`) for Row 1, Light gray (`FFF2F2F2`) for Row 2, Medium blue (`FF2E75B6`) for the header row, and alternating Light green (`FFE2EFDA`) and White (`FFFFFFFF`) for data rows.
   - Fonts: Arial for standard cells, Calibri 11pt blue/underlined for hyperlink-styled cells.
   - Alignments: Centered for headers/metadata, top-left aligned with text wrap for data cells.
   - Borders: Thin light-gray (`FFBFBFBF`) gridlines.
6. Attach hyperlink targets (`mailto:` for emails, `https://...` for web links) matching the original.
7. Configure sheet view: Zoom scale 110%, tab selected.
