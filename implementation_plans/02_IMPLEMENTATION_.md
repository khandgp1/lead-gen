# Remove Data Population but Keep Formatting Logistics

This implementation plan outlines the steps to clean the spreadsheet duplicate creator script so it does not populate the actual lead data (cell values and hyperlink objects) for rows 4 to 18, while preserving all layouts, sheet view settings, row/column dimensions, headers, alternating fills, cell alignments, fonts, and border stylings for future use.

## Proposed Changes

### Lead Generation Pipeline

#### [MODIFY] [create_duplicate.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/create_duplicate.py)
Modify the script to:
1. Retain the sheet `"Ranked Leads"`, sheet views, column widths, row heights, and merged ranges (`A1:H1`, `A2:H2`).
2. Retain headers in Row 3 (e.g. `'Business Name'`, `'Location'`, etc.) along with their font, fill, border, alignment, and format properties.
3. For rows 4 to 18:
   - Remove `cell.value = ...` and `cell.hyperlink = ...` lines.
   - Retain the style allocations (`cell.font = ...`, `cell.fill = ...`, `cell.alignment = ...`, `cell.border = ...`) so the cells remain styled (alternating green/white fills, thin gray borders, text alignments, wrap settings, and link style structures) but are empty of data.
