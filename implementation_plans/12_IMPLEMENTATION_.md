# Batch-Format All Input Lead Spreadsheets to Standard Column Schema

This update extends `format_leads.py` to batch-process **all** `.xlsx` files in the `input/` folder, reformatting each one to match the standard 8-column schema from `01_Roofing_Cold_Outreach_Leads_June2026.xlsx`.

## Background

There are **19 input files** across distinct structural variants:

| Files | Sheet Name | Columns | Header Row | Notes |
|---|---|---|---|---|
| 01–02 | `Ranked Leads` | 8 | Row 3 | Already standard — **skip** |
| 03–04 | `Ranked Leads` | 10 | Row 3 | Has `Rank` + `Website Status`; `Validated` present |
| 05–13 | `Ranked Leads` | 10 | Row 3 | Has `Rank` + `Website Status` + `State License Board`; no `Validated` |
| 14–19 | `Batch ## Leads` | 10 | Row 4 | Extra blank row 3; `State License Board` present; no `Validated` |

### Target Schema (8 columns)
`Business Name` · `Location` · `Website / Online Presence` · `Phone` · `Email` · `Opportunity Reason` · `Sources Verified` · `Validated`

### Column Mapping by Variant

**Files 03–04** (10-col, header row 3, has `Validated`):  
Drop `Rank` (A), `Website Status` (D). Map: B→A, C→B, E→C, F→D, G→E, H→F, I→G, J→H

**Files 05–13** (10-col, header row 3, no `Validated`):  
Drop `Rank` (A), `Website Status` (D), `State License Board` (H). Map: B→A, C→B, E→C, F→D, G→E, I→F, J→G; H=blank

**Files 14–19** (10-col, header row 4, no `Validated`):  
Same as 05–13 but data starts at row 5. Shift all data up one row (drop blank row 3). Rename sheet to `Ranked Leads`.

> [!IMPORTANT]
> Files `01` and `02` already conform to the target schema and must be **skipped** by the script.

> [!NOTE]
> Files 05–19 have `State License Board` which has no equivalent in the target schema — it will be silently dropped. `Validated` (col H) will be left blank for these files.

## Proposed Changes

### Lead Generation Pipeline

#### [MODIFY] [format_leads.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/format_leads.py)
Refactor from single-file to batch processor:
1. Scan `input/` for all `*.xlsx` files using `glob`.
2. Skip files already in the 8-column standard schema.
3. Auto-detect variant by checking column count and header row position.
4. Apply the appropriate column mapping per variant.
5. For files 14–19: rename sheet to `Ranked Leads`, shift rows up by 1 to drop the blank row 3.
6. Apply standard column widths, row heights, and merges (`A1:H1`, `A2:H2`) to all outputs.
7. Save each file in-place with a per-file progress log.

## Verification Plan

### Automated Tests
Extend `verify_leads_format.py` to loop over all files 03–19 and assert:
- 8 columns, header row 3 matches target, merges = `A1:H1`/`A2:H2`, column widths match `01`

```bash
.venv/bin/python format_leads.py
.venv/bin/python verify_leads_format.py
```

## Checklist

- `[x]` Refactor `format_leads.py` into a batch loop over all `input/*.xlsx` files
- `[x]` Auto-detect variant (skip 01-02, map variant A for 03-04, B for 05-13, C for 14-19)
- `[x]` Handle variant A (03–04): drop Rank + Website Status
- `[x]` Handle variant B (05–13): drop Rank + Website Status + State License Board; blank Validated
- `[x]` Handle variant C (14–19): same as B + shift data up one row; rename sheet
- `[/]` Update `verify_leads_format.py` to validate all non-standard files
- `[ ]` Run and verify all files pass
