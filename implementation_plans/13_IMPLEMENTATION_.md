# Batch-Process All Input Files in process.py

This update extends `process.py` to loop over **all** `.xlsx` files in the `input/` folder (rather than a single hardcoded file), extracting leads with valid emails from each, and consolidating them all into a single output file: `output/Roofing_Outreach_Leads.xlsx`.

## Background

### Current Behavior
`process.py` currently:
- Hard-codes a single input file: `01_Roofing_Cold_Outreach_Leads_June2026.xlsx`
- Reads the `Ranked Leads` sheet
- Filters rows by valid email (skips `None`, empty, `"No email found"`)
- Extracts columns: Business Name (A), Location (B), Website (C), Phone (D), Email (E), Sources Verified (G)
- Writes a styled output spreadsheet to `output/Roofing_Outreach_Leads.xlsx`

### After This Change
- Scan `input/` for all `*.xlsx` files (sorted)
- Skip the `.gitkeep` file
- Read from each file's `Ranked Leads` sheet (all files are now standardized to this name after `format_leads.py` runs)
- Apply the same email filter and column extraction per file
- Accumulate all valid leads from all files into a single consolidated output

### Key Details
- The **column mapping stays the same** — all input files now share the same 8-column schema (after `format_leads.py` runs first)
- The output file structure (title row, subtitle row, column headers, styling) remains unchanged
- The output tab is `Leads`, the file is `Roofing_Outreach_Leads.xlsx`
- A per-file and grand-total summary should be printed to the console

## Proposed Changes

### Lead Generation Pipeline

#### [MODIFY] [process.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/process.py)

1. Replace the hardcoded `filename` with a `glob` scan of `input/*.xlsx`, sorted.
2. Wrap the extraction logic in a loop over all files.
3. Accumulate leads into a single list across all files.
4. Print per-file extraction counts and a final total.
5. The output writing logic (styling, headers, data rows) remains unchanged.

> [!NOTE]
> `format_leads.py` must be run first to standardize all input files. `process.py` assumes all files use the standard 8-column schema on the `Ranked Leads` sheet. If a file cannot be opened or has no `Ranked Leads` sheet, log a warning and skip it.

## Verification Plan

### Automated Tests
```bash
.venv/bin/python format_leads.py   # ensure all input files are standardized first
.venv/bin/python process.py        # run batch processor
```
Verify:
- Console output shows per-file lead counts and a combined total
- `output/Roofing_Outreach_Leads.xlsx` exists and contains all consolidated leads
- No duplicate or missing rows

## Checklist

- `[x]` Replace hardcoded filename with `glob` scan of `input/*.xlsx`
- `[x]` Wrap extraction in a loop; accumulate all leads into one list
- `[x]` Add per-file logging and a final total summary
- `[x]` Handle missing/unreadable files gracefully (warn and skip)
- `[x]` Run and verify consolidated output is correct
