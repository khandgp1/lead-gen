# Remove Explicit Row Heights from Output Spreadsheet

## Background

`process.py` currently sets every data row in `output/Roofing_Outreach_Leads.xlsx` to an explicit height of `72.0` pts:

```python
ws_out.row_dimensions[r].height = 72.0
```

The user wants data rows to use Excel's **default row height** instead, so the output looks clean and consistent without forcing a fixed tall row size.

## Proposed Changes

### Lead Generation Pipeline

#### [MODIFY] [process.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/process.py)

Remove the single line on line 187:
```python
ws_out.row_dimensions[r].height = 72.0
```

No other changes needed. The header row heights (row 1 = 30.0, row 2 = 16.0, row 3 = 27.75) are intentional and stay.

## Verification Plan

### Automated Tests
```bash
.venv/bin/python process.py
```
Open `output/Roofing_Outreach_Leads.xlsx` and confirm data rows (row 4+) have default Excel height.

## Checklist

- `[x]` Remove `ws_out.row_dimensions[r].height = 72.0` from the data row loop in `process.py`
- `[x]` Re-run `process.py` and verify output
