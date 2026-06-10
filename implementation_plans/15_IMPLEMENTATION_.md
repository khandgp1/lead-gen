# Update Email Column (E) Alignment in Output Spreadsheet

## Background

In `process.py`, column C (`Website / Online Presence`) already has wrap text disabled and shrink-to-fit enabled:

```python
wrap_data = False if col_letter == 'C' else True
shrink_data = True if col_letter == 'C' else None
```

The user wants the same treatment applied to column **E** (`Email`) — no wrap text, shrink to fit.

## Proposed Changes

### Lead Generation Pipeline

#### [MODIFY] [process.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/process.py)

Update the two alignment lines in the data row loop to also handle column `E`:

```python
# Before
wrap_data  = False if col_letter == 'C' else True
shrink_data = True if col_letter == 'C' else None

# After
wrap_data  = False if col_letter in ('C', 'E') else True
shrink_data = True if col_letter in ('C', 'E') else None
```

Also update the **column header row 3** alignment for `E` to match (currently `wrap_text=True`):

```python
# Before
wrap_hdr  = False if col == 'C' else True
shrink_hdr = True if col == 'C' else None

# After
wrap_hdr  = False if col in ('C', 'E') else True
shrink_hdr = True if col in ('C', 'E') else None
```

## Verification Plan

### Automated Tests
```bash
.venv/bin/python process.py
```
Open `output/Roofing_Outreach_Leads.xlsx` and confirm column E cells (header + data) have wrap text off and shrink-to-fit on.

## Checklist

- `[ ]` Update data row alignment logic: extend `'C'` check to `('C', 'E')` for `wrap_data` and `shrink_data`
- `[ ]` Update header row 3 alignment logic: extend `'C'` check to `('C', 'E')` for `wrap_hdr` and `shrink_hdr`
- `[ ]` Re-run `process.py` and verify output
