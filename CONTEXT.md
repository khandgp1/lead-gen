# Lead Gen — Excel Processing Pipeline

## Project Goal

Process a collection of `.xlsx` spreadsheets by cleaning, reformatting, transforming, and combining them into a single unified output file.

---

## Stack

| Tool | Role |
|---|---|
| **Python 3** | Core language |
| **pandas** | Data manipulation (clean, merge, transform, dedupe, filter) |
| **openpyxl** | Read/write `.xlsx` files |

**Rationale:** Minimal dependencies, no web framework or database required. Industry-standard tooling for this kind of batch spreadsheet work. Handles 10–50 files with ease.

---

## Project Structure

```
lead-gen/
├── input/              # Drop source .xlsx files here before running
├── output/             # Combined result is written here
├── process.py          # Main entry-point script
└── requirements.txt    # pandas, openpyxl
```

---

## How to Run

```bash
python process.py
```

---

## Processing Pipeline (per run)

1. Scan `/input/` for all `.xlsx` files
2. For each file:
   - Load into a pandas DataFrame
   - Normalize column names (handle minor discrepancies across files)
   - Clean data (trim whitespace, fix types/formats)
   - Add or derive any new columns
   - Filter unwanted rows
   - Deduplicate rows
3. Combine all cleaned DataFrames into one
4. Write final result to `/output/combined.xlsx`

---

## Transformation Rules

Transformations are **hard-coded** in `process.py` (no external config file). Rules include:

- **Column normalization:** Reconcile minor naming inconsistencies across source files
- **Cleaning:** Trim whitespace, normalize casing, fix date/phone formats
- **Derived columns:** TBD — to be defined per data review
- **Filtering:** TBD — conditions to be defined per data review
- **Deduplication:** TBD — key column(s) to be defined per data review

---

## Output

- **Format:** Single `.xlsx` file
- **Location:** `/output/combined.xlsx`
- One combined file produced per run (individual cleaned files are not saved separately)
