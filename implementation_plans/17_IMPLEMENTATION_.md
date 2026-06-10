# Add Website Column to Email Cards

## Background

The HTML email viewer (`output/outreach_emails.html`) currently shows three fields per card:
- **To** (email address)
- **Subject** (merged template line)
- **Body** (merged template body)

The user wants the **Website / Online Presence** value from column C of the spreadsheet also shown on each card as additional context. It should display as plain text (raw cell value, no link rendering), always — regardless of whether it's a URL or text like "No website found".

---

## Proposed Changes

### Lead Generation Pipeline

#### [MODIFY] [generate_emails.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/generate_emails.py)

**1. Read column C in `main()`**

Add `raw_website = ws.cell(row=r, column=3).value` alongside the existing name/email reads, and include `"website"` in the lead dict:

```python
raw_website = ws.cell(row=r, column=3).value
website = str(raw_website).strip() if raw_website else "—"
leads.append({"name": name, "email": email, "website": website, "subject": subject, "body": body})
```

**2. Render it in `build_html()` / card template**

Add a new field row between `To` and `Subject`, displayed with the same `label` / `value` styling:

```html
<div class="field-row">
  <span class="label">Website</span>
  <span class="value">{website}</span>
</div>
```

No changes to CSS needed — reuses existing `.label` / `.value` styles.

---

## Verification Plan

### Automated
```bash
source .venv/bin/activate
python generate_emails.py
```
- Confirm script exits cleanly, still shows 104 leads
- Spot-check: row with "No website found" → shows that text on card
- Spot-check: row with a Facebook URL → shows the raw URL string on card

### Manual
- Open `output/outreach_emails.html` in browser
- Verify each card now shows a **Website** field row below **To**
- Confirm plain-text rendering (no link tags, no truncation)

---

## Checklist

- `[x]` Read column C (`Website / Online Presence`) in `main()` and include in lead dict
- `[x]` Add `"Website"` field row to card HTML in `build_html()` (plain text, existing styles)
- `[x]` Re-run `python generate_emails.py` and confirm 104 leads, no errors
- `[ ]` Manual browser check: website field visible and renders as plain text on all card types
