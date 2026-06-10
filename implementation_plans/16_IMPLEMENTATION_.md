# Generate Outreach Emails from VALID_Roofing_Outreach_Leads.xlsx

## Background

We have 104 verified roofing leads in `manual/VALID_Roofing_Outreach_Leads.xlsx` (sheet: `Leads`, data starts row 3). Each row has:

| Col | Field |
|---|---|
| A | Business Name |
| B | Location |
| C | Website / Online Presence |
| D | Phone |
| E | Email |
| F | VALID status |

We want to merge each lead's `Business Name` and `Email` into a pre-written cold outreach email template, then render every email as a card in a single HTML file the user can open in a browser, read, and copy with one click.

---

## Email Template

**Subject:**
```
Simple Website Demo for {{business_name}} Facebook Page
```

**Body:**
```
Hi,

I noticed your {{business_name}} Facebook page doesn't link to a website.

Potential customers search Google for a roofing business near them, and those that find a Facebook page commonly look for a website. When they don't find one, they go back to Google, find a business with a website, and use it to contact the next guy.

Want to see a simple website demo for {{business_name}}?
```

The only merge field is `{{business_name}}` → column A value.

---

## Proposed Changes

### Lead Generation Pipeline

#### [NEW] [generate_emails.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/generate_emails.py)

Standalone Python script that:

1. Opens `manual/VALID_Roofing_Outreach_Leads.xlsx`, sheet `Leads`
2. Iterates rows 3–106, skipping any row with no Business Name or no Email
3. For each lead, merges `{{business_name}}` into the subject + body template
4. Writes `output/outreach_emails.html` — a self-contained HTML file (no external dependencies, all CSS/JS inline)

**HTML output design:**
- Dark-mode card layout with a search/filter bar at the top
- Each card shows:
  - Lead number badge + Business Name (large)
  - `To:` email address
  - `Subject:` line (pre-filled)
  - Email body text block
  - **One "Copy All" button** — clicking it copies a formatted string containing the `To:` address, `Subject:` line, and body to the clipboard
  - Button shows "Copied ✓" feedback for 2 seconds after click
- Total count shown in the header (e.g. "104 emails")
- Card counter visible on each card (e.g. "1 of 104")

---

## Verification Plan

### Automated
```bash
source .venv/bin/activate
python generate_emails.py
```
- Confirm `output/outreach_emails.html` is created
- Confirm lead count in terminal output matches 104
- Confirm no crash on rows with `None` location or unusual characters

### Manual
- Open `output/outreach_emails.html` in a browser
- Verify `{{business_name}}` is fully replaced on every visible card
- Click "Copy All" on a card and paste into an email client — confirm To, Subject, and body are all present and formatted correctly

---

## Checklist

- `[x]` Create `generate_emails.py`
  - `[x]` Read `manual/VALID_Roofing_Outreach_Leads.xlsx`, sheet `Leads`, rows 3+
  - `[x]` Skip rows missing Business Name or Email
  - `[x]` Merge `{{business_name}}` into subject and body
  - `[x]` Build HTML string with all cards
  - `[x]` Write to `output/outreach_emails.html`
- `[x]` HTML card design
  - `[x]` Dark-mode layout, card grid
  - `[x]` To / Subject / Body fields visible per card
  - `[x]` "Copy All" button with clipboard API + "Copied ✓" feedback
  - `[x]` Lead number badge + total count header
- `[x]` Run script and verify output file (104 leads, 0 skipped)
- `[ ]` Manual browser check: copy-paste test into email client
