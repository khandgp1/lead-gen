# Email Sequence #2 — HTML Viewer Generator

Generate a new `output/sequence_2_emails.html` file with personalized Email #2 bodies for all
leads, matching the exact UI/UX of the existing `output/outreach_emails.html`.

---

## Background

Email #1 (`generate_emails.py` → `outreach_emails.html`) used the hook:
> "I noticed your {business_name} Facebook page doesn't link to a website."

Email #2 is a follow-up with a tighter value prop, referencing the city and web-presence platform:

**Subject:**
> Website Demo for {business_name}

**Body:**
> Hi,
>
> The {business_name} {web_presence} page is missing a link to a website.
>
> Last month there were thousands of Google searches for roofers in {city} alone.
>
> Want a website demo for {business_name}?

---

## Data Mapping (VALID_Roofing_Outreach_Leads.xlsx → "Leads" sheet)

| Variable | Column | Transform |
|---|---|---|
| `{business_name}` | A | as-is |
| `{city}` | B | text before the first comma, stripped |
| `{web_presence}` | D | capitalize first letter only (FACEBOOK → Facebook) |
| `{email}` | F | as-is |
| `{website_url}` | C | shown in card (for reference), as-is |

Rows start at row **3** (rows 1–2 are blank). Skip rows missing name or email.

---

## Proposed Changes

### [NEW] generate_emails_2.py

A standalone script that:
1. Reads `manual/VALID_Roofing_Outreach_Leads.xlsx`
2. Extracts cols A, B (city = text before comma), C, D (title-cased), F
3. Applies the Email #2 subject + body templates
4. Builds identical HTML card UI as `generate_emails.py`
5. Writes to `output/sequence_2_emails.html`

Key differences from `generate_emails.py`:
- New `SUBJECT_TEMPLATE` and `BODY_TEMPLATE` (Email #2 copy)
- Reads column **B** for city (split on `,`, take `[0].strip()`)
- Reads column **D** for web presence (`.capitalize()`)
- Output file name is `sequence_2_emails.html`
- Page title / header badge updated to "Sequence 2"

---

## Checklist

- [x] Create `generate_emails_2.py`
  - [x] Data extraction: cols A, B (city), C (website), D (web presence), F (email)
  - [x] City transform: split on `,`, take first part, strip whitespace
  - [x] Web presence transform: `.capitalize()` (FACEBOOK → Facebook)
  - [x] Email #2 subject template
  - [x] Email #2 body template
  - [x] HTML builder (copy design, update title/header to "Sequence 2")
  - [x] Write to `output/sequence_2_emails.html`
- [x] Run script and verify output
- [x] Spot-check 3–5 cards for correct city, web presence, business name

---

## Verification Plan

### Automated
```bash
cd lead-gen && source .venv/bin/activate && python3 generate_emails_2.py
```
Expected: `✅  Output written to: output/sequence_2_emails.html`

### Manual
- Open `output/sequence_2_emails.html` in a browser
- Confirm card #1 (Wilkins Roofing, Joplin, Facebook) reads correctly
- Confirm a NEXTDOOR lead shows "Nextdoor" (not "NEXTDOOR")
- Test "Copy All" button pastes clean plain text
- Confirm search/filter still works
