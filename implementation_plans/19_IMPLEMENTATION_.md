# Include Website in Clipboard Payload (Sequence 2)

Add the `Website` field to the copy/paste payload in `sequence_2_emails.html`.

---

## Proposed Changes

### [MODIFY] [generate_emails_2.py](file:///Users/khandpv1/Desktop/.AntiGrav/lead-gen/generate_emails_2.py)

Modify the HTML payload construction to include the `Website` link under the recipient email address.

```python
# Before
copy_text = html_escape(
    f"To: {lead['email']}\nSubject: {lead['subject']}\n\n{lead['body']}"
)

# After
copy_text = html_escape(
    f"To: {lead['email']}\nWebsite: {lead['website']}\nSubject: {lead['subject']}\n\n{lead['body']}"
)
```

---

## Checklist

- [x] Modify `generate_emails_2.py` clipboard text format
- [x] Run `generate_emails_2.py` to update the output
- [x] Open `output/sequence_2_emails.html` in browser and test clicking "Copy All" to verify "Website: <url>" is copied alongside To, Subject, and Body

---

## Verification Plan

### Automated
```bash
python3 generate_emails_2.py
```

### Manual
- Click the "Copy All" button on card #1.
- Paste clipboard contents in a scratch space.
- Confirm it looks like:
  ```text
  To: wilkinshomeexteriors@gmail.com
  Website: https://www.facebook.com/WilkinsHomeExteriors/
  Subject: Website Demo for Wilkins Roofing

  Hi,
  ...
  ```
