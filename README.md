# Google Docs → Gmail Newsletter (Apps Script)

Send a nicely formatted newsletter by **copy/pasting from Google Docs into a Gmail Draft**, then use **Google Apps Script** to personalize and send to a subscriber list in Google Sheets — while preserving **inline embedded images** (not attachments), and supporting **one-click unsubscribe**.

## What this solves

- ✅ Keep authoring in **Google Docs** (tables, images, spacing, formatting)
- ✅ Paste into **Gmail Draft** (WYSIWYG)
- ✅ Personalize per recipient (`{{first_name}}`, unsubscribe link)
- ✅ Preserve **inline embedded images** by extracting `cid:` images via Gmail API + sending using `inlineImages`
- ✅ One-click unsubscribe with Apps Script Web App endpoint

---

## Requirements

- Google Sheet with a tab named `Subscribers` (or update `SHEET_NAME`)
- Google Apps Script project attached to the sheet
- Gmail Draft created by pasting the formatted content from Google Docs
- **Advanced Gmail Service enabled** in Apps Script (Gmail API)

---

## Sheet schema (no change)

Required columns:

- `email`
- `name`
- `first_name`
- `status`
- `token`

Recommended optional column:

- `unsubscribed_at` (timestamp when user unsubscribed)

Example header row:
