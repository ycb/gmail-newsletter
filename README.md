# Google Docs → Gmail Newsletter (Apps Script)

Send a nicely formatted newsletter by **copy/pasting from Google Docs into a Gmail Draft**, then use **Google Apps Script** to personalize and send to a subscriber list in Google Sheets — while preserving **inline embedded images**, supporting **one-click unsubscribe**, and logging **open events** to a sheet.

---

## What this solves

- ✅ Author in **Google Docs** (tables, images, spacing, formatting)
- ✅ Paste into a **Gmail Draft** (WYSIWYG)
- ✅ Personalize per recipient (`{{first_name}}`, unsubscribe link)
- ✅ Preserve **inline embedded images** by extracting `cid:` images via Gmail API and sending via `inlineImages`
- ✅ One-click unsubscribe via Apps Script Web App endpoint
- ✅ Open tracking (pixel) → logs to an `Events` sheet

---

## Requirements

- Google Sheet with a tab named `Subscribers` (or update `SHEET_NAME`)
- Google Apps Script project (typically bound to the sheet)
- Gmail Draft created by pasting formatted content from Google Docs
- **Advanced Gmail Service enabled** in Apps Script (Gmail API)
- Web App deployment (used for unsubscribe + open tracking)

---

## Sheet schema

### Subscribers tab

Required columns:

- `email`
- `first_name`
- `status`
- `token`

Optional:

- `unsubscribed_at` (timestamp)

Example header row:

email | first_name | status | token | unsubscribed_at

Notes:
- `status` should be `subscribed` or `unsubscribed`
- `token` must be present (the script can generate tokens for missing rows)

### Events tab

No strict header requirement, but recommended columns:

ts | event | campaign_id | token | email | url

---

## Setup

### 1) Create your Gmail Draft (template)

1. Write your newsletter in **Google Docs**
2. Copy/paste into a new **Gmail Draft**
3. In the draft body, include:
   - `{{first_name}}`
   - `{{unsub_link}}` (preferred)

If Gmail/Docs link tooling is annoying, you can also include a placeholder URL and have the script replace it (see `PLACEHOLDER_UNSUB_URL`).

The draft subject becomes the subject used when sending.

---

### 2) Apps Script

1. Open the Google Sheet
2. Extensions → Apps Script
3. Paste the code into `Code.gs`
4. Update these constants:
   - `SPREADSHEET_ID`
   - `DRAFT_ID`
   - `WEB_APP_URL` (must be the deployed `/exec` URL)
   - `DEFAULT_CAMPAIGN_ID` (optional)

---

### 3) Enable Gmail API (Advanced Gmail Service)

In Apps Script Editor:
- Services → Add a service → **Gmail API** → Add

(If prompted in Google Cloud console, also ensure the Gmail API is enabled there.)

---

### 4) Get your Draft ID

Run:

PRINT_DRAFT_IDS()

Copy the correct ID and set:

const DRAFT_ID = "r-...";

---

### 5) Deploy Web App (unsubscribe + open tracking)

Deploy → New deployment → Type: Web app

- Execute as: Me  
- Who has access: Anyone (or Anyone with link)

Copy the deployed **/exec** URL and set:

const WEB_APP_URL = "https://script.google.com/macros/s/.../exec";

---

## Test

Set:

const TEST_TARGET_EMAIL = "you@domain.com";

Run:

sendNewsletterToOne()

Verify:
- Email renders correctly (including inline images)
- Unsubscribe link works
- Opens create rows in the `Events` sheet

---

## Send to all

Run:

sendNewsletter()

---

## Unsubscribe behavior

Each send injects an unsubscribe link like:

<WEB_APP_URL>?t=<token>

When a subscriber clicks:
- `status` becomes `unsubscribed`
- `unsubscribed_at` is set if the column exists

---

## Open tracking

Each email includes a 1×1 SVG pixel:

<WEB_APP_URL>?mode=track_open&t=<token>&cid=<campaign_id>

When the email client loads images, an `open` event is appended to the `Events` sheet.

Note: Open tracking is approximate. Some clients block images by default; others may prefetch.

---

## Common gotchas

- **Inline images missing**
  - Ensure images are truly embedded in the Gmail Draft (not attachments).
  - Confirm Gmail API is enabled and Advanced Gmail Service is on.
- **Layout differences between Gmail and other clients**
  - Gmail draft composer may render tables/styles differently than Apple Mail or Outlook.
  - The script adds conservative inline styling to `cid:` images to reduce breakage.
