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

email | name | first_name | status | token | unsubscribed_at

Notes:
- `status` should be `subscribed` or `unsubscribed`
- `token` must be present (script can generate tokens for missing rows)

---

## Setup

### 1) Create your Gmail Draft (template)
1. Write your newsletter in **Google Docs**
2. Copy/paste into a new **Gmail Draft**
3. In the draft body, include:
   - `{{first_name}}`
   - Either:
     - `{{unsub_link}}` (preferred), OR
     - a placeholder URL like `https://example.com/unsub` (if Gmail/Docs link tooling is annoying)

The draft subject becomes the subject used when sending.

### 2) Apps Script
1. Open the Google Sheet
2. Extensions → Apps Script
3. Paste the code into `Code.gs`

### 3) Enable Gmail API (Advanced Gmail Service)
In Apps Script Editor:
- **Services** (left sidebar) → Add a service → **Gmail API** → Add

If prompted in Google Cloud console, also ensure the Gmail API is enabled there.

### 4) Get your Draft ID
Run:

- `PRINT_DRAFT_IDS()`

Copy the correct ID and set:

- js
- const DRAFT_ID = "r-...";

### 5) Deploy Web App (for unsubscribe URL)

Deploy → New deployment → Type: Web app
	•	Execute as: Me
	•	Who has access: Anyone (or Anyone with link)

This enables:

ScriptApp.getService().getUrl()

### 6) Test send

Set:

const TEST_TARGET_EMAIL = "you@domain.com";

Run:
	•	sendNewsletterToOne()

### 7) Send to all

Run:
	•	sendNewsletter()

⸻

### Unsubscribe behavior

Each send creates a link like:

<webapp_url>?t=<token>

When someone clicks it:
	•	status becomes unsubscribed
	•	unsubscribed_at is set if the column exists

⸻

### Editing / personalization

The script replaces:
	•	{{first_name}}
	•	{{unsub_link}} OR PLACEHOLDER_UNSUB_URL

It also hardens inline cid images for better layout behavior inside tables:
	•	forces display:block; width:100%; height:auto; ...

⸻

### Common gotchas
	•	Inline images missing
	•	Ensure you pasted into Gmail so images are truly embedded in the draft body.
	•	Ensure Gmail API is enabled and Advanced Gmail Service is on.
	•	Some email clients ignore table auto-fill behavior
	•	Gmail draft composer may visually “fill” table cells in ways certain clients won’t replicate.
	•	This is why the script adds conservative inline styles to cid: images.

