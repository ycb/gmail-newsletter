/**
 * Newsletter sender + one-click unsubscribe (Gmail Draft template WITH inline embedded images)
 *
 * Sheet columns required:
 * email, name, first_name, status, token
 *
 * Optional column:
 * unsubscribed_at
 *
 * Usage:
 * 1) Create a Gmail DRAFT by pasting from Google Docs (keeps inline images).
 *    - Put placeholders in the draft body:
 *        {{first_name}}
 *        {{unsub_link}}   OR use a placeholder URL (e.g. https://example.com/unsub)
 * 2) Run PRINT_DRAFT_IDS() once to find your draft's ID
 * 3) Paste that ID into DRAFT_ID below
 * 4) Deploy as Web App (doGet) so ScriptApp.getService().getUrl() works
 * 5) Run sendNewsletterToOne() to test
 * 6) Run sendNewsletter() to send
 *
 * Requirements:
 * - Advanced Gmail Service enabled in Apps Script (Gmail API)
 *   Apps Script Editor → Services → add Gmail API
 */

const SHEET_NAME = "Subscribers";

// ====== CONFIG (safe to commit) ======

// Display name for the sender shown in recipients' inboxes
const FROM_NAME = "Your Name";

// Set this after running PRINT_DRAFT_IDS()
const DRAFT_ID = "r-REPLACE_WITH_YOUR_DRAFT_ID";

// If your draft can't contain {{unsub_link}} as a real link, put a real placeholder URL
// in the draft (e.g. https://example.com/unsub) and set this to that exact value.
// If you DO use {{unsub_link}}, you can leave this blank.
const PLACEHOLDER_UNSUB_URL = "https://example.com/unsub";

// Test helper (set to your own email when testing)
const TEST_TARGET_EMAIL = "you@example.com";

// =====================================

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Missing sheet tab: "${SHEET_NAME}"`);
  return sh;
}

function getHeaderMap_(headers) {
  const m = {};
  headers.forEach((h, i) => (m[String(h || "").trim()] = i));
  return m;
}

function getWebAppUrl_() {
  return ScriptApp.getService().getUrl();
}

/**
 * Run once to find your draft's ID.
 */
function PRINT_DRAFT_IDS() {
  const drafts = GmailApp.getDrafts();
  Logger.log(`Found ${drafts.length} drafts.`);
  drafts.forEach((d, i) => {
    const msg = d.getMessage();
    Logger.log(`${i + 1}. ID=${d.getId()} | Subject="${msg.getSubject()}" | Updated=${msg.getDate()}`);
  });
  Logger.log("Copy the correct ID and set DRAFT_ID at the top of the script.");
}

function getDraftById_() {
  if (!DRAFT_ID || !String(DRAFT_ID).trim() || String(DRAFT_ID).includes("REPLACE_WITH_YOUR_DRAFT_ID")) {
    throw new Error("DRAFT_ID is not set. Run PRINT_DRAFT_IDS() and paste the correct ID into DRAFT_ID.");
  }
  return GmailApp.getDraft(String(DRAFT_ID).trim());
}

function assertHasPlaceholders_(html) {
  const missing = [];
  if (!html || !html.trim()) missing.push("(empty body)");
  if (html && !html.includes("{{first_name}}")) missing.push("{{first_name}}");

  const hasUnsubToken = html && html.includes("{{unsub_link}}");
  const hasPlaceholderUrl = !!PLACEHOLDER_UNSUB_URL && html && html.includes(PLACEHOLDER_UNSUB_URL);
  if (!hasUnsubToken && !hasPlaceholderUrl) {
    missing.push("{{unsub_link}} (or set PLACEHOLDER_UNSUB_URL to a URL present in the draft)");
  }

  if (missing.length) throw new Error("Draft template missing required placeholders: " + missing.join(", "));
}

function escapeHtml_(s) {
  return String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function applyTemplate_(templateHtml, firstName, unsubLink) {
  let html = templateHtml.replaceAll("{{first_name}}", escapeHtml_(firstName));

  if (templateHtml.includes("{{unsub_link}}")) {
    html = html.replaceAll("{{unsub_link}}", unsubLink);
  } else if (PLACEHOLDER_UNSUB_URL && templateHtml.includes(PLACEHOLDER_UNSUB_URL)) {
    html = html.replaceAll(PLACEHOLDER_UNSUB_URL, unsubLink);
  }

  // Email-hardening: make cid images behave better in table layouts
  html = hardenCidImagesForEmail_(html);

  return html;
}

/**
 * Adds/overrides inline styles for cid: images so they fill table cells more reliably.
 */
function hardenCidImagesForEmail_(html) {
  if (!html) return html;

  // Only touch <img> tags that reference cid:
  return html.replace(/<img\b([^>]*\bsrc=["']cid:[^"']+["'][^>]*)>/gi, (full, attrs) => {
    const hasStyle = /\bstyle=["'][^"']*["']/.test(attrs);

    const desired = "display:block;width:100%;height:auto;max-width:100%;border:0;";

    if (hasStyle) {
      // Append our desired styles (safe-ish; last wins for duplicates)
      attrs = attrs.replace(/\bstyle=["']([^"']*)["']/, (m, s) => {
        const merged = (s || "").trim();
        const sep = merged && !merged.endsWith(";") ? ";" : "";
        return `style="${merged}${sep}${desired}"`;
      });
    } else {
      attrs += ` style="${desired}"`;
    }

    // Also set width attribute if absent (helps some clients)
    if (!/\bwidth=/.test(attrs)) attrs += ` width="100%"`;

    return `<img${attrs}>`;
  });
}

/**
 * ============================
 * Draft MIME extraction (Gmail API)
 * ============================
 *
 * Requires Advanced Gmail Service enabled:
 * Apps Script Editor → Services → add Gmail API
 */
function getDraftTemplate_() {
  const draft = getDraftById_();
  const msg = draft.getMessage();

  const subject = msg.getSubject();
  const msgId = msg.getId(); // message id used with Gmail API

  // Fetch full message payload (contains multipart + attachments ids)
  const full = Gmail.Users.Messages.get("me", msgId, { format: "full" });
  const payload = full.payload;

  const html = extractHtmlFromPayload_(payload);
  if (!html) throw new Error("Could not find an HTML part in the draft message.");
  assertHasPlaceholders_(html);

  const inlineImages = extractInlineImages_(payload, msgId);

  // NOTE: If the email has cid: references but we didn't find any inline images,
  // something is off (draft not actually storing cid attachments).
  if ((html.match(/src=["']cid:/gi) || []).length > 0 && Object.keys(inlineImages).length === 0) {
    Logger.log("WARNING: HTML contains cid: images but no inline images were extracted from payload.");
  }

  return { subject, html, inlineImages };
}

function extractHtmlFromPayload_(payload) {
  if (!payload) return "";

  // Sometimes HTML is directly in payload.body.data
  if (payload.mimeType === "text/html" && payload.body && payload.body.data) {
    return decodeBodyToString_(payload.body.data);
  }

  // Otherwise traverse parts
  const parts = payload.parts || [];
  for (const p of parts) {
    const found = extractHtmlFromPayload_(p);
    if (found) return found;
  }

  return "";
}

function extractInlineImages_(payload, msgId) {
  const inline = {};
  walkParts_(payload, (part) => {
    if (!part) return;

    const mime = part.mimeType || "";
    const isImage = mime.startsWith("image/");
    const filename = part.filename || "";
    const hasAttachmentId = part.body && part.body.attachmentId;

    if (!isImage || !hasAttachmentId) return;

    // Extract Content-ID header so we can map to cid: references
    const contentId = getHeader_(part.headers, "Content-ID"); // usually like "<ii_abc123>" or "<ii_abc123@mail.gmail.com>"
    const xAttachId = getHeader_(part.headers, "X-Attachment-Id"); // sometimes matches cid: too

    // If neither header exists, we can't reliably map this image to cid: references.
    if (!contentId && !xAttachId) return;

    const attachment = Gmail.Users.Messages.Attachments.get("me", msgId, part.body.attachmentId);
    const bytes = decodeBase64UrlToBytes_(attachment.data);

    const blob = Utilities.newBlob(bytes, mime, filename || "inline");

    // Build keys that might match HTML's cid: reference
    const keys = [];
    if (contentId) keys.push(...normalizeContentIdVariants_(contentId));
    if (xAttachId) keys.push(String(xAttachId).trim());

    // Store the blob under all variants (safe: same blob)
    keys
      .map((k) => String(k || "").trim())
      .filter((k) => k)
      .forEach((k) => {
        inline[k] = blob;
      });
  });

  return inline;
}

function walkParts_(part, fn) {
  fn(part);
  const parts = part && part.parts ? part.parts : [];
  for (const p of parts) walkParts_(p, fn);
}

function getHeader_(headers, name) {
  if (!headers || !headers.length) return "";
  const target = String(name || "").toLowerCase();
  for (const h of headers) {
    if (String(h.name || "").toLowerCase() === target) return String(h.value || "");
  }
  return "";
}

/**
 * Gmail can emit Content-ID values like:
 *  "<ii_abc123>"
 *  "<ii_abc123@mail.gmail.com>"
 * And HTML may reference either:
 *  cid:ii_abc123
 *  cid:ii_abc123@mail.gmail.com
 *
 * Return multiple plausible keys so cid mapping doesn't fail.
 */
function normalizeContentIdVariants_(cid) {
  const raw = String(cid || "").trim().replace(/^<|>$/g, "");
  if (!raw) return [];
  const variants = [raw];

  // If there's a domain suffix, add local-part too.
  const at = raw.indexOf("@");
  if (at > 0) variants.push(raw.slice(0, at));

  // Sometimes Gmail inserts angle brackets in HTML (rare), so include bracketed too.
  variants.push(`<${raw}>`);

  return Array.from(new Set(variants));
}

/**
 * Decode Gmail body data -> string, preserving UTF-8 punctuation.
 * Falls back to ISO-8859-1 only if UTF-8 produced lots of replacement chars.
 */
function decodeBodyToString_(data) {
  const bytes = decodeBase64UrlToBytes_(data);

  // First try UTF-8 (correct for Gmail/Docs paste)
  const sUtf8 = Utilities.newBlob(bytes).getDataAsString("UTF-8");
  const badCount = (sUtf8.match(/\uFFFD/g) || []).length;
  if (badCount > 5) {
    return Utilities.newBlob(bytes).getDataAsString("ISO-8859-1");
  }
  return sUtf8;
}

/**
 * Robust decoder for Gmail API payload fields:
 * - Accepts base64url/base64 strings
 * - Accepts byte arrays (Apps Script sometimes surfaces as [..] or "1,2,3" including NEGATIVES)
 * - Accepts raw HTML strings that already start with "<"
 *
 * CRITICAL FIX:
 * - Handle NEGATIVE byte values (Java signed bytes). Without this, smart quotes become �.
 */
function decodeBase64UrlToBytes_(data) {
  if (data == null) return [];

  // If it's already an array-like of numbers, normalize and return.
  if (Array.isArray(data)) {
    return data
      .map((n) => Number(n))
      .filter((n) => Number.isFinite(n))
      .map((n) => ((n % 256) + 256) % 256);
  }

  let s = String(data).trim();

  // Trim outer quotes if present
  if ((s.startsWith('"') && s.endsWith('"')) || (s.startsWith("'") && s.endsWith("'"))) {
    s = s.slice(1, -1).trim();
  }

  // Case 0: Already plain HTML/text (starts like "<div" etc). Return UTF-8 bytes.
  if (s.startsWith("<")) {
    return Utilities.newBlob(s, "text/plain", "raw").getBytes();
  }

  // Case 1: Byte list in messy forms, including NEGATIVES:
  //   "60,100,105,118,32,..." OR "[-30,-128,-103,...]"
  // IMPORTANT: keep the minus sign when parsing.
  if (s.includes(",") && /-?\d/.test(s)) {
    const nums = s.match(/-?\d{1,4}/g); // allow negatives and up to 4 digits just in case
    if (nums && nums.length > 10) {
      const bytes = nums
        .map((x) => parseInt(x, 10))
        .filter((n) => Number.isFinite(n))
        .map((n) => ((n % 256) + 256) % 256) // convert signed -> unsigned
        .filter((n) => n >= 0 && n <= 255);
      if (bytes.length) return bytes;
    }
  }

  // Case 2: base64url/base64
  let b64 = s.replace(/-/g, "+").replace(/_/g, "/").replace(/\s/g, "");

  // pad to multiple of 4
  const pad = b64.length % 4;
  if (pad) b64 += "====".slice(pad);

  try {
    return Utilities.base64Decode(b64);
  } catch (e) {
    Logger.log(`Base64 decode failed. len=${b64.length} head="${b64.slice(0, 80)}"`);
    throw new Error("Could not decode string (base64/base64url). " + e);
  }
}

/**
 * Sanity check: show who would receive the email without sending.
 */
function DRY_RUN_previewRecipients() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = getHeaderMap_(headers);

  ["email", "first_name", "status", "token"].forEach((col) => {
    if (idx[col] === undefined) throw new Error(`Missing column: ${col}`);
  });

  const recipients = [];
  for (let r = 1; r < values.length; r++) {
    const email = String(values[r][idx.email] || "").trim().toLowerCase();
    const status = String(values[r][idx.status] || "").trim().toLowerCase();
    const token = String(values[r][idx.token] || "").trim();
    const firstName = String(values[r][idx.first_name] || "").trim();

    if (!email || !token) continue;
    if (status === "unsubscribed") continue;

    recipients.push({ email, firstName: firstName || "there" });
  }

  Logger.log(`Would send to ${recipients.length} recipients.`);
  Logger.log(JSON.stringify(recipients.slice(0, 20), null, 2));
}

/**
 * Send the newsletter to all subscribers (one email per person).
 * Draft controls the subject line + inline images.
 */
function sendNewsletter() {
  const tmpl = getDraftTemplate_(); // {subject, html, inlineImages}

  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = getHeaderMap_(headers);

  ["email", "first_name", "status", "token"].forEach((col) => {
    if (idx[col] === undefined) throw new Error(`Missing column: ${col}`);
  });

  const webAppUrl = getWebAppUrl_();
  if (!webAppUrl) throw new Error("Web App URL missing. Deploy as Web App first.");

  let sent = 0;
  let skipped = 0;

  for (let r = 1; r < values.length; r++) {
    const email = String(values[r][idx.email] || "").trim().toLowerCase();
    const status = String(values[r][idx.status] || "").trim().toLowerCase();
    const token = String(values[r][idx.token] || "").trim();
    const firstName = String(values[r][idx.first_name] || "").trim() || "there";

    if (!email || !token) {
      skipped++;
      continue;
    }
    if (status === "unsubscribed") {
      skipped++;
      continue;
    }

    const unsubLink = `${webAppUrl}?t=${encodeURIComponent(token)}`;
    const htmlBody = applyTemplate_(tmpl.html, firstName, unsubLink);

    GmailApp.sendEmail(email, tmpl.subject, "Your email client requires HTML.", {
      htmlBody,
      inlineImages: tmpl.inlineImages, // preserves embedded images
      name: FROM_NAME,
    });

    sent++;
  }

  Logger.log(`Done. Sent: ${sent}. Skipped: ${skipped}.`);
}

/**
 * Send test email to one address.
 */
function sendNewsletterToOne() {
  const tmpl = getDraftTemplate_();

  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = getHeaderMap_(headers);

  ["email", "first_name", "status", "token"].forEach((col) => {
    if (idx[col] === undefined) throw new Error(`Missing column: ${col}`);
  });

  const webAppUrl = getWebAppUrl_();
  if (!webAppUrl) throw new Error("Web App URL missing. Deploy as Web App first.");

  const target = TEST_TARGET_EMAIL.trim().toLowerCase();
  if (!target || target === "you@example.com") {
    throw new Error("Set TEST_TARGET_EMAIL to your email address before running sendNewsletterToOne().");
  }

  for (let r = 1; r < values.length; r++) {
    const email = String(values[r][idx.email] || "").trim().toLowerCase();
    if (email !== target) continue;

    const status = String(values[r][idx.status] || "").trim().toLowerCase();
    if (status === "unsubscribed") throw new Error("Target email is unsubscribed in the sheet.");

    const token = String(values[r][idx.token] || "").trim();
    if (!token) throw new Error("Target row has no token. Generate a token first.");

    const firstName = String(values[r][idx.first_name] || "").trim() || "there";
    const unsubLink = `${webAppUrl}?t=${encodeURIComponent(token)}`;
    const htmlBody = applyTemplate_(tmpl.html, firstName, unsubLink);

    GmailApp.sendEmail(email, tmpl.subject, "Your email client requires HTML.", {
      htmlBody,
      inlineImages: tmpl.inlineImages,
      name: FROM_NAME,
    });

    Logger.log("Sent test newsletter to " + email);
    return;
  }

  throw new Error("Target email not found in sheet: " + target);
}

/**
 * Web App endpoint — one-click unsubscribe.
 * Visiting: <webapp_url>?t=<token>
 */
function doGet(e) {
  const token = e && e.parameter && e.parameter.t ? String(e.parameter.t).trim() : "";
  if (!token) return HtmlService.createHtmlOutput("Missing unsubscribe token.");

  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = getHeaderMap_(headers);

  if (idx.token === undefined || idx.status === undefined) {
    return HtmlService.createHtmlOutput("Sheet is missing required columns.");
  }

  const now = new Date();
  const unsubAtIdx = idx.unsubscribed_at; // optional

  for (let r = 1; r < values.length; r++) {
    const rowToken = String(values[r][idx.token] || "").trim();
    if (rowToken === token) {
      sh.getRange(r + 1, idx.status + 1).setValue("unsubscribed");
      if (unsubAtIdx !== undefined) sh.getRange(r + 1, unsubAtIdx + 1).setValue(now);

      return HtmlService.createHtmlOutput(
        `<p style="font-family:system-ui, -apple-system, Segoe UI, Roboto, sans-serif;">
           You’ve been unsubscribed. ✅
         </p>`
      );
    }
  }

  return HtmlService.createHtmlOutput("Unsubscribe token not found.");
}

/**
 * Token-generation (UUIDs) for rows missing tokens.
 */
function generateRandomTokensForMissing() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = getHeaderMap_(headers);

  ["email", "token"].forEach((col) => {
    if (idx[col] === undefined) throw new Error("Missing column: " + col);
  });

  let updated = 0;
  for (let r = 1; r < values.length; r++) {
    const email = String(values[r][idx.email] || "").trim().toLowerCase();
    const token = String(values[r][idx.token] || "").trim();
    if (!email || token) continue;

    sh.getRange(r + 1, idx.token + 1).setValue(Utilities.getUuid());
    updated++;
  }

  Logger.log("Generated " + updated + " UUID tokens.");
}
