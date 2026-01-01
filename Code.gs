/**
 * Google Docs → Gmail Newsletter (Apps Script)
 * - Author in Google Docs → paste into a Gmail Draft (keeps formatting + embedded images)
 * - Per-recipient send with mail merge
 * - One-click unsubscribe via Apps Script Web App endpoint
 * - Open tracking via 1x1 SVG pixel (logs to Events sheet)
 * - Preserves inline embedded images (cid:) by extracting via Gmail API + sending as inlineImages
 *
 * SECURITY / PRIVACY NOTE
 * This file is GitHub-safe: it contains NO personal emails, spreadsheet IDs, or deployment URLs.
 * Provide your values in the CONFIG section (or Script Properties) before running.
 *
 * REQUIRED SHEETS
 * 1) Subscribers tab (default: "Subscribers")
 *    Required columns: email, first_name, status, token
 *    Optional column: unsubscribed_at
 *
 * 2) Events tab (default: "Events")
 *    Recommended columns:
 *      ts | event | campaign_id | token | email | url
 *
 * REQUIRED SETUP
 * - Enable Advanced Gmail Service:
 *   Apps Script editor → Services → Add Gmail API
 *
 * - Deploy as Web App:
 *   Execute as: Me
 *   Who has access: Anyone (or Anyone with link)
 *   Copy the /exec URL (it becomes your WEB_APP_URL)
 *
 * TEST FLOW
 * - Run generateRandomTokensForMissing()
 * - Run PRINT_DRAFT_IDS() and set DRAFT_ID
 * - Run sendNewsletterToOne() to test
 * - Verify:
 *   - Unsubscribe link works
 *   - Opens are logged (Events tab)
 *   - Images render inline (not as attachments)
 */

/* =========================
 * CONFIG
 * ========================= */

// If you commit this to GitHub, keep these as placeholders.
// You can also move these into Script Properties (recommended for real use).

const CONFIG = {
  // Google Sheets
  SPREADSHEET_ID: "REPLACE_WITH_YOUR_SPREADSHEET_ID",
  SHEET_NAME: "Subscribers",
  EVENTS_SHEET_NAME: "Events",

  // Email
  FROM_NAME: "Your Name",
  TEST_TARGET_EMAIL: "you@example.com", // used only by sendNewsletterToOne()

  // Draft
  DRAFT_ID: "r-REPLACE_WITH_YOUR_DRAFT_ID",

  // Web App /exec URL (post-deploy). Used for unsubscribe + open tracking pixel URLs.
  WEB_APP_URL: "https://script.google.com/macros/s/REPLACE_WITH_DEPLOYMENT_ID/exec",

  // Unsubscribe placeholder (only if your Gmail draft can't include {{unsub_link}} as a link)
  PLACEHOLDER_UNSUB_URL: "https://example.com/unsub",

  // Tracking: campaign id
  DEFAULT_CAMPAIGN_ID: "campaign_YYYY_MM_DD",
};

/* =========================
 * Sheet helpers
 * ========================= */

function getSpreadsheet_() {
  const id = String(CONFIG.SPREADSHEET_ID || "").trim();
  if (!id || id.includes("REPLACE_WITH")) throw new Error("CONFIG.SPREADSHEET_ID is missing.");
  return SpreadsheetApp.openById(id);
}

function getSheet_() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error(`Missing sheet tab: "${CONFIG.SHEET_NAME}"`);
  return sh;
}

function getEventsSheet_() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(CONFIG.EVENTS_SHEET_NAME);
  if (!sh) throw new Error(`Missing sheet tab: "${CONFIG.EVENTS_SHEET_NAME}"`);
  return sh;
}

function getHeaderMap_(headers) {
  const m = {};
  headers.forEach((h, i) => (m[String(h || "").trim()] = i));
  return m;
}

function getWebAppUrl_() {
  const url = String(CONFIG.WEB_APP_URL || "").trim();
  if (!url || url.includes("REPLACE_WITH")) throw new Error("CONFIG.WEB_APP_URL is missing (/exec URL).");
  if (!/^https:\/\/script\.google\.com\/macros\/s\//i.test(url) || !url.includes("/exec")) {
    throw new Error(`CONFIG.WEB_APP_URL must be a valid Apps Script /exec URL. Got: ${url}`);
  }
  return url;
}

/* =========================
 * Draft helpers
 * ========================= */

function PRINT_DRAFT_IDS() {
  const drafts = GmailApp.getDrafts();
  Logger.log(`Found ${drafts.length} drafts.`);
  drafts.forEach((d, i) => {
    const msg = d.getMessage();
    Logger.log(`${i + 1}. ID=${d.getId()} | Subject="${msg.getSubject()}" | Updated=${msg.getDate()}`);
  });
  Logger.log("Copy the correct ID and set CONFIG.DRAFT_ID.");
}

function getDraftById_() {
  const id = String(CONFIG.DRAFT_ID || "").trim();
  if (!id || id.includes("REPLACE_WITH")) throw new Error("CONFIG.DRAFT_ID is missing. Run PRINT_DRAFT_IDS().");
  return GmailApp.getDraft(id);
}

function assertHasPlaceholders_(html) {
  const missing = [];
  if (!html || !html.trim()) missing.push("(empty body)");
  if (html && !html.includes("{{first_name}}")) missing.push("{{first_name}}");

  const hasUnsubToken = html && html.includes("{{unsub_link}}");
  const hasPlaceholderUrl =
    !!CONFIG.PLACEHOLDER_UNSUB_URL && html && html.includes(String(CONFIG.PLACEHOLDER_UNSUB_URL || ""));
  if (!hasUnsubToken && !hasPlaceholderUrl) {
    missing.push("{{unsub_link}} (or set CONFIG.PLACEHOLDER_UNSUB_URL to a URL present in the draft)");
  }

  if (missing.length) throw new Error("Draft template missing required placeholders: " + missing.join(", "));
}

/**
 * Attachment fetch can occasionally return an "Empty response" transiently.
 * Retry with short backoff.
 */
function getAttachmentWithRetry_(userId, messageId, attachmentId) {
  let lastErr = null;
  for (let i = 0; i < 3; i++) {
    try {
      const res = Gmail.Users.Messages.Attachments.get(userId, messageId, attachmentId);
      if (res && res.data) return res;
      throw new Error("Empty response");
    } catch (e) {
      lastErr = e;
      Utilities.sleep(250 * (i + 1));
    }
  }
  throw lastErr;
}

/**
 * Draft template loader
 * - HTML + subject via GmailApp (most reliable for unicode/emoji)
 * - Inline cid: images via Gmail API Drafts.get + attachments.get
 *
 * Requires Advanced Gmail Service enabled (Gmail API).
 */
function getDraftTemplate_() {
  // 1) HTML + subject (emoji-safe)
  const draftApp = getDraftById_();
  const msgApp = draftApp.getMessage();
  const subject = msgApp.getSubject() || "(no subject)";
  const html = msgApp.getBody() || "";
  if (!html) throw new Error("Draft body is empty (GmailApp.getBody()).");
  assertHasPlaceholders_(html);

  // 2) Inline images via Gmail API
  const draftId = String(CONFIG.DRAFT_ID || "").trim();
  let inlineImages = {};

  try {
    const draftApi = Gmail.Users.Drafts.get("me", draftId, { format: "full" });
    const msgApi = draftApi && draftApi.message ? draftApi.message : null;

    if (!msgApi || !msgApi.id || !msgApi.payload) {
      Logger.log("WARNING: Could not load draft via Gmail API. Sending without inline images.");
      return { subject, html, inlineImages: {} };
    }

    const msgId = String(msgApi.id).trim(); // Gmail API message id
    inlineImages = extractInlineImages_(msgApi.payload, msgId) || {};
  } catch (e) {
    Logger.log("WARNING: Inline image extraction failed; sending without inline images. Error: " + e);
    inlineImages = {};
  }

  // If HTML references cid: but we have no images, warn.
  if ((html.match(/src=["']cid:/gi) || []).length > 0 && Object.keys(inlineImages).length === 0) {
    Logger.log("WARNING: HTML contains cid: images but no inline images were extracted.");
  }

  return { subject, html, inlineImages };
}

/* =========================
 * Inline image extraction (cid:)
 * ========================= */

function extractInlineImages_(payload, msgId) {
  const inline = {};
  walkParts_(payload, (part) => {
    if (!part) return;

    const mime = String(part.mimeType || "");
    const isImage = mime.startsWith("image/");
    const filename = part.filename || "";
    const attachmentId = part.body && part.body.attachmentId ? String(part.body.attachmentId) : "";

    if (!isImage || !attachmentId) return;

    const contentId = getHeader_(part.headers, "Content-ID");
    const xAttachId = getHeader_(part.headers, "X-Attachment-Id");
    if (!contentId && !xAttachId) return;

    const attachment = getAttachmentWithRetry_("me", msgId, attachmentId);
    const bytes = decodeBase64UrlToBytes_(attachment.data);
    const blob = Utilities.newBlob(bytes, mime, filename || "inline");

    const keys = [];
    if (contentId) keys.push(...normalizeContentIdVariants_(contentId));
    if (xAttachId) keys.push(String(xAttachId).trim());

    keys
      .map((k) => String(k || "").trim())
      .filter(Boolean)
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

function normalizeContentIdVariants_(cid) {
  const raw = String(cid || "").trim().replace(/^<|>$/g, "");
  if (!raw) return [];
  const variants = [raw];

  const at = raw.indexOf("@");
  if (at > 0) variants.push(raw.slice(0, at));

  variants.push(`<${raw}>`);
  return Array.from(new Set(variants));
}

/* =========================
 * HTML helpers
 * ========================= */

function escapeHtml_(s) {
  return String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

/**
 * Adds/overrides inline styles for cid: images so they fill table cells more reliably.
 */
function hardenCidImagesForEmail_(html) {
  if (!html) return html;

  return html.replace(/<img\b([^>]*\bsrc=["']cid:[^"']+["'][^>]*)>/gi, (full, attrs) => {
    const hasStyle = /\bstyle=["'][^"']*["']/.test(attrs);
    const desired = "display:block;width:100%;height:auto;max-width:100%;border:0;";

    if (hasStyle) {
      attrs = attrs.replace(/\bstyle=["']([^"']*)["']/, (m, s) => {
        const merged = (s || "").trim();
        const sep = merged && !merged.endsWith(";") ? ";" : "";
        return `style="${merged}${sep}${desired}"`;
      });
    } else {
      attrs += ` style="${desired}"`;
    }

    if (!/\bwidth=/.test(attrs)) attrs += ` width="100%"`;
    return `<img${attrs}>`;
  });
}

/**
 * Optional hardening: convert non-ASCII to numeric HTML entities.
 * Use this ONLY if you see charset issues in a client that can't handle UTF-8.
 * Most modern clients are fine without it, especially for Gmail→Gmail.
 */
function encodeNonAsciiToEntities_(input) {
  const s = String(input || "");
  let out = "";
  for (let i = 0; i < s.length; i++) {
    const codePoint = s.codePointAt(i);
    if (codePoint > 0xffff) i++;
    if (codePoint <= 0x7f) out += String.fromCodePoint(codePoint);
    else out += `&#x${codePoint.toString(16).toUpperCase()};`;
  }
  return out;
}

/* =========================
 * Tracking URLs + HTML mutation
 * ========================= */

function extractTokenFromUnsubLink_(unsubLink) {
  const s = String(unsubLink || "");
  const m = s.match(/[?&]t=([^&]+)/i);
  return m ? safeDecodeURIComponent_(m[1]) : "";
}

function buildTrackOpenUrl_(token, campaignId) {
  const base = getWebAppUrl_();
  return `${base}?mode=track_open&t=${encodeURIComponent(token || "")}&cid=${encodeURIComponent(campaignId || CONFIG.DEFAULT_CAMPAIGN_ID)}`;
}

function injectOpenPixel_(html, token, campaignId) {
  if (!html || !token) return html;

  const pixelUrl = buildTrackOpenUrl_(token, campaignId);
  const tag = `<img src="${pixelUrl}" width="1" height="1" style="display:block;border:0;outline:none;text-decoration:none;" alt="">`;

  if (/<\/body>/i.test(html)) return html.replace(/<\/body>/i, `${tag}</body>`);
  if (/<\/html>/i.test(html)) return html.replace(/<\/html>/i, `${tag}</html>`);
  return html + tag;
}

/**
 * SINGLE SOURCE OF TRUTH for mail merge + tracking injection.
 */
function applyTemplate_(templateHtml, firstName, unsubLink, campaignIdOpt) {
  const campaignId = String(campaignIdOpt || CONFIG.DEFAULT_CAMPAIGN_ID || "").trim() || "default";

  let html = templateHtml.replaceAll("{{first_name}}", escapeHtml_(firstName));

  if (templateHtml.includes("{{unsub_link}}")) {
    html = html.replaceAll("{{unsub_link}}", unsubLink);
  } else if (CONFIG.PLACEHOLDER_UNSUB_URL && templateHtml.includes(String(CONFIG.PLACEHOLDER_UNSUB_URL))) {
    html = html.replaceAll(String(CONFIG.PLACEHOLDER_UNSUB_URL), unsubLink);
  }

  // Hardening for cid: images
  html = hardenCidImagesForEmail_(html);

  // Open tracking pixel
  const token = extractTokenFromUnsubLink_(unsubLink);
  html = injectOpenPixel_(html, token, campaignId);

  // Optional: uncomment only if you hit weird charset issues in some client
  // html = encodeNonAsciiToEntities_(html);

  return html;
}

/* =========================
 * Sending
 * ========================= */

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

function sendNewsletter() {
  const tmpl = getDraftTemplate_();

  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = getHeaderMap_(headers);

  ["email", "first_name", "status", "token"].forEach((col) => {
    if (idx[col] === undefined) throw new Error(`Missing column: ${col}`);
  });

  const webAppUrl = getWebAppUrl_();

  let sent = 0;
  let skipped = 0;

  for (let r = 1; r < values.length; r++) {
    const email = String(values[r][idx.email] || "").trim().toLowerCase();
    const status = String(values[r][idx.status] || "").trim().toLowerCase();
    const token = String(values[r][idx.token] || "").trim();
    const firstName = String(values[r][idx.first_name] || "").trim() || "there";

    if (!email || !token) { skipped++; continue; }
    if (status === "unsubscribed") { skipped++; continue; }

    const unsubLink = `${webAppUrl}?t=${encodeURIComponent(token)}`;
    const htmlBody = applyTemplate_(tmpl.html, firstName, unsubLink, CONFIG.DEFAULT_CAMPAIGN_ID);

    GmailApp.sendEmail(email, tmpl.subject, "Your email client requires HTML.", {
      htmlBody,
      inlineImages: tmpl.inlineImages,
      name: CONFIG.FROM_NAME,
    });

    sent++;
  }

  Logger.log(`Done. Sent: ${sent}. Skipped: ${skipped}.`);
}

function sendNewsletterToOne() {
  const tmpl = getDraftTemplate_();

  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = getHeaderMap_(headers);

  ["email", "first_name", "status", "token"].forEach((col) => {
    if (idx[col] === undefined) throw new Error(`Missing column: ${col}`);
  });

  const target = String(CONFIG.TEST_TARGET_EMAIL || "").trim().toLowerCase();
  if (!target || target.includes("you@")) throw new Error("Set CONFIG.TEST_TARGET_EMAIL before running sendNewsletterToOne().");

  const webAppUrl = getWebAppUrl_();

  for (let r = 1; r < values.length; r++) {
    const email = String(values[r][idx.email] || "").trim().toLowerCase();
    if (email !== target) continue;

    const status = String(values[r][idx.status] || "").trim().toLowerCase();
    if (status === "unsubscribed") throw new Error("Target email is unsubscribed in the sheet.");

    const token = String(values[r][idx.token] || "").trim();
    if (!token) throw new Error("Target row has no token. Generate a token first.");

    const firstName = String(values[r][idx.first_name] || "").trim() || "there";
    const unsubLink = `${webAppUrl}?t=${encodeURIComponent(token)}`;
    const htmlBody = applyTemplate_(tmpl.html, firstName, unsubLink, CONFIG.DEFAULT_CAMPAIGN_ID);

    GmailApp.sendEmail(email, tmpl.subject, "Your email client requires HTML.", {
      htmlBody,
      inlineImages: tmpl.inlineImages,
      name: CONFIG.FROM_NAME,
    });

    Logger.log("Sent test newsletter to " + email);
    return;
  }

  throw new Error("Target email not found in sheet: " + target);
}

/* =========================
 * Tokens
 * ========================= */

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

/* =========================
 * Web App endpoint (single doGet)
 * ========================= */

function doGet(e) {
  try {
    const p = (e && e.parameter) ? e.parameter : {};
    const mode = String(p.mode || "").trim().toLowerCase();

    const token = String(p.t || "").trim();
    const campaignId = String(p.cid || CONFIG.DEFAULT_CAMPAIGN_ID).trim() || CONFIG.DEFAULT_CAMPAIGN_ID;

    // DEBUG: /exec?mode=echo&t=...&cid=...&u=...
    if (mode === "echo") {
      const rawU = (p.u != null) ? String(p.u) : "";
      const decU = safeDecodeURIComponent_(rawU);

      const out =
        "doGet reached\n" +
        "mode=" + mode + "\n" +
        "t=" + token + "\n" +
        "cid=" + campaignId + "\n" +
        "u(raw)=" + rawU + "\n" +
        "u(decoded)=" + decU + "\n";

      return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.TEXT);
    }

    // OPEN TRACKING: /exec?mode=track_open&t=...&cid=...
    if (mode === "track_open") {
      if (token) {
        try { logEvent_("open", token, campaignId, "", p); } catch (err) {}
      }

      return ContentService
        .createTextOutput(`<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"></svg>`)
        .setMimeType(ContentService.MimeType.SVG);
    }

    // UNSUBSCRIBE FALLBACK: /exec?t=<token>
    if (token) {
      const sh = getSheet_();
      const values = sh.getDataRange().getValues();
      const headers = values[0];
      const idx = getHeaderMap_(headers);

      if (idx.token === undefined || idx.status === undefined) {
        return HtmlService.createHtmlOutput("Sheet missing required columns: token, status");
      }

      const now = new Date();
      const unsubAtIdx = idx.unsubscribed_at; // optional

      for (let r = 1; r < values.length; r++) {
        const rowToken = String(values[r][idx.token] || "").trim();
        if (rowToken === token) {
          sh.getRange(r + 1, idx.status + 1).setValue("unsubscribed");
          if (unsubAtIdx !== undefined) sh.getRange(r + 1, unsubAtIdx + 1).setValue(now);

          return HtmlService.createHtmlOutput(
            `<p style="font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif;">
              You’ve been unsubscribed. ✅
            </p>`
          );
        }
      }

      return HtmlService.createHtmlOutput("Unsubscribe token not found.");
    }

    return HtmlService.createHtmlOutput("OK (no mode).");
  } catch (err) {
    return ContentService
      .createTextOutput("doGet ERROR:\n" + (err && err.stack ? err.stack : err))
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function safeDecodeURIComponent_(s) {
  try { return decodeURIComponent(String(s || "")); }
  catch (e) { return String(s || ""); }
}

/* =========================
 * Events logging
 * ========================= */

function logEvent_(eventName, token, campaignId, url, paramsObj) {
  const shEvents = getEventsSheet_();
  const now = new Date();
  const email = lookupEmailByToken_(token) || "";
  shEvents.appendRow([now, eventName, campaignId, token, email, url || ""]);
}

function lookupEmailByToken_(token) {
  if (!token) return "";
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = getHeaderMap_(headers);
  if (idx.token === undefined || idx.email === undefined) return "";

  for (let r = 1; r < values.length; r++) {
    const rowToken = String(values[r][idx.token] || "").trim();
    if (rowToken === token) return String(values[r][idx.email] || "").trim().toLowerCase();
  }
  return "";
}

/* =========================
 * Debug helpers
 * ========================= */

function DEBUG_dumpDraftHtml() {
  const tmpl = getDraftTemplate_();
  Logger.log(tmpl.html.slice(0, 2000));
}

function AUTH_testWrite() {
  const sh = getEventsSheet_();
  sh.appendRow([new Date(), "auth_test", CONFIG.DEFAULT_CAMPAIGN_ID, "TEST", "test@example.com", ""]);
  Logger.log("AUTH_testWrite OK");
}

/* =========================
 * Scheduling helpers (Time-driven triggers)
 * ========================= */

/**
 * Schedule a one-time send at a specific local datetime.
 * This deletes any existing triggers for sendNewsletter to avoid duplicate sends.
 *
 * Usage:
 *  - Set whenStr e.g. "2026-01-01 09:00"
 *  - Run scheduleSendAt_()
 */
function scheduleSendAt_() {
  const whenStr = "YYYY-MM-DD HH:MM"; // <-- EDIT ME

  const m = String(whenStr).match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2})$/);
  if (!m) throw new Error('Bad format. Use "YYYY-MM-DD HH:MM"');

  const when = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]), 0);

  // Remove existing triggers for sendNewsletter
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === "sendNewsletter") ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger("sendNewsletter").timeBased().at(when).create();
  Logger.log("Scheduled sendNewsletter for: " + when);
}

/**
 * Convenience: schedule for tomorrow at 09:00 local time.
 */
function scheduleNewsletterForTomorrowMorning_() {
  const now = new Date();
  const when = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1, 9, 0, 0);

  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === "sendNewsletter") ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger("sendNewsletter").timeBased().at(when).create();
  Logger.log("Scheduled sendNewsletter for: " + when);
}

/* =========================
 * Base64 decode helper for Gmail API attachment.data
 * ========================= */

/**
 * Robust decoder for Gmail API payload fields:
 * - Accepts base64url/base64 strings
 * - Accepts byte arrays / stringified byte lists (including negatives)
 * - Accepts raw HTML strings that already start with "<"
 */
function decodeBase64UrlToBytes_(data) {
  if (data == null) return [];

  if (Array.isArray(data)) {
    return data
      .map((n) => Number(n))
      .filter((n) => Number.isFinite(n))
      .map((n) => ((n % 256) + 256) % 256);
  }

  let s = String(data).trim();
  if ((s.startsWith('"') && s.endsWith('"')) || (s.startsWith("'") && s.endsWith("'"))) {
    s = s.slice(1, -1).trim();
  }

  if (s.startsWith("<")) {
    return Utilities.newBlob(s, "text/plain", "raw").getBytes();
  }

  if (s.includes(",") && /-?\d/.test(s)) {
    const nums = s.match(/-?\d{1,4}/g);
    if (nums && nums.length > 10) {
      const bytes = nums
        .map((x) => parseInt(x, 10))
        .filter((n) => Number.isFinite(n))
        .map((n) => ((n % 256) + 256) % 256)
        .filter((n) => n >= 0 && n <= 255);
      if (bytes.length) return bytes;
    }
  }

  let b64 = s.replace(/-/g, "+").replace(/_/g, "/").replace(/\s/g, "");
  const pad = b64.length % 4;
  if (pad) b64 += "====".slice(pad);

  return Utilities.base64Decode(b64);
}
