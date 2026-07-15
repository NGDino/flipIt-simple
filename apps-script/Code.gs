/**
 * FlipIt Denver — Waitlist backend (Google Apps Script Web App)
 *
 * Receives POSTs from the landing-page form and appends rows to a Google Sheet.
 * Response JSON: { status: "success" | "duplicate" | "error" }
 *
 * ── SETUP ────────────────────────────────────────────────────────────────
 * 1. Create a Google Sheet. Rename the first tab to "Waitlist" (or change
 *    SHEET_NAME below to match your tab name).
 * 2. In that Sheet: Extensions ▸ Apps Script. Delete the sample code and
 *    paste this file in.
 * 3. Deploy ▸ New deployment ▸ type "Web app".
 *      - Execute as:  Me
 *      - Who has access:  Anyone
 *    Click Deploy, authorize when prompted, and copy the Web app URL
 *    (ends in /exec).
 * 4. Paste that URL into SCRIPT_URL at the top of the <script> in index.html.
 *
 * NOTE: after editing this script you must redeploy for changes to take
 * effect — Deploy ▸ Manage deployments ▸ edit (pencil) ▸ Version: New version,
 * or create a New deployment. The /exec URL stays the same when you edit an
 * existing deployment.
 * ─────────────────────────────────────────────────────────────────────────
 */

const SHEET_NAME = "Waitlist";
const HEADERS = ["Timestamp", "Name", "Email"];
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const name = (payload.name || "").trim();
    const email = (payload.email || "").trim();
    const company = (payload.company || "").trim(); // honeypot

    // Honeypot: bots autofill this hidden field. Pretend success, write nothing.
    if (company) {
      return jsonResponse({ status: "success" });
    }

    // Server-side validation (mirrors the client checks).
    if (!name || !email || !EMAIL_RE.test(email)) {
      return jsonResponse({ status: "error" });
    }

    const sheet = getSheet();

    // Ensure the header row exists on a fresh sheet.
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
    }

    // Dedupe against the Email column (case-insensitive, trimmed).
    if (emailExists(sheet, email)) {
      return jsonResponse({ status: "duplicate" });
    }

    sheet.appendRow([new Date(), name, email]);
    return jsonResponse({ status: "success" });
  } catch (err) {
    return jsonResponse({ status: "error" });
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function emailExists(sheet, email) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false; // only header (or nothing) present
  // Email is column 3; skip the header row.
  const values = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  const target = email.toLowerCase();
  return values.some((row) => String(row[0]).trim().toLowerCase() === target);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
