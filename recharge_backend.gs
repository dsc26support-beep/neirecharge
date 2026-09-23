/**
 * Kiribati recharge system — backend Web App (v15)
 * ---------------------------------------------------
 * Change from v14: both client-facing voucher emails now end with a
 * shared footer -- pill-button links to Terms/Privacy/Refund
 * Policy/Contact (HTML) or plain URLs (plain-text fallback), the
 * support address (neirecharge@gmail.com), and a professional notice
 * that replies to this email are not monitored and get archived
 * unread. The standard voucher email is now HTML+plain-text (was
 * plain-text only). SITE_BASE_URL assumes the site is still hosted
 * at https://dsc26support-beep.github.io/topup/ -- update it there
 * if that ever changes.
 *
 * Change from v13: NEW OPTIONAL SHEET TAB -- "Used Vouchers" (same
 * columns as Vouchers: A code | B amount | C used-marker). Once a
 * claimed voucher's email has actually sent, its row moves out of
 * Vouchers and into "Used Vouchers" -- keeping Vouchers limited to
 * still-available codes. If the email send fails, the claim is
 * reverted in place as before (the row is never moved in that case).
 * Silently does nothing if the tab doesn't exist yet.
 *
 * Change from v12: NEW REQUIRED SHEET TAB -- "Archive" (same columns
 * as Responses: A Timestamp | B Reference | C Name | D Email |
 * E Topup Amount | F Cost Paid | G Method | H Screenshot URL |
 * I Screenshot Hash | J Status | K Voucher Sent | L OCR Notes).
 * Add a header row matching Responses, place it wherever you like
 * (tab order is cosmetic only). Once it exists, any row that's fully
 * successful (Approved + voucher actually emailed) is automatically
 * moved out of Responses and into Archive -- keeping Responses down
 * to just Pending Review / unresolved rows. If the Archive tab
 * doesn't exist yet, archiving is silently skipped (rows just stay
 * in Responses as before -- nothing breaks). Reference and
 * screenshot-hash dedup checks now scan both sheets, so an archived
 * row still blocks a duplicate resubmission.
 *
 * Change from v11: wired in a real TIP_CELEBRATION_GIF_URL (Giphy
 * fireworks GIF) -- not verified live from this environment (no
 * general web access here), worth a manual check after deploying.
 *
 * Change from v10: dropped the "141...#" dial framing from both
 * voucher emails -- just shows the bare code now. The tip email is
 * now a full over-the-top HTML celebration (background GIF, VIP
 * "breaking news" copy) with a plain-text fallback for clients that
 * don't render HTML.
 *
 * Change from v9: amount checking now tolerates a small underpayment
 * (up to 5 cents under still counts as a match, but always forces
 * Pending Review, never auto-approves) and treats overpayment as a
 * tip -- eligible for auto-approve like an exact match, with a
 * celebratory thank-you email showing the tip amount. Also renamed
 * the email sender display name to "AM TOPUP (No-Reply)" and changed
 * the voucher line from a dial instruction to "This is your Recharge
 * Card Number".
 *
 * Change from v8: fixed "Invalid argument" from the v3 Files.create
 * branch -- v3 doesn't accept an "ocr" parameter (only v2 does); it
 * triggers OCR conversion by setting the target mimeType to a Google
 * Doc instead, with just "ocrLanguage" as the optional arg.
 *
 * Change from v7: ocrImage() now works whether the Drive advanced
 * service was added as v2 or v3 -- previously it hardcoded the v2
 * method name (Drive.Files.insert), which throws "is not a function"
 * if v3 (Drive.Files.create) was added instead.
 *
 * Change from v6: the EXIF ("looks like a photo, not a screenshot")
 * check no longer blocks submission -- it was rejecting legitimate
 * screenshots that picked up EXIF data after being forwarded/re-saved
 * through another app. It's now recorded as "Exif:true/false" in the
 * OCR Notes column for visibility only, and no longer affects
 * auto-approval eligibility.
 *
 * Change from v5: the phone number field was removed from the form
 * (the voucher code isn't tied to any specific number — the buyer
 * dials it on whichever phone they're topping up). Rate limiting and
 * screenshot filenames now key off email instead of phone.
 *
 * SETUP — Script Properties (unchanged from v5):
 *   AUTO_APPROVE_MAX, SCREENSHOT_FOLDER_ID, BANK_KEYWORDS,
 *   RATE_LIMIT_PER_HOUR, MIN_IMAGE_BYTES, MAX_TRANSACTION_AGE_HOURS,
 *   ADMIN_EMAIL
 *   Recommended BANK_KEYWORDS value based on your screenshot: "ANZ"
 *
 * IMPORTANT — Responses sheet columns changed (Phone column removed):
 * A Timestamp | B Reference | C Name | D Email | E Topup Amount
 * F Cost Paid | G Method | H Screenshot URL | I Screenshot Hash
 * J Status | K Voucher Sent | L OCR Notes
 * If you already have a live sheet from v5, either delete the old
 * Phone column (D) and shift the rest left, or start a fresh sheet.
 *
 * Other setup, same as before:
 *   - Services > + > Drive API (advanced service, enables OCR -- either
 *     v2 or v3 works, ocrImage() below detects which one is enabled)
 *   - Run createApprovalTrigger() once for the manual-approval fallback
 *   - Run createDailyDigestTrigger() once for the pending-review digest
 *   - Deploy > New deployment > Web app, Execute as Me, Access: Anyone
 */

const RESPONSES_SHEET_NAME = "Responses";
const VOUCHERS_SHEET_NAME = "Vouchers";
const ARCHIVE_SHEET_NAME = "Archive";
const USED_VOUCHERS_SHEET_NAME = "Used Vouchers";
const ACCOUNT_NUMBER = "786149";

const COL = {
  TIMESTAMP: 1, REFERENCE: 2, NAME: 3, EMAIL: 4,
  TOPUP_AMOUNT: 5, COST_AMOUNT: 6, METHOD: 7, SCREENSHOT_URL: 8,
  SCREENSHOT_HASH: 9, STATUS: 10, VOUCHER_SENT: 11, OCR_NOTES: 12,
};

// ---- Availability check (front end hides sold-out denominations) ----

function doGet(e) {
  const vSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOUCHERS_SHEET_NAME);
  const data = vSheet.getDataRange().getValues();
  const counts = {};
  for (let i = 1; i < data.length; i++) {
    const amt = Number(data[i][1]);
    const used = data[i][2];
    if (amt && !used) counts[amt] = (counts[amt] || 0) + 1;
  }
  const available = Object.keys(counts).map(Number).sort(function (a, b) { return a - b; });
  return jsonResponse({ availableAmounts: available });
}

function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ status: "error", message: "Bad request body." });
  }

  const reference = normalize(String(payload.reference || ""));
  const name = String(payload.name || "").trim();
  const email = String(payload.email || "").trim();
  const topupAmount = Number(payload.topupAmount);
  const costAmount = Number(payload.costAmount);
  const method = String(payload.method || "Internet Banking Kiribati").trim();
  const base64 = payload.screenshotBase64;
  const mimeType = payload.screenshotMimeType || "image/jpeg";

  if (!reference || !name || !email || !topupAmount || !costAmount || !base64) {
    return jsonResponse({ status: "error", message: "Missing required fields." });
  }
  if (!isValidEmail(email)) {
    return jsonResponse({ status: "error", message: "Invalid email address." });
  }

  try {
    if (isRateLimited(email)) {
      return jsonResponse({ status: "error", message: "Too many submissions recently. Please try again later." });
    }
    if (isReferenceAlreadyUsed(reference)) {
      return jsonResponse({ status: "error", message: "This reference has already been submitted." });
    }

    const imageBytes = Utilities.base64Decode(base64);

    if (!isValidImageType(imageBytes, mimeType)) {
      return jsonResponse({ status: "error", message: "That file doesn't look like a valid image. Please upload a photo/screenshot file." });
    }
    if (isTooSmall(imageBytes)) {
      return jsonResponse({ status: "error", message: "That image looks too small or empty. Please re-upload the screenshot." });
    }
    const isLikelyPhoto = hasExifMarker(imageBytes);

    const screenshotHash = computeImageHash(imageBytes);
    if (isScreenshotAlreadyUsed(screenshotHash)) {
      return jsonResponse({ status: "error", message: "This screenshot has already been used for a previous submission." });
    }

    const screenshotUrl = saveScreenshot(base64, mimeType, email);
    const ocrText = ocrImage(base64, mimeType);

    const refMatched = ocrTextContains(ocrText, reference);
    const acctMatched = ocrTextContains(ocrText, ACCOUNT_NUMBER);
    const successMatched = ocrContainsSuccessWord(ocrText);
    const bankMatched = ocrContainsBankKeyword(ocrText);
    const recency = checkTransactionRecency(ocrText);

    const amountCheck = checkAmountPaid(ocrText, costAmount);
    const amountMatched = amountCheck.matched;
    const tipAmount = amountCheck.tipAmount;

    const looksValid = refMatched && amountMatched && acctMatched &&
      successMatched && bankMatched && recency.ok;

    const props = PropertiesService.getScriptProperties();
    const autoMax = Number(props.getProperty("AUTO_APPROVE_MAX") || "0");
    const eligibleForAuto = looksValid && costAmount <= autoMax && !amountCheck.underTolerance;

    const notes = [
      "Ref:" + refMatched, "Cost:" + amountMatched, "Acct:" + acctMatched,
      "Success word:" + successMatched, "Bank name:" + bankMatched,
      "Recency:" + recency.ok + " (" + recency.note + ")",
      "Exif:" + isLikelyPhoto,
      "Paid:" + (amountCheck.paidAmount !== null ? amountCheck.paidAmount.toFixed(2) : "n/a"),
      "Tip:" + tipAmount.toFixed(2),
    ].join(" | ");

    const row = appendResponseRow({
      reference: reference, name: name, email: email,
      topupAmount: topupAmount, costAmount: costAmount, method: method,
      screenshotUrl: screenshotUrl, screenshotHash: screenshotHash,
      status: eligibleForAuto ? "Approved" : "Pending Review",
      ocrNotes: notes,
    });

    if (eligibleForAuto) {
      const sent = processApprovedRow(row);
      return jsonResponse({
        status: sent ? "approved" : "pending",
        message: sent ? "Auto-approved and email sent." : "Auto-approval passed but no matching vouchers left.",
      });
    }

    return jsonResponse({ status: "pending", message: "Submitted for manual review." });
  } catch (err) {
    return jsonResponse({ status: "error", message: "Server error: " + err.message });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function normalize(s) {
  return s.toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function u(b) {
  return b < 0 ? b + 256 : b;
}

// ---- Rate limiting (now email-only) ----

function isRateLimited(email) {
  const limit = Number(PropertiesService.getScriptProperties().getProperty("RATE_LIMIT_PER_HOUR") || "3");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESPONSES_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const oneHourAgo = new Date(Date.now() - 60 * 60 * 1000);
  const emailLower = email.toLowerCase();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const ts = new Date(data[i][COL.TIMESTAMP - 1]);
    const rowEmail = String(data[i][COL.EMAIL - 1] || "").toLowerCase();
    if (ts >= oneHourAgo && rowEmail === emailLower) count++;
  }
  return count >= limit;
}

// ---- Reference dedup ----

// Rows that were Approved + emailed get moved out of Responses and into
// Archive (see archiveRow()), so dedup checks must scan both sheets --
// otherwise a reference/screenshot from an already-completed transaction
// would look unused once it's archived, letting someone claim a second
// voucher for the same real payment.
function getResponsesAndArchiveRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responsesData = ss.getSheetByName(RESPONSES_SHEET_NAME).getDataRange().getValues().slice(1);
  const archiveSheet = ss.getSheetByName(ARCHIVE_SHEET_NAME);
  const archiveData = archiveSheet ? archiveSheet.getDataRange().getValues().slice(1) : [];
  return responsesData.concat(archiveData);
}

function isReferenceAlreadyUsed(reference) {
  const rows = getResponsesAndArchiveRows();
  for (let i = 0; i < rows.length; i++) {
    if (normalize(String(rows[i][COL.REFERENCE - 1] || "")) === reference) return true;
  }
  return false;
}

// ---- Image type validity (magic bytes) ----

function isValidImageType(bytes, mimeType) {
  const allowed = ["image/png", "image/jpeg", "image/jpg", "image/webp"];
  if (allowed.indexOf(mimeType) === -1) return false;
  if (bytes.length < 12) return false;
  const b0 = u(bytes[0]), b1 = u(bytes[1]), b2 = u(bytes[2]), b3 = u(bytes[3]);
  const isPng = b0 === 0x89 && b1 === 0x50 && b2 === 0x4e && b3 === 0x47;
  const isJpeg = b0 === 0xff && b1 === 0xd8 && b2 === 0xff;
  const isWebp = b0 === 0x52 && b1 === 0x49 && b2 === 0x46 && b3 === 0x46; // "RIFF"
  return isPng || isJpeg || isWebp;
}

// ---- Minimum size ----

function isTooSmall(bytes) {
  const minBytes = Number(PropertiesService.getScriptProperties().getProperty("MIN_IMAGE_BYTES") || "15000");
  return bytes.length < minBytes;
}

// ---- EXIF (camera photo) check ----

function hasExifMarker(bytes) {
  const marker = [0x45, 0x78, 0x69, 0x66]; // "Exif"
  const limit = bytes.length - marker.length;
  for (let i = 0; i < limit; i++) {
    if (bytes[i] === marker[0] && bytes[i + 1] === marker[1] &&
        bytes[i + 2] === marker[2] && bytes[i + 3] === marker[3]) {
      return true;
    }
  }
  return false;
}

// ---- Duplicate screenshot hash ----

function computeImageHash(bytes) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
  return digest.map(function (b) {
    const hex = u(b).toString(16);
    return hex.length === 1 ? "0" + hex : hex;
  }).join("");
}

function isScreenshotAlreadyUsed(hash) {
  const rows = getResponsesAndArchiveRows();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][COL.SCREENSHOT_HASH - 1] || "") === hash) return true;
  }
  return false;
}

// ---- OCR content checks ----

function ocrTextContains(ocrText, needle) {
  return normalize(ocrText).indexOf(normalize(needle)) !== -1;
}

// Under-payment tolerance: this much under the required cost still counts
// as a match, but always forces manual review (never auto-approves).
const UNDERPAY_TOLERANCE = 0.05;

function extractPaidAmountFromText(ocrText) {
  const match = ocrText.match(/(?:AUD|NZD|USD|\$)\s?([0-9]+\.[0-9]{2})/i);
  if (!match) return null;
  const val = parseFloat(match[1]);
  return isNaN(val) ? null : val;
}

// Checks the amount actually paid against the required cost. Exact match or
// overpayment (tip) both count as matched; a small underpayment still
// counts as matched but is flagged so it can never auto-approve. Falls back
// to an exact-string search when no dollar figure could be parsed from the
// OCR'd text at all.
function checkAmountPaid(ocrText, costAmount) {
  const paidAmount = extractPaidAmountFromText(ocrText);

  if (paidAmount === null) {
    return {
      matched: ocrTextContains(ocrText, costAmount.toFixed(2)),
      paidAmount: null, tipAmount: 0, underTolerance: false,
    };
  }

  const diff = Math.round((paidAmount - costAmount) * 100) / 100;
  if (diff >= 0) {
    return { matched: true, paidAmount: paidAmount, tipAmount: diff, underTolerance: false };
  }
  if (diff >= -UNDERPAY_TOLERANCE) {
    return { matched: true, paidAmount: paidAmount, tipAmount: 0, underTolerance: true };
  }
  return { matched: false, paidAmount: paidAmount, tipAmount: 0, underTolerance: false };
}

function ocrContainsSuccessWord(ocrText) {
  return /(successful|completed|confirmed|approved|receipt|success|posted)/i.test(ocrText);
}

function ocrContainsBankKeyword(ocrText) {
  const raw = PropertiesService.getScriptProperties().getProperty("BANK_KEYWORDS") || "";
  const keywords = raw.split(",").map(function (k) { return k.trim(); }).filter(Boolean);
  if (keywords.length === 0) return true;
  const text = ocrText.toLowerCase();
  return keywords.some(function (k) { return text.indexOf(k.toLowerCase()) !== -1; });
}

// ---- Transaction recency ----

function checkTransactionRecency(ocrText) {
  const maxAgeHours = Number(PropertiesService.getScriptProperties().getProperty("MAX_TRANSACTION_AGE_HOURS") || "48");
  const found = extractDateFromText(ocrText);
  if (!found) return { ok: true, note: "no date detected" };
  const ageHours = (Date.now() - found.getTime()) / 3600000;
  if (ageHours > maxAgeHours) return { ok: false, note: "older than " + maxAgeHours + "h" };
  return { ok: true, note: "within " + maxAgeHours + "h" };
}

function extractDateFromText(text) {
  const isoMatch = text.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (isoMatch) {
    const d = new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3]));
    if (!isNaN(d.getTime())) return d;
  }
  const slashOrDash = text.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (slashOrDash) {
    const d = new Date(Number(slashOrDash[3]), Number(slashOrDash[2]) - 1, Number(slashOrDash[1]));
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

// ---- Screenshot storage + OCR ----

function saveScreenshot(base64, mimeType, email) {
  const props = PropertiesService.getScriptProperties();
  const folder = DriveApp.getFolderById(props.getProperty("SCREENSHOT_FOLDER_ID"));
  const safeEmail = email.replace(/[^a-zA-Z0-9]/g, "_");
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType,
    "payment_" + safeEmail + "_" + new Date().getTime());
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function ocrImage(base64, mimeType) {
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, "ocr_temp");
  const name = "OCR_temp_" + new Date().getTime();
  const GOOGLE_DOC_MIME = "application/vnd.google-apps.document";

  // Drive API v3 renamed Files.insert -> Files.create, title -> name, and
  // dropped the "ocr" flag -- v3 triggers the OCR conversion by setting the
  // target mimeType instead, and rejects an unrecognized "ocr" argument
  // with "Invalid argument". v2 keeps the original ocr/ocrLanguage flags.
  const file = Drive.Files.create
    ? Drive.Files.create({ name: name, mimeType: GOOGLE_DOC_MIME }, blob, { ocrLanguage: "en" })
    : Drive.Files.insert({ title: name }, blob, { ocr: true, ocrLanguage: "en" });

  let text = "";
  try {
    text = DocumentApp.openById(file.id).getBody().getText();
  } finally {
    DriveApp.getFileById(file.id).setTrashed(true);
  }
  return text;
}

function appendResponseRow(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESPONSES_SHEET_NAME);
  sheet.appendRow([
    new Date(), data.reference, data.name, data.email,
    data.topupAmount, data.costAmount, data.method, data.screenshotUrl,
    data.screenshotHash, data.status, "", data.ocrNotes,
  ]);
  return sheet.getLastRow();
}

// ---- Voucher assignment + email send ----

const EMAIL_SENDER_NAME = "AM TOPUP (No-Reply)";

function processApprovedRow(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESPONSES_SHEET_NAME);
  const name = sheet.getRange(row, COL.NAME).getValue();
  const email = sheet.getRange(row, COL.EMAIL).getValue();
  const topupAmount = sheet.getRange(row, COL.TOPUP_AMOUNT).getValue();
  const ocrNotes = String(sheet.getRange(row, COL.OCR_NOTES).getValue() || "");
  const voucherSentCell = sheet.getRange(row, COL.VOUCHER_SENT);

  if (voucherSentCell.getValue()) return true;

  const voucher = claimNextVoucher(topupAmount);
  if (!voucher) {
    voucherSentCell.setValue("ERROR: no $" + topupAmount + " vouchers left");
    return false;
  }

  const tipMatch = ocrNotes.match(/Tip:([0-9]+\.[0-9]{2})/);
  const tipAmount = tipMatch ? parseFloat(tipMatch[1]) : 0;

  try {
    if (tipAmount > 0) {
      sendTipEmail(email, name, topupAmount, voucher.code, tipAmount);
    } else {
      sendStandardVoucherEmail(email, name, topupAmount, voucher.code);
    }
    voucherSentCell.setValue(voucher.code + " (emailed)");
    archiveRow(row);
    archiveUsedVoucher(voucher.rowIndex);
    return true;
  } catch (err) {
    voucherSentCell.setValue("ERROR: " + err.message);
    markVoucherUnused(voucher.rowIndex);
    return false;
  }
}

// Moves a fully successful row (Approved + voucher actually emailed) out of
// Responses and into Archive, keeping the active sheet limited to Pending
// Review / unresolved rows. Silently does nothing if the Archive tab hasn't
// been created yet, so a missing tab never breaks voucher delivery.
function archiveRow(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESPONSES_SHEET_NAME);
  const archiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ARCHIVE_SHEET_NAME);
  if (!archiveSheet) return;
  const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  archiveSheet.appendRow(rowValues);
  sheet.deleteRow(row);
}

// Moves a claimed voucher's row out of Vouchers and into "Used Vouchers",
// but only after its email has actually sent -- called from the same spot
// as archiveRow(), never before. If the send fails, markVoucherUnused()
// reverts the claim in place instead (this function is never reached), so
// the row is never moved out from under a claim that gets rolled back.
// Silently does nothing if the "Used Vouchers" tab hasn't been created yet.
function archiveUsedVoucher(rowIndex) {
  const vSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOUCHERS_SHEET_NAME);
  const usedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USED_VOUCHERS_SHEET_NAME);
  if (!usedSheet) return;
  const rowValues = vSheet.getRange(rowIndex, 1, 1, vSheet.getLastColumn()).getValues()[0];
  usedSheet.appendRow(rowValues);
  vSheet.deleteRow(rowIndex);
}

// Site pages linked from every client email -- pill buttons in the HTML
// version, plain URLs in the plain-text fallback for clients that don't
// render HTML.
const SITE_BASE_URL = "https://dsc26support-beep.github.io/topup/";
const SUPPORT_EMAIL = "neirecharge@gmail.com";

function buildEmailFooterHtml() {
  const pages = [
    ["Terms", "terms.html"], ["Privacy", "privacy.html"],
    ["Refund Policy", "refund.html"], ["Contact / Support", "contact.html"],
  ];
  const pills = pages.map(function (p) {
    return '<a href="' + SITE_BASE_URL + p[1] + '" style="display:inline-block;margin:4px 4px;' +
      'padding:8px 16px;border-radius:999px;background:#7c3aed;color:#fff;' +
      'font-size:0.85rem;font-weight:600;text-decoration:none;">' + p[0] + '</a>';
  }).join("");

  return (
    '<div style="margin-top:24px;padding-top:16px;border-top:1px solid #e5e7eb;text-align:center;font-family:sans-serif;">' +
    '<p style="font-size:0.85rem;color:#6b7280;margin:0 0 10px;">Tap a button below for our Terms, Privacy Policy, Refund Policy, or Support.</p>' +
    '<div>' + pills + '</div>' +
    '<p style="font-size:0.85rem;color:#6b7280;margin:16px 0 0;">' +
    'Questions or an issue with your top-up? Email our support team directly at ' +
    '<a href="mailto:' + SUPPORT_EMAIL + '" style="color:#7c3aed;">' + SUPPORT_EMAIL + '</a>.</p>' +
    '<p style="font-size:0.78rem;color:#9ca3af;margin:12px 0 0;">' +
    'This is an automated message. Replies sent directly to this email address are not monitored ' +
    'and will be archived without response. For assistance, please contact us using the address above.</p>' +
    '</div>'
  );
}

function buildEmailFooterPlainText() {
  return (
    "\n\n---\n" +
    "Terms: " + SITE_BASE_URL + "terms.html\n" +
    "Privacy: " + SITE_BASE_URL + "privacy.html\n" +
    "Refund Policy: " + SITE_BASE_URL + "refund.html\n" +
    "Contact / Support: " + SITE_BASE_URL + "contact.html\n\n" +
    "Questions or an issue with your top-up? Email our support team directly at " + SUPPORT_EMAIL + ".\n\n" +
    "This is an automated message. Replies sent directly to this email address are not monitored " +
    "and will be archived without response. For assistance, please contact us using the address above.\n"
  );
}

function sendStandardVoucherEmail(email, name, topupAmount, code) {
  const plainBody =
    "Hi " + name + ",\n\n" +
    "Your $" + topupAmount + " top-up is confirmed.\n\n" +
    "This is your Recharge Card Number:\n" +
    code + "\n\n" +
    "Ko rabwa\nNei Recharge.\n" +
    buildEmailFooterPlainText();

  const htmlBody =
    '<div style="font-family:sans-serif;max-width:420px;margin:0 auto;padding:24px;">' +
    '<p>Hi ' + name + ',</p>' +
    '<p>Your $' + topupAmount + ' top-up is confirmed.</p>' +
    '<p style="font-size:0.9rem;color:#6b7280;margin-bottom:4px;">This is your Recharge Card Number</p>' +
    '<p style="font-size:1.3rem;font-weight:700;letter-spacing:2px;">' + code + '</p>' +
    '<p>Ko rabwa<br>Nei Recharge.</p>' +
    buildEmailFooterHtml() +
    '</div>';

  MailApp.sendEmail(String(email), "Your phone top-up code", plainBody, {
    htmlBody: htmlBody,
    name: EMAIL_SENDER_NAME,
  });
}

const TIP_CELEBRATION_GIF_URL = "https://media.giphy.com/media/TmT51OyQLFD7a/giphy.gif";

function sendTipEmail(email, name, topupAmount, code, tipAmount) {
  const subject = "🚨 BREAKING: " + name + " IS OFFICIALLY A TOP-UP VIP 🚨";

  const plainBody =
    "Hi " + name + ",\n\n" +
    "🚨 BREAKING NEWS 🚨\n\n" +
    "Your $" + topupAmount + " top-up is CONFIRMED -- and you tipped $" +
    tipAmount.toFixed(2) + " on top. This is your Recharge Card Number:\n" +
    code + "\n\n" +
    "By order of the Ministry of Generosity, you have been promoted to " +
    "OFFICIAL VIP TOP-UP LEGEND. Your tip goes straight into keeping this " +
    "page alive and improving for everyone. We are, frankly, emotional.\n\n" +
    "Ko rabwa\nNei Recharge.\n" +
    buildEmailFooterPlainText();

  const htmlBody =
    '<div style="font-family:sans-serif;text-align:center;padding:24px;' +
    'background:url(\'' + TIP_CELEBRATION_GIF_URL + '\') center/cover;">' +
    '<div style="background:rgba(255,255,255,0.92);border-radius:12px;padding:24px;max-width:420px;margin:0 auto;">' +
    '<h1 style="margin:0 0 8px;font-size:1.4rem;">🚨 BREAKING NEWS 🚨</h1>' +
    '<p style="font-size:1.1rem;font-weight:700;margin:0 0 16px;">' +
    name + ' IS OFFICIALLY A TOP-UP VIP</p>' +
    '<p>Your <b>$' + topupAmount + '</b> top-up is <b>CONFIRMED</b> -- and you tipped an extra ' +
    '<b>$' + tipAmount.toFixed(2) + '</b> on top!</p>' +
    '<p style="font-size:0.9rem;color:#6b7280;">This is your Recharge Card Number</p>' +
    '<p style="font-size:1.3rem;font-weight:700;letter-spacing:2px;">' + code + '</p>' +
    '<p>By order of the Ministry of Generosity, you have been promoted to ' +
    '<b>OFFICIAL VIP TOP-UP LEGEND</b>. Your tip goes straight into keeping ' +
    'this page alive and improving for everyone. We are, frankly, emotional. 🎉🎈</p>' +
    '<p>Ko rabwa<br>Nei Recharge.</p>' +
    buildEmailFooterHtml() +
    '</div></div>';

  MailApp.sendEmail(String(email), subject, plainBody, {
    htmlBody: htmlBody,
    name: EMAIL_SENDER_NAME,
  });
}

function claimNextVoucher(topupAmount) {
  const vSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOUCHERS_SHEET_NAME);
  const data = vSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowAmount = Number(data[i][1]);
    const used = data[i][2];
    if (data[i][0] && rowAmount === Number(topupAmount) && !used) {
      vSheet.getRange(i + 1, 3).setValue("used " + new Date().toISOString());
      return { code: data[i][0], rowIndex: i + 1 };
    }
  }
  return null;
}

function markVoucherUnused(rowIndex) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOUCHERS_SHEET_NAME)
    .getRange(rowIndex, 3).setValue("");
}

// ---- Manual-approval fallback: type "Approved" in the Status column ----

function createApprovalTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "onStatusEdit") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("onStatusEdit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}

function onStatusEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== RESPONSES_SHEET_NAME) return;
  if (range.getColumn() !== COL.STATUS) return;
  if (String(range.getValue()).trim().toLowerCase() !== "approved") return;
  processApprovedRow(range.getRow());
}

// ---- Daily "pending review" digest ----

function sendPendingReviewDigest() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESPONSES_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const pending = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL.STATUS - 1]).toLowerCase() === "pending review") {
      pending.push(
        "- " + data[i][COL.REFERENCE - 1] + " | " + data[i][COL.NAME - 1] +
        " | $" + data[i][COL.TOPUP_AMOUNT - 1] + " (paid $" + data[i][COL.COST_AMOUNT - 1] + ")" +
        " | " + data[i][COL.SCREENSHOT_URL - 1]
      );
    }
  }
  if (pending.length === 0) return;

  const adminEmail = PropertiesService.getScriptProperties().getProperty("ADMIN_EMAIL")
    || Session.getEffectiveUser().getEmail();
  MailApp.sendEmail(
    adminEmail,
    "Recharge system: " + pending.length + " pending review(s)",
    "The following submissions need manual review:\n\n" + pending.join("\n")
  );
}

function createDailyDigestTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "sendPendingReviewDigest") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("sendPendingReviewDigest")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
}
