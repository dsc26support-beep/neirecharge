/**
 * Kiribati recharge system — backend Web App (v9)
 * ---------------------------------------------------
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
    const amountMatched = ocrTextContains(ocrText, costAmount.toFixed(2));
    const acctMatched = ocrTextContains(ocrText, ACCOUNT_NUMBER);
    const successMatched = ocrContainsSuccessWord(ocrText);
    const bankMatched = ocrContainsBankKeyword(ocrText);
    const recency = checkTransactionRecency(ocrText);

    const looksValid = refMatched && amountMatched && acctMatched &&
      successMatched && bankMatched && recency.ok;

    const props = PropertiesService.getScriptProperties();
    const autoMax = Number(props.getProperty("AUTO_APPROVE_MAX") || "0");
    const eligibleForAuto = looksValid && costAmount <= autoMax;

    const notes = [
      "Ref:" + refMatched, "Cost:" + amountMatched, "Acct:" + acctMatched,
      "Success word:" + successMatched, "Bank name:" + bankMatched,
      "Recency:" + recency.ok + " (" + recency.note + ")",
      "Exif:" + isLikelyPhoto,
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

function isReferenceAlreadyUsed(reference) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESPONSES_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normalize(String(data[i][COL.REFERENCE - 1] || "")) === reference) return true;
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESPONSES_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL.SCREENSHOT_HASH - 1] || "") === hash) return true;
  }
  return false;
}

// ---- OCR content checks ----

function ocrTextContains(ocrText, needle) {
  return normalize(ocrText).indexOf(normalize(needle)) !== -1;
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

function processApprovedRow(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESPONSES_SHEET_NAME);
  const name = sheet.getRange(row, COL.NAME).getValue();
  const email = sheet.getRange(row, COL.EMAIL).getValue();
  const topupAmount = sheet.getRange(row, COL.TOPUP_AMOUNT).getValue();
  const voucherSentCell = sheet.getRange(row, COL.VOUCHER_SENT);

  if (voucherSentCell.getValue()) return true;

  const voucher = claimNextVoucher(topupAmount);
  if (!voucher) {
    voucherSentCell.setValue("ERROR: no $" + topupAmount + " vouchers left");
    return false;
  }

  const subject = "Your phone top-up code";
  const body =
    "Hi " + name + ",\n\n" +
    "Your $" + topupAmount + " top-up is confirmed.\n\n" +
    "On the phone you're topping up, dial:\n" +
    "141" + voucher.code + "#\n\n" +
    "This applies the credit to your balance.\n";

  try {
    MailApp.sendEmail(String(email), subject, body);
    voucherSentCell.setValue(voucher.code + " (emailed)");
    return true;
  } catch (err) {
    voucherSentCell.setValue("ERROR: " + err.message);
    markVoucherUnused(voucher.rowIndex);
    return false;
  }
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
