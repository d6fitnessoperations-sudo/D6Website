/**
 * D6 Fitness Studio - Trial Session Form + Health & Fitness Assessment -> Google Sheets
 *
 * Setup (replace the existing Code.gs with this file, then redeploy):
 *  1. Open the Google Sheet. Extensions > Apps Script. Replace Code.gs with this file. Save.
 *  2. Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy.
 *     (Keep the same Web app URL, so the website does not need changing.)
 *
 * Trial form (no type parameter):  appends Timestamp, Name, Phone to the "Responses" tab.
 * Assessment (type=assessment):     appends a row to the "Assessment" tab.
 *   Only sent when the visitor presses "Show Detailed Report" and agrees to the notice.
 *
 * Tabs are created with bold, frozen headers if they are missing.
 */

const SHEET_NAME = "Responses";
const HEADERS = ["Timestamp", "Name", "Phone"];

const ASSESSMENT_SHEET_NAME = "Assessment";
const ASSESSMENT_HEADERS = [
  "Timestamp", "Name", "Mobile", "Overall Score", "Stage",
  "Physical Activity", "Sleep & Recovery", "Endurance", "Nutrition",
  "Assessment Version"
];

function doGet(e) {
  return handleRequest_(e);
}

function doPost(e) {
  return handleRequest_(e);
}

function handleRequest_(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const params = getParams_(e);

    if (clean_(params.type) === "assessment") {
      return handleAssessment_(params);
    }

    // ---- Trial session form (unchanged) ----
    const name = clean_(params.name);
    const phone = clean_(params.phone);

    if (!name || !phone) {
      return json_({ success: false, error: "Name and phone are required." });
    }

    const sheet = getOrCreateSheet_(SHEET_NAME, HEADERS);
    sheet.appendRow([new Date(), name, phone]);

    return json_({ success: true });
  } catch (err) {
    return json_({ success: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/**
 * Saves one assessment result. Scores are sent by the page and checked here
 * (each must be a number from 0 to 100). Nothing is stored unless the visitor
 * opted in on the site.
 */
function handleAssessment_(params) {
  const name = clean_(params.name);
  const mobile = clean_(params.mobile);
  const stage = clean_(params.stage);
  const version = clean_(params.version);

  if (!name) {
    return json_({ success: false, error: "Name is required." });
  }
  if (!/^[+\d][\d\s-]{7,}$/.test(mobile)) {
    return json_({ success: false, error: "A valid mobile number is required." });
  }

  const scores = [
    params.overall,
    params.activity,
    params.sleep,
    params.endurance,
    params.nutrition
  ].map(function (value) { return Number(clean_(value)); });

  const valid = scores.every(function (n) { return isFinite(n) && n >= 0 && n <= 100; });
  if (!valid) {
    return json_({ success: false, error: "Scores must be numbers between 0 and 100." });
  }

  const sheet = getOrCreateSheet_(ASSESSMENT_SHEET_NAME, ASSESSMENT_HEADERS);
  sheet.appendRow([
    new Date(), name, mobile,
    scores[0], stage,
    scores[1], scores[2], scores[3], scores[4],
    version
  ]);

  return json_({ success: true });
}

function getOrCreateSheet_(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const existing = headerRange.getValues()[0];
  const headersMissing = existing.every(function (cell) { return cell === ""; });

  if (headersMissing) {
    headerRange.setValues([headers]);
    headerRange.setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function getParams_(e) {
  const params = {};
  if (e && e.parameter) {
    Object.keys(e.parameter).forEach(function (k) { params[k] = e.parameter[k]; });
  }
  // Also accept a JSON body if one is sent
  if (e && e.postData && e.postData.contents) {
    try {
      const body = JSON.parse(e.postData.contents);
      Object.keys(body).forEach(function (k) { params[k] = body[k]; });
    } catch (_) {
      // Not JSON - ignore, query/form params already captured
    }
  }
  return params;
}

function clean_(value) {
  return value === undefined || value === null ? "" : String(value).trim();
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
