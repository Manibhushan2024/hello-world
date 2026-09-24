/**
 * Shiprocket label generation — BLANK ROWS ONLY.
 *
 * Reads shipment IDs from column A of "data sheet" and generates labels only
 * for rows where column H (label URL) is still empty. Rows that already have a
 * label are skipped, so you can safely re-run this after partial failures.
 *
 * All names here are prefixed with BL_ / blank so this file can live in the
 * same Apps Script project as the original generateLabels() script without
 * clashing (Apps Script shares one global scope across all .gs files).
 */

const BL_EMAIL = 'placeholder';   // <-- consider moving to PropertiesService
const BL_PASSWORD = 'abc';        // <-- consider moving to PropertiesService

const BL_SHEET_NAME = 'data sheet';
const BL_ID_COL = 1;     // Column A = shipment ID
const BL_LABEL_COL = 8;  // Column H = label URL
const BL_FIRST_ROW = 2;  // Row 1 is the header

/* -------------------- Utilities -------------------- */
function blankLog(msg) {
  Logger.log("[" + new Date().toISOString() + "] " + msg);
}

/* -------------------- Auth -------------------- */
function blankGetBearerToken() {
  blankLog("AUTH: Starting token request");
  var url = 'https://apiv2.shiprocket.in/v1/external/auth/login';
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ email: BL_EMAIL, password: BL_PASSWORD }),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  blankLog("AUTH: HTTP " + response.getResponseCode());
  var text = response.getContentText();

  var json;
  try {
    json = JSON.parse(text);
  } catch (e) {
    blankLog("AUTH: JSON parse error -> " + e);
    throw new Error("Failed to parse auth response");
  }

  if (!json || !json.token) {
    blankLog("AUTH: Missing token in response -> " + text);
    throw new Error("Token not found in auth response");
  }

  blankLog("AUTH: Token obtained successfully");
  return json.token;
}

/* -------------------- Main -------------------- */
function generateLabelsForBlankRows() {
  blankLog("MAIN: Execution started (blank rows only)");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BL_SHEET_NAME);
  if (!sheet) {
    throw new Error("Sheet '" + BL_SHEET_NAME + "' not found");
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < BL_FIRST_ROW) {
    blankLog("MAIN: No data rows. Exiting.");
    return;
  }

  // Read columns A..H in one call so row numbers stay aligned with the sheet.
  var numRows = lastRow - BL_FIRST_ROW + 1;
  var values = sheet.getRange(BL_FIRST_ROW, 1, numRows, BL_LABEL_COL).getValues();

  // Collect only rows that have a shipment ID and no label yet.
  var pending = [];
  for (var r = 0; r < values.length; r++) {
    var shipmentId = String(values[r][BL_ID_COL - 1]).trim();
    var existingLabel = String(values[r][BL_LABEL_COL - 1]).trim();
    if (shipmentId !== "" && existingLabel === "") {
      pending.push({ shipmentId: shipmentId, row: BL_FIRST_ROW + r });
    }
  }

  blankLog("MAIN: Rows needing labels = " + pending.length + " (of " + numRows + " data rows)");
  if (pending.length === 0) {
    blankLog("MAIN: Nothing to do. All rows already have labels.");
    return;
  }

  var token = blankGetBearerToken();
  var successCount = 0;
  var failCount = 0;

  for (var i = 0; i < pending.length; i++) {
    var item = pending[i];
    blankLog("LOOP: [" + (i + 1) + "/" + pending.length + "] Shipment ID = " + item.shipmentId + ", Row = " + item.row);

    try {
      var labelUrl = blankGenerateLabelForShipment(item.shipmentId, token);
      if (labelUrl) {
        sheet.getRange(item.row, BL_LABEL_COL).setValue(labelUrl);
        successCount++;
      } else {
        failCount++; // H stays blank, so the next run will retry this row
      }
    } catch (e) {
      blankLog("LOOP: Error for shipment " + item.shipmentId + " -> " + e);
      failCount++;
    }
  }

  blankLog("MAIN: Completed. Success = " + successCount + ", Failed = " + failCount + ", Attempted = " + pending.length);
}

/* -------------------- Per-shipment -------------------- */
function blankGenerateLabelForShipment(shipmentId, token) {
  var url = 'https://apiv2.shiprocket.in/v1/external/courier/generate/label';
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify({ shipment_id: [shipmentId] }),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  blankLog("SHIPMENT: HTTP " + response.getResponseCode() + " for " + shipmentId);
  var text = response.getContentText();

  var json;
  try {
    json = JSON.parse(text);
  } catch (e) {
    throw new Error("Failed to parse label response for " + shipmentId);
  }

  if (json && json.label_created === 1 && json.label_url) {
    blankLog("SHIPMENT: Label created OK for " + shipmentId + " | URL=" + json.label_url);
    return json.label_url;
  }

  blankLog("SHIPMENT: Label not created for " + shipmentId + " | Response=" + text);
  return null;
}
