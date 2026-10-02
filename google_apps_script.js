/**
 * ============================================================================
 * DISPUTE TRACKING — GOOGLE DRIVE IMAGE SYNC & GOOGLE SHEET MONTH-WISE SYNC
 * Main Drive Folder ID : 15c6PLPDtr9I3F8POfHUdwYv-ttdFtesA
 * Google Spreadsheet ID: 1e9OyvaLMVyRh84adrx-zxuWwz2ZmT-WddZB6VY3fw7Y
 * Unique Row Key       : F & G & H & J (Invoice No. | Order id | Item SKU | AWB No.)
 * ============================================================================
 *
 * HOW TO UPDATE / DEPLOY:
 * 1. Open https://script.google.com (your existing project or New Project).
 * 2. Paste this entire code.
 * 3. Click "Deploy" -> "New deployment" (or "Manage deployments" -> Edit -> New version).
 * 4. Type: "Web app", Execute as: "Me", Who has access: "Anyone" -> Click "Deploy".
 * 5. Authorize Drive + Google Sheets permissions and copy the Web app URL.
 */

const MAIN_FOLDER_ID = "15c6PLPDtr9I3F8POfHUdwYv-ttdFtesA";
const SPREADSHEET_ID = "1e9OyvaLMVyRh84adrx-zxuWwz2ZmT-WddZB6VY3fw7Y";

/**
 * Select "authorizePermissions" in the top bar and click "Run" once
 * to grant Google Sheets + Google Drive permissions!
 */
function authorizePermissions() {
  var folder = DriveApp.getFolderById(MAIN_FOLDER_ID);
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log("Authorized Drive Folder: " + folder.getName());
  Logger.log("Authorized Google Sheet: " + ss.getName());
}

function buildFileName(awb, courier) {
  var cleanAwb = String(awb || "").trim().replace(/[<>:"/\\|?*]+/g, "_");
  var cleanCourier = String(courier || "").trim().replace(/[<>:"/\\|?*]+/g, "_");
  if (cleanCourier && cleanAwb.toLowerCase().indexOf("_" + cleanCourier.toLowerCase()) === -1) {
    return cleanAwb + "_" + cleanCourier + ".png";
  }
  return cleanAwb + ".png";
}

/**
 * Builds composite key F&G&H&J from row values array
 */
function buildFGHJKey(rowValues, headers) {
  var idxF = 5, idxG = 6, idxH = 7, idxJ = 9;
  if (headers && headers.length) {
    for (var i = 0; i < headers.length; i++) {
      var hl = String(headers[i] || "").trim().toLowerCase();
      if (hl.indexOf("invoice") !== -1) idxF = i;
      else if (hl.indexOf("order id") !== -1 || hl === "order") idxG = i;
      else if (hl.indexOf("sku") !== -1) idxH = i;
      else if (hl.indexOf("awb") !== -1 || hl.indexOf("tracking id") !== -1) idxJ = i;
    }
  }
  var cleanVal = function(v) {
    var s = String(v || "").trim();
    if (s.endsWith(".0")) s = s.substring(0, s.length - 2);
    return s.toLowerCase();
  };
  var f = rowValues.length > idxF ? cleanVal(rowValues[idxF]) : "";
  var g = rowValues.length > idxG ? cleanVal(rowValues[idxG]) : "";
  var h = rowValues.length > idxH ? cleanVal(rowValues[idxH]) : "";
  var j = rowValues.length > idxJ ? cleanVal(rowValues[idxJ]) : "";
  return f + "|" + g + "|" + h + "|" + j;
}

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || "ping";

    if (action === "check") {
      const awb = (e.parameter.awb || "").trim();
      const courier = (e.parameter.courier || "").trim();
      const channel = (e.parameter.channel || "").trim();
      if (!awb) {
        return jsonResponse({ ok: false, error: "Missing awb parameter" });
      }
      const found = findAwbImageInDrive(awb, courier, channel);
      return jsonResponse({ ok: true, found: !!found, data: found });
    }

    if (action === "get_sheets_data") {
      return jsonResponse(readAllMonthSheetsFromSpreadsheet());
    }

    return jsonResponse({
      ok: true,
      message: "Dispute Tracking Drive & Sheet Sync Apps Script is active.",
      folderId: MAIN_FOLDER_ID,
      spreadsheetId: SPREADSHEET_ID
    });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || "{}");
    const action = payload.action || "upload";

    if (action === "check") {
      const found = findAwbImageInDrive(payload.awb, payload.courier, payload.channel);
      return jsonResponse({ ok: true, found: !!found, data: found });
    }

    if (action === "get_sheets_data") {
      return jsonResponse(readAllMonthSheetsFromSpreadsheet());
    }

    if (action === "save_sheet_rows") {
      return jsonResponse(saveRowsToGoogleSheet(payload));
    }

    if (action === "upload") {
      const awb = String(payload.awb || "").trim();
      const courier = String(payload.courier || "").trim();
      const channel = String(payload.channel || "General").trim() || "General";
      const base64Data = payload.base64Data || "";
      const imageUrl = payload.imageUrl || "";

      if (!awb) {
        return jsonResponse({ ok: false, error: "AWB number is required" });
      }

      const fileName = payload.fileName ? String(payload.fileName).trim() : buildFileName(awb, courier);
      const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
      const channelFolder = getOrCreateChannelFolder(mainFolder, channel);

      const existingInChannel = findFileByAwbInFolder(channelFolder, payload.rawAwb || awb, fileName);
      if (existingInChannel) {
        return jsonResponse({
          ok: true,
          status: "already_exists",
          awb: payload.rawAwb || awb,
          courier: courier,
          fileName: existingInChannel.getName(),
          channel: channelFolder.getName(),
          fileId: existingInChannel.getId(),
          viewUrl: "https://drive.google.com/file/d/" + existingInChannel.getId() + "/view",
          directUrl: "https://drive.google.com/uc?export=view&id=" + existingInChannel.getId()
        });
      }

      var blob = null;
      if (base64Data) {
        var decoded = Utilities.base64Decode(base64Data);
        blob = Utilities.newBlob(decoded, "image/png", fileName);
      } else if (imageUrl) {
        var response = UrlFetchApp.fetch(imageUrl, { muteHttpExceptions: true });
        if (response.getResponseCode() !== 200) {
          return jsonResponse({
            ok: false,
            error: "Failed to fetch image URL, HTTP " + response.getResponseCode()
          });
        }
        blob = response.getBlob().setName(fileName);
      } else {
        return jsonResponse({ ok: false, error: "Neither base64Data nor imageUrl provided" });
      }

      const createdFile = channelFolder.createFile(blob);
      try {
        createdFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (shareErr) {}

      const fileId = createdFile.getId();
      return jsonResponse({
        ok: true,
        status: "uploaded",
        awb: payload.rawAwb || awb,
        courier: courier,
        fileName: fileName,
        channel: channelFolder.getName(),
        fileId: fileId,
        viewUrl: "https://drive.google.com/file/d/" + fileId + "/view",
        directUrl: "https://drive.google.com/uc?export=view&id=" + fileId
      });
    }

    return jsonResponse({ ok: false, error: "Unknown action: " + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

/**
 * ============================================================================
 * GOOGLE SPREADSHEET MONTH-YEAR UPSERT & READ FUNCTIONS
 * ============================================================================
 */
function getOrCreateMonthSheet(ss, targetSheetName, headers) {
  var sheet = ss.getSheetByName(targetSheetName);

  // If sheet like "September" exists and target is "September-2026", rename it!
  if (!sheet && targetSheetName.indexOf("-") !== -1) {
    var parts = targetSheetName.split("-");
    var monthOnly = parts[0];
    var shortMonth = monthOnly.substring(0, 3) + "-" + parts[1]; // e.g., Sep-2026
    var legacySheet = ss.getSheetByName(monthOnly);
    if (legacySheet) {
      legacySheet.setName(targetSheetName);
      sheet = legacySheet;
    }
  }

  if (!sheet) {
    sheet = ss.insertSheet(targetSheetName);
  }

  // Ensure headers exist in Row 1
  if (sheet.getLastRow() === 0 && headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#e2ecfe");
    headerRange.setFontColor("#111827");
    headerRange.setFontWeight("normal");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function saveRowsToGoogleSheet(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var headers = payload.headers || [];
  var rowsBySheet = payload.rows_by_sheet || {};
  var totalInserted = 0;
  var totalUpdated = 0;
  var touchedSheets = [];

  var imageColIdx = -1;
  var ticketColIdx = -1;
  var ticketStatusColIdx = -1;
  var reasonColIdx = -1;
  var remarkColIdx = -1;
  var orderColIdx = -1;

  for (var c = 0; c < headers.length; c++) {
    var hl = String(headers[c] || "").toLowerCase().trim();
    if (hl.indexOf("image") !== -1 || hl.indexOf("photo") !== -1) imageColIdx = c;
    if (hl.indexOf("ticket status") !== -1) ticketStatusColIdx = c;
    else if (hl.indexOf("ticket") !== -1 && ticketColIdx === -1) ticketColIdx = c;
    if (hl.indexOf("dispute") !== -1 || hl.indexOf("reason") !== -1) reasonColIdx = c;
    if (hl.indexOf("remark") !== -1) remarkColIdx = c;
    if (hl.indexOf("order id") !== -1 || hl === "order") orderColIdx = c;
  }

  for (var sheetName in rowsBySheet) {
    var incomingRows = rowsBySheet[sheetName] || [];
    if (incomingRows.length === 0) continue;

    var sheet = getOrCreateMonthSheet(ss, sheetName, headers);
    touchedSheets.push(sheetName);

    var lastRow = sheet.getLastRow();
    var lastCol = Math.max(sheet.getLastColumn(), headers.length);

    // If sheet headers are narrower than incoming headers, update header row
    if (headers.length > sheet.getLastColumn()) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    var existingData = [];
    var keyToRowIndex = {}; // key -> 1-based sheet row number (row 2+)
    var ticketToRowIndex = {};
    var orderToRowIndex = {};

    if (lastRow >= 2) {
      existingData = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
      for (var r = 0; r < existingData.length; r++) {
        var rKey = buildFGHJKey(existingData[r], headers);
        if (rKey && rKey !== "|||") {
          keyToRowIndex[rKey] = r + 2;
        }
        if (ticketColIdx !== -1 && existingData[r][ticketColIdx]) {
          var tVal = String(existingData[r][ticketColIdx]).trim().toLowerCase();
          if (tVal) ticketToRowIndex[tVal] = r + 2;
        }
        if (orderColIdx !== -1 && existingData[r][orderColIdx]) {
          var oVal = String(existingData[r][orderColIdx]).trim().toLowerCase();
          if (oVal) orderToRowIndex[oVal] = r + 2;
        }
      }
    }

    for (var i = 0; i < incomingRows.length; i++) {
      var item = incomingRows[i];
      var vals = (item.values || []).slice();
      while (vals.length < headers.length) vals.push("");

      var rowKey = item.key ? String(item.key).toLowerCase() : buildFGHJKey(vals, headers);
      var itemTicket = item.ticket_id ? String(item.ticket_id).trim().toLowerCase() : (ticketColIdx !== -1 ? String(vals[ticketColIdx] || "").trim().toLowerCase() : "");
      var itemOrder = item.order_id ? String(item.order_id).trim().toLowerCase() : (orderColIdx !== -1 ? String(vals[orderColIdx] || "").trim().toLowerCase() : "");
      var imgUrl = item.image_url || "";

      // Match target row: by ticket_id first, then order_id, then composite key
      var targetRowNum = null;
      if (itemTicket && ticketToRowIndex[itemTicket]) {
        targetRowNum = ticketToRowIndex[itemTicket];
      } else if (itemOrder && orderToRowIndex[itemOrder]) {
        targetRowNum = orderToRowIndex[itemOrder];
      } else if (keyToRowIndex[rowKey]) {
        targetRowNum = keyToRowIndex[rowKey];
      }

      if (targetRowNum) {
        // UPDATE EXISTING ROW
        var oldRow = existingData[targetRowNum - 2] || [];

        // Preserve existing Ticket No / Dispute Reason if incoming is blank
        if (ticketColIdx !== -1 && !String(vals[ticketColIdx] || "").trim() && oldRow[ticketColIdx]) {
          vals[ticketColIdx] = oldRow[ticketColIdx];
        }
        if (reasonColIdx !== -1 && !String(vals[reasonColIdx] || "").trim() && oldRow[reasonColIdx]) {
          vals[reasonColIdx] = oldRow[reasonColIdx];
        }
        // Preserve existing Ticket Status / Remark only if incoming is blank
        if (ticketStatusColIdx !== -1 && !String(vals[ticketStatusColIdx] || "").trim() && oldRow[ticketStatusColIdx]) {
          vals[ticketStatusColIdx] = oldRow[ticketStatusColIdx];
        }
        if (remarkColIdx !== -1 && !String(vals[remarkColIdx] || "").trim() && oldRow[remarkColIdx]) {
          vals[remarkColIdx] = oldRow[remarkColIdx];
        }

        sheet.getRange(targetRowNum, 1, 1, vals.length).setValues([vals]);
        if (imageColIdx !== -1 && imgUrl && imgUrl.indexOf("http") === 0) {
          sheet
            .getRange(targetRowNum, imageColIdx + 1)
            .setFormula('=HYPERLINK("' + imgUrl.replace(/"/g, '""') + '", "View Image")');
        }
        totalUpdated++;
      } else {
        // INSERT NEW ROW
        var newRowNum = sheet.getLastRow() + 1;
        sheet.getRange(newRowNum, 1, 1, vals.length).setValues([vals]);
        if (imageColIdx !== -1 && imgUrl && imgUrl.indexOf("http") === 0) {
          sheet
            .getRange(newRowNum, imageColIdx + 1)
            .setFormula('=HYPERLINK("' + imgUrl.replace(/"/g, '""') + '", "View Image")');
        }
        keyToRowIndex[rowKey] = newRowNum;
        if (itemTicket) ticketToRowIndex[itemTicket] = newRowNum;
        if (itemOrder) orderToRowIndex[itemOrder] = newRowNum;
        existingData.push(vals);
        totalInserted++;
      }
    }
  }

  return {
    ok: true,
    inserted: totalInserted,
    updated: totalUpdated,
    sheets: touchedSheets
  };
}

function readAllMonthSheetsFromSpreadsheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var allSheets = ss.getSheets();
  var monthPattern = /^(January|February|March|April|May|June|July|August|September|October|November|December)-\d{4}$/i;

  var resultSheets = {};
  var sheetNames = [];

  for (var s = 0; s < allSheets.length; s++) {
    var sh = allSheets[s];
    var sName = sh.getName();

    // Auto-rename "September" -> "September-2026" if found
    if (/^(January|February|March|April|May|June|July|August|September|October|November|December)$/i.test(sName.trim())) {
      sName = sName.trim() + "-" + new Date().getFullYear();
      sh.setName(sName);
    }

    if (!monthPattern.test(sName.trim())) {
      continue;
    }

    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 1 || lastCol < 1) continue;

    var displayVals = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    var formulas = sh.getRange(1, 1, lastRow, lastCol).getFormulas();
    var headers = displayVals[0];
    var rows = [];

    for (var r = 1; r < displayVals.length; r++) {
      var rVals = displayVals[r];
      var rForms = formulas[r];
      var hasContent = false;
      var links = {};

      for (var c = 0; c < rVals.length; c++) {
        if (String(rVals[c] || "").trim() !== "") hasContent = true;
        var fStr = String(rForms[c] || "");
        if (fStr.toUpperCase().indexOf("=HYPERLINK") === 0) {
          var m = fStr.match(/=HYPERLINK\(\s*"([^"]+)"/i);
          if (m && m[1]) links[String(c)] = m[1];
        }
      }
      if (!hasContent) continue;

      rows.push({
        _row_num: rows.length + 1,
        values: rVals,
        links: links,
        key: buildFGHJKey(rVals, headers)
      });
    }

    sheetNames.push(sName);
    resultSheets[sName] = {
      headers: headers,
      rows: rows,
      total_rows: rows.length,
      total_columns: headers.length
    };
  }

  return {
    ok: true,
    filename: "Google Sheet (Saved Dispute Data)",
    sheet_names: sheetNames,
    active_sheet: sheetNames.length > 0 ? sheetNames[sheetNames.length - 1] : null,
    sheets: resultSheets
  };
}

/**
 * ============================================================================
 * GOOGLE DRIVE IMAGE HELPER FUNCTIONS
 * ============================================================================
 */
function getOrCreateChannelFolder(mainFolder, channelName) {
  const cleanName = String(channelName || "General").trim() || "General";
  const targetLower = cleanName.toLowerCase();

  const subFolders = mainFolder.getFolders();
  while (subFolders.hasNext()) {
    var folder = subFolders.next();
    if (folder.getName().trim().toLowerCase() === targetLower) {
      return folder;
    }
  }
  return mainFolder.createFolder(cleanName);
}

function findFileByAwbInFolder(folder, awb, exactFileName) {
  if (exactFileName) {
    var exactFiles = folder.getFilesByName(exactFileName);
    if (exactFiles.hasNext()) {
      return exactFiles.next();
    }
  }
  var cleanAwb = String(awb || "").trim().toLowerCase();
  if (!cleanAwb) return null;

  var allFiles = folder.getFiles();
  while (allFiles.hasNext()) {
    var f = allFiles.next();
    var nameLower = f.getName().trim().toLowerCase();
    if (nameLower === cleanAwb + ".png" || nameLower.indexOf(cleanAwb + "_") === 0) {
      return f;
    }
  }
  return null;
}

function buildChannelFolderMap(mainFolder) {
  const map = {};
  const subFolders = mainFolder.getFolders();
  while (subFolders.hasNext()) {
    var f = subFolders.next();
    map[f.getName().trim().toLowerCase()] = f;
  }
  return map;
}

function findAwbImageInDrive(awb, courier, preferredChannel) {
  const exactFileName = buildFileName(awb, courier);
  const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
  const folderMap = buildChannelFolderMap(mainFolder);

  if (preferredChannel) {
    var key = preferredChannel.trim().toLowerCase();
    if (folderMap[key]) {
      var f = findFileByAwbInFolder(folderMap[key], awb, exactFileName);
      if (f) return formatFileInfo(f, folderMap[key].getName());
    }
  }
  for (var k in folderMap) {
    var file = findFileByAwbInFolder(folderMap[k], awb, exactFileName);
    if (file) return formatFileInfo(file, folderMap[k].getName());
  }
  var rootFile = findFileByAwbInFolder(mainFolder, awb, exactFileName);
  if (rootFile) return formatFileInfo(rootFile, mainFolder.getName());
  return null;
}

function formatFileInfo(file, channelName) {
  var id = file.getId();
  return {
    fileId: id,
    fileName: file.getName(),
    channel: channelName,
    viewUrl: "https://drive.google.com/file/d/" + id + "/view",
    directUrl: "https://drive.google.com/uc?export=view&id=" + id
  };
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
