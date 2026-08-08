var HUD_WA_SHEET_ID = "1i_3360NECjPHsd687Uts4i3xmD4h67xu--PMazmKlw0";
var HUD_WA_INPUT_TAB = "Input";
var HUD_WA_REP_TAB = "Reparaturauftrag";
var HUD_NB_SHEET_ID = "1VGCAHUbOPgsInQICA1GnrtKg1EPK1d1zWB-GkLi6iVE";
var HUD_NB_TAB = "Input Exit";
var HUD_DRIVE_FOLDER_ID = "1yEwrUVpS2nDA9qL3ht0p_XF7S0M-RGH2";
var HUD_CAROL_BASE = "https://carol.autohero.com/en-GB/refurbishment?rsv=";
var HUD_PETRONAS_URL = "https://de.pli-petronas.com/de/schmierstoffe";
var HUD_CAROL_CACHE_PREFIX = "carol_data_";
var HUD_CAROL_CACHE_SECONDS = 21600;

var HUD_NB_ALIASES = {
  stockId: ["stock id", "stockid"],
  grund: ["grund"],
  art: ["art der nachbestellung", "art", "typ"],
  bearbeiter: ["bearbeiter"],
  status: ["status"],
  mappe: ["werkstattmappe erstellt", "werkstattmappe"]
};

function doGet(e) {
  var t = HtmlService.createTemplateFromFile("HUD_App");
  return t.evaluate()
    .setTitle("EXIT Werkstattmappe HUD")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function doPost(e) {
  var out = { success: false };
  try {
    var body = JSON.parse(e.postData.contents || "{}");
    var secret = String(PropertiesService.getScriptProperties().getProperty("CAROL_BRIDGE_SECRET") || "");
    if (!secret || String(body.secret || "") !== secret) {
      out.message = "Ungültiges Secret";
      return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
    }
    var entries = body.entries || [];
    if (body.stockId) entries.push({ stockId: body.stockId, modell: body.modell, vin: body.vin });
    var cache = CacheService.getScriptCache();
    var stored = 0;
    for (var i = 0; i < entries.length; i++) {
      var sid = hudNormalizeStockId_(entries[i].stockId);
      if (!sid) continue;
      cache.put(HUD_CAROL_CACHE_PREFIX + sid, JSON.stringify({
        modell: String(entries[i].modell || "").trim(),
        vin: String(entries[i].vin || "").trim(),
        ts: Date.now()
      }), HUD_CAROL_CACHE_SECONDS);
      stored++;
    }
    out.success = true;
    out.stored = stored;
  } catch (err) {
    out.message = String(err.message || err);
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function hudNormalizeStockId_(v) {
  return String(v || "").toUpperCase().replace(/\s+/g, "").trim();
}

function hudNormHeader_(v) {
  return String(v || "").toLowerCase().replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
}

function hudGetNbLayout_(sheet) {
  var lastCol = Math.max(1, sheet.getLastColumn());
  var scanRows = Math.min(10, Math.max(1, sheet.getLastRow()));
  var rows = sheet.getRange(1, 1, scanRows, lastCol).getValues();
  var keys = Object.keys(HUD_NB_ALIASES);
  var best = null;
  for (var r = 0; r < rows.length; r++) {
    var normToCol = {};
    for (var c = 0; c < lastCol; c++) {
      var n = hudNormHeader_(rows[r][c]);
      if (n && normToCol[n] === undefined) normToCol[n] = c + 1;
    }
    var cols = {};
    var score = 0;
    for (var k = 0; k < keys.length; k++) {
      var aliases = HUD_NB_ALIASES[keys[k]];
      for (var a = 0; a < aliases.length; a++) {
        if (normToCol[aliases[a]]) {
          cols[keys[k]] = normToCol[aliases[a]];
          score++;
          break;
        }
      }
    }
    if (!best || score > best.score) {
      best = { score: score, headerRow: r + 1, cols: cols };
    }
    if (score === keys.length) break;
  }
  if (best && !best.cols.mappe) {
    var hdr = rows[best.headerRow - 1];
    for (var mc = 0; mc < lastCol; mc++) {
      if (hudNormHeader_(hdr[mc]).indexOf("werkstattmappe") !== -1) {
        best.cols.mappe = mc + 1;
        break;
      }
    }
  }
  if (!best || !best.cols.stockId || !best.cols.mappe) {
    throw new Error("Nachbestellung-Header nicht gefunden (Stock ID / Werkstattmappe erstellt)");
  }
  return { headerRow: best.headerRow, dataStartRow: best.headerRow + 1, cols: best.cols, lastCol: lastCol };
}

function hudIsChecked_(v) {
  if (v === true) return true;
  var s = String(v || "").toLowerCase().trim();
  return s === "true" || s === "wahr" || s === "x" || s === "ja";
}

function getOpenMappen() {
  try {
    var sheet = SpreadsheetApp.openById(HUD_NB_SHEET_ID).getSheetByName(HUD_NB_TAB);
    if (!sheet) return { success: false, message: "Tab '" + HUD_NB_TAB + "' nicht gefunden", entries: [] };
    var layout = hudGetNbLayout_(sheet);
    var lastRow = sheet.getLastRow();
    if (lastRow < layout.dataStartRow) return { success: true, entries: [] };
    var numRows = lastRow - layout.dataStartRow + 1;
    var data = sheet.getRange(layout.dataStartRow, 1, numRows, layout.lastCol).getValues();
    var cols = layout.cols;
    var byStock = {};
    var order = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var stockId = hudNormalizeStockId_(row[cols.stockId - 1]);
      if (!stockId) continue;
      var art = cols.art ? String(row[cols.art - 1] || "").trim() : "";
      if (hudIsChecked_(row[cols.mappe - 1])) continue;
      if (!byStock[stockId]) {
        byStock[stockId] = {
          stockId: stockId,
          carolUrl: HUD_CAROL_BASE + encodeURIComponent(stockId),
          rows: []
        };
        order.push(stockId);
      }
      byStock[stockId].rows.push({
        sheetRow: layout.dataStartRow + i,
        grund: cols.grund ? String(row[cols.grund - 1] || "").trim() : "",
        art: art,
        bearbeiter: cols.bearbeiter ? String(row[cols.bearbeiter - 1] || "").trim() : "",
        status: cols.status ? String(row[cols.status - 1] || "").trim() : ""
      });
    }
    var cache = CacheService.getScriptCache();
    var entries = [];
    for (var o = 0; o < order.length; o++) {
      var ent = byStock[order[o]];
      var cached = cache.get(HUD_CAROL_CACHE_PREFIX + ent.stockId);
      if (cached) {
        try {
          var cd = JSON.parse(cached);
          ent.carolModell = cd.modell || "";
          ent.carolVin = cd.vin || "";
        } catch (e2) {}
      }
      entries.push(ent);
    }
    return { success: true, entries: entries, petronasUrl: HUD_PETRONAS_URL };
  } catch (err) {
    return { success: false, message: String(err.message || err), entries: [] };
  }
}

function getInputMeta() {
  try {
    var sheet = SpreadsheetApp.openById(HUD_WA_SHEET_ID).getSheetByName(HUD_WA_INPUT_TAB);
    if (!sheet) return { success: false, message: "Tab '" + HUD_WA_INPUT_TAB + "' nicht gefunden" };
    var aVals = sheet.getRange("A10:A41").getValues();
    var cVals = sheet.getRange("C29:C41").getValues();
    var dVals = sheet.getRange("D2:D41").getValues();
    var workItems = [];
    var oils = [];
    var dItems = [];
    for (var i = 0; i < aVals.length; i++) {
      var rowNum = 10 + i;
      var label = String(aVals[i][0] || "").trim();
      if (!label) continue;
      if (rowNum === 13) continue;
      if (rowNum <= 28) {
        workItems.push({ row: rowNum, label: label });
      } else {
        oils.push({ row: rowNum, label: label, nr: String(cVals[rowNum - 29][0] || "").trim() });
      }
    }
    for (var d = 0; d < dVals.length; d++) {
      var dRow = 2 + d;
      var dLabel = String(dVals[d][0] || "").trim();
      if (!dLabel) continue;
      if ((dRow >= 24 && dRow <= 28) || (dRow >= 32 && dRow <= 36)) continue;
      if (dLabel.toLowerCase().indexOf("sonstige") === 0) continue;
      dItems.push({ row: dRow, label: dLabel });
    }
    return { success: true, workItems: workItems, oils: oils, dItems: dItems };
  } catch (err) {
    return { success: false, message: String(err.message || err) };
  }
}

function getCarolData(stockId) {
  var sid = hudNormalizeStockId_(stockId);
  if (!sid) return { success: false };
  var cached = CacheService.getScriptCache().get(HUD_CAROL_CACHE_PREFIX + sid);
  if (!cached) return { success: true, found: false };
  try {
    var cd = JSON.parse(cached);
    return { success: true, found: true, modell: cd.modell || "", vin: cd.vin || "" };
  } catch (e) {
    return { success: true, found: false };
  }
}

function hudResetInput_(sheet) {
  sheet.getRange("B2:B45").clearContent();
  sheet.getRange("D24:D28").clearContent();
  sheet.getRange("D32:D36").clearContent();
  sheet.getRange("E2:E45").clearContent();
  sheet.getRange("H2:H45").clearContent();
  sheet.getRange("J2:O45").clearContent();
  sheet.getRange("Q2:V45").clearContent();
}

function hudApplyOilToReparaturauftrag_(ss) {
  var inputSheet = ss.getSheetByName(HUD_WA_INPUT_TAB);
  var outputSheet = ss.getSheetByName(HUD_WA_REP_TAB);
  var xRange = inputSheet.getRange("B29:B41").getValues();
  var contentRange = inputSheet.getRange("C29:C41").getValues();
  var values = [];
  for (var i = 0; i < xRange.length; i++) {
    if (xRange[i][0] === "x") values.push(contentRange[i][0]);
  }
  outputSheet.getRange("G9:G12").clearContent().clearFormat();
  var startIndex = Math.floor((4 - values.length) / 2);
  for (var v = 0; v < values.length; v++) {
    var cell = outputSheet.getRange(startIndex + v + 9, 7);
    cell.setValue(values[v]);
    cell.setFontFamily("Arial");
    cell.setFontSize(33);
    cell.setHorizontalAlignment("center");
    cell.setVerticalAlignment("middle");
  }
}

function hudExportRepPdf_(ss) {
  var repSheet = ss.getSheetByName(HUD_WA_REP_TAB);
  if (!repSheet) return { success: false, message: "Tab '" + HUD_WA_REP_TAB + "' nicht gefunden" };
  var exportUrl = "https://docs.google.com/spreadsheets/d/" + HUD_WA_SHEET_ID + "/export?exportFormat=pdf&format=pdf"
    + "&gid=" + repSheet.getSheetId()
    + "&portrait=true&size=A4&fitw=true&gridlines=false"
    + "&top_margin=0.3&bottom_margin=0.3&left_margin=0.3&right_margin=0.3"
    + "&sheetnames=false&printtitle=false&pagenumbers=false";
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true,
    followRedirects: true
  });
  var code = response.getResponseCode();
  var blob = response.getBlob();
  var bytes = blob.getBytes();
  var isPdf = bytes && bytes.length >= 4 && bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46;
  if (code !== 200 || !isPdf) {
    return { success: false, message: "PDF-Export fehlgeschlagen (HTTP " + code + ")" };
  }
  return { success: true, blob: blob, bytes: bytes };
}

function hudMarkMappeErstellt_(stockId) {
  var sheet = SpreadsheetApp.openById(HUD_NB_SHEET_ID).getSheetByName(HUD_NB_TAB);
  if (!sheet) return 0;
  var layout = hudGetNbLayout_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < layout.dataStartRow) return 0;
  var numRows = lastRow - layout.dataStartRow + 1;
  var data = sheet.getRange(layout.dataStartRow, 1, numRows, layout.lastCol).getValues();
  var cols = layout.cols;
  var marked = 0;
  for (var i = 0; i < data.length; i++) {
    var sid = hudNormalizeStockId_(data[i][cols.stockId - 1]);
    if (sid !== stockId) continue;
    if (hudIsChecked_(data[i][cols.mappe - 1])) continue;
    sheet.getRange(layout.dataStartRow + i, cols.mappe).setValue(true);
    marked++;
  }
  return marked;
}

function createMappe(payload) {
  var lock = LockService.getScriptLock();
  var gotLock = false;
  try {
    gotLock = lock.tryLock(45000);
    if (!gotLock) return { success: false, message: "Werkstattauftrag gerade in Arbeit – bitte kurz erneut versuchen" };

    payload = payload || {};
    var stockId = hudNormalizeStockId_(payload.stockId);
    var modell = String(payload.modell || "").trim();
    var vin = String(payload.vin || "").trim();
    if (!stockId) return { success: false, message: "Keine Stock-ID" };
    if (!modell) return { success: false, message: "Modell fehlt" };
    if (!vin) return { success: false, message: "VIN fehlt" };

    var workRows = payload.workRows || [];
    var dRows = payload.dRows || [];
    var freeMech = payload.freeMech || [];
    var freeParts = payload.freeParts || [];
    var oilRow = parseInt(payload.oilRow, 10) || 0;
    var liters = String(payload.liters || "").replace(",", ".").trim();
    var litersNum = liters ? parseFloat(liters) : 0;

    if (!workRows.length && !dRows.length && !freeMech.length && !freeParts.length) {
      return { success: false, message: "Keine Arbeiten ausgewählt" };
    }

    var ss = SpreadsheetApp.openById(HUD_WA_SHEET_ID);
    var inputSheet = ss.getSheetByName(HUD_WA_INPUT_TAB);
    if (!inputSheet) return { success: false, message: "Tab '" + HUD_WA_INPUT_TAB + "' nicht gefunden" };

    var oilRequired = false;
    var wLabels = inputSheet.getRange("A10:A28").getValues();
    for (var w = 0; w < workRows.length; w++) {
      var wr = parseInt(workRows[w], 10);
      if (wr >= 10 && wr <= 28) {
        var lbl = String(wLabels[wr - 10][0] || "").toLowerCase();
        if (lbl.indexOf("motorölwechsel") !== -1 || lbl.indexOf("motoroelwechsel") !== -1) oilRequired = true;
      }
    }
    if (oilRequired) {
      if (!(oilRow >= 29 && oilRow <= 41)) return { success: false, message: "Bitte Öl auswählen" };
      var oilLabel = String(inputSheet.getRange(oilRow, 1).getValue() || "").toLowerCase();
      if (oilLabel.indexOf("mitgeliefert") === -1 && !(litersNum > 0)) {
        return { success: false, message: "Bitte Füllmenge (Liter) angeben" };
      }
    }

    hudResetInput_(inputSheet);

    inputSheet.getRange("B2").setValue(stockId);
    inputSheet.getRange("B3").setValue(modell);
    inputSheet.getRange("B4").setValue(vin);

    for (var i = 0; i < workRows.length; i++) {
      var r = parseInt(workRows[i], 10);
      if (r >= 10 && r <= 28 && r !== 13) inputSheet.getRange(r, 2).setValue("x");
    }
    if (litersNum > 0) inputSheet.getRange("B13").setValue(litersNum);
    if (oilRow >= 29 && oilRow <= 41) inputSheet.getRange(oilRow, 2).setValue("x");

    for (var d = 0; d < dRows.length; d++) {
      var dr = parseInt(dRows[d], 10);
      if (dr >= 2 && dr <= 41) inputSheet.getRange(dr, 5).setValue("x");
    }
    for (var m = 0; m < freeMech.length && m < 5; m++) {
      var txtM = String(freeMech[m] || "").trim();
      if (txtM) inputSheet.getRange(24 + m, 4).setValue(txtM);
    }
    for (var p = 0; p < freeParts.length && p < 5; p++) {
      var txtP = String(freeParts[p] || "").trim();
      if (txtP) inputSheet.getRange(32 + p, 4).setValue(txtP);
    }

    hudApplyOilToReparaturauftrag_(ss);
    SpreadsheetApp.flush();
    Utilities.sleep(1500);

    var pdf = hudExportRepPdf_(ss);
    if (!pdf.success) return { success: false, message: pdf.message };

    var fileName = stockId + " Werkstattauftrag EXIT.pdf";
    var pdfUrl = "";
    var driveMsg = "";
    try {
      var folder = DriveApp.getFolderById(HUD_DRIVE_FOLDER_ID);
      var file = folder.createFile(pdf.blob.setName(fileName));
      pdfUrl = file.getUrl();
    } catch (driveErr) {
      driveMsg = "Drive-Ablage fehlgeschlagen: " + String(driveErr.message || driveErr);
    }

    var marked = 0;
    var nbMsg = "";
    try {
      marked = hudMarkMappeErstellt_(stockId);
    } catch (nbErr) {
      nbMsg = "Nachbestellung-Checkbox fehlgeschlagen: " + String(nbErr.message || nbErr);
    }

    var msgs = ["Werkstattmappe für " + stockId + " erstellt"];
    if (pdfUrl) msgs.push("PDF in Drive gespeichert");
    if (driveMsg) msgs.push(driveMsg);
    if (marked > 0) msgs.push(marked + " Nachbestellung(en) abgehakt");
    if (nbMsg) msgs.push(nbMsg);

    return {
      success: true,
      message: msgs.join(" | "),
      printB64: Utilities.base64Encode(pdf.bytes),
      pdfUrl: pdfUrl,
      fileName: fileName,
      markedRows: marked
    };
  } catch (err) {
    return { success: false, message: String(err.message || err) };
  } finally {
    if (gotLock) {
      try { lock.releaseLock(); } catch (e) {}
    }
  }
}
