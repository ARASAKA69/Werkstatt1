var KVA_WA_VERSION = "1.5.0";
var KVA_WA_SHEET_ID = "1oVUT3aQh1mdm2PgZV04EN_PUHrp4T5SHCNwcNKNGQ8g";
var KVA_WA_INPUT_TAB = "Input";
var KVA_WA_REP_TAB = "Reparaturauftrag";
var KVA_WA_DRIVE_FOLDER_ID = "1SLGE3ZXL7PFQ4dwwxGX9tqr_vSoAz98I";
var KVA_WA_DRIVE_FOLDER_NAME = "AG - Werkstatt Aufträge";
var KVA_NACHTRAG_SHEET_ID = "1NHhS7vSJB0S9NwqoCducgjw3pcIulCQlEO_JpLg-v4g";
var KVA_NACHTRAG_TAB = "Tabellenblatt1";
var KVA_BRIDGE_SECRET = "ARASAKA69";
var KVA_META_CACHE_KEY = "kva_wa_input_meta_v3";
var KVA_META_CACHE_TTL = 1800;

function kvaGetSpreadsheet_(optSs) {
  if (optSs) return optSs;
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getSheetByName(KVA_WA_INPUT_TAB)) return active;
  } catch (e) {}
  return SpreadsheetApp.openById(KVA_WA_SHEET_ID);
}

function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("KVA Werkstattauftrag")
      .addItem("KVA HUD öffnen", "openKvaAuftragHud")
      .addSeparator()
      .addItem("Nachtrag exportieren (Übertrag)", "exportNachtrag")
      .addItem("Input zurücksetzen (Reset)", "resetInput")
      .addSeparator()
      .addItem("Berechtigungen prüfen", "authorizeKvaScopes")
      .addToUi();
  } catch (e) {}
}

function openKvaAuftragHud() {
  var url = "";
  try { url = ScriptApp.getService().getUrl() || ""; } catch (e) { url = ""; }
  if (!url) {
    SpreadsheetApp.getUi().alert("Bitte die Web-App zuerst über 'Bereitstellen > Neue Bereitstellung' veröffentlichen.");
    return;
  }
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><body style="font-family:Segoe UI,sans-serif;padding:14px;background:#0d1117;color:#e6edf3;">'
    + '<p>KVA Werkstattauftrag HUD wird geöffnet…</p>'
    + '<p><a href="' + url + '" target="_blank" rel="noopener" style="color:#f2cc60;">Hier klicken, falls sich kein Tab öffnet</a></p>'
    + '<script>window.open(' + JSON.stringify(url) + ',"_blank");google.script.host.close();<\/script>'
    + '</body></html>'
  ).setWidth(360).setHeight(140);
  SpreadsheetApp.getUi().showModalDialog(html, "KVA Werkstattauftrag HUD");
}

function authorizeKvaScopes() {
  var out = [];
  try {
    var ss = kvaGetSpreadsheet_();
    out.push("Sheet: " + ss.getName());
  } catch (e1) { out.push("Sheet Fehler: " + e1.message); }
  try {
    var folder = DriveApp.getFolderById(KVA_WA_DRIVE_FOLDER_ID);
    out.push("Drive-Ordner: " + folder.getName());
  } catch (e2) { out.push("Drive Fehler: " + e2.message); }
  try {
    var resp = UrlFetchApp.fetch("https://www.googleapis.com/discovery/v1/apis?preferred=true", { muteHttpExceptions: true });
    out.push("UrlFetch: HTTP " + resp.getResponseCode());
  } catch (e3) { out.push("UrlFetch Fehler: " + e3.message); }
  try {
    out.push("Token: " + (ScriptApp.getOAuthToken() ? "ok" : "leer"));
  } catch (e4) { out.push("Token Fehler: " + e4.message); }
  var msg = out.join(" | ");
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e5) {}
  return msg;
}

function onEdit(e) {
  if (!e || !e.range) return;
  var range = e.range;
  var sheet = range.getSheet();
  if (sheet.getName() !== KVA_WA_INPUT_TAB) return;
  if (range.getA1Notation().match(/^B(2[9-9]|3[0-9]|4[0-1])$/)) {
    updateReparaturauftrag(e.source);
  }
}

function updateReparaturauftrag(optSs) {
  var ss = kvaGetSpreadsheet_(optSs);
  var inputSheet = ss.getSheetByName(KVA_WA_INPUT_TAB);
  var outputSheet = ss.getSheetByName(KVA_WA_REP_TAB);
  if (!inputSheet || !outputSheet) return;

  var xRange = inputSheet.getRange("B29:B41").getValues();
  var contentRange = inputSheet.getRange("C29:C41").getValues();

  var values = [];
  for (var i = 0; i < xRange.length; i++) {
    if (String(xRange[i][0] || "").toLowerCase().trim() === "x") {
      values.push(contentRange[i][0]);
    }
  }

  var totalCells = 4;
  outputSheet.getRange("G9:G12").clearContent().clearFormat();

  if (!values.length) return;

  var startIndex = Math.max(0, Math.floor((totalCells - values.length) / 2));
  for (var v = 0; v < values.length && (startIndex + v) < totalCells; v++) {
    var row = startIndex + v + 9;
    var cell = outputSheet.getRange(row, 7);
    cell.setValue(values[v]);
    cell.setFontFamily("Arial");
    cell.setFontSize(33);
    cell.setHorizontalAlignment("center");
    cell.setVerticalAlignment("middle");
  }
}

function resetInput(optSs) {
  var ss = kvaGetSpreadsheet_(optSs);
  var sheet = ss.getSheetByName(KVA_WA_INPUT_TAB);
  if (sheet) {
    sheet.getRange("B2:B45").clearContent();
    sheet.getRange("D24:D29").clearContent();
    sheet.getRange("D31:D36").clearContent();
    sheet.getRange("E2:E45").clearContent();
    sheet.getRange("H2:H45").clearContent();
    sheet.getRange("J2:O45").clearContent();
    sheet.getRange("Q2:V45").clearContent();
  }
  var rep = ss.getSheetByName(KVA_WA_REP_TAB);
  if (rep) {
    try { rep.getRange("G9:G12").clearContent().clearFormat(); } catch (e) {}
  }
  return { success: true };
}

function exportNachtrag(optSs, silent) {
  var ss = kvaGetSpreadsheet_(optSs);
  var inputSheet = ss.getSheetByName(KVA_WA_INPUT_TAB);
  var repSheet = ss.getSheetByName(KVA_WA_REP_TAB);

  var targetSpreadsheet = SpreadsheetApp.openById(KVA_NACHTRAG_SHEET_ID);
  var targetSheet = targetSpreadsheet.getSheetByName(KVA_NACHTRAG_TAB);

  if (!inputSheet || !repSheet || !targetSheet) {
    if (!silent) {
      try { SpreadsheetApp.getUi().alert("Ein Blatt wurde nicht gefunden!"); } catch (e) {}
    }
    return { success: false, message: "Ein Blatt wurde nicht gefunden!" };
  }

  var headerValue = String(inputSheet.getRange("B2").getValue() || "").trim();
  var values = repSheet.getRange("C15:C35").getValues().flat().filter(function(v) { return String(v || "").trim() !== ""; });
  if (values.length === 0) {
    values = repSheet.getRange("C14:C35").getValues().flat().filter(function(v) { return String(v || "").trim() !== ""; });
  }

  if (values.length === 0) {
    if (!silent) {
      try { SpreadsheetApp.getUi().alert("Keine Daten zum Export vorhanden!"); } catch (e2) {}
    }
    return { success: false, message: "Keine Daten zum Export vorhanden!" };
  }

  var combinedText = values.join("\n");
  var columnAValues = targetSheet.getRange("A:A").getValues().flat();
  var nextRow = columnAValues.findIndex(function(v) { return v === "" || v === null; }) + 1;
  if (nextRow === 0) nextRow = targetSheet.getLastRow() + 1;

  targetSheet.getRange(nextRow, 1, 1, 2).setValues([[headerValue, combinedText]]);
  targetSheet.getRange(nextRow, 2).setWrap(true);

  var reifenTexte = [
    "Reifen vorne rechts erneuern",
    "Reifen vorne links erneuern",
    "Reifen hinten rechts erneuern",
    "Reifen hinten links erneuern"
  ];

  var reifenVorhanden = values.some(function(v) { return reifenTexte.indexOf(String(v).trim()) !== -1; });
  if (reifenVorhanden) {
    targetSheet.getRange(nextRow, 3).setValue("Extern + Reifen Seng");
  }

  if (!silent) {
    try { SpreadsheetApp.getUi().alert("Export erfolgreich in Zeile " + nextRow + "!"); } catch (e3) {}
  }
  return {
    success: true,
    row: nextRow,
    stockId: headerValue,
    reifenVorhanden: reifenVorhanden,
    message: "Export erfolgreich in Zeile " + nextRow + "!"
  };
}

function doGet(e) {
  e = e || {};
  var p = e.parameter || {};
  var action = String(p.action || "").trim().toLowerCase();

  if (action === "ping") {
    return kvaJsonOut_({ success: true, version: KVA_WA_VERSION });
  }
  if (action === "meta") {
    return kvaJsonOut_(getInputMeta(p.force === "1"));
  }
  if (action === "state") {
    return kvaJsonOut_(getCurrentSheetState());
  }

  return HtmlService.createHtmlOutputFromFile("KVA_HUD")
    .setTitle("KVA Werkstattauftrag HUD")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (body.secret && String(body.secret) !== KVA_BRIDGE_SECRET) {
      return kvaJsonOut_({ success: false, message: "Ungültiges Secret" });
    }
    var action = String(body.action || "create").trim();
    if (action === "meta") {
      return kvaJsonOut_(getInputMeta(!!body.force));
    }
    if (action === "reset") {
      return kvaJsonOut_(resetInput());
    }
    if (action === "exportNachtrag") {
      return kvaJsonOut_(exportNachtrag(null, true));
    }
    if (action === "state") {
      return kvaJsonOut_(getCurrentSheetState());
    }
    return kvaJsonOut_(createKvaAuftrag(body));
  } catch (err) {
    return kvaJsonOut_({ success: false, message: String(err && err.message || err) });
  }
}

function kvaJsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function kvaNormalizeStockId_(v) {
  return String(v || "").toUpperCase().replace(/\s+/g, "").trim();
}

function getInputMeta(force) {
  try {
    var cache = CacheService.getScriptCache();
    if (!force) {
      var raw = cache.get(KVA_META_CACHE_KEY);
      if (raw) {
        try {
          var parsed = JSON.parse(raw);
          parsed.fromCache = true;
          return parsed;
        } catch (e1) {}
      }
    }

    var ss = kvaGetSpreadsheet_();
    var sheet = ss.getSheetByName(KVA_WA_INPUT_TAB);
    if (!sheet) return { success: false, message: "Tab '" + KVA_WA_INPUT_TAB + "' nicht gefunden" };

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
      if (rowNum === 10 || rowNum === 13) continue;
      if (rowNum <= 28) {
        workItems.push({ row: rowNum, label: label });
      } else {
        oils.push({
          row: rowNum,
          label: label,
          nr: String(cVals[rowNum - 29][0] || "").trim()
        });
      }
    }

    for (var d = 0; d < dVals.length; d++) {
      var dRow = 2 + d;
      var dLabel = String(dVals[d][0] || "").trim();
      if (!dLabel) continue;
      if ((dRow >= 23 && dRow <= 37)) continue;
      if (dLabel.toLowerCase().indexOf("sonstige") === 0) continue;
      dItems.push({ row: dRow, label: dLabel });
    }

    var actVals = sheet.getRange("F2:G7").getValues();
    var actions = [];
    for (var a = 0; a < actVals.length; a++) {
      var aNr = parseInt(actVals[a][0], 10);
      var aLabel = String(actVals[a][1] || "").trim();
      if (aNr >= 1 && aNr <= 6 && aLabel) actions.push({ nr: aNr, label: aLabel });
    }

    var parts = [];
    var p1Vals = sheet.getRange("I2:I42").getValues();
    for (var p1 = 0; p1 < p1Vals.length; p1++) {
      var p1Label = String(p1Vals[p1][0] || "").trim();
      if (p1Label) parts.push({ side: 1, row: 2 + p1, label: p1Label });
    }
    var p2Vals = sheet.getRange("P2:P42").getValues();
    for (var p2 = 0; p2 < p2Vals.length; p2++) {
      var p2Label = String(p2Vals[p2][0] || "").trim();
      if (p2Label) parts.push({ side: 2, row: 2 + p2, label: p2Label });
    }
    parts.sort(function(x, y) { return x.label.localeCompare(y.label, "de"); });

    var metaOut = {
      success: true,
      version: KVA_WA_VERSION,
      workItems: workItems,
      oils: oils,
      dItems: dItems,
      actions: actions,
      parts: parts
    };
    try { cache.put(KVA_META_CACHE_KEY, JSON.stringify(metaOut), KVA_META_CACHE_TTL); } catch (e2) {}
    return metaOut;
  } catch (err) {
    return { success: false, message: String(err && err.message || err) };
  }
}

function getCurrentSheetState() {
  try {
    var ss = kvaGetSpreadsheet_();
    var inputSheet = ss.getSheetByName(KVA_WA_INPUT_TAB);
    var repSheet = ss.getSheetByName(KVA_WA_REP_TAB);
    if (!inputSheet) return { success: false, message: "Input nicht gefunden" };
    var head = inputSheet.getRange("B2:B8").getDisplayValues().flat();
    var lines = [];
    if (repSheet) {
      var repVals = repSheet.getRange("A14:C39").getDisplayValues();
      for (var i = 0; i < repVals.length; i++) {
        var cat = String(repVals[i][0] || repVals[i][1] || "").trim();
        var txt = String(repVals[i][2] || "").trim();
        if (txt && txt.indexOf("#N/A") === -1 && txt.indexOf("#NV") === -1) {
          lines.push({ category: cat, text: txt });
        }
      }
    }
    return {
      success: true,
      stockId: head[0] || "",
      modell: head[1] || "",
      vin: head[2] || "",
      erstzulassung: head[3] || "",
      kba: head[4] || "",
      getriebe: head[5] || "",
      km: head[6] || "",
      lines: lines
    };
  } catch (err) {
    return { success: false, message: String(err && err.message || err) };
  }
}

function kvaExportRepPdf_(ss) {
  var repSheet = ss.getSheetByName(KVA_WA_REP_TAB);
  if (!repSheet) return { success: false, message: "Tab '" + KVA_WA_REP_TAB + "' nicht gefunden" };
  try { repSheet.showSheet(); } catch (eShow) {}
  var sheetId = ss.getId();
  var exportUrl = "https://docs.google.com/spreadsheets/d/" + sheetId + "/export?exportFormat=pdf&format=pdf"
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

function kvaSavePdfToDrive_(pdfBlob, fileName) {
  var folder = DriveApp.getFolderById(KVA_WA_DRIVE_FOLDER_ID);
  try {
    var existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) {
      existing.next().setTrashed(true);
    }
  } catch (eOld) {}
  var file = folder.createFile(pdfBlob.setName(fileName));
  return {
    fileId: file.getId(),
    pdfUrl: file.getUrl(),
    fileName: fileName
  };
}

function createKvaAuftrag(payload) {
  var lock = LockService.getScriptLock();
  var gotLock = false;
  try {
    gotLock = lock.tryLock(45000);
    if (!gotLock) {
      return { success: false, message: "Werkstattauftrag gerade in Arbeit – bitte kurz erneut versuchen" };
    }

    payload = payload || {};
    var stockId = kvaNormalizeStockId_(payload.stockId);
    var modell = String(payload.modell || "").trim();
    var vin = String(payload.vin || "").trim().toUpperCase();

    if (!stockId) return { success: false, message: "Stock-ID fehlt" };
    if (!modell) return { success: false, message: "Modell fehlt" };
    if (!vin) return { success: false, message: "VIN fehlt" };

    var workRows = payload.workRows || [];
    var dRows = payload.dRows || [];
    var freeMech = (payload.freeMech || []).map(function(s) { return String(s || "").trim(); }).filter(Boolean);
    var freeParts = (payload.freeParts || []).map(function(s) { return String(s || "").trim(); }).filter(Boolean);
    var bodyWorks = payload.bodyWorks || [];
    var oilRows = [];
    if (Array.isArray(payload.oilRows)) {
      oilRows = payload.oilRows.map(function(r) { return parseInt(r, 10); }).filter(function(r) { return r >= 29 && r <= 41; });
    } else if (payload.oilRow) {
      var singleOil = parseInt(payload.oilRow, 10);
      if (singleOil >= 29 && singleOil <= 41) oilRows.push(singleOil);
    }

    var litersRaw = String(payload.liters || "").replace(",", ".").trim();
    var litersNum = litersRaw ? parseFloat(litersRaw) : 0;
    var nach = String(payload.nach || "").toLowerCase().trim();
    var nachDatum = String(payload.nachDatum || "").trim();

    var ss = kvaGetSpreadsheet_();
    var inputSheet = ss.getSheetByName(KVA_WA_INPUT_TAB);
    var repSheet = ss.getSheetByName(KVA_WA_REP_TAB);
    if (!inputSheet || !repSheet) {
      return { success: false, message: "Blatt 'Input' oder 'Reparaturauftrag' nicht gefunden" };
    }

    var oilSelected = oilRows.length > 0;
    var oilChangeChecked = false;
    for (var wc = 0; wc < workRows.length; wc++) {
      var wRow = parseInt(workRows[wc], 10);
      if (wRow === 12 || wRow === 14) oilChangeChecked = true;
    }

    if (oilChangeChecked && !oilSelected) {
      return { success: false, message: "Motorölwechsel ist ausgewählt – bitte ein Öl auswählen!" };
    }

    if (oilSelected) {
      var onlyMitgeliefert = oilRows.length === 1 && oilRows[0] === 41;
      if (!onlyMitgeliefert && !(litersNum > 0)) {
        return {
          success: false,
          message: "Füllmenge (Liter) fehlt! Wenn ein Öl ausgewählt ist, muss die Füllmenge (B13) eingetragen werden (sonst #N/A auf dem Werkstattauftrag)."
        };
      }
      if (workRows.indexOf(14) === -1) workRows.push(14);
    }

    if (nach === "ja" && !nachDatum) {
      return {
        success: false,
        message: "Nachuntersuchung Ja ausgewählt – bitte das Nachuntersuchungs-Datum (B45) eintragen!"
      };
    }

    resetInput(ss);

    inputSheet.getRange("B2").setValue(stockId);
    inputSheet.getRange("B3").setValue(modell);
    inputSheet.getRange("B4").setValue(vin);

    if (workRows.length) {
      var workMarks = inputSheet.getRange("B10:B28").getValues();
      for (var i = 0; i < workRows.length; i++) {
        var r = parseInt(workRows[i], 10);
        if (r >= 10 && r <= 28 && r !== 13) workMarks[r - 10][0] = "x";
      }
      inputSheet.getRange("B10:B28").setValues(workMarks);
    }

    if (litersNum > 0) {
      inputSheet.getRange("B13").setValue(litersNum);
    }

    if (oilRows.length) {
      var oilMarks = inputSheet.getRange("B29:B41").getValues();
      for (var o = 0; o < oilRows.length; o++) {
        var orow = oilRows[o];
        if (orow >= 29 && orow <= 41) oilMarks[orow - 29][0] = "x";
      }
      inputSheet.getRange("B29:B41").setValues(oilMarks);
    }

    if (dRows.length) {
      var dMarks = inputSheet.getRange("E2:E41").getValues();
      for (var d = 0; d < dRows.length; d++) {
        var dr = parseInt(dRows[d], 10);
        if (dr >= 2 && dr <= 41) dMarks[dr - 2][0] = "x";
      }
      inputSheet.getRange("E2:E41").setValues(dMarks);
    }

    for (var m = 0; m < freeMech.length && m < 6; m++) {
      var mVal = String(freeMech[m] || "").trim();
      if (mVal) {
        inputSheet.getRange(24 + m, 4).setValue(mVal);
        inputSheet.getRange(24 + m, 5).setValue("x");
      }
    }

    for (var p = 0; p < freeParts.length && p < 6; p++) {
      var pVal = String(freeParts[p] || "").trim();
      if (pVal) {
        inputSheet.getRange(31 + p, 4).setValue(pVal);
        inputSheet.getRange(31 + p, 5).setValue("x");
      }
    }

    for (var b = 0; b < bodyWorks.length; b++) {
      var bw = bodyWorks[b] || {};
      var bRow = parseInt(bw.row, 10);
      var bSide = parseInt(bw.side, 10);
      if (!(bRow >= 2 && bRow <= 42)) continue;
      var partCol = bSide === 2 ? 16 : 9;
      var acts = bw.actions || [];
      for (var ac = 0; ac < acts.length; ac++) {
        var an = parseInt(acts[ac], 10);
        if (an >= 1 && an <= 6) {
          inputSheet.getRange(bRow, partCol + an).setValue("x");
        }
      }
    }

    if (nach === "ja") {
      inputSheet.getRange("B44").setValue("x");
      if (nachDatum) inputSheet.getRange("B45").setValue(nachDatum);
    } else if (nach === "nein") {
      inputSheet.getRange("B43").setValue("x");
    }

    updateReparaturauftrag(ss);
    SpreadsheetApp.flush();
    Utilities.sleep(600);

    var repDisplay = repSheet.getRange("A14:C39").getDisplayValues();
    var printLines = [];
    for (var rl = 0; rl < repDisplay.length; rl++) {
      var cTxt = String(repDisplay[rl][2] || "").trim();
      var catTxt = String(repDisplay[rl][0] || repDisplay[rl][1] || "").trim();
      if (cTxt && cTxt.indexOf("#N/A") === -1 && cTxt.indexOf("#NV") === -1) {
        printLines.push((catTxt ? (catTxt + ": ") : "") + cTxt);
      }
    }

    var pdf = kvaExportRepPdf_(ss);
    if (!pdf.success) {
      return { success: false, message: pdf.message };
    }

    var fileName = stockId + ".pdf";
    var driveSaved = false;
    var driveError = "";
    var pdfUrl = "";
    var fileId = "";

    try {
      var savedDrive = kvaSavePdfToDrive_(pdf.blob, fileName);
      driveSaved = true;
      pdfUrl = savedDrive.pdfUrl;
      fileId = savedDrive.fileId;
    } catch (dErr) {
      driveError = String(dErr && dErr.message || dErr);
    }

    var uebertragRes = null;
    if (payload.runUebertrag) {
      try {
        uebertragRes = exportNachtrag(ss, true);
      } catch (uErr) {
        uebertragRes = { success: false, message: String(uErr && uErr.message || uErr) };
      }
    }

    if (payload.resetAfterCreate) {
      try {
        resetInput(ss);
        SpreadsheetApp.flush();
      } catch (rErr) {}
    }

    return {
      success: true,
      version: KVA_WA_VERSION,
      stockId: stockId,
      modell: modell,
      vin: vin,
      fileName: fileName,
      pdfB64: Utilities.base64Encode(pdf.bytes),
      driveSaved: driveSaved,
      driveError: driveError,
      pdfUrl: pdfUrl,
      fileId: fileId,
      printLines: printLines,
      uebertrag: uebertragRes
    };
  } catch (err) {
    return { success: false, message: String(err && err.message || err) };
  } finally {
    if (gotLock) {
      try { lock.releaseLock(); } catch (eLock) {}
    }
  }
}
