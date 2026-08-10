var HUD_WA_SHEET_ID = "1i_3360NECjPHsd687Uts4i3xmD4h67xu--PMazmKlw0";
var HUD_WA_INPUT_TAB = "Input";
var HUD_WA_REP_TAB = "Reparaturauftrag";
var HUD_NB_SHEET_ID = "1VGCAHUbOPgsInQICA1GnrtKg1EPK1d1zWB-GkLi6iVE";
var HUD_NB_TAB = "Input Exit";
var HUD_DRIVE_FOLDER_ID = "1yEwrUVpS2nDA9qL3ht0p_XF7S0M-RGH2";
var HUD_DRIVE_FOLDER_NAME = "EXIT Werkstattaufträge";
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
  mappe: ["werkstattmappe erstellt", "werkstattmappe"],
  entryId: ["entryid", "entry id", "entry_id", "eintrag id", "eintragid"],
  datum: ["datum", "bestelldatum", "datum der bestellung", "bestellt am", "erstellt am", "eingetragen am", "date", "timestamp", "zeitstempel"]
};

var HUD_INSPECTION_TOKENS = { "tüv": 1, "tuv": 1, "tuev": 1, "hu": 1, "hauptuntersuchung": 1 };
var HUD_CONNECTOR_TOKENS = { "und": 1, "u": 1, "oder": 1 };

function hudIsExcludedGrund_(grund) {
  var s = String(grund || "").toLowerCase();
  var tokens = s.split(/[^a-zäöüß]+/);
  var hasInspection = false;
  for (var i = 0; i < tokens.length; i++) {
    var t = tokens[i];
    if (!t) continue;
    if (HUD_INSPECTION_TOKENS[t]) { hasInspection = true; continue; }
    if (HUD_CONNECTOR_TOKENS[t]) continue;
    return false;
  }
  return hasInspection;
}

var HUD_WEB_APP_URL = "https://script.google.com/a/macros/auto1.com/s/AKfycbw-f8igKkOObQAwYAKk00U2PPHs10VhOxOC2QyplnbcsFW5sg0PS6OgXAmTvje7OHK-Jw/exec";

function authorizeScopes() {
  var out = [];
  try {
    var resp = UrlFetchApp.fetch("https://www.googleapis.com/discovery/v1/apis?preferred=true", { muteHttpExceptions: true });
    out.push("UrlFetch: HTTP " + resp.getResponseCode());
  } catch (e) { out.push("UrlFetch Fehler: " + e.message); }
  try {
    var folder = hudResolveDriveFolder_();
    out.push("Drive-Ordner: " + folder.getName() + " (" + folder.getId() + ")");
  } catch (e2) { out.push("Drive Fehler: " + e2.message); }
  try {
    out.push("Token: " + (ScriptApp.getOAuthToken() ? "ok" : "leer"));
  } catch (e3) { out.push("Token Fehler: " + e3.message); }
  Logger.log(out.join(" | "));
  return out.join(" | ");
}

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("EXIT HUD")
      .addItem("EXIT HUD öffnen", "openExitHud")
      .addToUi();
  } catch (e) {}
}

function openExitHud() {
  var html = HtmlService.createHtmlOutput(
    '<html><body style="font-family:Segoe UI,sans-serif;padding:14px;">'
    + '<p>EXIT HUD wird geöffnet…</p>'
    + '<p>Falls sich kein Tab öffnet: <a href="' + HUD_WEB_APP_URL + '" target="_blank" rel="noopener">hier klicken</a>.</p>'
    + '<script>'
    + 'window.open("' + HUD_WEB_APP_URL + '","_blank");'
    + 'google.script.host.close();'
    + '</script>'
    + '</body></html>'
  ).setWidth(320).setHeight(140);
  SpreadsheetApp.getUi().showModalDialog(html, "EXIT HUD");
}

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

function hudFormatDate_(v) {
  if (v === null || v === undefined || v === "") return "";
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, "Europe/Berlin", "dd.MM.yyyy");
  }
  return String(v).trim();
}

var HUD_OPEN_CACHE_KEY = "open_mappen_base_v3";
var HUD_OPEN_CACHE_SECONDS = 600;

function hudBuildOpenMappen_() {
  var sheet = SpreadsheetApp.openById(HUD_NB_SHEET_ID).getSheetByName(HUD_NB_TAB);
  if (!sheet) throw new Error("Tab '" + HUD_NB_TAB + "' nicht gefunden");
  var layout = hudGetNbLayout_(sheet);
  var lastRow = sheet.getLastRow();
  var order = [];
  var byStock = {};
  if (lastRow >= layout.dataStartRow) {
    var numRows = lastRow - layout.dataStartRow + 1;
    var data = sheet.getRange(layout.dataStartRow, 1, numRows, layout.lastCol).getValues();
    var cols = layout.cols;
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var stockId = hudNormalizeStockId_(row[cols.stockId - 1]);
      if (!stockId) continue;
      if (hudIsChecked_(row[cols.mappe - 1])) continue;
      var grund = cols.grund ? String(row[cols.grund - 1] || "").trim() : "";
      if (hudIsExcludedGrund_(grund)) continue;
      var art = cols.art ? String(row[cols.art - 1] || "").trim() : "";
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
        entryId: cols.entryId ? String(row[cols.entryId - 1] || "").trim() : "",
        grund: grund,
        art: art,
        bearbeiter: cols.bearbeiter ? String(row[cols.bearbeiter - 1] || "").trim() : "",
        status: cols.status ? String(row[cols.status - 1] || "").trim() : "",
        datum: cols.datum ? hudFormatDate_(row[cols.datum - 1]) : ""
      });
    }
  }
  var entries = [];
  for (var o = 0; o < order.length; o++) entries.push(byStock[order[o]]);
  return entries;
}

function hudApplyCarol_(entries) {
  var cache = CacheService.getScriptCache();
  for (var o = 0; o < entries.length; o++) {
    var ent = entries[o];
    var cached = cache.get(HUD_CAROL_CACHE_PREFIX + ent.stockId);
    if (cached) {
      try {
        var cd = JSON.parse(cached);
        ent.carolModell = cd.modell || "";
        ent.carolVin = cd.vin || "";
      } catch (e2) {}
    } else {
      ent.carolModell = "";
      ent.carolVin = "";
    }
  }
  return entries;
}

function hudSignature_(entries) {
  var parts = [];
  for (var i = 0; i < entries.length; i++) {
    parts.push(entries[i].stockId + ":" + entries[i].rows.length);
  }
  return parts.join("|");
}

function hudInvalidateOpenCache_() {
  try { CacheService.getScriptCache().remove(HUD_OPEN_CACHE_KEY); } catch (e) {}
}

function getOpenMappen(force) {
  try {
    var cache = CacheService.getScriptCache();
    var entries = null;
    var fromCache = false;
    if (!force) {
      var raw = cache.get(HUD_OPEN_CACHE_KEY);
      if (raw) {
        try { entries = JSON.parse(raw); fromCache = true; } catch (e) { entries = null; }
      }
    }
    if (!entries) {
      entries = hudBuildOpenMappen_();
      try { cache.put(HUD_OPEN_CACHE_KEY, JSON.stringify(entries), HUD_OPEN_CACHE_SECONDS); } catch (ePut) {}
    }
    hudApplyCarol_(entries);
    return {
      success: true,
      entries: entries,
      petronasUrl: HUD_PETRONAS_URL,
      fromCache: fromCache,
      sig: hudSignature_(entries)
    };
  } catch (err) {
    return { success: false, message: String(err.message || err), entries: [] };
  }
}

var HUD_META_CACHE_KEY = "input_meta_v3";
var HUD_META_CACHE_SECONDS = 1800;

function getInputMeta() {
  try {
    var cache = CacheService.getScriptCache();
    var rawMeta = cache.get(HUD_META_CACHE_KEY);
    if (rawMeta) {
      try {
        var parsed = JSON.parse(rawMeta);
        parsed.fromCache = true;
        return parsed;
      } catch (eMeta) {}
    }
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
    var metaOut = { success: true, workItems: workItems, oils: oils, dItems: dItems, actions: actions, parts: parts };
    try { cache.put(HUD_META_CACHE_KEY, JSON.stringify(metaOut), HUD_META_CACHE_SECONDS); } catch (ePutMeta) {}
    return metaOut;
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
  try { hudClearBodySummariesOnSheet_(sheet); } catch (e) {}
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

function hudNormLabel_(v) {
  return String(v || "").toLowerCase()
    .replace(/ä/g, "ae").replace(/ö/g, "oe").replace(/ü/g, "ue").replace(/ß/g, "ss")
    .replace(/\s+/g, " ").trim();
}

function hudClearBodySummariesOnSheet_(sheet) {
  if (!sheet) return;
  var values = sheet.getDataRange().getValues();
  var keys = [
    "spaltmass einstellen", "spaltmaß einstellen",
    "ganzes bauteil lackieren", "beilackieren",
    "bauteil reparieren", "bauteil polieren",
    "delle druecken", "delle drücken",
    "polieren gesamtes fahrzeug"
  ];
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var n = hudNormLabel_(values[r][c]);
      if (!n) continue;
      for (var k = 0; k < keys.length; k++) {
        if (n.indexOf(keys[k]) === -1) continue;
        if (c + 1 < values[r].length) {
          try { sheet.getRange(r + 1, c + 2).clearContent(); } catch (e) {}
        }
        break;
      }
    }
  }
}

function hudWriteBodySummariesOnSheet_(sheet, byAction) {
  if (!sheet) return;
  var labelMap = {
    1: ["ganzes bauteil lackieren"],
    2: ["beilackieren"],
    3: ["delle druecken", "delle drücken"],
    4: ["spaltmass einstellen", "spaltmaß einstellen"],
    5: ["bauteil reparieren"],
    6: ["bauteil polieren", "polieren gesamtes fahrzeug"]
  };
  var values = sheet.getDataRange().getValues();
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var raw = String(values[r][c] || "");
      var n = hudNormLabel_(raw);
      if (!n) continue;
      for (var an = 1; an <= 6; an++) {
        var keys = labelMap[an];
        var hit = false;
        for (var k = 0; k < keys.length; k++) {
          if (n.indexOf(keys[k]) !== -1) { hit = true; break; }
        }
        if (!hit) continue;
        var parts = byAction[an] || [];
        if (!parts.length) continue;
        var text = parts.join(", ");
        if (/:\s*$/.test(raw) || n.indexOf(keys[0]) === 0) {
          sheet.getRange(r + 1, c + 2).setValue(text);
        } else if (n.indexOf(":") !== -1) {
          sheet.getRange(r + 1, c + 1).setValue(raw.replace(/:\s*.*$/, ": ") + text);
        } else {
          sheet.getRange(r + 1, c + 2).setValue(text);
        }
      }
    }
  }
}

function hudApplyBodyWorks_(ss, bodyWorks) {
  var inputSheet = ss.getSheetByName(HUD_WA_INPUT_TAB);
  var repSheet = ss.getSheetByName(HUD_WA_REP_TAB);
  hudClearBodySummariesOnSheet_(inputSheet);
  hudClearBodySummariesOnSheet_(repSheet);
  if (!bodyWorks || !bodyWorks.length) return;

  var byAction = { 1: [], 2: [], 3: [], 4: [], 5: [], 6: [] };
  for (var b = 0; b < bodyWorks.length; b++) {
    var bw = bodyWorks[b] || {};
    var bRow = parseInt(bw.row, 10);
    var bSide = parseInt(bw.side, 10);
    if (!(bRow >= 2 && bRow <= 42)) continue;
    var partCol = bSide === 2 ? 16 : 9;
    var label = String(bw.label || "").trim();
    if (!label && inputSheet) label = String(inputSheet.getRange(bRow, partCol).getValue() || "").trim();
    var acts = bw.actions || [];
    for (var ac = 0; ac < acts.length; ac++) {
      var an = parseInt(acts[ac], 10);
      if (!(an >= 1 && an <= 6)) continue;
      inputSheet.getRange(bRow, partCol + an).setValue("x");
      if (label && byAction[an].indexOf(label) === -1) byAction[an].push(label);
    }
  }
  hudWriteBodySummariesOnSheet_(inputSheet, byAction);
  hudWriteBodySummariesOnSheet_(repSheet, byAction);
}

function hudResolveDriveFolder_() {
  try { PropertiesService.getScriptProperties().deleteProperty("HUD_DRIVE_FOLDER_ID_RESOLVED"); } catch (e0) {}
  try {
    return DriveApp.getFolderById(HUD_DRIVE_FOLDER_ID);
  } catch (e) {
    throw new Error("Kein Zugriff auf Ordner \"" + HUD_DRIVE_FOLDER_NAME + "\" (" + HUD_DRIVE_FOLDER_ID + "). " +
      "Der Ordner muss von ersatzteile.hemau@autohero.com für dein Konto als Bearbeiter freigegeben sein.");
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

function hudMarkMappeErstellt_(stockId, entryIds, sheetRows) {
  var sheet = SpreadsheetApp.openById(HUD_NB_SHEET_ID).getSheetByName(HUD_NB_TAB);
  if (!sheet) return 0;
  var layout = hudGetNbLayout_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < layout.dataStartRow) return 0;
  var numRows = lastRow - layout.dataStartRow + 1;
  var data = sheet.getRange(layout.dataStartRow, 1, numRows, layout.lastCol).getValues();
  var cols = layout.cols;

  var idSet = {};
  var haveIds = false;
  if (entryIds && entryIds.length) {
    for (var e = 0; e < entryIds.length; e++) {
      var idv = String(entryIds[e] || "").trim();
      if (idv) { idSet[idv] = 1; haveIds = true; }
    }
  }
  var rowSet = {};
  var haveRows = false;
  if (sheetRows && sheetRows.length) {
    for (var s = 0; s < sheetRows.length; s++) {
      var rv = parseInt(sheetRows[s], 10);
      if (rv > 0) { rowSet[rv] = 1; haveRows = true; }
    }
  }
  var selective = haveIds || haveRows;

  var marked = 0;
  for (var i = 0; i < data.length; i++) {
    var sid = hudNormalizeStockId_(data[i][cols.stockId - 1]);
    if (sid !== stockId) continue;
    if (hudIsChecked_(data[i][cols.mappe - 1])) continue;
    var sheetRow = layout.dataStartRow + i;
    if (selective) {
      var rowEntryId = cols.entryId ? String(data[i][cols.entryId - 1] || "").trim() : "";
      var matchId = haveIds && rowEntryId && idSet[rowEntryId] === 1;
      var matchRow = haveRows && rowSet[sheetRow] === 1;
      if (!matchId && !matchRow) continue;
    } else {
      if (hudIsExcludedGrund_(cols.grund ? data[i][cols.grund - 1] : "")) continue;
    }
    sheet.getRange(sheetRow, cols.mappe).setValue(true);
    marked++;
  }
  return marked;
}

function markOrdersDone(stockId, entryIds, sheetRows) {
  var lock = LockService.getScriptLock();
  var gotLock = false;
  try {
    gotLock = lock.tryLock(30000);
    if (!gotLock) return { success: false, message: "Gerade beschäftigt – bitte kurz erneut versuchen" };
    stockId = hudNormalizeStockId_(stockId);
    if (!stockId) return { success: false, message: "Keine Stock-ID" };
    if ((!entryIds || !entryIds.length) && (!sheetRows || !sheetRows.length)) {
      return { success: false, message: "Keine Auftrags-ID übergeben" };
    }
    var marked = hudMarkMappeErstellt_(stockId, entryIds, sheetRows);
    if (marked > 0) hudInvalidateOpenCache_();
    return { success: marked > 0, marked: marked, message: marked > 0 ? (marked + " Auftrag(e) erledigt") : "Nichts markiert (evtl. schon erledigt)" };
  } catch (err) {
    return { success: false, message: String(err.message || err) };
  } finally {
    if (gotLock) { try { lock.releaseLock(); } catch (e) {} }
  }
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
    var bodyWorks = payload.bodyWorks || [];
    var oilRow = parseInt(payload.oilRow, 10) || 0;
    var liters = String(payload.liters || "").replace(",", ".").trim();
    var litersNum = liters ? parseFloat(liters) : 0;

    if (!workRows.length && !dRows.length && !freeMech.length && !freeParts.length && !bodyWorks.length) {
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
    var erstzulassung = String(payload.erstzulassung || "").trim();
    var kba = String(payload.kba || "").trim();
    var getriebe = String(payload.getriebe || "").trim();
    var km = String(payload.km || "").trim();
    if (erstzulassung) inputSheet.getRange("B5").setValue(erstzulassung);
    if (kba) inputSheet.getRange("B6").setValue(kba);
    if (getriebe) inputSheet.getRange("B7").setValue(getriebe);
    if (km) inputSheet.getRange("B8").setValue(km);

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

    hudApplyBodyWorks_(ss, bodyWorks);

    var nach = String(payload.nach || "").toLowerCase();
    var nachDatum = String(payload.nachDatum || "").trim();
    if (nach === "ja") {
      inputSheet.getRange("B44").setValue("x");
      if (nachDatum) inputSheet.getRange("B45").setValue(nachDatum);
    } else if (nach === "nein") {
      inputSheet.getRange("B43").setValue("x");
    }

    hudApplyOilToReparaturauftrag_(ss);
    SpreadsheetApp.flush();
    Utilities.sleep(2500);
    SpreadsheetApp.flush();

    var pdf = hudExportRepPdf_(ss);
    if (!pdf.success) return { success: false, message: pdf.message };

    var fileName = stockId + " Werkstattauftrag EXIT.pdf";
    var pdfUrl = "";
    var driveMsg = "";
    var driveSaved = false;
    try {
      var folder = hudResolveDriveFolder_();
      var file = folder.createFile(pdf.blob.setName(fileName));
      pdfUrl = file.getUrl();
      driveSaved = true;
    } catch (driveErr) {
      driveMsg = String(driveErr.message || driveErr);
    }

    var marked = 0;
    var nbMsg = "";
    try {
      marked = hudMarkMappeErstellt_(stockId, payload.entryIds || [], payload.sheetRows || []);
      if (marked > 0) hudInvalidateOpenCache_();
    } catch (nbErr) {
      nbMsg = String(nbErr.message || nbErr);
    }

    try {
      hudResetInput_(inputSheet);
      hudClearBodySummariesOnSheet_(ss.getSheetByName(HUD_WA_REP_TAB));
      SpreadsheetApp.flush();
    } catch (resetErr) {}

    return {
      success: true,
      stockId: stockId,
      modell: modell,
      vin: vin,
      printB64: Utilities.base64Encode(pdf.bytes),
      driveSaved: driveSaved,
      driveError: driveMsg,
      pdfUrl: pdfUrl,
      fileName: fileName,
      markedRows: marked,
      nbError: nbMsg
    };
  } catch (err) {
    return { success: false, message: String(err.message || err) };
  } finally {
    if (gotLock) {
      try { lock.releaseLock(); } catch (e) {}
    }
  }
}
