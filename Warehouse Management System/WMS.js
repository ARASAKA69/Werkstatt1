const TRACKING_SHEET_URL = "https://docs.google.com/spreadsheets/d/1PuCLw8UmDjB_pBo_jCZ9rmSD3GJQESHzPoBVu_--MRo/edit?gid=1453769469#gid=1453769469";
const REIFEN_SHEET_ID = "1NTWkl4r40VUb8hM3Zk5BYWofdxn0FgtZh4DJpOufSd8";
const NACHBESTELL_SHEET_ID = "1VGCAHUbOPgsInQICA1GnrtKg1EPK1d1zWB-GkLi6iVE";
const NACHBESTELL_TAB = "Nachbestellung";
const NACHBESTELL_GID = 130741593;
const AUFTRAG_SHEET_ID = "1nE6SErc1-jmZYd_Ydviw28Pa5qdJmwNepXCiVbsdsVo";
const AUFTRAG_TAB = "BLANCO Reparaturauftrag";
const AUFTRAG_EMAIL = "francesco.berger@auto1.com";
const HEMAU_SHEET_ID = "13Oh7gDT8NAul2s0cwQUeaGwMcS3B2MYu0QOdFNMhXzM";
const HEMAU_DAILY_PLANNING_TAB = "Daily Planning List";
const TAGESLISTE_SHEET_ID = "1PuCLw8UmDjB_pBo_jCZ9rmSD3GJQESHzPoBVu_--MRo";
const TAGESLISTE_TAB = "Tagesliste";
const VASOLD_WSS_TAB = "Vasold WSS";
const NACHBESTELL_STATUS_COL = 11;
const NACHBESTELL_REGAL_COL = 13;
const NACHBESTELL_ENTRYID_COL = 15;
const INPUT_EXIT_TAB = "Input Exit";
const INPUT_EXIT_STATUS_COL = 11;
const INPUT_EXIT_STATUS_DATE_COL = 12;
const WMS_WEB_APP_URL = "https://script.google.com/a/macros/auto1.com/s/AKfycbz3tBqPKeNI4JPd0ytWxb_6hXpHd8sjgfHAPaHBewIgcHMHiQkNg13Xa30K5FAaGjIG/exec";
const WMS_APP_VERSION = "2.0.5";
const WMS_APP_CHANGELOG = "• Update öffnet nur noch den sauberen /exec Link (kein wmsboot mehr)";
const GMAIL_LOOKUP_SHEET_ID = "16QFzXPUkxvpTHwSSAtjRAeKYb5YdrQPhUrBWInygASE";
const GMAIL_LOOKUP_TAB = "Lookup";
const PACKZETTEL_TAB = "Packzettel";
const KOMMENTAR_VERLAUF_SHEET_ID = "11d2YPM4wqLbGMkTCL7-lZJ1GXcHydNiNjOvRmkSxPEM";
const KOMMENTAR_VERLAUF_TAB = "Kommentar Verlauf";
const INFO_LAGER_EXIT_WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/AAQA5uO7fLU/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=I-lsD4vjmGQdTIecI9l7hyg7XO3let1u9VRS4ofAIT8";

function normalizeRegalKeyForCount(val) {
  if (val === "" || val == null) return "";
  if (Object.prototype.toString.call(val) === "[object Date]" || val instanceof Date) {
    var dm = Utilities.formatDate(new Date(val), "Europe/Berlin", "d.M");
    var dp = dm.split(".");
    if (dp.length >= 2) {
      var da = parseInt(dp[0], 10);
      var mo = parseInt(dp[1], 10);
      if (!isNaN(da) && !isNaN(mo) && da >= 1 && da <= 9 && mo >= 1 && mo <= 8) return "Regal " + da + "." + mo;
    }
    return "";
  }
  if (typeof val === "number" && isFinite(val)) {
    var a = Math.floor(val);
    var rest = val - a;
    var b = Math.round(rest * 10 + 1e-6);
    if (a >= 1 && a <= 9 && b >= 1 && b <= 8) return "Regal " + a + "." + b;
    return "";
  }
  var s = String(val || "").trim().replace(/,/g, ".");
  if (!s) return "";
  var low = s.toLowerCase();
  if (low === "tagesliste" || low === "lack" || low === "exit") return "";
  var m = s.match(/^regal\s+(\d+)\s*\.\s*(\d+)$/i);
  if (m) return "Regal " + parseInt(m[1], 10) + "." + parseInt(m[2], 10);
  m = s.match(/^(\d+)\s*\.\s*(\d+)$/);
  if (m) {
    var a2 = parseInt(m[1], 10);
    var b2 = parseInt(m[2], 10);
    if (a2 >= 1 && a2 <= 9 && b2 >= 1 && b2 <= 8) return "Regal " + a2 + "." + b2;
  }
  return "";
}

function nachbestellungRegalUiFromCell(val) {
  var k = normalizeRegalKeyForCount(val);
  if (k) return k.replace(/^Regal /i, "");
  return String(val || "").trim();
}

function nachbestellungLagerortVerifyMatch(expected, cellValue) {
  var e = String(expected || "").trim();
  if (e === "" && (cellValue === "" || cellValue == null)) return true;
  if (cellValue === "" || cellValue == null) return e === "";
  if (typeof cellValue === "string" && e.toLowerCase() === cellValue.trim().toLowerCase()) return true;
  var lowE = e.toLowerCase();
  if ((lowE === "tagesliste" || lowE.indexOf("tagesliste") !== -1 || lowE === "lack" || lowE === "exit") && String(cellValue || "").trim().toLowerCase() === lowE) return true;
  if (lowE.indexOf("tagesliste") !== -1 && lowE.indexOf("automatik") !== -1) {
    var cvLow = String(cellValue || "").trim().toLowerCase();
    if (cvLow.indexOf("tagesliste") !== -1 && cvLow.indexOf("automatik") !== -1) return true;
  }
  var ne = normalizeRegalKeyForCount(e);
  var nv = normalizeRegalKeyForCount(cellValue);
  if (ne && nv && ne === nv) return true;
  return false;
}

function getColIndexExact(headerRow, exactTerms) {
  if (!headerRow) return -1;
  for (var i = 0; i < headerRow.length; i++) {
    var cellText = String(headerRow[i] || "").toLowerCase().replace(/[^a-z0-9äöüß]/g, "");
    for (var j = 0; j < exactTerms.length; j++) {
      var term = String(exactTerms[j] || "").toLowerCase().replace(/[^a-z0-9äöüß]/g, "");
      if (cellText === term) return i + 1;
    }
  }
  return -1;
}

function findNachbestellungSheetLayout(sheet) {
  var fallback = {
    headerRow: 2,
    dataStartRow: 3,
    statusCol: NACHBESTELL_STATUS_COL,
    lagerortCol: NACHBESTELL_REGAL_COL
  };
  if (!sheet) return fallback;
  var lastRow = Math.max(1, sheet.getLastRow());
  var lastCol = Math.max(1, Math.min(80, sheet.getLastColumn()));
  var headerScan = sheet.getRange(1, 1, Math.min(10, lastRow), lastCol).getValues();
  var h, bestIdx = -1, bestScore = 0;
  for (h = 0; h < headerScan.length; h++) {
    var row = headerScan[h];
    var stockCol = getColIndex(row, ["stockid", "stock"]);
    if (stockCol === -1) continue;
    var score = 1;
    var lagerortCol = getColIndex(row, ["lagerort", "regal"]);
    if (lagerortCol !== -1) score += 3;
    var statusCol = getColIndexExact(row, ["status"]);
    if (statusCol !== -1) score += 3;
    if (getColIndex(row, ["datum", "date"]) !== -1) score += 1;
    if (score > bestScore) {
      bestScore = score;
      bestIdx = h;
    }
  }
  if (bestIdx === -1) return fallback;
  var hdr = headerScan[bestIdx];
  var loCol = getColIndex(hdr, ["lagerort", "regal"]);
  var stCol = getColIndexExact(hdr, ["status"]);
  return {
    headerRow: bestIdx + 1,
    dataStartRow: bestIdx + 2,
    statusCol: stCol !== -1 ? stCol : NACHBESTELL_STATUS_COL,
    lagerortCol: loCol !== -1 ? loCol : NACHBESTELL_REGAL_COL
  };
}

function findNachbestellungLagerortColumn(sheet) {
  return findNachbestellungSheetLayout(sheet).lagerortCol;
}

function nachbestellungLagerortFallbackList() {
  var o = ["Tagesliste"];
  for (var i = 1; i <= 9; i++) {
    for (var j = 1; j <= 8; j++) {
      o.push("Regal " + i + "." + j);
    }
  }
  o.push("LACK");
  o.push("Exit");
  return o;
}

function extractDataValidationList(dv) {
  if (!dv) return null;
  if (dv.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return null;
  var cv = dv.getCriteriaValues();
  if (!cv || cv.length === 0) return null;
  var first = cv[0];
  var out = [];
  if (Object.prototype.toString.call(first) === "[object Array]") {
    for (var i = 0; i < first.length; i++) {
      var cellVal = first[i];
      if (cellVal !== "" && cellVal != null) out.push(String(cellVal));
    }
  } else if (first !== "" && first != null) {
    out.push(String(first));
  }
  return out.length ? out : null;
}

function getNachbestellungLagerortAllowedList(sheet, col) {
  var lastRow = Math.max(2, sheet.getLastRow());
  var samples = [2, 3, 4, 5, 10, 15, 20, 50, 100, 500, 1000, 2000, 3000, 4000, 5000];
  for (var s = 0; s < samples.length; s++) {
    var rr = samples[s];
    if (rr > lastRow) continue;
    var list = extractDataValidationList(sheet.getRange(rr, col).getDataValidation());
    if (list && list.length) return list;
  }
  return nachbestellungLagerortFallbackList();
}

function nachbestellungStatusFallbackList() {
  return [
    "bestellt",
    "nicht bestellt",
    "teilweise angeliefert",
    "komplett angeliefert",
    "Lieferrückstand",
    "B2A1",
    "Fertiggestellt",
    "Konsi-Lager"
  ];
}

function getNachbestellungStatusAllowedList(sheet, col) {
  var lastRow = Math.max(2, sheet.getLastRow());
  var samples = [2, 3, 4, 5, 10, 15, 20, 50, 100, 500, 1000, 2000, 3000, 4000, 5000];
  for (var s = 0; s < samples.length; s++) {
    var rr = samples[s];
    if (rr > lastRow) continue;
    var list = extractDataValidationList(sheet.getRange(rr, col).getDataValidation());
    if (list && list.length) return list;
  }
  return nachbestellungStatusFallbackList();
}

function nachbestellungStatusToSheetValue(raw, allowedList) {
  raw = String(raw || "").trim();
  if (!raw) return "";
  var i;
  for (i = 0; i < allowedList.length; i++) {
    var opt = String(allowedList[i] != null ? allowedList[i] : "").trim();
    if (raw.toLowerCase() === opt.toLowerCase()) return String(allowedList[i]);
  }
  return null;
}

function nachbestellungIsClosedStatus(statusVal) {
  var s = String(statusVal || "").trim().toLowerCase();
  if (!s) return false;
  if (s === "angeliefert" || s === "fahrzeug rr") return true;
  if (s.indexOf("angeliefert/bereit") !== -1) return true;
  if (s.indexOf("komplett angeliefert") !== -1) return true;
  if (s === "fertiggestellt") return true;
  if (s === "konsi-lager" || s === "konsilager") return true;
  return false;
}

function nachbestellungTageslisteAutomatikSheetValue(allowedList) {
  var candidates = ["Tagesliste automatik", "Tagesliste Automatik", "Tagesliste"];
  var c;
  for (c = 0; c < candidates.length; c++) {
    var v = nachbestellungLagerortToSheetValue(candidates[c], allowedList);
    if (v !== null) return v;
  }
  for (c = 0; c < allowedList.length; c++) {
    var opt = String(allowedList[c] || "").trim().toLowerCase();
    if (opt.indexOf("tagesliste") !== -1 && opt.indexOf("automatik") !== -1) return String(allowedList[c]);
  }
  for (c = 0; c < allowedList.length; c++) {
    var opt2 = String(allowedList[c] || "").trim().toLowerCase();
    if (opt2 === "tagesliste" || opt2.indexOf("tagesliste") === 0) return String(allowedList[c]);
  }
  return null;
}

function nachbestellungLagerortToSheetValue(raw, allowedList) {
  raw = String(raw || "").trim();
  if (!raw) return "";
  var i;
  for (i = 0; i < allowedList.length; i++) {
    var opt = String(allowedList[i] != null ? allowedList[i] : "").trim();
    if (raw.toLowerCase() === opt.toLowerCase()) return String(allowedList[i]);
  }
  var rawNorm = normalizeRegalKeyForCount(raw);
  for (i = 0; i < allowedList.length; i++) {
    var opt2 = String(allowedList[i] != null ? allowedList[i] : "").trim();
    var optNorm = normalizeRegalKeyForCount(opt2);
    if (rawNorm && optNorm && rawNorm === optNorm) return String(allowedList[i]);
  }
  return null;
}

function normalizeStockId(value) {
    return String(value || "").replace(/\s+/g, "").toUpperCase();
  }

function cellMatchesStockId(cellVal, stockId) {
    var cv = normalizeStockId(cellVal);
    var sid = normalizeStockId(stockId);
    return sid !== "" && cv === sid;
  }

function getColIndex(headerRow, searchTerms) {
    if (!headerRow) return -1;
    for (var i = 0; i < headerRow.length; i++) {
      var cellText = String(headerRow[i] || "").toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
      for (var j = 0; j < searchTerms.length; j++) {
        var term = String(searchTerms[j] || "").toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
        if (cellText === term || cellText.indexOf(term) !== -1) return i + 1;
      }
    }
    return -1;
  }

function findHeaderRow(data, searchTerms) {
    for (var i = 0; i < Math.min(30, data.length); i++) {
      if (getColIndex(data[i], searchTerms) !== -1) return i;
    }
    return -1;
  }

function findRowFast(sheet, searchTermsHeader, stockId) {
    stockId = normalizeStockId(stockId);
    var lastRow = Math.max(1, sheet.getLastRow());
    var lastCol = Math.max(1, Math.min(80, sheet.getLastColumn()));
    var headerData = sheet.getRange(1, 1, Math.min(30, lastRow), lastCol).getValues();
    var headerIdx = findHeaderRow(headerData, searchTermsHeader);
    if (headerIdx === -1) return { row: -1, headerIdx: -1, stockCol: -1 };

    var stockCol = getColIndex(headerData[headerIdx], searchTermsHeader);
    if (stockCol === -1) return { row: -1, headerIdx: headerIdx, stockCol: -1 };

    var colData = sheet.getRange(1, stockCol, lastRow, 1).getValues();
    for (var i = headerIdx + 1; i < colData.length; i++) {
      if (cellMatchesStockId(colData[i][0], stockId)) {
        return { row: i + 1, headerIdx: headerIdx, stockCol: stockCol };
      }
    }
    return { row: -1, headerIdx: headerIdx, stockCol: stockCol };
  }

function canonicalTireSizeServer(str) {
    var m = String(str || "").match(/(\d{3})\s*\/\s*(\d{2})\s*[A-Z]{0,2}\s*R?\s*(\d{2})/i);
    return m ? (m[1] + "/" + m[2] + "/" + m[3]) : "";
  }

function findReifenStockRowsDetailed(sheet, stockId) {
    stockId = normalizeStockId(stockId);
    var lastRow = Math.max(1, sheet.getLastRow());
    var lastCol = Math.max(1, Math.min(80, sheet.getLastColumn()));
    var headerData = sheet.getRange(1, 1, Math.min(30, lastRow), lastCol).getValues();
    var headerIdx = findHeaderRow(headerData, ["stockid", "stock"]);
    if (headerIdx === -1) return { headerIdx: -1, stockCol: -1, header: null, angeliefertCol: -1, matches: [] };
    var header = headerData[headerIdx];
    var stockCol = getColIndex(header, ["stockid", "stock"]);
    if (stockCol === -1) return { headerIdx: headerIdx, stockCol: -1, header: header, angeliefertCol: -1, matches: [] };
    var angeliefertCol = getColIndex(header, ["angeliefert"]);
    var groesseCol = getColIndex(header, ["größe", "groesse"]);
    var lastIndexCol = getColIndex(header, ["lastindex", "last"]);
    var gwIndexCol = getColIndex(header, ["gwindex", "gw"]);
    var mengeCol = getColIndex(header, ["menge", "anzahl"]);
    var startRow = headerIdx + 2;
    var numRows = lastRow - startRow + 1;
    var matches = [];
    if (numRows > 0) {
      var data = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
      for (var i = 0; i < data.length; i++) {
        if (!cellMatchesStockId(data[i][stockCol - 1], stockId)) continue;
        var mengeVal = mengeCol !== -1 ? (parseInt(data[i][mengeCol - 1], 10) || 0) : 0;
        if (mengeVal < 1) mengeVal = 2;
        matches.push({
          row: startRow + i,
          status: angeliefertCol !== -1 ? String(data[i][angeliefertCol - 1] || "").trim().toLowerCase() : "",
          groesse: groesseCol !== -1 ? String(data[i][groesseCol - 1] || "").trim() : "",
          lastindex: lastIndexCol !== -1 ? String(data[i][lastIndexCol - 1] || "").trim() : "",
          gwindex: gwIndexCol !== -1 ? String(data[i][gwIndexCol - 1] || "").trim() : "",
          menge: mengeVal
        });
      }
    }
    return { headerIdx: headerIdx, stockCol: stockCol, header: header, angeliefertCol: angeliefertCol, matches: matches };
  }

function pickReifenUnbookedRow(matches, sizeHint) {
    var unbooked = [];
    for (var i = 0; i < matches.length; i++) {
      if (matches[i].status !== "ja" && matches[i].status !== "nein") unbooked.push(matches[i]);
    }
    var sizeCanon = canonicalTireSizeServer(sizeHint);
    var chosen = null;
    if (sizeCanon) {
      for (var u = 0; u < unbooked.length; u++) {
        if (canonicalTireSizeServer(unbooked[u].groesse) === sizeCanon) { chosen = unbooked[u]; break; }
      }
    }
    if (!chosen && unbooked.length) chosen = unbooked[0];
    return { unbooked: unbooked, chosen: chosen };
  }

function getReifenSheet() {
    return SpreadsheetApp.openById(REIFEN_SHEET_ID);
  }

function getReifenSheetTab(tabName) {
    var ss = getReifenSheet();
    var name = String(tabName || "").trim();
    var sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
    var sheets = ss.getSheets();
    var normalizedName = name.replace(/\s+/g, '').toLowerCase();
    for (var i = 0; i < sheets.length; i++) {
      var sn = sheets[i].getName().replace(/\s+/g, '').toLowerCase();
      if (sn === normalizedName) return sheets[i];
    }
    return null;
  }

function parseReifenTabDateMsFromName(name) {
    var trimmed = String(name || "").trim();
    var all = trimmed.match(/\d{2}\.\d{2}\.\d{4}/g);
    if (!all || !all.length) return null;
    var dateStr = all[all.length - 1];
    var p = dateStr.split(".");
    if (p.length !== 3) return null;
    var t = new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10)).getTime();
    return isNaN(t) ? null : t;
  }

function filterReifenSheetNamesByTabDateRange(startBound, endBound) {
    var ss = getReifenSheet();
    var sheetNames = ss.getSheets().map(function(sheet) {
      return sheet.getName();
    });
    var filtered = [];
    var i;
    for (i = 0; i < sheetNames.length; i++) {
      var name = sheetNames[i];
      var tabDate = parseReifenTabDateMsFromName(name);
      if (tabDate === null) continue;
      if (tabDate < startBound || tabDate > endBound) continue;
      filtered.push(name);
    }
    filtered.sort(function(a, b) {
      var ta = parseReifenTabDateMsFromName(a);
      var tb = parseReifenTabDateMsFromName(b);
      if (ta === null || tb === null) return 0;
      return tb - ta;
    });
    return filtered;
  }

function getReifenTabNamesForStockHudWindow() {
    var now = new Date();
    var y = now.getFullYear();
    var mo = now.getMonth();
    var d = now.getDate();
    var endBound = new Date(y, mo, d, 23, 59, 59, 999).getTime();
    var startBound = new Date(y, mo, d - 7, 0, 0, 0, 0).getTime();
    return filterReifenSheetNamesByTabDateRange(startBound, endBound);
  }

function getReifenTabNamesForPaketHudWindow() {
    var now = new Date();
    var y = now.getFullYear();
    var mo = now.getMonth();
    var d = now.getDate();
    var endBound = new Date(y, mo, d, 23, 59, 59, 999).getTime();
    var startBound = new Date(y, mo, d - 21, 0, 0, 0, 0).getTime();
    return filterReifenSheetNamesByTabDateRange(startBound, endBound);
  }

function getReifenTabOptions() {
    try {
      return { success: true, tabs: getReifenTabNamesForStockHudWindow() };
    } catch (err) {
      return { success: false, message: err.message, tabs: [] };
    }
  }

function appendPaketdienstRowsFromSheet_(sh, results) {
    var tabName = sh.getName();
    var lastRow = Math.max(1, sh.getLastRow());
    var lastCol = Math.max(1, Math.min(80, sh.getLastColumn()));
    if (lastRow < 3) return;

    var headerData = sh.getRange(1, 1, Math.min(30, lastRow), lastCol).getValues();
    var headerIdx = findHeaderRow(headerData, ["stockid", "stock"]);
    if (headerIdx === -1) return;
    var header = headerData[headerIdx];

    var stockCol = getColIndex(header, ["stockid", "stock"]);
    var groesseCol = getColIndex(header, ["größe", "groesse"]);
    var lastIndexCol = getColIndex(header, ["lastindex", "last"]);
    var gwIndexCol = getColIndex(header, ["gwindex", "gw"]);
    var mengeCol = getColIndex(header, ["menge", "anzahl"]);
    var angeliefertCol = getColIndex(header, ["angeliefert"]);
    if (stockCol === -1 || groesseCol === -1) return;

    var startRow = headerIdx + 2;
    var numRows = lastRow - startRow + 1;
    if (numRows <= 0) return;
    var data = sh.getRange(startRow, 1, numRows, lastCol).getValues();

    var i;
    for (i = 0; i < data.length; i++) {
      var row = data[i];
      var hasPaket = false;
      var c;
      for (c = 0; c < row.length; c++) {
        if (String(row[c] || "").toLowerCase().indexOf("paketdienst") !== -1) { hasPaket = true; break; }
      }
      if (!hasPaket) continue;

      if (angeliefertCol !== -1) {
        var aStatus = String(row[angeliefertCol - 1] || "").trim().toLowerCase();
        if (aStatus === "ja" || aStatus === "nein") continue;
      }

      var cellGroesse = String(row[groesseCol - 1] || "").trim();
      var cellLast = lastIndexCol !== -1 ? String(row[lastIndexCol - 1] || "").trim() : "";
      var cellGw = gwIndexCol !== -1 ? String(row[gwIndexCol - 1] || "").trim() : "";
      var cellStock = stockCol !== -1 ? normalizeStockId(row[stockCol - 1]) : "";
      var cellMenge = mengeCol !== -1 ? (parseInt(row[mengeCol - 1], 10) || 1) : 1;
      if (!cellStock) continue;

      results.push({
        tabName: tabName,
        stockId: cellStock,
        groesse: cellGroesse,
        lastindex: cellLast,
        gwIndex: cellGw,
        menge: cellMenge,
        row: startRow + i
      });
    }
  }

function sortPaketdienstResults_(results) {
    results.sort(function(a, b) {
      var ma = String(a.tabName).match(/(\d{2})\.(\d{2})\.(\d{4})$/);
      var mb = String(b.tabName).match(/(\d{2})\.(\d{2})\.(\d{4})$/);
      if (ma && mb) {
        var da = new Date(parseInt(ma[3], 10), parseInt(ma[2], 10) - 1, parseInt(ma[1], 10)).getTime();
        var db = new Date(parseInt(mb[3], 10), parseInt(mb[2], 10) - 1, parseInt(mb[1], 10)).getTime();
        return db - da;
      }
      return 0;
    });
  }

function collectPaketdienstReifenRowsForTabNames(tabNames) {
    var results = [];
    var si;
    for (si = 0; si < tabNames.length; si++) {
      var sh = getReifenSheetTab(tabNames[si]);
      if (!sh) continue;
      appendPaketdienstRowsFromSheet_(sh, results);
    }
    sortPaketdienstResults_(results);
    return results;
  }

function collectPaketdienstReifenRows() {
    return collectPaketdienstReifenRowsForTabNames(getReifenTabNamesForPaketHudWindow());
  }

function searchReifenBySize(query) {
    try {
      query = String(query || "").trim();
      if (!query || query.length < 3) return { success: false, message: "Suchbegriff zu kurz (min. 3 Zeichen)", results: [] };

      var queryParts = query.replace(/\s+/g, ' ').split(' ');
      var sizeQuery = "";
      var lastQuery = "";
      var gwQuery = "";
      for (var q = 0; q < queryParts.length; q++) {
        var p = queryParts[q];
        if (p.indexOf('/') !== -1 || /^r\d+/i.test(p) || /^\d{3}\//.test(p)) {
          sizeQuery += (sizeQuery ? " " : "") + p;
        } else if (/^\d{2,3}$/.test(p) && !lastQuery) {
          lastQuery = p;
        } else if (/^[A-Za-z]{1,2}$/.test(p) && !gwQuery) {
          gwQuery = p.toUpperCase();
        } else {
          sizeQuery += (sizeQuery ? " " : "") + p;
        }
      }
      if (!sizeQuery) sizeQuery = query;
      var sizeNorm = sizeQuery.replace(/\s+/g, '').toUpperCase();

      var all = collectPaketdienstReifenRows();
      var results = [];
      for (var ri = 0; ri < all.length; ri++) {
        var item = all[ri];
        var cellGroesseNorm = String(item.groesse || "").replace(/\s+/g, '').toUpperCase();
        if (cellGroesseNorm.indexOf(sizeNorm) === -1) continue;
        var cellLast = String(item.lastindex || "").trim();
        var cellGw = String(item.gwIndex || "").trim();
        if (lastQuery && cellLast.indexOf(lastQuery) === -1) continue;
        if (gwQuery && cellGw.toUpperCase().indexOf(gwQuery) === -1) continue;
        results.push(item);
      }

      return { success: true, results: results, message: results.length + " Reifen gefunden" };
    } catch (err) {
      return { success: false, message: "Fehler: " + err.message, results: [] };
    }
  }

function getPaketdienstReifenCachePayload() {
    try {
      var rows = collectPaketdienstReifenRows();
      return { success: true, version: Date.now(), paketdienstRows: rows };
    } catch (err) {
      return { success: false, message: err.message, paketdienstRows: [] };
    }
  }

function getReifenWindowStocksCachePayload() {
    try {
      var tabsRes = getReifenTabOptions();
      if (!tabsRes.success) {
        return { success: false, message: tabsRes.message || "Tabs", tabs: [], stocksByTab: {} };
      }
      var tabs = tabsRes.tabs || [];
      var stocksByTab = {};
      var ti;
      for (ti = 0; ti < tabs.length; ti++) {
        var tn = tabs[ti];
        var idRes = getAvailableReifenStockIds(tn);
        stocksByTab[tn] = (idRes.success && idRes.ids) ? idRes.ids : [];
      }
      return { success: true, version: Date.now(), tabs: tabs, stocksByTab: stocksByTab };
    } catch (err) {
      return { success: false, message: err.message, tabs: [], stocksByTab: {} };
    }
  }

function getAvailableReifenStockIds(tabName) {
    try {
      var sheet = getReifenSheetTab(tabName);
      if (!sheet) return { success: false, message: "Tabellenblatt nicht gefunden!", ids: [] };
      var search = findRowFast(sheet, ["stockid", "stock"], "___NEVER_MATCH___");
      if (search.headerIdx === -1 || search.stockCol === -1) {
        return { success: false, message: "Kopfzeile 'Stock ID' nicht gefunden!", ids: [] };
      }

      var lastRow = Math.max(1, sheet.getLastRow());
      var startRow = search.headerIdx + 2;
      var numRows = lastRow - startRow + 1;
      if (numRows <= 0) return { success: true, ids: [] };

      var headerRow = sheet.getRange(search.headerIdx + 1, 1, 1, Math.max(1, Math.min(80, sheet.getLastColumn()))).getValues()[0];
      var angeliefertCol = getColIndex(headerRow, ["angeliefert"]);
      var groesseCol = getColIndex(headerRow, ["größe", "groesse"]);
      var lastIndexCol = getColIndex(headerRow, ["lastindex", "last"]);
      var gwIndexCol = getColIndex(headerRow, ["gwindex", "gw"]);
      var mengeCol = getColIndex(headerRow, ["menge", "anzahl"]);
      var stockData = sheet.getRange(startRow, search.stockCol, numRows, 1).getValues();
      var statusData = angeliefertCol !== -1 ? sheet.getRange(startRow, angeliefertCol, numRows, 1).getValues() : [];
      var groesseData = groesseCol !== -1 ? sheet.getRange(startRow, groesseCol, numRows, 1).getValues() : [];
      var lastIndexData = lastIndexCol !== -1 ? sheet.getRange(startRow, lastIndexCol, numRows, 1).getValues() : [];
      var gwIndexData = gwIndexCol !== -1 ? sheet.getRange(startRow, gwIndexCol, numRows, 1).getValues() : [];
      var mengeData = mengeCol !== -1 ? sheet.getRange(startRow, mengeCol, numRows, 1).getValues() : [];
      var ids = [];
      for (var i = 0; i < stockData.length; i++) {
        var val = normalizeStockId(stockData[i][0]);
        if (val) {
          var mengeVal = mengeCol !== -1 ? (parseInt(mengeData[i][0], 10) || 0) : 0;
          if (mengeVal < 1) mengeVal = 2;
          ids.push({
            id: val,
            status: angeliefertCol !== -1 ? String(statusData[i][0] || "").trim().toLowerCase() : "",
            groesse: groesseCol !== -1 ? String(groesseData[i][0] || "").trim() : "",
            lastindex: lastIndexCol !== -1 ? String(lastIndexData[i][0] || "").trim() : "",
            gwindex: gwIndexCol !== -1 ? String(gwIndexData[i][0] || "").trim() : "",
            menge: mengeVal
          });
        }
      }
      return { success: true, ids: ids };
    } catch (err) {
      return { success: false, message: err.message, ids: [] };
    }
  }

function checkReifenStock(tabName, stockId) {
    try {
      stockId = normalizeStockId(stockId);
      if (!stockId) return { found: false, message: "Bitte eine Stock-ID eingeben." };

      var sheet = getReifenSheetTab(tabName);
      if (!sheet) return { found: false, message: "Bitte ein gültiges Tabellenblatt auswählen." };

      var det = findReifenStockRowsDetailed(sheet, stockId);
      if (det.headerIdx === -1) return { found: false, message: "Kopfzeile 'Stock ID' in Reifenliste nicht gefunden!" };
      if (!det.matches.length) return { found: false, message: "Stock-ID '" + stockId + "' in '" + sheet.getName() + "' nicht gefunden!" };
      var pick = pickReifenUnbookedRow(det.matches, null);
      if (!pick.chosen) {
        return { found: false, message: "Stock-ID '" + stockId + "' wurde in '" + sheet.getName() + "' bereits verbucht!" };
      }
      return { found: true, message: "Stock-ID gefunden! Bitte Status auswählen:" };
    } catch (err) {
      return { found: false, message: "Systemfehler: " + err.message };
    }
  }

function formatReifenRegalLabel(val) {
    var s = String(val || "").trim();
    if (!s || String(s).toLowerCase() === "tagesliste") return "";
    var m = s.match(/^Regal\s+(\d+)\.(\d+)$/i);
    if (m) return "Regal " + m[1] + "." + m[2];
    m = s.match(/^(\d+)\.(\d+)$/);
    if (m) return "Regal " + m[1] + "." + m[2];
    return "";
  }

function processReifenStock(tabName, stockId, isDelivered, sizeHint) {
    try {
      stockId = normalizeStockId(stockId);
      var sheetSeng = getReifenSheetTab(tabName);
      if (!sheetSeng) return { success: false, message: "Bitte ein gültiges Tabellenblatt auswählen." };

      var sheetHemau = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
      if (!sheetHemau) return { success: false, message: "Reiter 'Refurbisment List' fehlt!" };

      var tireInfo = "UNBEKANNT _X";
      var mengeValNum = 1;
      var det = findReifenStockRowsDetailed(sheetSeng, stockId);
      if (det.headerIdx === -1) return { success: false, message: "Kopfzeile 'Stock ID' in Reifenliste nicht gefunden!" };
      if (!det.matches.length) return { success: false, message: "Stock-ID '" + stockId + "' in '" + sheetSeng.getName() + "' nicht gefunden!" };
      var pick = pickReifenUnbookedRow(det.matches, sizeHint);
      if (!pick.chosen) return { success: false, message: "Stock-ID '" + stockId + "' wurde in '" + sheetSeng.getName() + "' bereits verbucht!" };

      var search = { row: pick.chosen.row, headerIdx: det.headerIdx, stockCol: det.stockCol };
      var headerRow = det.header;
      var angeliefertCol = det.angeliefertCol;
      var mengeCol = getColIndex(headerRow, ["menge", "anzahl"]);
      var groesseCol = getColIndex(headerRow, ["größe", "groesse"]);
      var lastIndexCol = getColIndex(headerRow, ["lastindex", "last"]);
      var gwIndexCol = getColIndex(headerRow, ["gwindex", "gw"]);

      var remaining = [];
      for (var ri = 0; ri < pick.unbooked.length; ri++) {
        if (pick.unbooked[ri].row !== search.row) remaining.push(pick.unbooked[ri]);
      }
      var remainingNext = remaining.length ? remaining[0] : null;

      var mengeVal = mengeCol !== -1 ? sheetSeng.getRange(search.row, mengeCol).getValue() : "1";
      mengeValNum = parseInt(mengeVal, 10) || 1;
      var groesseVal = groesseCol !== -1 ? sheetSeng.getRange(search.row, groesseCol).getValue() : "GRÖSSE";
      var lastIndexVal = lastIndexCol !== -1 ? sheetSeng.getRange(search.row, lastIndexCol).getValue() : "";
      var gwIndexVal = gwIndexCol !== -1 ? sheetSeng.getRange(search.row, gwIndexCol).getValue() : "";
      tireInfo = String(groesseVal).trim() + " " + String(lastIndexVal).trim() + String(gwIndexVal).trim() + " _" + String(mengeVal).trim();

      if (angeliefertCol !== -1) {
        var statusSeng = isDelivered ? "Ja" : "Nein";
        var colorSeng = isDelivered ? "#00FF00" : "#FF0000";
        sheetSeng.getRange(search.row, angeliefertCol).setValue(statusSeng);
        var startCol = Math.min(search.stockCol, angeliefertCol);
        var numCols = Math.abs(angeliefertCol - search.stockCol) + 1;
        sheetSeng.getRange(search.row, startCol, 1, numCols).setBackground(colorSeng);
      }

      var hemauDataStock = sheetHemau.getRange(1, 2, Math.max(1, sheetHemau.getLastRow()), 1).getValues();
      var hemauRow = -1;
      for (var i = 0; i < hemauDataStock.length; i++) {
        if (cellMatchesStockId(hemauDataStock[i][0], stockId)) {
          hemauRow = i + 1;
          break;
        }
      }

      var hemauMsg = "";
      var locationText = "Lagerplatz unbekannt";
      var regalLabel = "";
      if (hemauRow !== -1) {
        var oldLocation = String(sheetHemau.getRange(hemauRow, 28).getValue() || "").trim();
        regalLabel = formatReifenRegalLabel(oldLocation);
        if (regalLabel) locationText = "Kiste steht in " + regalLabel;
        else if (oldLocation !== "") locationText = "Kiste steht in " + oldLocation;

        var currentComment = String(sheetHemau.getRange(hemauRow, 25).getValue() || "");
        if (isDelivered) {
          if (currentComment.indexOf("Reifen da //") === -1) {
            var newComment = currentComment ? "Reifen da // " + currentComment : "Reifen da // ";
            sheetHemau.getRange(hemauRow, 25).setValue(newComment);
          }
          sheetHemau.getRange(hemauRow, 30).setValue("Werkstatt 1");
          hemauMsg = "Reifen als da gebucht & Refurbishment aktualisiert!";
        } else {
          sheetHemau.getRange(hemauRow, 30).setValue("Reifen nicht vorhanden");
          hemauMsg = "Reifen als fehlend gebucht & Refurbishment aktualisiert!";
        }
      } else {
        hemauMsg = "Reifen gebucht (Refurbishment übersprungen: ID nicht gefunden)";
      }

      SpreadsheetApp.flush();
      return {
        success: true,
        message: hemauMsg,
        stockId: stockId,
        tireInfo: tireInfo,
        locationText: locationText,
        regal: regalLabel,
        menge: mengeValNum,
        remainingUnbooked: remaining.length,
        remainingSize: remainingNext ? remainingNext.groesse : "",
        remainingLast: remainingNext ? remainingNext.lastindex : "",
        remainingGw: remainingNext ? remainingNext.gwindex : "",
        remainingMenge: remainingNext ? (parseInt(remainingNext.menge, 10) || 2) : 0
      };
    } catch (err) {
      return { success: false, message: "Fehler: " + err.message };
    }
  }

function extractHuDateOnlyFromNachuntersuchungSegment(segment) {
  var s = String(segment || "").trim();
  if (!s) return "";
  var mNum = s.match(/(\d{1,2}\.\s*\d{1,2}\.\s*\d{2,4})/);
  if (mNum) {
    var p = mNum[1].match(/(\d{1,2})\s*\.\s*(\d{1,2})\s*\.\s*(\d{2,4})/);
    if (p) return p[1] + "." + p[2] + "." + p[3];
  }
  var reDotMo = /(\d{1,2}\.\s*(?:Jan|Feb|Mär|Mrz|Apr|Mai|Jun|Jul|Aug|Sep|Okt|Nov|Dez|Januar|Februar|März|April|Juni|Juli|September|Oktober|November|Dezember)[a-zä]*\.?\s*\d{4})/i;
  var mDot = s.match(reDotMo);
  if (mDot) return mDot[1].replace(/\s+/g, " ").trim();
  var reSpMo = /(\d{1,2}\s+(?:Jan|Feb|Mär|Mrz|Apr|Mai|Jun|Jul|Aug|Sep|Okt|Nov|Dez|Januar|Februar|März|April|Juni|Juli|September|Oktober|November|Dezember)[a-zä]*\.?\s*\d{4})/i;
  var mSp = s.match(reSpMo);
  if (mSp) return mSp[1].replace(/\s+/g, " ").trim();
  return "";
}

function huVasoldValueFromSchaedenText(wText) {
  var dash = "---------------";
  var m = String(wText || "").match(/Nachuntersuchung\s*bis\s*:\s*([^\n\r]+)/i);
  if (!m) return dash;
  var rest = String(m[1] || "").trim();
  if (!rest) return dash;
  var dateOnly = extractHuDateOnlyFromNachuntersuchungSegment(rest);
  if (!dateOnly) return dash;
  return dateOnly;
}

function processWssVasoldBooking(stockId, carolUrlOpt, markeOpt, gummiVorhanden) {
  try {
    stockId = normalizeStockId(stockId);
    if (!stockId) return { success: false, message: "Keine Stock-ID" };
    carolUrlOpt = String(carolUrlOpt || "").trim();
    markeOpt = String(markeOpt || "").trim();
    var gummiYes = gummiVorhanden === true || gummiVorhanden === "true";

    var sheetRef = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
    if (!sheetRef) return { success: false, message: "Reiter 'Refurbisment List' fehlt!" };

    var lastRef = Math.max(2, sheetRef.getLastRow());
    var stockCol = sheetRef.getRange(1, 2, lastRef, 1).getValues();
    var refurbRow = -1;
    for (var ir = 1; ir < stockCol.length; ir++) {
      if (cellMatchesStockId(stockCol[ir][0], stockId)) {
        refurbRow = ir + 1;
        break;
      }
    }
    if (refurbRow === -1) return { success: false, message: "Stock-ID in Refurbisment List nicht gefunden!" };

    var wText = String(sheetRef.getRange(refurbRow, 23).getValue() || "");
    if (!/wss\s*ers/i.test(wText)) {
      return { success: false, message: "Kein „WSS ers“ in Schaden (Spalte W)." };
    }

    var carol = carolUrlOpt || String(sheetRef.getRange(refurbRow, 3).getValue() || "").trim();
    var marke = markeOpt || String(sheetRef.getRange(refurbRow, 13).getValue() || "").trim();
    if (!carol) return { success: false, message: "Carol-Link fehlt oder konnte nicht gelesen werden — bitte aus Carol kopieren und im Dialog eintragen." };
    if (!marke) return { success: false, message: "Marke & Model fehlt oder konnte nicht gelesen werden — bitte aus Carol kopieren und im Dialog eintragen." };

    var huVasold = huVasoldValueFromSchaedenText(wText);

    var ssV = SpreadsheetApp.openById(TAGESLISTE_SHEET_ID);
    var sheetV = ssV.getSheetByName(VASOLD_WSS_TAB);
    if (!sheetV) return { success: false, message: "Tab „" + VASOLD_WSS_TAB + "“ in Hemau Tageslisten nicht gefunden!" };

    var searchV = findRowFast(sheetV, ["stockid", "stock"], stockId);
    if (searchV.headerIdx === -1) {
      return { success: false, message: "Vasold WSS: Kopfzeile mit Stock-ID nicht gefunden!" };
    }

    var hdr1 = searchV.headerIdx + 1;
    var targetRow;
    var isNew = false;
    if (searchV.row === -1) {
      isNew = true;
      var lrV = sheetV.getLastRow();
      targetRow = Math.max(lrV, hdr1) + 1;
    } else {
      targetRow = searchV.row;
    }

    sheetV.getRange(targetRow, 1).setValue(stockId);
    sheetV.getRange(targetRow, 2).setValue(carol);
    sheetV.getRange(targetRow, 3).setValue(marke);
    sheetV.getRange(targetRow, 4).setValue(huVasold);
    sheetV.getRange(targetRow, 5).setValue("Ja");
    if (gummiYes) sheetV.getRange(targetRow, 6).setValue("Vorhanden");

    var curCom = String(sheetRef.getRange(refurbRow, 25).getValue() || "");
    if (!/wss\s+da\b/i.test(curCom)) {
      var wssCom = curCom ? "WSS da // " + curCom : "WSS da // ";
      sheetRef.getRange(refurbRow, 25).setValue(wssCom);
      logKommentarVerlauf_(stockId, curCom, wssCom, "wss");
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      message: (isNew ? "Vasold WSS: neue Zeile. " : "Vasold WSS: Zeile aktualisiert. ") + "Kommentar Anlieferung: WSS da //",
      stockId: stockId,
      isNew: isNew,
      row: targetRow
    };
  } catch (err) {
    return { success: false, message: "Fehler: " + err.message };
  }
}

function getVasoldWssSyncState(stockId) {
  try {
    stockId = normalizeStockId(stockId);
    if (!stockId) return { success: false, rowFound: false, message: "Keine Stock-ID" };

    var ssV = SpreadsheetApp.openById(TAGESLISTE_SHEET_ID);
    var sheetV = ssV.getSheetByName(VASOLD_WSS_TAB);
    if (!sheetV) return { success: false, rowFound: false, message: "Tab fehlt" };

    var searchV = findRowFast(sheetV, ["stockid", "stock"], stockId);
    if (searchV.headerIdx === -1) {
      return { success: true, rowFound: false, frontscheibeJa: false, gummileisteDa: false };
    }
    if (searchV.row === -1) {
      return { success: true, rowFound: false, frontscheibeJa: false, gummileisteDa: false };
    }

    var eRaw = sheetV.getRange(searchV.row, 5).getValue();
    var fRaw = sheetV.getRange(searchV.row, 6).getValue();
    var eVal = String(eRaw != null ? eRaw : "").trim().toLowerCase();
    var fVal = String(fRaw != null ? fRaw : "").trim().toLowerCase();

    var frontscheibeJa = eVal === "ja";
    var gummileisteDa = fVal === "vorhanden" || fVal === "da" || fVal.indexOf("vorhanden") !== -1;

    return {
      success: true,
      rowFound: true,
      frontscheibeJa: frontscheibeJa,
      gummileisteDa: gummileisteDa
    };
  } catch (err) {
    return { success: false, rowFound: false, message: err.message };
  }
}

function processNachbestellVasoldWss(stockId, toggleWssJa, toggleGummi) {
  try {
    stockId = normalizeStockId(stockId);
    if (!stockId) return { success: false, message: "Keine Stock-ID" };
    var wssYes = toggleWssJa === true || toggleWssJa === "true";
    var gummiYes = toggleGummi === true || toggleGummi === "true";
    if (!wssYes && !gummiYes) {
      return { success: false, message: "Keine Option zum Übernehmen gewählt." };
    }

    var sheetRef = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
    if (!sheetRef) return { success: false, message: "Reiter 'Refurbisment List' fehlt!" };

    var lastRef = Math.max(2, sheetRef.getLastRow());
    var stockCol = sheetRef.getRange(1, 2, lastRef, 1).getValues();
    var refurbRow = -1;
    for (var ir = 1; ir < stockCol.length; ir++) {
      if (cellMatchesStockId(stockCol[ir][0], stockId)) {
        refurbRow = ir + 1;
        break;
      }
    }
    if (refurbRow === -1) return { success: false, message: "Stock-ID in Refurbisment List nicht gefunden!" };

    var wText = String(sheetRef.getRange(refurbRow, 23).getValue() || "");
    var carol = String(sheetRef.getRange(refurbRow, 3).getValue() || "").trim();
    var marke = String(sheetRef.getRange(refurbRow, 13).getValue() || "").trim();
    if (!carol) return { success: false, message: "Carol-Link fehlt in Refurbishment." };
    if (!marke) return { success: false, message: "Marke & Model fehlt in Refurbishment." };

    var huVasold = huVasoldValueFromSchaedenText(wText);

    var ssV = SpreadsheetApp.openById(TAGESLISTE_SHEET_ID);
    var sheetV = ssV.getSheetByName(VASOLD_WSS_TAB);
    if (!sheetV) return { success: false, message: "Tab „" + VASOLD_WSS_TAB + "“ in Hemau Tageslisten nicht gefunden!" };

    var searchV = findRowFast(sheetV, ["stockid", "stock"], stockId);
    if (searchV.headerIdx === -1) {
      return { success: false, message: "Vasold WSS: Kopfzeile mit Stock-ID nicht gefunden!" };
    }

    var hdr1 = searchV.headerIdx + 1;
    var targetRow;
    var isNew = false;
    if (searchV.row === -1) {
      isNew = true;
      var lrV = sheetV.getLastRow();
      targetRow = Math.max(lrV, hdr1) + 1;
    } else {
      targetRow = searchV.row;
    }

    var eRaw = sheetV.getRange(targetRow, 5).getValue();
    var fRaw = sheetV.getRange(targetRow, 6).getValue();
    var eVal = String(eRaw != null ? eRaw : "").trim().toLowerCase();
    var fVal = String(fRaw != null ? fRaw : "").trim().toLowerCase();
    var eHasJa = eVal === "ja";
    var fHasGummi = fVal === "vorhanden" || fVal === "da" || fVal.indexOf("vorhanden") !== -1;

    var needWssWrite = wssYes && (isNew || !eHasJa);
    var needGummiWrite = gummiYes && (isNew || !fHasGummi);
    if (!needWssWrite && !needGummiWrite) {
      return {
        success: true,
        skipped: true,
        message: "Vasold hatte die gewählten Werte bereits — nichts geändert."
      };
    }

    sheetV.getRange(targetRow, 1).setValue(stockId);
    sheetV.getRange(targetRow, 2).setValue(carol);
    sheetV.getRange(targetRow, 3).setValue(marke);
    sheetV.getRange(targetRow, 4).setValue(huVasold);
    if (needWssWrite) {
      sheetV.getRange(targetRow, 5).setValue("Ja");
    }
    if (needGummiWrite) {
      sheetV.getRange(targetRow, 6).setValue("Vorhanden");
    }

    var wroteEJa = needWssWrite;
    if (wroteEJa) {
      var curCom = String(sheetRef.getRange(refurbRow, 25).getValue() || "");
      if (!/wss\s+da\b/i.test(curCom)) {
        var wssCom2 = curCom ? "WSS da // " + curCom : "WSS da // ";
        sheetRef.getRange(refurbRow, 25).setValue(wssCom2);
        logKommentarVerlauf_(stockId, curCom, wssCom2, "wss");
      }
    }

    SpreadsheetApp.flush();
    var parts = [];
    if (needWssWrite) parts.push("WSS Ja");
    if (needGummiWrite) parts.push("Gummileiste");
    return {
      success: true,
      skipped: false,
      message: "Vasold WSS: " + parts.join(", ") + (isNew ? " (neue Zeile)" : " — Zeile aktualisiert"),
      stockId: stockId,
      isNew: isNew,
      row: targetRow
    };
  } catch (err) {
    return { success: false, message: "Fehler: " + err.message };
  }
}

  function onOpen() {
    SpreadsheetApp.getUi().createMenu('WMS')
      .addItem('Öffne Warehouse Management System', 'openWMS')
      .addToUi();
  }
  
  function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  }

  function compareWmsAppVersion_(a, b) {
    var pa = String(a || "").split(".");
    var pb = String(b || "").split(".");
    var n = Math.max(pa.length, pb.length);
    for (var i = 0; i < n; i++) {
      var x = parseInt(pa[i], 10);
      var y = parseInt(pb[i], 10);
      if (isNaN(x)) x = 0;
      if (isNaN(y)) y = 0;
      if (x > y) return 1;
      if (x < y) return -1;
    }
    return 0;
  }

  function getWmsWebAppUrl_() {
    return (WMS_WEB_APP_URL && String(WMS_WEB_APP_URL).trim()) || ScriptApp.getService().getUrl() || "";
  }

  function checkWmsAppUpdate(clientVersion) {
    var serverVersion = String(WMS_APP_VERSION || "");
    return {
      updateAvailable: compareWmsAppVersion_(serverVersion, clientVersion) > 0,
      latestVersion: serverVersion,
      clientVersion: String(clientVersion || ""),
      url: getWmsWebAppUrl_(),
      changelog: String(WMS_APP_CHANGELOG || "")
    };
  }

  function doGet(e) {
    var p = e && e.parameter || {};
    var mode = String(p.mode || 'standalone').toLowerCase();
    if (mode !== 'overlay' && mode !== 'standalone') mode = 'standalone';
    var t = HtmlService.createTemplateFromFile('WMS_App');
    t.mode = mode;
    t.appVersion = String(WMS_APP_VERSION || "");
    t.appUrl = getWmsWebAppUrl_();
    t.cacheNonce = String(Date.now());
    return t.evaluate()
      .setTitle('Warehouse Management System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  function openWMS() {
    var url = getWmsWebAppUrl_();
    if (!url) {
      SpreadsheetApp.getUi().alert("Keine Web-App-URL. Bitte WMS_WEB_APP_URL setzen oder Web-App bereitstellen.");
      return;
    }
    var html = HtmlService.createHtmlOutput(
      "<!DOCTYPE html><html><body style=\"margin:0\"><script>" +
      "window.onload=function(){window.open(" + JSON.stringify(url) + ",\"_blank\");google.script.host.close();};" +
      "</script></body></html>"
    );
    SpreadsheetApp.getUi().showModalDialog(html, "Warehouse Management System");
  }

  function applyTrackingDateIfEmpty(stockId) {
    stockId = normalizeStockId(stockId);
    var trackingSs = SpreadsheetApp.openByUrl(TRACKING_SHEET_URL);
    var sheet = trackingSs.getSheetByName("Stock ID extern Tracking");
    if (!sheet) return { success: false, updated: false, message: "Reiter 'Stock ID extern Tracking' fehlt!" };

    var lastRow = Math.max(2, sheet.getLastRow());
    var data = sheet.getRange(1, 1, lastRow, 9).getValues();

    for (var i = 0; i < data.length; i++) {
      if (cellMatchesStockId(data[i][0], stockId)) {
        var currentDate = String(data[i][8] || "").trim();
        if (currentDate !== "") {
          return { success: true, updated: false, message: "Datum bereits gesetzt" };
        }
        var dateStr = Utilities.formatDate(new Date(), "Europe/Berlin", "dd.MM.yyyy");
        sheet.getRange(i + 1, 9).setValue(dateStr);
        SpreadsheetApp.flush();
        return { success: true, updated: true, message: "Datum gesetzt" };
      }
    }

    return { success: false, updated: false, message: "Stock-ID in Stock ID extern Tracking nicht gefunden!" };
  }
  
  function searchOrderNumberInGmailLookup_(query) {
    query = String(query || "").replace(/\s+/g, "").toUpperCase();
    if (!query) return { found: false, message: "Keine Suchanfrage" };

    var cleanQuery = query.replace(/^N4P/i, "");
    if (!cleanQuery || cleanQuery.length < 4) {
      return { found: false, message: "Suchanfrage zu kurz (min. 4 Zeichen)" };
    }

    var ss = SpreadsheetApp.openById(GMAIL_LOOKUP_SHEET_ID);
    var sheet = ss.getSheetByName(GMAIL_LOOKUP_TAB);
    if (!sheet || sheet.getLastRow() < 2) {
      return { found: false, message: "Gmail Lookup leer — Sync auf lager.hemau starten." };
    }

    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(2, 1, lastRow, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      var key = String(data[i][0] || "").replace(/\s+/g, "").toUpperCase();
      if (!key || (key !== query && key !== cleanQuery)) continue;
      var stockId = String(data[i][1] || "").trim();
      if (!stockId) continue;
      return {
        found: true,
        stockId: normalizeStockId(stockId),
        message: "Gefunden via Gmail Lookup",
        source: "gmail"
      };
    }

    return { found: false, message: "Bestellnummer '" + query + "' nicht in Gmail Lookup gefunden." };
  }

  function testGmailLookupFromWms() {
    var result = searchOrderNumberInGmailLookup_("2611233477");
    Logger.log(JSON.stringify(result, null, 2));
    return result;
  }

  function searchByOrderNumber(query) {
    try {
      query = String(query || "").replace(/\s+/g, '').toUpperCase();
      if (!query) return { found: false, message: "Keine Suchanfrage" };

      var cleanQuery = query.replace(/^N4P/i, '');
      if (!cleanQuery || cleanQuery.length < 4) return { found: false, message: "Suchanfrage zu kurz (min. 4 Zeichen)" };

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("Refurbisment List");
      if (!sheet) return { found: false, message: "Reiter 'Refurbisment List' fehlt!" };

      var lastRow = Math.max(2, sheet.getLastRow());
      var kommData = sheet.getRange(1, 24, lastRow, 1).getValues();
      var stockData = sheet.getRange(1, 2, lastRow, 1).getValues();

      for (var i = 1; i < kommData.length; i++) {
        var cellText = String(kommData[i][0] || "").replace(/\s+/g, '').toUpperCase();
        if (!cellText) continue;
        if (cellText.indexOf(cleanQuery) !== -1 || cellText.indexOf(query) !== -1) {
          var stockId = String(stockData[i][0] || "").trim();
          if (stockId) {
            return { found: true, stockId: stockId, message: "Gefunden via Bestellnummer in Zeile " + (i + 1), source: "sheet" };
          }
        }
      }

      var gmailHit = searchOrderNumberInGmailLookup_(query);
      if (gmailHit && gmailHit.found && gmailHit.stockId) return gmailHit;

      return { found: false, message: "Bestellnummer '" + query + "' weder im Sheet noch in Gmail Lookup gefunden." };
    } catch (err) {
      return { found: false, message: "Fehler: " + err.message };
    }
  }

  function fetchWmsData(stockId) {
    try {
      stockId = normalizeStockId(stockId);
      if (!stockId) return { success: false, message: "Keine Stock-ID" };
  
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("Refurbisment List");
      if (!sheet) return { success: false, message: "Reiter 'Refurbisment List' fehlt!" };
  
      var lastRow = Math.max(2, sheet.getLastRow());
      var stockColData = sheet.getRange(1, 2, lastRow, 1).getValues();
      var result = { success: false };
      var hitRow = -1;

      for (var s = 1; s < stockColData.length; s++) {
        if (cellMatchesStockId(stockColData[s][0], stockId)) {
          hitRow = s + 1;
          break;
        }
      }
  
      if (hitRow > 0) {
        var rowData = sheet.getRange(hitRow, 1, 1, 30).getValues()[0];
        result.success = true;
        result.carolUrl = String(rowData[2] || "");
        result.schaeden = String(rowData[22] || "");
        result.kommBestellung = String(rowData[23] || "");
        result.kommAnlieferung = String(rowData[24] || "");
        result.status = String(rowData[25] || "");
        result.regal = String(rowData[27] || "");
        result.reifenStatus = String(rowData[29] || "");
        result.markeModel = String(rowData[12] || "");
      } else {
        result.message = "Stock-ID in Refurbisment List nicht gefunden!";
      }
  
      return result;
    } catch (err) {
      return { success: false, message: err.message };
    }
  }

function getActiveUserEmail_() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (email) return email;
    return Session.getEffectiveUser().getEmail() || "unbekannt";
  } catch (e) {
    return "unbekannt";
  }
}

var WMS_STOCK_LOCK_TTL_SEC = 60;
var WMS_STOCK_LOCK_STALE_MS = 45000;
var WMS_STOCK_LOCK_PREFIX = "wmsLock:";

function stockEditLockKey_(stockId) {
  return WMS_STOCK_LOCK_PREFIX + normalizeStockId(stockId);
}

function parseStockEditLock_(raw) {
  var s = String(raw || "");
  if (!s) return null;
  var pipe = s.indexOf("|");
  if (pipe < 0) return { email: s, ts: 0 };
  return {
    email: s.substring(0, pipe),
    ts: parseInt(s.substring(pipe + 1), 10) || 0
  };
}

function encodeStockEditLock_(email) {
  return String(email || "unbekannt") + "|" + Date.now();
}

function isStockEditLockStale_(existing) {
  if (!existing) return true;
  if (!existing.ts) return true;
  return (Date.now() - existing.ts) > WMS_STOCK_LOCK_STALE_MS;
}

function claimStockEditLock(stockId) {
  try {
    stockId = normalizeStockId(stockId);
    if (!stockId) return { success: false, message: "Keine Stock-ID" };
    var email = getActiveUserEmail_();
    var cache = CacheService.getScriptCache();
    var key = stockEditLockKey_(stockId);
    var existing = parseStockEditLock_(cache.get(key));
    if (existing && existing.email && existing.email !== email && !isStockEditLockStale_(existing)) {
      return { success: true, occupied: true, email: existing.email, stockId: stockId };
    }
    cache.put(key, encodeStockEditLock_(email), WMS_STOCK_LOCK_TTL_SEC);
    return { success: true, occupied: false, email: email, stockId: stockId };
  } catch (err) {
    return { success: false, message: err.message || String(err) };
  }
}

function heartbeatStockEditLock(stockId) {
  try {
    stockId = normalizeStockId(stockId);
    if (!stockId) return { success: false, message: "Keine Stock-ID" };
    var email = getActiveUserEmail_();
    var cache = CacheService.getScriptCache();
    var key = stockEditLockKey_(stockId);
    var existing = parseStockEditLock_(cache.get(key));
    if (existing && existing.email && existing.email !== email && !isStockEditLockStale_(existing)) {
      return { success: true, occupied: true, email: existing.email, stockId: stockId };
    }
    cache.put(key, encodeStockEditLock_(email), WMS_STOCK_LOCK_TTL_SEC);
    return { success: true, occupied: false, email: email, stockId: stockId };
  } catch (err) {
    return { success: false, message: err.message || String(err) };
  }
}

function releaseStockEditLock(stockId) {
  try {
    stockId = normalizeStockId(stockId);
    if (!stockId) return { success: false, message: "Keine Stock-ID" };
    var email = getActiveUserEmail_();
    var cache = CacheService.getScriptCache();
    var key = stockEditLockKey_(stockId);
    var existing = parseStockEditLock_(cache.get(key));
    if (existing && existing.email && existing.email !== email && !isStockEditLockStale_(existing)) {
      return { success: true, released: false, email: existing.email, stockId: stockId };
    }
    cache.remove(key);
    return { success: true, released: true, stockId: stockId };
  } catch (err) {
    return { success: false, message: err.message || String(err) };
  }
}

function buildConcurrencyConflict_(currentKommentar, currentRegal) {
  return {
    success: false,
    conflict: true,
    currentKommentar: String(currentKommentar == null ? "" : currentKommentar),
    currentRegal: String(currentRegal == null ? "" : currentRegal).trim(),
    message: "Konflikt: Ein Kollege hat Kommentar/Regal bereits geändert. Daten neu geladen – bitte prüfen und erneut speichern."
  };
}

function baselinesMatch_(sheetKommentar, sheetRegal, expectedKommentar, expectedRegal) {
  if (expectedKommentar === undefined && expectedRegal === undefined) return true;
  var expC = String(expectedKommentar == null ? "" : expectedKommentar);
  var expR = String(expectedRegal == null ? "" : expectedRegal).trim();
  var curC = String(sheetKommentar == null ? "" : sheetKommentar);
  var curR = String(sheetRegal == null ? "" : sheetRegal).trim();
  return expC === curC && expR === curR;
}

function withRefurbDocumentLock_(fn) {
  var lock = LockService.getDocumentLock();
  var got = false;
  try {
    got = lock.tryLock(15000);
    if (!got) {
      return { success: false, message: "Speichern beschäftigt – bitte kurz erneut versuchen" };
    }
    return fn();
  } catch (err) {
    return { success: false, message: "Fehler: " + (err.message || String(err)) };
  } finally {
    if (got) {
      try { lock.releaseLock(); } catch (e) {}
    }
  }
}

function formatVerlaufTimestamp_(date) {
  return Utilities.formatDate(date || new Date(), "Europe/Berlin", "dd.MM.yyyy HH:mm");
}

function verlaufTimestampStr_(v) {
  if (v instanceof Date || Object.prototype.toString.call(v) === "[object Date]") {
    return formatVerlaufTimestamp_(v);
  }
  var s = String(v || "").trim();
  if (!s) return "";
  if (/^\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}$/.test(s)) return s;
  try {
    var d = new Date(s);
    if (!isNaN(d.getTime())) return formatVerlaufTimestamp_(d);
  } catch (e) {}
  return s.replace(/\s*GMT[^\)]*(\([^)]*\))?/gi, "").trim();
}

function computeKommentarVerlaufDelta_(oldText, newText) {
  var oldT = String(oldText == null ? "" : oldText);
  var newT = String(newText == null ? "" : newText);
  if (oldT === newT) return null;

  if (newT.length > oldT.length && newT.indexOf(oldT) === 0) {
    return newT.substring(oldT.length);
  }
  if (oldT.length > newT.length && oldT.indexOf(newT) === 0) {
    return "-" + oldT.substring(newT.length);
  }
  if (oldT.length > 0 && newT.length > oldT.length && newT.slice(-oldT.length) === oldT) {
    return "+" + newT.substring(0, newT.length - oldT.length);
  }
  if (oldT.length > 0 && newT.length < oldT.length && oldT.slice(-newT.length) === newT) {
    return "-" + oldT.substring(0, oldT.length - newT.length);
  }

  var pi = 0;
  var maxP = Math.min(oldT.length, newT.length);
  while (pi < maxP && oldT.charAt(pi) === newT.charAt(pi)) pi++;

  var oRem = oldT.length - pi;
  var nRem = newT.length - pi;
  var si = 0;
  while (si < oRem && si < nRem && oldT.charAt(oldT.length - 1 - si) === newT.charAt(newT.length - 1 - si)) si++;

  var removed = oldT.substring(pi, oldT.length - si);
  var added = newT.substring(pi, newT.length - si);
  var parts = [];
  if (removed) parts.push("-" + removed);
  if (added) parts.push("+" + added);
  if (parts.length) return parts.join(" ");
  return newT;
}

function getKommentarVerlaufSpreadsheet_() {
  if (!KOMMENTAR_VERLAUF_SHEET_ID) return null;
  return SpreadsheetApp.openById(KOMMENTAR_VERLAUF_SHEET_ID);
}

function getKommentarVerlaufSheet_() {
  var ss = getKommentarVerlaufSpreadsheet_();
  if (!ss) return null;
  var sheet = ss.getSheetByName(KOMMENTAR_VERLAUF_TAB);
  if (!sheet) sheet = ss.getSheets()[0];
  return sheet || null;
}

function ensureKommentarVerlaufSheet_() {
  var sheet = getKommentarVerlaufSheet_();
  if (!sheet) return null;
  var header = String(sheet.getRange(1, 1).getValue() || "").trim();
  if (sheet.getLastRow() < 1 || header === "") {
    sheet.getRange(1, 1, 1, 5).setValues([["Zeitstempel", "Stock-ID", "Email", "Änderung", "Aktion"]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function logKommentarVerlauf_(stockId, oldText, newText, action) {
  try {
    stockId = normalizeStockId(stockId);
    action = String(action || "speichern");
    if (!stockId) return null;
    var delta = computeKommentarVerlaufDelta_(oldText, newText);
    if (delta === null || delta === "") return null;
    var sheet = ensureKommentarVerlaufSheet_();
    if (!sheet) return null;
    var tsStr = formatVerlaufTimestamp_(new Date());
    var email = getActiveUserEmail_();
    sheet.appendRow([tsStr, stockId, email, delta, action]);
    return { ts: tsStr, stockId: stockId, email: email, text: delta, action: action };
  } catch (e) {
    return null;
  }
}

function getKommentarVerlaufCachePayload() {
  try {
    if (!KOMMENTAR_VERLAUF_SHEET_ID) {
      return { success: true, version: Date.now(), rows: [], configured: false };
    }
    var sheet = getKommentarVerlaufSheet_();
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, version: Date.now(), rows: [], configured: true };
    }
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(2, 1, lastRow, 5).getValues();
    var rows = [];
    var i;
    for (i = data.length - 1; i >= 0; i--) {
      var r = data[i];
      var sid = normalizeStockId(r[1]);
      if (!sid) continue;
      rows.push([
        sid,
        verlaufTimestampStr_(r[0]),
        String(r[2] || ""),
        String(r[3] || ""),
        String(r[4] || "speichern")
      ]);
    }
    return { success: true, version: Date.now(), rows: rows, configured: true };
  } catch (err) {
    return { success: false, message: err.message, rows: [], configured: !!KOMMENTAR_VERLAUF_SHEET_ID };
  }
}

function getRefurbishmentCachePayload() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Refurbisment List");
    if (!sheet) return { success: false, message: "Reiter 'Refurbisment List' fehlt!" };

    var lastRow = Math.max(2, sheet.getLastRow());
    var rawRows = [];
    if (lastRow >= 2) {
      rawRows = sheet.getRange(2, 1, lastRow, 30).getValues();
    }
    var rows = [];
    var ri;
    for (ri = 0; ri < rawRows.length; ri++) {
      var r = rawRows[ri];
      rows.push([
        r[1], r[2], r[22], r[23], r[24], r[25], r[27], r[29], r[12]
      ]);
    }

    var gmailLookup = [];
    try {
      var lookupSs = SpreadsheetApp.openById(GMAIL_LOOKUP_SHEET_ID);
      var lookupSheet = lookupSs.getSheetByName(GMAIL_LOOKUP_TAB);
      if (lookupSheet && lookupSheet.getLastRow() >= 2) {
        var lookupLastRow = lookupSheet.getLastRow();
        gmailLookup = lookupSheet.getRange(2, 1, lookupLastRow, 2).getValues();
      }
    } catch (lookupErr) {}

    return {
      success: true,
      version: Date.now(),
      lastRow: lastRow,
      rows: rows,
      gmailLookup: gmailLookup
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

  function packzettelDateStr_(v) {
    if (v instanceof Date || Object.prototype.toString.call(v) === "[object Date]") {
      return Utilities.formatDate(new Date(v), "Europe/Berlin", "dd.MM.yyyy HH:mm");
    }
    return String(v || "");
  }

  function packzettelRowLight_(row, rowNumber) {
    return {
      row: rowNumber,
      messageDate: packzettelDateStr_(row[0]),
      source: String(row[1] || ""),
      kind: String(row[2] || ""),
      orderNumber: String(row[3] || "").trim(),
      referenceNumber: String(row[4] || "").trim(),
      stockId: String(row[5] || "").trim(),
      kennzeichen: String(row[6] || "").trim(),
      orderDate: String(row[7] || "").trim(),
      subject: String(row[8] || "").trim(),
      previewUrl: String(row[10] || "").trim(),
      downloadUrl: String(row[11] || "").trim()
    };
  }

  function pzDedupByOrder_(entries) {
    var seen = {};
    var out = [];
    for (var i = 0; i < entries.length; i++) {
      var on = String(entries[i].orderNumber || "").toUpperCase().replace(/\s+/g, "");
      if (on) {
        if (seen[on]) continue;
        seen[on] = true;
      }
      out.push(entries[i]);
    }
    return out;
  }

  function readPackzettelSheet_() {
    var ss = SpreadsheetApp.openById(GMAIL_LOOKUP_SHEET_ID);
    var sheet = ss.getSheetByName(PACKZETTEL_TAB);
    if (!sheet || sheet.getLastRow() < 2) return { sheet: null, values: [] };
    var lastRow = sheet.getLastRow();
    return { sheet: sheet, values: sheet.getRange(2, 1, lastRow - 1, 15).getValues() };
  }

  function packzettelRowMatchesQuery_(row, query) {
    if (!query) return true;
    var haystack = [
      row[3], row[4], row[5], row[6], row[8], row[1], row[7], row[13]
    ].join(" ").toLowerCase().replace(/\s+/g, "");
    return haystack.indexOf(query) !== -1;
  }

  function getPackzettelList(query) {
    try {
      var q = String(query || "").toLowerCase().replace(/\s+/g, "");
      var data = readPackzettelSheet_();
      var values = data.values;
      var matched = [];
      for (var i = 0; i < values.length; i++) {
        var row = values[i];
        if (!String(row[8] || "").trim() && !String(row[9] || "").trim() && !String(row[3] || "").trim()) continue;
        if (!packzettelRowMatchesQuery_(row, q)) continue;
        matched.push(packzettelRowLight_(row, i + 2));
      }
      matched.sort(function(a, b) {
        return String(b.messageDate).localeCompare(String(a.messageDate));
      });
      matched = pzDedupByOrder_(matched);
      return { success: true, entries: matched.slice(0, 200), total: matched.length };
    } catch (err) {
      return { success: false, message: err.message, entries: [] };
    }
  }

  function getPackzettelForStock(stockId) {
    try {
      var want = normalizeStockId(stockId);
      if (!want) return { success: true, entries: [], count: 0 };
      var data = readPackzettelSheet_();
      var values = data.values;
      var out = [];
      for (var i = 0; i < values.length; i++) {
        var row = values[i];
        var idNorm = normalizeStockId(row[5]);
        var rawHas = row[13] && normalizeStockId(row[13]).indexOf(want) !== -1;
        var subjHas = row[8] && normalizeStockId(row[8]).indexOf(want) !== -1;
        if (idNorm === want || rawHas || subjHas) out.push(packzettelRowLight_(row, i + 2));
      }
      out.sort(function(a, b) {
        return String(b.messageDate).localeCompare(String(a.messageDate));
      });
      out = pzDedupByOrder_(out);
      return { success: true, entries: out, count: out.length };
    } catch (err) {
      return { success: false, message: err.message, entries: [] };
    }
  }

  function getPackzettelForKeys(keys) {
    try {
      var normKeys = [];
      var seen = {};
      for (var k = 0; k < (keys || []).length; k++) {
        var nk = String(keys[k] || "").toUpperCase().replace(/\s+/g, "");
        if (nk.length >= 4 && !seen[nk]) { seen[nk] = true; normKeys.push(nk); }
      }
      if (!normKeys.length) return { success: true, entries: [], count: 0 };

      var data = readPackzettelSheet_();
      var values = data.values;
      var out = [];
      for (var i = 0; i < values.length; i++) {
        var row = values[i];
        var haystack = [row[3], row[4], row[5], row[8], row[13]]
          .join(" ").toUpperCase().replace(/\s+/g, "");
        var hit = false;
        for (var j = 0; j < normKeys.length; j++) {
          if (haystack.indexOf(normKeys[j]) !== -1) { hit = true; break; }
        }
        if (hit) out.push(packzettelRowLight_(row, i + 2));
      }
      out.sort(function(a, b) {
        return String(b.messageDate).localeCompare(String(a.messageDate));
      });
      out = pzDedupByOrder_(out);
      return { success: true, entries: out, count: out.length };
    } catch (err) {
      return { success: false, message: err.message, entries: [], count: 0 };
    }
  }

  function getPackzettelDoc(rowNumber) {
    try {
      var n = parseInt(rowNumber, 10);
      if (!n || n < 2) return { success: false, message: "Ungültige Zeile" };
      var ss = SpreadsheetApp.openById(GMAIL_LOOKUP_SHEET_ID);
      var sheet = ss.getSheetByName(PACKZETTEL_TAB);
      if (!sheet || n > sheet.getLastRow()) return { success: false, message: "Beleg nicht gefunden" };
      var row = sheet.getRange(n, 1, 1, 15).getValues()[0];
      return {
        success: true,
        kind: String(row[2] || ""),
        previewUrl: String(row[10] || "").trim(),
        downloadUrl: String(row[11] || "").trim(),
        bodyHtml: String(row[12] || ""),
        rawText: String(row[13] || "")
      };
    } catch (err) {
      return { success: false, message: err.message };
    }
  }

  function lagerKistenIsRegalHeader_(v) {
    return /^\s*regal\s+\d+\.\d+\s*$/i.test(String(v == null ? "" : v));
  }

  function lagerKistenCategoryLabel_(v) {
    var s = String(v == null ? "" : v).trim();
    if (!s) return "";
    var up = s.toUpperCase();
    var cats = ["B2A1", "COMPLETE", "TAGESLISTE", "KONTROLLIEREN", "NACHBESTELLT"];
    if (cats.indexOf(up) !== -1) return s;
    if (up.indexOf("RÜCKFRAGE") !== -1 || up.indexOf("RUCKFRAGE") !== -1) return s;
    if (up.indexOf("ÜBERSICHT") !== -1 || up.indexOf("UBERSICHT") !== -1) return s;
    return "";
  }

  function lagerKistenNormalizeRegal_(v) {
    var s = String(v == null ? "" : v).trim();
    var m = s.match(/(\d+)\.(\d+)/);
    if (m) return "Regal " + m[1] + "." + m[2];
    return s;
  }

  function lagerKistenLooksLikeStockId_(v) {
    var s = String(v == null ? "" : v).replace(/\s+/g, "").toUpperCase();
    return /^[A-Z]{2}\d{4,8}$/.test(s);
  }

  function getLagerKistenCachePayload() {
    try {
      var ss = SpreadsheetApp.openById(TAGESLISTE_SHEET_ID);
      var sheet = ss.getSheetByName("LAGER");
      if (!sheet) return { success: false, message: "Tab 'LAGER' nicht gefunden!", items: [], notes: [] };

      var range = sheet.getDataRange();
      var values = range.getValues();
      var cellNotes = range.getNotes();
      var backgrounds = range.getBackgrounds();
      var numRows = values.length;
      var numCols = numRows ? values[0].length : 0;

      var colorLabelMap = {};
      for (var lr = 0; lr < numRows; lr++) {
        for (var lc = 0; lc < numCols; lc++) {
          var label = lagerKistenCategoryLabel_(values[lr][lc]);
          if (!label) continue;
          var hex = String(backgrounds[lr][lc] || "").toLowerCase();
          if (hex && hex !== "#ffffff" && hex !== "#fff" && !colorLabelMap[hex]) {
            colorLabelMap[hex] = label;
          }
        }
      }

      var items = [];
      var looseNotes = [];
      for (var c = 0; c < numCols; c++) {
        var currentRegal = "";
        for (var r = 0; r < numRows; r++) {
          var raw = values[r][c];
          var note = (cellNotes[r] && cellNotes[r][c]) ? String(cellNotes[r][c]).trim() : "";
          if (lagerKistenIsRegalHeader_(raw)) {
            currentRegal = lagerKistenNormalizeRegal_(raw);
            continue;
          }
          if (lagerKistenCategoryLabel_(raw)) continue;
          var txt = String(raw == null ? "" : raw).trim();
          if (!txt) continue;
          if (lagerKistenLooksLikeStockId_(txt)) {
            var bg = String(backgrounds[r][c] || "").toLowerCase();
            items.push({
              stockId: txt.replace(/\s+/g, "").toUpperCase(),
              regal: currentRegal,
              kategorie: colorLabelMap[bg] || "",
              note: note
            });
          } else if (txt.length >= 4 || note) {
            looseNotes.push({
              regal: currentRegal,
              text: txt + (note ? " [Notiz: " + note + "]" : "")
            });
          }
        }
      }

      return { success: true, version: Date.now(), items: items, notes: looseNotes };
    } catch (err) {
      return { success: false, message: err.message, items: [], notes: [] };
    }
  }

  function saveKommentar(stockId, text, action, expectedKommentar, expectedRegal) {
    return withRefurbDocumentLock_(function() {
      stockId = normalizeStockId(stockId);
      action = String(action || "speichern");
      text = String(text == null ? "" : text);
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
      if (!sheet) return { success: false, message: "Reiter 'Refurbisment List' fehlt!" };
      var lastRow = Math.max(2, sheet.getLastRow());
      var data = sheet.getRange(1, 2, lastRow, 1).getValues();

      for (var i = 1; i < data.length; i++) {
        if (cellMatchesStockId(data[i][0], stockId)) {
          var row = i + 1;
          var oldText = String(sheet.getRange(row, 25).getValue() || "");
          var sheetRegal = String(sheet.getRange(row, 28).getValue() || "").trim();
          if (!baselinesMatch_(oldText, sheetRegal, expectedKommentar, expectedRegal)) {
            return buildConcurrencyConflict_(oldText, sheetRegal);
          }
          sheet.getRange(row, 25).setValue(text);
          SpreadsheetApp.flush();
          var check = sheet.getRange(row, 25).getValue();
          if (check != text) return { success: false, message: "Fehler beim Verifizieren!" };

          var verlaufEntry = logKommentarVerlauf_(stockId, oldText, text, action);
          var dateResult = applyTrackingDateIfEmpty(stockId);
          var msg = "Kommentar gespeichert!";
          if (dateResult.updated) msg += " Datum gesetzt!";
          if (!dateResult.success) msg += " " + dateResult.message;
          return {
            success: true,
            message: msg,
            verlaufEntry: verlaufEntry,
            savedKommentar: text,
            savedRegal: sheetRegal
          };
        }
      }
      return { success: false, message: "Stock-ID nicht gefunden!" };
    });
  }
  
  function einlagern(stockId, regal, expectedKommentar, expectedRegal) {
    return withRefurbDocumentLock_(function() {
      stockId = normalizeStockId(stockId);
      regal = String(regal || "").trim();
      if (!stockId) return { success: false, message: "Keine Stock-ID" };
      if (!regal) return { success: false, message: "Bitte Regal auswählen!" };
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
      if (!sheet) return { success: false, message: "Reiter 'Refurbisment List' fehlt!" };
      var lastRow = Math.max(2, sheet.getLastRow());
      var data = sheet.getRange(1, 2, lastRow, 1).getValues();

      for (var i = 1; i < data.length; i++) {
        if (cellMatchesStockId(data[i][0], stockId)) {
          var row = i + 1;
          var sheetKommentar = String(sheet.getRange(row, 25).getValue() || "");
          var sheetRegal = String(sheet.getRange(row, 28).getValue() || "").trim();
          if (!baselinesMatch_(sheetKommentar, sheetRegal, expectedKommentar, expectedRegal)) {
            return buildConcurrencyConflict_(sheetKommentar, sheetRegal);
          }
          sheet.getRange(row, 28).setValue(regal);
          SpreadsheetApp.flush();
          var check = sheet.getRange(row, 28).getValue();
          if (check != regal) return { success: false, message: "Fehler beim Verifizieren!" };
          return {
            success: true,
            message: "In " + regal + " eingelagert!",
            savedKommentar: sheetKommentar,
            savedRegal: regal
          };
        }
      }
      return { success: false, message: "Stock-ID nicht gefunden!" };
    });
  }

  function saveKommentarUndRegal(stockId, text, regal, expectedKommentar, expectedRegal) {
    return withRefurbDocumentLock_(function() {
      stockId = normalizeStockId(stockId);
      regal = String(regal || "").trim();
      text = String(text || "");
      if (!stockId) return { success: false, message: "Keine Stock-ID" };
      if (!text.trim()) return { success: false, message: "Bitte erst Kommentar eintragen!" };

      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
      if (!sheet) return { success: false, message: "Reiter 'Refurbisment List' fehlt!" };

      var lastRow = Math.max(2, sheet.getLastRow());
      var data = sheet.getRange(1, 2, lastRow, 1).getValues();

      for (var i = 1; i < data.length; i++) {
        if (cellMatchesStockId(data[i][0], stockId)) {
          var row = i + 1;
          var oldText = String(sheet.getRange(row, 25).getValue() || "");
          var sheetRegal = String(sheet.getRange(row, 28).getValue() || "").trim();
          if (!baselinesMatch_(oldText, sheetRegal, expectedKommentar, expectedRegal)) {
            return buildConcurrencyConflict_(oldText, sheetRegal);
          }
          sheet.getRange(row, 25).setValue(text);
          sheet.getRange(row, 25).setBackground("#ff0000");
          sheet.getRange(row, 26).setValue("Teilweise angeliefert");
          if (regal) {
            sheet.getRange(row, 28).setValue(regal);
          }
          SpreadsheetApp.flush();

          var commentCheck = sheet.getRange(row, 25).getValue();
          var statusCheck = String(sheet.getRange(row, 26).getValue() || "").trim();
          if (commentCheck != text || statusCheck !== "Teilweise angeliefert") {
            return { success: false, message: "Fehler beim Verifizieren!" };
          }
          var finalRegal = regal || sheetRegal;
          if (regal) {
            var regalCheck = String(sheet.getRange(row, 28).getValue() || "").trim();
            if (regalCheck !== regal) return { success: false, message: "Fehler beim Verifizieren!" };
            finalRegal = regalCheck;
          }

          var verlaufAction = regal ? "speichern+regal" : "speichern+status";
          var verlaufEntry = logKommentarVerlauf_(stockId, oldText, text, verlaufAction);
          var dateResult = applyTrackingDateIfEmpty(stockId);
          var msg = regal
            ? "Kommentar und Regal gespeichert! Status auf Teilweise angeliefert gesetzt."
            : "Kommentar gespeichert! Status auf Teilweise angeliefert gesetzt.";
          if (dateResult.updated) msg += " Datum gesetzt!";
          if (!dateResult.success) msg += " " + dateResult.message;
          return {
            success: true,
            message: msg,
            verlaufEntry: verlaufEntry,
            savedKommentar: text,
            savedRegal: finalRegal
          };
        }
      }

      return { success: false, message: "Stock-ID nicht gefunden!" };
    });
  }

  function getNachbestellungRegalOverviewEntries() {
    var entries = [];
    try {
      var ss = SpreadsheetApp.openById(NACHBESTELL_SHEET_ID);
      var sheet = ss.getSheetByName(NACHBESTELL_TAB);
      if (!sheet) return entries;

      var layout = findNachbestellungSheetLayout(sheet);
      var lagerortCol = layout.lagerortCol;
      var statusCol = layout.statusCol;
      var lastRow = Math.max(2, sheet.getLastRow());
      var lastCol = Math.max(1, Math.min(80, sheet.getLastColumn()));
      if (lagerortCol > lastCol) lagerortCol = Math.min(NACHBESTELL_REGAL_COL, lastCol);
      if (statusCol > lastCol) statusCol = Math.min(NACHBESTELL_STATUS_COL, lastCol);

      var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      var headerIdx = layout.headerRow - 1;
      var header = data[headerIdx];

      var stockCol = getColIndex(header, ["stockid", "stock"]);
      if (stockCol === -1) stockCol = 2;
      var typCol = getColIndex(header, ["art", "nachbestellung", "typ"]);
      if (typCol === -1) typCol = 3;
      var teilCol = getColIndex(header, ["ersatzteil", "teil", "benennung"]);
      if (teilCol === -1) teilCol = 5;
      var artikelCol = getColIndex(header, ["artikelnr", "artikelnummer", "teilenr"]);
      if (artikelCol === -1) artikelCol = 7;

      for (var i = layout.dataStartRow - 1; i < data.length; i++) {
        var stockId = String(data[i][stockCol - 1] || "").trim();
        if (!stockId) continue;

        var regalNorm = normalizeRegalKeyForCount(data[i][lagerortCol - 1]);
        if (!regalNorm || !/^Regal\s+\d+\.\d+$/i.test(regalNorm)) continue;

        entries.push({
          stockId: stockId,
          regal: regalNorm,
          status: String(data[i][statusCol - 1] || "").trim(),
          schaeden: typCol > 0 ? String(data[i][typCol - 1] || "") : "",
          kommBestellung: artikelCol > 0 ? String(data[i][artikelCol - 1] || "") : "",
          kommAnlieferung: teilCol > 0 ? String(data[i][teilCol - 1] || "") : "",
          regalReifen: "",
          source: "nachbestellung"
        });
      }
    } catch (err) {}
    return entries;
  }

  function getStockRegalOverview() {
    try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
      if (!sheet) return { success: false, message: "Reiter 'Refurbisment List' fehlt!", entries: [] };

      var lastRow = Math.max(2, sheet.getLastRow());
      var data = sheet.getRange(1, 1, lastRow, 30).getValues();
      var entries = [];
      var existingKeys = {};

      for (var i = 1; i < data.length; i++) {
        var stockId = String(data[i][1] || "").trim();
        var regal = String(data[i][27] || "").trim();
        if (!stockId) continue;
        if (!/^Regal\s+\d+\.\d+$/i.test(regal)) continue;
        entries.push({
          stockId: stockId,
          regal: regal,
          schaeden: String(data[i][22] || ""),
          kommBestellung: String(data[i][23] || ""),
          kommAnlieferung: String(data[i][24] || ""),
          regalReifen: String(data[i][29] || ""),
          status: String(data[i][25] || ""),
          source: "refurbishment"
        });
        existingKeys[normalizeStockId(stockId) + "|" + regal] = true;
      }

      var nbEntries = getNachbestellungRegalOverviewEntries();
      for (var n = 0; n < nbEntries.length; n++) {
        var nbKey = normalizeStockId(nbEntries[n].stockId) + "|" + nbEntries[n].regal;
        if (!existingKeys[nbKey]) {
          entries.push(nbEntries[n]);
          existingKeys[nbKey] = true;
        }
      }

      entries.sort(function(a, b) {
        var regalMatchA = String(a.regal).match(/^Regal\s+(\d+)\.(\d+)$/i);
        var regalMatchB = String(b.regal).match(/^Regal\s+(\d+)\.(\d+)$/i);

        if (regalMatchA && regalMatchB) {
          var aMain = parseInt(regalMatchA[1], 10);
          var bMain = parseInt(regalMatchB[1], 10);
          if (aMain !== bMain) return aMain - bMain;
          var aSub = parseInt(regalMatchA[2], 10);
          var bSub = parseInt(regalMatchB[2], 10);
          if (aSub !== bSub) return aSub - bSub;
        } else if (regalMatchA) {
          return -1;
        } else if (regalMatchB) {
          return 1;
        } else {
          var regalCompare = String(a.regal).localeCompare(String(b.regal), "de");
          if (regalCompare !== 0) return regalCompare;
        }

        return normalizeStockId(a.stockId).localeCompare(normalizeStockId(b.stockId), "de");
      });

      return { success: true, entries: entries };
    } catch (err) {
      return { success: false, message: err.message, entries: [] };
    }
  }
  
  function triggerCarolMission(stockId) {
    try {
      stockId = normalizeStockId(stockId);
      if (!stockId) return { success: false, message: "Keine Stock-ID", oldRegal: "LEER", carolUrl: "" };

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var log = [];
      var result = { success: false, oldRegal: "LEER", carolUrl: "" };

      var dateResult = applyTrackingDateIfEmpty(stockId);
      if (dateResult.updated) log.push("Datum gesetzt");
      if (dateResult.success && !dateResult.updated) log.push("Datum bereits vorhanden");
      if (!dateResult.success) log.push(dateResult.message);

      var sheetRefurb = ss.getSheetByName("Refurbisment List");
      if (!sheetRefurb) return { success: false, message: "Reiter Refurbisment List fehlt!", oldRegal: "LEER", carolUrl: "" };

      var lastRowRefurb = Math.max(2, sheetRefurb.getLastRow());
      var dataRefurb = sheetRefurb.getRange(1, 2, lastRowRefurb, 1).getValues();
      var foundRow = -1;
      for (var j = 0; j < dataRefurb.length; j++) {
        if (cellMatchesStockId(dataRefurb[j][0], stockId)) {
          foundRow = j + 1;
          break;
        }
      }
      if (foundRow === -1) {
        return { success: false, message: "Stock-ID " + stockId + " in Refurbisment List nicht gefunden!", oldRegal: "LEER", carolUrl: "" };
      }

      result.oldRegal = String(sheetRefurb.getRange(foundRow, 28).getValue() || "LEER");
      result.carolUrl = String(sheetRefurb.getRange(foundRow, 3).getValue() || "");

      sheetRefurb.getRange(foundRow, 25).setBackground("#00FF00");
      sheetRefurb.getRange(foundRow, 26).setValue("Herausgegeben");
      sheetRefurb.getRange(foundRow, 28).setValue("Tagesliste");
      SpreadsheetApp.flush();

      result.success = true;
      log.push("Zeile " + foundRow + " auf Tagesliste gesetzt");
      result.message = log.join(" | ");
      return result;
    } catch (err) {
      return { success: false, message: "Fehler: " + err.message, oldRegal: "LEER", carolUrl: "" };
    }
  }

function getNachbestellungen() {
    try {
      var ss = SpreadsheetApp.openById(NACHBESTELL_SHEET_ID);
      var sheet = ss.getSheetByName(NACHBESTELL_TAB);
      if (!sheet) return { success: false, message: "Tab '" + NACHBESTELL_TAB + "' nicht gefunden!", entries: [], lagerortOptions: [], statusOptions: [] };

      var layout = findNachbestellungSheetLayout(sheet);
      var lagerortCol = layout.lagerortCol;
      var statusCol = layout.statusCol;
      var lastColBound = Math.max(1, Math.min(80, sheet.getLastColumn()));
      if (lagerortCol > lastColBound) lagerortCol = Math.min(NACHBESTELL_REGAL_COL, lastColBound);
      if (statusCol > lastColBound) statusCol = Math.min(NACHBESTELL_STATUS_COL, lastColBound);
      var lagerortOptions = getNachbestellungLagerortAllowedList(sheet, lagerortCol);
      var statusOptions = getNachbestellungStatusAllowedList(sheet, statusCol);

      var lastRow = Math.max(2, sheet.getLastRow());
      var lastCol = Math.max(1, Math.min(80, sheet.getLastColumn()));
      var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

      var headerIdx = layout.headerRow - 1;
      var header = data[headerIdx];
      var cols = {};
      for (var c = 0; c < header.length; c++) {
        var txt = String(header[c] || "").toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
        if (!cols.date && (txt === "datum" || txt === "date")) cols.date = c;
        if (!cols.stock && (txt.indexOf("stock") !== -1)) cols.stock = c;
        if (!cols.url && (txt.indexOf("carol") !== -1 || txt.indexOf("url") !== -1 || txt.indexOf("link") !== -1)) cols.url = c;
        if (!cols.typ && (txt.indexOf("typ") !== -1 || txt.indexOf("art") !== -1 || txt.indexOf("bestellung") !== -1 || txt.indexOf("beschreibung") !== -1)) cols.typ = c;
        if (!cols.person && (txt.indexOf("team") !== -1 || txt.indexOf("name") !== -1 || txt.indexOf("person") !== -1 || txt.indexOf("mechaniker") !== -1)) cols.person = c;
        if (!cols.teil && (txt.indexOf("teil") !== -1 || txt.indexOf("article") !== -1 || txt.indexOf("ersatzteil") !== -1 || txt.indexOf("benennung") !== -1)) cols.teil = c;
        if (!cols.preis && (txt.indexOf("preis") !== -1 || txt.indexOf("kosten") !== -1 || txt.indexOf("price") !== -1)) cols.preis = c;
        if (!cols.artikel && (txt.indexOf("artikelnr") !== -1 || txt.indexOf("artikelnummer") !== -1 || txt.indexOf("article") !== -1 || txt.indexOf("teilenr") !== -1)) cols.artikel = c;
      }
      cols.status = statusCol - 1;

      if (cols.stock === undefined) {
        for (var bc = 0; bc < header.length; bc++) {
          var sample = String(data[headerIdx + 1] ? data[headerIdx + 1][bc] : "").trim();
          if (/^[A-Z]{2}\d{4,}/.test(sample)) { cols.stock = bc; break; }
        }
      }

      if (cols.stock === undefined) return { success: false, message: "Spalte 'Stock ID' nicht gefunden!", entries: [], lagerortOptions: lagerortOptions, statusOptions: statusOptions };

      var entryIdCol = getSheetEntryIdCol(sheet, NACHBESTELL_ENTRYID_COL);

      var entries = [];
      
      for (var i = headerIdx + 1; i < data.length; i++) {
        var stockId = String(data[i][cols.stock] || "").trim();
        
        if (!stockId) continue;

        var rawStatus = cols.status !== undefined ? String(data[i][cols.status] || "").trim() : "";
        if (nachbestellungIsClosedStatus(rawStatus)) continue;

        var dateVal = cols.date !== undefined ? data[i][cols.date] : "";
        var dateStr = "";
        if (dateVal instanceof Date) {
          dateStr = Utilities.formatDate(dateVal, "Europe/Berlin", "dd.MM.yyyy");
        } else if (Object.prototype.toString.call(dateVal) === '[object Date]') {
          dateStr = Utilities.formatDate(dateVal, "Europe/Berlin", "dd.MM.yyyy");
        } else {
          var rawDate = String(dateVal || "");
          var jsDateMatch = rawDate.match(/(\w+)\s(\w+)\s(\d{1,2})\s(\d{4})/);
          if (jsDateMatch) {
            var months = {Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
            var mm = months[jsDateMatch[2]] || 0;
            var dd = parseInt(jsDateMatch[3], 10);
            var yyyy = parseInt(jsDateMatch[4], 10);
            if (mm > 0) {
              dateStr = (dd < 10 ? "0" : "") + dd + "." + (mm < 10 ? "0" : "") + mm + "." + yyyy;
            } else {
              dateStr = rawDate;
            }
          } else {
            dateStr = rawDate;
          }
        }

        var typ = cols.typ !== undefined ? String(data[i][cols.typ] || "").trim() : "";

        var regalVal = "";
        if (lastCol >= lagerortCol) {
          var rawL = data[i][lagerortCol - 1];
          regalVal = nachbestellungRegalUiFromCell(rawL);
        }

        var entryIdVal = (entryIdCol > 0 && (lastCol >= entryIdCol)) ? String(data[i][entryIdCol - 1] || "").trim() : "";

        entries.push({
          row: i + 1,
          entryId: entryIdVal,
          date: dateStr,
          stockId: stockId,
          url: cols.url !== undefined ? String(data[i][cols.url] || "").trim() : "",
          typ: typ,
          person: cols.person !== undefined ? String(data[i][cols.person] || "").trim() : "",
          teil: cols.teil !== undefined ? String(data[i][cols.teil] || "").trim() : "",
          preis: cols.preis !== undefined ? String(data[i][cols.preis] || "").trim() : "",
          artikel: cols.artikel !== undefined ? String(data[i][cols.artikel] || "").trim() : "",
          status: cols.status !== undefined ? String(data[i][cols.status] || "").trim() : "",
          regal: regalVal
        });
      }

      entries.sort(function(a, b) {
        var ma = String(a.date).match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
        var mb = String(b.date).match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
        if (ma && mb) {
          var da = new Date(parseInt(ma[3], 10), parseInt(ma[2], 10) - 1, parseInt(ma[1], 10)).getTime();
          var db = new Date(parseInt(mb[3], 10), parseInt(mb[2], 10) - 1, parseInt(mb[1], 10)).getTime();
          if (da !== db) return db - da;
        }
        return b.row - a.row;
      });

      return { success: true, entries: entries, lagerortOptions: lagerortOptions, statusOptions: statusOptions };
    } catch (err) {
      return { success: false, message: err.message, entries: [], lagerortOptions: [], statusOptions: [] };
    }
  }

function getNachbestellungenForStock(stockId) {
    try {
      var wantStock = normalizeStockId(stockId);
      if (!wantStock) return { success: true, entries: [], openCount: 0, closedCount: 0 };

      var ss = SpreadsheetApp.openById(NACHBESTELL_SHEET_ID);
      var sheet = ss.getSheetByName(NACHBESTELL_TAB);
      if (!sheet) return { success: false, message: "Tab '" + NACHBESTELL_TAB + "' nicht gefunden!", entries: [] };

      var layout = findNachbestellungSheetLayout(sheet);
      var statusCol = layout.statusCol;
      var lagerortCol = layout.lagerortCol;
      var lastRow = Math.max(2, sheet.getLastRow());
      var lastCol = Math.max(1, Math.min(80, sheet.getLastColumn()));
      if (lagerortCol > lastCol) lagerortCol = Math.min(NACHBESTELL_REGAL_COL, lastCol);
      if (statusCol > lastCol) statusCol = Math.min(NACHBESTELL_STATUS_COL, lastCol);
      var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

      var headerIdx = layout.headerRow - 1;
      var header = data[headerIdx];
      var cols = {};
      for (var c = 0; c < header.length; c++) {
        var txt = String(header[c] || "").toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
        if (!cols.date && (txt === "datum" || txt === "date")) cols.date = c;
        if (!cols.stock && (txt.indexOf("stock") !== -1)) cols.stock = c;
        if (!cols.url && (txt.indexOf("carol") !== -1 || txt.indexOf("url") !== -1 || txt.indexOf("link") !== -1)) cols.url = c;
        if (!cols.typ && (txt.indexOf("typ") !== -1 || txt.indexOf("art") !== -1 || txt.indexOf("bestellung") !== -1 || txt.indexOf("beschreibung") !== -1)) cols.typ = c;
        if (!cols.teil && (txt.indexOf("teil") !== -1 || txt.indexOf("article") !== -1 || txt.indexOf("ersatzteil") !== -1 || txt.indexOf("benennung") !== -1)) cols.teil = c;
        if (!cols.artikel && (txt.indexOf("artikelnr") !== -1 || txt.indexOf("artikelnummer") !== -1 || txt.indexOf("article") !== -1 || txt.indexOf("teilenr") !== -1)) cols.artikel = c;
      }
      cols.status = statusCol - 1;

      if (cols.stock === undefined) {
        for (var bc = 0; bc < header.length; bc++) {
          var sample = String(data[headerIdx + 1] ? data[headerIdx + 1][bc] : "").trim();
          if (/^[A-Z]{2}\d{4,}/.test(sample)) { cols.stock = bc; break; }
        }
      }
      if (cols.stock === undefined) return { success: false, message: "Spalte 'Stock ID' nicht gefunden!", entries: [] };

      var entryIdCol = getSheetEntryIdCol(sheet, NACHBESTELL_ENTRYID_COL);

      var entries = [];
      var openCount = 0;
      var closedCount = 0;
      for (var i = headerIdx + 1; i < data.length; i++) {
        var rowStock = normalizeStockId(data[i][cols.stock]);
        if (!rowStock || rowStock !== wantStock) continue;

        var rawStatus = cols.status !== undefined ? String(data[i][cols.status] || "").trim() : "";
        var isClosed = nachbestellungIsClosedStatus(rawStatus);
        if (isClosed) closedCount++; else openCount++;

        var dateVal = cols.date !== undefined ? data[i][cols.date] : "";
        var dateStr = "";
        if (dateVal instanceof Date) {
          dateStr = Utilities.formatDate(dateVal, "Europe/Berlin", "dd.MM.yyyy");
        } else if (Object.prototype.toString.call(dateVal) === '[object Date]') {
          dateStr = Utilities.formatDate(dateVal, "Europe/Berlin", "dd.MM.yyyy");
        } else {
          dateStr = String(dateVal || "");
        }

        var regalVal = "";
        if (lastCol >= lagerortCol) regalVal = nachbestellungRegalUiFromCell(data[i][lagerortCol - 1]);

        entries.push({
          row: i + 1,
          entryId: (entryIdCol > 0 && lastCol >= entryIdCol) ? String(data[i][entryIdCol - 1] || "").trim() : "",
          date: dateStr,
          stockId: String(data[i][cols.stock] || "").trim(),
          typ: cols.typ !== undefined ? String(data[i][cols.typ] || "").trim() : "",
          teil: cols.teil !== undefined ? String(data[i][cols.teil] || "").trim() : "",
          artikel: cols.artikel !== undefined ? String(data[i][cols.artikel] || "").trim() : "",
          status: rawStatus,
          regal: regalVal,
          closed: isClosed
        });
      }

      return { success: true, stockId: stockId, entries: entries, openCount: openCount, closedCount: closedCount };
    } catch (err) {
      return { success: false, message: err.message, entries: [] };
    }
  }

function resolveNachbestellungTargetRow(sheet, nbLayout, sheetRow, expectedStockId, expectedEntryId) {
  var lastRow = Math.max(1, sheet.getLastRow());
  var stockCol = getNachbestellungStockIdCol(sheet, nbLayout);
  var eidCol = getSheetEntryIdCol(sheet, NACHBESTELL_ENTRYID_COL);
  var expStock = normalizeStockId(expectedStockId);
  var expEntry = String(expectedEntryId || "").trim();
  var dataStartIdx = Math.max(0, nbLayout.dataStartRow - 1);

  if (!expStock && !expEntry) {
    return { row: sheetRow, ok: true };
  }

  if (sheetRow >= 1 && sheetRow <= lastRow) {
    var rowStock = stockCol > 0 ? normalizeStockId(sheet.getRange(sheetRow, stockCol).getValue()) : "";
    var rowEntry = eidCol > 0 ? String(sheet.getRange(sheetRow, eidCol).getValue() || "").trim() : "";
    var entryOk = expEntry ? (rowEntry === expEntry) : true;
    var stockOk = expStock ? (rowStock === expStock) : true;
    if (entryOk && stockOk) {
      return { row: sheetRow, ok: true };
    }
  }

  if (expEntry && eidCol > 0) {
    var eids = sheet.getRange(1, eidCol, lastRow, 1).getValues();
    var eidMatches = [];
    for (var i = dataStartIdx; i < eids.length; i++) {
      if (String(eids[i][0] || "").trim() === expEntry) eidMatches.push(i + 1);
    }
    if (eidMatches.length === 1) return { row: eidMatches[0], ok: true, relocated: true };
    if (eidMatches.length > 1 && expStock && stockCol > 0) {
      for (var m = 0; m < eidMatches.length; m++) {
        var rs = normalizeStockId(sheet.getRange(eidMatches[m], stockCol).getValue());
        if (rs === expStock) return { row: eidMatches[m], ok: true, relocated: true };
      }
    }
  }

  if (expStock && stockCol > 0) {
    var stocks = sheet.getRange(1, stockCol, lastRow, 1).getValues();
    var stockMatches = [];
    for (var j = dataStartIdx; j < stocks.length; j++) {
      if (normalizeStockId(stocks[j][0]) === expStock) stockMatches.push(j + 1);
    }
    if (stockMatches.length === 1) return { row: stockMatches[0], ok: true, relocated: true };
    if (stockMatches.length > 1) {
      return { row: -1, ok: false, message: "Mehrere Zeilen mit Stock-ID " + expStock + " gefunden – bitte Liste aktualisieren (🔄) und erneut speichern." };
    }
  }

  return { row: -1, ok: false, message: "Zeile für '" + (expStock || expEntry) + "' nicht mehr gefunden – Liste hat sich geändert. Bitte aktualisieren (🔄)." };
}

function updateNachbestellung(sheetRow, fieldName, value, expectedStockId, expectedEntryId) {
    try {
      var autoLagerortWrite = null;
      var ss = SpreadsheetApp.openById(NACHBESTELL_SHEET_ID);
      var sheet = ss.getSheetByName(NACHBESTELL_TAB);
      if (!sheet) return { success: false, message: "Tab nicht gefunden!" };

      var nbLayout = findNachbestellungSheetLayout(sheet);

      var resolved = resolveNachbestellungTargetRow(sheet, nbLayout, sheetRow, expectedStockId, expectedEntryId);
      if (!resolved.ok) return { success: false, message: resolved.message };
      sheetRow = resolved.row;

      if (fieldName === "regal") {
        var lagerortCol = nbLayout.lagerortCol;
        var lcW = Math.max(1, sheet.getLastColumn());
        if (lagerortCol > lcW) lagerortCol = Math.min(NACHBESTELL_REGAL_COL, lcW);
        var allowedLv = getNachbestellungLagerortAllowedList(sheet, lagerortCol);
        var rawIn = String(value || "").trim();
        var regalWrite = rawIn === "" ? "" : nachbestellungLagerortToSheetValue(rawIn, allowedLv);
        if (rawIn !== "" && regalWrite === null) {
          return { success: false, message: "Lagerort nicht in der Sheet-Liste. Bitte passenden Eintrag wählen." };
        }
        var lr = sheet.getRange(sheetRow, lagerortCol);
        lr.setNumberFormat("@");
        lr.setValue(regalWrite);
        SpreadsheetApp.flush();
        var verifyCell = lr.getValue();
        if (!nachbestellungLagerortVerifyMatch(regalWrite, verifyCell)) return { success: false, message: "Lagerort konnte nicht verifiziert werden!" };
        return { success: true, message: "Lagerort gespeichert!" };
      }

      if (fieldName === "status") {
        var statusColW = nbLayout.statusCol;
        var scW = Math.max(1, sheet.getLastColumn());
        if (statusColW > scW) statusColW = Math.min(NACHBESTELL_STATUS_COL, scW);
        var allowedSt = getNachbestellungStatusAllowedList(sheet, statusColW);
        var rawSt = String(value || "").trim();
        var statusWrite = rawSt === "" ? "" : nachbestellungStatusToSheetValue(rawSt, allowedSt);
        if (rawSt !== "" && statusWrite === null) {
          return { success: false, message: "Status nicht in der Sheet-Liste. Bitte passenden Eintrag wählen." };
        }
        var stCell = sheet.getRange(sheetRow, statusColW);
        var oldStatusLc = String(stCell.getValue() || "").trim().toLowerCase();
        stCell.setValue(statusWrite);
        SpreadsheetApp.flush();
        var verifySt = String(stCell.getValue() || "").trim();
        if (rawSt !== "" && verifySt.toLowerCase() !== String(statusWrite || "").trim().toLowerCase()) {
          return { success: false, message: "Status konnte nicht verifiziert werden!" };
        }
        value = statusWrite;
        var newStatusLc = String(statusWrite || "").trim().toLowerCase();
        if (oldStatusLc.indexOf("teilweise angeliefert") !== -1 && newStatusLc.indexOf("komplett angeliefert") !== -1) {
          var lagerortColAuto = nbLayout.lagerortCol;
          var lcAuto = Math.max(1, sheet.getLastColumn());
          if (lagerortColAuto > lcAuto) lagerortColAuto = Math.min(NACHBESTELL_REGAL_COL, lcAuto);
          var currentLagerort = sheet.getRange(sheetRow, lagerortColAuto).getValue();
          if (normalizeRegalKeyForCount(currentLagerort)) {
            var allowedLvAuto = getNachbestellungLagerortAllowedList(sheet, lagerortColAuto);
            autoLagerortWrite = nachbestellungTageslisteAutomatikSheetValue(allowedLvAuto);
            if (autoLagerortWrite !== null) {
              var lrAuto = sheet.getRange(sheetRow, lagerortColAuto);
              lrAuto.setNumberFormat("@");
              lrAuto.setValue(autoLagerortWrite);
              SpreadsheetApp.flush();
            }
          }
        }
      } else {
        var lastCol = Math.max(1, Math.min(80, sheet.getLastColumn()));
        var headerData = sheet.getRange(1, 1, Math.min(10, sheet.getLastRow()), lastCol).getValues();
        var headerIdx = nbLayout.headerRow - 1;
        var header = headerData[headerIdx];
        var colMap = {};
        for (var c = 0; c < header.length; c++) {
          var txt = String(header[c] || "").toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
          if (txt.indexOf("teil") !== -1 || txt.indexOf("benennung") !== -1 || txt.indexOf("ersatzteil") !== -1) colMap["teil"] = c + 1;
          if (txt.indexOf("artikelnr") !== -1 || txt.indexOf("artikelnummer") !== -1 || txt.indexOf("teilenr") !== -1) colMap["artikel"] = c + 1;
        }
        var targetCol = colMap[fieldName];
        if (!targetCol) return { success: false, message: "Spalte '" + fieldName + "' nicht gefunden!" };
        sheet.getRange(sheetRow, targetCol).setValue(value);
        SpreadsheetApp.flush();
      }

      var lastCol = Math.max(1, Math.min(80, sheet.getLastColumn()));
      var extraMsgs = [];
      var printB64 = "";
      if (fieldName === "status") {
        var rowTyp = getNachbestellungRowTyp(sheet, sheetRow, nbLayout, lastCol);
        var rowStockId = normalizeStockId(sheet.getRange(sheetRow, getNachbestellungStockIdCol(sheet, nbLayout)).getValue());
        var valLcStatus = String(value || "").toLowerCase();
        var isExitTyp = rowTyp.toLowerCase().indexOf("exit") !== -1;
        var isKomplettAngeliefert = valLcStatus.indexOf("komplett angeliefert") !== -1;
        var wasAlreadyKomplett = oldStatusLc.indexOf("komplett angeliefert") !== -1;

        if (isExitTyp) {
          extraMsgs.push(syncNachbestellungStatusToInputExit(ss, sheet, sheetRow, value, nbLayout));
        }

        if (isExitTyp && isKomplettAngeliefert && !wasAlreadyKomplett) {
          var exitBeschreibung = "";
          var exitHdr = sheet.getRange(nbLayout.headerRow, 1, 1, lastCol).getValues()[0];
          var exitTeilCol = getColIndex(exitHdr, ["ersatzteil", "teil", "benennung"]);
          if (exitTeilCol !== -1) exitBeschreibung = String(sheet.getRange(sheetRow, exitTeilCol).getValue() || "").trim();
          if (!rowStockId) {
            var exitRowData = sheet.getRange(sheetRow, 1, 1, lastCol).getValues()[0];
            for (var er = 0; er < exitRowData.length; er++) {
              var exitCell = String(exitRowData[er] || "").trim();
              if (!rowStockId && /^[A-Z]{2}\d{4,}/i.test(exitCell)) rowStockId = normalizeStockId(exitCell);
            }
          }
          extraMsgs.push(sendInfoLagerExitChat_(rowStockId, exitBeschreibung));
        }

        if (valLcStatus.indexOf("angeliefert") !== -1) {
          var rowBeschreibung = "";
          var hdrData = sheet.getRange(nbLayout.headerRow, 1, 1, lastCol).getValues()[0];
          var teilCol = getColIndex(hdrData, ["ersatzteil", "teil", "benennung"]);
          if (teilCol !== -1) rowBeschreibung = String(sheet.getRange(sheetRow, teilCol).getValue() || "").trim();
          if (!rowStockId) {
            var rowData = sheet.getRange(sheetRow, 1, 1, lastCol).getValues()[0];
            for (var r = 0; r < rowData.length; r++) {
              var cellVal = String(rowData[r] || "").trim();
              if (!rowStockId && /^[A-Z]{2}\d{4,}/i.test(cellVal)) rowStockId = normalizeStockId(cellVal);
            }
          }

          var valStr = String(value || "");
          var valLc = valStr.toLowerCase();
          var shouldWerkstattAuftrag = nachbestellungTypShouldPrintWerkstattauftrag(rowTyp);
          if (shouldWerkstattAuftrag && rowStockId && (valStr.indexOf("Angeliefert/Bereit") !== -1 || valLc.indexOf("komplett angeliefert") !== -1)) {
            var waPrep = buildWerkstattauftragPrint_(rowStockId, rowBeschreibung, true);
            if (!waPrep.success) {
              waPrep = buildWerkstattauftragPrint_(rowStockId, rowBeschreibung, false);
            }
            if (waPrep.message) extraMsgs.push(waPrep.message);
            if (waPrep.success && waPrep.printB64) {
              printB64 = waPrep.printB64;
            }
          }
        }
      }

      var msg = "Gespeichert!";
      if (autoLagerortWrite) extraMsgs.push("Lagerort auf " + autoLagerortWrite + " gesetzt");
      for (var m = 0; m < extraMsgs.length; m++) {
        if (extraMsgs[m]) msg += " | " + extraMsgs[m];
      }
      var result = { success: true, message: msg };
      if (printB64) {
        result.printB64 = printB64;
      }
      if (autoLagerortWrite) {
        result.newRegal = nachbestellungRegalUiFromCell(autoLagerortWrite) || autoLagerortWrite;
      }
      return result;
    } catch (err) {
      return { success: false, message: err.message };
    }
  }

function getSheetEntryIdCol(sheet, fallbackCol) {
  if (!sheet) return fallbackCol || 0;
  var lastCol = Math.max(1, Math.min(80, sheet.getLastColumn()));
  for (var hr = 1; hr <= 5; hr++) {
    var header = sheet.getRange(hr, 1, 1, lastCol).getValues()[0];
    for (var c = 0; c < header.length; c++) {
      var txt = String(header[c] || "").toLowerCase().replace(/[^a-z0-9äöüß_]/g, "");
      if (txt === "entryid" || txt.indexOf("entryid") !== -1) return c + 1;
    }
  }
  return fallbackCol || 0;
}

function getNachbestellungStockIdCol(sheet, nbLayout) {
  var hdr = sheet.getRange(nbLayout.headerRow, 1, 1, Math.min(80, sheet.getLastColumn())).getValues()[0];
  var col = getColIndex(hdr, ["stockid", "stock"]);
  return col !== -1 ? col : 2;
}

function findInputExitRowByKeys(exitSheet, stockId, entryId) {
  var lastRow = Math.max(1, exitSheet.getLastRow());
  if (lastRow < 2) return { row: -1, message: "Input Exit: leer" };
  var eidCol = getSheetEntryIdCol(exitSheet, 27);
  var stockCol = 2;
  var targetE = String(entryId || "").trim();

  if (targetE && eidCol > 0) {
    var eids = exitSheet.getRange(2, eidCol, lastRow - 1, 1).getValues();
    var eidMatches = [];
    for (var i = 0; i < eids.length; i++) {
      if (String(eids[i][0] || "").trim() === targetE) eidMatches.push(i + 2);
    }
    if (eidMatches.length === 1) return { row: eidMatches[0] };
    if (eidMatches.length > 1) {
      if (stockId) {
        for (var k = 0; k < eidMatches.length; k++) {
          if (cellMatchesStockId(exitSheet.getRange(eidMatches[k], stockCol).getValue(), stockId)) {
            return { row: eidMatches[k] };
          }
        }
      }
      return { row: -1, message: "Input Exit: EntryID mehrfach vorhanden – bitte prüfen." };
    }
    return { row: -1, message: "Input Exit: EntryID nicht gefunden – nichts geändert (kein Fallback, um falsche Zeile zu vermeiden)." };
  }

  if (stockId) {
    var stocks = exitSheet.getRange(2, stockCol, lastRow - 1, 1).getValues();
    var stockMatches = [];
    for (var j = 0; j < stocks.length; j++) {
      if (cellMatchesStockId(stocks[j][0], stockId)) stockMatches.push(j + 2);
    }
    if (stockMatches.length === 1) return { row: stockMatches[0] };
    if (stockMatches.length > 1) {
      return { row: -1, message: "Input Exit: Stock-ID mehrfach vorhanden und keine EntryID – bitte EntryID pflegen." };
    }
  }
  return { row: -1, message: "Input Exit: Zeile nicht gefunden" };
}

function sendInfoLagerExitChat_(stockId, beschreibung) {
  try {
    stockId = String(stockId || "").trim();
    if (!stockId) return "Info Lager: keine Stock-ID";
    if (!INFO_LAGER_EXIT_WEBHOOK_URL) return "Info Lager: Webhook fehlt";
    var text = stockId + "\n-> " + String(beschreibung || "").trim() + "\nEXIT";
    var res = UrlFetchApp.fetch(INFO_LAGER_EXIT_WEBHOOK_URL, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ text: text }),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code >= 200 && code < 300) return "EXIT an Info Lager gesendet";
    return "Info Lager Fehler (" + code + ")";
  } catch (err) {
    return "Info Lager Fehler: " + err.message;
  }
}

function syncNachbestellungStatusToInputExit(ss, nbSheet, sheetRow, statusValue, nbLayout) {
  try {
    var exitSheet = ss.getSheetByName(INPUT_EXIT_TAB);
    if (!exitSheet) return "Input Exit: Tab nicht gefunden";
    var stockCol = getNachbestellungStockIdCol(nbSheet, nbLayout);
    var eidCol = getSheetEntryIdCol(nbSheet, NACHBESTELL_ENTRYID_COL);
    var stockId = normalizeStockId(nbSheet.getRange(sheetRow, stockCol).getValue());
    var entryId = eidCol > 0 ? String(nbSheet.getRange(sheetRow, eidCol).getValue() || "").trim() : "";
    if (!stockId && !entryId) return "Input Exit: Keine Stock-ID/EntryID";
    var exitMatch = findInputExitRowByKeys(exitSheet, stockId, entryId);
    var exitRow = exitMatch.row;
    if (exitRow === -1) return exitMatch.message || "Input Exit: Zeile nicht gefunden";
    var statusCell = exitSheet.getRange(exitRow, INPUT_EXIT_STATUS_COL);
    var allowedSt = getNachbestellungStatusAllowedList(exitSheet, INPUT_EXIT_STATUS_COL);
    var rawSt = String(statusValue || "").trim();
    var statusWrite = rawSt === "" ? "" : nachbestellungStatusToSheetValue(rawSt, allowedSt);
    if (rawSt !== "" && statusWrite === null) statusWrite = rawSt;
    var validation = statusCell.getDataValidation();
    statusCell.setValue(statusWrite);
    var stamp = (statusWrite !== "" && statusWrite != null) ? new Date() : "";
    exitSheet.getRange(exitRow, INPUT_EXIT_STATUS_DATE_COL).setValue(stamp);
    if (nbLayout.statusCol > 0) {
      nbSheet.getRange(sheetRow, nbLayout.statusCol + 1).setValue(stamp);
    }
    SpreadsheetApp.flush();
    var verifySt = String(statusCell.getValue() || "").trim();
    if (rawSt !== "" && verifySt.toLowerCase() !== String(statusWrite || "").trim().toLowerCase()) {
      statusCell.clearDataValidations();
      statusCell.setValue(statusWrite);
      if (validation) statusCell.setDataValidation(validation);
      SpreadsheetApp.flush();
      verifySt = String(statusCell.getValue() || "").trim();
    }
    return "Input Exit Z" + exitRow + " → " + verifySt;
  } catch (err) {
    return "Input Exit Fehler: " + err.message;
  }
}

function nachbestellungTypShouldPrintWerkstattauftrag(typ) {
  var t = String(typ || "").trim().toLowerCase();
  if (!t) return false;
  if (t.indexOf("exit") !== -1) return false;
  if (t.indexOf("erstbestellung") !== -1 && t.indexOf("falsch") !== -1) return true;
  if (t.indexOf("mechanik") !== -1 && t.indexOf("nachbestellung") !== -1) return true;
  if (t.indexOf("q-check") !== -1 || t.indexOf("qcheck") !== -1) return true;
  return false;
}

function getNachbestellungRowTyp(sheet, sheetRow, nbLayout, lastCol) {
  var hdrData = sheet.getRange(nbLayout.headerRow, 1, 1, lastCol).getValues()[0];
  var typCol = getColIndex(hdrData, ["typ"]);
  if (typCol === -1) typCol = getColIndex(hdrData, ["art", "bestellung"]);
  if (typCol === -1) return "";
  return String(sheet.getRange(sheetRow, typCol).getValue() || "").trim();
}

function getWmsUserCapabilities() {
    return {
      werkstattAuftragPrint: true
    };
  }

function formatWerkstattauftragDate_() {
  return Utilities.formatDate(new Date(), "Europe/Berlin", "dd.MM.yyyy");
}

function fillWerkstattauftragSheet_(zielSheet, stockId, beschreibung) {
  var dateStr = formatWerkstattauftragDate_();
  var desc = String(beschreibung || "").trim();
  zielSheet.getRange("B10").setValue(dateStr);
  zielSheet.getRange("D10").setValue(stockId);
  zielSheet.getRange("D18").setValue(desc);
  var markeModell = lookupHemauMarkeModell(stockId);
  if (markeModell) {
    zielSheet.getRange("B13:D13").setValue(markeModell);
  } else {
    zielSheet.getRange("B13:D13").clearContent();
  }
  return { dateStr: dateStr, markeModell: markeModell, desc: desc };
}

function verifyWerkstattauftragSheet_(zielSheet, stockId, beschreibung, strictDesc) {
  if (!zielSheet) return { ok: false, message: "Kein Werkstattauftrag-Blatt" };
  var filledStock = String(zielSheet.getRange("D10").getValue() || "").trim();
  var filledDesc = String(zielSheet.getRange("D18").getValue() || "").trim();
  if (!cellMatchesStockId(filledStock, stockId)) return { ok: false, message: "Stock-ID nicht verifiziert" };
  if (!filledDesc) return { ok: false, message: "Beschreibung leer" };
  if (strictDesc !== false && beschreibung && filledDesc !== String(beschreibung).trim()) {
    return { ok: false, message: "Beschreibung nicht verifiziert" };
  }
  return { ok: true };
}

function exportWerkstattauftragPdf_(spreadsheetId, zielSheet) {
  var gid = zielSheet ? zielSheet.getSheetId() : "";
  var exportUrl = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?exportFormat=pdf&format=pdf"
    + (gid !== "" && gid != null ? "&gid=" + gid : "")
    + "&portrait=true&size=A4&fitw=true&gridlines=false"
    + "&top_margin=0.3&bottom_margin=0.3&left_margin=0.3&right_margin=0.3"
    + "&sheetnames=false&printtitle=false&pagenumbers=false"
    + "&r1=0&c1=0&r2=46&c2=15";
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true,
    followRedirects: true
  });
  var code = response.getResponseCode();
  var blob = response.getBlob();
  var bytes = blob.getBytes();
  var isPdf = bytes && bytes.length >= 4
    && bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46;
  if (code !== 200 || !isPdf) {
    var hint = "HTTP " + code;
    try {
      var body = String(response.getContentText() || "").replace(/\s+/g, " ").slice(0, 120);
      if (body) hint += " · " + body;
    } catch (e) {}
    return { success: false, message: "PDF-Export fehlgeschlagen (" + hint + ")", printB64: "" };
  }
  if (!bytes || bytes.length < 100) {
    return { success: false, message: "PDF leer oder ungültig", printB64: "" };
  }
  return {
    success: true,
    message: "",
    printB64: Utilities.base64Encode(bytes)
  };
}

function trashWerkstattauftragTempFile_(fileId) {
  if (!fileId) return;
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (err) {
    try {
      SpreadsheetApp.openById(fileId).getSheets();
    } catch (e2) {}
  }
}

function buildWerkstattauftragPrint_(stockId, beschreibung, strictDesc) {
  var lock = LockService.getScriptLock();
  var gotLock = false;
  var tempId = "";
  try {
    gotLock = lock.tryLock(45000);
    if (!gotLock) {
      return { success: false, message: "Werkstattauftrag beschäftigt – bitte kurz erneut versuchen", printB64: "" };
    }
    stockId = normalizeStockId(stockId);
    if (!stockId) return { success: false, message: "Keine Stock-ID", printB64: "" };

    var sourceSs = SpreadsheetApp.openById(AUFTRAG_SHEET_ID);
    var template = sourceSs.getSheetByName(AUFTRAG_TAB);
    if (!template) return { success: false, message: "Tab '" + AUFTRAG_TAB + "' nicht gefunden", printB64: "" };

    var filled = null;
    var pdf = { success: false, message: "", printB64: "" };

    try {
      var tempSs = SpreadsheetApp.create("WMS_WA_" + stockId + "_" + Date.now());
      tempId = tempSs.getId();
      var copied = template.copyTo(tempSs);
      copied.setName(AUFTRAG_TAB);
      var sheets = tempSs.getSheets();
      for (var si = 0; si < sheets.length; si++) {
        if (sheets[si].getSheetId() !== copied.getSheetId()) {
          tempSs.deleteSheet(sheets[si]);
        }
      }
      filled = fillWerkstattauftragSheet_(copied, stockId, beschreibung);
      SpreadsheetApp.flush();
      Utilities.sleep(800);
      var verifyTemp = verifyWerkstattauftragSheet_(copied, stockId, beschreibung, strictDesc);
      if (!verifyTemp.ok) {
        return { success: false, message: "Auftrag: " + verifyTemp.message, printB64: "" };
      }
      pdf = exportWerkstattauftragPdf_(tempId, copied);
      if (!pdf.success) pdf = exportWerkstattauftragPdf_(tempId, null);
    } catch (tempErr) {
      pdf = { success: false, message: "Temp: " + tempErr.message, printB64: "" };
    }

    if (!pdf.success) {
      filled = fillWerkstattauftragSheet_(template, stockId, beschreibung);
      SpreadsheetApp.flush();
      Utilities.sleep(500);
      var verifyMain = verifyWerkstattauftragSheet_(template, stockId, beschreibung, strictDesc);
      if (!verifyMain.ok) {
        return { success: false, message: "Auftrag: " + verifyMain.message, printB64: "" };
      }
      pdf = exportWerkstattauftragPdf_(AUFTRAG_SHEET_ID, template);
      if (!pdf.success) {
        return { success: false, message: pdf.message || "PDF-Export fehlgeschlagen", printB64: "" };
      }
    }

    var msg = "Werkstattauftrag befüllt (" + stockId + ")";
    if (filled && filled.dateStr) msg += " · " + filled.dateStr;
    if (filled && filled.markeModell) msg += " + " + filled.markeModell;
    return { success: true, message: msg, printB64: pdf.printB64 };
  } catch (err) {
    return { success: false, message: "Auftrag Fehler: " + err.message, printB64: "" };
  } finally {
    trashWerkstattauftragTempFile_(tempId);
    if (gotLock) {
      try { lock.releaseLock(); } catch (e3) {}
    }
  }
}

function deleteWerkstattauftragTempSheet_(ss, tempSheet) {
  if (!ss || !tempSheet) return;
  try {
    ss.deleteSheet(tempSheet);
  } catch (err) {
    try { tempSheet.setName("TMP_WA_DEL_" + Date.now()); tempSheet.hideSheet(); } catch (e2) {}
  }
}

function verifyWerkstattauftragFill(stockId, beschreibung, strictDesc) {
    var zielSs = SpreadsheetApp.openById(AUFTRAG_SHEET_ID);
    var zielSheet = zielSs.getSheetByName(AUFTRAG_TAB);
    if (!zielSheet) return { ok: false, message: "Tab '" + AUFTRAG_TAB + "' nicht gefunden" };
    var check = verifyWerkstattauftragSheet_(zielSheet, stockId, beschreibung, strictDesc);
    if (!check.ok) return check;
    return { ok: true, sheet: zielSheet };
  }

function lookupHemauMarkeModell(stockId) {
    try {
      stockId = normalizeStockId(stockId);
      if (!stockId) return "";
      var ss = SpreadsheetApp.openById(HEMAU_SHEET_ID);
      var sheet = ss.getSheetByName(HEMAU_DAILY_PLANNING_TAB);
      if (!sheet) return "";
      var search = findRowFast(sheet, ["stockid", "stock"], stockId);
      if (search.row === -1) return "";
      var lastCol = Math.max(1, Math.min(80, sheet.getLastColumn()));
      var headerData = sheet.getRange(search.headerIdx + 1, 1, 1, lastCol).getValues()[0];
      var markeCol = getColIndex(headerData, ["markemodell", "marke", "modell"]);
      if (markeCol === -1) markeCol = 9;
      return String(sheet.getRange(search.row, markeCol).getValue() || "").trim();
    } catch (err) {
      return "";
    }
  }

function autoFillWerkstattauftrag(stockId, beschreibung) {
    var res = buildWerkstattauftragPrint_(stockId, beschreibung, true);
    if (res.success) return res.message;
    return res.message || "";
  }

function prepareWerkstattauftragPrint(stockId, beschreibung, strictDesc) {
    return buildWerkstattauftragPrint_(stockId, beschreibung, strictDesc);
  }

function requestWerkstattauftragPrint(stockId, beschreibung, typ) {
    try {
      if (!nachbestellungTypShouldPrintWerkstattauftrag(typ)) {
        return { success: false, message: "Für diesen Typ kein Werkstattauftrag" };
      }
      var prep = buildWerkstattauftragPrint_(stockId, beschreibung, true);
      if (prep.success) return prep;
      var relaxed = buildWerkstattauftragPrint_(stockId, beschreibung, false);
      if (relaxed.success) return relaxed;
      return { success: false, message: prep.message || relaxed.message || "Druck nicht möglich" };
    } catch (err) {
      return { success: false, message: err.message };
    }
  }

function getHeutigeReifenausgabe() {
  try {
    var sheet = SpreadsheetApp.openById(TAGESLISTE_SHEET_ID).getSheetByName(TAGESLISTE_TAB);
    if (!sheet) return { success: false, message: "Tab '" + TAGESLISTE_TAB + "' nicht gefunden!", entries: [] };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, entries: [] };

    var headerIdx = -1;
    for (var h = 0; h < Math.min(10, data.length); h++) {
      var rowLc = data[h].map(function(c) { return String(c || "").toLowerCase(); });
      if (rowLc.some(function(c) { return c.indexOf("stock") !== -1; })) { headerIdx = h; break; }
    }
    if (headerIdx === -1) headerIdx = 0;

    var header = data[headerIdx];
    var cols = {};
    for (var c = 0; c < header.length; c++) {
      var txt = String(header[c] || "").toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
      if (!cols.date && (txt === "datum" || txt === "date")) cols.date = c;
      if (!cols.stock && (txt.indexOf("stock") !== -1)) cols.stock = c;
      if (!cols.reifen && (txt.indexOf("reifen") !== -1)) cols.reifen = c;
    }
    if (cols.stock === undefined) return { success: false, message: "Spalte 'Stock ID' nicht gefunden!", entries: [] };
    if (cols.reifen === undefined) return { success: false, message: "Spalte 'Reifen' nicht gefunden!", entries: [] };

    var today = new Date(); today.setHours(0,0,0,0);
    var tomorrow = new Date(today); tomorrow.setDate(tomorrow.getDate() + 1);
    var dayAfter = new Date(tomorrow); dayAfter.setDate(dayAfter.getDate() + 1);

    var entries = [];
    for (var i = headerIdx + 1; i < data.length; i++) {
      var stockId = String(data[i][cols.stock] || "").trim();
      if (!stockId) continue;

      var reifenVal = String(data[i][cols.reifen] || "").trim().toLowerCase();
      var isReifen   = (reifenVal === "2 reifen"  || reifenVal === "4 reifen");
      var isGestellt = (reifenVal === "2 gestellt" || reifenVal === "4 gestellt");
      if (!isReifen && !isGestellt) continue;

      var dateVal = cols.date !== undefined ? data[i][cols.date] : null;
      var dateObj = null, dateStr = "";
      if (dateVal instanceof Date || Object.prototype.toString.call(dateVal) === '[object Date]') {
        dateObj = new Date(dateVal); dateObj.setHours(0,0,0,0);
        dateStr = Utilities.formatDate(dateVal, "Europe/Berlin", "dd.MM.yyyy");
      } else {
        var raw = String(dateVal || "");
        var m = raw.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
        if (m) { dateObj = new Date(+m[3], +m[2]-1, +m[1]); dateObj.setHours(0,0,0,0); dateStr = raw; }
      }
      if (!dateObj || dateObj < today || dateObj >= dayAfter) continue;

      entries.push({
        row: i + 1,
        date: dateStr,
        stockId: stockId,
        reifen: String(data[i][cols.reifen] || "").trim(),
        amount: reifenVal.indexOf("4") !== -1 ? 4 : 2,
        gestellt: isGestellt,
        reifenCol: cols.reifen + 1
      });
    }

    entries.sort(function(a, b) {
      if (a.gestellt !== b.gestellt) return a.gestellt ? 1 : -1;
      var ma = String(a.date).match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
      var mb = String(b.date).match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
      if (ma && mb) {
        var da = new Date(parseInt(ma[3], 10), parseInt(ma[2], 10) - 1, parseInt(ma[1], 10)).getTime();
        var db = new Date(parseInt(mb[3], 10), parseInt(mb[2], 10) - 1, parseInt(mb[1], 10)).getTime();
        if (da !== db) return da - db;
      }
      return a.row - b.row;
    });

    return { success: true, entries: entries };
  } catch (err) {
    return { success: false, message: err.message, entries: [] };
  }
}

function updateReifenGestellt(sheetRow, reifenCol, newValue) {
  try {
    var ss = SpreadsheetApp.openById(TAGESLISTE_SHEET_ID);
    var sheet = ss.getSheetByName(TAGESLISTE_TAB);
    if (!sheet) return { success: false, message: "Tab nicht gefunden!" };
    sheet.getRange(sheetRow, reifenCol).setValue(newValue);
    SpreadsheetApp.flush();
    var verify = String(sheet.getRange(sheetRow, reifenCol).getValue() || "").trim().toLowerCase();
    if (verify === newValue.toLowerCase()) {
      return { success: true, message: "Gespeichert!" };
    }
    return { success: false, message: "Verifikation fehlgeschlagen: '" + verify + "'" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function stelleAlleReifen(entries) {
  try {
    var ss = SpreadsheetApp.openById(TAGESLISTE_SHEET_ID);
    var sheet = ss.getSheetByName(TAGESLISTE_TAB);
    if (!sheet) return { success: false, message: "Tab nicht gefunden!" };
    var updated = 0;
    var failed = 0;
    for (var i = 0; i < entries.length; i++) {
      var e = entries[i];
      if (e.gestellt) continue;
      var newVal = e.amount === 4 ? "4 gestellt" : "2 gestellt";
      sheet.getRange(e.row, e.reifenCol).setValue(newVal);
      updated++;
    }
    SpreadsheetApp.flush();
    for (var j = 0; j < entries.length; j++) {
      var en = entries[j];
      if (en.gestellt) continue;
      var expected = en.amount === 4 ? "4 gestellt" : "2 gestellt";
      var actual = String(sheet.getRange(en.row, en.reifenCol).getValue() || "").trim().toLowerCase();
      if (actual !== expected.toLowerCase()) failed++;
    }
    if (failed > 0) return { success: false, message: updated + " aktualisiert, " + failed + " fehlgeschlagen!" };
    return { success: true, message: "Alle " + updated + " Reifen auf gestellt gesetzt!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}