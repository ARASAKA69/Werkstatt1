const TRACKING_SHEET_URL = "https://docs.google.com/spreadsheets/d/1PuCLw8UmDjB_pBo_jCZ9rmSD3GJQESHzPoBVu_--MRo/edit?gid=1453769469#gid=1453769469";
const REIFEN_SHEET_ID = "1NTWkl4r40VUb8hM3Zk5BYWofdxn0FgtZh4DJpOufSd8";
const NACHBESTELL_SHEET_ID = "1PuCLw8UmDjB_pBo_jCZ9rmSD3GJQESHzPoBVu_--MRo";
const NACHBESTELL_TAB = "Nachbestellungen";

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

function getReifenSheet() {
    return SpreadsheetApp.openById(REIFEN_SHEET_ID);
  }

function getReifenTabOptions() {
    try {
      var ss = getReifenSheet();
      var sheets = ss.getSheets().map(function(sheet) {
        return sheet.getName();
      });
      sheets.sort(function(a, b) {
        var matchA = String(a).match(/(\d{2})\.(\d{2})\.(\d{4})$/);
        var matchB = String(b).match(/(\d{2})\.(\d{2})\.(\d{4})$/);
        if (matchA && matchB) {
          var dateA = new Date(parseInt(matchA[3], 10), parseInt(matchA[2], 10) - 1, parseInt(matchA[1], 10)).getTime();
          var dateB = new Date(parseInt(matchB[3], 10), parseInt(matchB[2], 10) - 1, parseInt(matchB[1], 10)).getTime();
          return dateB - dateA;
        }
        if (matchA) return -1;
        if (matchB) return 1;
        return String(a).localeCompare(String(b), "de");
      });
      return { success: true, tabs: sheets };
    } catch (err) {
      return { success: false, message: err.message, tabs: [] };
    }
  }

function getAvailableReifenStockIds(tabName) {
    try {
      var sheet = getReifenSheet().getSheetByName(String(tabName || "").trim());
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
      var stockData = sheet.getRange(startRow, search.stockCol, numRows, 1).getValues();
      var statusData = angeliefertCol !== -1 ? sheet.getRange(startRow, angeliefertCol, numRows, 1).getValues() : [];
      var ids = [];
      for (var i = 0; i < stockData.length; i++) {
        var val = normalizeStockId(stockData[i][0]);
        if (val) {
          ids.push({
            id: val,
            status: angeliefertCol !== -1 ? String(statusData[i][0] || "").trim().toLowerCase() : ""
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

      var sheet = getReifenSheet().getSheetByName(String(tabName || "").trim());
      if (!sheet) return { found: false, message: "Bitte ein gültiges Tabellenblatt auswählen." };

      var search = findRowFast(sheet, ["stockid", "stock"], stockId);
      if (search.headerIdx === -1) return { found: false, message: "Kopfzeile 'Stock ID' in Reifenliste nicht gefunden!" };
      if (search.row === -1) return { found: false, message: "Stock-ID '" + stockId + "' in '" + sheet.getName() + "' nicht gefunden!" };
      var headerRow = sheet.getRange(search.headerIdx + 1, 1, 1, Math.max(1, Math.min(80, sheet.getLastColumn()))).getValues()[0];
      var angeliefertCol = getColIndex(headerRow, ["angeliefert"]);
      if (angeliefertCol !== -1) {
        var currentStatus = String(sheet.getRange(search.row, angeliefertCol).getValue() || "").trim().toLowerCase();
        if (currentStatus === "ja" || currentStatus === "nein") {
          return { found: false, message: "Stock-ID '" + stockId + "' wurde in '" + sheet.getName() + "' bereits verbucht!" };
        }
      }
      return { found: true, message: "Stock-ID gefunden! Bitte Status auswählen:" };
    } catch (err) {
      return { found: false, message: "Systemfehler: " + err.message };
    }
  }

function processReifenStock(tabName, stockId, isDelivered) {
    try {
      stockId = normalizeStockId(stockId);
      var sheetSeng = getReifenSheet().getSheetByName(String(tabName || "").trim());
      if (!sheetSeng) return { success: false, message: "Bitte ein gültiges Tabellenblatt auswählen." };

      var sheetHemau = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
      if (!sheetHemau) return { success: false, message: "Reiter 'Refurbisment List' fehlt!" };

      var tireInfo = "UNBEKANNT _X";
      var mengeValNum = 1;
      var search = findRowFast(sheetSeng, ["stockid", "stock"], stockId);
      if (search.headerIdx === -1) return { success: false, message: "Kopfzeile 'Stock ID' in Reifenliste nicht gefunden!" };
      if (search.row === -1) return { success: false, message: "Stock-ID '" + stockId + "' in '" + sheetSeng.getName() + "' nicht gefunden!" };

      var headerRow = sheetSeng.getRange(search.headerIdx + 1, 1, 1, Math.max(1, Math.min(80, sheetSeng.getLastColumn()))).getValues()[0];
      var angeliefertCol = getColIndex(headerRow, ["angeliefert"]);
      var mengeCol = getColIndex(headerRow, ["menge", "anzahl"]);
      var groesseCol = getColIndex(headerRow, ["größe", "groesse"]);
      var lastIndexCol = getColIndex(headerRow, ["lastindex", "last"]);
      var gwIndexCol = getColIndex(headerRow, ["gwindex", "gw"]);
      if (angeliefertCol !== -1) {
        var existingStatus = String(sheetSeng.getRange(search.row, angeliefertCol).getValue() || "").trim().toLowerCase();
        if (existingStatus === "ja" || existingStatus === "nein") {
          return { success: false, message: "Stock-ID '" + stockId + "' wurde in '" + sheetSeng.getName() + "' bereits verbucht!" };
        }
      }

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
      if (hemauRow !== -1) {
        var oldLocation = String(sheetHemau.getRange(hemauRow, 28).getValue() || "").trim();
        if (oldLocation !== "") locationText = "Kiste steht in Regal " + oldLocation;

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
        menge: mengeValNum
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
  
  function openWMS() {
    var html = HtmlService.createHtmlOutputFromFile('WMS_HUD')
      .setWidth(1900)
      .setHeight(1100);
    SpreadsheetApp.getUi().showModelessDialog(html, 'Warehouse Management System');
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
  
  function fetchWmsData(stockId) {
    try {
      stockId = normalizeStockId(stockId);
      if (!stockId) return { success: false, message: "Keine Stock-ID" };
  
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("Refurbisment List");
      if (!sheet) return { success: false, message: "Reiter 'Refurbisment List' fehlt!" };
  
      var lastRow = Math.max(2, sheet.getLastRow());
      var stockColData = sheet.getRange(1, 2, lastRow, 1).getValues();
      var regalColData = sheet.getRange(1, 28, lastRow, 1).getValues();
      var result = { success: false };
      var hitRow = -1;
  
      var counts = {};
      for (var i = 1; i <= 9; i++) {
        for (var j = 1; j <= 8; j++) {
          counts["Regal " + i + "." + j] = 0;
        }
      }
  
      for (var r = 1; r < regalColData.length; r++) {
        var regalVal = String(regalColData[r][0] || "").trim();
        if (counts.hasOwnProperty(regalVal)) counts[regalVal]++;
      }

      for (var s = 1; s < stockColData.length; s++) {
        if (cellMatchesStockId(stockColData[s][0], stockId)) {
          hitRow = s + 1;
          break;
        }
      }
  
      var availableShelves = [];
      for (var shelf in counts) {
        if (counts[shelf] < 5) {
          availableShelves.push({ name: shelf, count: counts[shelf] });
        }
      }
      result.freeShelves = availableShelves;
  
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
        result.currentShelfCount = counts[result.regal] || 0;
        result.currentShelfCapacity = 5;
      } else {
        result.message = "Stock-ID in Refurbisment List nicht gefunden!";
      }
  
      return result;
    } catch (err) {
      return { success: false, message: err.message };
    }
  }
  
  function saveKommentar(stockId, text) {
    try {
      stockId = normalizeStockId(stockId);
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
      var lastRow = Math.max(2, sheet.getLastRow());
      var data = sheet.getRange(1, 2, lastRow, 1).getValues();

      for (var i = 1; i < data.length; i++) {
        if (cellMatchesStockId(data[i][0], stockId)) {
          var row = i + 1;
          sheet.getRange(row, 25).setValue(text);
          SpreadsheetApp.flush();
          var check = sheet.getRange(row, 25).getValue();
          if (check != text) return { success: false, message: "Fehler beim Verifizieren!" };

          var dateResult = applyTrackingDateIfEmpty(stockId);
          var msg = "Kommentar gespeichert!";
          if (dateResult.updated) msg += " Datum gesetzt!";
          if (!dateResult.success) msg += " " + dateResult.message;
          return { success: true, message: msg };
        }
      }
      return { success: false, message: "Stock-ID nicht gefunden!" };
    } catch (err) {
      return { success: false, message: "Fehler: " + err.message };
    }
  }
  
  function einlagern(stockId, regal) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
    var lastRow = Math.max(2, sheet.getLastRow());
    var data = sheet.getRange(1, 2, lastRow, 1).getValues();

    for (var i = 1; i < data.length; i++) {
      if (cellMatchesStockId(data[i][0], stockId)) {
        var row = i + 1;
        sheet.getRange(row, 28).setValue(regal);
        SpreadsheetApp.flush();
        var check = sheet.getRange(row, 28).getValue();
        return (check == regal) ? { success: true, message: "In " + regal + " eingelagert!" } : { success: false, message: "Fehler beim Verifizieren!" };
      }
    }
    return { success: false, message: "Stock-ID nicht gefunden!" };
  }

  function saveKommentarUndRegal(stockId, text, regal) {
    try {
      stockId = normalizeStockId(stockId);
      regal = String(regal || "").trim();
      text = String(text || "");
      if (!stockId) return { success: false, message: "Keine Stock-ID" };
      if (!text.trim()) return { success: false, message: "Bitte erst Kommentar eintragen!" };
      if (!regal) return { success: false, message: "Bitte Regal auswählen!" };

      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
      if (!sheet) return { success: false, message: "Reiter 'Refurbisment List' fehlt!" };

      var lastRow = Math.max(2, sheet.getLastRow());
      var data = sheet.getRange(1, 2, lastRow, 1).getValues();

      for (var i = 1; i < data.length; i++) {
        if (cellMatchesStockId(data[i][0], stockId)) {
          var row = i + 1;
          sheet.getRange(row, 25).setValue(text);
          sheet.getRange(row, 25).setBackground("#ff0000");
          sheet.getRange(row, 26).setValue("Teilweise angeliefert");
          sheet.getRange(row, 28).setValue(regal);
          SpreadsheetApp.flush();

          var commentCheck = sheet.getRange(row, 25).getValue();
          var statusCheck = String(sheet.getRange(row, 26).getValue() || "").trim();
          var regalCheck = String(sheet.getRange(row, 28).getValue() || "").trim();
          if (commentCheck != text || statusCheck !== "Teilweise angeliefert" || regalCheck !== regal) {
            return { success: false, message: "Fehler beim Verifizieren!" };
          }

          var dateResult = applyTrackingDateIfEmpty(stockId);
          var msg = "Kommentar und Regal gespeichert! Status auf Teilweise angeliefert gesetzt.";
          if (dateResult.updated) msg += " Datum gesetzt!";
          if (!dateResult.success) msg += " " + dateResult.message;
          return { success: true, message: msg };
        }
      }

      return { success: false, message: "Stock-ID nicht gefunden!" };
    } catch (err) {
      return { success: false, message: "Fehler: " + err.message };
    }
  }

  function getStockRegalOverview() {
    try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
      if (!sheet) return { success: false, message: "Reiter 'Refurbisment List' fehlt!", entries: [] };

      var lastRow = Math.max(2, sheet.getLastRow());
      var data = sheet.getRange(1, 1, lastRow, 30).getValues();
      var entries = [];

      for (var i = 1; i < data.length; i++) {
        var stockId = String(data[i][1] || "").trim();
        var regal = String(data[i][27] || "").trim();
        if (!stockId) continue;
        if (!/^Regal\s+\d+\.\d+$/i.test(regal)) continue;
        entries.push({
          stockId: stockId,
          regal: regal,
          kommBestellung: String(data[i][23] || ""),
          kommAnlieferung: String(data[i][24] || ""),
          regalReifen: String(data[i][29] || ""),
          status: String(data[i][25] || "")
        });
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
      if (!sheet) return { success: false, message: "Tab '" + NACHBESTELL_TAB + "' nicht gefunden!", entries: [] };

      var lastRow = Math.max(2, sheet.getLastRow());
      var lastCol = Math.max(1, Math.min(15, sheet.getLastColumn()));
      var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

      var headerIdx = -1;
      for (var h = 0; h < Math.min(10, data.length); h++) {
        var row = data[h].map(function(c) { return String(c || "").toLowerCase(); });
        if (row.some(function(c) { return c.indexOf("stock") !== -1; })) {
          headerIdx = h;
          break;
        }
      }
      if (headerIdx === -1) headerIdx = 0;

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
        if (!cols.status && (txt.indexOf("status") !== -1 || txt.indexOf("bestellt") !== -1 || txt.indexOf("angeliefert") !== -1)) cols.status = c;
      }

      if (cols.stock === undefined) {
        for (var bc = 0; bc < header.length; bc++) {
          var sample = String(data[headerIdx + 1] ? data[headerIdx + 1][bc] : "").trim();
          if (/^[A-Z]{2}\d{4,}/.test(sample)) { cols.stock = bc; break; }
        }
      }

      if (cols.stock === undefined) return { success: false, message: "Spalte 'Stock ID' nicht gefunden!", entries: [] };

      var entries = [];
      for (var i = headerIdx + 1; i < data.length; i++) {
        var stockId = String(data[i][cols.stock] || "").trim();
        if (!stockId) continue;

        var statusVal = cols.status !== undefined ? String(data[i][cols.status] || "").trim().toLowerCase() : "";
        if (statusVal === "angeliefert") continue;

        var dateVal = cols.date !== undefined ? data[i][cols.date] : "";
        var dateStr = "";
        if (dateVal instanceof Date) {
          dateStr = Utilities.formatDate(dateVal, "Europe/Berlin", "dd.MM.yyyy");
        } else {
          dateStr = String(dateVal || "");
        }

        var typ = cols.typ !== undefined ? String(data[i][cols.typ] || "").trim() : "";

        entries.push({
          row: i + 1,
          date: dateStr,
          stockId: stockId,
          url: cols.url !== undefined ? String(data[i][cols.url] || "").trim() : "",
          typ: typ,
          person: cols.person !== undefined ? String(data[i][cols.person] || "").trim() : "",
          teil: cols.teil !== undefined ? String(data[i][cols.teil] || "").trim() : "",
          preis: cols.preis !== undefined ? String(data[i][cols.preis] || "").trim() : "",
          artikel: cols.artikel !== undefined ? String(data[i][cols.artikel] || "").trim() : "",
          status: cols.status !== undefined ? String(data[i][cols.status] || "").trim() : ""
        });
      }

      entries.sort(function(a, b) {
        var ma = String(a.date).match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
        var mb = String(b.date).match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
        if (ma && mb) {
          var da = new Date(parseInt(ma[3], 10), parseInt(ma[2], 10) - 1, parseInt(ma[1], 10)).getTime();
          var db = new Date(parseInt(mb[3], 10), parseInt(mb[2], 10) - 1, parseInt(mb[1], 10)).getTime();
          return db - da;
        }
        return 0;
      });

      return { success: true, entries: entries };
    } catch (err) {
      return { success: false, message: err.message, entries: [] };
    }
  }

function updateNachbestellung(sheetRow, fieldName, value) {
    try {
      var ss = SpreadsheetApp.openById(NACHBESTELL_SHEET_ID);
      var sheet = ss.getSheetByName(NACHBESTELL_TAB);
      if (!sheet) return { success: false, message: "Tab nicht gefunden!" };

      var lastCol = Math.max(1, Math.min(15, sheet.getLastColumn()));
      var headerData = sheet.getRange(1, 1, Math.min(10, sheet.getLastRow()), lastCol).getValues();
      var headerIdx = -1;
      for (var h = 0; h < headerData.length; h++) {
        var row = headerData[h].map(function(c) { return String(c || "").toLowerCase(); });
        if (row.some(function(c) { return c.indexOf("stock") !== -1; })) { headerIdx = h; break; }
      }
      if (headerIdx === -1) headerIdx = 0;

      var header = headerData[headerIdx];
      var colMap = {};
      for (var c = 0; c < header.length; c++) {
        var txt = String(header[c] || "").toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
        if (txt.indexOf("teil") !== -1 || txt.indexOf("benennung") !== -1 || txt.indexOf("ersatzteil") !== -1) colMap["teil"] = c + 1;
        if (txt.indexOf("artikelnr") !== -1 || txt.indexOf("artikelnummer") !== -1 || txt.indexOf("teilenr") !== -1) colMap["artikel"] = c + 1;
        if (txt.indexOf("status") !== -1 || txt.indexOf("bestellt") !== -1 || txt.indexOf("angeliefert") !== -1) colMap["status"] = c + 1;
      }

      var targetCol = colMap[fieldName];
      if (!targetCol) return { success: false, message: "Spalte '" + fieldName + "' nicht gefunden!" };

      sheet.getRange(sheetRow, targetCol).setValue(value);
      SpreadsheetApp.flush();
      return { success: true, message: "Gespeichert!" };
    } catch (err) {
      return { success: false, message: err.message };
    }
  }