const TRACKING_SHEET_URL = "https://docs.google.com/spreadsheets/d/1PuCLw8UmDjB_pBo_jCZ9rmSD3GJQESHzPoBVu_--MRo/edit?gid=1453769469#gid=1453769469";

function normalizeStockId(value) {
    return String(value || "").replace(/\s+/g, "").toUpperCase();
  }

function cellMatchesStockId(cellVal, stockId) {
    var cv = normalizeStockId(cellVal);
    var sid = normalizeStockId(stockId);
    return sid !== "" && cv === sid;
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
      var data = sheet.getRange(1, 1, lastRow, 28).getValues();
      var entries = [];

      for (var i = 1; i < data.length; i++) {
        var stockId = String(data[i][1] || "").trim();
        var regal = String(data[i][27] || "").trim();
        if (!stockId) continue;
        if (!/^Regal\s+\d+\.\d+$/i.test(regal)) continue;
        entries.push({
          stockId: stockId,
          regal: regal,
          sortKey: regal
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