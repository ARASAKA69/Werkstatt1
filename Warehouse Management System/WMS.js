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
      var data = sheet.getRange(1, 1, lastRow, 30).getValues();
      var hitRow = -1;
      var result = { success: false };
  
      var counts = {};
      for (var i = 1; i <= 9; i++) {
        for (var j = 1; j <= 8; j++) {
          counts["Regal " + i + "." + j] = 0;
        }
      }
  
      for (var r = 1; r < data.length; r++) {
        var regalVal = String(data[r][27] || "").trim();
        if (counts.hasOwnProperty(regalVal)) counts[regalVal]++;
        if (cellMatchesStockId(data[r][1], stockId)) hitRow = r;
      }
  
      var availableShelves = [];
      for (var shelf in counts) {
        if (counts[shelf] < 5) {
          availableShelves.push({ name: shelf, count: counts[shelf] });
        }
      }
      result.freeShelves = availableShelves;
  
      if (hitRow !== -1) {
        result.success = true;
        result.carolUrl = String(data[hitRow][2] || "");
        result.schaeden = String(data[hitRow][22] || "");
        result.kommBestellung = String(data[hitRow][23] || "");
        result.kommAnlieferung = String(data[hitRow][24] || "");
        result.status = String(data[hitRow][25] || "");
        result.regal = String(data[hitRow][27] || "");
        result.reifenStatus = String(data[hitRow][29] || "");
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