// Nichts anfassen danke.

const HEMAU_SHEET_ID = "13Oh7gDT8NAul2s0cwQUeaGwMcS3B2MYu0QOdFNMhXzM";
const HEMAU_TAB_NAME = "Refurbisment List";

function setupAuth() {
  SpreadsheetApp.openById(HEMAU_SHEET_ID);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Reifen Annahme Lager')
    .addItem('HUD öffnen', 'openReifenHUD')
    .addToUi();
}

function openReifenHUD() {
  var html = HtmlService.createHtmlOutputFromFile('ReifenHUD')
    .setWidth(420)
    .setHeight(620);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Reifen Annahme HUD');
}

function getAvailableStockIds() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = Math.max(1, sheet.getLastRow());
    var lastCol = Math.max(1, sheet.getLastColumn());
    var headerData = sheet.getRange(1, 1, Math.min(30, lastRow), lastCol).getValues();
    
    var headerIdx = findHeaderRow(headerData, ["stockid", "stock"]);
    if (headerIdx === -1) return [];
    
    var stockCol = getColIndex(headerData[headerIdx], ["stockid", "stock"]);
    if (stockCol === -1) return [];
    
    var numRows = lastRow - headerIdx - 1;
    if (numRows <= 0) return [];
    
    var colData = sheet.getRange(headerIdx + 2, stockCol, numRows, 1).getValues();
    var ids = [];
    for (var i = 0; i < colData.length; i++) {
      var val = String(colData[i][0] || "").replace(/\s+/g, '').toUpperCase();
      if (val !== "") {
        ids.push(val);
      }
    }
    return ids;
  } catch (err) {
    return [];
  }
}

function getColIndex(headerRow, searchTerms) {
  if (!headerRow) return -1;
  for (var i = 0; i < headerRow.length; i++) {
    var cellText = String(headerRow[i]).toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
    for (var j = 0; j < searchTerms.length; j++) {
      var term = searchTerms[j].toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
      if (cellText === term || cellText.indexOf(term) !== -1) {
        return i + 1;
      }
    }
  }
  return -1;
}

function findHeaderRow(data, searchTerms) {
  for (var i = 0; i < Math.min(30, data.length); i++) {
    if (getColIndex(data[i], searchTerms) !== -1) {
      return i;
    }
  }
  return -1;
}

function findRowFast(sheet, searchTermsHeader, stockId) {
  var lastRow = Math.max(1, sheet.getLastRow());
  var headerData = sheet.getRange(1, 1, Math.min(30, lastRow), 80).getValues();
  
  var headerIdx = findHeaderRow(headerData, searchTermsHeader);
  if (headerIdx === -1) return { row: -1, headerIdx: -1, stockCol: -1 };
  
  var stockCol = getColIndex(headerData[headerIdx], searchTermsHeader);
  if (stockCol === -1) return { row: -1, headerIdx: headerIdx, stockCol: -1 };
  
  var colData = sheet.getRange(1, stockCol, lastRow, 1).getValues();
  for (var i = headerIdx + 1; i < colData.length; i++) {
    var currentStock = String(colData[i][0] || "").replace(/\s+/g, '').toUpperCase();
    if (currentStock === stockId) {
      return { row: i + 1, headerIdx: headerIdx, stockCol: stockCol };
    }
  }
  return { row: -1, headerIdx: headerIdx, stockCol: stockCol };
}

function checkStock(stockId) {
  try {
    stockId = stockId.replace(/\s+/g, '').toUpperCase();
    if (!stockId) return { found: false, message: "Bitte eine Stock-ID eingeben." };

    var ssSeng = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSeng = ssSeng.getActiveSheet();
    
    var sengSearch = findRowFast(sheetSeng, ["stockid", "stock"], stockId);
    
    if (sengSearch.headerIdx === -1) return { found: false, message: "Kopfzeile 'Stock ID' in Reifen Seng Liste nicht gefunden!" };
    if (sengSearch.row === -1) return { found: false, message: "Stock ID '" + stockId + "' in der aktuellen Reifen Seng Liste NICHT gefunden!" };

    return { found: true, message: "Stock ID gefunden! Bitte Status auswählen:" };

  } catch (err) {
    return { found: false, message: "Systemfehler: " + err.message };
  }
}

function processStock(stockId, isDelivered) {
  try {
    stockId = stockId.replace(/\s+/g, '').toUpperCase();
    
    var ssSeng = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSeng = ssSeng.getActiveSheet();
    
    var ssHemau = SpreadsheetApp.openById(HEMAU_SHEET_ID);
    var sheetHemau = ssHemau.getSheetByName(HEMAU_TAB_NAME);
    
    var tireInfo = "UNBEKANNT _X";
    
    var sengSearch = findRowFast(sheetSeng, ["stockid", "stock"], stockId);
    if (sengSearch.row !== -1) {
        var headerDataSeng = sheetSeng.getRange(sengSearch.headerIdx + 1, 1, 1, 80).getValues();
        var angeliefertCol = getColIndex(headerDataSeng[0], ["angeliefert"]);
        
        var mengeCol = getColIndex(headerDataSeng[0], ["menge", "anzahl"]);
        var groesseCol = getColIndex(headerDataSeng[0], ["größe", "groesse"]);
        var lastIndexCol = getColIndex(headerDataSeng[0], ["lastindex", "last"]);
        var gwIndexCol = getColIndex(headerDataSeng[0], ["gwindex", "gw"]);
        
        var mengeVal = mengeCol !== -1 ? sheetSeng.getRange(sengSearch.row, mengeCol).getValue() : "X";
        var groesseVal = groesseCol !== -1 ? sheetSeng.getRange(sengSearch.row, groesseCol).getValue() : "GRÖSSE";
        var lastIndexVal = lastIndexCol !== -1 ? sheetSeng.getRange(sengSearch.row, lastIndexCol).getValue() : "";
        var gwIndexVal = gwIndexCol !== -1 ? sheetSeng.getRange(sengSearch.row, gwIndexCol).getValue() : "";
        
        tireInfo = String(groesseVal).trim() + " " + String(lastIndexVal).trim() + String(gwIndexVal).trim() + " _" + String(mengeVal).trim();
        
        if (angeliefertCol !== -1) {
            var statusSeng = isDelivered ? "Ja" : "Nein";
            var colorSeng = isDelivered ? "#00FF00" : "#FF0000";
            sheetSeng.getRange(sengSearch.row, angeliefertCol).setValue(statusSeng);
            
            var startCol = Math.min(sengSearch.stockCol, angeliefertCol);
            var numCols = Math.abs(angeliefertCol - sengSearch.stockCol) + 1;
            sheetSeng.getRange(sengSearch.row, startCol, 1, numCols).setBackground(colorSeng);
        }
    }

    var hemauMsg = "";
    var locationText = "Kiste hat keinen Lagerplatz";
    
    if (sheetHemau) {
        var lastRowHemau = Math.max(1, sheetHemau.getLastRow());
        var hemauDataStock = sheetHemau.getRange(1, 2, lastRowHemau, 1).getValues();
        var hemauRow = -1;

        for (var i = 0; i < hemauDataStock.length; i++) {
            if (String(hemauDataStock[i][0]).replace(/\s+/g, '').toUpperCase() === stockId) {
                hemauRow = i + 1;
                break;
            }
        }

        if (hemauRow !== -1) {
            var oldLocationCol = 28;
            var oldLocation = String(sheetHemau.getRange(hemauRow, oldLocationCol).getValue() || "").trim();
            if (oldLocation !== "") {
                locationText = "Kiste steht in Regal " + oldLocation;
            }
            
            var kommentarCol = 25;
            var regalCol = 30;
            
            if (isDelivered) {
                var currentComment = String(sheetHemau.getRange(hemauRow, kommentarCol).getValue() || "");
                if (currentComment.indexOf("Reifen da //") === -1) {
                    var newComment = currentComment ? "Reifen da // " + currentComment : "Reifen da // ";
                    sheetHemau.getRange(hemauRow, kommentarCol).setValue(newComment);
                }
                sheetHemau.getRange(hemauRow, regalCol).setValue("Werkstatt 1");
            } else {
                sheetHemau.getRange(hemauRow, regalCol).setValue("Reifen nicht vorhanden");
            }
            hemauMsg = "& Refurbishment aktualisiert!";
        } else {
            hemauMsg = "(Refurbishment übersprungen: ID in Liste nicht gefunden)";
            locationText = "Lagerplatz unbekannt";
        }
    } else {
        hemauMsg = "(Refurbishment Tabellenblatt nicht gefunden)";
        locationText = "Lagerplatz unbekannt";
    }

    return { 
        success: true, 
        message: "Gebucht " + hemauMsg,\n        stockId: stockId,
        tireInfo: tireInfo,
        locationText: locationText
    };

  } catch (err) {
    return { success: false, message: "Fehler: " + err.message };
  }
}
