const ZIEL_TABELLEN_ID = "1nE6SErc1-jmZYd_Ydviw28Pa5qdJmwNepXCiVbsdsVo";
const ZIEL_TABELLENBLATT_NAME = "BLANCO Reparaturauftrag";
const MEINE_EMAIL = "francesco.berger@auto1.com";
const HEMAU_SHEET_ID = "13Oh7gDT8NAul2s0cwQUeaGwMcS3B2MYu0QOdFNMhXzM";
const HEMAU_DAILY_PLANNING_TAB = "Daily Planning List";

function normalizeStockIdLocal(value) {
  return String(value || "").replace(/\s+/g, "").toUpperCase();
}

function cellMatchesStockIdLocal(cellVal, stockId) {
  var cv = normalizeStockIdLocal(cellVal);
  var sid = normalizeStockIdLocal(stockId);
  return sid !== "" && cv === sid;
}

function lookupHemauMarkeModellLocal(stockId) {
  try {
    stockId = normalizeStockIdLocal(stockId);
    if (!stockId) return "";
    var ss = SpreadsheetApp.openById(HEMAU_SHEET_ID);
    var sheet = ss.getSheetByName(HEMAU_DAILY_PLANNING_TAB);
    if (!sheet) return "";
    var lastRow = Math.max(1, sheet.getLastRow());
    var lastCol = Math.max(1, Math.min(80, sheet.getLastColumn()));
    var headerData = sheet.getRange(1, 1, Math.min(30, lastRow), lastCol).getValues();
    var headerIdx = -1;
    var stockCol = -1;
    var markeCol = -1;
    for (var h = 0; h < headerData.length; h++) {
      for (var c = 0; c < headerData[h].length; c++) {
        var txt = String(headerData[h][c] || "").toLowerCase().replace(/[^a-z0-9äöüß]/g, "");
        if (stockCol === -1 && txt.indexOf("stock") !== -1) {
          headerIdx = h;
          stockCol = c + 1;
        }
        if (markeCol === -1 && (txt.indexOf("marke") !== -1 || txt.indexOf("modell") !== -1)) {
          markeCol = c + 1;
        }
      }
      if (stockCol !== -1) break;
    }
    if (stockCol === -1) stockCol = 2;
    if (markeCol === -1) markeCol = 9;
    if (headerIdx === -1) headerIdx = 0;
    var colData = sheet.getRange(1, stockCol, lastRow, 1).getValues();
    for (var i = headerIdx + 1; i < colData.length; i++) {
      if (cellMatchesStockIdLocal(colData[i][0], stockId)) {
        return String(sheet.getRange(i + 1, markeCol).getValue() || "").trim();
      }
    }
    return "";
  } catch (err) {
    return "";
  }
}

function nachbestellungTypShouldPrintWerkstattauftragLocal(typ) {
  var t = String(typ || "").trim().toLowerCase();
  if (!t) return false;
  if (t.indexOf("exit") !== -1) return false;
  if (t.indexOf("erstbestellung") !== -1 && t.indexOf("falsch") !== -1) return true;
  if (t.indexOf("mechanik") !== -1 && t.indexOf("nachbestellung") !== -1) return true;
  if (t.indexOf("q-check") !== -1 || t.indexOf("qcheck") !== -1) return true;
  return false;
}

function getNachbestellungTypLocal(sheet, row) {
  var lastCol = Math.max(1, Math.min(80, sheet.getLastColumn()));
  for (var hr = 1; hr <= 3; hr++) {
    var header = sheet.getRange(hr, 1, 1, lastCol).getValues()[0];
    for (var c = 0; c < header.length; c++) {
      var txt = String(header[c] || "").toLowerCase().replace(/[^a-z0-9äöüß]/g, "");
      if (txt.indexOf("typ") !== -1) {
        return String(sheet.getRange(row, c + 1).getValue() || "").trim();
      }
    }
  }
  return "";
}

function autoFillAuftrag(e) {
  if (!e || !e.range) return;

  var bearbeiter = Session.getActiveUser().getEmail();
  if (bearbeiter !== MEINE_EMAIL) return;

  var sheet = e.range.getSheet();
  if (sheet.getName() !== "Nachbestellung") return;

  var col = e.range.getColumn();
  var row = e.range.getRow();

  if (col === 11 && row >= 4) {
    var newValue = String(e.range.getValue());

    if (newValue.indexOf("komplett angeliefert") !== -1) {
      var stockId = sheet.getRange(row, 2).getValue();
      var beschreibung = sheet.getRange(row, 5).getValue();
      var typ = getNachbestellungTypLocal(sheet, row);
      if (!nachbestellungTypShouldPrintWerkstattauftragLocal(typ)) return;

      try {
        var zielSs = SpreadsheetApp.openById(ZIEL_TABELLEN_ID);
        var zielSheet = zielSs.getSheetByName(ZIEL_TABELLENBLATT_NAME);

        if (zielSheet) {
          zielSheet.getRange("D10").setValue(stockId);
          zielSheet.getRange("D18").setValue(beschreibung);
          var markeModell = lookupHemauMarkeModellLocal(stockId);
          if (markeModell) {
            zielSheet.getRange("B13:D13").setValue(markeModell);
          } else {
            zielSheet.getRange("B13:D13").clearContent();
          }
        }
      } catch (err) {}
    }
  }
}
