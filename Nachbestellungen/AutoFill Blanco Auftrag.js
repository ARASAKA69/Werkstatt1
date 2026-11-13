const ZIEL_TABELLEN_ID = "1nE6SErc1-jmZYd_Ydviw28Pa5qdJmwNepXCiVbsdsVo";
const ZIEL_TABELLENBLATT_NAME = "BLANCO Reparaturauftrag";
const MEINE_EMAIL = "francesco.berger@auto1.com";
const HEMAU_SHEET_ID = "13Oh7gDT8NAul2s0cwQUeaGwMcS3B2MYu0QOdFNMhXzM";
const HEMAU_REFURB_TAB = "Refurbisment List";
const HEMAU_REFURB_STOCK_COL = 2;
const HEMAU_REFURB_MARKE_COL = 13;

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
    var sheet = ss.getSheetByName(HEMAU_REFURB_TAB);
    if (!sheet) return "";
    var lastRow = Math.max(2, sheet.getLastRow());
    var stockColData = sheet.getRange(1, HEMAU_REFURB_STOCK_COL, lastRow, 1).getValues();
    var markeColData = sheet.getRange(1, HEMAU_REFURB_MARKE_COL, lastRow, 1).getValues();
    var found = "";
    for (var s = 1; s < stockColData.length; s++) {
      if (!cellMatchesStockIdLocal(stockColData[s][0], stockId)) continue;
      var marke = String((markeColData[s] && markeColData[s][0]) || "").trim();
      if (marke) found = marke;
    }
    return found;
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
  if (t.indexOf("verschwunden") !== -1) return true;
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
