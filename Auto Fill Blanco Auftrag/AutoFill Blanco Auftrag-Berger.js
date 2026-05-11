const ZIEL_TABELLEN_ID = "1nE6SErc1-jmZYd_Ydviw28Pa5qdJmwNepXCiVbsdsVo";
const ZIEL_TABELLENBLATT_NAME = "BLANCO Reparaturauftrag";
const MEINE_EMAIL = "francesco.berger@auto1.com";

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
      var ui = SpreadsheetApp.getActiveSpreadsheet();
      
      try {
        var zielSs = SpreadsheetApp.openById(ZIEL_TABELLEN_ID);
        var zielSheet = zielSs.getSheetByName(ZIEL_TABELLENBLATT_NAME);
        
        if (zielSheet) {
          zielSheet.getRange("D10").setValue(stockId);
          zielSheet.getRange("D18").setValue(beschreibung);
          ui.toast("✅ StockID " + stockId + " ans Auftragsblatt gesendet!", "AS", 4);
        } else {
          ui.toast("❌ Fehler: Auftragsblatt nicht gefunden!", "AS", 4);
        }
      } catch (err) {
        ui.toast("❌ Fehler beim Senden: " + err.message, "AS", 5);
      }
    }
  }
}
