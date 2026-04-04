// Nicht anfassen, gehört zu meiner Automatisierung vom Carol für Fertige Aufträge. Danke, Francesco B.


function doGet(e) {
  var stockIdRaw = e.parameter.stock;
  if (!stockIdRaw) return ContentService.createTextOutput("Fehler: Keine Stock-ID empfangen");

  var stockId = String(stockIdRaw).trim().toUpperCase(); 
  var log = [];

  try {
    var ss1 = SpreadsheetApp.getActiveSpreadsheet();
    var sheet1 = ss1.getSheetByName("Stock ID extern Tracking");
    
    if (sheet1) {
      var data1 = sheet1.getRange("A:A").getValues();
      var idGefunden1 = false;
      
      for (var i = 0; i < data1.length; i++) {
        var cellValue = String(data1[i][0]).trim().toUpperCase(); 
        
        if (cellValue === stockId) {
          var dateStr = Utilities.formatDate(new Date(), "Europe/Berlin", "dd.MM.yyyy");
          sheet1.getRange(i + 1, 9).setValue(dateStr);
          log.push("Job 1 (Datum): ERFOLG");
          idGefunden1 = true;
          break;
        }
      }
      if (!idGefunden1) log.push("Job 1 Fehler: Stock-ID '" + stockId + "' in Spalte A nicht gefunden!");
    } else {
      log.push("Job 1 Fehler: Reiter 'Stock ID extern Tracking' nicht gefunden!");
    }
  } catch (err) {
    log.push("Job 1 Crash: " + err.message);
  }

  var urlZweiteDatei = "https://docs.google.com/spreadsheets/d/13Oh7gDT8NAul2s0cwQUeaGwMcS3B2MYu0QOdFNMhXzM/edit?pli=1&gid=1006156704#gid=1006156704"; 
  
  try {
    var ss2 = SpreadsheetApp.openByUrl(urlZweiteDatei);
    var sheet2 = ss2.getSheetByName("Refurbisment List"); 
    
    if (sheet2) {
      var data2 = sheet2.getRange("B:B").getValues();
      var idGefunden2 = false;
      
      for (var j = 0; j < data2.length; j++) {
        var cellValue2 = String(data2[j][0]).trim().toUpperCase();
        
        if (cellValue2 === stockId) {
          var row = j + 1;
          
          var oldRegal = sheet2.getRange(row, 28).getValue();
          if(!oldRegal) oldRegal = "LEER";
          
          sheet2.getRange(row, 25).setBackground("#00FF00");
          sheet2.getRange(row, 26).setValue("Herausgegeben");
          sheet2.getRange(row, 28).setValue("Tagesliste");
          
          log.push("Job 2 (Status): ERFOLG");
          log.push("OLD_REGAL:" + oldRegal);
          
          idGefunden2 = true;
          break;
        }
      }
      if (!idGefunden2) log.push("Job 2 Fehler: Stock-ID '" + stockId + "' in Spalte B nicht gefunden!");
    } else {
      log.push("Job 2 Fehler: Reiter 'Refurbisment List' nicht gefunden!");
    }
  } catch (err) {
    log.push("Job 2 Crash: " + err.message);
  }

  return ContentService.createTextOutput(log.join(" | "));
}