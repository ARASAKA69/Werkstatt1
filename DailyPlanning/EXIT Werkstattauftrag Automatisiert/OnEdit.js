function onEdit(e) {
    var ss = e.source;
    var range = e.range;
    var sheet = range.getSheet();
    
    // Nur reagieren, wenn wir auf dem Blatt "Input" sind
    if(sheet.getName() !== "Input") return;
    
    // Prüfen, ob die geänderte Zelle im Bereich B29:B41 liegt
    if(range.getA1Notation().match(/^B(2[9-9]|3[0-9]|4[0-1])$/)) {
      updateReparaturauftrag();
    }
  }
  
  // Funktion, die die Werte zentriert auf Reparaturauftrag schreibt
  function updateReparaturauftrag() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var inputSheet = ss.getSheetByName("Input");
    var outputSheet = ss.getSheetByName("Reparaturauftrag");
    
    // Bereich der X-Markierungen und Inhalte
    var xRange = inputSheet.getRange("B29:B41").getValues();
    var contentRange = inputSheet.getRange("C29:C41").getValues();
    
    // Alle Inhalte sammeln, die ein "x" haben
    var values = [];
    for (var i = 0; i < xRange.length; i++) {
      if (xRange[i][0] === "x") {
        values.push(contentRange[i][0]);
      }
    }
    
    // Zielbereich: G9:G12 (4 Zeilen, 1 Spalte)
    var rows = 4;
    var totalCells = rows;
    
    // Clear vorheriger Inhalt + Formatierung
    outputSheet.getRange("G9:G12").clearContent().clearFormat();
    
    // Berechnen der Startposition für zentrierte Anzeige
    var startIndex = Math.floor((totalCells - values.length) / 2);
    
    for (var i = 0; i < values.length; i++) {
      var cellIndex = startIndex + i;
      var row = cellIndex + 9; // +9, da G9=Zeile 9
      var cell = outputSheet.getRange(row, 7); // Spalte G=7
      cell.setValue(values[i]);
      cell.setFontFamily("Arial");
      cell.setFontSize(33);
      cell.setHorizontalAlignment("center");
      cell.setVerticalAlignment("middle");
    }
  }
  