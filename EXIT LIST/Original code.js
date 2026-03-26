function onEdit(e) {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    const range = e.range;
  
    const row = range.getRow();
    const col = range.getColumn();
  
    // ==========================================
    // EXIT REPAIR (Farblogik + Datum bei BESTELLT / KOMPLETT ANGELIEFERT)
    // ==========================================
    if (sheetName === "Exit Repair") {
  
      if (row < 2) return;
  
      const valueG = sheet.getRange(row, 7).getDisplayValue().trim().toUpperCase(); // G
      const valueH = sheet.getRange(row, 8).getDisplayValue().trim().toUpperCase(); // H
      const valueJ = sheet.getRange(row, 10).getDisplayValue().trim().toUpperCase(); // J
  
      const targetRange = sheet.getRange(row, 1, 1, 10); // A bis J
      const timestampCell = sheet.getRange(row, 9); // Spalte I
  
      // ============================
      // Datum setzen/löschen
      // ============================
      if (col === 8) { // nur reagieren, wenn H geändert wird
        const now = new Date();
  
        if (valueH === "BESTELLT") {
          // nur setzen, wenn leer
          if (timestampCell.getValue() === "") {
            timestampCell.setValue(now);
            timestampCell.setNumberFormat("dd.MM.yyyy HH:mm");
          }
        } else if (valueH === "komplett angeliefert") {
          // immer aktualisieren
          timestampCell.setValue(now);
          timestampCell.setNumberFormat("dd.MM.yyyy HH:mm");
        } else {
          // alle anderen Werte -> löschen
          timestampCell.setValue("");
        }
      }
  
      // ============================
      // FARBE (läuft bei jeder Änderung in der Zeile)
      // ============================
      if (valueJ === "JA") {
        targetRange.setBackground("#b6fcb6");
        return;
      }
  
      if (valueH === "komplett angeliefert") {
        targetRange.setBackground("#ffa500");
        return;
      }
  
      if (valueH === "BESTELLT") {
        targetRange.setBackground("#fff2b3");
        return;
      }
  
      if (valueG === "B2A1") {
        targetRange.setBackground("#ff0000");
        return;
      }
  
      if (valueG === "PUSH TO SENIOR") {
        targetRange.setBackground("#d9b3ff");
        return;
      }
  
      if (valueG === "DURCH SENIOR GENEHMIGT") {
        targetRange.setBackground("#b3d9ff");
        return;
      }
  
      targetRange.setBackground(null);
    }
  
    // ==========================================
    // TIMESTAMP LOGIK (Spalte B -> Zeit in A)
    // ==========================================
    if (col === 2 && row >= 2) {
      const targetCell = sheet.getRange(row, 1);
  
      if (range.getDisplayValue().trim() === "") {
        targetCell.setValue("");
      } else {
        const now = new Date();
        targetCell.setValue(now);
        targetCell.setNumberFormat("dd.MM.yyyy HH:mm");
      }
    }
  }  