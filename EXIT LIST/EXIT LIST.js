function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  var range = e.range;
  var startRow = range.getRow();
  var endRow = range.getLastRow();
  var startCol = range.getColumn();
  var endCol = range.getLastColumn();
  var r;
  var touchesH = startCol <= 8 && endCol >= 8;
  var touchesB = startCol <= 2 && endCol >= 2;

  if (sheetName === "Exit Repair") {
    for (r = startRow; r <= endRow; r++) {
      if (r < 2) continue;
      applyExitRepairRowColors_(sheet, r, touchesH);
    }
  }

  if (sheetName === "Exit Repair" && touchesB) {
    for (r = startRow; r <= endRow; r++) {
      if (r < 2) continue;
      var bCell = sheet.getRange(r, 2);
      var aCell = sheet.getRange(r, 1);
      if (bCell.getDisplayValue().trim() === "") {
        aCell.setValue("");
      } else {
        var nowB = new Date();
        aCell.setValue(nowB);
        aCell.setNumberFormat("dd.MM.yyyy HH:mm");
      }
    }
  }
}

function exitNorm_(v) {
  return String(v || "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function applyExitRepairRowColors_(sheet, row, applyTimestampForH) {
  var valueG = exitNorm_(sheet.getRange(row, 7).getDisplayValue());
  var valueH = exitNorm_(sheet.getRange(row, 8).getDisplayValue());
  if (!valueH) {
    valueH = exitNorm_(sheet.getRange(row, 8).getValue());
  }
  var valueJ = exitNorm_(sheet.getRange(row, 10).getDisplayValue());
  if (!valueJ) {
    valueJ = exitNorm_(sheet.getRange(row, 10).getValue());
  }

  var targetRange = sheet.getRange("A" + row + ":J" + row);
  var timestampCell = sheet.getRange(row, 9);

  if (valueJ === "JA") {
    targetRange.setBackground("#b6fcb6");
  } else if (valueH === "KOMPLETT ANGELIEFERT") {
    targetRange.setBackground("#ffa500");
  } else if (valueH === "BESTELLT") {
    targetRange.setBackground("#fff2b3");
  } else if (valueG === "B2A1") {
    targetRange.setBackground("#ff0000");
  } else if (valueG === "PUSH TO SENIOR") {
    targetRange.setBackground("#d9b3ff");
  } else if (valueG === "DURCH SENIOR GENEHMIGT") {
    targetRange.setBackground("#b3d9ff");
  } else {
    targetRange.setBackground("#ffffff");
  }

  if (applyTimestampForH) {
    var now = new Date();
    if (valueH === "BESTELLT") {
      if (timestampCell.getValue() === "") {
        timestampCell.setValue(now);
        timestampCell.setNumberFormat("dd.MM.yyyy HH:mm");
      }
    } else if (valueH === "KOMPLETT ANGELIEFERT") {
      timestampCell.setValue(now);
      timestampCell.setNumberFormat("dd.MM.yyyy HH:mm");
    } else {
      timestampCell.setValue("");
    }
  }
}
