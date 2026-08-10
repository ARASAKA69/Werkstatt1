function resetInput() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Input");
  if (sheet) {
    sheet.getRange("B2:B45").clearContent();
    sheet.getRange("D24:D28").clearContent();
    sheet.getRange("D32:D36").clearContent();
    sheet.getRange("E2:E45").clearContent();
    sheet.getRange("H2:H45").clearContent();
    sheet.getRange("J2:O45").clearContent();
    sheet.getRange("Q2:V45").clearContent();
  }

  var rep = ss.getSheetByName("Reparaturauftrag");
  if (!rep) return;

  var saved = [];
  for (var r = 13; r <= 39; r++) {
    for (var c = 1; c <= 9; c++) {
      try {
        var fml = rep.getRange(r, c).getFormula();
        if (fml) saved.push({ row: r, col: c, formula: fml });
      } catch (e1) {}
    }
  }

  for (var rr = 14; rr <= 39; rr++) {
    for (var cc = 1; cc <= 9; cc++) {
      try {
        var cell = rep.getRange(rr, cc);
        var merges = cell.getMergedRanges();
        if (merges && merges.length) {
          if (merges[0].getRow() !== rr || merges[0].getColumn() !== cc) continue;
          merges[0].clearContent();
        } else {
          cell.clearContent();
        }
      } catch (e2) {}
    }
  }

  for (var i = 0; i < saved.length; i++) {
    try { rep.getRange(saved[i].row, saved[i].col).setFormula(saved[i].formula); } catch (e3) {}
  }

  try { rep.getRange("G9:G12").clearContent(); } catch (e4) {}
}
