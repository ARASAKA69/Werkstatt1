function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (sheetName === "Nachbestellungen" && row >= 2) {
    if (col === 2) {
      var valB = e.range.getValue();
      if (valB !== "" && valB != null) {
        sheet.getRange(row, 1).setValue(new Date());
      } else {
        sheet.getRange(row, 1).clearContent();
      }
    }

    if (col === 4 || col === 10 || col === 12) {
      var valD = sheet.getRange(row, 4).getValue();
      var valJ = sheet.getRange(row, 10).getValue();
      var valL = sheet.getRange(row, 12).getValue();

      var D = (valD || "").toString().toLowerCase();
      var J = (valJ || "").toString().toLowerCase();
      var L = (valL || "").toString().toLowerCase();

      var color = null;
      if (J.includes("push to senior")) color = "#e5d4ff";
      else if (L.includes("angeliefert") || L.includes("bereit für tagesliste")) color = "#d4ffd4";
      else if (J.includes("anfrage alfah")) color = "#ffd8a8";
      else if (L.includes("nicht bestellt")) color = "#ffffff";
      else if (L.includes("bestellt")) color = "#fff5a5";

      sheet.getRange(row, 1, 1, 14).setBackground(color);

      if (col === 4) {
        if (D.includes("exit bestellung")) {
          sheet.getRange(row, 2).setBackground("#FFC4C4");
        } else {
          sheet.getRange(row, 2).setBackground(null);
        }
      }
    }
  }
}
