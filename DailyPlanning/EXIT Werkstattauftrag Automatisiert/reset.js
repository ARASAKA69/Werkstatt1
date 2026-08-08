function resetInput() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input");
    
    sheet.getRange("B2:B45").clearContent();
    sheet.getRange("D24:D28").clearContent();
    sheet.getRange("D32:D36").clearContent();
    sheet.getRange("E2:E45").clearContent();
    sheet.getRange("H2:H45").clearContent();
    sheet.getRange("J2:O45").clearContent();
    sheet.getRange("Q2:V45").clearContent();
  }