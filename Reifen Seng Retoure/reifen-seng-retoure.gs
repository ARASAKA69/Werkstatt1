var CONFIG = {
  senderEmail: 'francesco.berger@auto1.com',
  recipientEmail: 'francesco.berger@auto1.com',
  recipientEmail2: '',
  stockIdCol: 1,
  lieferscheinNrCol: 2,
  artikelCol: 3,
  beschreibungCol: 4,
  anzahlCol: 5,
  statusCol: 6,
  headerRow: 1
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Per Email senden an Seng')
    .addItem('Per Email senden', 'sendFilledRowsByEmail')
    .addToUi();
}

function sendFilledRowsByEmail() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = collectFilledRows_(sheet);

  if (!rows.length) {
    ui.alert('Keine vollstaendig ausgefuellten Zeilen gefunden.');
    return;
  }

  var recipients = buildRecipientList_();
  if (!recipients.length) {
    ui.alert('Bitte zuerst CONFIG.recipientEmail im Script eintragen.');
    return;
  }

  var pdfBlob = buildFilledRowsPdf_(sheet, rows);
  var subject = 'Reifen Seng Retoure - ' + sheet.getName();
  var body = 'Anbei die Retourenliste als PDF.\n\n' +
    'Bei Rueckfragen wenden Sie sich bitte an ersatzteile.hemau@autohero.com.';

  GmailApp.sendEmail(recipients.join(','), subject, body, {
    attachments: [pdfBlob],
    name: 'Reifen Seng Retoure'
  });

  ui.alert('E-Mail mit ' + rows.length + ' Position(en) an ' + recipients.join(', ') + ' gesendet.');
}

function buildRecipientList_() {
  var list = [];
  if (CONFIG.recipientEmail && CONFIG.recipientEmail.indexOf('@') !== -1) {
    list.push(CONFIG.recipientEmail);
  }
  if (CONFIG.recipientEmail2 && CONFIG.recipientEmail2.indexOf('@') !== -1) {
    list.push(CONFIG.recipientEmail2);
  }
  return list;
}

function collectFilledRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.headerRow + 1) return [];

  var numRows = lastRow - CONFIG.headerRow;
  var values = sheet.getRange(CONFIG.headerRow + 1, 1, numRows, CONFIG.anzahlCol).getValues();
  var rows = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var complete = true;
    for (var c = 0; c < CONFIG.anzahlCol; c++) {
      if (row[c] === '' || row[c] === null || row[c] === undefined) {
        complete = false;
        break;
      }
    }
    if (complete) rows.push(row);
  }

  return rows;
}

function buildFilledRowsPdf_(sheet, rows) {
  var ss = sheet.getParent();
  var temp = ss.insertSheet('PDF_TEMP_' + Date.now());

  try {
    var titleRange = temp.getRange(1, 1, 1, 5);
    titleRange.merge();
    titleRange.setValue('Reifen Retoure (' + Utilities.formatDate(new Date(), 'Europe/Berlin', 'dd.MM.yyyy') + ')');
    titleRange.setFontWeight('bold');
    titleRange.setFontSize(14);
    titleRange.setHorizontalAlignment('center');

    var header = [['StockID', 'Lieferschein Nr.', 'Artikel/Reifen', 'Beschreibung', 'Anzahl']];
    temp.getRange(2, 1, 1, 5).setValues(header).setFontWeight('bold');
    temp.getRange(3, 1, rows.length, 5).setValues(rows);

    var total = 0;
    for (var i = 0; i < rows.length; i++) {
      total += Number(rows[i][4]) || 0;
    }
    var totalRow = rows.length + 3;
    temp.getRange(totalRow, 4).setValue('TOTAL:').setFontWeight('bold').setHorizontalAlignment('right');
    temp.getRange(totalRow, 5).setValue(total).setFontWeight('bold');

    var tableRange = temp.getRange(2, 1, totalRow - 1, 5);
    tableRange.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);
    tableRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    tableRange.setVerticalAlignment('middle');

    temp.setColumnWidth(1, 90);
    temp.setColumnWidth(2, 110);
    temp.setColumnWidth(3, 160);
    temp.setColumnWidth(4, 320);
    temp.setColumnWidth(5, 70);
    SpreadsheetApp.flush();

    var pdfBlob = exportSheetAsPdf_(ss, temp);
    return pdfBlob.setName(buildPdfFileName_());
  } finally {
    ss.deleteSheet(temp);
  }
}

function buildPdfFileName_() {
  var dateStr = Utilities.formatDate(new Date(), 'Europe/Berlin', 'dd.MM.yyyy');
  return 'Autohero Hemau Retourenliste (' + dateStr + ').pdf';
}

function exportSheetAsPdf_(ss, sheet) {
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export' +
    '?format=pdf' +
    '&gid=' + sheet.getSheetId() +
    '&size=A4&portrait=false&scale=1' +
    '&sheetnames=false&printtitle=false&pagenumbers=false' +
    '&gridlines=false&fzr=false' +
    '&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4';

  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  return response.getBlob();
}

function installEditTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onStockIdEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('onStockIdEdit').forSpreadsheet(ss).onEdit().create();
  SpreadsheetApp.getUi().alert('Trigger eingerichtet. Ab jetzt wird automatisch importiert, wenn du eine StockID in Spalte A eintippst.');
}

function manualRefreshActiveRow() {
  var range = SpreadsheetApp.getActiveRange();
  if (!range) return;
  processRow_(range.getSheet(), range.getRow());
}

function onStockIdEdit(e) {
  try {
    var range = e.range;
    var sheet = range.getSheet();
    var startCol = range.getColumn();
    var endCol = startCol + range.getNumColumns() - 1;
    if (CONFIG.stockIdCol < startCol || CONFIG.stockIdCol > endCol) return;

    var startRow = range.getRow();
    var endRow = startRow + range.getNumRows() - 1;
    for (var row = startRow; row <= endRow; row++) {
      if (row <= CONFIG.headerRow) continue;
      var stockId = String(sheet.getRange(row, CONFIG.stockIdCol).getValue() || '').trim();
      if (!stockId) continue;
      processRow_(sheet, row);
    }
  } catch (err) {
    Logger.log('onStockIdEdit error: ' + err.message);
  }
}

function processRow_(sheet, row) {
  var stockId = String(sheet.getRange(row, CONFIG.stockIdCol).getValue() || '').trim();
  if (!stockId) return;

  ensureStatusHeader_(sheet);
  setStatus_(sheet, row, 'Suche laeuft...');

  var message = findLatestLieferscheinMessage_(stockId);
  if (!message) {
    setStatus_(sheet, row, 'Keine E-Mail gefunden');
    return;
  }

  var attachment = findPdfAttachment_(message);
  if (!attachment) {
    setStatus_(sheet, row, 'Kein PDF-Anhang gefunden');
    return;
  }

  var data;
  try {
    data = extractLieferscheinData_(attachment.copyBlob());
  } catch (err) {
    setStatus_(sheet, row, 'PDF-Verarbeitung fehlgeschlagen');
    return;
  }

  var lieferscheinNr = extractLieferscheinNr_(data.rawText);
  var position = data.tablePosition || extractFirstPositionFromText_(data.rawText);

  sheet.getRange(row, CONFIG.lieferscheinNrCol).setValue(lieferscheinNr || '');
  sheet.getRange(row, CONFIG.artikelCol).setValue(position ? position.artikel : '');
  sheet.getRange(row, CONFIG.beschreibungCol).setValue(position ? position.beschreibung : '');
  sheet.getRange(row, CONFIG.anzahlCol).setValue(position ? position.anzahl : '');

  if (lieferscheinNr && position && position.anzahl !== '') {
    setStatus_(sheet, row, 'OK');
  } else {
    setStatus_(sheet, row, 'Bitte pruefen');
  }
}

function ensureStatusHeader_(sheet) {
  var cell = sheet.getRange(CONFIG.headerRow, CONFIG.statusCol);
  if (String(cell.getValue() || '') === '') {
    cell.setValue('Status');
    cell.setFontWeight('bold');
  }
}

function setStatus_(sheet, row, text) {
  sheet.getRange(row, CONFIG.statusCol).setValue(text);
  SpreadsheetApp.flush();
}

function findLatestLieferscheinMessage_(stockId) {
  var query = 'from:(' + CONFIG.senderEmail + ') "' + stockId + '" has:attachment';
  var threads = GmailApp.search(query, 0, 10);
  var stockIdUpper = String(stockId).toUpperCase();
  var latestAny = null;
  var latestSubjectMatch = null;

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();
    for (var m = 0; m < messages.length; m++) {
      var msg = messages[m];
      if (!latestAny || msg.getDate().getTime() > latestAny.getDate().getTime()) {
        latestAny = msg;
      }
      var subject = String(msg.getSubject() || '').toUpperCase();
      if (subject.indexOf(stockIdUpper) === -1) continue;
      if (!latestSubjectMatch || msg.getDate().getTime() > latestSubjectMatch.getDate().getTime()) {
        latestSubjectMatch = msg;
      }
    }
  }

  return latestSubjectMatch || latestAny;
}

function findPdfAttachment_(message) {
  var attachments = message.getAttachments({ includeInlineImages: false, includeAttachments: true }) || [];
  for (var i = 0; i < attachments.length; i++) {
    var contentType = String(attachments[i].getContentType() || '').toLowerCase();
    var name = String(attachments[i].getName() || '').toLowerCase();
    if (contentType.indexOf('pdf') !== -1 || name.indexOf('.pdf') !== -1) {
      return attachments[i];
    }
  }
  return null;
}

function extractLieferscheinData_(blob) {
  var docId = null;
  var rawText = '';
  var tablePosition = null;

  try {
    var inserted = driveConvertToDoc_(blob);
    docId = inserted.id;
  } catch (createErr) {
    return { rawText: '', tablePosition: null };
  }

  try {
    var doc = DocumentApp.openById(docId);
    rawText = doc.getBody().getText();
    tablePosition = extractPositionFromTables_(doc);
  } catch (readErr) {
    rawText = '';
  }

  if (!rawText) {
    rawText = exportDocAsText_(docId);
  }

  driveRemove_(docId);
  return { rawText: rawText || '', tablePosition: tablePosition };
}

function driveConvertToDoc_(blob) {
  var stamp = 'rsr_ocr_temp_' + Date.now();
  if (Drive.Files && typeof Drive.Files.create === 'function') {
    return Drive.Files.create({ name: stamp, mimeType: 'application/vnd.google-apps.document' }, blob, { ocrLanguage: 'de' });
  }
  return Drive.Files.insert({ title: stamp, mimeType: 'application/vnd.google-apps.document' }, blob, { ocr: true, ocrLanguage: 'de', convert: true });
}

function driveRemove_(fileId) {
  try {
    if (Drive.Files && typeof Drive.Files.remove === 'function') {
      Drive.Files.remove(fileId);
    } else if (Drive.Files && typeof Drive.Files.trash === 'function') {
      Drive.Files.trash(fileId);
    }
  } catch (e) {}
}

function exportDocAsText_(docId) {
  try {
    var url = 'https://www.googleapis.com/drive/v3/files/' + docId + '/export?mimeType=text%2Fplain';
    var resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() === 200) return resp.getContentText();
  } catch (e) {}
  return '';
}

function extractPositionFromTables_(doc) {
  var tables = doc.getBody().getTables();
  for (var t = 0; t < tables.length; t++) {
    var table = tables[t];
    var numRows = table.getNumRows();
    var headerRow = -1;
    var col = {};

    for (var r = 0; r < numRows; r++) {
      var row = table.getRow(r);
      var numCells = row.getNumCells();
      var idxAnzahl = -1, idxArtikel = -1, idxBeschreibung = -1;
      for (var c = 0; c < numCells; c++) {
        var cellText = row.getCell(c).getText().trim();
        if (/^Anzahl$/i.test(cellText)) idxAnzahl = c;
        if (/Artikel/i.test(cellText)) idxArtikel = c;
        if (/Beschreibung/i.test(cellText)) idxBeschreibung = c;
      }
      if (idxAnzahl > -1 && idxBeschreibung > -1) {
        headerRow = r;
        col = { anzahl: idxAnzahl, artikel: idxArtikel, beschreibung: idxBeschreibung };
        break;
      }
    }

    if (headerRow > -1 && headerRow + 1 < numRows) {
      var dataRow = table.getRow(headerRow + 1);
      var anzahlText = dataRow.getCell(col.anzahl).getText().trim();
      var artikelText = col.artikel > -1 ? dataRow.getCell(col.artikel).getText().trim() : '';
      var beschreibungText = dataRow.getCell(col.beschreibung).getText().trim();
      var anzahlNum = parseAnzahl_(anzahlText);
      if (anzahlNum !== '') {
        return { anzahl: anzahlNum, artikel: artikelText, beschreibung: beschreibungText };
      }
    }
  }
  return null;
}

function extractFirstPositionFromText_(text) {
  var lines = String(text || '').split('\n');
  for (var i = 0; i < lines.length; i++) lines[i] = lines[i].trim();
  lines = lines.filter(function(l) { return l.length > 0; });

  var headerIdx = -1;
  for (var h = 0; h < lines.length; h++) {
    if (/Anzahl/i.test(lines[h]) && /Beschreibung/i.test(lines[h])) {
      headerIdx = h;
      break;
    }
  }
  var scanLines = headerIdx >= 0 ? lines.slice(headerIdx + 1) : lines;

  for (var j = 0; j < scanLines.length; j++) {
    var line = scanLines[j];
    var anzahlMatch = line.match(/^(\d+[.,]\d+)\s+(.+)$/);
    if (!anzahlMatch) continue;

    var anzahl = parseAnzahl_(anzahlMatch[1]);
    var rest = anzahlMatch[2];

    var sameLineMatch = rest.match(/^(\d{3}\/\d{2}\s?R\d{2}\s+\d{2,3}(?:\/\d{2,3})?\s*[A-Z]{1,3})\s+(.+)$/);
    if (sameLineMatch) {
      return { anzahl: anzahl, artikel: sameLineMatch[1].trim(), beschreibung: sameLineMatch[2].trim() };
    }

    var artikel = rest;
    var next = j + 1;
    var continuationPattern = /^[A-Z0-9()\/\s]{1,10}$/;
    var maxContinuations = 3;
    while (next < scanLines.length && maxContinuations > 0 && continuationPattern.test(scanLines[next])) {
      artikel += ' ' + scanLines[next];
      next++;
      maxContinuations--;
    }

    var beschreibung = next < scanLines.length ? scanLines[next] : '';

    return { anzahl: anzahl, artikel: artikel.trim(), beschreibung: beschreibung.trim() };
  }

  return null;
}

function extractLieferscheinNr_(text) {
  var match = String(text || '').match(/Lieferschein\s*Nr\.?\s*:?\s*(\d+)/i);
  return match ? match[1] : '';
}

function parseAnzahl_(raw) {
  var n = parseFloat(String(raw || '').replace(',', '.'));
  return isNaN(n) ? '' : n;
}

function debugGmailAccount() {
  Logger.log('Effective user: ' + Session.getEffectiveUser().getEmail());
  Logger.log('Active user: ' + Session.getActiveUser().getEmail());

  var broad = GmailApp.search('from:' + CONFIG.senderEmail, 0, 5);
  Logger.log('Treffer nur nach Sender (' + CONFIG.senderEmail + '): ' + broad.length);
  for (var i = 0; i < broad.length; i++) {
    var msgs = broad[i].getMessages();
    var last = msgs[msgs.length - 1];
    Logger.log('Thread ' + i + ': ' + msgs.length + ' Nachricht(en), letzte: "' + last.getSubject() + '" (' + last.getDate() + ')');
  }

  var withStock = GmailApp.search('from:' + CONFIG.senderEmail + ' "AM11142"', 0, 5);
  Logger.log('Treffer mit StockID-Filter ("AM11142"): ' + withStock.length);

  var withAttachment = GmailApp.search('from:' + CONFIG.senderEmail + ' "AM11142" has:attachment', 0, 5);
  Logger.log('Treffer mit StockID + has:attachment: ' + withAttachment.length);
}

function testExtractionAM11142() {
  testExtraction('AM11142');
}

function testExtraction(stockId) {
  var message = findLatestLieferscheinMessage_(stockId);
  if (!message) {
    Logger.log('Keine E-Mail gefunden fuer ' + stockId);
    return;
  }
  Logger.log('E-Mail gefunden: ' + message.getSubject() + ' (' + message.getDate() + ')');

  var attachment = findPdfAttachment_(message);
  if (!attachment) {
    Logger.log('Kein PDF-Anhang gefunden');
    return;
  }

  var data = extractLieferscheinData_(attachment.copyBlob());
  Logger.log('Lieferschein Nr: ' + extractLieferscheinNr_(data.rawText));
  Logger.log('Tabellen-Position: ' + JSON.stringify(data.tablePosition));
  Logger.log('Text-Position: ' + JSON.stringify(extractFirstPositionFromText_(data.rawText)));
  Logger.log('--- RAW TEXT ---');
  Logger.log(data.rawText);
}
