var CONFIG = {
  senderEmail: 'info@rv-seng.de',
  recipientEmail: 'Info@rv-seng.de',
  ccEmails: [
    'lena.scholz@auto1.com',
    'joerg.eichenseher@auto1.com',
    'bastian.frankhauser@auto1.com',
    'francesco.berger@auto1.com',
    'aimad.eddidi@auto1.com',
    'robert.gawron@auto1.com',
    'emir.jasarevic@auto1.com',
    'g.bell@rv-seng.de'
  ],
  stockIdCol: 1,
  lieferscheinNrCol: 2,
  artikelCol: 3,
  beschreibungCol: 4,
  anzahlCol: 5,
  statusCol: 6,
  headerRow: 1
};

var LOGO_BASE64 = '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';

function getLogoBlob_() {
  return Utilities.newBlob(Utilities.base64Decode(LOGO_BASE64), 'image/png', 'autohero-logo.png');
}

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

  if (!CONFIG.recipientEmail || CONFIG.recipientEmail.indexOf('@') === -1) {
    ui.alert('Bitte zuerst CONFIG.recipientEmail im Script eintragen.');
    return;
  }

  var pdfBlob = buildFilledRowsPdf_(sheet, rows);
  var subject = 'Reifen Seng Retoure - ' + sheet.getName();
  var signature = 'Mit Freundlichen Gruessen\nAutoHero GmbH\nNordhausstr. 1\n93155 Hemau';
  var body = 'Anbei die Retourenliste als PDF.\n\n' +
    'Bei Rueckfragen wenden Sie sich bitte an ersatzteile.hemau@autohero.com.\n\n' +
    signature;
  var htmlBody =
    '<div style="text-align:center;"><img src="cid:logo" width="220"></div>' +
    '<p>Anbei die Retourenliste als PDF.</p>' +
    '<p>Bei Rueckfragen wenden Sie sich bitte an ersatzteile.hemau@autohero.com.</p>' +
    '<p>Mit Freundlichen Gruessen<br>AutoHero GmbH<br>Nordhausstr. 1<br>93155 Hemau</p>';

  GmailApp.sendEmail(CONFIG.recipientEmail, subject, body, {
    cc: (CONFIG.ccEmails || []).join(','),
    htmlBody: htmlBody,
    inlineImages: { logo: getLogoBlob_() },
    attachments: [pdfBlob],
    name: 'Reifen Seng Retoure'
  });

  ui.alert('E-Mail mit ' + rows.length + ' Position(en) an ' + CONFIG.recipientEmail +
    ' (CC: ' + CONFIG.ccEmails.join(', ') + ') gesendet.');
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
    var tableWidth = 750;
    var logoWidth = 220;
    var logoHeight = Math.round(logoWidth * 59 / 475);

    temp.setRowHeight(1, logoHeight + 16);
    temp.insertImage(getLogoBlob_(), 1, 1, Math.round((tableWidth - logoWidth) / 2), 8)
      .setWidth(logoWidth)
      .setHeight(logoHeight);

    var titleRange = temp.getRange(2, 1, 1, 5);
    titleRange.merge();
    titleRange.setValue('Reifen Retoure (' + Utilities.formatDate(new Date(), 'Europe/Berlin', 'dd.MM.yyyy') + ')');
    titleRange.setFontWeight('bold');
    titleRange.setFontSize(14);
    titleRange.setHorizontalAlignment('center');

    var header = [['StockID', 'Lieferschein Nr.', 'Artikel/Reifen', 'Beschreibung', 'Anzahl']];
    temp.getRange(3, 1, 1, 5).setValues(header).setFontWeight('bold');
    temp.getRange(4, 1, rows.length, 5).setValues(rows);

    var total = 0;
    for (var i = 0; i < rows.length; i++) {
      total += Number(rows[i][4]) || 0;
    }
    var totalRow = rows.length + 4;
    temp.getRange(totalRow, 4).setValue('TOTAL:').setFontWeight('bold').setHorizontalAlignment('right');
    temp.getRange(totalRow, 5).setValue(total).setFontWeight('bold');

    var tableRange = temp.getRange(3, 1, totalRow - 2, 5);
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
  var senderLower = CONFIG.senderEmail.toLowerCase();
  var bestSenderMatch = null;
  var bestSubjectMatch = null;

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();
    for (var m = 0; m < messages.length; m++) {
      var msg = messages[m];
      var fromLower = String(msg.getFrom() || '').toLowerCase();
      if (fromLower.indexOf(senderLower) === -1) continue;
      if (!findPdfAttachment_(msg)) continue;

      if (!bestSenderMatch || msg.getDate().getTime() > bestSenderMatch.getDate().getTime()) {
        bestSenderMatch = msg;
      }

      var subject = String(msg.getSubject() || '').toUpperCase();
      if (subject.indexOf(stockIdUpper) === -1) continue;
      if (!bestSubjectMatch || msg.getDate().getTime() > bestSubjectMatch.getDate().getTime()) {
        bestSubjectMatch = msg;
      }
    }
  }

  return bestSubjectMatch || bestSenderMatch;
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
  Logger.log('Autorisiertes Konto (effective user): ' + Session.getEffectiveUser().getEmail());
  Logger.log('Aktives Konto (active user): ' + Session.getActiveUser().getEmail());

  var labelCheck = GmailApp.search('label:"Reifen Seng"', 0, 3);
  Logger.log('Treffer mit label:"Reifen Seng": ' + labelCheck.length + ' (sollte > 0 sein, wenn dies das richtige Postfach ist)');

  var broad = GmailApp.search('from:' + CONFIG.senderEmail, 0, 5);
  Logger.log('Treffer nur nach Sender (' + CONFIG.senderEmail + '): ' + broad.length);
  for (var i = 0; i < broad.length; i++) {
    var msgs = broad[i].getMessages();
    var last = msgs[msgs.length - 1];
    Logger.log('Thread ' + i + ': ' + msgs.length + ' Nachricht(en), letzte: "' + last.getSubject() + '" (' + last.getDate() + ')');
  }
}

function debugStockSearch_(stockId) {
  Logger.log('=== Diagnose fuer ' + stockId + ' ===');

  var idOnly = GmailApp.search('"' + stockId + '" has:attachment', 0, 5);
  Logger.log('Nur StockID + has:attachment (ohne Sender-Filter): ' + idOnly.length + ' Thread(s)');
  for (var i = 0; i < idOnly.length; i++) {
    var msgs = idOnly[i].getMessages();
    for (var m = 0; m < msgs.length; m++) {
      Logger.log('  From: "' + msgs[m].getFrom() + '" | Subject: "' + msgs[m].getSubject() + '" | Datum: ' + msgs[m].getDate());
    }
  }

  var withSender = GmailApp.search('from:(' + CONFIG.senderEmail + ') "' + stockId + '"', 0, 5);
  Logger.log('Sender(' + CONFIG.senderEmail + ') + StockID, ohne attachment-Filter: ' + withSender.length + ' Thread(s)');

  var full = GmailApp.search('from:(' + CONFIG.senderEmail + ') "' + stockId + '" has:attachment', 0, 5);
  Logger.log('Sender + StockID + has:attachment: ' + full.length + ' Thread(s)');
}

function testExtractionAM11142() {
  testExtraction('AM11142');
}

function debugStockSearchXU27308() {
  debugStockSearch_('XU27308');
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
