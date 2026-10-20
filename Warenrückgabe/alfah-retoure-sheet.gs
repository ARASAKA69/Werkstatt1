var ALFAH_RETOURE_TAB = 'Alfah';

function alfahRetoureJson_(value) {
  return ContentService.createTextOutput(JSON.stringify(value || {}))
    .setMimeType(ContentService.MimeType.JSON);
}

function alfahRetoureSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Skript aus der Tabelle Warehouse Returns Spare Parts heraus anlegen.');
  var sheets = ss.getSheets();
  var i;
  for (i = 0; i < sheets.length; i++) {
    if (String(sheets[i].getName()).toLowerCase() === ALFAH_RETOURE_TAB.toLowerCase()) return sheets[i];
  }
  throw new Error('Tabellenblatt Alfah fehlt. Vorhanden: ' + sheets.map(function (item) { return item.getName(); }).join(', '));
}

function alfahRetoureHeader_(sheet) {
  var lastCol = Math.max(10, sheet.getLastColumn());
  var scanRows = Math.min(6, Math.max(2, sheet.getLastRow()));
  var scan = sheet.getRange(1, 1, scanRows, lastCol).getDisplayValues();
  var r;
  for (r = 0; r < scan.length; r++) {
    var map = {};
    var c;
    for (c = 0; c < scan[r].length; c++) {
      var name = String(scan[r][c] || '').replace(/\s+/g, ' ').trim().toLowerCase();
      if (name) map[name] = c + 1;
    }
    if (map['r-nummer'] && map['auftrag']) return { row: r + 1, cols: map, width: Math.max(lastCol, scan[r].length) };
  }
  throw new Error('Kopfzeile im Blatt Alfah nicht gefunden.');
}

function alfahRetoureFind_(sheet, header, rNumber, article) {
  var rCol = header.cols['r-nummer'];
  var aCol = header.cols['teilenummer'];
  var last = sheet.getLastRow();
  if (!rCol || last <= header.row) return 0;
  var count = last - header.row;
  var rValues = sheet.getRange(header.row + 1, rCol, count, 1).getDisplayValues();
  var aValues = aCol ? sheet.getRange(header.row + 1, aCol, count, 1).getDisplayValues() : [];
  var wantedR = String(rNumber || '').trim();
  var wantedArticle = String(article || '').trim();
  var i;
  for (i = 0; i < rValues.length; i++) {
    if (String(rValues[i][0] || '').trim() !== wantedR) continue;
    if (!wantedArticle || !aCol || String((aValues[i] || [])[0] || '').trim() === wantedArticle) return header.row + 1 + i;
  }
  return 0;
}

function alfahRetourePut_(values, cols, name, value) {
  var col = cols[name];
  if (col) values[col - 1] = value;
}

function doGet(e) {
  try {
    var params = (e && e.parameter) || {};
    var sheet = alfahRetoureSheet_();
    var header = alfahRetoureHeader_(sheet);
    var rNumber = String(params.r || params.case || params.vorgang || '').trim();
    if (!rNumber) return alfahRetoureJson_({ ok: true, ping: true, sheet: sheet.getName() });
    var article = String(params.article || params.artikel || params.teilenummer || '').trim();
    var existing = alfahRetoureFind_(sheet, header, rNumber, article);
    if (existing) return alfahRetoureJson_({ ok: true, duplicate: true, row: existing, sheet: sheet.getName() });
    var now = new Date();
    var zone = Session.getScriptTimeZone() || 'Europe/Berlin';
    var values = [];
    var c;
    for (c = 0; c < header.width; c++) values.push('');
    alfahRetourePut_(values, header.cols, 'auftrag', String(params.order || params.bestellung || params.auftrag || '').trim());
    alfahRetourePut_(values, header.cols, 'stock-id', String(params.stock || params.stockid || '').trim());
    alfahRetourePut_(values, header.cols, 'r-nummer', rNumber);
    alfahRetourePut_(values, header.cols, 'teilenummer', article);
    alfahRetourePut_(values, header.cols, 'name', String(params.name || params.bezeichnung || '').trim());
    alfahRetourePut_(values, header.cols, 'lieferant', String(params.supplier || params.lieferant || '').trim());
    alfahRetourePut_(values, header.cols, 'hersteller', String(params.brand || params.hersteller || '').trim());
    alfahRetourePut_(values, header.cols, 'datum', Utilities.formatDate(now, zone, 'dd.MM.yyyy'));
    alfahRetourePut_(values, header.cols, 'uhrzeit', Utilities.formatDate(now, zone, 'HH:mm'));
    alfahRetourePut_(values, header.cols, 'wer', String(params.who || params.wer || 'Berger').trim());
    sheet.appendRow(values);
    SpreadsheetApp.flush();
    return alfahRetoureJson_({ ok: true, row: sheet.getLastRow(), sheet: sheet.getName() });
  } catch (error) {
    return alfahRetoureJson_({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}
