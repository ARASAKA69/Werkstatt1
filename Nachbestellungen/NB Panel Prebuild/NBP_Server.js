function doGet() {
  if (!nbpAllowed_()) {
    return HtmlService.createHtmlOutput("<p style=\"font-family:sans-serif;padding:24px\">Kein Zugriff.</p>").setTitle("NB Panel");
  }
  var t = HtmlService.createTemplateFromFile("NBP_App");
  t.appVersion = NBP_VERSION;
  return t
    .evaluate()
    .setTitle("NB Panel")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function nbpUser_() {
  try {
    return String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  } catch (e) {
    return "";
  }
}

function nbpAllowed_() {
  var u = nbpUser_();
  if (!u) return false;
  for (var i = 0; i < NBP_ALLOWED.length; i++) {
    if (String(NBP_ALLOWED[i]).toLowerCase() === u) return true;
  }
  return false;
}

function nbpSS_() {
  return SpreadsheetApp.openById(NBP_MAIN_ID);
}

function nbpNorm_(s) {
  return String(s == null ? "" : s)
    .replace(/\u00a0/g, " ")
    .replace(/[\u200b-\u200d\ufeff]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function nbpIsDate_(v) {
  return Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime());
}

function nbpToMs_(v) {
  if (nbpIsDate_(v)) return v.getTime();
  if (typeof v === "number" && v > 20000 && v < 80000) return Math.round((v - 25569) * 864e5);
  var s = String(v == null ? "" : v).trim();
  var m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    var y = Number(m[3]);
    if (y < 100) y += 2000;
    return new Date(y, Number(m[2]) - 1, Number(m[1]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0)).getTime();
  }
  return 0;
}

function nbpDayIndex_(ms) {
  if (!ms) return null;
  var k = Utilities.formatDate(new Date(ms), NBP_TZ, "yyyy-MM-dd").split("-");
  return Math.floor(Date.UTC(Number(k[0]), Number(k[1]) - 1, Number(k[2])) / 864e5);
}

function nbpDaysSince_(ms, todayIdx) {
  var d = nbpDayIndex_(ms);
  return d == null ? null : todayIdx - d;
}

function nbpHeaderMatches_(h, alias) {
  return h === alias || h.indexOf(alias + " ") === 0 || h.indexOf(alias + "(") === 0;
}

function nbpMapFields_(headerRow, fields) {
  var cols = {};
  var used = {};
  var score = 0;
  for (var f = 0; f < fields.length; f++) {
    var fd = fields[f];
    for (var c = 0; c < headerRow.length; c++) {
      if (used[c]) continue;
      var h = nbpNorm_(headerRow[c]);
      if (!h) continue;
      var hit = false;
      for (var a = 0; a < fd.aliases.length; a++) {
        if (nbpHeaderMatches_(h, fd.aliases[a])) {
          hit = true;
          break;
        }
      }
      if (hit) {
        cols[fd.key] = c + 1;
        used[c] = true;
        score++;
        break;
      }
    }
  }
  return { cols: cols, score: score };
}

function nbpListFromValidation_(dv) {
  if (!dv) return null;
  var out = [];
  var seen = {};
  function add(v) {
    var s = String(v == null ? "" : v).trim();
    if (!s || seen[s.toLowerCase()]) return;
    seen[s.toLowerCase()] = true;
    out.push(s);
  }
  try {
    var type = dv.getCriteriaType();
    var cv = dv.getCriteriaValues();
    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      var arr = cv[0];
      if (Object.prototype.toString.call(arr) === "[object Array]") {
        for (var i = 0; i < arr.length; i++) add(arr[i]);
      } else {
        add(arr);
      }
    } else if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      var vals = cv[0].getDisplayValues();
      for (var r = 0; r < vals.length; r++) {
        for (var c = 0; c < vals[r].length; c++) add(vals[r][c]);
      }
    }
  } catch (e) {}
  return out.length ? out : null;
}

function nbpColumnOptions_(sheet, col, dataStart) {
  if (!sheet || !col) return null;
  var last = sheet.getLastRow();
  if (last < dataStart) return null;
  var n = Math.min(80, last - dataStart + 1);
  var dvs = sheet.getRange(dataStart, col, n, 1).getDataValidations();
  for (var i = dvs.length - 1; i >= 0; i--) {
    var list = nbpListFromValidation_(dvs[i][0]);
    if (list) return list;
  }
  return null;
}

function nbpNbLayout_(sheet) {
  var lastCol = Math.max(16, sheet.getLastColumn());
  var lastRow = Math.max(1, sheet.getLastRow());
  var scan = sheet.getRange(1, 1, Math.min(15, lastRow), lastCol).getDisplayValues();
  var best = null;
  for (var r = 0; r < scan.length; r++) {
    var m = nbpMapFields_(scan[r], NBP_NB_FIELDS);
    if (!m.cols.stockId || !m.cols.status) continue;
    if (!best || m.score > best.score) {
      best = { headerRow: r + 1, cols: m.cols, score: m.score, headers: scan[r] };
    }
  }
  if (!best) {
    var hr = Math.min(3, lastRow);
    best = { headerRow: hr, cols: {}, score: 0, headers: sheet.getRange(hr, 1, 1, lastCol).getDisplayValues()[0] };
  }
  var usedCols = {};
  for (var k in best.cols) usedCols[best.cols[k]] = true;
  for (var f = 0; f < NBP_NB_FIELDS.length; f++) {
    var fd = NBP_NB_FIELDS[f];
    if (best.cols[fd.key]) continue;
    if (usedCols[fd.col]) continue;
    best.cols[fd.key] = fd.col;
    usedCols[fd.col] = true;
  }
  best.dataStart = best.headerRow + 1;
  best.lastCol = lastCol;
  return best;
}

function nbpInputLayout_(sheet) {
  var lastCol = Math.max(7, sheet.getLastColumn());
  var lastRow = Math.max(1, sheet.getLastRow());
  var scan = sheet.getRange(1, 1, Math.min(8, lastRow), lastCol).getDisplayValues();
  var hdr = -1;
  var eidCol = 0;
  for (var r = 0; r < scan.length; r++) {
    var hasStock = false;
    var hasGrund = false;
    for (var c = 0; c < scan[r].length; c++) {
      var h = nbpNorm_(scan[r][c]);
      if (h === "stock id") hasStock = true;
      if (h === "grund") hasGrund = true;
      if (!eidCol && (h === "entryid" || h === "entry id" || h === "entry_id")) eidCol = c + 1;
    }
    if (hdr < 0 && hasStock && hasGrund) hdr = r + 1;
  }
  if (hdr < 0) hdr = 2;
  if (!eidCol) eidCol = NBP_INPUT_EID_FALLBACK;
  var width = Math.max(lastCol, eidCol);
  var headers = sheet.getRange(hdr, 1, 1, width).getDisplayValues()[0];
  var hasBem = /bemerkung/i.test(String(headers[6] || ""));
  var extras = [];
  for (var col = hasBem ? 8 : 7; col <= width; col++) {
    if (col === eidCol) continue;
    var label = String(headers[col - 1] || "").replace(/\s+/g, " ").trim();
    if (!label) continue;
    extras.push({ col: col, label: label, editable: /claim|liefert/i.test(label) });
  }
  return {
    headerRow: hdr,
    dataStart: hdr + 1,
    eidCol: eidCol,
    width: width,
    headers: headers,
    hasBem: hasBem,
    extras: extras
  };
}

function nbpReadNb_(ss, todayIdx) {
  var sh = ss.getSheetByName(NBP_NB_TAB);
  if (!sh) throw new Error("Tabellenblatt " + NBP_NB_TAB + " fehlt.");
  var lo = nbpNbLayout_(sh);
  var keyByCol = {};
  for (var k in lo.cols) keyByCol[lo.cols[k]] = k;
  var nbcCols = [];
  var extraCols = [];
  for (var c = 1; c <= lo.lastCol; c++) {
    if (keyByCol[c]) continue;
    var h = String(lo.headers[c - 1] || "").replace(/\s+/g, " ").trim();
    if (!h) continue;
    if (/^nbc/i.test(h)) nbcCols.push({ col: c, label: h });
    else extraCols.push({ col: c, label: h });
  }
  var rows = [];
  var last = sh.getLastRow();
  if (last >= lo.dataStart) {
    var n = last - lo.dataStart + 1;
    var rng = sh.getRange(lo.dataStart, 1, n, lo.lastCol);
    var vals = rng.getValues();
    var disp = rng.getDisplayValues();
    for (var i = 0; i < n; i++) {
      var stock = String(disp[i][lo.cols.stockId - 1] || "").replace(/\s+/g, "").toUpperCase();
      if (!stock) continue;
      var rec = { row: lo.dataStart + i, f: {}, nbc: [], extra: [] };
      for (var f = 0; f < NBP_NB_FIELDS.length; f++) {
        var key = NBP_NB_FIELDS[f].key;
        var col = lo.cols[key];
        rec.f[key] = col ? String(disp[i][col - 1] || "").trim() : "";
      }
      rec.f.stockId = stock;
      rec.entryId = rec.f.entryId;
      rec.datumMs = nbpToMs_(vals[i][lo.cols.datum - 1]);
      rec.statusMs = nbpToMs_(vals[i][lo.cols.statusAenderung - 1]);
      rec.daysOpen = nbpDaysSince_(rec.datumMs, todayIdx);
      rec.daysStatus = nbpDaysSince_(rec.statusMs, todayIdx);
      rec.leadDays = rec.datumMs && rec.statusMs ? nbpDayIndex_(rec.statusMs) - nbpDayIndex_(rec.datumMs) : null;
      for (var a = 0; a < nbcCols.length; a++) {
        rec.nbc.push(String(disp[i][nbcCols[a].col - 1] || "").trim());
      }
      for (var b = 0; b < extraCols.length; b++) {
        rec.extra.push(String(disp[i][extraCols[b].col - 1] || "").trim());
      }
      rows.push(rec);
    }
  }
  var options = {};
  for (var o = 0; o < NBP_NB_FIELDS.length; o++) {
    var fd = NBP_NB_FIELDS[o];
    if (fd.edit !== "list") continue;
    options[fd.key] = nbpColumnOptions_(sh, lo.cols[fd.key], lo.dataStart);
  }
  return {
    rows: rows,
    options: options,
    nbcLabels: nbcCols.map(function (x) { return x.label; }),
    extraLabels: extraCols.map(function (x) { return x.label; })
  };
}

function nbpReadDashboard_(ss) {
  var map = {};
  var sh = ss.getSheetByName(NBP_DASH_TAB);
  if (!sh) return map;
  var lastCol = Math.max(15, sh.getLastColumn());
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return map;
  var scan = sh.getRange(1, 1, Math.min(15, lastRow), lastCol).getDisplayValues();
  var best = null;
  for (var r = 0; r < scan.length; r++) {
    var m = nbpMapFields_(scan[r], NBP_DASH_FIELDS);
    if (!m.cols.entryId || !m.cols.stockId) continue;
    if (!best || m.score > best.score) best = { headerRow: r + 1, cols: m.cols, score: m.score };
  }
  if (!best) return map;
  var start = best.headerRow + 1;
  if (lastRow < start) return map;
  var disp = sh.getRange(start, 1, lastRow - start + 1, lastCol).getDisplayValues();
  for (var i = 0; i < disp.length; i++) {
    var eid = String(disp[i][best.cols.entryId - 1] || "").trim();
    if (!eid) continue;
    var rec = {};
    for (var f = 0; f < NBP_DASH_FIELDS.length; f++) {
      var key = NBP_DASH_FIELDS[f].key;
      var col = best.cols[key];
      rec[key] = col ? String(disp[i][col - 1] || "").trim() : "";
    }
    map[eid] = rec;
  }
  return map;
}

function nbpReadSection_(ss, sec, todayIdx) {
  var base = { key: sec.key, label: sec.label, sheet: sec.sheet, code: sec.code, rows: [] };
  var sh = ss.getSheetByName(sec.sheet);
  if (!sh) {
    base.missing = true;
    return base;
  }
  var lay = nbpInputLayout_(sh);
  var last = sh.getLastRow();
  if (last >= lay.dataStart) {
    var n = last - lay.dataStart + 1;
    var rng = sh.getRange(lay.dataStart, 1, n, lay.width);
    var vals = rng.getValues();
    var disp = rng.getDisplayValues();
    for (var i = 0; i < n; i++) {
      var stock = String(disp[i][1] || "").replace(/\s+/g, "").toUpperCase();
      if (!stock) continue;
      var ms = nbpToMs_(vals[i][0]);
      var extra = [];
      for (var x = 0; x < lay.extras.length; x++) {
        extra.push(String(disp[i][lay.extras[x].col - 1] || "").trim());
      }
      base.rows.push({
        row: lay.dataStart + i,
        entryId: String(disp[i][lay.eidCol - 1] || "").trim(),
        datumMs: ms,
        datum: String(disp[i][0] || "").trim(),
        days: nbpDaysSince_(ms, todayIdx),
        stockId: stock,
        et: String(disp[i][2] || "").trim(),
        grund: String(disp[i][3] || "").trim(),
        art: String(disp[i][4] || "").trim(),
        bearbeiter: String(disp[i][5] || "").trim(),
        bemerkung: lay.hasBem ? String(disp[i][6] || "").trim() : "",
        extra: extra
      });
    }
  }
  var extraOptions = {};
  for (var e = 0; e < lay.extras.length; e++) {
    if (!lay.extras[e].editable) continue;
    extraOptions[lay.extras[e].col] = nbpColumnOptions_(sh, lay.extras[e].col, lay.dataStart);
  }
  base.layout = { hasBem: lay.hasBem, extras: lay.extras };
  base.options = {
    et: nbpColumnOptions_(sh, 3, lay.dataStart),
    art: nbpColumnOptions_(sh, 5, lay.dataStart),
    bearbeiter: nbpColumnOptions_(sh, 6, lay.dataStart),
    extra: extraOptions
  };
  return base;
}

function getPanelData() {
  if (!nbpAllowed_()) return { ok: false, message: "Kein Zugriff." };
  try {
    var ss = nbpSS_();
    var now = new Date();
    var todayIdx = nbpDayIndex_(now.getTime());
    var nb = nbpReadNb_(ss, todayIdx);
    var dash = nbpReadDashboard_(ss);
    for (var i = 0; i < nb.rows.length; i++) {
      var eid = nb.rows[i].entryId;
      nb.rows[i].dash = eid && dash[eid] ? dash[eid] : null;
    }
    var sections = [];
    for (var s = 0; s < NBP_SECTIONS.length; s++) {
      sections.push(nbpReadSection_(ss, NBP_SECTIONS[s], todayIdx));
    }
    return {
      ok: true,
      version: NBP_VERSION,
      user: nbpUser_(),
      loadedAt: Utilities.formatDate(now, NBP_TZ, "dd.MM.yyyy HH:mm"),
      loadedClock: Utilities.formatDate(now, NBP_TZ, "HH:mm"),
      sheetUrl: ss.getUrl(),
      nb: nb,
      sections: sections
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}
