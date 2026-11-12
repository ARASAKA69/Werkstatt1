function rtpWithLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(RTP_LOCK_MS)) throw new Error("Sheet ist gerade belegt, bitte gleich nochmal versuchen.");
  try {
    return fn();
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function rtpRun_(fn) {
  if (!rtpAllowed_()) return { ok: false, message: "Kein Zugriff." };
  try {
    var res = rtpWithLock_(fn) || {};
    res.ok = true;
    return res;
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function rtpText_(value) {
  var text = String(value == null ? "" : value).replace(/\s+/g, " ").trim();
  if (text.length > 180) text = text.substring(0, 180);
  if (text.charAt(0) === "=") text = "'" + text;
  return text;
}

function rtpCanonR_(value) {
  var s = rtpText_(value).toUpperCase();
  if (!s) return "";
  if (/^\d{4,7}$/.test(s)) return "R-" + s;
  var m = s.match(/R-\d{4,}/);
  return m ? m[0] : s;
}

function rtpStampNow_() {
  var text = Utilities.formatDate(new Date(), RTP_TZ, "dd.MM.yyyy HH:mm");
  var parts = text.split(" ");
  return { date: parts[0] || "", time: parts[1] || "" };
}

function rtpNormPart_(s) {
  return String(s == null ? "" : s).toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function rtpRowMatchesIdent_(dispRow, cfg, ident) {
  if (!ident) return true;
  if (cfg.kind === "alfah") {
    var rRow = String(dispRow[2] || "").trim().toUpperCase();
    var pRow = rtpNormPart_(dispRow[3]);
    var sRow = String(dispRow[1] || "").trim().toUpperCase();
    var oRow = String(dispRow[0] || "").trim();
    var rWant = String(ident.rNumber || "").trim().toUpperCase();
    var pWant = rtpNormPart_(ident.article);
    var sWant = String(ident.stockId || "").trim().toUpperCase();
    var oWant = String(ident.orderId || "").trim();
    if (rWant && rRow !== rWant) return false;
    if (pWant && pRow !== pWant) return false;
    if (!rWant && sWant && sRow !== sWant) return false;
    if (!rWant && !pWant && oWant && oRow !== oWant) return false;
    return !!(rRow || pRow || sRow || oRow);
  }
  var sSup = String(dispRow[1] || "").trim().toUpperCase();
  var pSup = rtpNormPart_(dispRow[2]);
  var oSup = String(dispRow[0] || "").trim();
  var nSup = rtpNorm_(dispRow[3]);
  var sWant2 = String(ident.stockId || "").trim().toUpperCase();
  var pWant2 = rtpNormPart_(ident.article);
  var oWant2 = String(ident.orderId || "").trim();
  var nWant2 = rtpNorm_(ident.name);
  if (sWant2 && sSup !== sWant2) return false;
  if (pWant2 && pSup !== pWant2) return false;
  if (!sWant2 && !pWant2 && oWant2 && oSup !== oWant2) return false;
  if (!sWant2 && !pWant2 && nWant2 && nSup !== nWant2) return false;
  return !!(sSup || pSup || oSup || nSup);
}

function rtpLocateRow_(sheet, cfg, rowHint, ident) {
  var last = sheet.getLastRow();
  if (last < 3) throw new Error("Tabellenblatt ist leer.");
  var cols = cfg.kind === "alfah" ? 5 : 4;
  var hint = Number(rowHint || 0);
  if (hint >= 3 && hint <= last) {
    var single = sheet.getRange(hint, 1, 1, cols).getDisplayValues()[0];
    if (rtpRowMatchesIdent_(single, cfg, ident)) return hint;
  }
  if (ident) {
    var all = sheet.getRange(3, 1, last - 2, cols).getDisplayValues();
    for (var i = 0; i < all.length; i++) {
      if (rtpRowMatchesIdent_(all[i], cfg, ident)) return 3 + i;
    }
  }
  throw new Error("Eintrag nicht mehr in dieser Zeile gefunden (evtl. bereits verschoben). Daten werden aktualisiert.");
}

function rtpPutLink_(cell, url, label, color) {
  if (!url) {
    cell.clearContent();
    return;
  }
  var style = SpreadsheetApp.newTextStyle()
    .setForegroundColor(color || RTP_ALFAH_BLUE)
    .setUnderline(true)
    .setFontFamily("Arial")
    .setFontSize(RTP_FONT)
    .build();
  var rich = SpreadsheetApp.newRichTextValue()
    .setText(label)
    .setLinkUrl(url)
    .setTextStyle(style)
    .build();
  cell.setRichTextValue(rich);
}

function rtpPutAlfahChecks_(sheet, rowNumber) {
  var boxes = sheet.getRange(rowNumber, 15, 1, 2);
  boxes.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  boxes.setHorizontalAlignment("center");
  boxes.setVerticalAlignment("middle");
  boxes.setFontSize(RTP_FONT);
}

function rtpPaintSupplierRow_(sheet, rowNumber, accent) {
  var rowRange = sheet.getRange(rowNumber, 1, 1, 10);
  rowRange.setBackground("#ffffff");
  rowRange.setFontColor("#1f2933");
  rowRange.setFontFamily("Arial");
  rowRange.setFontSize(RTP_FONT);
  rowRange.setVerticalAlignment("middle");
  rowRange.setWrap(true);
  rowRange.setBorder(true, true, true, true, true, true, "#d6deea", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(rowNumber, 2).setFontWeight("bold");
  sheet.getRange(rowNumber, 3).setFontWeight("bold").setFontColor(accent || RTP_WM_GREEN);
  try { sheet.autoResizeRows(rowNumber, 1); } catch (e) {}
}

function rtpStatePaintSpec_(stateKey) {
  if (stateKey === "bereit" || stateKey === "geliefert") return { text: "Rücksendung", bg: RTP_ALFAH_GREEN, fg: "#ffffff", defaultInfo: "Alfah System" };
  if (stateKey === "veranlasst") return { text: "Rücksendung", bg: RTP_ALFAH_ORANGE, fg: "#ffffff", defaultInfo: "Retoure veranlasst" };
  if (stateKey === "vorbereitet_mail") return { text: "VORBEREITET", bg: RTP_ALFAH_YELLOW, fg: "#1f2933", defaultInfo: "" };
  if (stateKey === "abgelehnt") return { text: "ABGELEHNT", bg: RTP_ALFAH_RED, fg: "#ffffff", defaultInfo: "" };
  if (stateKey === "storno") return { text: "STORNO", bg: RTP_ALFAH_RED, fg: "#ffffff", defaultInfo: "" };
  return null;
}

function rtpApplyAlfahState_(sheet, rowNumber, stateKey, infoOpt, keepMail) {
  var rNumber = rtpCanonR_(sheet.getRange(rowNumber, 3).getDisplayValue());
  var curMailRich = sheet.getRange(rowNumber, 13).getRichTextValue();
  var curMailUrl = rtpRichLink_(curMailRich);
  var searchUrl = rNumber ? rtpSearchUrl_(rNumber) : "";
  var rowRange = sheet.getRange(rowNumber, 1, 1, 14);
  var spec = rtpStatePaintSpec_(stateKey);
  if (!spec) {
    sheet.getRange(rowNumber, 11).setValue("");
    sheet.getRange(rowNumber, 12).setValue(infoOpt != null ? rtpText_(infoOpt) : "");
    if (!keepMail) sheet.getRange(rowNumber, 13).clearContent();
    else if (curMailUrl) rtpPutLink_(sheet.getRange(rowNumber, 13), curMailUrl, "Mail", RTP_ALFAH_BLUE);
    if (searchUrl) rtpPutLink_(sheet.getRange(rowNumber, 14), searchUrl, "Suche", RTP_ALFAH_BLUE);
    rowRange.setBackground("#ffffff");
    rowRange.setFontColor("#1f2933");
    rowRange.setFontFamily("Arial");
    rowRange.setFontSize(RTP_FONT);
    rowRange.setWrap(true);
    rowRange.setVerticalAlignment("middle");
    rowRange.setBorder(true, true, true, true, true, true, "#d6deea", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(rowNumber, 2).setFontWeight("bold");
    sheet.getRange(rowNumber, 3).setFontWeight("bold").setFontColor(RTP_ALFAH_BLUE);
    rtpPutAlfahChecks_(sheet, rowNumber);
    return;
  }
  var curInfo = String(sheet.getRange(rowNumber, 12).getDisplayValue() || "").trim();
  var nextInfo = infoOpt != null ? rtpText_(infoOpt) : (curInfo || spec.defaultInfo);
  sheet.getRange(rowNumber, 11, 1, 2).setValues([[spec.text, nextInfo]]);
  if (curMailUrl) rtpPutLink_(sheet.getRange(rowNumber, 13), curMailUrl, "Mail", spec.fg);
  if (searchUrl) rtpPutLink_(sheet.getRange(rowNumber, 14), searchUrl, "Suche", spec.fg);
  rowRange.setBackground(spec.bg);
  rowRange.setFontColor(spec.fg);
  rowRange.setFontFamily("Arial");
  rowRange.setFontSize(RTP_FONT);
  rowRange.setWrap(true);
  rowRange.setVerticalAlignment("middle");
  rowRange.setBorder(true, true, true, true, true, true, "#d6deea", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(rowNumber, 2).setFontWeight("bold");
  sheet.getRange(rowNumber, 3).setFontWeight("bold");
  sheet.getRange(rowNumber, 11).setFontWeight("bold");
  rtpPutAlfahChecks_(sheet, rowNumber);
}

function rtpRowBlank_(sheet, rowNumber, width) {
  var last = sheet.getLastRow();
  if (rowNumber < 1 || rowNumber > last) return false;
  var vals = sheet.getRange(rowNumber, 1, 1, width || 16).getDisplayValues()[0];
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i] || "").trim()) return false;
  }
  return true;
}

function rtpEnsureAlfahTab_(ss, cfg) {
  var sh = rtpFindSheet_(ss, cfg);
  if (!sh) sh = ss.insertSheet(cfg.sheet);
  sh.setTabColor(cfg.color);
  if (!String(sh.getRange(1, 1).getValue() || "")) {
    var titleRange = sh.getRange(1, 1, 1, 16);
    if (titleRange.isPartOfMerge()) titleRange.breakApart();
    titleRange.merge();
    titleRange
      .setValue(cfg.title)
      .setBackground(cfg.color)
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily("Arial")
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle");
    var head = sh.getRange(2, 1, 1, 16);
    head.setValues([RTP_ALFAH_HEADERS]);
    head.setBackground(cfg.color);
    head.setFontColor("#ffffff");
    head.setFontWeight("bold");
    head.setFontSize(RTP_FONT);
    head.setFontFamily("Arial");
    head.setHorizontalAlignment("center");
    head.setVerticalAlignment("middle");
    head.setWrap(true);
    sh.setFrozenRows(2);
    sh.setRowHeight(1, 22);
    sh.setRowHeight(2, 32);
    for (var c = 0; c < RTP_ALFAH_WIDTHS.length; c++) sh.setColumnWidth(c + 1, RTP_ALFAH_WIDTHS[c]);
  }
  return sh;
}

function rtpEnsureSupplierTab_(ss, cfg) {
  var sh = rtpFindSheet_(ss, cfg);
  if (!sh) sh = ss.insertSheet(cfg.sheet);
  sh.setTabColor(cfg.color);
  if (!String(sh.getRange(1, 1).getValue() || "")) {
    var titleRange = sh.getRange(1, 1, 1, 10);
    if (titleRange.isPartOfMerge()) titleRange.breakApart();
    titleRange.merge();
    titleRange
      .setValue(cfg.title)
      .setBackground(cfg.color)
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily("Arial")
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle");
    var head = sh.getRange(2, 1, 1, 10);
    head.setValues([RTP_SUPPLIER_HEADERS]);
    head.setBackground(cfg.color);
    head.setFontColor("#ffffff");
    head.setFontWeight("bold");
    head.setFontSize(RTP_FONT);
    head.setFontFamily("Arial");
    head.setHorizontalAlignment("center");
    head.setVerticalAlignment("middle");
    head.setWrap(true);
    sh.setFrozenRows(2);
    sh.setRowHeight(1, 22);
    var widths = cfg.widths || RTP_SUPPLIER_WIDTHS;
    for (var c = 0; c < widths.length; c++) sh.setColumnWidth(c + 1, widths[c]);
  }
  return sh;
}

function rtpMoveAlfahRows_(sourceSheet, rows, destSheet, forceBg) {
  if (!rows || !rows.length) return { moved: 0, removed: 0 };
  var sorted = rows.slice().sort(function (a, b) { return a - b; });
  var srcLast = sourceSheet.getLastRow();
  var pairs = srcLast >= 3 ? sourceSheet.getRange(3, 3, srcLast - 2, 2).getValues() : [];
  var have = {};
  var destLast = destSheet.getLastRow();
  if (destLast >= 3) {
    var existing = destSheet.getRange(3, 3, destLast - 2, 2).getValues();
    for (var i = 0; i < existing.length; i++) {
      var known = String(existing[i][0] || "").toUpperCase() + "|" + String(existing[i][1] || "").toUpperCase();
      if (known !== "|") have[known] = true;
    }
  }
  var cursor = 3;
  if (destLast >= 3) {
    var tail = destSheet.getRange(destLast, 1, 1, 4).getDisplayValues()[0].join("");
    cursor = String(tail).trim() ? destLast + 2 : destLast + 1;
  }
  var moved = 0;
  for (var m = 0; m < sorted.length; m++) {
    var rowNumber = sorted[m];
    var pairRow = pairs[rowNumber - 3] || ["", ""];
    var key = String(pairRow[0] || "").toUpperCase() + "|" + String(pairRow[1] || "").toUpperCase();
    if (key !== "|" && have[key]) continue;
    sourceSheet.getRange(rowNumber, 1, 1, 16).copyTo(destSheet.getRange(cursor, 1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    if (forceBg) destSheet.getRange(cursor, 1, 1, 16).setBackground(forceBg);
    try { destSheet.setRowHeight(cursor, sourceSheet.getRowHeight(rowNumber)); } catch (e) {}
    if (key !== "|") have[key] = true;
    moved++;
    cursor += 2;
  }
  for (var d = sorted.length - 1; d >= 0; d--) {
    var drop = sorted[d];
    var above = drop <= 3 || rtpRowBlank_(sourceSheet, drop - 1, 16);
    var span = above && rtpRowBlank_(sourceSheet, drop + 1, 16) ? 2 : 1;
    sourceSheet.deleteRows(drop, span);
  }
  return { moved: moved, removed: sorted.length };
}

function rtpIsRueckText_(text) {
  var t = String(text || "").trim();
  return t === "Rücksendung" || t === "Rücksendung eingeleitet";
}

function rtpAutoMoveSingleAlfahRow_(ss, alfahSheet, rowNumber) {
  var status = String(alfahSheet.getRange(rowNumber, 11).getDisplayValue() || "").trim();
  var bg = String(alfahSheet.getRange(rowNumber, 1).getBackground() || "").trim().toLowerCase();
  var statusUp = status.toUpperCase();
  if (statusUp === "ABGELEHNT" || statusUp === "STORNO" || bg === RTP_ALFAH_RED) {
    var rejDest = rtpEnsureAlfahTab_(ss, rtpTabConfig_("abgelehnt"));
    rtpMoveAlfahRows_(alfahSheet, [rowNumber], rejDest, RTP_ALFAH_RED);
    return "abgelehnt";
  }
  var boxes = alfahSheet.getRange(rowNumber, 15, 1, 2).getValues()[0];
  if (rtpIsRueckText_(status) && bg === RTP_ALFAH_GREEN && rtpChecked_(boxes[0]) && rtpChecked_(boxes[1])) {
    var doneDest = rtpEnsureAlfahTab_(ss, rtpTabConfig_("geliefert"));
    rtpMoveAlfahRows_(alfahSheet, [rowNumber], doneDest, "");
    return "geliefert";
  }
  return "";
}

function rtpSweepAlfahMoves_(ss, alfahSheet) {
  var rejRows = [];
  var last = alfahSheet.getLastRow();
  if (last >= 3) {
    var count = last - 2;
    var status = alfahSheet.getRange(3, 11, count, 1).getDisplayValues();
    var bg = alfahSheet.getRange(3, 1, count, 1).getBackgrounds();
    for (var i = 0; i < count; i++) {
      var st = String(status[i][0] || "").trim().toUpperCase();
      var col = String(bg[i][0] || "").trim().toLowerCase();
      if (st === "ABGELEHNT" || st === "STORNO" || col === RTP_ALFAH_RED) rejRows.push(3 + i);
    }
  }
  var rejRes = rejRows.length
    ? rtpMoveAlfahRows_(alfahSheet, rejRows, rtpEnsureAlfahTab_(ss, rtpTabConfig_("abgelehnt")), RTP_ALFAH_RED)
    : { moved: 0 };

  var delRows = [];
  var last2 = alfahSheet.getLastRow();
  if (last2 >= 3) {
    var count2 = last2 - 2;
    var status2 = alfahSheet.getRange(3, 11, count2, 1).getDisplayValues();
    var bg2 = alfahSheet.getRange(3, 1, count2, 1).getBackgrounds();
    var boxes2 = alfahSheet.getRange(3, 15, count2, 2).getValues();
    for (var j = 0; j < count2; j++) {
      if (!rtpIsRueckText_(status2[j][0])) continue;
      if (String(bg2[j][0] || "").trim().toLowerCase() !== RTP_ALFAH_GREEN) continue;
      if (!rtpChecked_(boxes2[j][0]) || !rtpChecked_(boxes2[j][1])) continue;
      delRows.push(3 + j);
    }
  }
  var delRes = delRows.length
    ? rtpMoveAlfahRows_(alfahSheet, delRows, rtpEnsureAlfahTab_(ss, rtpTabConfig_("geliefert")), "")
    : { moved: 0 };

  return { rejectedMoved: rejRes.moved, deliveredMoved: delRes.moved };
}

function rtpToggleCheck(tabKey, row, ident, field, checked) {
  return rtpRun_(function () {
    var cfg = rtpTabConfig_(tabKey);
    if (!cfg) throw new Error("Unbekannter Tab: " + tabKey);
    var ss = rtpSS_();
    var sh = rtpFindSheet_(ss, cfg);
    if (!sh) throw new Error("Tabellenblatt " + cfg.sheet + " fehlt.");
    var targetRow = rtpLocateRow_(sh, cfg, row, ident);
    var nextVal = !!checked;
    if (cfg.kind === "alfah") {
      var col = field === "geliefert" ? 16 : 15;
      rtpPutAlfahChecks_(sh, targetRow);
      sh.getRange(targetRow, col).setValue(nextVal);
      var movedTo = "";
      if (cfg.key === "alfah") {
        movedTo = rtpAutoMoveSingleAlfahRow_(ss, sh, targetRow);
      }
      var bothCheckedHint = false;
      if (cfg.key === "alfah" && !movedTo) {
        var b = sh.getRange(targetRow, 15, 1, 2).getValues()[0];
        if (rtpChecked_(b[0]) && rtpChecked_(b[1])) bothCheckedHint = true;
      }
      return { row: targetRow, checked: nextVal, movedTo: movedTo, bothCheckedNotGreen: bothCheckedHint };
    }
    return { row: targetRow, checked: nextVal, movedTo: "" };
  });
}

function rtpUpdateEntry(tabKey, row, ident, patch) {
  return rtpRun_(function () {
    patch = patch || {};
    var cfg = rtpTabConfig_(tabKey);
    if (!cfg) throw new Error("Unbekannter Tab: " + tabKey);
    var ss = rtpSS_();
    var sh = rtpFindSheet_(ss, cfg);
    if (!sh) throw new Error("Tabellenblatt " + cfg.sheet + " fehlt.");
    var targetRow = rtpLocateRow_(sh, cfg, row, ident);

    if (cfg.kind === "alfah") {
      var textMap = [
        ["orderId", 1],
        ["stockId", 2],
        ["rNumber", 3],
        ["article", 4],
        ["name", 5],
        ["supplier", 6],
        ["brand", 7],
        ["date", 8],
        ["time", 9],
        ["who", 10]
      ];
      for (var i = 0; i < textMap.length; i++) {
        var k = textMap[i][0];
        var col = textMap[i][1];
        if (!Object.prototype.hasOwnProperty.call(patch, k)) continue;
        var val = rtpText_(patch[k]);
        if (k === "stockId") val = val.toUpperCase();
        if (k === "rNumber") val = rtpCanonR_(val);
        var cell = sh.getRange(targetRow, col);
        cell.setNumberFormat("@");
        cell.setValue(val);
      }
      if (Object.prototype.hasOwnProperty.call(patch, "vorbereitet")) {
        rtpPutAlfahChecks_(sh, targetRow);
        sh.getRange(targetRow, 15).setValue(!!patch.vorbereitet);
      }
      if (Object.prototype.hasOwnProperty.call(patch, "geliefert")) {
        rtpPutAlfahChecks_(sh, targetRow);
        sh.getRange(targetRow, 16).setValue(!!patch.geliefert);
      }
      var hasState = Object.prototype.hasOwnProperty.call(patch, "state");
      var hasInfo = Object.prototype.hasOwnProperty.call(patch, "info");
      if (hasState) {
        rtpApplyAlfahState_(sh, targetRow, patch.state, hasInfo ? patch.info : null, true);
      } else {
        if (hasInfo) {
          var infoCell = sh.getRange(targetRow, 12);
          infoCell.setNumberFormat("@");
          infoCell.setValue(rtpText_(patch.info));
        }
        if (Object.prototype.hasOwnProperty.call(patch, "rNumber")) {
          var curBg = String(sh.getRange(targetRow, 1).getBackground() || "").toLowerCase();
          var curStatus = String(sh.getRange(targetRow, 11).getDisplayValue() || "");
          var curState = rtpClassifyAlfah_(curStatus, "", curBg, cfg.key);
          var spec = rtpStatePaintSpec_(curState);
          var rNum = rtpCanonR_(sh.getRange(targetRow, 3).getDisplayValue());
          rtpPutLink_(sh.getRange(targetRow, 14), rNum ? rtpSearchUrl_(rNum) : "", "Suche", spec ? spec.fg : RTP_ALFAH_BLUE);
        }
      }
      var movedTo = "";
      if (cfg.key === "alfah") {
        movedTo = rtpAutoMoveSingleAlfahRow_(ss, sh, targetRow);
      }
      return { row: targetRow, movedTo: movedTo };
    }

    var supMap = [
      ["orderId", 1],
      ["stockId", 2],
      ["article", 3],
      ["name", 4],
      ["supplier", 5],
      ["brand", 6],
      ["date", 7],
      ["time", 8],
      ["who", 9],
      ["info", 10]
    ];
    for (var s = 0; s < supMap.length; s++) {
      var sk = supMap[s][0];
      var scol = supMap[s][1];
      if (!Object.prototype.hasOwnProperty.call(patch, sk)) continue;
      var sval = rtpText_(patch[sk]);
      if (sk === "stockId") sval = sval.toUpperCase();
      var scell = sh.getRange(targetRow, scol);
      scell.setNumberFormat("@");
      scell.setValue(sval);
    }
    return { row: targetRow, movedTo: "" };
  });
}

function rtpSupplierStartRow_(sheet, checkCols) {
  var last = sheet.getLastRow();
  if (last < 3) return 3;
  var filled = sheet.getRange(last, 1, 1, checkCols || 5).getDisplayValues()[0].join("");
  return String(filled).trim() ? last + 2 : last + 1;
}

function rtpCreateEntry(tabKey, form) {
  return rtpRun_(function () {
    form = form || {};
    var cfg = rtpTabConfig_(tabKey);
    if (!cfg || cfg.archive) throw new Error("In diesem Tab können keine neuen Einträge angelegt werden.");
    var ss = rtpSS_();
    var stamp = rtpStampNow_();
    var date = rtpText_(form.date) || stamp.date;
    var time = rtpText_(form.time) || stamp.time;
    var who = rtpText_(form.who) || rtpDefaultWho_();

    if (cfg.kind === "alfah") {
      var sh = rtpEnsureAlfahTab_(ss, cfg);
      var rNumber = rtpCanonR_(form.rNumber);
      var stockId = rtpText_(form.stockId).toUpperCase();
      var orderId = rtpText_(form.orderId);
      var article = rtpText_(form.article);
      var name = rtpText_(form.name);
      if (!rNumber && !stockId && !article) {
        throw new Error("Bitte mindestens R-Nummer, Stock-ID oder Teilenummer angeben.");
      }
      var last = sh.getLastRow();
      if (rNumber && article && last >= 3) {
        var have = sh.getRange(3, 3, last - 2, 2).getDisplayValues();
        var wantKey = rNumber + "|" + rtpNormPart_(article);
        for (var i = 0; i < have.length; i++) {
          var k = rtpCanonR_(have[i][0]) + "|" + rtpNormPart_(have[i][1]);
          if (k === wantKey) {
            throw new Error(rNumber + " mit Teilenummer " + article + " steht bereits in Alfah (Zeile " + (3 + i) + ").");
          }
        }
      }
      var cursor = rtpSupplierStartRow_(sh, 5);
      var rowVals = [
        orderId,
        stockId,
        rNumber,
        article,
        name,
        rtpText_(form.supplier),
        rtpText_(form.brand),
        date,
        time,
        who
      ];
      var rng = sh.getRange(cursor, 1, 1, 10);
      rng.setNumberFormat("@");
      rng.setValues([rowVals]);
      var initState = String(form.state || "offen");
      rtpApplyAlfahState_(sh, cursor, initState, form.info != null ? form.info : null, false);
      sh.getRange(cursor, 15, 1, 2).setValues([[!!form.vorbereitet, !!form.geliefert]]);
      var movedTo = rtpAutoMoveSingleAlfahRow_(ss, sh, cursor);
      return { row: cursor, movedTo: movedTo };
    }

    var supSh = rtpEnsureSupplierTab_(ss, cfg);
    var sStock = rtpText_(form.stockId).toUpperCase();
    var sArticle = rtpText_(form.article);
    var sName = rtpText_(form.name);
    var sOrder = rtpText_(form.orderId);
    if (!sStock && !sArticle && !sName) {
      throw new Error("Bitte mindestens Stock-ID, Teilenummer oder Name angeben.");
    }
    var sCursor = rtpSupplierStartRow_(supSh, 5);
    var sVals = [
      sOrder,
      sStock,
      sArticle,
      sName,
      rtpText_(form.supplier) || cfg.defaultSupplier || cfg.label,
      rtpText_(form.brand),
      date,
      time,
      who,
      rtpText_(form.info)
    ];
    var sRng = supSh.getRange(sCursor, 1, 1, 10);
    sRng.setNumberFormat("@");
    sRng.setValues([sVals]);
    rtpPaintSupplierRow_(supSh, sCursor, cfg.color);
    return { row: sCursor, movedTo: "" };
  });
}

function rtpQuickSetGreen(row, ident) {
  return rtpRun_(function () {
    var cfg = rtpTabConfig_("alfah");
    var ss = rtpSS_();
    var sh = rtpFindSheet_(ss, cfg);
    if (!sh) throw new Error("Tabellenblatt Alfah fehlt.");
    var targetRow = rtpLocateRow_(sh, cfg, row, ident);
    var curInfo = String(sh.getRange(targetRow, 12).getDisplayValue() || "").trim();
    var nextInfo = (!curInfo || /veranlasst|eingeleitet/i.test(curInfo)) ? "Alfah System" : curInfo;
    rtpApplyAlfahState_(sh, targetRow, "bereit", nextInfo, true);
    var movedTo = rtpAutoMoveSingleAlfahRow_(ss, sh, targetRow);
    return { row: targetRow, movedTo: movedTo };
  });
}

function rtpWantedRIds_(raw) {
  var text = String(raw || "").toUpperCase();
  var out = [];
  var seen = {};
  var re = /R-\d{4,}|\b\d{4,6}\b/g;
  var m;
  while ((m = re.exec(text))) {
    var id = m[0].indexOf("R-") === 0 ? m[0] : "R-" + m[0];
    if (seen[id]) continue;
    seen[id] = true;
    out.push(id);
  }
  return out;
}

function rtpApplyRuecksendung(raw) {
  return rtpRun_(function () {
    var wanted = rtpWantedRIds_(raw);
    if (!wanted.length) return { done: [], skipped: [], missing: [], deliveredMoved: 0 };
    var ss = rtpSS_();
    var sh = rtpEnsureAlfahTab_(ss, rtpTabConfig_("alfah"));
    var last = sh.getLastRow();
    var foundRows = {};
    for (var w = 0; w < wanted.length; w++) foundRows[wanted[w]] = [];
    if (last >= 3) {
      var cases = sh.getRange(3, 3, last - 2, 1).getDisplayValues();
      for (var i = 0; i < cases.length; i++) {
        var id = String(cases[i][0] || "").toUpperCase().match(/R-\d{4,}/);
        if (!id || !foundRows[id[0]]) continue;
        foundRows[id[0]].push(3 + i);
      }
    }
    var done = [];
    var skipped = [];
    var missing = [];
    for (var k = 0; k < wanted.length; k++) {
      var key = wanted[k];
      var rows = foundRows[key] || [];
      if (!rows.length) {
        missing.push(key);
        continue;
      }
      var painted = 0;
      for (var r = 0; r < rows.length; r++) {
        var status = String(sh.getRange(rows[r], 11).getDisplayValue() || "").trim();
        var info = String(sh.getRange(rows[r], 12).getDisplayValue() || "").trim();
        if (status === "ABGELEHNT" || status === "STORNO") {
          skipped.push(key + " ist " + status);
          continue;
        }
        var rowBg = String(sh.getRange(rows[r], 1).getBackground() || "").toLowerCase();
        if (rtpIsRueckText_(status) && info && info !== "Alfah System" && rowBg !== RTP_ALFAH_ORANGE) {
          skipped.push(key + " schon grün (" + info + ")");
          continue;
        }
        rtpApplyAlfahState_(sh, rows[r], "bereit", "Alfah System", true);
        painted++;
      }
      if (painted) done.push(key);
    }
    var sweep = rtpSweepAlfahMoves_(ss, sh);
    return {
      done: done,
      skipped: skipped,
      missing: missing,
      deliveredMoved: sweep.deliveredMoved
    };
  });
}

function rtpRestoreAlfahRow(fromTabKey, row, ident) {
  return rtpRun_(function () {
    if (fromTabKey !== "geliefert" && fromTabKey !== "abgelehnt") {
      throw new Error("Wiederherstellen ist nur aus Alfah Geliefert oder Alfah Abgelehnt möglich.");
    }
    var ss = rtpSS_();
    var srcCfg = rtpTabConfig_(fromTabKey);
    var src = rtpFindSheet_(ss, srcCfg);
    if (!src) throw new Error("Quell-Tab fehlt.");
    var targetRow = rtpLocateRow_(src, srcCfg, row, ident);
    var dest = rtpEnsureAlfahTab_(ss, rtpTabConfig_("alfah"));
    var cursor = rtpSupplierStartRow_(dest, 5);
    src.getRange(targetRow, 1, 1, 16).copyTo(dest.getRange(cursor, 1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    if (fromTabKey === "geliefert") {
      rtpPutAlfahChecks_(dest, cursor);
      dest.getRange(cursor, 16).setValue(false);
    } else {
      rtpApplyAlfahState_(dest, cursor, "offen", "", true);
    }
    var above = targetRow <= 3 || rtpRowBlank_(src, targetRow - 1, 16);
    var span = above && rtpRowBlank_(src, targetRow + 1, 16) ? 2 : 1;
    src.deleteRows(targetRow, span);
    return { restoredRow: cursor };
  });
}
