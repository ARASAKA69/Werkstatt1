function nbpWithLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(NBP_LOCK_MS)) throw new Error("Sheet ist gerade belegt, bitte gleich nochmal versuchen.");
  try {
    return fn();
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function nbpRun_(fn) {
  if (!nbpAllowed_()) return { ok: false, message: "Kein Zugriff." };
  try {
    var res = nbpWithLock_(fn) || {};
    res.ok = true;
    return res;
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

function nbpSection_(key) {
  for (var i = 0; i < NBP_SECTIONS.length; i++) {
    if (NBP_SECTIONS[i].key === key) return NBP_SECTIONS[i];
  }
  throw new Error("Unbekannter Bereich: " + key);
}

function nbpFilled_(v) {
  return v !== null && v !== undefined && String(v).trim() !== "";
}

function nbpClean_(v) {
  return String(v == null ? "" : v).replace(/\u00a0/g, " ").trim();
}

function nbpCanon_(value, options) {
  var s = nbpClean_(value);
  if (!s || !options || !options.length) return s;
  var n = nbpNorm_(s);
  for (var i = 0; i < options.length; i++) {
    if (nbpNorm_(options[i]) === n) return options[i];
  }
  throw new Error("Ungültiger Wert \"" + s + "\". Erlaubt: " + options.join(", "));
}

function nbpParsePrice_(v) {
  var s = nbpClean_(v).replace(/€/g, "").replace(/\s/g, "");
  if (!s) return "";
  if (/^-?\d{1,3}(\.\d{3})*(,\d+)?$/.test(s)) s = s.replace(/\./g, "").replace(",", ".");
  else s = s.replace(",", ".");
  var n = Number(s);
  if (isNaN(n)) throw new Error("Preis ist keine Zahl: " + v);
  return n;
}

function nbpParseDateInput_(v) {
  var s = nbpClean_(v);
  if (!s) return "";
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  var ms = nbpToMs_(s);
  return ms ? new Date(ms) : s;
}

function nbpLastFilledRow_(sheet, col, start) {
  var last = sheet.getLastRow();
  if (last < start) return start - 1;
  var vals = sheet.getRange(start, col, last - start + 1, 1).getDisplayValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0] || "").trim()) return start + i;
  }
  return start - 1;
}

function nbpFindRowByValue_(sheet, col, start, value) {
  var target = String(value || "").trim();
  if (!target) return -1;
  var last = sheet.getLastRow();
  if (last < start) return -1;
  var vals = sheet.getRange(start, col, last - start + 1, 1).getDisplayValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || "").trim() === target) return start + i;
  }
  return -1;
}

function nbpProtectStock_(ss, sheet, row) {
  try {
    var range = sheet.getRange(row, 2);
    var a1 = range.getA1Notation();
    var prots = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var i = prots.length - 1; i >= 0; i--) {
      if (prots[i].getRange().getA1Notation() === a1) prots[i].remove();
    }
    var p = range.protect();
    p.setDescription("Nur berechtigte Personen dürfen diese Stock ID ändern");
    p.removeEditors(p.getEditors());
    try {
      p.addEditor(ss.getOwner().getEmail());
    } catch (eo) {}
    for (var e = 0; e < NBP_PROTECT_EDITORS.length; e++) {
      try {
        p.addEditor(NBP_PROTECT_EDITORS[e]);
      } catch (ee) {}
    }
    p.setWarningOnly(false);
  } catch (err) {}
}

function nbpUpsertNb_(ss, rowData, entryId) {
  var nb = ss.getSheetByName(NBP_NB_TAB);
  if (!nb) return "noop";
  var lo = nbpNbLayout_(nb);
  var et = nbpNorm_(rowData[2]);
  var complete = true;
  for (var i = 1; i <= 5; i++) {
    if (!nbpFilled_(rowData[i])) complete = false;
  }
  if (et === "ja" && complete) {
    var key = String(rowData[1]).trim();
    var found = nbpFindRowByValue_(nb, lo.cols.entryId, 2, entryId);
    var width = Math.max(14, lo.cols.entryId - 1);
    var old = found !== -1 ? nb.getRange(found, 1, 1, width).getValues()[0] : [];
    var out = new Array(width);
    out[0] = rowData[0] ? new Date(rowData[0]) : new Date();
    out[1] = key;
    out[2] = rowData[4] || "";
    out[3] = rowData[5] || "";
    out[4] = rowData[3] || "";
    for (var c = 5; c < width; c++) out[c] = old[c] != null ? old[c] : "";
    var target = found !== -1 ? found : Math.max(2, nb.getLastRow() + 1);
    nb.getRange(target, 1, 1, width).setValues([out]);
    nb.getRange(target, lo.cols.entryId).setValue(entryId);
    return found !== -1 ? "nb-updated" : "nb-created";
  }
  if (et === "nein" || et === "diagnose") {
    var last = nb.getLastRow();
    if (last < 2) return "noop";
    var vals = nb.getRange(2, lo.cols.entryId, last - 1, 1).getValues();
    var removed = 0;
    for (var r = vals.length - 1; r >= 0; r--) {
      if (String(vals[r][0]) === String(entryId)) {
        nb.deleteRow(r + 2);
        removed++;
      }
    }
    return removed ? "nb-deleted" : "noop";
  }
  return "noop";
}

function nbpEnqueue_(ss, sheet, sec, row, entryId) {
  try {
    var q = ss.getSheetByName(NBP_QUEUE_TAB);
    if (!q) return;
    var data = sheet.getRange(row, 1, 1, 6).getValues()[0];
    var key = String(data[1] || "").trim();
    if (!entryId && !key) return;
    q.appendRow([
      new Date(),
      sec.sheet,
      row,
      key ? "UPDATE" : "DELETE",
      key,
      JSON.stringify(data),
      "NEW",
      sec.code,
      entryId,
      nbpUser_()
    ]);
  } catch (err) {}
}

function nbpSyncRepairStatus_(ss, row) {
  var exit = ss.getSheetByName(NBP_EXIT_TAB);
  if (!exit) return "";
  var stock = String(exit.getRange(row, 2).getValue() || "").trim().toUpperCase();
  var m = exit.getRange(row, 13).getValue();
  var result = "";
  if (m === "" || m === null) {
    exit.getRange(row, 14).setValue("");
    return "";
  }
  try {
    if (!stock) {
      result = "Kein Wert/StockID";
    } else {
      var ziel = SpreadsheetApp.openById(NBP_ZIEL_ID).getSheetByName(NBP_ZIEL_TAB);
      if (!ziel) {
        result = "Zielsheet nicht gefunden";
      } else {
        var data = ziel.getDataRange().getValues();
        result = "StockID nicht gefunden";
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][0] || "").trim().toUpperCase() === stock) {
            ziel.getRange(i + 1, 5).setValue(m);
            result = "Übertragen";
            break;
          }
        }
      }
    }
  } catch (err) {
    result = "Fehler: " + err;
  }
  exit.getRange(row, 14).setValue(result);
  return result;
}

function nbpPushNbToExit_(ss, nb, lo, row) {
  var art = nbpNorm_(nb.getRange(row, lo.cols.art).getValue());
  if (art.indexOf("exit mechanik") === -1) return false;
  var exit = ss.getSheetByName(NBP_EXIT_TAB);
  if (!exit) return false;
  var stock = String(nb.getRange(row, lo.cols.stockId).getValue() || "").trim().toUpperCase();
  var eid = String(nb.getRange(row, lo.cols.entryId).getValue() || "").trim();
  var exLay = nbpInputLayout_(exit);
  var exitRow = -1;
  if (eid) {
    exitRow = nbpFindRowByValue_(exit, exLay.eidCol, exLay.dataStart, eid);
  } else if (stock) {
    var last = exit.getLastRow();
    if (last >= exLay.dataStart) {
      var st = exit.getRange(exLay.dataStart, 2, last - exLay.dataStart + 1, 1).getValues();
      for (var i = st.length - 1; i >= 0; i--) {
        if (String(st[i][0] || "").trim().toUpperCase() === stock) {
          exitRow = exLay.dataStart + i;
          break;
        }
      }
    }
  }
  if (exitRow === -1) return false;
  var src = {
    f: nb.getRange(row, lo.cols.preis).getValue(),
    i: nb.getRange(row, lo.cols.costGate).getValue(),
    j: nb.getRange(row, lo.cols.carol).getValue(),
    k: nb.getRange(row, lo.cols.status).getValue(),
    l: nb.getRange(row, lo.cols.statusAenderung).getValue()
  };
  var cur = exit.getRange(exitRow, 7, 1, 6).getValues()[0];
  var out = cur.slice();
  var map = [["f", 0], ["i", 2], ["j", 3], ["k", 4], ["l", 5]];
  var changed = false;
  for (var m = 0; m < map.length; m++) {
    var v = src[map[m][0]];
    var idx = map[m][1];
    if (!nbpFilled_(v)) continue;
    var same = nbpIsDate_(v) && nbpIsDate_(cur[idx]) ? v.getTime() === cur[idx].getTime() : String(v) === String(cur[idx]);
    if (same) continue;
    out[idx] = v;
    changed = true;
  }
  if (!changed) return false;
  exit.getRange(exitRow, 7, 1, 6).setValues([out]);
  return true;
}

function nbpReadForm_(sec, sheet, lay, form, isNew) {
  form = form || {};
  var stock = nbpClean_(form.stockId).replace(/\s+/g, "").toUpperCase();
  if (isNew && !stock) throw new Error("Stock ID fehlt.");
  var etOpts = nbpColumnOptions_(sheet, 3, lay.dataStart);
  var artOpts = nbpColumnOptions_(sheet, 5, lay.dataStart);
  var bOpts = nbpColumnOptions_(sheet, 6, lay.dataStart);
  var out = {
    stockId: stock,
    et: nbpCanon_(form.et, etOpts),
    grund: nbpClean_(form.grund),
    art: nbpCanon_(form.art, artOpts),
    bearbeiter: nbpCanon_(form.bearbeiter, bOpts),
    bemerkung: nbpClean_(form.bemerkung),
    extra: {}
  };
  if (isNew) {
    if (!out.et) throw new Error("ET/Diagnose fehlt.");
    if (!out.grund) throw new Error("Grund fehlt.");
    if (!out.bearbeiter) throw new Error("Bearbeiter fehlt.");
    if (nbpNorm_(out.et) === "ja" && !out.art) throw new Error("Art der Nachbestellung fehlt (ET = Ja).");
  }
  var fx = form.extra || {};
  for (var i = 0; i < lay.extras.length; i++) {
    var ex = lay.extras[i];
    if (!ex.editable) continue;
    if (!Object.prototype.hasOwnProperty.call(fx, String(ex.col))) continue;
    var raw = fx[String(ex.col)];
    var opts = nbpColumnOptions_(sheet, ex.col, lay.dataStart);
    var val;
    if (opts) val = nbpCanon_(raw, opts);
    else if (/liefert/i.test(ex.label)) val = nbpParseDateInput_(raw);
    else if (/claim/i.test(ex.label)) val = nbpClean_(raw) ? nbpParsePrice_(raw) : "";
    else val = nbpClean_(raw);
    out.extra[ex.col] = val;
  }
  return out;
}

function nbpAfterInputWrite_(ss, sheet, sec, row, entryId, liefertTouched) {
  var data = sheet.getRange(row, 1, 1, 6).getValues()[0];
  var nbResult = nbpUpsertNb_(ss, data, entryId);
  nbpEnqueue_(ss, sheet, sec, row, entryId);
  var zielResult = "";
  if (sec.sheet === NBP_EXIT_TAB && liefertTouched) zielResult = nbpSyncRepairStatus_(ss, row);
  return { row: row, entryId: entryId, nb: nbResult, ziel: zielResult };
}

function nbpCreateInputEntry(sectionKey, form) {
  return nbpRun_(function () {
    var sec = nbpSection_(sectionKey);
    var ss = nbpSS_();
    var sheet = ss.getSheetByName(sec.sheet);
    if (!sheet) throw new Error("Tabellenblatt " + sec.sheet + " fehlt.");
    var lay = nbpInputLayout_(sheet);
    var f = nbpReadForm_(sec, sheet, lay, form, true);
    var lastFilled = nbpLastFilledRow_(sheet, 2, lay.dataStart);
    var row = Math.max(lay.dataStart, lastFilled + 1);
    if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 1);
    if (lastFilled >= lay.dataStart) {
      var srcRow = sheet.getRange(lastFilled, 1, 1, lay.width);
      var dst = sheet.getRange(row, 1, 1, lay.width);
      srcRow.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      srcRow.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }
    var entryId = Utilities.getUuid();
    var vals = [new Date(), f.stockId, f.et, f.grund, f.art, f.bearbeiter];
    if (lay.hasBem) vals.push(f.bemerkung);
    sheet.getRange(row, 1, 1, vals.length).setValues([vals]);
    sheet.getRange(row, lay.eidCol).setValue(entryId);
    var liefert = false;
    for (var col in f.extra) {
      sheet.getRange(row, Number(col)).setValue(f.extra[col]);
      if (Number(col) === 13) liefert = true;
    }
    nbpProtectStock_(ss, sheet, row);
    SpreadsheetApp.flush();
    return nbpAfterInputWrite_(ss, sheet, sec, row, entryId, liefert);
  });
}

function nbpUpdateInputEntry(sectionKey, row, entryId, patch) {
  return nbpRun_(function () {
    var sec = nbpSection_(sectionKey);
    var ss = nbpSS_();
    var sheet = ss.getSheetByName(sec.sheet);
    if (!sheet) throw new Error("Tabellenblatt " + sec.sheet + " fehlt.");
    var lay = nbpInputLayout_(sheet);
    row = Number(row);
    var eid = String(entryId || "").trim();
    if (eid && String(sheet.getRange(row, lay.eidCol).getValue() || "").trim() !== eid) {
      row = nbpFindRowByValue_(sheet, lay.eidCol, lay.dataStart, eid);
      if (row === -1) throw new Error("Eintrag nicht mehr gefunden. Bitte neu laden.");
    }
    if (row < lay.dataStart) throw new Error("Ungültige Zeile.");
    var stockNow = String(sheet.getRange(row, 2).getValue() || "").trim();
    if (!stockNow) throw new Error("Zeile hat keine Stock ID mehr. Bitte neu laden.");
    patch = patch || {};
    var f = nbpReadForm_(sec, sheet, lay, patch, false);
    var cur = sheet.getRange(row, 1, 1, lay.hasBem ? 7 : 6).getValues()[0];
    var map = { et: 2, grund: 3, art: 4, bearbeiter: 5 };
    if (lay.hasBem) map.bemerkung = 6;
    var core = false;
    for (var k in map) {
      if (!Object.prototype.hasOwnProperty.call(patch, k)) continue;
      cur[map[k]] = f[k];
      if (k !== "bemerkung") core = true;
    }
    if (core && nbpNorm_(cur[2]) === "ja" && !nbpFilled_(cur[4])) {
      throw new Error("Art der Nachbestellung fehlt (ET = Ja).");
    }
    if (!nbpFilled_(cur[0])) cur[0] = new Date();
    sheet.getRange(row, 1, 1, cur.length).setValues([cur]);
    if (!eid) {
      eid = String(sheet.getRange(row, lay.eidCol).getValue() || "").trim();
      if (!eid) {
        eid = Utilities.getUuid();
        sheet.getRange(row, lay.eidCol).setValue(eid);
      }
    }
    var liefert = false;
    for (var col in f.extra) {
      sheet.getRange(row, Number(col)).setValue(f.extra[col]);
      if (Number(col) === 13) liefert = true;
    }
    SpreadsheetApp.flush();
    if (!core) {
      var res = { row: row, entryId: eid, nb: "noop", ziel: "" };
      if (sec.sheet === NBP_EXIT_TAB && liefert) res.ziel = nbpSyncRepairStatus_(ss, row);
      return res;
    }
    return nbpAfterInputWrite_(ss, sheet, sec, row, eid, liefert);
  });
}

function nbpUpdateNbEntry(row, entryId, patch) {
  return nbpRun_(function () {
    var ss = nbpSS_();
    var nb = ss.getSheetByName(NBP_NB_TAB);
    if (!nb) throw new Error("Tabellenblatt " + NBP_NB_TAB + " fehlt.");
    var lo = nbpNbLayout_(nb);
    row = Number(row);
    var eid = String(entryId || "").trim();
    if (eid && String(nb.getRange(row, lo.cols.entryId).getValue() || "").trim() !== eid) {
      row = nbpFindRowByValue_(nb, lo.cols.entryId, lo.dataStart, eid);
      if (row === -1) throw new Error("Eintrag nicht mehr gefunden. Bitte neu laden.");
    }
    if (row < lo.dataStart) throw new Error("Ungültige Zeile.");
    patch = patch || {};
    var written = [];
    var statusTouched = false;
    var mirrorTouched = false;
    for (var i = 0; i < NBP_NB_FIELDS.length; i++) {
      var fd = NBP_NB_FIELDS[i];
      if (!fd.edit) continue;
      if (!Object.prototype.hasOwnProperty.call(patch, fd.key)) continue;
      var col = lo.cols[fd.key];
      if (!col) continue;
      var val;
      if (fd.edit === "list") {
        val = nbpCanon_(patch[fd.key], nbpColumnOptions_(nb, col, lo.dataStart));
      } else if (fd.key === "preis") {
        val = nbpParsePrice_(patch[fd.key]);
      } else {
        val = nbpClean_(patch[fd.key]);
      }
      var before = nb.getRange(row, col).getValue();
      var same = String(before == null ? "" : before) === String(val);
      if (same) continue;
      nb.getRange(row, col).setValue(val);
      written.push(fd.key);
      if (fd.key === "status") statusTouched = true;
      if (fd.key === "preis" || fd.key === "costGate" || fd.key === "carol" || fd.key === "status") mirrorTouched = true;
    }
    if (statusTouched) {
      var k = nb.getRange(row, lo.cols.status).getValue();
      nb.getRange(row, lo.cols.statusAenderung).setValue(nbpFilled_(k) ? new Date() : "");
    }
    var exitPushed = false;
    if (mirrorTouched) {
      SpreadsheetApp.flush();
      try {
        exitPushed = nbpPushNbToExit_(ss, nb, lo, row);
      } catch (err) {}
    }
    return { row: row, written: written, exitPushed: exitPushed };
  });
}
