const FERTIG_ARCHIVE_MAIN_ID = "1VGCAHUbOPgsInQICA1GnrtKg1EPK1d1zWB-GkLi6iVE";
const FERTIG_ARCHIVE_STATUSES = {
  fertiggestellt: true,
  b2a1: true
};
const FERTIG_ARCHIVE_LOCK_MS = 120000;
const FERTIG_ARCHIVE_TRIGGER_FN = "archiveFertiggestelltRows";
const FERTIG_ARCHIVE_TRIGGER_TZ = "Europe/Berlin";
const FERTIG_ARCHIVE_TRIGGER_HOUR = 23;

function normFertigHeader_(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .replace(/\u00a0/g, " ")
    .replace(/[\u200b-\u200d\ufeff]/g, "")
    .replace(/\s+/g, " ");
}

function fertigStatusNorm_(val) {
  if (val == null || val === "") return "";
  if (Object.prototype.toString.call(val) === "[object Date]") return "";
  if (typeof val === "object") {
    try {
      const t = String(val);
      if (t && t !== "[object Object]") return normFertigHeader_(t);
    } catch (e) {}
    return "";
  }
  return normFertigHeader_(val);
}

function isFertiggestelltStatus_(value, displayValue) {
  const a = fertigStatusNorm_(value);
  const b = fertigStatusNorm_(displayValue);
  if (a && FERTIG_ARCHIVE_STATUSES[a]) return true;
  if (b && FERTIG_ARCHIVE_STATUSES[b]) return true;
  return false;
}

function fertigSheetRowRange_(sheet, row, lastCol) {
  return sheet.getRange(row, 1, 1, lastCol);
}

function getFertigArchiveSpreadsheet_() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}
  return SpreadsheetApp.openById(FERTIG_ARCHIVE_MAIN_ID);
}

function findStatusColInHeaderRow_(headerValues) {
  for (let c = 0; c < headerValues.length; c++) {
    if (normFertigHeader_(headerValues[c]) === "status") return c + 1;
  }
  return 11;
}

function nachbestellungFertigBounds_(sheet) {
  const lastCol = Math.max(1, sheet.getLastColumn());
  const lastRow = sheet.getLastRow();
  const scanRows = Math.min(40, Math.max(1, lastRow));
  let headerRow = 1;
  let statusCol = 11;
  let found = false;
  for (let r = 1; r <= scanRows; r++) {
    const row = fertigSheetRowRange_(sheet, r, lastCol).getValues()[0];
    let hasStatus = false;
    let hasStock = false;
    let rowStatusCol = 11;
    for (let c = 0; c < row.length; c++) {
      const n = normFertigHeader_(row[c]);
      if (n === "status") {
        hasStatus = true;
        rowStatusCol = c + 1;
      }
      if (n === "stock id") hasStock = true;
    }
    if (hasStatus && hasStock) {
      headerRow = r;
      statusCol = rowStatusCol;
      found = true;
      break;
    }
  }
  const dataStart = found ? headerRow + 1 : 2;
  const trim = 1;
  const dataEnd = lastRow > trim ? lastRow - trim : lastRow;
  if (dataEnd < dataStart) return null;
  return { dataStart: dataStart, dataEnd: dataEnd, statusCol: statusCol, headerRow: headerRow };
}

function inputExitFertigBounds_(sheet) {
  const lastCol = Math.max(1, sheet.getLastColumn());
  const headerRow = 2;
  let statusCol = 11;
  if (sheet.getLastRow() >= headerRow) {
    const hdr = fertigSheetRowRange_(sheet, headerRow, lastCol).getValues()[0];
    statusCol = findStatusColInHeaderRow_(hdr);
  }
  const dataStart = 3;
  const last = sheet.getLastRow();
  if (last < dataStart) return null;
  const dataEnd = last === dataStart ? last : last - 1;
  if (dataEnd < dataStart) return null;
  return { dataStart: dataStart, dataEnd: dataEnd, statusCol: statusCol, headerRow: headerRow };
}

function ensureFertigArchiveSheet_(ss, sourceName, archiveName, headerRow) {
  let archive = ss.getSheetByName(archiveName);
  if (archive) return archive;
  const source = ss.getSheetByName(sourceName);
  archive = ss.insertSheet(archiveName);
  if (!source) return archive;
  const lastCol = Math.max(1, source.getLastColumn());
  const hr = headerRow > 0 ? headerRow : 1;
  fertigSheetRowRange_(source, hr, lastCol).copyTo(fertigSheetRowRange_(archive, 1, lastCol));
  return archive;
}

function collectFertiggestelltRows_(sheet, bounds) {
  const numRows = bounds.dataEnd - bounds.dataStart + 1;
  if (numRows <= 0) return [];
  const statusRange = sheet.getRange(bounds.dataStart, bounds.statusCol, numRows, 1);
  const values = statusRange.getValues();
  const displays = statusRange.getDisplayValues();
  const rows = [];
  for (let i = 0; i < numRows; i++) {
    if (!isFertiggestelltStatus_(values[i][0], displays[i][0])) continue;
    rows.push(bounds.dataStart + i);
  }
  return rows;
}


function moveFertiggestelltRowsOnSheet_(ss, sourceName, archiveName, boundsFn) {
  const sheet = ss.getSheetByName(sourceName);
  if (!sheet) return 0;
  const bounds = boundsFn(sheet);
  if (!bounds) return 0;
  const archive = ensureFertigArchiveSheet_(ss, sourceName, archiveName, bounds.headerRow);
  const lastCol = Math.max(1, sheet.getLastColumn());
  const toMove = collectFertiggestelltRows_(sheet, bounds);
  let moved = 0;
  for (let i = toMove.length - 1; i >= 0; i--) {
    const row = toMove[i];
    try {
      if (typeof archiveCopyFormattedRowToBottom === "function") {
        archiveCopyFormattedRowToBottom(sheet, row, archive, lastCol);
      } else {
        const destRow = archive.getLastRow() + 1;
        fertigSheetRowRange_(sheet, row, lastCol).copyTo(fertigSheetRowRange_(archive, destRow, lastCol));
      }
      sheet.deleteRow(row);
      moved++;
    } catch (e) {
      try {
        if (typeof logDebug === "function") {
          logDebug("moveFertiggestelltRowsOnSheet_ " + sourceName + " row " + row + ": " + e);
        }
      } catch (e2) {}
    }
  }
  return moved;
}

function archiveFertiggestelltRows() {
  let lock = null;
  if (typeof acquireScriptLock_ === "function") {
    lock = acquireScriptLock_(FERTIG_ARCHIVE_LOCK_MS);
  } else {
    lock = LockService.getScriptLock();
    if (!lock.tryLock(FERTIG_ARCHIVE_LOCK_MS)) lock = null;
  }
  if (!lock) {
    try {
      if (typeof logDebug === "function") {
        logDebug("archiveFertiggestelltRows: LOCK nicht frei – Lauf übersprungen (Cleanup/Queue/Sync aktiv?)");
      }
    } catch (e) {}
    return { ok: false, fehler: "Lock nicht frei – anderer Script-Lauf aktiv. Kurz warten und runFertiggestelltArchiveNow erneut ausführen." };
  }
  try {
    const ss = getFertigArchiveSpreadsheet_();
    if (!ss) return { ok: false, fehler: "Hauptmappe nicht erreichbar." };
    const nbMoved = moveFertiggestelltRowsOnSheet_(ss, "Nachbestellung", "Nachbestellung_Archiv", nachbestellungFertigBounds_);
    const exitMoved = moveFertiggestelltRowsOnSheet_(ss, "Input Exit", "Input Exit_Archiv", inputExitFertigBounds_);
    const total = nbMoved + exitMoved;
    try {
      if (typeof logDebug === "function") {
        logDebug(
          "archiveFertiggestelltRows: Nachbestellung=" + nbMoved + " Input Exit=" + exitMoved + " total=" + total
        );
      }
    } catch (e) {}
    return { ok: true, nbMoved: nbMoved, exitMoved: exitMoved, total: total };
  } catch (err) {
    return { ok: false, fehler: String(err && err.message ? err.message : err) };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

function installFertiggestelltArchiveTrigger_(silent) {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === FERTIG_ARCHIVE_TRIGGER_FN) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger(FERTIG_ARCHIVE_TRIGGER_FN)
    .timeBased()
    .atHour(FERTIG_ARCHIVE_TRIGGER_HOUR)
    .everyDays(1)
    .inTimezone(FERTIG_ARCHIVE_TRIGGER_TZ)
    .create();
  if (silent) return;
  try {
    SpreadsheetApp.getUi().alert(
      "Archiv-Trigger installiert (Status: Fertiggestellt, B2A1).\n" +
        "Läuft täglich um " +
        FERTIG_ARCHIVE_TRIGGER_HOUR +
        ":00 Uhr (" +
        FERTIG_ARCHIVE_TRIGGER_TZ +
        ")."
    );
  } catch (e) {}
}

function installFertiggestelltArchiveTrigger() {
  installFertiggestelltArchiveTrigger_(false);
}

function uninstallFertiggestelltArchiveTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === FERTIG_ARCHIVE_TRIGGER_FN) {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  try {
    SpreadsheetApp.getUi().alert(
      removed > 0
        ? "Fertiggestellt-Archiv Trigger entfernt (" + removed + ")."
        : "Kein Fertiggestellt-Archiv Trigger gefunden."
    );
  } catch (e) {}
}

function runFertiggestelltArchiveNow() {
  const res = archiveFertiggestelltRows();
  try {
    if (res && res.ok) {
      SpreadsheetApp.getUi().alert(
        "Archivierung fertig.\n" +
          "Nachbestellung: " + (res.nbMoved || 0) + " Zeilen\n" +
          "Input Exit: " + (res.exitMoved || 0) + " Zeilen\n" +
          "(Status: Fertiggestellt oder B2A1)"
      );
    } else {
      SpreadsheetApp.getUi().alert((res && res.fehler) || "Archivierung hat nicht geklappt.");
    }
  } catch (e) {}
  return res;
}

function ensureFertiggestelltArchiveTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === FERTIG_ARCHIVE_TRIGGER_FN) return;
  }
  installFertiggestelltArchiveTrigger_(true);
}
