var CONFIG = {
    refurbSpreadsheetId: "13Oh7gDT8NAul2s0cwQUeaGwMcS3B2MYu0QOdFNMhXzM",
    refurbSheetName: "Refurbisment List",
    refurbStockCol: 2,
    refurbStatusCol: 28,
    nachbestellSpreadsheetId: "1VGCAHUbOPgsInQICA1GnrtKg1EPK1d1zWB-GkLi6iVE",
    nachbestellSheetName: "Nachbestellung",
    nachbestellStockCol: 2,
    nachbestellStatusCol: 13,
    nachbestellColor: "#9900ff",
    triggerStatus: "Tagesliste",
    lagerSheetName: "LAGER",
    tageslisteSheetName: "Tagesliste",
    greenColor: "#00ff00",
    manualMoveColor: "#46bdc6",
    removeFromLager: true,
    addMissingToLager: true,
    intervalMinutes: 2,
    activeFromHour: 6,
    activeToHour: 21,
    timezone: "Europe/Berlin",
    ignoreColors: ["#00ffff", "#ffff00", "#ff9900", "#ff0000", "#4a86e8"],
    ignoreNoteKeyword: "NACHBESTELLUNG"
};

function setupTrigger() {
    removeTriggers_();
    ScriptApp.newTrigger("syncTagesliste").timeBased().everyMinutes(1).create();
}

function removeTriggers_() {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === "syncTagesliste") {
            ScriptApp.deleteTrigger(triggers[i]);
        }
    }
}

function isWithinActiveHours_() {
    var hour = Number(Utilities.formatDate(new Date(), CONFIG.timezone, "H"));
    return hour >= CONFIG.activeFromHour && hour < CONFIG.activeToHour;
}

function syncTagesliste() {
    if (!isWithinActiveHours_()) return;
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return;
    try {
        var props = PropertiesService.getScriptProperties();
        var now = Date.now();
        var last = Number(props.getProperty("lagerSyncLastRun") || 0);
        if (now - last < (CONFIG.intervalMinutes * 60 - 5) * 1000) return;
        props.setProperty("lagerSyncLastRun", String(now));
        runSync_();
    } finally {
        lock.releaseLock();
    }
}

function manualSync() {
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(60000)) return;
    try {
        runSync_();
    } finally {
        lock.releaseLock();
    }
}

function runSync_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var lager = ss.getSheetByName(CONFIG.lagerSheetName);
    var tagesliste = ss.getSheetByName(CONFIG.tageslisteSheetName);
    if (!lager || !tagesliste) return;

    var entries = collectEntries_();
    var statusMap = buildStatusMap_(entries);

    var lagerData = readSheetData_(lager);
    var tagesData = readSheetData_(tagesliste);
    var tagesRegale = getRegalColumnsFromValues_(tagesData.values);
    var existingTages = getRegalIdsFromValues_(tagesData.values);

    var tagesMoveColor = {};
    for (var i = 0; i < entries.length; i++) {
        if (isTageslisteStatus_(entries[i].status)) {
            tagesMoveColor[normalizeId_(entries[i].id)] = entries[i].tagesColor;
        }
    }

    var manualColor = String(CONFIG.manualMoveColor).toLowerCase();
    var lagerItems = scanRegaleFromData_(lagerData);
    for (var li = 0; li < lagerItems.length; li++) {
        var item = lagerItems[li];
        var key = normalizeId_(item.stockId);
        var moveColor = null;

        if (item.background === manualColor) {
            moveColor = CONFIG.greenColor;
        } else if (tagesMoveColor.hasOwnProperty(key) && !isIgnored_(item)) {
            moveColor = tagesMoveColor[key];
        } else {
            continue;
        }

        if (!existingTages[key]) {
            var destRow = placeInGap_(tagesliste, tagesData, tagesRegale, item.regal, item.stockId, moveColor);
            if (destRow) {
                existingTages[key] = true;
                var destCol = tagesRegale[item.regal].col;
                copyNote_(lager, item.row, item.col, tagesliste, destRow, destCol);
            }
        }

        if (!existingTages[key]) continue;

        clearCell_(lager, item.row, item.col);
        sheetDataClear_(lagerData, item.row, item.col);
    }

    var tagesItems = scanRegaleFromData_(tagesData);
    for (var j = 0; j < tagesItems.length; j++) {
        var t = tagesItems[j];
        var tkey = normalizeId_(t.stockId);
        if (statusMap.hasOwnProperty(tkey) && !isTageslisteStatus_(statusMap[tkey])) {
            clearCell_(tagesliste, t.row, t.col);
            sheetDataClear_(tagesData, t.row, t.col);
        }
    }

    if (CONFIG.addMissingToLager) {
        var lagerRegale = getRegalColumnsFromValues_(lagerData.values);
        var lagerLocations = getRegalIdLocations_(lagerData.values);
        for (var k = 0; k < entries.length; k++) {
            var regalKey = parseRegalStatus_(entries[k].status);
            if (!regalKey) continue;
            var idn = normalizeId_(entries[k].id);
            var loc = lagerLocations[idn];
            if (loc && loc.regal === regalKey) continue;
            var destRow = placeInGap_(lager, lagerData, lagerRegale, regalKey, entries[k].id, entries[k].lagerColor);
            if (!destRow) continue;
            if (loc) {
                var oldCell = lager.getRange(loc.row, loc.col);
                oldCell.clearContent();
                oldCell.setBackground(null);
                oldCell.clearNote();
                sheetDataClear_(lagerData, loc.row, loc.col);
            }
            lagerLocations[idn] = { regal: regalKey, row: destRow, col: lagerRegale[regalKey].col };
        }
    }
}

function collectEntries_() {
    var list = [];
    var refurb = getSourceEntries_(CONFIG.refurbSpreadsheetId, CONFIG.refurbSheetName, CONFIG.refurbStockCol, CONFIG.refurbStatusCol);
    for (var i = 0; i < refurb.length; i++) {
        list.push({ id: refurb[i].id, status: refurb[i].status, lagerColor: null, tagesColor: CONFIG.greenColor });
    }
    var nach = getSourceEntries_(CONFIG.nachbestellSpreadsheetId, CONFIG.nachbestellSheetName, CONFIG.nachbestellStockCol, CONFIG.nachbestellStatusCol);
    for (var j = 0; j < nach.length; j++) {
        list.push({ id: nach[j].id, status: nach[j].status, lagerColor: CONFIG.nachbestellColor, tagesColor: CONFIG.nachbestellColor });
    }
    return list;
}

function buildStatusMap_(entries) {
    var map = {};
    for (var e = 0; e < entries.length; e++) {
        var id = normalizeId_(entries[e].id);
        var status = entries[e].status;
        if (!map[id] || isTageslisteStatus_(status)) map[id] = status;
    }
    return map;
}

function readSheetData_(sheet) {
    var range = sheet.getDataRange();
    return {
        sheet: sheet,
        values: range.getValues(),
        backgrounds: range.getBackgrounds(),
        notes: range.getNotes()
    };
}

function sheetDataClear_(data, row, col) {
    var row0 = row - 1;
    var col0 = col - 1;
    if (row0 < 0 || col0 < 0 || row0 >= data.values.length) return;
    if (col0 >= data.values[row0].length) return;
    data.values[row0][col0] = "";
    data.backgrounds[row0][col0] = null;
    data.notes[row0][col0] = "";
}

function getSourceEntries_(spreadsheetId, sheetName, stockCol, statusCol) {
    var entries = [];
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return entries;
    var lastRow = sh.getLastRow();
    if (lastRow < 1) return entries;
    var maxCol = Math.max(stockCol, statusCol);
    var colLetter = columnToLetter_(maxCol);
    var values = sh.getRange("A1:" + colLetter + lastRow).getValues();
    var headerRow0 = -1;
    for (var r = 0; r < values.length; r++) {
        var head = String(values[r][stockCol - 1] || "").trim().toLowerCase();
        if (head === "stock id" || head === "stockid") {
            headerRow0 = r;
            break;
        }
    }
    var start = headerRow0 >= 0 ? headerRow0 + 1 : 0;
    for (var r2 = start; r2 < values.length; r2++) {
        var rawId = values[r2][stockCol - 1];
        var status = values[r2][statusCol - 1];
        if (rawId === "" || rawId === null) continue;
        if (isRegalHeader_(rawId)) continue;
        entries.push({ id: rawId, status: String(status).trim() });
    }
    return entries;
}

function isTageslisteStatus_(status) {
    if (status === undefined || status === null || status === "") return false;
    return String(status).trim().toLowerCase() === CONFIG.triggerStatus.toLowerCase();
}

function parseRegalStatus_(status) {
    var s = String(status).trim();
    if (!/regal/i.test(s)) return null;
    var m = s.match(/(\d+)\.(\d+)/);
    if (!m) return null;
    var a = Number(m[1]);
    var b = Number(m[2]);
    if (a >= 1 && a <= 5 && b >= 1 && b <= 8) return "REGAL " + a + "." + b;
    return null;
}

function getRegalEndRow0_(values, headerRow0, col0) {
    for (var rr = headerRow0 + 1; rr < values.length; rr++) {
        if (isRegalHeader_(values[rr][col0])) return rr;
    }
    return values.length;
}

function findFirstEmptyRow_(values, headerRow0, col0, endRow0) {
    for (var r = headerRow0 + 1; r < endRow0; r++) {
        var v = values[r][col0];
        if (v === "" || v === null) return r + 1;
    }
    return endRow0 + 1;
}

function ensureSheetRow_(data, row0, col0) {
    var cols = data.values.length ? data.values[0].length : col0 + 1;
    while (data.values.length <= row0) {
        data.values.push(new Array(cols).fill(""));
        data.backgrounds.push(new Array(cols).fill(null));
        data.notes.push(new Array(cols).fill(""));
    }
}

function placeInGap_(sheet, sheetData, regalMap, regalKey, stockId, color) {
    var pos = regalMap[regalKey];
    if (!pos) return 0;
    var col0 = pos.col - 1;
    var headerRow0 = pos.row - 1;
    var endRow0 = getRegalEndRow0_(sheetData.values, headerRow0, col0);
    var row = findFirstEmptyRow_(sheetData.values, headerRow0, col0, endRow0);
    var row0 = row - 1;
    if (row0 <= headerRow0) return 0;
    if (row0 >= endRow0 && endRow0 < sheetData.values.length) return 0;
    ensureSheetRow_(sheetData, row0, col0);
    var cell = sheet.getRange(row, pos.col);
    cell.setValue(stockId);
    cell.setBackground(color ? color : null);
    cell.clearNote();
    sheetData.values[row0][col0] = stockId;
    sheetData.backgrounds[row0][col0] = color ? color : null;
    sheetData.notes[row0][col0] = "";
    return row;
}

function copyNote_(fromSheet, fromRow, fromCol, toSheet, toRow, toCol) {
    var note = "";
    try { note = fromSheet.getRange(fromRow, fromCol).getNote() || ""; } catch (e) {}
    if (note) toSheet.getRange(toRow, toCol).setNote(note);
}

function clearCell_(sheet, row, col) {
    if (!CONFIG.removeFromLager) {
        sheet.getRange(row, col).setBackground(CONFIG.greenColor);
        return;
    }
    var cell = sheet.getRange(row, col);
    cell.clearContent();
    cell.setBackground(null);
    cell.clearNote();
}

function scanRegaleFromData_(data) {
    var items = [];
    var values = data.values;
    var backgrounds = data.backgrounds;
    var notes = data.notes;
    var numRows = values.length;
    var numCols = numRows ? values[0].length : 0;
    for (var c = 0; c < numCols; c++) {
        for (var r = 0; r < numRows; r++) {
            if (!isRegalHeader_(values[r][c])) continue;
            var regalKey = normalizeRegal_(values[r][c]);
            var endRow0 = getRegalEndRow0_(values, r, c);
            for (var rr = r + 1; rr < endRow0; rr++) {
                var v = values[rr][c];
                if (v === "" || v === null) continue;
                if (isRegalHeader_(v)) break;
                items.push({
                    stockId: v,
                    regal: regalKey,
                    row: rr + 1,
                    col: c + 1,
                    background: String(backgrounds[rr][c]).toLowerCase(),
                    note: String(notes[rr][c] || "")
                });
            }
        }
    }
    return items;
}

function isIgnored_(item) {
    for (var i = 0; i < CONFIG.ignoreColors.length; i++) {
        if (item.background === String(CONFIG.ignoreColors[i]).toLowerCase()) return true;
    }
    if (CONFIG.ignoreNoteKeyword) {
        if (item.note.toUpperCase().indexOf(CONFIG.ignoreNoteKeyword.toUpperCase()) !== -1) return true;
    }
    return false;
}

function getRegalColumnsFromValues_(values) {
    var map = {};
    for (var r = 0; r < values.length; r++) {
        for (var c = 0; c < values[r].length; c++) {
            if (isRegalHeader_(values[r][c])) {
                map[normalizeRegal_(values[r][c])] = { row: r + 1, col: c + 1 };
            }
        }
    }
    return map;
}

function getRegalIdsFromValues_(values) {
    var set = {};
    var numRows = values.length;
    var numCols = numRows ? values[0].length : 0;
    for (var c = 0; c < numCols; c++) {
        for (var r = 0; r < numRows; r++) {
            if (!isRegalHeader_(values[r][c])) continue;
            var endRow0 = getRegalEndRow0_(values, r, c);
            for (var rr = r + 1; rr < endRow0; rr++) {
                var v = values[rr][c];
                if (v === "" || v === null) continue;
                if (isRegalHeader_(v)) break;
                set[normalizeId_(v)] = true;
            }
        }
    }
    return set;
}

function isRegalHeader_(v) {
    if (v === "" || v === null) return false;
    return /^\s*regal\s+\d+\.\d+\s*$/i.test(String(v));
}

function normalizeRegal_(v) {
    var m = String(v).match(/(\d+)\.(\d+)/);
    if (!m) return String(v).trim().toUpperCase();
    return "REGAL " + m[1] + "." + m[2];
}

function normalizeId_(v) {
    return String(v).replace(/\s+/g, "").toUpperCase();
}

function columnToLetter_(col) {
    var letter = "";
    while (col > 0) {
        var mod = (col - 1) % 26;
        letter = String.fromCharCode(65 + mod) + letter;
        col = Math.floor((col - 1) / 26);
    }
    return letter;
}

function debugAllToLager() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var lager = ss.getSheetByName(CONFIG.lagerSheetName);
    if (!lager) {
        Logger.log("LAGER sheet not found.");
        return;
    }
    var lagerData = readSheetData_(lager);
    var lagerRegale = getRegalColumnsFromValues_(lagerData.values);
    var lagerLocations = getRegalIdLocations_(lagerData.values);

    Logger.log("Recognized LAGER regals: " + Object.keys(lagerRegale).join(", "));

    var entries = collectEntries_();
    var withRegal = 0;
    var wouldPlace = 0;
    var skipNoRegal = 0;
    var dupSameRegal = 0;
    var dupWrongRegal = 0;
    var skipNoHeader = 0;
    var skipNoSlot = 0;

    for (var i = 0; i < entries.length; i++) {
        var status = entries[i].status;
        var regalKey = parseRegalStatus_(status);
        if (!regalKey) {
            skipNoRegal++;
            continue;
        }
        withRegal++;
        var idn = normalizeId_(entries[i].id);

        var loc = lagerLocations[idn];
        if (loc) {
            if (loc.regal === regalKey) {
                dupSameRegal++;
            } else {
                dupWrongRegal++;
                Logger.log("WRONG REGAL: id=" + entries[i].id + " currently in " + loc.regal + " but Refurbishment says " + regalKey);
            }
            continue;
        }

        var pos = lagerRegale[regalKey];
        if (!pos) {
            skipNoHeader++;
            Logger.log("SKIP no header in LAGER: id=" + entries[i].id + " status='" + status + "' regal=" + regalKey);
            continue;
        }

        var col0 = pos.col - 1;
        var headerRow0 = pos.row - 1;
        var endRow0 = getRegalEndRow0_(lagerData.values, headerRow0, col0);
        var row = findFirstEmptyRow_(lagerData.values, headerRow0, col0, endRow0);
        var row0 = row - 1;
        var blocked = (row0 <= headerRow0) || (row0 >= endRow0 && endRow0 < lagerData.values.length);
        if (blocked) {
            skipNoSlot++;
            Logger.log("SKIP no free slot: id=" + entries[i].id + " status='" + status + "' regal=" + regalKey);
            continue;
        }
        wouldPlace++;
    }

    Logger.log("---- SUMMARY ----");
    Logger.log("total source entries: " + entries.length);
    Logger.log("with valid Regal: " + withRegal + " | without parseable Regal: " + skipNoRegal);
    Logger.log("would be placed (new): " + wouldPlace);
    Logger.log("already in correct regal: " + dupSameRegal);
    Logger.log("present but in WRONG regal (would need move): " + dupWrongRegal);
    Logger.log("skipped missing header: " + skipNoHeader);
    Logger.log("skipped no free slot: " + skipNoSlot);
}

function getRegalIdLocations_(values) {
    var map = {};
    var numRows = values.length;
    var numCols = numRows ? values[0].length : 0;
    for (var c = 0; c < numCols; c++) {
        for (var r = 0; r < numRows; r++) {
            if (!isRegalHeader_(values[r][c])) continue;
            var regalKey = normalizeRegal_(values[r][c]);
            var endRow0 = getRegalEndRow0_(values, r, c);
            for (var rr = r + 1; rr < endRow0; rr++) {
                var v = values[rr][c];
                if (v === "" || v === null) continue;
                if (isRegalHeader_(v)) break;
                var idn = normalizeId_(v);
                if (!map.hasOwnProperty(idn)) {
                    map[idn] = { regal: regalKey, row: rr + 1, col: c + 1 };
                }
            }
        }
    }
    return map;
}

function debugStockId(stockId) {
    var target = normalizeId_(stockId);
    var entries = collectEntries_();
    var found = null;
    for (var i = 0; i < entries.length; i++) {
        if (normalizeId_(entries[i].id) === target) {
            found = entries[i];
            break;
        }
    }
    if (!found) {
        Logger.log("[" + stockId + "] not found in Refurbishment/Nachbestellung sources.");
        return;
    }
    Logger.log("[" + stockId + "] source status = '" + found.status + "'");

    var regalKey = parseRegalStatus_(found.status);
    Logger.log("parseRegalStatus_ => " + regalKey);
    Logger.log("isTageslisteStatus_ => " + isTageslisteStatus_(found.status));

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var lager = ss.getSheetByName(CONFIG.lagerSheetName);
    if (!lager) {
        Logger.log("LAGER sheet not found.");
        return;
    }
    var lagerData = readSheetData_(lager);
    var lagerRegale = getRegalColumnsFromValues_(lagerData.values);
    var lagerIds = getRegalIdsFromValues_(lagerData.values);

    Logger.log("Already present in LAGER? " + (lagerIds[target] ? "YES => skipped by duplicate check" : "no"));

    if (!regalKey) {
        Logger.log("Status does not parse to a valid Regal (range 1.1-5.8, must contain word 'regal').");
        return;
    }

    var pos = lagerRegale[regalKey];
    if (!pos) {
        Logger.log("Regal header '" + regalKey + "' NOT recognized in LAGER => box can never be placed.");
        Logger.log("Recognized LAGER regals: " + Object.keys(lagerRegale).join(", "));
        return;
    }

    var col0 = pos.col - 1;
    var headerRow0 = pos.row - 1;
    var endRow0 = getRegalEndRow0_(lagerData.values, headerRow0, col0);
    var row = findFirstEmptyRow_(lagerData.values, headerRow0, col0, endRow0);
    var row0 = row - 1;
    var blocked = (row0 <= headerRow0) || (row0 >= endRow0 && endRow0 < lagerData.values.length);
    Logger.log("Regal '" + regalKey + "' header at row " + pos.row + ", col " + pos.col +
        ", blockEndsBeforeRow " + (endRow0 + 1) + ", firstEmptyRow " + row +
        (blocked ? " => NO FREE SLOT, box skipped" : " => free slot found, box would be placed"));
}
