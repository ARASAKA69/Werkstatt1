function setupAuth() {
    var ss = SpreadsheetApp.openById("1PuCLw8UmDjB_pBo_jCZ9rmSD3GJQESHzPoBVu_--MRo");
    var sheet = ss.getSheetByName("Tagesliste");
    return sheet.getRange("A1").getValue();
}

function appendTextLogInFolder_(folderId, fileName, line) {
    var folder = DriveApp.getFolderById(folderId);
    var logFiles = folder.getFilesByName(fileName);
    var now = Utilities.formatDate(new Date(), "Europe/Berlin", "dd.MM.yyyy HH:mm:ss");
    var logText = "[" + now + "] " + line + "\n";
    if (logFiles.hasNext()) {
        var logFile = logFiles.next();
        logFile.setContent(logFile.getBlob().getDataAsString() + logText);
    } else {
        folder.createFile(fileName, logText);
    }
}

function doGet(e) {
    var action = e.parameter.action;
    var key = e.parameter.key;
    if (key !== "ARASAKA_2026") return ContentService.createTextOutput("Zugriff verweigert!");

    var folderOffenId = "1e_U2mQlCyFR7KJIVNlnNFVVe7sAW9OL5";
    var folderErledigtId = "1HsVdfJSmMICYGSjRJu_QzOD8-4HxEGKH";
    var folderFehlerId = "1ej_qqIa_G5mLe8ZMRu1haxagziJKD3Jc";
    var folderRetoureId = "1KYgI1ZLG2x2OwLNoJw3eDCs_UBE6EZfs";
    var sheetId = "1PuCLw8UmDjB_pBo_jCZ9rmSD3GJQESHzPoBVu_--MRo";
    var kistenLogName = "Kisten_Ausgabe_Log.txt";

    if (action === "getBatch") {
        var folder = DriveApp.getFolderById(folderOffenId);
        var files = folder.getFiles();
        var data = {};
        while (files.hasNext()) {
            var file = files.next();
            var name = file.getName();
            var nameWithoutExt = name.split('.')[0].toUpperCase();
            var stockId = nameWithoutExt.split(/[-_ ]/)[0];
            if (!data[stockId]) data[stockId] = [];
            data[stockId].push({
                id: file.getId(),
                name: name,
                mimeType: file.getMimeType()
            });
        }
        return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "getFileData") {
        var fileIdGet = e.parameter.fileId;
        var fileGet = DriveApp.getFileById(fileIdGet);
        var b64 = Utilities.base64Encode(fileGet.getBlob().getBytes());
        return ContentService.createTextOutput(b64);
    }

    if (action === "moveFile") {
        var fileIdMove = e.parameter.fileId;
        var isRetoure = e.parameter.isRetoure === "true";
        var fileMove = DriveApp.getFileById(fileIdMove);
        var targetFolder = DriveApp.getFolderById(isRetoure ? folderRetoureId : folderErledigtId);
        fileMove.moveTo(targetFolder);
        var logKind = e.parameter.logKind;
        if (logKind) {
            var logStockId = e.parameter.logStockId || "";
            var logFileName = e.parameter.logFileName || "";
            var logDetail = e.parameter.logDetail || "";
            var logLine = "moveFile | " + logKind + " | stock=" + logStockId + " | file=" + logFileName + " | detail=" + logDetail + " | retoure=" + isRetoure;
            appendTextLogInFolder_(folderErledigtId, kistenLogName, logLine);
        }
        return ContentService.createTextOutput("OK");
    }

    if (action === "moveFileError") {
        var fileIdError = e.parameter.fileId;
        var stockIdError = e.parameter.stockId || "Unbekannt";
        var reason = e.parameter.reason || "grund unbekannt";
        var fileError = DriveApp.getFileById(fileIdError);
        var targetFolderError = DriveApp.getFolderById(folderFehlerId);
        fileError.moveTo(targetFolderError);
        var logName = "Fehler_Log.txt";
        var logFiles = targetFolderError.getFilesByName(logName);
        var now = Utilities.formatDate(new Date(), "Europe/Berlin", "dd.MM.yyyy HH:mm");
        var logText = stockIdError + " -> " + reason + " (" + now + ")\n";
        if (logFiles.hasNext()) {
            var logFile = logFiles.next();
            logFile.setContent(logFile.getBlob().getDataAsString() + logText);
        } else {
            targetFolderError.createFile(logName, logText);
        }
        return ContentService.createTextOutput("OK");
    }

    if (action === "markSheet") {
        try {
            var stockToMark = String(e.parameter.stockId).toUpperCase().replace(/\s+/g, '');
            var skippedDup = e.parameter.skippedDup || "0";
            var skippedComment = e.parameter.skippedComment || "0";
            var batchFiles = e.parameter.batchFiles || "0";
            var uniqueFiles = e.parameter.uniqueFiles || "0";
            var skipDupDetail = e.parameter.skipDupDetail || "";
            var skipCommentDetail = e.parameter.skipCommentDetail || "";

            var ss = SpreadsheetApp.openById(sheetId);
            var sheet = ss.getSheetByName("Tagesliste");
            if (!sheet) {
                appendTextLogInFolder_(folderErledigtId, kistenLogName, "markSheet | " + stockToMark + " | SHEET_NOT_FOUND | dup=" + skippedDup + " | comment_skip=" + skippedComment + " | batch=" + batchFiles + " | unique=" + uniqueFiles);
                return ContentService.createTextOutput("SHEET_NOT_FOUND");
            }

            var data = sheet.getDataRange().getValues();
            var matchRow = -1;
            var latestDate = -1;

            for (var i = 1; i < data.length; i++) {
                var rowStock = String(data[i][4] || "").toUpperCase().replace(/\s+/g, '');
                if (rowStock === stockToMark) {
                    var dateVal = data[i][1];
                    var timeMs = 0;
                    if (dateVal instanceof Date) {
                        timeMs = dateVal.getTime();
                    } else if (dateVal) {
                        var parts = String(dateVal).split('.');
                        if (parts.length === 3) {
                            timeMs = new Date(parts[2], parts[1] - 1, parts[0]).getTime();
                        }
                    }
                    if (timeMs >= latestDate) {
                        latestDate = timeMs;
                        matchRow = i + 1;
                    }
                }
            }

            var markResult = "STOCK_NOT_FOUND";
            if (matchRow !== -1) {
                sheet.getRange(matchRow, 14).setValue(true);
                markResult = "OK";
            }

            var markLogLine = "markSheet | " + stockToMark + " | " + markResult + " | dup_skipped=" + skippedDup + " | comment_skip=" + skippedComment + " | batch=" + batchFiles + " | unique=" + uniqueFiles;
            if (skipDupDetail) markLogLine += " | dup_detail=" + skipDupDetail;
            if (skipCommentDetail) markLogLine += " | comment_skip_detail=" + skipCommentDetail;
            appendTextLogInFolder_(folderErledigtId, kistenLogName, markLogLine);

            return ContentService.createTextOutput(matchRow !== -1 ? "OK" : "STOCK_NOT_FOUND");
        } catch (err) {
            appendTextLogInFolder_(folderErledigtId, kistenLogName, "markSheet | ERROR | " + err.message);
            return ContentService.createTextOutput("ERROR_" + err.message);
        }
    }

    return ContentService.createTextOutput("Invalid Action");
}
