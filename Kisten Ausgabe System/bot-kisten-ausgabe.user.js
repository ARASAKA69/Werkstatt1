// ==UserScript==
// @name         ARASAKA Master-Bot (Upload)
// @namespace    http://tampermonkey.net/
// @version      1.37
// @description  Live-Version
// @author       ARASAKA
// @match        *://carol.autohero.com/*
// @grant        GM_xmlhttpRequest
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @connect      googleusercontent.com
// ==/UserScript==

(function() {
    'use strict';

    const DRIVE_WEB_APP_URL = "https://script.google.com/a/macros/autohero.com/s/AKfycbz0yz1BdUx4ZXgT4V4rqfif8KM3D76rNDjWXY2DZD9JIP0D4y9cjsGsFooOZqaGlm1c/exec";
    const API_KEY = "ARASAKA_2026";
    const ARASAKA_DEBUG = true;

    function dbg() {
        if (!ARASAKA_DEBUG) return;
        var a = ['[ARASAKA]'];
        for (var i = 0; i < arguments.length; i++) a.push(arguments[i]);
        console.log.apply(console, a);
    }

    function bridgeBodyLooksLikeHtml(t) {
        var s = String(t || '').trim();
        return s.length > 0 && (s.slice(0, 9).toLowerCase() === '<!doctype' || s.slice(0, 5).toLowerCase() === '<html' || s.slice(0, 6).toLowerCase() === '<head>');
    }

    function bridgeHtmlErrorHint(html) {
        var h = String(html || '');
        var out = '';
        var mt = h.match(/<title[^>]*>([^<]+)<\/title>/i);
        if (mt) out = mt[1].trim();
        var me = h.match(/class="errorMessage"[^>]*>([^<]+)/i);
        if (me) out = (out ? out + ' — ' : '') + me[1].trim();
        if (!out) out = h.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 120);
        return out;
    }

    function bridgeAppsScriptHtmlPopupText() {
        return 'Die Web-App antwortet mit einer Google-Fehlerseite (HTML), nicht mit JSON.\n\n'
            + 'Typisch: Deployment-Zugriff, fehlende Autorisierung, oder falscher Google-Account.\n\n'
            + '1) Im Browser mit dem Workspace-Konto einloggen (z. B. @autohero.com).\n'
            + '2) Apps Script öffnen → Deploy → Verwaltung → neue Version, Zugriff z. B. "Alle innerhalb der Domain" oder "Jeder".\n'
            + '3) Die /exec-URL einmal im gleichen Browser öffnen und Berechtigungen erlauben.\n'
            + '4) Falls im Skript eine E-Mail-Whitelist aktiv ist: actor-Parameter setzen oder Liste prüfen.\n'
            + '5) Neues Apps-Skript deployen (doPost muss enthalten sein — aktueller drive-bridge.gs).';
    }

    function bridgeExtraHintForDriveError(msg) {
        var m = String(msg || '');
        if (/angegebenen ID|specified ID|kein Element|not have the required permission|Berechtigung|permission/i.test(m)) {
            return '\n\nDrive: In Apps Script die Variable folderOffenId (Kisten-Ordner) prüfen — ID aus der Browser-URL des Ordners kopieren. Ordner für das Konto aus "Deploy → Als ausführen" freigeben.';
        }
        return '';
    }

    function bridgeGetJson(payload, timeoutMs) {
        var params = Object.assign({ key: API_KEY }, payload);
        var qs = Object.keys(params).map(function(k) {
            return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
        }).join('&');
        var url = DRIVE_WEB_APP_URL + '?' + qs;
        return new Promise(function(resolve) {
            GM_xmlhttpRequest({
                method: 'GET',
                url: url,
                timeout: timeoutMs || 15000,
                onload: function(r) { resolve(r); },
                onerror: function(err) {
                    dbg('bridgeGetJson', 'onerror', 'status', err && err.status, 'statusText', err && err.statusText, 'finalUrl', err && err.finalUrl, 'responseText', err && String(err.responseText || '').slice(0, 300));
                    resolve(null);
                },
                ontimeout: function(err) {
                    dbg('bridgeGetJson', 'ontimeout', 'status', err && err.status, 'finalUrl', err && err.finalUrl);
                    resolve(null);
                }
            });
        });
    }

    function bridgePostJson(payload, timeoutMs) {
        var body = JSON.stringify(Object.assign({ key: API_KEY }, payload));
        return new Promise(function(resolve) {
            GM_xmlhttpRequest({
                method: 'POST',
                url: DRIVE_WEB_APP_URL,
                headers: { 'Content-Type': 'application/json' },
                data: body,
                timeout: timeoutMs || 15000,
                onload: function(r) { resolve(r); },
                onerror: function(err) {
                    dbg('bridgePostJson', 'onerror', 'status', err && err.status, 'statusText', err && err.statusText, 'finalUrl', err && err.finalUrl, 'responseText', err && String(err.responseText || '').slice(0, 300));
                    resolve(null);
                },
                ontimeout: function(err) {
                    dbg('bridgePostJson', 'ontimeout', 'status', err && err.status, 'finalUrl', err && err.finalUrl);
                    resolve(null);
                }
            });
        });
    }

    let isProcessing = false;
    let abortMission = false;

    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            abortMission = true;
            isProcessing = false;
            sessionStorage.removeItem('arasaka_batch_data');
            sessionStorage.removeItem('arasaka_batch_keys');
            sessionStorage.removeItem('arasaka_batch_current_idx');
            showCustomPopup("ARASAKA", "PROZESS GESTOPPT & SPEICHER GELÖSCHT", true);
        }

        if (e.altKey && e.key.toLowerCase() === 'b') {
            e.preventDefault();
            if (!isProcessing) {
                isProcessing = true;
                abortMission = false;
                startBatchProcess();
            }
        }
    });

    setTimeout(() => {
        let keys = JSON.parse(sessionStorage.getItem('arasaka_batch_keys') || "[]");
        if (keys.length > 0) {
            isProcessing = true;
            abortMission = false;
            setTimeout(processNextStock, 1500);
        }
    }, 1500);

    function playSuccessSound() {
        try {
            let audio = new Audio("https://actions.google.com/sounds/v1/alarms/beep_short.ogg");
            audio.volume = 0.8;
            audio.play().catch(function(e) { dbg('audio', 'autoplay', e); });
        } catch (e) {}
    }

    function forceClick(el) {
        if (!el) return;
        el.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
        el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
        el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        try { el.click(); } catch(e) {}
    }

    async function waitForElementByText(texts, selector = '*', timeout = 15000) {
        const start = Date.now();
        while (Date.now() - start < timeout) {
            if (abortMission) return null;
            const elements = Array.from(document.querySelectorAll(selector));
            const el = elements.find(e => texts.some(t => e.textContent && e.textContent.includes(t)) && !e.disabled);
            if (el) return el;
            await new Promise(r => setTimeout(r, 200));
        }
        return null;
    }

    async function waitForExactText(texts, selector = 'button', timeout = 10000) {
        const start = Date.now();
        while (Date.now() - start < timeout) {
            if (abortMission) return null;
            const elements = Array.from(document.querySelectorAll(selector));
            const el = elements.reverse().find(e => texts.some(t => e.textContent.trim() === t) && !e.disabled);
            if (el) return el;
            await new Promise(r => setTimeout(r, 200));
        }
        return null;
    }

    async function waitForText(text, timeout = 15000) {
        const start = Date.now();
        const searchStr = text.toLowerCase();
        while (Date.now() - start < timeout) {
            if (abortMission) return false;
            if (document.body.innerText.toLowerCase().includes(searchStr)) return true;
            await new Promise(r => setTimeout(r, 500));
        }
        return false;
    }

    function basenameLower(path) {
        const s = String(path || '').replace(/\\/g, '/');
        const i = s.lastIndexOf('/');
        return (i >= 0 ? s.slice(i + 1) : s).toLowerCase();
    }

    function isDocumentFileNameAlreadyOnPage(fileName) {
        const base = basenameLower(fileName);
        if (!base) return false;
        return (document.body.innerText || '').toLowerCase().includes(base);
    }

    function classifyUploadFile(nameLower) {
        const base = basenameLower(nameLower);
        if (base.includes('retoure')) return 'retoure';
        if (/(?:^|[\s_-])na(?:\s*\(\d+\))?\.[a-z0-9]{2,5}$/i.test(base)) return 'nachbestellung';
        if (/ na\b/.test(base)) return 'nachbestellung';
        return 'ausgabe';
    }

    function dedupeFilesByName(files) {
        const byKey = new Map();
        const duplicates = [];
        for (let i = 0; i < files.length; i++) {
            const f = files[i];
            const key = basenameLower(f.name);
            const mt = f.modifiedTime != null ? Number(f.modifiedTime) : 0;
            if (!byKey.has(key)) {
                byKey.set(key, f);
                continue;
            }
            const prev = byKey.get(key);
            const prevMt = prev.modifiedTime != null ? Number(prev.modifiedTime) : 0;
            if (mt > prevMt) {
                duplicates.push({ file: prev, firstName: prev.name, reason: 'batch_older_mtime' });
                byKey.set(key, f);
            } else if (mt < prevMt) {
                duplicates.push({ file: f, firstName: prev.name, reason: 'batch_older_mtime' });
            } else {
                duplicates.push({ file: f, firstName: prev.name, reason: 'batch_same_mtime' });
            }
        }
        const unique = Array.from(byKey.values());
        return { unique, duplicates };
    }

    function moveFileRequest(fileInfo, isRetoure, extraParams) {
        const maxParamLen = 450;
        let payload = { action: 'moveFile', fileId: fileInfo.id, isRetoure: isRetoure };
        if (extraParams && typeof extraParams === 'object') {
            for (const [k, v] of Object.entries(extraParams)) {
                if (v === true || v === false) {
                    payload[k] = v;
                } else if (v != null && v !== '') {
                    let s = String(v);
                    if (s.length > maxParamLen) s = s.slice(0, maxParamLen - 3) + '...';
                    payload[k] = s;
                }
            }
        }
        return bridgePostJson(payload, 10000).then(function(r) {
            if (!r) { dbg('moveFile', 'noResponse', fileInfo.id); return r; }
            dbg('moveFile', 'http', r.status, 'id', fileInfo.id, 'body', String(r.responseText || '').slice(0, 160));
            return r;
        });
    }

    function sleep(ms) {
        return new Promise(resolve => {
            let elapsed = 0;
            let interval = setInterval(() => {
                if (abortMission) {
                    clearInterval(interval);
                    resolve();
                }
                elapsed += 100;
                if (elapsed >= ms) {
                    clearInterval(interval);
                    resolve();
                }
            }, 100);
        });
    }

    async function findSearchBar() {
        let start = Date.now();
        while (Date.now() - start < 10000) {
            if (abortMission) return null;
            let exactMatch = document.querySelector('[data-qa-selector="vin_stock_number_filter"]');
            if (exactMatch) return exactMatch;
            let contentEditables = Array.from(document.querySelectorAll('div[contenteditable="true"]'));
            if (contentEditables.length > 0) return contentEditables[0];
            let inputs = Array.from(document.querySelectorAll('input:not([type="hidden"]):not([type="checkbox"]):not([type="radio"]):not([type="file"])'));
            let visibleInputs = inputs.filter(i => i.offsetWidth > 0 && i.offsetHeight > 0);
            if (visibleInputs.length > 0) return visibleInputs[0];
            await new Promise(r => setTimeout(r, 500));
        }
        return null;
    }

    async function handleStockError(stockId, idx, message) {
        dbg('handleStockError', 'stockId', stockId, 'idx', idx, 'message', message);
        showCustomPopup("ARASAKA FEHLER", `${message}\nVerschiebe in Kisten Falsche Stock Ordner...`, false);
        let allData = JSON.parse(sessionStorage.getItem('arasaka_batch_data'));
        let files = allData[stockId];
        let movePromises = files.map(f => bridgePostJson({
            action: 'moveFileError',
            fileId: f.id,
            stockId: stockId,
            reason: message
        }, 10000).then(function(r) {
            if (!r) { dbg('moveFileError', 'noResponse', f.id); return; }
            dbg('moveFileError', 'http', r.status, 'fileId', f.id, 'body', String(r.responseText || '').slice(0, 120));
        }));
        await Promise.all(movePromises);
        sessionStorage.setItem('arasaka_batch_current_idx', (idx + 1).toString());
        window.location.href = '/';
    }

    async function startBatchProcess() {
        dbg('startBatchProcess');
        showCustomPopup("ARASAKA ONLINE", "Prüfe Kisten im Google Drive...", false);
        if (abortMission) return;
        let response = await bridgeGetJson({ action: 'getBatch' }, 30000);
        if (!response) {
            dbg('getBatch', 'noResponse');
            showCustomPopup("FEHLER", "Keine Verbindung zu Google Drive möglich.", true);
            isProcessing = false;
            return;
        }
        var rt = response.responseText || '';
        dbg('getBatch', 'http', response.status, 'finalUrl', response.finalUrl || '', 'len', rt.length, 'head', rt.slice(0, 240));
        if (bridgeBodyLooksLikeHtml(rt)) {
            dbg('getBatch', 'appsScriptHtml', bridgeHtmlErrorHint(rt));
            dbg('getBatch', 'appsScriptHtmlHelp', bridgeAppsScriptHtmlPopupText());
            showCustomPopup("FEHLER", "Die Verbindung zur Google-Web-App ist fehlgeschlagen (HTML statt Daten). Bei Bedarf Konsole (F12) für Details.", true);
            isProcessing = false;
            return;
        }
        var parsedErr = null;
        try { parsedErr = JSON.parse(rt); } catch (e0) {}
        if (parsedErr && typeof parsedErr.error === 'string') {
            dbg('getBatch', 'serverError', parsedErr.error, bridgeExtraHintForDriveError(parsedErr.error));
            showCustomPopup("FEHLER", "Server- oder Drive-Fehler. Konsole (F12) für den genauen Text.", true);
            isProcessing = false;
            return;
        }
        try {
            let data = JSON.parse(rt);
            let stockIds = Object.keys(data);
            dbg('getBatch', 'parsed', 'stockCount', stockIds.length, 'keys', stockIds);
            if (stockIds.length === 0) {
                playSuccessSound();
                showCustomPopup("Mahlzeit!", "Kisten Offen ist leer. Keine neuen Bilder gefunden.\n\n[ALT + B] drücken, um später neu zu scannen.\nZeit, den Falsche Stock-Ordner zu checken!", true);
                isProcessing = false;
                return;
            }
            sessionStorage.setItem('arasaka_batch_data', JSON.stringify(data));
            sessionStorage.setItem('arasaka_batch_keys', JSON.stringify(stockIds));
            sessionStorage.setItem('arasaka_batch_current_idx', '0');
            processNextStock();
        } catch (err) {
            dbg('getBatch', 'parseError', err && err.message, 'rawHead', rt.slice(0, 400));
            showCustomPopup("FEHLER", "Google Drive antwortet nicht richtig. Skript-URL und API-Key prüfen!", true);
            isProcessing = false;
        }
    }

    async function processNextStock() {
        if (abortMission) return;
        let keys = JSON.parse(sessionStorage.getItem('arasaka_batch_keys') || "[]");
        let idx = parseInt(sessionStorage.getItem('arasaka_batch_current_idx') || "0");
        dbg('processNextStock', 'idx', idx, 'totalKeys', keys.length, 'keys', keys);

        if (idx >= keys.length) {
            sessionStorage.removeItem('arasaka_batch_data');
            sessionStorage.removeItem('arasaka_batch_keys');
            sessionStorage.removeItem('arasaka_batch_current_idx');
            showCustomPopup("ARASAKA", "Stapel fertig. Kurzer Check im Drive...", false);
            (async function() {
                if (abortMission) return;
                let response = await bridgeGetJson({ action: 'getBatch' }, 30000);
                if (!response) {
                    dbg('getBatchRecheck', 'noResponse');
                    isProcessing = false;
                    return;
                }
                var rt = response.responseText || '';
                dbg('getBatchRecheck', 'http', response.status, 'finalUrl', response.finalUrl || '', 'len', rt.length, 'head', rt.slice(0, 240));
                if (bridgeBodyLooksLikeHtml(rt)) {
                    dbg('getBatchRecheck', 'appsScriptHtml', bridgeHtmlErrorHint(rt));
                    dbg('getBatchRecheck', 'appsScriptHtmlHelp', bridgeAppsScriptHtmlPopupText());
                    showCustomPopup("FEHLER", "Die Verbindung zur Google-Web-App ist fehlgeschlagen (HTML statt Daten). Bei Bedarf Konsole (F12) für Details.", true);
                    isProcessing = false;
                    return;
                }
                var parsedErr2 = null;
                try { parsedErr2 = JSON.parse(rt); } catch (e1) {}
                if (parsedErr2 && typeof parsedErr2.error === 'string') {
                    dbg('getBatchRecheck', 'serverError', parsedErr2.error, bridgeExtraHintForDriveError(parsedErr2.error));
                    showCustomPopup("FEHLER", "Server- oder Drive-Fehler. Konsole (F12) für den genauen Text.", true);
                    isProcessing = false;
                    return;
                }
                try {
                    let data = JSON.parse(rt);
                    let stockIds = Object.keys(data);
                    dbg('getBatchRecheck', 'parsed', 'stockCount', stockIds.length);
                    if (stockIds.length === 0) {
                        playSuccessSound();
                        showCustomPopup("Mahlzeit!", "Alle Bilder sind sauber hochgeladen!\n\nKeine Kisten mehr da.\n\n[ALT + B] für den nächsten Scan.", true);
                        isProcessing = false;
                    } else {
                        sessionStorage.setItem('arasaka_batch_data', JSON.stringify(data));
                        sessionStorage.setItem('arasaka_batch_keys', JSON.stringify(stockIds));
                        sessionStorage.setItem('arasaka_batch_current_idx', '0');
                        processNextStock();
                    }
                } catch (err) {
                    dbg('getBatchRecheck', 'parseError', err && err.message);
                    showCustomPopup("FEHLER", "Fehler beim Auto-Recheck.", true);
                    isProcessing = false;
                }
            })();
            return;
        }

        let stockId = keys[idx].toUpperCase();
        showCustomPopup("ARASAKA LÄUFT", `Suche nach Stock ID: ${stockId}...`, false);

        let searchInput = await findSearchBar();
        if (!searchInput) {
            dbg('findSearchBar', 'null');
            await handleStockError(stockId, idx, "Suchleiste auf Startseite nicht gefunden");
            return;
        }

        searchInput.style.transition = "all 0.2s";
        searchInput.style.border = "3px solid red";
        searchInput.focus();

        if (searchInput.isContentEditable) {
            document.execCommand('selectAll', false, null);
            document.execCommand('insertText', false, stockId);
        } else {
            let nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value")?.set;
            if (nativeInputValueSetter) {
                nativeInputValueSetter.call(searchInput, stockId);
            } else {
                searchInput.value = stockId;
            }
        }

        searchInput.dispatchEvent(new Event('input', { bubbles: true }));
        searchInput.dispatchEvent(new Event('change', { bubbles: true }));
        await sleep(300);
        searchInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
        searchInput.dispatchEvent(new KeyboardEvent('keypress', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
        searchInput.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
        setTimeout(() => { if(searchInput) searchInput.style.border = ""; }, 1000);

        showCustomPopup("ARASAKA LÄUFT", `Warte auf Suchergebnis für ${stockId}...`, false);
        await sleep(2000);

        if (abortMission) return;

        let resultRow = null;
        for (let attempt = 0; attempt < 25; attempt++) {
            if (abortMission) return;
            let links = Array.from(document.querySelectorAll('a[href*="/refurbishment/"]'));
            let validLinks = links.filter(l => l.href.match(/[0-9a-f]{8}-[0-9a-f]{4}/i));
            if (validLinks.length > 0) { resultRow = validLinks[0]; break; }
            else if (links.length > 0) { resultRow = links[0]; break; }

            let rows = Array.from(document.querySelectorAll('tr.clickable-row, .rt-tr-group, tr[data-test-id*="row"]'));
            let validRows = rows.filter(r => !r.querySelector('th'));
            if (validRows.length > 0) { resultRow = validRows[0]; break; }

            let isNoData = Array.from(document.querySelectorAll('div, span, td')).some(el =>
                ['no data', 'keine daten', 'no results', 'keine ergebnisse', '0 results'].some(t => el.textContent.toLowerCase().trim() === t)
            );
            if (isNoData && attempt > 5) {
                break;
            }
            await sleep(200);
        }

        if (abortMission) return;

        if (resultRow) {
            dbg('searchResult', 'rowFound', true);
            showCustomPopup("ARASAKA NAVIGATION", `Öffne Auftrag...`, false);
            forceClick(resultRow);
            let uploadReady = await waitForElementByText(['Upload Document', 'Dokument hochladen'], 'button', 15000);
            if (abortMission) return;

            if (uploadReady) {
                showCustomPopup("ARASAKA VERIFIKATION", `Prüfe, ob Auftrag ${stockId} der richtige ist...`, false);
                let isCorrectPage = await waitForText(stockId, 5000);
                dbg('pageVerify', 'stockId', stockId, 'isCorrectPage', isCorrectPage);
                if (isCorrectPage) {
                    executeUploadsForStock(stockId);
                } else {
                    await handleStockError(stockId, idx, "Falscher Auftrag geladen (Stock-ID fehlt im Auftrag)");
                }
            } else {
                dbg('uploadButton', 'missing');
                await handleStockError(stockId, idx, "Auftrag hat nicht geöffnet oder Upload Button fehlt");
            }
        } else {
            dbg('searchResult', 'rowFound', false);
            await handleStockError(stockId, idx, "Auftrag nicht gefunden");
        }
    }

    async function executeUploadsForStock(stockId) {
        let allData = JSON.parse(sessionStorage.getItem('arasaka_batch_data'));
        let rawFiles = allData[stockId];
        let idx = parseInt(sessionStorage.getItem('arasaka_batch_current_idx') || "0");

        let deduped = dedupeFilesByName(rawFiles);
        let files = deduped.unique;
        let dupByName = deduped.duplicates;
        dbg('executeUploadsForStock', 'stockId', stockId, 'batchIdx', idx, 'rawCount', rawFiles.length, 'uniqueCount', files.length, 'dupCount', dupByName.length);

        for (let d = 0; d < dupByName.length; d++) {
            if (abortMission) return;
            let dup = dupByName[d];
            let fn = dup.file.name;
            let cls = classifyUploadFile(fn);
            let isRetoureDup = cls === 'retoure';
            dbg('dupFilename', fn, 'first', dup.firstName, 'reason', dup.reason);
            showCustomPopup("ARASAKA SKIP", "Doppelter Dateiname in dieser Ladung — wird übersprungen und verschoben.", false);
            await moveFileRequest(dup.file, isRetoureDup, {
                toDuplicate: true,
                logKind: 'skip_duplicate_filename',
                logStockId: stockId,
                logFileName: fn,
                logDetail: (dup.reason || 'same_name') + '_vs_' + dup.firstName
            });
        }

        let totalFiles = files.length;
        let pageText = document.body.innerText;
        let existingAusgabe = (pageText.match(/Ausgabe \d+\/\d+/g) || []).length;
        let existingRetoure = (pageText.match(/Retoure \d+\/\d+/g) || []).length;
        let existingNachbestellung = (pageText.match(/Nachbestellung \d+\/\d+/g) || []).length;

        let newAusgabe = files.filter(f => classifyUploadFile(f.name) === 'ausgabe').length;
        let newRetoure = files.filter(f => classifyUploadFile(f.name) === 'retoure').length;
        let newNachbestellung = files.filter(f => classifyUploadFile(f.name) === 'nachbestellung').length;

        let ausgabeTotal = existingAusgabe + newAusgabe;
        let retoureTotal = existingRetoure + newRetoure;
        let nachbestellungTotal = existingNachbestellung + newNachbestellung;
        let ausgabeIdx = existingAusgabe + 1;
        let retoureIdx = existingRetoure + 1;
        let nachbestellungIdx = existingNachbestellung + 1;
        let skipCommentCount = 0;
        let skipFilenameOnPageCount = 0;
        dbg('commentPlan', 'existingA/R/N', existingAusgabe, existingRetoure, existingNachbestellung, 'newA/R/N', newAusgabe, newRetoure, newNachbestellung, 'totalA/R/N', ausgabeTotal, retoureTotal, nachbestellungTotal);

        for (let i = 0; i < totalFiles; i++) {
            if (abortMission) return;

            let fileInfo = files[i];
            let cls = classifyUploadFile(fileInfo.name);
            let isRetoure = cls === 'retoure';
            let isNachbestellung = cls === 'nachbestellung';

            let currentComment;
            if (isRetoure) {
                currentComment = `Retoure ${retoureIdx}/${retoureTotal}`;
            } else if (isNachbestellung) {
                currentComment = `Nachbestellung ${nachbestellungIdx}/${nachbestellungTotal}`;
            } else {
                currentComment = `Ausgabe ${ausgabeIdx}/${ausgabeTotal}`;
            }

            const mt = fileInfo.modifiedTime != null ? Number(fileInfo.modifiedTime) : 0;
            const stored = fileInfo.lastUploadedStored != null ? Number(fileInfo.lastUploadedStored) : null;

            if (stored != null && mt === stored) {
                skipFilenameOnPageCount++;
                dbg('skipSameMtimeAsStored', fileInfo.name, 'mt', mt, 'stored', stored);
                showCustomPopup("ARASAKA SKIP", "Gleicher Zeitstempel wie beim letzten Upload — Duplikat, übersprungen.", false);
                await moveFileRequest(fileInfo, isRetoure, {
                    toDuplicate: true,
                    logKind: 'skip_same_mtime_as_stored',
                    logStockId: stockId,
                    logFileName: fileInfo.name,
                    logDetail: currentComment
                });
                continue;
            }

            if (stored != null && mt < stored) {
                skipFilenameOnPageCount++;
                dbg('skipOlderThanStored', fileInfo.name, 'mt', mt, 'stored', stored);
                showCustomPopup("ARASAKA SKIP", "Datei ist älter als der letzte Upload — übersprungen.", false);
                await moveFileRequest(fileInfo, isRetoure, {
                    toDuplicate: true,
                    logKind: 'skip_older_than_stored',
                    logStockId: stockId,
                    logFileName: fileInfo.name,
                    logDetail: currentComment
                });
                continue;
            }

            if (stored == null && isDocumentFileNameAlreadyOnPage(fileInfo.name)) {
                skipFilenameOnPageCount++;
                dbg('skipFilenameOnPageNoStored', fileInfo.name, 'planned', currentComment);
                showCustomPopup("ARASAKA SKIP", `Dateiname ${fileInfo.name} steht schon in der Dokumentenliste. Überspringe...`, false);
                await moveFileRequest(fileInfo, isRetoure, {
                    toDuplicate: true,
                    logKind: 'skip_filename_on_page',
                    logStockId: stockId,
                    logFileName: fileInfo.name,
                    logDetail: currentComment
                });
                continue;
            }

            if (stored != null && mt > stored && isDocumentFileNameAlreadyOnPage(fileInfo.name)) {
                dbg('uploadNewerDespiteFilenameOnPage', fileInfo.name, 'mt', mt, 'stored', stored);
            }

            if (document.body.innerText.includes(currentComment)) {
                skipCommentCount++;
                dbg('skipExistingComment', currentComment, fileInfo.name);
                showCustomPopup("ARASAKA SKIP", "Dieser Kommentar existiert bereits — übersprungen.", false);
                if (isRetoure) retoureIdx++;
                else if (isNachbestellung) nachbestellungIdx++;
                else ausgabeIdx++;
                await moveFileRequest(fileInfo, isRetoure, {
                    toDuplicate: true,
                    logKind: 'skip_comment_exists',
                    logStockId: stockId,
                    logFileName: fileInfo.name,
                    logDetail: currentComment
                });
                continue;
            }

            if (isRetoure) retoureIdx++;
            else if (isNachbestellung) nachbestellungIdx++;
            else ausgabeIdx++;

            dbg('uploadStep', i + 1, totalFiles, cls, currentComment, fileInfo.name, 'stockId', stockId);
            showCustomPopup("ARASAKA DOWNLOAD", "Lade Bild " + (i + 1) + " von " + totalFiles + " aus Drive...", false);
            let b64 = await new Promise((resolve) => {
                bridgePostJson({ action: 'getFileData', fileId: fileInfo.id }, 45000).then(function(res) {
                    if (!res) { dbg('getFileData', 'noResponse', fileInfo.id); resolve(null); return; }
                    dbg('getFileData', 'http', res.status, 'fileId', fileInfo.id, 'b64len', (res.responseText || '').length);
                    resolve(res.responseText);
                });
            });

            if (!b64 || bridgeBodyLooksLikeHtml(b64)) {
                dbg('getFileData', 'emptyOrHtml', fileInfo.id);
                await handleStockError(stockId, idx, `Bild ${i+1} konnte nicht geladen werden`);
                return;
            }

            showCustomPopup("ARASAKA UPLOAD", "Lade Bild " + (i + 1) + " hoch...", false);

            let uploadBtn = await waitForElementByText(['Upload Document', 'Dokument hochladen'], 'button', 10000);
            if (abortMission) return;
            if (uploadBtn) forceClick(uploadBtn);

            let selectType = await waitForElementByText(['Other', 'Andere'], 'select, option', 5000);
            if (abortMission) return;

            let selectElement = document.querySelector('select');
            if (selectElement) {
                let options = Array.from(selectElement.options);
                let otherOpt = options.find(o => o.text === 'Other' || o.text === 'Andere');
                if (otherOpt) {
                    selectElement.value = otherOpt.value;
                    selectElement.dispatchEvent(new Event('change', { bubbles: true }));
                }
            }

            let commentInput = document.querySelector('input[name="comment"], input[placeholder*="Comment"]');
            if (!commentInput) {
                await sleep(1000);
                commentInput = document.querySelector('input[name="comment"], input[placeholder*="Comment"]');
            }
            if (abortMission) return;
            if (commentInput) {
                let nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value")?.set;
                if(nativeInputValueSetter) {
                    nativeInputValueSetter.call(commentInput, currentComment);
                } else {
                    commentInput.value = currentComment;
                }
                commentInput.dispatchEvent(new Event('input', { bubbles: true }));
            }

            let fileInput = document.querySelector('input[type="file"]');
            if (!fileInput) {
                await sleep(1000);
                fileInput = document.querySelector('input[type="file"]');
            }
            if (abortMission) return;
            if (fileInput) {
                let blob = b64toBlob(b64, fileInfo.mimeType);
                let file = new File([blob], fileInfo.name, { type: fileInfo.mimeType });
                let dataTransfer = new DataTransfer();
                dataTransfer.items.add(file);
                fileInput.files = dataTransfer.files;
                fileInput.dispatchEvent(new Event('change', { bubbles: true }));
            }

            await sleep(1000);

            let submitBtn = await waitForExactText(['Upload', 'Hochladen', 'Save', 'Speichern', 'Add', 'Hinzufügen'], 'button', 10000);
            if (abortMission) return;
            if (submitBtn) forceClick(submitBtn);

            showCustomPopup("ARASAKA VERIFIKATION", `Warte auf Bestätigung im System...`, false);
            let verifySuccess = await waitForText(currentComment, 20000);
            dbg('postUploadVerify', 'comment', currentComment, 'verifySuccess', verifySuccess);
            if (abortMission) return;

            if(!verifySuccess) {
                dbg('postUploadVerify', 'noMatchExtraWait');
                await sleep(2000);
            }

            if (abortMission) return;
            await moveFileRequest(fileInfo, isRetoure, {
                recordLastUpload: true,
                logStockId: stockId,
                logFileName: fileInfo.name
            });
        }

        if (abortMission) return;

        showCustomPopup("ARASAKA SYNC", `Setze Haken in der Tagesliste...`, false);

        let syncDone = false;
        let warningTimer = setTimeout(() => {
            if (!syncDone && !abortMission) {
                showCustomPopup("ARASAKA WARTET", "Warte noch kurz...\nGoogle braucht gerade etwas länger für den Haken. Geht gleich weiter!", false);
            }
        }, 5000);

        dbg('markSheet', 'post', 'skippedDup', dupByName.length, 'skippedComment', skipCommentCount, 'skippedFilenamePage', skipFilenameOnPageCount);
        let markRes = await bridgePostJson({
            action: 'markSheet',
            stockId: stockId,
            skippedDup: String(dupByName.length),
            skippedComment: String(skipCommentCount),
            skippedFilenamePage: String(skipFilenameOnPageCount),
            batchFiles: String(rawFiles.length),
            uniqueFiles: String(files.length)
        }, 45000);
        let syncStatus = 'HTTP_ERROR';
        if (markRes) {
            var mrt = String(markRes.responseText || '').trim();
            dbg('markSheet', 'http', markRes.status, 'body', mrt.slice(0, 200));
            if (bridgeBodyLooksLikeHtml(mrt)) {
                syncStatus = 'HTML_FEHLER';
                dbg('markSheet', 'appsScriptHtml', bridgeHtmlErrorHint(mrt));
            } else {
                syncStatus = mrt;
            }
        } else {
            dbg('markSheet', 'noResponse');
            syncStatus = 'HTTP_ERROR';
        }

        syncDone = true;
        clearTimeout(warningTimer);

        dbg('markSheet', 'syncStatus', syncStatus);

        if (syncStatus === "OK") {
            dbg('markSheetOk', stockId);
            showCustomPopup("ARASAKA", "Fertig — Tagesliste aktualisiert. Nächste Seite...", false);
            await sleep(1500);
        } else {
            dbg('markSheetHudWarn', 'stockId', stockId, 'syncStatus', syncStatus);
            showCustomPopup("ARASAKA WARNUNG", "Upload fertig, aber der Haken in der Tagesliste konnte nicht gesetzt werden. Konsole (F12) für Details. Nächste Seite...", false);
            await sleep(4000);
        }

        dbg('executeUploadsForStock', 'done', 'nextIdx', idx + 1);
        sessionStorage.setItem('arasaka_batch_current_idx', (idx + 1).toString());
        window.location.href = '/';
    }

    function b64toBlob(b64Data, contentType = '', sliceSize = 512) {
        const byteCharacters = atob(b64Data);
        const byteArrays = [];
        for (let offset = 0; offset < byteCharacters.length; offset += sliceSize) {
            const slice = byteCharacters.slice(offset, offset + sliceSize);
            const byteNumbers = new Array(slice.length);
            for (let i = 0; i < slice.length; i++) {
                byteNumbers[i] = slice.charCodeAt(i);
            }
            const byteArray = new Uint8Array(byteNumbers);
            byteArrays.push(byteArray);
        }
        return new Blob(byteArrays, {type: contentType});
    }

    function showCustomPopup(title, message, isEnd) {
        let existing = document.getElementById('arasaka-batch-popup');
        if (existing) existing.remove();
        let popup = document.createElement('div');
        popup.id = 'arasaka-batch-popup';
        popup.style.position = 'fixed';
        popup.style.top = '20px';
        popup.style.right = '20px';
        popup.style.backgroundColor = 'rgba(10, 10, 10, 0.95)';
        popup.style.border = isEnd ? '2px solid #00FF00' : '2px solid #00ffcc';
        popup.style.color = isEnd ? '#00FF00' : '#00ffcc';
        popup.style.padding = '20px';
        popup.style.borderRadius = '10px';
        popup.style.zIndex = '9999999';
        popup.style.fontFamily = 'monospace';
        popup.style.boxShadow = isEnd ? '0 0 20px #00FF00' : '0 0 20px #00ffcc';
        popup.style.whiteSpace = 'pre-wrap';
        popup.innerHTML = `
            <div style="font-size: 16px; font-weight: bold; margin-bottom: 10px;">[ ${title} ]</div>
            <div style="font-size: 14px; max-width: 300px; line-height: 1.5;">${message}</div>
            ${isEnd ? '<div style="margin-top: 15px; font-size: 12px; opacity: 0.7;">(Klick zum Schließen)</div>' : ''}
        `;
        if (isEnd) {
            popup.style.cursor = 'pointer';
            popup.onclick = () => popup.remove();
        }
        document.body.appendChild(popup);
    }
})();