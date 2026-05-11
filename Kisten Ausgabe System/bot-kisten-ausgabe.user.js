// ==UserScript==
// @name         ARASAKA Master-Bot (Upload)
// @namespace    http://tampermonkey.net/
// @version      1.41
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
    const ARASAKA_BOT_VERSION = "1.41";
    const ARASAKA_BRIDGE_VERSION = "15";
    const ARASAKA_HUD_POS_KEY = "arasaka_hud_position";

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

    function normalizeBase64Response(value) {
        var s = String(value || '').trim().replace(/\s+/g, '');
        if (/[-_]/.test(s) && /^[A-Za-z0-9_-]+={0,2}$/.test(s)) {
            s = s.replace(/-/g, '+').replace(/_/g, '/');
        }
        var mod = s.length % 4;
        if (mod > 1) s += new Array(5 - mod).join('=');
        return s;
    }

    function base64ResponseProblem(value) {
        var raw = String(value || '').trim();
        if (!raw) return 'EMPTY_RESPONSE';
        if (bridgeBodyLooksLikeHtml(raw)) return 'HTML_RESPONSE: ' + bridgeHtmlErrorHint(raw);
        if (raw.charAt(0) === '{') {
            try {
                var parsed = JSON.parse(raw);
                if (parsed && typeof parsed.error === 'string') return 'BRIDGE_ERROR: ' + parsed.error;
            } catch (e) {
                return 'JSON_PARSE_ERROR: ' + e.message;
            }
        }
        var b64 = normalizeBase64Response(raw);
        if (b64.length < 64) return 'BASE64_TOO_SHORT';
        if (b64.length % 4 === 1) return 'BASE64_BAD_LENGTH';
        if (!/^[A-Za-z0-9+/]*={0,2}$/.test(b64)) return 'BASE64_BAD_CHARS';
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

    async function moveFileRequest(fileInfo, isRetoure, extraParams) {
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
        for (let attempt = 1; attempt <= 3; attempt++) {
            let r = await bridgePostJson(payload, 20000);
            if (!r) {
                dbg('moveFile', 'noResponse', fileInfo.id, 'attempt', attempt);
            } else {
                let body = String(r.responseText || '').trim();
                dbg('moveFile', 'http', r.status, 'id', fileInfo.id, 'attempt', attempt, 'body', body.slice(0, 160));
                if (r.status >= 200 && r.status < 300 && body === 'OK') return r;
            }
            if (attempt < 3) await sleep(1000 * attempt);
        }
        return null;
    }

    async function moveFileOrStop(fileInfo, isRetoure, extraParams) {
        let moved = await moveFileRequest(fileInfo, isRetoure, extraParams);
        if (moved) return true;
        showCustomPopup("ARASAKA FEHLER", "Drive-Datei konnte nach dem Upload nicht verschoben werden. Stoppe, damit nichts doppelt hochgeladen wird. Konsole (F12) für Details.", true);
        isProcessing = false;
        return false;
    }

    function continueWithNextStock(nextIdx) {
        sessionStorage.setItem('arasaka_batch_current_idx', String(nextIdx));
        if (window.location.pathname === '/' || window.location.pathname === '') {
            setTimeout(processNextStock, 1000);
            return;
        }
        let target = window.location.origin + '/';
        window.location.assign(target);
        setTimeout(function() {
            if (!abortMission && window.location.pathname !== '/') window.location.replace(target);
        }, 4000);
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
        continueWithNextStock(idx + 1);
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
                    try {
                        await executeUploadsForStock(stockId);
                    } catch (err) {
                        dbg('executeUploadsForStock', 'unhandledError', err && err.message);
                        showCustomPopup("ARASAKA FEHLER", "Unerwarteter Fehler beim Upload. Stoppe, damit nichts doppelt hochgeladen wird. Konsole (F12) prüfen.", true);
                        isProcessing = false;
                    }
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
            if (!await moveFileOrStop(dup.file, isRetoureDup, {
                toDuplicate: true,
                logKind: 'skip_duplicate_filename',
                logStockId: stockId,
                logFileName: fn,
                logDetail: (dup.reason || 'same_name') + '_vs_' + dup.firstName
            })) return;
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
                if (!await moveFileOrStop(fileInfo, isRetoure, {
                    toDuplicate: true,
                    logKind: 'skip_same_mtime_as_stored',
                    logStockId: stockId,
                    logFileName: fileInfo.name,
                    logDetail: currentComment
                })) return;
                continue;
            }

            if (stored != null && mt < stored) {
                skipFilenameOnPageCount++;
                dbg('skipOlderThanStored', fileInfo.name, 'mt', mt, 'stored', stored);
                showCustomPopup("ARASAKA SKIP", "Datei ist älter als der letzte Upload — übersprungen.", false);
                if (!await moveFileOrStop(fileInfo, isRetoure, {
                    toDuplicate: true,
                    logKind: 'skip_older_than_stored',
                    logStockId: stockId,
                    logFileName: fileInfo.name,
                    logDetail: currentComment
                })) return;
                continue;
            }

            if (stored == null && isDocumentFileNameAlreadyOnPage(fileInfo.name)) {
                skipFilenameOnPageCount++;
                dbg('skipFilenameOnPageNoStored', fileInfo.name, 'planned', currentComment);
                showCustomPopup("ARASAKA SKIP", `Dateiname ${fileInfo.name} steht schon in der Dokumentenliste. Überspringe...`, false);
                if (!await moveFileOrStop(fileInfo, isRetoure, {
                    toDuplicate: true,
                    logKind: 'skip_filename_on_page',
                    logStockId: stockId,
                    logFileName: fileInfo.name,
                    logDetail: currentComment
                })) return;
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
                if (!await moveFileOrStop(fileInfo, isRetoure, {
                    toDuplicate: true,
                    logKind: 'skip_comment_exists',
                    logStockId: stockId,
                    logFileName: fileInfo.name,
                    logDetail: currentComment
                })) return;
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

            let b64Problem = base64ResponseProblem(b64);
            if (b64Problem) {
                dbg('getFileData', 'invalidBody', fileInfo.id, b64Problem, 'bodyHead', String(b64 || '').slice(0, 240));
                await handleStockError(stockId, idx, `Bild ${i+1} konnte nicht geladen werden`);
                return;
            }
            b64 = normalizeBase64Response(b64);

            showCustomPopup("ARASAKA UPLOAD", "Lade Bild " + (i + 1) + " hoch...", false);

            let uploadBtn = await waitForElementByText(['Upload Document', 'Dokument hochladen'], 'button', 10000);
            if (abortMission) return;
            if (!uploadBtn) {
                dbg('uploadStep', 'uploadButtonMissing', fileInfo.name);
                showCustomPopup("ARASAKA FEHLER", "Upload-Button nicht gefunden. Stoppe, damit nichts doppelt hochgeladen wird.", true);
                isProcessing = false;
                return;
            }
            forceClick(uploadBtn);

            await waitForElementByText(['Other', 'Andere'], 'select, option', 5000);
            if (abortMission) return;

            let selectElement = document.querySelector('select');
            if (!selectElement) {
                dbg('uploadStep', 'typeSelectMissing', fileInfo.name);
                showCustomPopup("ARASAKA FEHLER", "Upload-Auswahl nicht gefunden. Stoppe, damit nichts doppelt hochgeladen wird.", true);
                isProcessing = false;
                return;
            }
            let options = Array.from(selectElement.options);
            let otherOpt = options.find(o => o.text === 'Other' || o.text === 'Andere');
            if (otherOpt) {
                selectElement.value = otherOpt.value;
                selectElement.dispatchEvent(new Event('change', { bubbles: true }));
            }

            let commentInput = document.querySelector('input[name="comment"], input[placeholder*="Comment"]');
            if (!commentInput) {
                await sleep(1000);
                commentInput = document.querySelector('input[name="comment"], input[placeholder*="Comment"]');
            }
            if (abortMission) return;
            if (!commentInput) {
                dbg('uploadStep', 'commentInputMissing', fileInfo.name);
                showCustomPopup("ARASAKA FEHLER", "Kommentarfeld nicht gefunden. Stoppe, damit nichts doppelt hochgeladen wird.", true);
                isProcessing = false;
                return;
            }
            let nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value")?.set;
            if(nativeInputValueSetter) {
                nativeInputValueSetter.call(commentInput, currentComment);
            } else {
                commentInput.value = currentComment;
            }
            commentInput.dispatchEvent(new Event('input', { bubbles: true }));

            let fileInput = document.querySelector('input[type="file"]');
            if (!fileInput) {
                await sleep(1000);
                fileInput = document.querySelector('input[type="file"]');
            }
            if (abortMission) return;
            if (!fileInput) {
                dbg('uploadStep', 'fileInputMissing', fileInfo.name);
                showCustomPopup("ARASAKA FEHLER", "Dateifeld nicht gefunden. Stoppe, damit nichts doppelt hochgeladen wird.", true);
                isProcessing = false;
                return;
            }
            let blob;
            try {
                blob = b64toBlob(b64, fileInfo.mimeType);
            } catch (e2) {
                dbg('getFileData', 'decodeFailed', fileInfo.id, e2 && e2.message);
                await handleStockError(stockId, idx, `Bild ${i+1} konnte nicht geladen werden`);
                return;
            }
            let file = new File([blob], fileInfo.name, { type: fileInfo.mimeType });
            let dataTransfer = new DataTransfer();
            dataTransfer.items.add(file);
            fileInput.files = dataTransfer.files;
            fileInput.dispatchEvent(new Event('change', { bubbles: true }));

            await sleep(1000);

            let submitBtn = await waitForExactText(['Upload', 'Hochladen', 'Save', 'Speichern', 'Add', 'Hinzufügen'], 'button', 10000);
            if (abortMission) return;
            if (!submitBtn) {
                dbg('uploadStep', 'submitButtonMissing', fileInfo.name);
                showCustomPopup("ARASAKA FEHLER", "Upload-Bestätigung nicht gefunden. Stoppe, damit nichts doppelt hochgeladen wird.", true);
                isProcessing = false;
                return;
            }
            forceClick(submitBtn);

            showCustomPopup("ARASAKA VERIFIKATION", `Warte auf Bestätigung im System...`, false);
            let verifySuccess = await waitForText(currentComment, 20000);
            dbg('postUploadVerify', 'comment', currentComment, 'verifySuccess', verifySuccess);
            if (abortMission) return;

            if(!verifySuccess) {
                dbg('postUploadVerify', 'noMatchExtraWait');
                await sleep(2000);
                verifySuccess = await waitForText(currentComment, 10000);
                dbg('postUploadVerify', 'afterExtraWait', currentComment, 'verifySuccess', verifySuccess);
            }

            if(!verifySuccess) {
                showCustomPopup("ARASAKA FEHLER", "Upload wurde nicht bestätigt. Stoppe, damit nichts doppelt hochgeladen wird. Bitte Auftrag prüfen.", true);
                isProcessing = false;
                return;
            }

            if (abortMission) return;
            if (!await moveFileOrStop(fileInfo, isRetoure, {
                recordLastUpload: true,
                logStockId: stockId,
                logFileName: fileInfo.name
            })) return;
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

        if (syncStatus !== "OK") {
            let checkRes = await bridgePostJson({
                action: 'checkSheetMark',
                stockId: stockId
            }, 30000);
            if (checkRes) {
                var crt = String(checkRes.responseText || '').trim();
                dbg('checkSheetMark', 'http', checkRes.status, 'body', crt.slice(0, 200));
                if (crt === "OK") syncStatus = "OK";
            } else {
                dbg('checkSheetMark', 'noResponse');
            }
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
        continueWithNextStock(idx + 1);
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

    function hudEscape(value) {
        return String(value == null ? '' : value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function hudMessageHtml(message) {
        return hudEscape(message).replace(/\n/g, '<br>');
    }

    function hudTone(title, message, isEnd) {
        var text = String(title || '') + ' ' + String(message || '');
        var lower = text.toLowerCase();
        if (/fehler|error|konnte nicht|keine verbindung|nicht richtig|html_fehler|http_error|fehlgeschlagen/.test(lower)) return 'error';
        if (/warnung|wartet|länger/.test(lower)) return 'warn';
        if (/skip|übersprungen|duplicate|duplikat|doppelt/.test(lower)) return 'skip';
        if (/mahlzeit|fertig|sauber|gestoppt/.test(lower) || isEnd) return 'done';
        if (/download|upload|sync|prüfe|suche|warte|läuft|online|navigation|verifikation|stapel|drive/.test(lower)) return 'active';
        return 'neutral';
    }

    function hudToneLabel(tone) {
        if (tone === 'error') return 'FEHLER';
        if (tone === 'warn') return 'WARNUNG';
        if (tone === 'skip') return 'SKIP';
        if (tone === 'done') return 'FERTIG';
        if (tone === 'active') return 'LÄUFT';
        return 'INFO';
    }

    function hudProgressChips(message) {
        var chips = [];
        try {
            var keys = JSON.parse(sessionStorage.getItem('arasaka_batch_keys') || '[]');
            var idx = parseInt(sessionStorage.getItem('arasaka_batch_current_idx') || '0', 10);
            if (keys.length > 0 && idx < keys.length) {
                chips.push('Stock ' + (idx + 1) + '/' + keys.length + ' · ' + keys[idx]);
            } else if (keys.length > 0) {
                chips.push('Recheck · ' + keys.length + ' offen');
            }
        } catch (e) {}
        var img = String(message || '').match(/Bild\s+(\d+)\s+von\s+(\d+)/i);
        if (img) chips.push('Bild ' + img[1] + '/' + img[2]);
        if (ARASAKA_DEBUG) chips.push('Debug an');
        return chips;
    }

    function hudPosition() {
        try {
            var raw = localStorage.getItem(ARASAKA_HUD_POS_KEY);
            if (!raw) return null;
            var pos = JSON.parse(raw);
            if (typeof pos.left !== 'number' || typeof pos.top !== 'number') return null;
            return pos;
        } catch (e) {
            return null;
        }
    }

    function hudClamp(value, min, max) {
        return Math.max(min, Math.min(max, value));
    }

    function hudApplyPosition(popup) {
        var pos = hudPosition();
        if (!pos) {
            popup.style.left = '50%';
            popup.style.top = '50%';
            popup.style.transform = 'translate(-50%, -50%)';
            return;
        }
        requestAnimationFrame(function() {
            var maxLeft = Math.max(12, window.innerWidth - popup.offsetWidth - 12);
            var maxTop = Math.max(12, window.innerHeight - popup.offsetHeight - 12);
            popup.style.left = hudClamp(pos.left, 12, maxLeft) + 'px';
            popup.style.top = hudClamp(pos.top, 12, maxTop) + 'px';
            popup.style.transform = 'none';
        });
    }

    function hudSavePosition(left, top) {
        try {
            localStorage.setItem(ARASAKA_HUD_POS_KEY, JSON.stringify({ left: left, top: top }));
        } catch (e) {}
    }

    function hudStopProcess() {
        abortMission = true;
        isProcessing = false;
        sessionStorage.removeItem('arasaka_batch_data');
        sessionStorage.removeItem('arasaka_batch_keys');
        sessionStorage.removeItem('arasaka_batch_current_idx');
        showCustomPopup("ARASAKA STOP", "Prozess gestoppt. Speicher gelöscht.", true);
    }

    function hudEnsureStyles() {
        if (document.getElementById('arasaka-hud-style')) return;
        var style = document.createElement('style');
        style.id = 'arasaka-hud-style';
        style.textContent = `
            #arasaka-batch-popup {
                position: fixed;
                z-index: 9999999;
                width: min(560px, calc(100vw - 32px));
                color: #c9d1d9;
                font-family: "Segoe UI", Arial, sans-serif;
                border-radius: 22px;
                overflow: hidden;
                background:
                    radial-gradient(circle at top left, rgba(61, 158, 220, 0.14), transparent 34%),
                    radial-gradient(circle at top right, rgba(242, 116, 32, 0.1), transparent 32%),
                    linear-gradient(180deg, rgba(21, 27, 35, 0.98) 0%, rgba(14, 19, 26, 0.98) 100%);
                border: 1px solid #2d3642;
                box-shadow: 0 24px 56px rgba(0, 0, 0, 0.52);
                backdrop-filter: blur(12px);
                -webkit-backdrop-filter: blur(12px);
                user-select: none;
            }
            #arasaka-batch-popup[data-tone="active"] { border-color: rgba(61, 158, 220, 0.52); box-shadow: 0 24px 56px rgba(0, 0, 0, 0.52), 0 0 34px rgba(61, 158, 220, 0.18); }
            #arasaka-batch-popup[data-tone="done"] { border-color: rgba(86, 211, 100, 0.55); box-shadow: 0 24px 56px rgba(0, 0, 0, 0.52), 0 0 34px rgba(86, 211, 100, 0.16); }
            #arasaka-batch-popup[data-tone="skip"] { border-color: rgba(242, 116, 32, 0.62); box-shadow: 0 24px 56px rgba(0, 0, 0, 0.52), 0 0 34px rgba(242, 116, 32, 0.18); }
            #arasaka-batch-popup[data-tone="warn"] { border-color: rgba(227, 180, 60, 0.64); box-shadow: 0 24px 56px rgba(0, 0, 0, 0.52), 0 0 34px rgba(227, 180, 60, 0.16); }
            #arasaka-batch-popup[data-tone="error"] { border-color: rgba(248, 81, 73, 0.68); box-shadow: 0 24px 56px rgba(0, 0, 0, 0.52), 0 0 38px rgba(248, 81, 73, 0.22); }
            .arasaka-hud-head {
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 18px;
                padding: 18px 20px 16px;
                background: linear-gradient(180deg, rgba(22, 27, 34, 0.99) 0%, rgba(18, 23, 30, 0.99) 100%);
                border-bottom: 1px solid #242d39;
                cursor: move;
            }
            .arasaka-hud-eyebrow {
                color: #8b949e;
                font-size: 12px;
                font-weight: 800;
                letter-spacing: 1.6px;
                text-transform: uppercase;
                margin-bottom: 4px;
            }
            .arasaka-hud-title {
                color: #f0f6fc;
                font-size: 20px;
                font-weight: 900;
                letter-spacing: 0.5px;
                line-height: 1.15;
                text-transform: uppercase;
            }
            .arasaka-hud-actions {
                display: flex;
                align-items: center;
                gap: 8px;
                flex-shrink: 0;
            }
            .arasaka-hud-btn {
                border: 1px solid rgba(255,255,255,0.12);
                border-radius: 14px;
                background: rgba(13, 17, 23, 0.72);
                color: #c9d1d9;
                min-width: 42px;
                height: 40px;
                padding: 0 14px;
                font-size: 13px;
                font-weight: 900;
                cursor: pointer;
                transition: transform 0.14s, border-color 0.14s, background 0.14s;
            }
            .arasaka-hud-btn:hover {
                transform: translateY(-1px);
                border-color: rgba(61, 158, 220, 0.46);
                background: rgba(61, 158, 220, 0.14);
            }
            .arasaka-hud-stop:hover {
                border-color: rgba(248, 81, 73, 0.48);
                background: rgba(248, 81, 73, 0.14);
                color: #ff7b72;
            }
            .arasaka-hud-body {
                padding: 22px;
            }
            .arasaka-hud-main {
                display: flex;
                flex-direction: column;
                gap: 16px;
                align-items: start;
            }
            .arasaka-hud-badge {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                gap: 10px;
                min-height: 40px;
                padding: 9px 16px;
                border-radius: 999px;
                background: rgba(13, 17, 23, 0.7);
                border: 2px solid rgba(61, 158, 220, 0.38);
                color: #3FA0DB;
                font-size: 13px;
                font-weight: 900;
                letter-spacing: 1.2px;
                text-transform: uppercase;
                box-shadow: inset 0 0 0 1px rgba(255,255,255,0.03);
            }
            .arasaka-hud-badge::before {
                content: "";
                width: 10px;
                height: 10px;
                border-radius: 50%;
                background: currentColor;
                box-shadow: 0 0 16px currentColor;
            }
            #arasaka-batch-popup[data-tone="active"] .arasaka-hud-badge { animation: arasakaHudPulse 1.35s ease-in-out infinite; }
            #arasaka-batch-popup[data-tone="done"] .arasaka-hud-badge { color: #56d364; border-color: rgba(86, 211, 100, 0.48); }
            #arasaka-batch-popup[data-tone="skip"] .arasaka-hud-badge { color: #F27420; border-color: rgba(242, 116, 32, 0.54); }
            #arasaka-batch-popup[data-tone="warn"] .arasaka-hud-badge { color: #e3b43c; border-color: rgba(227, 180, 60, 0.54); }
            #arasaka-batch-popup[data-tone="error"] .arasaka-hud-badge { color: #ff7b72; border-color: rgba(248, 81, 73, 0.58); }
            .arasaka-hud-message {
                color: #d7dee8;
                font-size: 16px;
                font-weight: 700;
                line-height: 1.52;
                white-space: normal;
                word-break: break-word;
                user-select: text;
            }
            .arasaka-hud-chips {
                display: flex;
                flex-wrap: wrap;
                gap: 8px;
                margin-top: 18px;
            }
            .arasaka-hud-chip {
                display: inline-flex;
                align-items: center;
                min-height: 32px;
                padding: 6px 12px;
                border-radius: 999px;
                border: 1px solid rgba(61, 158, 220, 0.24);
                background: rgba(13, 17, 23, 0.56);
                color: #8bcef7;
                font-size: 13px;
                font-weight: 850;
                letter-spacing: 0.15px;
            }
            .arasaka-hud-reset {
                margin-top: 16px;
                padding: 12px 13px;
                border-radius: 14px;
                border: 1px solid rgba(248, 81, 73, 0.28);
                background: rgba(248, 81, 73, 0.09);
                color: #ffb3ad;
                font-size: 14px;
                font-weight: 800;
                line-height: 1.42;
                user-select: text;
            }
            .arasaka-hud-footer {
                display: flex;
                justify-content: space-between;
                gap: 10px;
                margin-top: 16px;
                padding-top: 12px;
                border-top: 1px solid rgba(255,255,255,0.07);
                color: #8b949e;
                font-size: 12px;
                font-weight: 800;
                letter-spacing: 0.25px;
            }
            @keyframes arasakaHudPulse {
                0%, 100% { transform: scale(1); box-shadow: 0 0 0 0 rgba(61, 158, 220, 0.32); }
                50% { transform: scale(1.06); box-shadow: 0 0 18px 4px rgba(61, 158, 220, 0.22); }
            }
        `;
        document.head.appendChild(style);
    }

    function hudEnableDrag(popup, handle) {
        var state = null;
        function move(e) {
            if (!state) return;
            var maxLeft = Math.max(12, window.innerWidth - popup.offsetWidth - 12);
            var maxTop = Math.max(12, window.innerHeight - popup.offsetHeight - 12);
            var left = hudClamp(e.clientX - state.dx, 12, maxLeft);
            var top = hudClamp(e.clientY - state.dy, 12, maxTop);
            popup.style.left = left + 'px';
            popup.style.top = top + 'px';
            popup.style.transform = 'none';
            state.left = left;
            state.top = top;
        }
        function up() {
            if (state) hudSavePosition(state.left, state.top);
            state = null;
            document.removeEventListener('pointermove', move);
        }
        handle.addEventListener('pointerdown', function(e) {
            if (e.button !== 0 || e.target.closest('button')) return;
            var rect = popup.getBoundingClientRect();
            state = { dx: e.clientX - rect.left, dy: e.clientY - rect.top, left: rect.left, top: rect.top };
            popup.style.left = rect.left + 'px';
            popup.style.top = rect.top + 'px';
            popup.style.transform = 'none';
            document.addEventListener('pointermove', move);
            document.addEventListener('pointerup', up, { once: true });
            e.preventDefault();
        });
    }

    function showCustomPopup(title, message, isEnd) {
        hudEnsureStyles();
        let existing = document.getElementById('arasaka-batch-popup');
        if (existing) existing.remove();
        let tone = hudTone(title, message, isEnd);
        let chips = hudProgressChips(message);
        let chipHtml = chips.map(function(chip) {
            return '<span class="arasaka-hud-chip">' + hudEscape(chip) + '</span>';
        }).join('');
        let resetHtml = tone === 'error'
            ? '<div class="arasaka-hud-reset">Erst hart neu laden: Strg+F5 oder Ctrl+Shift+R. Danach ALT+B nochmal versuchen. Wenn es wieder kommt, Konsole (F12) offen lassen.</div>'
            : '';
        let popup = document.createElement('div');
        popup.id = 'arasaka-batch-popup';
        popup.setAttribute('data-tone', tone);
        popup.innerHTML = `
            <div class="arasaka-hud-head" id="arasaka-hud-drag">
                <div>
                    <div class="arasaka-hud-eyebrow">Kisten Ausgabe · ${hudEscape(hudToneLabel(tone))}</div>
                    <div class="arasaka-hud-title">${hudEscape(title)}</div>
                </div>
                <div class="arasaka-hud-actions">
                    <button type="button" class="arasaka-hud-btn arasaka-hud-stop" id="arasaka-hud-stop">STOP</button>
                    <button type="button" class="arasaka-hud-btn" id="arasaka-hud-close">X</button>
                </div>
            </div>
            <div class="arasaka-hud-body">
                <div class="arasaka-hud-main">
                    <div class="arasaka-hud-badge">${hudEscape(hudToneLabel(tone))}</div>
                    <div class="arasaka-hud-message">${hudMessageHtml(message)}</div>
                </div>
                ${chipHtml ? '<div class="arasaka-hud-chips">' + chipHtml + '</div>' : ''}
                ${resetHtml}
                <div class="arasaka-hud-footer">
                    <span>Bot v${hudEscape(ARASAKA_BOT_VERSION)} · Bridge v${hudEscape(ARASAKA_BRIDGE_VERSION)}</span>
                    <span>ALT+B Start · ESC Stop</span>
                </div>
            </div>
        `;
        document.body.appendChild(popup);
        hudApplyPosition(popup);
        hudEnableDrag(popup, popup.querySelector('#arasaka-hud-drag'));
        popup.querySelector('#arasaka-hud-close').addEventListener('click', function(e) {
            e.stopPropagation();
            popup.remove();
        });
        popup.querySelector('#arasaka-hud-stop').addEventListener('click', function(e) {
            e.stopPropagation();
            hudStopProcess();
        });
    }
})();
