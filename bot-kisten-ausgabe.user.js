// ==UserScript==
// @name         ARASAKA Master-Bot (Upload)
// @namespace    http://tampermonkey.net/
// @version      1.16
// @description  Live-Version
// @author       ARASAKA
// @match        *://carol.autohero.com/*
// @grant        GM_xmlhttpRequest
// ==/UserScript==

(function() {
    'use strict';

    const DRIVE_WEB_APP_URL = "https://script.google.com/a/macros/autohero.com/s/AKfycbz0yz1BdUx4ZXgT4V4rqfif8KM3D76rNDjWXY2DZD9JIP0D4y9cjsGsFooOZqaGlm1c/exec";
    const API_KEY = "ARASAKA_2026";

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
            audio.play().catch(e => console.log('Audio Autoplay blockiert', e));
        } catch (e) {}
    }

    function forceClick(el) {
        if (!el) return;
        el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
        el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        el.click();
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
        showCustomPopup("ARASAKA FEHLER", `${message}\nVerschiebe in Kisten Falsche Stock Ordner...`, false);
        let allData = JSON.parse(sessionStorage.getItem('arasaka_batch_data'));
        let files = allData[stockId];

        // Alle Bilder gleichzeitig (parallel) an Google senden für maximalen Speed!
        let movePromises = files.map(f => new Promise(resolve => {
            GM_xmlhttpRequest({
                method: "GET",
                url: `${DRIVE_WEB_APP_URL}?action=moveFileError&fileId=${f.id}&stockId=${stockId}&reason=${encodeURIComponent(message)}&key=${API_KEY}`,
                timeout: 10000,
                onload: resolve,
                onerror: resolve,
                ontimeout: resolve
            });
        }));

        // Warten, bis Google für alle Bilder das "OK" zurückgibt, dann sofort weiter!
        await Promise.all(movePromises);

        sessionStorage.setItem('arasaka_batch_current_idx', (idx + 1).toString());
        window.location.href = '/';
    }

    async function startBatchProcess() {
        showCustomPopup("ARASAKA ONLINE", "Prüfe Kisten im Google Drive...", false);

        GM_xmlhttpRequest({
            method: "GET",
            url: `${DRIVE_WEB_APP_URL}?action=getBatch&key=${API_KEY}`,
            timeout: 15000,
            onload: async function(response) {
                if (abortMission) return;
                try {
                    let data = JSON.parse(response.responseText);
                    let stockIds = Object.keys(data);

                    if (stockIds.length === 0) {
                        playSuccessSound();
                        showCustomPopup("Mahlzeit!", "Kisten Offen ist leer. Keine neuen Bilder gefunden.\n\n[ALT + B] drücken, um später neu zu scannen.\nZeit, den Fehler-Ordner zu checken!", true);
                        isProcessing = false;
                        return;
                    }

                    sessionStorage.setItem('arasaka_batch_data', JSON.stringify(data));
                    sessionStorage.setItem('arasaka_batch_keys', JSON.stringify(stockIds));
                    sessionStorage.setItem('arasaka_batch_current_idx', '0');
                    processNextStock();
                } catch (err) {
                    showCustomPopup("FEHLER", "Google Drive antwortet nicht richtig. Skript-URL und API-Key prüfen!", true);
                    isProcessing = false;
                }
            },
            onerror: function() { showCustomPopup("FEHLER", "Keine Verbindung zu Google Drive möglich.", true); isProcessing = false; },
            ontimeout: function() { showCustomPopup("FEHLER", "Zeitüberschreitung bei Google Drive.", true); isProcessing = false; }
        });
    }

    async function processNextStock() {
        if (abortMission) return;

        let keys = JSON.parse(sessionStorage.getItem('arasaka_batch_keys') || "[]");
        let idx = parseInt(sessionStorage.getItem('arasaka_batch_current_idx') || "0");

        if (idx >= keys.length) {
            sessionStorage.removeItem('arasaka_batch_data');
            sessionStorage.removeItem('arasaka_batch_keys');
            sessionStorage.removeItem('arasaka_batch_current_idx');

            showCustomPopup("ARASAKA", "Stapel fertig. Kurzer Check im Drive...", false);

            GM_xmlhttpRequest({
                method: "GET",
                url: `${DRIVE_WEB_APP_URL}?action=getBatch&key=${API_KEY}`,
                timeout: 15000,
                onload: async function(response) {
                    if (abortMission) return;
                    try {
                        let data = JSON.parse(response.responseText);
                        let stockIds = Object.keys(data);

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
                        showCustomPopup("FEHLER", "Fehler beim Auto-Recheck.", true);
                        isProcessing = false;
                    }
                }
            });
            return;
        }

        let stockId = keys[idx].toUpperCase();
        showCustomPopup("ARASAKA LÄUFT", `Suche nach Stock ID: ${stockId}...`, false);

        let searchInput = await findSearchBar();
        if (!searchInput) {
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

        await sleep(300); // Nur noch ein winziger Moment statt 500ms

        searchInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
        searchInput.dispatchEvent(new KeyboardEvent('keypress', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
        searchInput.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));

        setTimeout(() => { if(searchInput) searchInput.style.border = ""; }, 1000);

        showCustomPopup("ARASAKA LÄUFT", `Prüfe Ergebnisse für ${stockId}...`, false);

        if (abortMission) return;

        let resultRow = null;
        // High-Speed Loop: Scannt 25x alle 200ms (max 5 Sekunden) statt feste Wartezeiten!
        for (let attempt = 0; attempt < 25; attempt++) {
            if (abortMission) return;

            let links = Array.from(document.querySelectorAll('a[href*="/refurbishment/"]'));
            let validLinks = links.filter(l => l.href.match(/[0-9a-f]{8}-[0-9a-f]{4}/i));
            if (validLinks.length > 0) { resultRow = validLinks[0]; break; }
            else if (links.length > 0) { resultRow = links[0]; break; }

            let rows = Array.from(document.querySelectorAll('tr.clickable-row, .rt-tr-group, tr[data-test-id*="row"]'));
            let validRows = rows.filter(r => !r.querySelector('th'));
            if (validRows.length > 0) { resultRow = validRows[0]; break; }

            // Wenn die Seite aktiv meldet, dass sie leer ist ("no data", "0 results"), und 1 Sekunde vergangen ist -> sofortiger Abbruch!
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
            forceClick(resultRow);
            showCustomPopup("ARASAKA VERIFIKATION", `Prüfe, ob Auftrag ${stockId} geöffnet wurde...`, false);
            let isCorrectPage = await waitForText(stockId, 15000);
            if (abortMission) return;

            if (isCorrectPage) {
                let uploadReady = await waitForElementByText(['Upload Document', 'Dokument hochladen'], 'button', 15000);
                if (abortMission) return;
                if (uploadReady) {
                    executeUploadsForStock(stockId);
                } else {
                    await handleStockError(stockId, idx, "Upload Button im Auftrag fehlt");
                }
            } else {
                await handleStockError(stockId, idx, "Falscher Auftrag geladen");
            }

        } else {
            await handleStockError(stockId, idx, "Auftrag nicht gefunden");
        }
    }

    async function executeUploadsForStock(stockId) {
        let allData = JSON.parse(sessionStorage.getItem('arasaka_batch_data'));
        let files = allData[stockId];
        let totalFiles = files.length;
        let idx = parseInt(sessionStorage.getItem('arasaka_batch_current_idx') || "0");

        for (let i = 0; i < totalFiles; i++) {
            if (abortMission) return;

            let fileInfo = files[i];
            let currentComment = `Ausgabe ${i + 1}/${totalFiles}`;

            // DOPPEL-UPLOAD-SCHUTZ
            if (document.body.innerText.includes(currentComment)) {
                showCustomPopup("ARASAKA SKIP", `Bild ${i + 1} (${currentComment}) existiert bereits. Überspringe...`, false);
                await new Promise(resolve => {
                    GM_xmlhttpRequest({
                        method: "GET",
                        url: `${DRIVE_WEB_APP_URL}?action=moveFile&fileId=${fileInfo.id}&key=${API_KEY}`,
                        timeout: 10000,
                        onload: resolve, onerror: resolve, ontimeout: resolve
                    });
                });
                continue;
            }

            showCustomPopup("ARASAKA DOWNLOAD", `Lade Bild ${i + 1} von ${totalFiles} für ${stockId} aus Drive...`, false);
            let b64 = await new Promise((resolve) => {
                GM_xmlhttpRequest({
                    method: "GET",
                    url: `${DRIVE_WEB_APP_URL}?action=getFileData&fileId=${fileInfo.id}&key=${API_KEY}`,
                    timeout: 45000,
                    onload: (res) => resolve(res.responseText),
                    onerror: () => resolve(null),
                    ontimeout: () => resolve(null)
                });
            });

            if (!b64) {
                await handleStockError(stockId, idx, `Bild ${i+1} konnte nicht geladen werden`);
                return;
            }

            showCustomPopup("ARASAKA UPLOAD", `Lade Bild ${i + 1} von ${totalFiles} in Carol hoch...`, false);

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
            if (abortMission) return;

            if(!verifySuccess) {
                await sleep(2000);
            }

            await new Promise(resolve => {
                if (abortMission) return resolve();
                GM_xmlhttpRequest({
                    method: "GET",
                    url: `${DRIVE_WEB_APP_URL}?action=moveFile&fileId=${fileInfo.id}&key=${API_KEY}`,
                    timeout: 10000,
                    onload: resolve, onerror: resolve, ontimeout: resolve
                });
            });
        }

        if (abortMission) return;

        sessionStorage.setItem('arasaka_batch_current_idx', (idx + 1).toString());
        showCustomPopup("ARASAKA", `${stockId} sauber hochgeladen. Lade nächste Seite...`, false);
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
