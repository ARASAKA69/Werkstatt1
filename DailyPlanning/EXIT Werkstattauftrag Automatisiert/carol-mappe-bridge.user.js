// ==UserScript==
// @name         Carol Werkstattmappe Bridge
// @namespace    arasaka
// @version      0.1
// @description  Sendet Modell + VIN aus Carol an das EXIT Werkstattmappe HUD
// @match        *://carol.autohero.com/*
// @grant        GM_xmlhttpRequest
// @grant        GM_setValue
// @grant        GM_getValue
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';

    var CONSOLE_PROBE_SNIPPET = [
        "performance.getEntriesByType('resource')",
        "  .filter(function(r){ return r.initiatorType === 'fetch' || r.initiatorType === 'xmlhttprequest'; })",
        "  .map(function(r){ return r.name; })",
        "  .filter(function(u, i, a){ return a.indexOf(u) === i; })",
        "  .forEach(function(u){ console.log(u); });"
    ].join("\n");

    var STORE_URL = 'wm_bridge_webapp_url';
    var STORE_SECRET = 'wm_bridge_secret';
    var STORE_AUTO = 'wm_bridge_auto';

    function gmGet(key, fallback) {
        try { return GM_getValue(key, fallback); } catch (e) {}
        try { return localStorage.getItem(key) || fallback; } catch (e2) {}
        return fallback;
    }

    function gmSet(key, value) {
        try { GM_setValue(key, value); return; } catch (e) {}
        try { localStorage.setItem(key, value); } catch (e2) {}
    }

    function extractVehicleData() {
        var result = { stockId: '', modell: '', vin: '' };
        var bodyText = document.body ? document.body.innerText : '';
        var vinMatch = bodyText.match(/VIN[:\s]*([A-HJ-NPR-Z0-9]{11,17})/i);
        if (vinMatch) result.vin = vinMatch[1].toUpperCase();
        var headings = document.querySelectorAll('h1, h2, h3, [class*="title"], [class*="Title"]');
        for (var i = 0; i < headings.length; i++) {
            var txt = String(headings[i].textContent || '').trim();
            var m = txt.match(/^([A-Z]{2}\d{4,})\s*[-–]\s*(.+)$/);
            if (m) {
                result.stockId = m[1].toUpperCase();
                result.modell = m[2].trim();
                break;
            }
        }
        if (!result.stockId) {
            var sidMatch = bodyText.match(/\b([A-Z]{2}\d{5,})\b\s*[-–]\s*([^\n]+)/);
            if (sidMatch) {
                result.stockId = sidMatch[1].toUpperCase();
                result.modell = String(sidMatch[2] || '').trim();
            }
        }
        return result;
    }

    function sendToHud(data, onDone) {
        var url = gmGet(STORE_URL, '');
        var secret = gmGet(STORE_SECRET, '');
        if (!url) { onDone(false, 'Web-App URL fehlt (im Panel eintragen)'); return; }
        if (!data.stockId || !data.vin || !data.modell) {
            onDone(false, 'Unvollständig: ' + JSON.stringify(data));
            return;
        }
        var payload = JSON.stringify({
            secret: secret,
            entries: [{ stockId: data.stockId, modell: data.modell, vin: data.vin }]
        });
        if (typeof GM_xmlhttpRequest === 'function') {
            GM_xmlhttpRequest({
                method: 'POST',
                url: url,
                data: payload,
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                onload: function (res) {
                    var ok = false;
                    var msg = 'HTTP ' + res.status;
                    try {
                        var parsed = JSON.parse(res.responseText);
                        ok = !!parsed.success;
                        if (parsed.message) msg = parsed.message;
                    } catch (e) {}
                    onDone(ok, ok ? 'Gesendet: ' + data.stockId : msg);
                },
                onerror: function () { onDone(false, 'Netzwerkfehler'); }
            });
        } else {
            fetch(url, {
                method: 'POST',
                mode: 'no-cors',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: payload
            }).then(function () {
                onDone(true, 'Gesendet (no-cors): ' + data.stockId);
            }).catch(function () {
                onDone(false, 'Netzwerkfehler');
            });
        }
    }

    var panel, statusEl;

    function buildPanel() {
        panel = document.createElement('div');
        panel.style.cssText = 'position:fixed;bottom:14px;right:14px;z-index:999999;background:#161b22;color:#e6edf3;'
            + 'border:1px solid #f97316;border-radius:10px;padding:10px 12px;font:12px/1.5 "Segoe UI",sans-serif;'
            + 'box-shadow:0 0 18px rgba(249,115,22,.35);width:250px;';
        panel.innerHTML = ''
            + '<div style="font-weight:700;margin-bottom:6px;color:#ffa94d;">Werkstattmappe Bridge</div>'
            + '<input id="wmb-url" placeholder="Web-App /exec URL" style="width:100%;margin-bottom:4px;padding:4px 6px;'
            + 'background:#0b0f14;border:1px solid #30363d;border-radius:6px;color:#e6edf3;font-size:11px;">'
            + '<input id="wmb-secret" placeholder="Secret" style="width:100%;margin-bottom:6px;padding:4px 6px;'
            + 'background:#0b0f14;border:1px solid #30363d;border-radius:6px;color:#e6edf3;font-size:11px;">'
            + '<label style="display:flex;align-items:center;gap:6px;margin-bottom:6px;cursor:pointer;">'
            + '<input type="checkbox" id="wmb-auto"> Auto-Senden auf Detailseite</label>'
            + '<div style="display:flex;gap:6px;">'
            + '<button id="wmb-send" style="flex:1;background:#f97316;border:0;border-radius:6px;padding:5px 8px;'
            + 'color:#10141a;font-weight:700;cursor:pointer;">Senden</button>'
            + '<button id="wmb-probe" style="background:#1c2330;border:1px solid #30363d;border-radius:6px;'
            + 'padding:5px 8px;color:#e6edf3;cursor:pointer;">API-Probe</button>'
            + '</div>'
            + '<div id="wmb-status" style="margin-top:6px;color:#9da7b3;min-height:14px;"></div>';
        document.body.appendChild(panel);
        statusEl = panel.querySelector('#wmb-status');

        var urlInput = panel.querySelector('#wmb-url');
        var secretInput = panel.querySelector('#wmb-secret');
        var autoCheck = panel.querySelector('#wmb-auto');
        urlInput.value = gmGet(STORE_URL, '');
        secretInput.value = gmGet(STORE_SECRET, '');
        autoCheck.checked = gmGet(STORE_AUTO, '') === '1';
        urlInput.addEventListener('change', function () { gmSet(STORE_URL, urlInput.value.trim()); });
        secretInput.addEventListener('change', function () { gmSet(STORE_SECRET, secretInput.value.trim()); });
        autoCheck.addEventListener('change', function () { gmSet(STORE_AUTO, autoCheck.checked ? '1' : ''); });

        panel.querySelector('#wmb-send').addEventListener('click', function () {
            setStatus('Lese Seite…');
            var data = extractVehicleData();
            sendToHud(data, function (ok, msg) {
                setStatus(msg, ok ? '#56d364' : '#f85149');
            });
        });
        panel.querySelector('#wmb-probe').addEventListener('click', function () {
            console.log('%cWerkstattmappe Bridge — API Probe', 'color:#f97316;font-weight:bold');
            console.log('Alle fetch/XHR URLs dieser Seite (Kandidaten für interne Carol-API):');
            performance.getEntriesByType('resource')
                .filter(function (r) { return r.initiatorType === 'fetch' || r.initiatorType === 'xmlhttprequest'; })
                .map(function (r) { return r.name; })
                .filter(function (u, i, a) { return a.indexOf(u) === i; })
                .forEach(function (u) { console.log(u); });
            console.log('Snippet zum manuellen Ausführen in der Konsole:');
            console.log(CONSOLE_PROBE_SNIPPET);
            setStatus('Probe in Konsole ausgegeben (F12)', '#58a6ff');
        });
    }

    function setStatus(msg, color) {
        if (!statusEl) return;
        statusEl.textContent = msg;
        statusEl.style.color = color || '#9da7b3';
    }

    var lastAutoUrl = '';

    function autoTick() {
        if (gmGet(STORE_AUTO, '') !== '1') return;
        var href = location.href;
        if (href === lastAutoUrl) return;
        if (!/\/refurbishment\/[0-9a-f-]{20,}/i.test(href)) return;
        var data = extractVehicleData();
        if (!data.stockId || !data.vin) return;
        lastAutoUrl = href;
        setStatus('Auto-Senden: ' + data.stockId + '…');
        sendToHud(data, function (ok, msg) {
            setStatus(msg, ok ? '#56d364' : '#f85149');
        });
    }

    function init() {
        if (!document.body) { setTimeout(init, 500); return; }
        buildPanel();
        setInterval(autoTick, 2000);
    }

    init();
})();
