// ==UserScript==
// @name         Carol Werkstattmappe Bridge
// @namespace    arasaka
// @version      0.3
// @description  Sendet Modell + VIN aus Carol an das EXIT Werkstattmappe HUD (GraphQL Sniffer + DOM Fallback + Autofill-Navigation)
// @match        *://carol.autohero.com/*
// @grant        GM_xmlhttpRequest
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        unsafeWindow
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @run-at       document-start
// ==/UserScript==

(function () {
    'use strict';

    var GRAPHQL_URL = '/api/v1/refurbishment-aggregation/graphql';
    var MSG_MARKER = '__wmBridgeGraphql';
    var STORE_URL = 'wm_bridge_webapp_url';
    var STORE_SECRET = 'wm_bridge_secret';
    var STORE_AUTO = 'wm_bridge_auto';

    var lastCapture = null;
    var captureLog = {};

    function gmGet(key, fallback) {
        try { return GM_getValue(key, fallback); } catch (e) {}
        try { return localStorage.getItem(key) || fallback; } catch (e2) {}
        return fallback;
    }

    function gmSet(key, value) {
        try { GM_setValue(key, value); return; } catch (e) {}
        try { localStorage.setItem(key, value); } catch (e2) {}
    }

    function injectPageHook() {
        var hookSrc = '(function(){'
            + 'if(window.__wmBridgeHooked)return;window.__wmBridgeHooked=true;'
            + 'var of=window.fetch;'
            + 'window.fetch=function(){var a=arguments;'
            + 'var u=String(a[0]&&a[0].url?a[0].url:a[0]||"");'
            + 'var p=of.apply(this,a);'
            + 'if(u.indexOf("graphql")!==-1){p.then(function(r){'
            + 'try{r.clone().json().then(function(d){'
            + 'window.postMessage({' + MSG_MARKER + ':1,url:u,data:d},"*");'
            + '}).catch(function(){});}catch(e){}return r;});}'
            + 'return p;};'
            + 'var ox=XMLHttpRequest.prototype.open;var os=XMLHttpRequest.prototype.send;'
            + 'XMLHttpRequest.prototype.open=function(m,u){this.__wmUrl=String(u||"");return ox.apply(this,arguments);};'
            + 'XMLHttpRequest.prototype.send=function(){var x=this;'
            + 'if(String(x.__wmUrl).indexOf("graphql")!==-1){x.addEventListener("load",function(){'
            + 'try{var d=JSON.parse(x.responseText);'
            + 'window.postMessage({' + MSG_MARKER + ':1,url:x.__wmUrl,data:d},"*");}catch(e){}});}'
            + 'return os.apply(this,arguments);};'
            + '})();';
        try {
            var s = document.createElement('script');
            s.textContent = hookSrc;
            (document.head || document.documentElement).appendChild(s);
            s.remove();
        } catch (e) {}
        try {
            if (typeof unsafeWindow !== 'undefined' && !unsafeWindow.__wmBridgeHooked) {
                var evalFn = unsafeWindow.eval;
                if (evalFn) evalFn(hookSrc);
            }
        } catch (e2) {}
    }

    function strOrName(v) {
        if (typeof v === 'string') return v.trim();
        if (v && typeof v === 'object' && typeof v.name === 'string') return v.name.trim();
        return '';
    }

    function buildModelName(node) {
        var direct = strOrName(node.fullName) || strOrName(node.title) || strOrName(node.displayName);
        if (direct) return direct;
        var parts = [];
        var make = strOrName(node.make) || strOrName(node.manufacturer) || strOrName(node.brand);
        var model = strOrName(node.model) || strOrName(node.modelName);
        var sub = strOrName(node.subModel) || strOrName(node.modelVariant) || strOrName(node.variant) || strOrName(node.trim);
        if (make) parts.push(make);
        if (model) parts.push(model);
        if (sub && sub !== model) parts.push(sub);
        return parts.join(' ').trim();
    }

    function firstStock(node) {
        var cands = ['stockNumber', 'stockId', 'stock_number', 'adId', 'refurbishmentNumber'];
        for (var i = 0; i < cands.length; i++) {
            var v = node[cands[i]];
            if (typeof v === 'string' && /^[A-Z]{2}\d{4,}$/i.test(v.trim())) return v.trim().toUpperCase();
        }
        return '';
    }

    function scanForVehicles(node, out, depth) {
        if (!node || typeof node !== 'object' || depth > 14) return;
        if (Array.isArray(node)) {
            for (var i = 0; i < node.length; i++) scanForVehicles(node[i], out, depth + 1);
            return;
        }
        if (typeof node.vin === 'string' && /^[A-HJ-NPR-Z0-9]{11,17}$/i.test(node.vin.trim())) {
            out.push({
                vin: node.vin.trim().toUpperCase(),
                stockId: firstStock(node),
                modell: buildModelName(node),
                raw: node
            });
        }
        var keys = Object.keys(node);
        for (var k = 0; k < keys.length; k++) scanForVehicles(node[keys[k]], out, depth + 1);
    }

    function mergeCapture(hit, parent) {
        if (!hit.stockId && parent) hit.stockId = firstStock(parent);
        if (!hit.modell && parent) hit.modell = buildModelName(parent);
        return hit;
    }

    function onGraphqlData(url, data) {
        var hits = [];
        scanForVehicles(data, hits, 0);
        if (!hits.length) return;
        for (var i = 0; i < hits.length; i++) {
            var hit = hits[i];
            if (!hit.stockId || !hit.modell) {
                var domHit = extractVehicleData();
                if (!hit.stockId && domHit.vin === hit.vin) hit.stockId = domHit.stockId;
                if (!hit.modell && domHit.vin === hit.vin) hit.modell = domHit.modell;
            }
            var key = hit.stockId || hit.vin;
            if (!captureLog[key]) {
                captureLog[key] = true;
                console.log('%c[WM Bridge] Fahrzeug aus GraphQL:', 'color:#56d364', hit.stockId, hit.modell, hit.vin);
                console.log('[WM Bridge] Rohobjekt (für Schema-Analyse):', hit.raw);
            }
            lastCapture = { vin: hit.vin, stockId: hit.stockId, modell: hit.modell };
        }
        setStatus('GraphQL erkannt: ' + (lastCapture.stockId || lastCapture.vin), '#58a6ff');
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

    function collectData() {
        var data = extractVehicleData();
        if (lastCapture) {
            if (!data.vin) data.vin = lastCapture.vin;
            if (!data.stockId && lastCapture.stockId) data.stockId = lastCapture.stockId;
            if (!data.modell && lastCapture.modell) data.modell = lastCapture.modell;
            if (data.vin === lastCapture.vin) {
                if (!data.stockId) data.stockId = lastCapture.stockId;
                if (!data.modell) data.modell = lastCapture.modell;
            }
        }
        return data;
    }

    function sendToHud(data, onDone) {
        var url = gmGet(STORE_URL, '');
        var secret = gmGet(STORE_SECRET, '');
        if (!url) { onDone(false, 'Web-App URL fehlt (im Panel eintragen)'); return; }
        if (!secret) { onDone(false, 'Secret fehlt (im Panel eintragen)'); return; }
        if (!data.stockId || !data.vin || !data.modell) {
            onDone(false, 'Unvollständig: ' + JSON.stringify({ stockId: data.stockId, modell: data.modell, vin: data.vin }));
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

    function runGraphqlProbe() {
        console.log('%c[WM Bridge] GraphQL Introspection Test auf ' + GRAPHQL_URL, 'color:#f97316;font-weight:bold');
        fetch(GRAPHQL_URL, {
            method: 'POST',
            credentials: 'include',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ query: 'query { __schema { queryType { fields { name description } } } }' })
        }).then(function (r) {
            return r.json().then(function (d) { return { status: r.status, data: d }; });
        }).then(function (res) {
            if (res.data && res.data.data && res.data.data.__schema) {
                var fields = res.data.data.__schema.queryType.fields;
                console.log('%c[WM Bridge] Introspection AKTIV — verfügbare Queries (' + fields.length + '):', 'color:#56d364');
                fields.forEach(function (f) { console.log('  ' + f.name + (f.description ? ' — ' + f.description : '')); });
                setStatus('Introspection OK: ' + fields.length + ' Queries (Konsole)', '#56d364');
            } else {
                console.log('[WM Bridge] Introspection Antwort (HTTP ' + res.status + '):', res.data);
                setStatus('Introspection blockiert/Fehler — Konsole prüfen', '#f85149');
            }
        }).catch(function (err) {
            console.log('[WM Bridge] Introspection Fehler:', err);
            setStatus('Introspection Fehler — Konsole prüfen', '#f85149');
        });
        console.log('[WM Bridge] Alle bisher gesnifften Fahrzeuge:', captureLog);
    }

    var panel, statusEl;

    function buildPanel() {
        panel = document.createElement('div');
        panel.style.cssText = 'position:fixed;bottom:14px;right:14px;z-index:999999;background:#161b22;color:#e6edf3;'
            + 'border:1px solid #f97316;border-radius:10px;padding:10px 12px;font:12px/1.5 "Segoe UI",sans-serif;'
            + 'box-shadow:0 0 18px rgba(249,115,22,.35);width:250px;';
        panel.innerHTML = ''
            + '<div style="font-weight:700;margin-bottom:6px;color:#ffa94d;">Werkstattmappe Bridge <span style="color:#9da7b3;font-weight:400;">v0.2</span></div>'
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
            setStatus('Lese Daten…');
            sendToHud(collectData(), function (ok, msg) {
                setStatus(msg, ok ? '#56d364' : '#f85149');
            });
        });
        panel.querySelector('#wmb-probe').addEventListener('click', runGraphqlProbe);
    }

    function setStatus(msg, color) {
        if (!statusEl) return;
        statusEl.textContent = msg;
        statusEl.style.color = color || '#9da7b3';
    }

    var lastAutoSent = '';
    var STORE_FORCE_AUTO = 'wm_bridge_force_auto';

    function qsParam(name) {
        try {
            return new URL(location.href).searchParams.get(name) || '';
        } catch (e) {
            return '';
        }
    }

    function normalizeSid(v) {
        return String(v || '').toUpperCase().replace(/\s+/g, '').trim();
    }

    function wantsAutofill() {
        return qsParam('wm_autofill') === '1' || sessionStorage.getItem(STORE_FORCE_AUTO) === '1';
    }

    function markAutofillSession(stockId) {
        try {
            sessionStorage.setItem(STORE_FORCE_AUTO, '1');
            if (stockId) sessionStorage.setItem('wm_bridge_autofill_sid', normalizeSid(stockId));
        } catch (e) {}
    }

    function clearAutofillSession() {
        try {
            sessionStorage.removeItem(STORE_FORCE_AUTO);
            sessionStorage.removeItem('wm_bridge_autofill_sid');
        } catch (e) {}
    }

    function targetAutofillStock() {
        return normalizeSid(qsParam('rsv') || sessionStorage.getItem('wm_bridge_autofill_sid') || '');
    }

    function isDetailPage() {
        return /\/refurbishment\/[0-9a-f-]{20,}/i.test(location.href);
    }

    function isListPage() {
        return /\/refurbishment\/?(\?|$)/i.test(location.pathname + (location.search ? '?' : '')) ||
            /\/refurbishment\?/i.test(location.href);
    }

    function findStockResultLink(stockId) {
        var sid = normalizeSid(stockId);
        var links = Array.from(document.querySelectorAll('a[href*="/refurbishment/"]'));
        links = links.filter(function (a) {
            return /\/refurbishment\/[0-9a-f-]{20,}/i.test(a.href || '');
        });
        if (!links.length) return null;
        if (!sid) return links[0];
        for (var i = 0; i < links.length; i++) {
            var row = links[i].closest('tr, [role="row"], .rt-tr-group, [class*="row"]') || links[i].parentElement;
            var txt = ((row && row.textContent) || links[i].textContent || '').toUpperCase().replace(/\s+/g, '');
            if (txt.indexOf(sid) !== -1) return links[i];
        }
        return links[0];
    }

    function clickEl(el) {
        if (!el) return;
        try { el.scrollIntoView({ block: 'center' }); } catch (e) {}
        try { el.click(); } catch (e2) {}
        try {
            el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
        } catch (e3) {}
    }

    var autofillBusy = false;
    function runAutofillNavigation() {
        if (autofillBusy) return;
        if (!wantsAutofill()) return;
        var sid = targetAutofillStock();
        if (sid) markAutofillSession(sid);

        if (isDetailPage()) {
            setStatus('Autofill: Detailseite — sende Daten…', '#58a6ff');
            gmSet(STORE_AUTO, '1');
            var autoCheck = panel && panel.querySelector('#wmb-auto');
            if (autoCheck) autoCheck.checked = true;
            autoTick(true);
            return;
        }

        if (!/refurbishment/i.test(location.href)) return;
        autofillBusy = true;
        setStatus('Autofill: öffne ' + (sid || 'Fahrzeug') + '…', '#ffa94d');

        var tries = 0;
        var timer = setInterval(function () {
            tries++;
            if (isDetailPage()) {
                clearInterval(timer);
                autofillBusy = false;
                runAutofillNavigation();
                return;
            }
            var link = findStockResultLink(sid);
            if (link) {
                clearInterval(timer);
                setStatus('Autofill: treffer gefunden — öffne Auftrag…', '#56d364');
                clickEl(link);
                setTimeout(function () { autofillBusy = false; }, 1500);
                return;
            }
            if (tries >= 40) {
                clearInterval(timer);
                autofillBusy = false;
                setStatus('Autofill: kein Treffer für ' + sid + ' — bitte manuell öffnen', '#f85149');
            }
        }, 400);
    }

    function autoTick(force) {
        if (!force && gmGet(STORE_AUTO, '') !== '1' && !wantsAutofill()) return;
        if (!isDetailPage()) return;
        var data = collectData();
        if (!data.stockId || !data.vin || !data.modell) return;
        var want = targetAutofillStock();
        if (want && normalizeSid(data.stockId) !== want) return;
        if (data.stockId === lastAutoSent) return;
        lastAutoSent = data.stockId;
        setStatus('Auto-Senden: ' + data.stockId + '…');
        sendToHud(data, function (ok, msg) {
            setStatus(msg, ok ? '#56d364' : '#f85149');
            if (ok) clearAutofillSession();
            else lastAutoSent = '';
        });
    }

    window.addEventListener('message', function (ev) {
        var d = ev && ev.data;
        if (!d || !d[MSG_MARKER] || !d.data) return;
        try { onGraphqlData(d.url, d.data); } catch (e) {}
    });

    injectPageHook();

    function init() {
        if (!document.body) { setTimeout(init, 300); return; }
        buildPanel();
        var ver = panel.querySelector('div');
        if (ver) ver.innerHTML = ver.innerHTML.replace('v0.2', 'v0.3');
        if (wantsAutofill()) {
            var sid0 = targetAutofillStock();
            if (sid0) markAutofillSession(sid0);
            gmSet(STORE_AUTO, '1');
            var autoCheck = panel.querySelector('#wmb-auto');
            if (autoCheck) autoCheck.checked = true;
            setTimeout(runAutofillNavigation, 600);
        }
        setInterval(function () { autoTick(false); }, 2000);
        setInterval(function () {
            if (wantsAutofill() && !isDetailPage()) runAutofillNavigation();
        }, 3000);
    }

    init();
})();
