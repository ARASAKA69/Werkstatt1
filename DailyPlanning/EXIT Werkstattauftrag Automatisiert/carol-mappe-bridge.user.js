// ==UserScript==
// @name         Carol Werkstattmappe Bridge
// @namespace    arasaka
// @version      0.4
// @description  Sendet Modell + VIN aus Carol automatisch an das EXIT Werkstattmappe HUD
// @match        *://carol.autohero.com/*
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @run-at       document-start
// ==/UserScript==

(function () {
    'use strict';

    var HUD_WEB_APP_URL = 'https://script.google.com/a/macros/auto1.com/s/AKfycbxAz9jQS2cpMcyaDr86zBx7pVaY0hxzdp7rupT_wcfxqU4jgwmvhvv-WtX0D7JcE1JsXA/exec';
    var BRIDGE_SECRET = 'ARASAKA69';
    var MSG_MARKER = '__wmBridgeGraphql';
    var STORE_FORCE_AUTO = 'wm_bridge_force_auto';
    var STORE_AUTOFILL_SID = 'wm_bridge_autofill_sid';

    var lastCapture = null;
    var lastAutoSent = '';
    var autofillBusy = false;

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
            lastCapture = { vin: hit.vin, stockId: hit.stockId, modell: hit.modell };
        }
        autoTick(true);
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
        if (!data.stockId || !data.vin || !data.modell) {
            onDone(false, 'unvollständig');
            return;
        }
        var payload = JSON.stringify({
            secret: BRIDGE_SECRET,
            entries: [{ stockId: data.stockId, modell: data.modell, vin: data.vin }]
        });
        if (typeof GM_xmlhttpRequest === 'function') {
            GM_xmlhttpRequest({
                method: 'POST',
                url: HUD_WEB_APP_URL,
                data: payload,
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                onload: function (res) {
                    var ok = false;
                    try {
                        var parsed = JSON.parse(res.responseText);
                        ok = !!parsed.success;
                    } catch (e) {}
                    onDone(ok, ok ? data.stockId : ('HTTP ' + res.status));
                },
                onerror: function () { onDone(false, 'netzwerk'); }
            });
        } else {
            fetch(HUD_WEB_APP_URL, {
                method: 'POST',
                mode: 'no-cors',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: payload
            }).then(function () {
                onDone(true, data.stockId);
            }).catch(function () {
                onDone(false, 'netzwerk');
            });
        }
    }

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
            if (stockId) sessionStorage.setItem(STORE_AUTOFILL_SID, normalizeSid(stockId));
        } catch (e) {}
    }

    function clearAutofillSession() {
        try {
            sessionStorage.removeItem(STORE_FORCE_AUTO);
            sessionStorage.removeItem(STORE_AUTOFILL_SID);
        } catch (e) {}
    }

    function targetAutofillStock() {
        return normalizeSid(qsParam('rsv') || sessionStorage.getItem(STORE_AUTOFILL_SID) || '');
    }

    function isDetailPage() {
        return /\/refurbishment\/[0-9a-f-]{20,}/i.test(location.href);
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

    function runAutofillNavigation() {
        if (autofillBusy) return;
        if (!wantsAutofill()) return;
        var sid = targetAutofillStock();
        if (sid) markAutofillSession(sid);

        if (isDetailPage()) {
            autoTick(true);
            return;
        }

        if (!/refurbishment/i.test(location.href)) return;
        autofillBusy = true;

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
                clickEl(link);
                setTimeout(function () { autofillBusy = false; }, 1500);
                return;
            }
            if (tries >= 40) {
                clearInterval(timer);
                autofillBusy = false;
            }
        }, 400);
    }

    function autoTick(force) {
        if (!force && !wantsAutofill() && !isDetailPage()) return;
        if (!isDetailPage()) return;
        var data = collectData();
        if (!data.stockId || !data.vin || !data.modell) return;
        var want = targetAutofillStock();
        if (want && normalizeSid(data.stockId) !== want) return;
        if (data.stockId === lastAutoSent) return;
        lastAutoSent = data.stockId;
        sendToHud(data, function (ok) {
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
        if (wantsAutofill()) {
            var sid0 = targetAutofillStock();
            if (sid0) markAutofillSession(sid0);
            setTimeout(runAutofillNavigation, 600);
        }
        setInterval(function () { autoTick(true); }, 2000);
        setInterval(function () {
            if (wantsAutofill() && !isDetailPage()) runAutofillNavigation();
        }, 3000);
    }

    init();
})();
