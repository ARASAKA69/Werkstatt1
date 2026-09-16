// ==UserScript==
// @name         Carol Retoure Bridge
// @namespace    retoure-lager
// @version      1.3
// @description  Öffnet den Carol-Auftrag, liest B2A1 / Fertiggestellt und sendet an Retoure Scan
// @match        *://carol.autohero.com/*
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @run-at       document-start
// ==/UserScript==

(function () {
    'use strict';

    var WEB_APP_URL = 'https://script.google.com/a/macros/auto1.com/s/AKfycbwsGB1o_1z0t9nCVXDx0lu3nQv8Ltj81Dgq5BVw8laLHPA4v4oLUpNvj-qx49iMjeVm/exec';
    var MSG = '__retoureCarolGql';
    var busy = false;
    var lastHref = '';
    var walkQueue = false;
    var gqlFlags = null;
    var gqlSid = '';

    function injectHook() {
        var src = '(function(){'
            + 'if(window.__retoureCarolHooked)return;window.__retoureCarolHooked=true;'
            + 'var of=window.fetch;'
            + 'window.fetch=function(){var a=arguments;'
            + 'var u=String(a[0]&&a[0].url?a[0].url:a[0]||"");'
            + 'var p=of.apply(this,a);'
            + 'if(u.indexOf("graphql")!==-1){p.then(function(r){'
            + 'try{r.clone().json().then(function(d){'
            + 'window.postMessage({' + MSG + ':1,data:d},"*");'
            + '}).catch(function(){});}catch(e){}return r;});}'
            + 'return p;};'
            + '})();';
        try {
            var s = document.createElement('script');
            s.textContent = src;
            (document.head || document.documentElement).appendChild(s);
            s.remove();
        } catch (e) {}
        try {
            if (typeof unsafeWindow !== 'undefined' && unsafeWindow.eval && !unsafeWindow.__retoureCarolHooked) {
                unsafeWindow.eval(src);
            }
        } catch (e2) {}
    }

    function qs(name) {
        try {
            return new URL(location.href).searchParams.get(name) || '';
        } catch (e) {
            return '';
        }
    }

    function batchId() {
        var fromQs = String(qs('retoure_batch') || '').trim();
        if (fromQs) {
            try { sessionStorage.setItem('retoure_batch', fromQs); } catch (e) {}
            return fromQs;
        }
        try { return String(sessionStorage.getItem('retoure_batch') || '').trim(); } catch (e2) { return ''; }
    }

    function pageText() {
        return String((document.body && document.body.innerText) || '').replace(/\u00a0/g, ' ');
    }

    function sidFromPage() {
        var fromQs = String(qs('rsv') || '').replace(/\s+/g, '').toUpperCase();
        if (/^[A-Z]{2}\d{4,8}$/.test(fromQs)) return fromQs;
        try {
            var stored = String(sessionStorage.getItem('retoure_sid') || '').replace(/\s+/g, '').toUpperCase();
            if (/^[A-Z]{2}\d{4,8}$/.test(stored) && isDetailPage()) return stored;
        } catch (e) {}
        if (gqlSid && /^[A-Z]{2}\d{4,8}$/.test(gqlSid)) return gqlSid;
        var nodes = document.querySelectorAll('h1, h2, h3, h4, [class*="title"], [class*="Title"], [class*="heading"]');
        for (var i = 0; i < nodes.length; i++) {
            var m = String(nodes[i].textContent || '').trim().match(/\b([A-Z]{2}\d{4,8})\b/i);
            if (m) return m[1].toUpperCase();
        }
        var t = pageText();
        var m2 = t.match(/\bStock(?:\s*Number)?[:\s]+([A-Z]{2}\d{4,8})\b/i) || t.match(/\b([A-Z]{2}\d{4,8})\s*[-–]/);
        if (m2) return m2[1].toUpperCase();
        var m3 = t.match(/\b([A-Z]{2}\d{4,8})\b/);
        return m3 ? m3[1].toUpperCase() : '';
    }

    function rememberSid(sid) {
        if (/^[A-Z]{2}\d{4,8}$/.test(sid)) {
            try { sessionStorage.setItem('retoure_sid', sid); } catch (e) {}
        }
        return sid;
    }

    function isDetailPage() {
        return /\/refurbishment\/[0-9a-f-]{8,}/i.test(location.pathname);
    }

    function clickEl(el) {
        if (!el) return;
        try { el.scrollIntoView({ block: 'center' }); } catch (e) {}
        try { el.click(); } catch (e2) {}
        try { el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window })); } catch (e3) {}
    }

    function fillSearch(sid) {
        var want = String(sid || '');
        if (!want) return false;
        var inputs = document.querySelectorAll('input');
        for (var i = 0; i < inputs.length; i++) {
            var inp = inputs[i];
            var typ = String(inp.type || 'text').toLowerCase();
            if (typ === 'hidden' || typ === 'checkbox' || typ === 'radio') continue;
            var ph = String(inp.placeholder || inp.name || inp.getAttribute('aria-label') || '').toLowerCase();
            if (ph && ph.indexOf('vin') === -1 && ph.indexOf('stock') === -1 && ph.indexOf('refurb') === -1 && ph.indexOf('suche') === -1) continue;
            try {
                var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value');
                if (setter && setter.set) setter.set.call(inp, want);
                else inp.value = want;
                inp.dispatchEvent(new Event('input', { bubbles: true }));
                inp.dispatchEvent(new Event('change', { bubbles: true }));
                inp.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', keyCode: 13, bubbles: true }));
                inp.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', keyCode: 13, bubbles: true }));
            } catch (e) {}
            var btns = document.querySelectorAll('button');
            for (var b = 0; b < btns.length; b++) {
                var lab = String(btns[b].textContent || '').trim().toLowerCase();
                if (lab === 'filter' || lab === 'suchen' || lab === 'search') {
                    clickEl(btns[b]);
                    break;
                }
            }
            return true;
        }
        return false;
    }

    function clickStockLink(sid) {
        var want = String(sid || '').toUpperCase();
        var links = document.querySelectorAll('a[href*="/refurbishment/"]');
        for (var i = 0; i < links.length; i++) {
            var href = links[i].getAttribute('href') || links[i].href || '';
            if (!/\/refurbishment\/[0-9a-f-]{8,}/i.test(href)) continue;
            var row = links[i].closest('article, li, tr, [role="row"], [class*="card"], [class*="Card"], [class*="row"]') || links[i];
            var txt = ((row && row.textContent) || links[i].textContent || '').toUpperCase().replace(/\s+/g, '');
            if (!want || txt.indexOf(want) !== -1) {
                clickEl(links[i]);
                return true;
            }
        }
        var all = document.querySelectorAll('a, button, [role="link"], article, [class*="card"], [class*="Card"]');
        for (var j = 0; j < all.length; j++) {
            var t = String(all[j].textContent || '').toUpperCase();
            if (!want || t.indexOf(want) === -1) continue;
            if (t.length > 400) continue;
            var go = all[j].closest('a') || all[j].querySelector('a[href*="/refurbishment/"]') || all[j];
            clickEl(go);
            return true;
        }
        return false;
    }

    function flagsFromText(t) {
        t = String(t || '');
        var b2a1 = /als\s*b2a1\s*markiert|flagged\s*for\s*return(\s*to\s*auto\s*1)?|return\s*to\s*auto\s*1/i.test(t);
        var fertig = /fertiggestellt(\s+am)?|completed\s+on\s+\d|handed\s*out|herausgegeben/i.test(t);
        var label = '';
        var m = t.match(/als\s*b2a1\s*markiert[^\n]{0,60}/i)
            || t.match(/flagged\s*for\s*return[^\n]{0,72}/i)
            || t.match(/fertiggestellt[^\n]{0,48}/i)
            || t.match(/completed\s+on[^\n]{0,48}/i);
        if (m) label = String(m[0] || '').replace(/\s+/g, ' ').trim();
        if (b2a1) fertig = false;
        return { b2a1: b2a1, fertig: fertig, label: label };
    }

    function readFlags() {
        var fromDom = flagsFromText(pageText());
        if (!fromDom.b2a1 && !fromDom.fertig && gqlFlags) {
            return gqlFlags;
        }
        return fromDom;
    }

    function scanGql(data) {
        var s = '';
        try { s = JSON.stringify(data); } catch (e) { return; }
        var sidM = s.match(/"(?:stockNumber|stockId|stock_number)"\s*:\s*"([A-Z]{2}\d{4,8})"/i);
        if (sidM) gqlSid = sidM[1].toUpperCase();
        var flags = flagsFromText(s);
        if (flags.b2a1 || flags.fertig) gqlFlags = flags;
    }

    window.addEventListener('message', function (ev) {
        if (!ev || !ev.data || !ev.data[MSG]) return;
        scanGql(ev.data.data);
    });

    function parseLoose(txt) {
        var s = String(txt || '');
        try { return JSON.parse(s); } catch (e) {}
        var pre = s.match(/<pre[^>]*>([\s\S]*?)<\/pre>/i);
        if (pre) {
            var inner = pre[1].replace(/&quot;/g, '"').replace(/&#34;/g, '"').replace(/&amp;/g, '&');
            try { return JSON.parse(inner); } catch (e2) {}
        }
        var m = s.match(/\{[\s\S]*\}/);
        if (m) {
            try { return JSON.parse(m[0]); } catch (e3) {}
        }
        return null;
    }

    function gmReq(method, url, body, onOk, onErr) {
        var done = false;
        function finishOk(data) {
            if (done) return;
            done = true;
            onOk(data);
        }
        function finishErr(msg) {
            if (done) return;
            done = true;
            onErr(msg);
        }
        if (typeof GM_xmlhttpRequest === 'function') {
            GM_xmlhttpRequest({
                method: method,
                url: url,
                data: body || undefined,
                headers: method === 'POST' ? { 'Content-Type': 'text/plain;charset=utf-8' } : {},
                anonymous: false,
                timeout: 8000,
                onload: function (res) {
                    var data = parseLoose(res.responseText);
                    if (!data) finishErr('kein JSON');
                    else finishOk(data);
                },
                onerror: function () { finishErr('netzwerk'); },
                ontimeout: function () { finishErr('timeout'); }
            });
            return;
        }
        var opt = { method: method, credentials: 'include' };
        if (method === 'POST') {
            opt.headers = { 'Content-Type': 'text/plain;charset=utf-8' };
            opt.body = body || '';
        }
        fetch(url, opt).then(function (r) { return r.text(); }).then(function (txt) {
            var data = parseLoose(txt);
            if (!data) finishErr('kein JSON');
            else finishOk(data);
        }).catch(function () { finishErr('netzwerk'); });
        setTimeout(function () { finishErr('timeout'); }, 8000);
    }

    function gmPost(body, onOk, onErr) {
        gmReq('POST', WEB_APP_URL, JSON.stringify(body || {}), onOk, function (err) {
            var q = WEB_APP_URL + '?page=' + encodeURIComponent(body.page || '')
                + '&batch=' + encodeURIComponent(body.batch || '')
                + '&sid=' + encodeURIComponent(body.sid || '')
                + '&b2a1=' + encodeURIComponent(body.b2a1 || '0')
                + '&fertig=' + encodeURIComponent(body.fertig || '0')
                + '&label=' + encodeURIComponent(body.label || '')
                + '&detail=' + encodeURIComponent(body.detail || '0');
            gmReq('GET', q, '', onOk, onErr);
        });
    }

    function toast(msg) {
        if (!document.body) return;
        var el = document.getElementById('retoure-carol-toast');
        if (!el) {
            el = document.createElement('div');
            el.id = 'retoure-carol-toast';
            el.style.cssText = 'position:fixed;bottom:18px;left:18px;z-index:999999;padding:12px 16px;border-radius:12px;font:13px/1.4 Segoe UI,sans-serif;background:#0f1720;color:#e6edf3;border:1px solid #2dd4bf;box-shadow:0 8px 24px rgba(0,0,0,.45);max-width:380px;';
            el.innerHTML = '<div style="font-weight:700;color:#2dd4bf;margin-bottom:4px;">Retoure Scan</div><div id="retoure-carol-toast-msg"></div>';
            document.body.appendChild(el);
        }
        var msgEl = document.getElementById('retoure-carol-toast-msg');
        if (msgEl) msgEl.textContent = msg;
        else el.textContent = 'Retoure Scan · ' + msg;
        el.style.display = 'block';
    }

    function langPrefix() {
        var m = location.pathname.match(/^\/([a-z]{2}-[A-Z]{2})/);
        return m ? m[1] : 'de-DE';
    }

    function withBatch(url, batch) {
        var u = String(url || '');
        if (!u) return '';
        if (batch && u.indexOf('retoure_batch=') === -1) u += (u.indexOf('?') >= 0 ? '&' : '?') + 'retoure_batch=' + encodeURIComponent(batch);
        return u;
    }

    function goNext(batch, sid, url) {
        if (url && /\/refurbishment\/[0-9a-f-]{8,}/i.test(url)) {
            location.assign(withBatch(url, batch));
            return;
        }
        if (!batch || !sid) return;
        location.assign(location.origin + '/' + langPrefix() + '/refurbishment?rsv=' + encodeURIComponent(sid) + '&retoure_batch=' + encodeURIComponent(batch));
    }

    function finishWalk(msg) {
        toast(msg || 'Carol-Check fertig');
        busy = false;
        if (walkQueue) {
            setTimeout(function () {
                try { window.close(); } catch (e) {}
            }, 600);
        }
    }

    function sendFlags(batch, sid, flags) {
        rememberSid(sid);
        toast(sid + (flags.b2a1 ? ' · B2A1' : (flags.fertig ? ' · Fertig' : ' · kein Badge')) + ' — sende…');
        gmPost({
            page: 'carolreport',
            batch: batch || '',
            sid: sid,
            b2a1: flags.b2a1 ? '1' : '0',
            fertig: flags.fertig ? '1' : '0',
            label: flags.label || '',
            detail: '1'
        }, function (res) {
            if (res && res.batchId) {
                try { sessionStorage.setItem('retoure_batch', res.batchId); } catch (e) {}
            }
            var next = walkQueue && res && res.nextId;
            if (next) {
                toast(sid + ' ok · weiter ' + next);
                busy = false;
                goNext(res.batchId || batch, next, res.carolUrl || '');
                return;
            }
            finishWalk((res && res.message) || 'Carol-Check fertig');
        }, function (err) {
            toast('Report fehlgeschlagen · ' + err);
            busy = false;
        });
    }

    function waitAndReport(batch, sid) {
        rememberSid(sid);
        var tries = 0;
        var searched = false;
        toast('Öffne Auftrag ' + sid + '…');
        var timer = setInterval(function () {
            tries++;
            var liveSid = sidFromPage() || sid;
            if (liveSid) sid = rememberSid(liveSid);
            if (!isDetailPage()) {
                if (tries === 3 || tries === 10) fillSearch(sid);
                clickStockLink(sid);
                searched = true;
                if (tries < 45) return;
                toast(sid + ' · Auftrag nicht offen — klicke das Fahrzeug');
                return;
            }
            var flags = readFlags();
            if (!flags.b2a1 && !flags.fertig && tries < 28) return;
            clearInterval(timer);
            sendFlags(batch, sid, flags);
        }, 250);
    }

    function start() {
        if (qs('retoure_auto') === '1' && !qs('retoure_batch') && !qs('rsv') && !isDetailPage()) return;
        if (!/\/refurbishment/i.test(location.pathname + location.search)) return;
        walkQueue = !!qs('retoure_batch') || !!batchId();
        if (!walkQueue && !isDetailPage() && !qs('rsv')) return;
        if (busy) return;
        busy = true;
        var tries = 0;
        function kick() {
            tries++;
            var sid = sidFromPage();
            var batch = batchId();
            if (!sid && tries < 24) {
                toast('Warte auf Auftrag…');
                setTimeout(kick, 250);
                return;
            }
            if (sid) rememberSid(sid);
            toast(sid ? ('Lese ' + sid + '…') : 'Carol-Queue…');
            if (sid) {
                waitAndReport(batch, sid);
                return;
            }
            gmPost({ page: 'carolq', batch: batch || '' }, function (res) {
                if (res && res.batchId) {
                    try { sessionStorage.setItem('retoure_batch', res.batchId); } catch (e) {}
                    batch = res.batchId;
                }
                if (!res || !res.success || !res.ids || !res.ids.length) {
                    finishWalk((res && res.message) || 'Carol-Check fertig');
                    return;
                }
                busy = false;
                goNext(batch, res.nextId || res.ids[0], res.carolUrl || '');
            }, function (err) {
                finishWalk('Queue fehlgeschlagen · ' + err);
            });
        }
        kick();
    }

    injectHook();
    function boot() {
        lastHref = location.href;
        start();
        setInterval(function () {
            if (location.href === lastHref) return;
            lastHref = location.href;
            busy = false;
            gqlFlags = null;
            start();
        }, 400);
    }
    if (document.readyState === 'loading') window.addEventListener('DOMContentLoaded', boot);
    else boot();
})();
