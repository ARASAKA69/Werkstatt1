// ==UserScript==
// @name         Carol Retoure Bridge
// @namespace    retoure-lager
// @version      1.6
// @description  Liest nur sichtbare Carol-Badges (B2A1 / Fertiggestellt) und sendet an Retoure Scan
// @match        *://carol.autohero.com/*
// @grant        GM_xmlhttpRequest
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';

    var VER = '1.6';
    if (window.__retoureBridgeVer) {
        return;
    }
    window.__retoureBridgeVer = VER;

    var WEB_APP_URL = 'https://script.google.com/a/macros/auto1.com/s/AKfycbwsGB1o_1z0t9nCVXDx0lu3nQv8Ltj81Dgq5BVw8laLHPA4v4oLUpNvj-qx49iMjeVm/exec';
    var MAX_TRIES = 34;
    var busy = false;
    var lastKey = '';
    var timer = null;

    function qs(name) {
        try {
            return new URL(location.href).searchParams.get(name) || '';
        } catch (e) {
            return '';
        }
    }

    function ssGet(key) {
        try { return String(sessionStorage.getItem(key) || ''); } catch (e) { return ''; }
    }

    function ssSet(key, val) {
        try { sessionStorage.setItem(key, String(val)); } catch (e) {}
    }

    function batchId() {
        var fromQs = String(qs('retoure_batch') || '').trim();
        if (fromQs) {
            ssSet('retoure_batch', fromQs);
            return fromQs;
        }
        return ssGet('retoure_batch').trim();
    }

    function isDetailPage() {
        return /\/refurbishment\/[0-9a-f-]{8,}/i.test(location.pathname);
    }

    function pageText() {
        return String((document.body && document.body.innerText) || '').replace(/\u00a0/g, ' ');
    }

    function detailSid() {
        var nodes = document.querySelectorAll('h1, h2, h3, h4, h5');
        for (var i = 0; i < nodes.length; i++) {
            var m = String(nodes[i].textContent || '').trim().match(/^([A-Z]{2}\d{4,8})\s*[-–]/);
            if (m) return m[1].toUpperCase();
        }
        var m2 = pageText().match(/\b([A-Z]{2}\d{4,8})\s*[-–]\s*\S/);
        return m2 ? m2[1].toUpperCase() : '';
    }

    function wantedSid() {
        var fromQs = String(qs('rsv') || '').replace(/\s+/g, '').toUpperCase();
        if (/^[A-Z]{2}\d{4,8}$/.test(fromQs)) return fromQs;
        return '';
    }

    function doneKey(batch) {
        return 'retoure_done_' + (batch || 'x');
    }

    function isDone(batch, sid) {
        return ssGet(doneKey(batch)).indexOf('|' + sid + '|') !== -1;
    }

    function markDone(batch, sid) {
        var cur = ssGet(doneKey(batch));
        if (cur.indexOf('|' + sid + '|') === -1) ssSet(doneKey(batch), cur + '|' + sid + '|');
    }

    function clickEl(el) {
        if (!el) return;
        try { el.scrollIntoView({ block: 'center' }); } catch (e) {}
        try { el.click(); } catch (e2) {}
        try { el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window })); } catch (e3) {}
    }

    function fillSearch(sid) {
        if (!sid) return false;
        var inputs = document.querySelectorAll('input');
        for (var i = 0; i < inputs.length; i++) {
            var inp = inputs[i];
            var typ = String(inp.type || 'text').toLowerCase();
            if (typ === 'hidden' || typ === 'checkbox' || typ === 'radio') continue;
            var ph = String(inp.placeholder || inp.name || inp.getAttribute('aria-label') || '').toLowerCase();
            if (ph && ph.indexOf('vin') === -1 && ph.indexOf('stock') === -1 && ph.indexOf('refurb') === -1 && ph.indexOf('suche') === -1) continue;
            try {
                var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value');
                if (setter && setter.set) setter.set.call(inp, sid);
                else inp.value = sid;
                inp.dispatchEvent(new Event('input', { bubbles: true }));
                inp.dispatchEvent(new Event('change', { bubbles: true }));
            } catch (e) {}
            return true;
        }
        return false;
    }

    function clickStockLink(sid) {
        var want = String(sid || '').toUpperCase();
        if (!want) return false;
        var links = document.querySelectorAll('a[href*="/refurbishment/"]');
        for (var i = 0; i < links.length; i++) {
            var href = links[i].getAttribute('href') || links[i].href || '';
            if (!/\/refurbishment\/[0-9a-f-]{8,}/i.test(href)) continue;
            var row = links[i].closest('article, li, tr, [class*="card"], [class*="row"]') || links[i];
            var txt = String((row && row.textContent) || '').toUpperCase().replace(/\s+/g, '');
            if (txt.indexOf(want) !== -1) {
                clickEl(links[i]);
                return true;
            }
        }
        return false;
    }

    function isB2a1Chip(t) {
        if (/refurbishment\s+started/i.test(t)) return false;
        return /als\s+b2a1\s+markiert/i.test(t) || /flagged\s+for\s+return\s+to\s+auto\s*1/i.test(t);
    }

    function isFertigChip(t) {
        if (/refurbishment\s+started/i.test(t)) return false;
        return /^fertiggestellt(\s+am\b|$)/i.test(t) || /^completed\s+on\s+\d/i.test(t);
    }

    function readFlags() {
        var nodes = document.querySelectorAll('span, div, p, li, small, strong, [class*="badge"], [class*="chip"], [class*="tag"], [class*="pill"], [class*="label"]');
        var b2a1 = '';
        var fertig = '';
        for (var i = 0; i < nodes.length; i++) {
            if (nodes[i].children && nodes[i].children.length) continue;
            var t = String(nodes[i].textContent || '').replace(/\s+/g, ' ').trim();
            if (!t || t.length > 70) continue;
            if (!b2a1 && isB2a1Chip(t)) b2a1 = t;
            if (!fertig && isFertigChip(t)) fertig = t;
        }
        if (b2a1) return { b2a1: true, fertig: false, label: b2a1 };
        if (fertig) return { b2a1: false, fertig: true, label: fertig };
        return { b2a1: false, fertig: false, label: '' };
    }

    function looksLikeAnswer(o) {
        return !!(o && typeof o === 'object' && (o.success === true || o.success === false || o.batchId || o.nextId || o.pending != null));
    }

    function parseLoose(txt) {
        var s = String(txt || '').trim();
        if (s.charAt(0) === '{') {
            try {
                var o = JSON.parse(s);
                if (looksLikeAnswer(o)) return o;
            } catch (e) {}
        }
        var pre = s.match(/<pre[^>]*>([\s\S]*?)<\/pre>/i);
        if (pre) {
            try {
                var inner = JSON.parse(pre[1].replace(/&quot;/g, '"').replace(/&amp;/g, '&').replace(/&lt;/g, '<'));
                if (looksLikeAnswer(inner)) return inner;
            } catch (e2) {}
        }
        return null;
    }

    function echoUrl(txt) {
        var m = String(txt || '').replace(/&amp;/g, '&').match(/https:\/\/script\.googleusercontent\.com\/macros\/echo[^"'\\\s<>]*/);
        return m ? m[0] : '';
    }

    function gmReq(method, url, body, onOk, onErr, hop) {
        hop = hop || 0;
        var done = false;
        function ok(d) { if (!done) { done = true; onOk(d); } }
        function err(m) { if (!done) { done = true; onErr(m); } }
        if (typeof GM_xmlhttpRequest !== 'function') {
            err('kein GM');
            return;
        }
        GM_xmlhttpRequest({
            method: method,
            url: url,
            data: body || undefined,
            headers: method === 'POST' ? { 'Content-Type': 'text/plain;charset=utf-8' } : {},
            anonymous: false,
            timeout: 20000,
            onload: function (res) {
                var txt = res && res.responseText != null ? res.responseText : '';
                var data = parseLoose(txt);
                if (data) {
                    ok(data);
                    return;
                }
                var jump = hop < 2 ? echoUrl(txt) : '';
                if (jump) {
                    done = true;
                    gmReq('GET', jump, '', onOk, onErr, hop + 1);
                    return;
                }
                if (res && res.status >= 200 && res.status < 400) ok({ success: true, accepted: true });
                else err('HTTP ' + (res && res.status));
            },
            onerror: function () { err('netzwerk'); },
            ontimeout: function () { err('timeout'); }
        });
    }

    function post(body, onOk, onErr) {
        gmReq('POST', WEB_APP_URL, JSON.stringify(body || {}), onOk, function () {
            var q = WEB_APP_URL + '?page=' + encodeURIComponent(body.page || '')
                + '&batch=' + encodeURIComponent(body.batch || '')
                + '&sid=' + encodeURIComponent(body.sid || '')
                + '&b2a1=' + encodeURIComponent(body.b2a1 || '0')
                + '&fertig=' + encodeURIComponent(body.fertig || '0')
                + '&notfound=' + encodeURIComponent(body.notfound || '0')
                + '&label=' + encodeURIComponent(body.label || '')
                + '&detail=' + encodeURIComponent(body.detail || '0');
            gmReq('GET', q, '', onOk, onErr);
        });
    }

    function toast(msg) {
        if (!document.body) return;
        var el = document.getElementById('retoure-carol-toast-' + VER);
        if (!el) {
            el = document.createElement('div');
            el.id = 'retoure-carol-toast-' + VER;
            el.style.cssText = 'position:fixed;bottom:18px;left:18px;z-index:2147483647;padding:12px 16px;border-radius:12px;font:13px/1.4 Segoe UI,sans-serif;background:#0f1720;color:#e6edf3;border:1px solid #2dd4bf;box-shadow:0 8px 24px rgba(0,0,0,.45);max-width:380px;';
            el.innerHTML = '<div style="font-weight:700;color:#2dd4bf;margin-bottom:4px;">Retoure Bridge v' + VER + '</div><div class="m"></div>';
            document.body.appendChild(el);
        }
        var m = el.querySelector('.m');
        if (m) m.textContent = msg;
    }

    function langPrefix() {
        var m = location.pathname.match(/^\/([a-z]{2}-[A-Z]{2})/);
        return m ? m[1] : 'de-DE';
    }

    function goTo(batch, sid, url) {
        if (url && /\/refurbishment\/[0-9a-f-]{8,}/i.test(url)) {
            var u = url;
            if (batch && u.indexOf('retoure_batch=') === -1) u += (u.indexOf('?') >= 0 ? '&' : '?') + 'retoure_batch=' + encodeURIComponent(batch);
            location.assign(u);
            return;
        }
        if (!sid) return;
        location.assign(location.origin + '/' + langPrefix() + '/refurbishment?rsv=' + encodeURIComponent(sid)
            + (batch ? '&retoure_batch=' + encodeURIComponent(batch) : ''));
    }

    function askQueue(batch, sid) {
        post({ page: 'carolq', batch: batch || '' }, function (res) {
            if (res && res.batchId) ssSet('retoure_batch', res.batchId);
            var next = res && res.nextId;
            if (!next) {
                toast('Carol-Check fertig');
                busy = false;
                return;
            }
            if (next === sid || isDone(res.batchId || batch, next)) {
                toast(next + ' bleibt offen — Scan-Seite übernimmt');
                busy = false;
                return;
            }
            toast('weiter · ' + next);
            busy = false;
            goTo(res.batchId || batch, next, res.carolUrl || '');
        }, function (e) {
            toast('Queue: ' + e);
            busy = false;
        });
    }

    function send(batch, sid, flags, notfound) {
        markDone(batch, sid);
        toast(sid + ' · ' + (notfound ? 'kein Auftrag' : (flags.b2a1 ? 'B2A1' : (flags.fertig ? 'Fertiggestellt' : 'kein Badge'))) + ' — sende…');
        post({
            page: 'carolreport',
            batch: batch || '',
            sid: sid,
            b2a1: flags.b2a1 ? '1' : '0',
            fertig: flags.fertig ? '1' : '0',
            notfound: notfound ? '1' : '0',
            label: flags.label || '',
            detail: '1'
        }, function (res) {
            if (res && res.batchId) ssSet('retoure_batch', res.batchId);
            var next = res && res.nextId;
            if (next && next !== sid && !isDone(res.batchId || batch, next)) {
                toast(sid + ' ok · weiter ' + next);
                busy = false;
                goTo(res.batchId || batch, next, res.carolUrl || '');
                return;
            }
            if (res && res.accepted) {
                askQueue(batch, sid);
                return;
            }
            toast(sid + ' ok · ' + ((res && res.pending) || 0) + ' offen');
            busy = false;
        }, function (e) {
            toast(sid + ' · Report ' + e);
            busy = false;
        });
    }

    function run() {
        var batch = batchId();
        var want = wantedSid();
        var tries = 0;
        if (timer) clearInterval(timer);
        timer = setInterval(function () {
            tries++;
            if (!isDetailPage()) {
                if (tries === 2 || tries === 12) fillSearch(want);
                clickStockLink(want);
                if (tries < MAX_TRIES) return;
                clearInterval(timer);
                if (want) send(batch, want, { b2a1: false, fertig: false, label: '' }, true);
                else {
                    busy = false;
                    toast('Kein Auftrag gefunden');
                }
                return;
            }
            var sid = detailSid();
            if (!sid) {
                if (tries < MAX_TRIES) return;
                clearInterval(timer);
                busy = false;
                toast('Stock-ID nicht gefunden');
                return;
            }
            if (want && sid !== want && tries < 14) return;
            if (isDone(batch, sid)) {
                clearInterval(timer);
                askQueue(batch, sid);
                return;
            }
            var flags = readFlags();
            if (!flags.b2a1 && !flags.fertig && tries < 18) return;
            clearInterval(timer);
            send(batch, sid, flags, false);
        }, 300);
    }

    function start() {
        if (!/\/refurbishment/i.test(location.pathname + location.search)) return;
        if (!isDetailPage() && !wantedSid()) return;
        if (busy) return;
        busy = true;
        toast(isDetailPage() ? 'Lese Auftrag…' : ('Öffne ' + wantedSid() + '…'));
        run();
    }

    lastKey = location.pathname + location.search;
    start();
    setInterval(function () {
        var key = location.pathname + location.search;
        if (key === lastKey) return;
        lastKey = key;
        busy = false;
        start();
    }, 400);
})();
