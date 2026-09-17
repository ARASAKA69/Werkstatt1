// ==UserScript==
// @name         Carol Retoure Bridge
// @namespace    retoure-lager
// @version      1.5
// @description  Liest nur sichtbare Carol-Badges (B2A1 / Fertiggestellt) und sendet an Retoure Scan
// @match        *://carol.autohero.com/*
// @grant        GM_xmlhttpRequest
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';

    var WEB_APP_URL = 'https://script.google.com/a/macros/auto1.com/s/AKfycbwsGB1o_1z0t9nCVXDx0lu3nQv8Ltj81Dgq5BVw8laLHPA4v4oLUpNvj-qx49iMjeVm/exec';
    var VER = '1.5';
    var busy = false;
    var lastHref = '';
    var walkQueue = false;

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
        var nodes = document.querySelectorAll('h1, h2, h3, h4, [class*="title"], [class*="Title"], [class*="heading"]');
        for (var i = 0; i < nodes.length; i++) {
            var m = String(nodes[i].textContent || '').trim().match(/\b([A-Z]{2}\d{4,8})\b/i);
            if (m) return m[1].toUpperCase();
        }
        var t = pageText();
        var m2 = t.match(/\b([A-Z]{2}\d{4,8})\s*[-–]/);
        return m2 ? m2[1].toUpperCase() : '';
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
            } catch (e) {}
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
            var row = links[i].closest('article, li, tr, [class*="card"], [class*="row"]') || links[i];
            var txt = ((row && row.textContent) || '').toUpperCase().replace(/\s+/g, '');
            if (!want || txt.indexOf(want) !== -1) {
                clickEl(links[i]);
                return true;
            }
        }
        return false;
    }

    function isB2a1Chip(t) {
        t = String(t || '').replace(/\s+/g, ' ').trim();
        if (/refurbishment\s+started/i.test(t)) return false;
        if (/auto-?ordered/i.test(t)) return false;
        return /als\s+b2a1\s+markiert/i.test(t) || /flagged\s+for\s+return\s+to\s+auto\s*1/i.test(t);
    }

    function isFertigChip(t) {
        t = String(t || '').replace(/\s+/g, ' ').trim();
        if (/refurbishment\s+started/i.test(t)) return false;
        return /fertiggestellt(\s+am)?/i.test(t) || /^completed\s+on\s+\d/i.test(t);
    }

    function readFlags() {
        var chips = [];
        var nodes = document.querySelectorAll('span, div, button, a, p, li, [class*="badge"], [class*="chip"], [class*="tag"], [class*="pill"]');
        for (var i = 0; i < nodes.length; i++) {
            if (nodes[i].children && nodes[i].children.length > 3) continue;
            var t = String(nodes[i].innerText || nodes[i].textContent || '').replace(/\s+/g, ' ').trim();
            if (!t || t.length > 90) continue;
            chips.push(t);
        }
        var b2a1 = false;
        var fertig = false;
        var label = '';
        for (var c = 0; c < chips.length; c++) {
            if (isB2a1Chip(chips[c])) {
                b2a1 = true;
                label = chips[c];
                break;
            }
        }
        if (!b2a1) {
            for (var f = 0; f < chips.length; f++) {
                if (isFertigChip(chips[f])) {
                    fertig = true;
                    label = chips[f];
                    break;
                }
            }
        }
        return { b2a1: b2a1, fertig: fertig, label: label };
    }

    function looksLikeReport(o) {
        return !!(o && typeof o === 'object' && (o.success === true || o.success === false || o.batchId || o.nextId || o.pending != null || o.accepted));
    }

    function parseLoose(txt) {
        var s = String(txt || '').trim();
        if (s.charAt(0) === '{') {
            try {
                var o = JSON.parse(s);
                if (looksLikeReport(o)) return o;
            } catch (e) {}
        }
        var pre = s.match(/<pre[^>]*>([\s\S]*?)<\/pre>/i);
        if (pre) {
            try {
                var inner = JSON.parse(pre[1].replace(/&quot;/g, '"').replace(/&amp;/g, '&').replace(/&lt;/g, '<'));
                if (looksLikeReport(inner)) return inner;
            } catch (e2) {}
        }
        return null;
    }

    function echoUrl(txt) {
        var s = String(txt || '').replace(/&amp;/g, '&');
        var m = s.match(/https:\/\/script\.googleusercontent\.com\/macros\/echo[^"'\\\s<>]*/);
        return m ? m[0] : '';
    }

    function gmReq(method, url, body, onOk, onErr, hop) {
        hop = hop || 0;
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
        function handle(res) {
            var txt = res && res.responseText != null ? res.responseText : String(res || '');
            var data = parseLoose(txt);
            if (data) {
                finishOk(data);
                return;
            }
            var jump = hop < 2 ? echoUrl(txt) : '';
            if (jump) {
                done = true;
                gmReq('GET', jump, '', onOk, onErr, hop + 1);
                return;
            }
            var st = res && res.status;
            if (st >= 200 && st < 400) {
                finishOk({ success: true, accepted: true });
                return;
            }
            finishErr('kein JSON');
        }
        if (typeof GM_xmlhttpRequest !== 'function') {
            finishErr('kein GM');
            return;
        }
        GM_xmlhttpRequest({
            method: method,
            url: url,
            data: body || undefined,
            headers: method === 'POST' ? { 'Content-Type': 'text/plain;charset=utf-8' } : {},
            anonymous: false,
            timeout: 20000,
            onload: handle,
            onerror: function () { finishErr('netzwerk'); },
            ontimeout: function () { finishErr('timeout'); }
        });
    }

    function gmPost(body, onOk, onErr) {
        gmReq('POST', WEB_APP_URL, JSON.stringify(body || {}), onOk, function () {
            var q = WEB_APP_URL + '?page=' + encodeURIComponent(body.page || '')
                + '&batch=' + encodeURIComponent(body.batch || '')
                + '&sid=' + encodeURIComponent(body.sid || '')
                + '&b2a1=' + encodeURIComponent(body.b2a1 || '0')
                + '&fertig=' + encodeURIComponent(body.fertig || '0')
                + '&label=' + encodeURIComponent(body.label || '')
                + '&detail=' + encodeURIComponent(body.detail || '0');
            gmReq('GET', q, '', onOk, function (err) {
                onOk({ success: true, accepted: true, message: err || 'accepted' });
            });
        });
    }

    function toast(msg) {
        if (!document.body) return;
        var el = document.getElementById('retoure-carol-toast');
        if (!el) {
            el = document.createElement('div');
            el.id = 'retoure-carol-toast';
            el.style.cssText = 'position:fixed;bottom:18px;left:18px;z-index:999999;padding:12px 16px;border-radius:12px;font:13px/1.4 Segoe UI,sans-serif;background:#0f1720;color:#e6edf3;border:1px solid #2dd4bf;box-shadow:0 8px 24px rgba(0,0,0,.45);max-width:380px;';
            el.innerHTML = '<div style="font-weight:700;color:#2dd4bf;margin-bottom:4px;">Retoure Scan ' + VER + '</div><div id="retoure-carol-toast-msg"></div>';
            document.body.appendChild(el);
        }
        var msgEl = document.getElementById('retoure-carol-toast-msg');
        if (msgEl) msgEl.textContent = msg;
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
    }

    function afterSend(batch, sid, res) {
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
        toast(sid + ' ok · Scan holt nächste ID');
        busy = false;
    }

    function sendFlags(batch, sid, flags) {
        rememberSid(sid);
        var kind = flags.b2a1 ? 'B2A1' : (flags.fertig ? 'Fertig' : 'kein Badge');
        toast(sid + ' · ' + kind + ' — sende…');
        gmPost({
            page: 'carolreport',
            batch: batch || '',
            sid: sid,
            b2a1: flags.b2a1 ? '1' : '0',
            fertig: flags.fertig ? '1' : '0',
            label: flags.label || '',
            detail: '1'
        }, function (res) {
            afterSend(batch, sid, res || {});
        }, function () {
            afterSend(batch, sid, { accepted: true });
        });
    }

    function waitAndReport(batch, sid) {
        rememberSid(sid);
        var tries = 0;
        toast('Öffne Auftrag ' + sid + '…');
        var timer = setInterval(function () {
            tries++;
            var liveSid = sidFromPage() || sid;
            if (liveSid) sid = rememberSid(liveSid);
            if (!isDetailPage()) {
                if (tries === 3 || tries === 10) fillSearch(sid);
                clickStockLink(sid);
                if (tries < 40) return;
                toast(sid + ' · Auftrag nicht offen — klicke das Fahrzeug');
                return;
            }
            var flags = readFlags();
            if (!flags.b2a1 && !flags.fertig && tries < 24) return;
            clearInterval(timer);
            sendFlags(batch, sid, flags);
        }, 280);
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
            if (!sid && tries < 20) {
                toast('Warte auf Auftrag…');
                setTimeout(kick, 250);
                return;
            }
            if (sid) {
                rememberSid(sid);
                waitAndReport(batch, sid);
                return;
            }
            gmPost({ page: 'carolq', batch: batch || '' }, function (res) {
                if (res && res.batchId) {
                    try { sessionStorage.setItem('retoure_batch', res.batchId); } catch (e) {}
                    batch = res.batchId;
                }
                if (res && res.nextId) {
                    busy = false;
                    goNext(batch, res.nextId, res.carolUrl || '');
                    return;
                }
                finishWalk('Carol-Check fertig');
            }, function () {
                finishWalk('Queue Ende');
            });
        }
        kick();
    }

    function boot() {
        lastHref = location.href;
        start();
        setInterval(function () {
            if (location.href === lastHref) return;
            lastHref = location.href;
            busy = false;
            start();
        }, 400);
    }

    if (document.readyState === 'loading') window.addEventListener('DOMContentLoaded', boot);
    else boot();
})();
