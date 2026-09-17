// ==UserScript==
// @name         Carol Retoure Bridge
// @namespace    retoure-lager
// @version      2.5
// @description  Öffnet den Carol-Auftrag, liest die Badges mit Datum und sendet sie an Retoure Scan
// @match        *://carol.autohero.com/*
// @grant        GM_xmlhttpRequest
// @grant        window.close
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';

    var VER = '2.5';
    if (window.__retoureBridgeVer) return;
    window.__retoureBridgeVer = VER;

    var WEB_APP_URL = 'https://script.google.com/a/macros/auto1.com/s/AKfycbwsGB1o_1z0t9nCVXDx0lu3nQv8Ltj81Dgq5BVw8laLHPA4v4oLUpNvj-qx49iMjeVm/exec';
    var MAX_TRIES = 40;
    var ORDER_RE = /\/refurbishment\/[0-9a-f-]{8,}/i;
    var BADGE_BOX = '[class*="refurbishmentBadges"], [class*="RefurbishmentBadges"]';
    var BADGE_QA = '[data-qa-selector*="efurbishmentStatus"]';
    var ACTION_EL = 'button, a, input, select, option, textarea, label, form, [role="button"], [role="menu"], [role="menuitem"], [role="tab"], [aria-haspopup]';
    var B2A1_RE = /als\s+b2a1\s+markiert|flagged\s+for\s+return\s+to\s+auto\s*1/i;
    var FERTIG_RE = /fertiggestellt|completed\s+on\s+\d/i;
    var STARTED_RE = /refurbishment\s+(started|gestartet)/i;

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

    function wantedSid() {
        var v = String(qs('rsv') || '').replace(/\s+/g, '').toUpperCase();
        return /^[A-Z]{2}\d{4,8}$/.test(v) ? v : '';
    }

    function isDetailPage() {
        return ORDER_RE.test(location.pathname);
    }

    function pageText() {
        return String((document.body && document.body.innerText) || '').replace(/\u00a0/g, ' ');
    }

    function detailSid() {
        var head = document.querySelector('[class*="vehicleHeader"], [class*="VehicleHeader"]');
        if (head) {
            var hm = String(head.innerText || head.textContent || '').match(/\b([A-Z]{2}\d{4,8})\b/);
            if (hm) return hm[1].toUpperCase();
        }
        var nodes = document.querySelectorAll('h1, h2, h3, h4, h5');
        for (var i = 0; i < nodes.length; i++) {
            var m = String(nodes[i].textContent || '').trim().match(/^([A-Z]{2}\d{4,8})\s*[-–]/);
            if (m) return m[1].toUpperCase();
        }
        var m2 = pageText().match(/\b([A-Z]{2}\d{4,8})\s*[-–]\s*\S/);
        return m2 ? m2[1].toUpperCase() : '';
    }

    function doneKey(batch) {
        return 'retoure_done_' + (batch || 'x');
    }

    function isDone(batch, sid) {
        if (qs('retoure_force') === '1') return false;
        return ssGet(doneKey(batch)).indexOf('|' + sid + '|') !== -1;
    }

    function markDone(batch, sid) {
        var cur = ssGet(doneKey(batch));
        if (cur.indexOf('|' + sid + '|') === -1) ssSet(doneKey(batch), cur + '|' + sid + '|');
    }

    function unmarkDone(batch, sid) {
        var cur = ssGet(doneKey(batch));
        ssSet(doneKey(batch), cur.split('|' + sid + '|').join('|'));
    }

    function shortLabel(s) {
        s = String(s || '').replace(/[·•]/g, '|').replace(/\s+/g, ' ').trim();
        if (s.length > 80) s = s.slice(0, 80);
        return s;
    }

    function pickNext(res, skipSid, batch) {
        skipSid = String(skipSid || '').toUpperCase();
        var next = String((res && res.nextId) || '').toUpperCase();
        if (next && next !== skipSid && !isDone(batch, next)) return next;
        var ids = (res && res.ids) || [];
        var i;
        for (i = 0; i < ids.length; i++) {
            var id = String(ids[i] || '').toUpperCase();
            if (id && id !== skipSid && !isDone(batch, id)) return id;
        }
        return '';
    }

    function isAction(el) {
        try { return !!el.closest(ACTION_EL); } catch (e) { return false; }
    }

    function textsIn(root, skipActions) {
        var out = [];
        var seen = {};
        function push(raw) {
            var t = String(raw || '').replace(/\s+/g, ' ').trim();
            if (!t || t.length > 90 || seen[t]) return;
            seen[t] = 1;
            out.push(t);
        }
        if (!root) return out;
        var leaves = root.querySelectorAll('span, div, p, small, strong, b');
        for (var i = 0; i < leaves.length; i++) {
            if (leaves[i].children && leaves[i].children.length) continue;
            if (skipActions && isAction(leaves[i])) continue;
            push(leaves[i].textContent);
        }
        if (!out.length) push(root.innerText || root.textContent);
        return out;
    }

    function badgeTextsDetail() {
        var out = [];
        var seen = {};
        function add(list) {
            for (var i = 0; i < list.length; i++) {
                if (seen[list[i]]) continue;
                seen[list[i]] = 1;
                out.push(list[i]);
            }
        }
        var boxes = document.querySelectorAll(BADGE_BOX);
        for (var b = 0; b < boxes.length; b++) add(textsIn(boxes[b], true));
        var qa = document.querySelectorAll(BADGE_QA);
        for (var q = 0; q < qa.length; q++) {
            if (isAction(qa[q])) continue;
            add(textsIn(qa[q], true));
        }
        return out;
    }

    function flagsFrom(texts) {
        for (var i = 0; i < texts.length; i++) {
            if (B2A1_RE.test(texts[i])) return { b2a1: true, fertig: false, label: texts[i] };
        }
        for (var f = 0; f < texts.length; f++) {
            if (FERTIG_RE.test(texts[f]) && !STARTED_RE.test(texts[f])) {
                return { b2a1: false, fertig: true, label: texts[f] };
            }
        }
        return { b2a1: false, fertig: false, label: shortLabel(texts.join(' | ')) };
    }

    function rowLooksComplete(el, raw) {
        var a = el.querySelector('a[href*="refurbishment/"]');
        if (a && ORDER_RE.test(a.getAttribute('href') || a.href || '')) return true;
        return /als\s+b2a1|flagged\s+for\s+return|fertiggestellt|completed\s+on|refurbishment\s+started|automatisch\s+beauftragt|auto-?ordered/i.test(raw);
    }

    function rowForSid(sid) {
        var want = String(sid || '').toUpperCase();
        if (!want) return null;
        var cands = document.querySelectorAll('article, li, tr, [class*="card"], [class*="Card"], [class*="row"], [class*="result"], [class*="Result"], [class*="entry"], [class*="item"]');
        var best = null;
        var bestLen = 1e9;
        var loose = null;
        var looseLen = 1e9;
        for (var i = 0; i < cands.length; i++) {
            var el = cands[i];
            var raw = String(el.textContent || '');
            if (raw.length > 2000) continue;
            if (raw.toUpperCase().replace(/\s+/g, '').indexOf(want) === -1) continue;
            if (rowLooksComplete(el, raw)) {
                if (raw.length < bestLen) {
                    best = el;
                    bestLen = raw.length;
                }
            } else if (raw.length < looseLen) {
                loose = el;
                looseLen = raw.length;
            }
        }
        if (best) return best;
        var up = loose;
        for (var h = 0; up && h < 6; h++) {
            var txt = String(up.textContent || '');
            if (txt.length < 2000 && rowLooksComplete(up, txt)) return up;
            up = up.parentElement;
        }
        return loose;
    }

    function orderHref(root) {
        var scopes = [root, document];
        for (var s = 0; s < scopes.length; s++) {
            if (!scopes[s]) continue;
            var links = scopes[s].querySelectorAll('a[href*="refurbishment/"]');
            for (var i = 0; i < links.length; i++) {
                var href = links[i].getAttribute('href') || links[i].href || '';
                if (ORDER_RE.test(href)) return href;
            }
        }
        return '';
    }

    function fireClick(el) {
        if (!el) return false;
        try { el.scrollIntoView({ block: 'center' }); } catch (e) {}
        var evts = ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click'];
        for (var i = 0; i < evts.length; i++) {
            try {
                el.dispatchEvent(new MouseEvent(evts[i], { bubbles: true, cancelable: true, view: window }));
            } catch (e2) {}
        }
        try { el.click(); } catch (e3) {}
        return true;
    }

    function openOrder(row, batch) {
        var href = orderHref(row);
        if (href) {
            var abs = href.indexOf('http') === 0 ? href : (location.origin + (href.charAt(0) === '/' ? '' : '/') + href);
            if (batch && abs.indexOf('retoure_batch=') === -1) {
                abs += (abs.indexOf('?') >= 0 ? '&' : '?') + 'retoure_batch=' + encodeURIComponent(batch);
            }
            location.assign(abs);
            return true;
        }
        if (!row) return false;
        var title = row.querySelector('[class*="title"], [class*="Title"], [class*="link"], [class*="Link"], h1, h2, h3, h4, h5, strong, b');
        if (title && fireClick(title)) return true;
        return fireClick(row);
    }

    function noResults() {
        var t = pageText();
        if (/\b0\s+(ergebnisse|results)\b/i.test(t)) return true;
        return /keine\s+(ergebnisse|fahrzeuge)|no\s+(results|vehicles)\s+found/i.test(t);
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
            anonymous: true,
            timeout: 12000,
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
                if (res && res.status >= 200 && res.status < 400) ok({ accepted: true });
                else err('HTTP ' + (res && res.status));
            },
            onerror: function () { err('netzwerk'); },
            ontimeout: function () { err('timeout'); }
        });
    }

    function post(body, onOk, onErr) {
        var payload = {
            page: body.page || '',
            batch: body.batch || '',
            sid: body.sid || '',
            b2a1: body.b2a1 || '0',
            fertig: body.fertig || '0',
            notfound: body.notfound || '0',
            label: shortLabel(body.label || ''),
            detail: body.detail || '0'
        };
        function tryPost(n) {
            gmReq('POST', WEB_APP_URL, JSON.stringify(payload), onOk, function (e) {
                if (n < 2) {
                    setTimeout(function () { tryPost(n + 1); }, 700 * (n + 1));
                    return;
                }
                var q = WEB_APP_URL + '?page=' + encodeURIComponent(payload.page)
                    + '&batch=' + encodeURIComponent(payload.batch)
                    + '&sid=' + encodeURIComponent(payload.sid)
                    + '&b2a1=' + encodeURIComponent(payload.b2a1)
                    + '&fertig=' + encodeURIComponent(payload.fertig)
                    + '&notfound=' + encodeURIComponent(payload.notfound)
                    + '&detail=' + encodeURIComponent(payload.detail)
                    + '&label=' + encodeURIComponent(payload.label);
                gmReq('GET', q, '', onOk, onErr || function () {});
            });
        }
        tryPost(0);
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
        if (url && ORDER_RE.test(url)) {
            var u = url;
            if (batch && u.indexOf('retoure_batch=') === -1) u += (u.indexOf('?') >= 0 ? '&' : '?') + 'retoure_batch=' + encodeURIComponent(batch);
            location.assign(u);
            return;
        }
        if (!sid) return;
        location.assign(location.origin + '/' + langPrefix() + '/refurbishment?rsv=' + encodeURIComponent(sid)
            + (batch ? '&retoure_batch=' + encodeURIComponent(batch) : ''));
    }

    function pingHud(data) {
        var targets = [];
        try { if (window.opener) targets.push(window.opener); } catch (e0) {}
        try { if (window.opener && window.opener.parent) targets.push(window.opener.parent); } catch (e1) {}
        try { if (window.opener && window.opener.top) targets.push(window.opener.top); } catch (e2) {}
        for (var i = 0; i < targets.length; i++) {
            try { targets[i].postMessage(data, '*'); } catch (e3) {}
        }
    }

    function isQueueDone(res) {
        return !!(res && res.success === true && Number(res.pending) === 0 && !res.accepted);
    }

    function goHome() {
        pingHud({ retoureCarolDone: 1 });
        try {
            if (window.opener && !window.opener.closed) {
                try { window.opener.focus(); } catch (e1) {}
            }
        } catch (e2) {}
        setTimeout(function () {
            try { window.close(); } catch (e3) {}
        }, 350);
    }

    function finish(msg) {
        toast(msg || 'Carol-Check fertig');
        busy = false;
        goHome();
    }

    function askQueue(batch, sid) {
        post({ page: 'carolq', batch: batch || '' }, function (res) {
            if (res && res.batchId) ssSet('retoure_batch', res.batchId);
            var b = res.batchId || batch;
            var next = pickNext(res, sid, b);
            if (next) {
                toast('weiter · ' + next);
                busy = false;
                goTo(b, next, res.carolUrl || '');
                return;
            }
            if (isQueueDone(res)) {
                finish('Carol-Check fertig');
                return;
            }
            if (res && res.accepted) {
                toast('Queue unklar — erneut');
                busy = false;
                setTimeout(function () { if (!busy) start(); }, 900);
                return;
            }
            if (String(res && res.nextId || '').toUpperCase() === String(sid || '').toUpperCase()) {
                unmarkDone(b, sid);
                toast(sid + ' · Report erneut');
                busy = false;
                setTimeout(function () { if (!busy) start(); }, 700);
                return;
            }
            if (res && Number(res.pending) > 0 && res.nextId) {
                unmarkDone(b, String(res.nextId).toUpperCase());
                toast('weiter · ' + res.nextId);
                busy = false;
                goTo(b, res.nextId, res.carolUrl || '');
                return;
            }
            toast('Warte auf Scan-Seite');
            busy = false;
        }, function (e) {
            toast('Queue: ' + e + ' — erneut');
            busy = false;
            setTimeout(function () { if (!busy) start(); }, 1200);
        });
    }

    function send(batch, sid, flags, notfound, attempt) {
        attempt = attempt || 0;
        toast(sid + ' · ' + (notfound ? 'kein Auftrag in Carol' : (flags.label || 'kein Badge')) + (attempt ? ' — erneut…' : ' — sende…'));
        post({
            page: 'carolreport',
            batch: batch || '',
            sid: sid,
            b2a1: flags.b2a1 ? '1' : '0',
            fertig: flags.fertig ? '1' : '0',
            notfound: notfound ? '1' : '0',
            label: shortLabel(flags.label || ''),
            detail: '1'
        }, function (res) {
            if (res && res.accepted && res.success !== true) {
                if (attempt < 4) {
                    toast(sid + ' · Report unklar — erneut');
                    setTimeout(function () { send(batch, sid, flags, notfound, attempt + 1); }, 800 * (attempt + 1));
                    return;
                }
                toast(sid + ' · Report unklar — Scan-Seite übernimmt');
                busy = false;
                return;
            }
            if (res && res.success === false && String(res.message || '') === 'Kein Auftrag offen') {
                toast(sid + ' · Scan-Seite übernimmt');
                busy = false;
                return;
            }
            if (!res || res.success === false) {
                if (attempt < 4) {
                    toast(sid + ' · ' + ((res && res.message) || 'Fehler') + ' — erneut');
                    setTimeout(function () { send(batch, sid, flags, notfound, attempt + 1); }, 800 * (attempt + 1));
                    return;
                }
                toast(sid + ' · Report hängt — Scan-Seite übernimmt');
                busy = false;
                return;
            }
            markDone(batch, sid);
            pingHud({
                retoureCarolUpdate: 1,
                stockId: sid,
                carolB2a1: !!flags.b2a1,
                carolFertig: !!flags.fertig,
                carolLabel: notfound ? 'Kein Carol-Auftrag' : shortLabel(flags.label || ''),
                notFound: !!notfound,
                pending: res && res.pending,
                nextId: res && res.nextId
            });
            if (res && res.batchId) ssSet('retoure_batch', res.batchId);
            var b = (res && res.batchId) || batch;
            var next = pickNext(res, sid, b);
            if (next) {
                toast(sid + ' ok · weiter ' + next);
                busy = false;
                goTo(b, next, res.carolUrl || '');
                return;
            }
            if (isQueueDone(res)) {
                finish(sid + ' ok · 0 offen');
                return;
            }
            toast(sid + ' ok · ' + ((res && res.pending) || '?') + ' offen');
            busy = false;
            askQueue(batch, sid);
        }, function (e) {
            if (attempt < 4) {
                toast(sid + ' · Report ' + e + ' — erneut');
                setTimeout(function () { send(batch, sid, flags, notfound, attempt + 1); }, 800 * (attempt + 1));
                return;
            }
            toast(sid + ' · Report ' + e + ' — Scan-Seite übernimmt');
            busy = false;
        });
    }

    function run() {
        var batch = batchId();
        var want = wantedSid();
        var tries = 0;
        var opened = false;
        if (timer) clearInterval(timer);
        timer = setInterval(function () {
            tries++;

            if (isDetailPage()) {
                var sid = detailSid() || want;
                if (!sid) {
                    if (tries < MAX_TRIES) return;
                    clearInterval(timer);
                    busy = false;
                    toast('Stock-ID nicht gefunden');
                    return;
                }
                if (isDone(batch, sid)) {
                    clearInterval(timer);
                    askQueue(batch, sid);
                    return;
                }
                var texts = badgeTextsDetail();
                if (!texts.length) {
                    if (tries < MAX_TRIES) return;
                    clearInterval(timer);
                    send(batch, sid, { b2a1: false, fertig: false, label: 'Badges nicht geladen' }, false);
                    return;
                }
                clearInterval(timer);
                send(batch, sid, flagsFrom(texts), false);
                return;
            }

            if (!want) {
                if (tries < MAX_TRIES) return;
                clearInterval(timer);
                busy = false;
                toast('Keine Stock-ID in der URL');
                return;
            }
            if (isDone(batch, want)) {
                clearInterval(timer);
                askQueue(batch, want);
                return;
            }
            if (tries === 2 || tries === 16) fillSearch(want);

            var row = rowForSid(want);
            if (row) {
                if (!opened) {
                    opened = true;
                    toast('Öffne Auftrag ' + want + '…');
                    if (openOrder(row, batch)) return;
                }
                if (tries % 12 === 0) opened = false;
                if (tries < MAX_TRIES) return;
                clearInterval(timer);
                send(batch, want, flagsFrom(textsIn(row, true)), false);
                return;
            }

            if (noResults() && tries > 6) {
                clearInterval(timer);
                send(batch, want, { b2a1: false, fertig: false, label: '' }, true);
                return;
            }
            if (tries < MAX_TRIES) return;
            clearInterval(timer);
            send(batch, want, { b2a1: false, fertig: false, label: '' }, true);
        }, 300);
    }

    function start() {
        if (!/\/refurbishment/i.test(location.pathname + location.search)) return;
        if (!isDetailPage() && !wantedSid()) return;
        if (busy) return;
        busy = true;
        toast(isDetailPage() ? 'Lese Auftrag…' : ('Suche ' + wantedSid() + '…'));
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
