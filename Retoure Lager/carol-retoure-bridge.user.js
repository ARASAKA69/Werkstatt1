// ==UserScript==
// @name         Carol Retoure Bridge
// @namespace    retoure-lager
// @version      1.2
// @description  Prüft automatisch alle Dump-IDs in Carol (B2A1 / Fertiggestellt) und sendet sie an Retoure Scan
// @match        *://carol.autohero.com/*
// @grant        GM_xmlhttpRequest
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';

    var WEB_APP_URL = 'https://script.google.com/a/macros/auto1.com/s/AKfycbwsGB1o_1z0t9nCVXDx0lu3nQv8Ltj81Dgq5BVw8laLHPA4v4oLUpNvj-qx49iMjeVm/exec';
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

    function currentSid() {
        var fromQs = String(qs('rsv') || '').replace(/\s+/g, '').toUpperCase();
        if (/^[A-Z]{2}\d{4,8}$/.test(fromQs)) {
            try { sessionStorage.setItem('retoure_sid', fromQs); } catch (e) {}
            return fromQs;
        }
        try { return String(sessionStorage.getItem('retoure_sid') || '').replace(/\s+/g, '').toUpperCase(); } catch (e2) { return ''; }
    }

    function sidFromPage() {
        var fromQs = currentSid();
        if (/^[A-Z]{2}\d{4,8}$/.test(fromQs)) return fromQs;
        var headings = document.querySelectorAll('h1, h2, h3, [class*="title"], [class*="Title"], [class*="heading"]');
        for (var i = 0; i < headings.length; i++) {
            var m = String(headings[i].textContent || '').trim().match(/^([A-Z]{2}\d{4,8})\b/i);
            if (m) return m[1].toUpperCase();
        }
        var t = pageText();
        var m2 = t.match(/\b([A-Z]{2}\d{4,8})\b/);
        return m2 ? m2[1].toUpperCase() : '';
    }

    function isDetailPage() {
        return /\/refurbishment\/[0-9a-f-]{20,}/i.test(location.pathname);
    }

    function clickStockLink(sid) {
        var want = String(sid || '').toUpperCase();
        var links = document.querySelectorAll('a[href*="/refurbishment/"]');
        for (var i = 0; i < links.length; i++) {
            var href = links[i].getAttribute('href') || '';
            if (!/\/refurbishment\/[0-9a-f-]{20,}/i.test(href)) continue;
            var row = links[i].closest('tr, [role="row"], .rt-tr-group, [class*="row"]') || links[i].parentElement;
            var txt = ((row && row.textContent) || links[i].textContent || '').toUpperCase().replace(/\s+/g, '');
            if (!want || txt.indexOf(want) !== -1) {
                try { links[i].click(); } catch (e) {}
                return true;
            }
        }
        return false;
    }

    function pageText() {
        return String((document.body && document.body.innerText) || '').replace(/\u00a0/g, ' ');
    }

    function badgeText() {
        var t = pageText();
        var nodes = document.querySelectorAll('[class*="badge"], [class*="Badge"], [class*="chip"], [class*="Chip"], [class*="tag"], [class*="Tag"], [class*="pill"], [class*="status"], [class*="Status"]');
        for (var i = 0; i < nodes.length; i++) {
            t += '\n' + String(nodes[i].innerText || nodes[i].textContent || '');
        }
        return t;
    }

    function readFlags() {
        var t = badgeText();
        var b2a1 = /als\s*b2a1\s*markiert|flagged\s*for\s*return(\s*to\s*auto\s*1)?/i.test(t);
        var fertig = /fertiggestellt(\s+am)?|completed\s+on\s+\d/i.test(t);
        var label = '';
        var m = t.match(/als\s*b2a1\s*markiert[^\n]{0,48}/i)
            || t.match(/flagged\s*for\s*return[^\n]{0,56}/i)
            || t.match(/fertiggestellt[^\n]{0,48}/i)
            || t.match(/completed\s+on[^\n]{0,48}/i);
        if (m) label = String(m[0] || '').replace(/\s+/g, ' ').trim();
        if (b2a1) fertig = false;
        return { b2a1: b2a1, fertig: fertig, label: label };
    }

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

    function gmPost(body, onOk, onErr) {
        var payload = JSON.stringify(body || {});
        if (typeof GM_xmlhttpRequest === 'function') {
            GM_xmlhttpRequest({
                method: 'POST',
                url: WEB_APP_URL,
                data: payload,
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                anonymous: false,
                onload: function (res) {
                    var data = parseLoose(res.responseText);
                    if (!data) {
                        onErr('kein JSON');
                        return;
                    }
                    onOk(data);
                },
                onerror: function () { onErr('netzwerk'); }
            });
            return;
        }
        fetch(WEB_APP_URL, {
            method: 'POST',
            mode: 'cors',
            credentials: 'include',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: payload
        }).then(function (r) { return r.text(); }).then(function (txt) {
            var data = parseLoose(txt);
            if (!data) { onErr('kein JSON'); return; }
            onOk(data);
        }).catch(function () { onErr('netzwerk'); });
    }

    function toast(msg) {
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

    function goNext(batch, sid) {
        if (!batch || !sid) return;
        location.assign(location.origin + '/' + langPrefix() + '/refurbishment?rsv=' + encodeURIComponent(sid) + '&retoure_batch=' + encodeURIComponent(batch));
    }

    function finishWalk(msg) {
        toast(msg || 'Carol-Check fertig');
        busy = false;
        if (walkQueue) {
            setTimeout(function () {
                try { window.close(); } catch (e) {}
            }, 500);
        }
    }

    function sendFlags(batch, sid, flags) {
        toast(sid + (flags.b2a1 ? ' · Als B2A1 markiert' : (flags.fertig ? ' · Fertiggestellt' : ' · kein Badge')) + ' — weiter…');
        gmPost({
            page: 'carolreport',
            batch: batch || '',
            sid: sid,
            b2a1: flags.b2a1 ? '1' : '0',
            fertig: flags.fertig ? '1' : '0',
            label: flags.label || ''
        }, function (res) {
            if (res && res.batchId) {
                try { sessionStorage.setItem('retoure_batch', res.batchId); } catch (e) {}
            }
            var next = walkQueue && res && res.nextId;
            if (next) {
                toast(sid + ' ok · weiter ' + next);
                busy = false;
                goNext(res.batchId || batch, next);
                return;
            }
            finishWalk((res && res.message) || 'Carol-Check fertig');
        }, function (err) {
            toast('Report fehlgeschlagen · ' + err);
            busy = false;
            if (!walkQueue) return;
            gmPost({ page: 'carolq', batch: batch || '' }, function (q) {
                if (q && q.nextId && q.nextId !== sid) goNext(q.batchId || batch, q.nextId);
                else finishWalk('Carol-Check fertig');
            }, function () { finishWalk('Carol-Check fertig'); });
        });
    }

    function waitAndReport(batch, sid) {
        var tries = 0;
        var timer = setInterval(function () {
            tries++;
            if (!isDetailPage()) {
                clickStockLink(sid);
                if (tries < 22) return;
                clearInterval(timer);
                sendFlags(batch, sid, { b2a1: false, fertig: false, label: '' });
                return;
            }
            var flags = readFlags();
            if (!flags.b2a1 && !flags.fertig && tries < 20) return;
            clearInterval(timer);
            sendFlags(batch, sid, flags);
        }, 280);
    }

    function start() {
        if (qs('retoure_auto') === '1' && !qs('retoure_batch') && !qs('rsv')) return;
        if (!/\/refurbishment/i.test(location.pathname + location.search)) return;
        var sid = sidFromPage();
        var batch = batchId();
        walkQueue = !!qs('retoure_batch') || !!batch;
        if (!sid && !batch) return;
        if (busy) return;
        busy = true;
        toast(sid ? ('Lese ' + sid + '…') : 'Carol-Queue…');
        gmPost({ page: 'carolq', batch: batch || '' }, function (res) {
            if (res && res.batchId) {
                try { sessionStorage.setItem('retoure_batch', res.batchId); } catch (e) {}
                batch = res.batchId;
            }
            walkQueue = walkQueue || !!(res && res.pending);
            if (sid) {
                waitAndReport(batch, sid);
                return;
            }
            if (!res || !res.success) {
                finishWalk((res && res.message) || 'Queue fehlgeschlagen');
                return;
            }
            if (!res.ids || !res.ids.length) {
                finishWalk('Carol-Check fertig');
                return;
            }
            busy = false;
            goNext(batch, res.nextId || res.ids[0]);
        }, function (err) {
            if (sid) {
                toast('Queue ' + err + ' — lese trotzdem ' + sid);
                waitAndReport(batch, sid);
                return;
            }
            finishWalk('Queue fehlgeschlagen · ' + err);
        });
    }

    function boot() {
        lastHref = location.href;
        start();
        setInterval(function () {
            if (location.href === lastHref) return;
            lastHref = location.href;
            busy = false;
            start();
        }, 500);
    }

    if (document.readyState === 'complete' || document.readyState === 'interactive') boot();
    else window.addEventListener('DOMContentLoaded', boot);
})();
