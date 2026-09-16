// ==UserScript==
// @name         Carol Retoure Bridge
// @namespace    retoure-lager
// @version      1.0
// @description  Liest Carol-Badges Als B2A1 markiert / Flagged for Return / Fertiggestellt / Completed und sendet sie an Retoure Scan
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
        if (fromQs) {
            try { sessionStorage.setItem('retoure_sid', fromQs); } catch (e) {}
            return fromQs;
        }
        try { return String(sessionStorage.getItem('retoure_sid') || '').replace(/\s+/g, '').toUpperCase(); } catch (e2) { return ''; }
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
            if (txt.indexOf(want) !== -1) {
                try { links[i].click(); } catch (e) {}
                return true;
            }
        }
        return false;
    }

    function pageText() {
        return String((document.body && document.body.innerText) || '').replace(/\u00a0/g, ' ');
    }

    function readFlags() {
        var t = pageText();
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

    function sidOnPage(sid) {
        var t = pageText().toUpperCase();
        return t.indexOf(String(sid || '').toUpperCase()) !== -1;
    }

    function gmGet(url, onOk, onErr) {
        if (typeof GM_xmlhttpRequest !== 'function') {
            fetch(url, { credentials: 'include' }).then(function (r) { return r.text(); }).then(function (txt) {
                var data = {};
                try { data = JSON.parse(txt); } catch (e) { onErr('kein JSON'); return; }
                onOk(data);
            }).catch(function () { onErr('netzwerk'); });
            return;
        }
        GM_xmlhttpRequest({
            method: 'GET',
            url: url,
            onload: function (res) {
                var data = {};
                try { data = JSON.parse(res.responseText); } catch (e) {
                    onErr('kein JSON');
                    return;
                }
                onOk(data);
            },
            onerror: function () { onErr('netzwerk'); }
        });
    }

    function toast(msg) {
        var el = document.getElementById('retoure-carol-toast');
        if (!el) {
            el = document.createElement('div');
            el.id = 'retoure-carol-toast';
            el.style.cssText = 'position:fixed;bottom:18px;left:18px;z-index:999999;padding:10px 14px;border-radius:10px;font:13px/1.35 Segoe UI,sans-serif;background:#0f1720;color:#e6edf3;border:1px solid #2dd4bf;box-shadow:0 8px 24px rgba(0,0,0,.4);max-width:360px;';
            document.body.appendChild(el);
        }
        el.textContent = 'Retoure Scan · ' + msg;
        el.style.display = 'block';
    }

    function queueUrl(batch) {
        return WEB_APP_URL + '?page=carolq&batch=' + encodeURIComponent(batch);
    }

    function reportUrl(batch, sid, flags) {
        return WEB_APP_URL + '?page=carolreport'
            + '&batch=' + encodeURIComponent(batch)
            + '&sid=' + encodeURIComponent(sid)
            + '&b2a1=' + (flags.b2a1 ? '1' : '0')
            + '&fertig=' + (flags.fertig ? '1' : '0')
            + '&label=' + encodeURIComponent(flags.label || '');
    }

    function langPrefix() {
        var m = location.pathname.match(/^\/([a-z]{2}-[A-Z]{2})/);
        return m ? m[1] : 'de-DE';
    }

    function goNext(batch, sid) {
        if (currentSid() === sid && batchId() === batch) return;
        location.assign(location.origin + '/' + langPrefix() + '/refurbishment?rsv=' + encodeURIComponent(sid) + '&retoure_batch=' + encodeURIComponent(batch));
    }

    function waitAndReport(batch, sid) {
        var tries = 0;
        var timer = setInterval(function () {
            tries++;
            if (!sidOnPage(sid) && tries < 20) return;
            if (!isDetailPage() && tries < 18) {
                clickStockLink(sid);
                return;
            }
            if (tries < 8 && !/VIN|Stock|Refurbish/i.test(pageText())) return;
            var flags = readFlags();
            if (!flags.b2a1 && !flags.fertig && tries < 24) return;
            clearInterval(timer);
            toast(sid + (flags.b2a1 ? ' · B2A1' : (flags.fertig ? ' · Fertig' : ' · kein Badge')));
            gmGet(reportUrl(batch, sid, flags), function (res) {
                var next = res && res.nextId;
                if (next) goNext(batch, next);
                else toast((res && res.message) || 'Carol-Check fertig');
            }, function (err) {
                toast('Report fehlgeschlagen · ' + err);
            });
        }, 400);
    }

    function start() {
        var batch = batchId();
        if (!batch || busy) return;
        busy = true;
        toast('Carol-Queue…');
        gmGet(queueUrl(batch), function (res) {
            if (!res || !res.success) {
                toast((res && res.message) || 'Queue fehlgeschlagen');
                busy = false;
                return;
            }
            if (!res.ids || !res.ids.length) {
                toast('Carol-Check fertig');
                return;
            }
            var sid = currentSid();
            if (!sid || res.ids.indexOf(sid) === -1) {
                goNext(batch, res.nextId || res.ids[0]);
                return;
            }
            waitAndReport(batch, sid);
        }, function (err) {
            toast('Queue fehlgeschlagen · ' + err);
            busy = false;
        });
    }

    if (batchId()) {
        if (document.readyState === 'complete' || document.readyState === 'interactive') start();
        else window.addEventListener('DOMContentLoaded', start);
    }
})();
