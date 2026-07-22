// ==UserScript==
// @name         N4Parts StockID Warenkorb Suche
// @namespace    http://tampermonkey.net/
// @version      3.2
// @description  StockID-Suche + Bestellung auslösen inkl. Packzettel-Scan (N4P)
// @author       ARASAKA
// @match        https://www.n4parts.net/*
// @match        https://n4parts.net/*
// @require      https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    const BOT_VERSION = '3.2';
    const PANEL_ID = 'n4-stockid-search-panel';
    const LAST_KEY = 'n4_stockid_last_query';
    const POS_KEY = 'n4_stockid_panel_pos';
    const FLOW_KEY = 'n4_stockid_order_flow';
    const PAGE_SIZE = 20;
    const GREEN = '#56d364';
    const ORANGE = '#f59e0b';

    let searching = false;
    let lastHits = [];
    let dragState = null;
    let flowBusy = false;
    let flowAbort = false;
    let pdfHookInstalled = false;

    function sleep(ms) {
        return new Promise((r) => setTimeout(r, ms));
    }

    function isWarenkoerbePage() {
        const h = (location.hash || '').split('?')[0];
        return h === '#/cart' || h === '#/cart/';
    }

    function defaultFlow() {
        return { stage: 'idle', step: '', bestellnummer: '', error: '', startedAt: 0 };
    }

    function getFlow() {
        try {
            const raw = sessionStorage.getItem(FLOW_KEY);
            if (!raw) return defaultFlow();
            return Object.assign(defaultFlow(), JSON.parse(raw));
        } catch (e) {
            return defaultFlow();
        }
    }

    function setFlow(patch) {
        const next = Object.assign(getFlow(), patch || {});
        sessionStorage.setItem(FLOW_KEY, JSON.stringify(next));
        renderOrderSection();
        return next;
    }

    function isFlowActive() {
        const f = getFlow();
        return f.stage === 'confirm' || f.stage === 'running' || (f.stage === 'done' && f.bestellnummer);
    }

    function shouldShowPanel() {
        return isWarenkoerbePage() || isFlowActive() || !!getFlow().bestellnummer;
    }

    function api(path, options) {
        const opts = options || {};
        const headers = {
            Accept: 'application/json',
            ...(opts.headers || {})
        };
        if (opts.body != null && !headers['Content-Type'] && !headers['content-type']) {
            headers['Content-Type'] = 'application/json';
        }
        return fetch(path, {
            credentials: 'include',
            ...opts,
            headers
        }).then(async (r) => {
            if (!r.ok) {
                const text = await r.text().catch(() => '');
                throw new Error(r.status + (text ? ': ' + text.slice(0, 120) : ''));
            }
            if (r.status === 204) return null;
            const ct = r.headers.get('content-type') || '';
            if (ct.includes('application/json')) return r.json();
            return r.text();
        });
    }

    function formatDate(iso) {
        if (!iso) return '';
        const d = new Date(iso);
        if (Number.isNaN(d.getTime())) return '';
        const dd = String(d.getDate()).padStart(2, '0');
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const yyyy = d.getFullYear();
        return dd + '.' + mm + '.' + yyyy;
    }

    async function findStock(query, onProgress) {
        const q = String(query || '').trim().toUpperCase();
        if (!q) return [];

        let page = 0;
        let total = Infinity;
        const hits = [];

        while (page * PAGE_SIZE < total) {
            if (onProgress) onProgress(page + 1, Math.ceil(total / PAGE_SIZE) || '?');
            const res = await api('/api/cart/list', {
                method: 'POST',
                body: JSON.stringify({ page, pageSize: PAGE_SIZE })
            });
            total = res.totalElements || 0;
            const list = res.cartLiteDTOS || [];
            for (const cart of list) {
                const name = String(cart.name || '').toUpperCase();
                if (name.includes(q)) {
                    hits.push({
                        page: page + 1,
                        id: cart.id,
                        name: String(cart.name || '').trim(),
                        created: cart.created,
                        count: cart.count || 0
                    });
                }
            }
            page += 1;
            if (!list.length) break;
        }
        return hits;
    }

    async function selectCart(id) {
        try {
            await api('/api/cart/select', {
                method: 'PUT',
                body: String(id)
            });
            return;
        } catch (err) {
            if (!String(err.message || '').startsWith('400')) throw err;
        }
        await api('/api/cart/select', {
            method: 'PUT',
            body: JSON.stringify({ cartId: id })
        });
    }

    function normalizeText(s) {
        return String(s || '').replace(/\s+/g, ' ').trim();
    }

    function isVisible(el) {
        if (!el || !el.getBoundingClientRect) return false;
        const r = el.getBoundingClientRect();
        const style = window.getComputedStyle(el);
        if (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0') return false;
        return r.width > 0 && r.height > 0;
    }

    function findControlByText(label) {
        const want = normalizeText(label).toLowerCase();
        const nodes = document.querySelectorAll('button, a, [role="button"], .btn, .order-btn, input[type="button"], input[type="submit"]');
        let best = null;
        let bestScore = Infinity;
        for (const el of nodes) {
            if (!isVisible(el)) continue;
            const txt = normalizeText(el.value || el.textContent || '').toLowerCase();
            if (!txt) continue;
            if (txt === want) return el;
            if (txt.includes(want) && txt.length < bestScore) {
                best = el;
                bestScore = txt.length;
            }
        }
        return best;
    }

    function panelTone() {
        const f = getFlow();
        if (f.stage === 'running') return 'active';
        if (f.stage === 'done' && f.bestellnummer) return 'done';
        if (f.stage === 'confirm') return 'warn';
        if (f.error || (f.stage === 'done' && !f.bestellnummer && f.error)) return 'error';
        return 'idle';
    }

    function syncPanelTone() {
        const panel = document.getElementById(PANEL_ID);
        if (!panel) return;
        panel.setAttribute('data-tone', panelTone());
        const badge = panel.querySelector('[data-tone-badge]');
        const map = { idle: 'BEREIT', active: 'LÄUFT', warn: 'CHECK', done: 'FERTIG', error: 'FEHLER' };
        if (badge) badge.textContent = map[panelTone()] || 'BEREIT';
    }

    function ensureStyles() {
        const styleId = PANEL_ID + '-style-v32';
        ['-style', '-style-v3', '-style-v31'].forEach((s) => {
            const old = document.getElementById(PANEL_ID + s);
            if (old) old.remove();
        });
        if (document.getElementById(styleId)) return;
        const style = document.createElement('style');
        style.id = styleId;
        style.textContent = `
@keyframes n4-click-blink {
  0%, 100% { outline-color: ${ORANGE}; box-shadow: 0 0 0 2px rgba(245,158,11,.8); }
  50% { outline-color: ${GREEN}; box-shadow: 0 0 0 3px rgba(86,211,100,.55); }
}
.n4-click-blink {
  outline: 2px solid ${ORANGE} !important;
  outline-offset: 1px !important;
  animation: n4-click-blink .4s ease-in-out 4 !important;
  position: relative !important;
  z-index: 2147483000 !important;
}
#${PANEL_ID}{
  position:fixed;top:72px;left:12px;z-index:2147483646;
  width:min(300px, calc(100vw - 24px));
  color:#c9d1d9;font-family:"Segoe UI",Arial,sans-serif;font-size:12px;
  border-radius:14px;overflow:hidden;
  background:
    radial-gradient(circle at top left, rgba(86,211,100,.12), transparent 36%),
    radial-gradient(circle at top right, rgba(245,158,11,.1), transparent 34%),
    linear-gradient(180deg, rgba(21,27,35,.98) 0%, rgba(14,19,26,.98) 100%);
  border:1px solid #2d3642;
  box-shadow:0 10px 28px rgba(0,0,0,.4);
  backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px);
  user-select:none;
}
#${PANEL_ID}[data-tone="active"]{border-color:rgba(245,158,11,.5)}
#${PANEL_ID}[data-tone="done"]{border-color:rgba(86,211,100,.5)}
#${PANEL_ID}[data-tone="warn"]{border-color:rgba(245,158,11,.62)}
#${PANEL_ID}[data-tone="error"]{border-color:rgba(248,81,73,.62)}
#${PANEL_ID}.n4dragging{opacity:.94;box-shadow:0 14px 32px rgba(0,0,0,.5)}
#${PANEL_ID} .n4h{
  display:flex;align-items:center;justify-content:space-between;gap:10px;
  padding:10px 12px;
  background:linear-gradient(180deg, rgba(22,27,34,.99) 0%, rgba(18,23,30,.99) 100%);
  border-bottom:1px solid #242d39;cursor:move;touch-action:none;
}
#${PANEL_ID} .n4h-eyebrow{
  color:#8b949e;font-size:9px;font-weight:800;letter-spacing:1.1px;
  text-transform:uppercase;margin-bottom:1px;
}
#${PANEL_ID} .n4h-title{
  color:#f0f6fc;font-size:13px;font-weight:900;letter-spacing:.3px;
  line-height:1.15;text-transform:uppercase;
}
#${PANEL_ID} .n4h-actions{display:flex;align-items:center;gap:6px;flex-shrink:0}
#${PANEL_ID} .n4h-close{
  border:1px solid rgba(255,255,255,.12);border-radius:10px;
  background:rgba(13,17,23,.72);color:#c9d1d9;
  min-width:30px;height:28px;padding:0 8px;
  font-size:11px;font-weight:900;cursor:pointer;
}
#${PANEL_ID} .n4b{padding:12px}
#${PANEL_ID} .n4b,#${PANEL_ID} input,#${PANEL_ID} .n4hit,#${PANEL_ID} .n4bnr strong{user-select:text}
#${PANEL_ID} .n4badge{
  display:inline-flex;align-items:center;justify-content:center;
  min-height:22px;padding:3px 9px;margin-bottom:8px;
  border-radius:999px;background:rgba(13,17,23,.7);
  border:1px solid rgba(245,158,11,.38);color:${ORANGE};
  font-size:9px;font-weight:900;letter-spacing:1px;text-transform:uppercase;
}
#${PANEL_ID}[data-tone="done"] .n4badge{border-color:rgba(86,211,100,.42);color:${GREEN}}
#${PANEL_ID}[data-tone="active"] .n4badge{border-color:rgba(245,158,11,.5);color:${ORANGE}}
#${PANEL_ID}[data-tone="error"] .n4badge{border-color:rgba(248,81,73,.5);color:#f85149}
#${PANEL_ID} .n4row{display:flex;gap:6px}
#${PANEL_ID} input[type="text"]{
  flex:1;min-width:0;border:1px solid #2d3642;border-radius:10px;
  padding:7px 10px;font:inherit;font-size:12px;font-weight:700;
  outline:none;background:rgba(13,17,23,.82);color:#f0f6fc;
}
#${PANEL_ID} input[type="text"]::placeholder{color:#8b949e;font-weight:600}
#${PANEL_ID} input[type="text"]:focus{border-color:rgba(245,158,11,.65);box-shadow:0 0 0 2px rgba(245,158,11,.12)}
#${PANEL_ID} .n4go{
  border:1px solid rgba(245,158,11,.55);border-radius:10px;
  background:rgba(245,158,11,.14);color:${ORANGE};
  padding:0 10px;cursor:pointer;font:inherit;font-size:11px;font-weight:900;
  white-space:nowrap;min-height:32px;
}
#${PANEL_ID} .n4go:hover{background:rgba(245,158,11,.22)}
#${PANEL_ID} .n4go:disabled{opacity:.55;cursor:default}
#${PANEL_ID} .n4status{
  margin-top:7px;min-height:14px;color:#8b949e;
  font-size:11px;font-weight:700;line-height:1.35;
}
#${PANEL_ID} .n4status.ok{color:${GREEN}}
#${PANEL_ID} .n4status.err{color:#f85149}
#${PANEL_ID} .n4hits{
  margin-top:8px;max-height:120px;overflow:auto;
  border-radius:10px;border:1px solid rgba(255,255,255,.07);
  background:rgba(13,17,23,.45);
}
#${PANEL_ID} .n4hit{
  display:block;width:100%;text-align:left;border:0;
  border-bottom:1px solid rgba(255,255,255,.06);
  background:transparent;padding:7px 9px;cursor:pointer;font:inherit;color:inherit;
}
#${PANEL_ID} .n4hit:last-child{border-bottom:0}
#${PANEL_ID} .n4hit:hover{background:rgba(86,211,100,.08)}
#${PANEL_ID} .n4hit strong{display:block;color:#f0f6fc;font-size:12px;font-weight:800}
#${PANEL_ID} .n4hit span{color:#8b949e;font-size:10px;font-weight:700}
#${PANEL_ID} .n4order{
  margin-top:10px;padding-top:10px;border-top:1px solid rgba(255,255,255,.07);
}
#${PANEL_ID} .n4order-title{
  color:#8b949e;font-size:9px;font-weight:800;letter-spacing:1px;
  text-transform:uppercase;margin-bottom:6px;
}
#${PANEL_ID} .n4obtn{
  display:block;width:100%;border-radius:10px;padding:8px 10px;
  margin-top:6px;cursor:pointer;font:inherit;font-size:12px;font-weight:900;
  text-align:center;border:1px solid rgba(255,255,255,.14);
  background:rgba(13,17,23,.82);color:#d7dee8;
}
#${PANEL_ID} .n4obtn-primary{
  border-color:rgba(245,158,11,.55);background:rgba(245,158,11,.12);color:${ORANGE};
}
#${PANEL_ID} .n4obtn-yes{
  border-color:rgba(86,211,100,.55);background:rgba(86,211,100,.12);color:${GREEN};
}
#${PANEL_ID} .n4obtn-no{
  border-color:rgba(255,255,255,.14);color:#d7dee8;
}
#${PANEL_ID} .n4obtn:disabled{opacity:.55;cursor:default}
#${PANEL_ID} .n4confirm{
  color:#d7dee8;font-size:12px;font-weight:800;line-height:1.35;margin-bottom:2px;
}
#${PANEL_ID} .n4bnr{
  margin-top:2px;margin-bottom:6px;padding:8px 10px;
  background:rgba(86,211,100,.1);border:1px solid rgba(86,211,100,.28);
  border-radius:10px;text-align:center;
}
#${PANEL_ID} .n4bnr label{
  display:block;color:#8b949e;font-size:9px;font-weight:800;
  letter-spacing:1px;text-transform:uppercase;margin-bottom:4px;
}
#${PANEL_ID} .n4bnr strong{
  font-size:14px;letter-spacing:.03em;color:#f0f6fc;font-weight:900;
}
#${PANEL_ID} .n4bnr-row{
  display:flex;align-items:center;justify-content:center;gap:8px;flex-wrap:wrap;margin-top:2px;
}
#${PANEL_ID} .n4copy{
  border:1px solid rgba(86,211,100,.45);border-radius:8px;
  background:rgba(13,17,23,.82);color:${GREEN};
  padding:5px 8px;cursor:pointer;font:inherit;font-size:10px;font-weight:900;
  white-space:nowrap;
}
#${PANEL_ID} .n4copy.ok{
  border-color:rgba(86,211,100,.7);background:rgba(86,211,100,.18);color:${GREEN};
}
#${PANEL_ID} .n4copy-hint{margin-top:5px;min-height:12px;font-size:10px;font-weight:700;color:${GREEN}}
#${PANEL_ID} .n4foot{
  display:flex;justify-content:space-between;gap:8px;
  margin-top:10px;padding-top:8px;border-top:1px solid rgba(255,255,255,.07);
  color:#8b949e;font-size:10px;font-weight:800;
}
#${PANEL_ID} .n4credit{
  text-align:right;margin-top:2px;color:#6e7681;font-size:9px;font-weight:600;letter-spacing:.2px;
}
`;
        document.head.appendChild(style);
    }

    async function clickWithBlink(el) {
        if (!el) throw new Error('Element fehlt');
        el.classList.add('n4-click-blink');
        try {
            el.scrollIntoView({ block: 'center', inline: 'nearest', behavior: 'smooth' });
        } catch (e) {}
        await sleep(550);
        el.click();
        await sleep(250);
        setTimeout(() => el.classList.remove('n4-click-blink'), 1800);
    }

    async function waitForControl(label, timeoutMs) {
        const timeout = timeoutMs || 45000;
        const start = Date.now();
        while (Date.now() - start < timeout) {
            if (flowAbort) throw new Error('Abgebrochen');
            const el = findControlByText(label);
            if (el) return el;
            await sleep(300);
        }
        throw new Error('Timeout: „' + label + '“ nicht gefunden');
    }

    function findN4pInText(text) {
        const t = String(text || '');
        let m = t.match(/Bestellnummer\s*[:\-]?\s*(N4P\d+)/i);
        if (m) return m[1].toUpperCase();
        m = t.match(/KNOLL[_\-\s]*N4P(\d+)/i);
        if (m) return ('N4P' + m[1]).toUpperCase();
        m = t.match(/\b(N4P\d{5,})\b/i);
        if (m) return m[1].toUpperCase();
        return '';
    }

    function findN4pInDom() {
        try {
            return findN4pInText(document.body && document.body.innerText);
        } catch (e) {
            return '';
        }
    }

    function pdfBytesToAscii(buffer) {
        const bytes = new Uint8Array(buffer);
        let out = '';
        for (let i = 0; i < bytes.length; i++) {
            const c = bytes[i];
            out += c >= 32 && c < 127 ? String.fromCharCode(c) : ' ';
        }
        return out;
    }

    function pdfBytesToUtf16Ascii(buffer) {
        const bytes = new Uint8Array(buffer);
        let out = '';
        for (let i = 0; i < bytes.length - 1; i++) {
            if (bytes[i] === 0 && bytes[i + 1] >= 32 && bytes[i + 1] < 127) {
                out += String.fromCharCode(bytes[i + 1]);
            } else if (bytes[i] >= 32 && bytes[i] < 127 && bytes[i + 1] === 0) {
                out += String.fromCharCode(bytes[i]);
                i += 1;
            }
        }
        return out;
    }

    async function inflateBytes(data) {
        for (const format of ['deflate', 'deflate-raw']) {
            try {
                const ds = new DecompressionStream(format);
                const ab = await new Response(new Blob([data]).stream().pipeThrough(ds)).arrayBuffer();
                return new Uint8Array(ab);
            } catch (e) {}
        }
        return null;
    }

    function collectPdfStreams(u8) {
        const bodies = [];
        const n = u8.length;
        for (let i = 0; i < n - 15; i++) {
            if (
                u8[i] === 0x73 && u8[i + 1] === 0x74 && u8[i + 2] === 0x72 &&
                u8[i + 3] === 0x65 && u8[i + 4] === 0x61 && u8[i + 5] === 0x6d
            ) {
                let start = i + 6;
                if (start < n && u8[start] === 0x0d) start += 1;
                if (start < n && u8[start] === 0x0a) start += 1;
                let end = -1;
                for (let j = start; j < n - 9; j++) {
                    if (
                        u8[j] === 0x65 && u8[j + 1] === 0x6e && u8[j + 2] === 0x64 &&
                        u8[j + 3] === 0x73 && u8[j + 4] === 0x74 && u8[j + 5] === 0x72 &&
                        u8[j + 6] === 0x65 && u8[j + 7] === 0x61 && u8[j + 8] === 0x6d
                    ) {
                        end = j;
                        break;
                    }
                }
                if (end > start) bodies.push(u8.subarray(start, end));
            }
        }
        return bodies;
    }

    async function extractViaInflate(buffer) {
        const u8 = new Uint8Array(buffer);
        const streams = collectPdfStreams(u8);
        let text = '';
        for (const body of streams) {
            const inflated = await inflateBytes(body);
            if (!inflated) continue;
            text += pdfBytesToAscii(inflated.buffer) + '\n';
            text += pdfBytesToUtf16Ascii(inflated.buffer) + '\n';
            const nr = findN4pInText(text);
            if (nr) return nr;
        }
        return findN4pInText(text);
    }

    async function extractViaPdfJs(buffer) {
        const pdfjsLib = window.pdfjsLib || window['pdfjs-dist/build/pdf'];
        if (!pdfjsLib || !pdfjsLib.getDocument) return '';
        try {
            pdfjsLib.GlobalWorkerOptions.workerSrc =
                'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
            const task = pdfjsLib.getDocument({
                data: new Uint8Array(buffer),
                useSystemFonts: true,
                isEvalSupported: false,
                disableFontFace: true
            });
            const doc = await task.promise;
            let text = '';
            const max = Math.min(doc.numPages || 1, 3);
            for (let i = 1; i <= max; i++) {
                const page = await doc.getPage(i);
                const content = await page.getTextContent();
                text += content.items.map((it) => it.str || '').join(' ') + '\n';
                const nr = findN4pInText(text);
                if (nr) {
                    try { doc.destroy(); } catch (e) {}
                    return nr;
                }
            }
            try { doc.destroy(); } catch (e) {}
            return findN4pInText(text);
        } catch (e) {
            return '';
        }
    }

    async function extractBestellnummer(buffer) {
        if (!buffer) return '';
        let nr = findN4pInText(pdfBytesToAscii(buffer));
        if (nr) return nr;
        nr = findN4pInText(pdfBytesToUtf16Ascii(buffer));
        if (nr) return nr;
        nr = await extractViaInflate(buffer);
        if (nr) return nr;
        nr = await extractViaPdfJs(buffer);
        if (nr) return nr;
        return '';
    }

    function rememberBestellnummer(nr) {
        if (!nr) return false;
        const cur = getFlow().bestellnummer;
        if (cur) return true;
        setFlow({ bestellnummer: String(nr).toUpperCase() });
        return true;
    }

    function resolveHref(el) {
        if (!el) return '';
        let href = el.getAttribute && (el.getAttribute('href') || '');
        if (!href && el.closest) {
            const a = el.closest('a');
            if (a) href = a.getAttribute('href') || '';
        }
        if (!href || href === '#' || href.toLowerCase().startsWith('javascript:')) return '';
        try {
            return new URL(href, location.origin).toString();
        } catch (e) {
            return '';
        }
    }

    function installPdfHook() {
        if (pdfHookInstalled) return;
        pdfHookInstalled = true;
        const origFetch = window.fetch.bind(window);

        function scanPayload(payload) {
            try {
                if (!payload) return;
                if (typeof payload === 'string') {
                    rememberBestellnummer(findN4pInText(payload));
                    return;
                }
                if (payload instanceof ArrayBuffer) {
                    extractBestellnummer(payload).then((nr) => rememberBestellnummer(nr));
                    return;
                }
                rememberBestellnummer(findN4pInText(JSON.stringify(payload)));
            } catch (e) {}
        }

        window.fetch = async function (input, init) {
            const res = await origFetch(input, init);
            try {
                const url = typeof input === 'string' ? input : (input && input.url) || '';
                const ct = (res.headers.get('content-type') || '').toLowerCase();
                const flow = getFlow();
                if (flow.stage !== 'running' && flow.stage !== 'done') return res;
                if (/\/cart\/order(?:\?|$)/i.test(url) || /\/documents\/packing\//i.test(url)) {
                    const clone = res.clone();
                    if (ct.includes('json')) {
                        clone.json().then(scanPayload).catch(() => {});
                    } else {
                        clone.arrayBuffer().then(scanPayload).catch(() => {});
                    }
                } else if (ct.includes('pdf') || /\.pdf(\?|$)/i.test(url)) {
                    res.clone().arrayBuffer().then(scanPayload).catch(() => {});
                }
            } catch (e) {}
            return res;
        };

        const OrigXHR = window.XMLHttpRequest;
        function WrappedXHR() {
            const xhr = new OrigXHR();
            let url = '';
            const open = xhr.open;
            xhr.open = function (method, u) {
                url = String(u || '');
                return open.apply(xhr, arguments);
            };
            xhr.addEventListener('load', function () {
                try {
                    const flow = getFlow();
                    if (flow.stage !== 'running' && flow.stage !== 'done') return;
                    if (!/\/cart\/order|\/documents\/packing|\.pdf/i.test(url)) return;
                    const ct = String(xhr.getResponseHeader('content-type') || '').toLowerCase();
                    if (ct.includes('json') || typeof xhr.response === 'string') {
                        scanPayload(xhr.responseText || xhr.response);
                    } else if (xhr.response instanceof ArrayBuffer) {
                        scanPayload(xhr.response);
                    }
                } catch (e) {}
            });
            return xhr;
        }
        WrappedXHR.prototype = OrigXHR.prototype;
        window.XMLHttpRequest = WrappedXHR;
    }

    async function downloadPdfBuffer(url) {
        const res = await fetch(url, {
            credentials: 'include',
            headers: { Accept: 'application/pdf,*/*' }
        });
        if (!res.ok) throw new Error('PDF Download ' + res.status);
        return res.arrayBuffer();
    }

    function offerPdfLocally(buffer, filename) {
        try {
            const blob = new Blob([buffer], { type: 'application/pdf' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename || 'Packzettel.pdf';
            a.rel = 'noopener';
            document.body.appendChild(a);
            a.click();
            a.remove();
            setTimeout(() => URL.revokeObjectURL(url), 60000);
        } catch (e) {}
    }

    async function handlePackzettelOpen(openBtn) {
        openBtn.classList.add('n4-click-blink');
        try {
            openBtn.scrollIntoView({ block: 'center', inline: 'nearest', behavior: 'smooth' });
        } catch (e) {}
        await sleep(400);

        let nr = findN4pInDom();
        if (nr) {
            rememberBestellnummer(nr);
            setTimeout(() => openBtn.classList.remove('n4-click-blink'), 1200);
            return nr;
        }

        const href = resolveHref(openBtn);
        if (!href) {
            openBtn.click();
            await sleep(1200);
            nr = findN4pInDom() || getFlow().bestellnummer;
            setTimeout(() => openBtn.classList.remove('n4-click-blink'), 1200);
            return nr || '';
        }

        setFlow({ step: 'Packzettel laden…' });
        const buf = await downloadPdfBuffer(href);
        setFlow({ step: 'Packzettel scannen…' });
        nr = await extractBestellnummer(buf);
        if (nr) rememberBestellnummer(nr);

        offerPdfLocally(buf, (nr || 'Packzettel') + '.pdf');
        setTimeout(() => openBtn.classList.remove('n4-click-blink'), 1200);
        return nr || getFlow().bestellnummer || '';
    }

    async function waitForBestellnummer(timeoutMs) {
        const timeout = timeoutMs || 12000;
        const start = Date.now();
        while (Date.now() - start < timeout) {
            if (flowAbort) throw new Error('Abgebrochen');
            const nr = getFlow().bestellnummer || findN4pInDom();
            if (nr) {
                rememberBestellnummer(nr);
                return nr;
            }
            await sleep(300);
        }
        return getFlow().bestellnummer || '';
    }

    async function runOrderFlow() {
        if (flowBusy) return;
        flowBusy = true;
        flowAbort = false;
        installPdfHook();
        const freshStart = getFlow().stage === 'confirm' || getFlow().stage === 'idle' || !getFlow().step;
        setFlow({
            stage: 'running',
            error: '',
            startedAt: Date.now(),
            ...(freshStart ? { bestellnummer: '', step: '' } : {})
        });
        try {
            const hasOpen = !!findControlByText('Packzettel öffnen');
            const hasMail = !!findControlByText('Packzettel per Email senden');
            const hasAbsenden = !!findControlByText('Bestellung absenden');
            const hasZur = !!findControlByText('Zur Bestellung');

            if (!hasOpen && !hasMail && !hasAbsenden) {
                setFlow({ step: 'Zur Bestellung…' });
                const zur = await waitForControl('Zur Bestellung');
                await clickWithBlink(zur);
            } else if (hasZur && !hasAbsenden && !hasMail && !hasOpen) {
                setFlow({ step: 'Zur Bestellung…' });
                await clickWithBlink(findControlByText('Zur Bestellung'));
            }

            if (!hasOpen && !hasMail) {
                setFlow({ step: 'Bestellung absenden…' });
                const absenden = await waitForControl('Bestellung absenden');
                await clickWithBlink(absenden);
            }

            setFlow({ step: 'Packzettel per Email…' });
            const mailBtn = await waitForControl('Packzettel per Email senden');
            await clickWithBlink(mailBtn);
            await sleep(900);

            setFlow({ step: 'Packzettel öffnen / scannen…' });
            const openBtn = await waitForControl('Packzettel öffnen');
            let nr = await handlePackzettelOpen(openBtn);
            if (!nr) nr = await waitForBestellnummer(8000);

            if (nr) {
                setFlow({ stage: 'done', step: 'Fertig', error: '', bestellnummer: nr });
            } else {
                setFlow({
                    stage: 'done',
                    step: 'Fertig (N4P nicht gelesen)',
                    error: 'PDF geladen, aber Bestellnummer (N4P…) nicht gefunden',
                    bestellnummer: ''
                });
            }
        } catch (err) {
            if (flowAbort) {
                setFlow({ stage: 'idle', step: '', error: '' });
            } else {
                setFlow({ stage: 'idle', step: '', error: String(err.message || err) });
            }
        } finally {
            flowBusy = false;
            flowAbort = false;
        }
    }

    function setStatus(text, kind) {
        const el = document.querySelector('#' + PANEL_ID + ' .n4status');
        if (!el) return;
        el.textContent = text || '';
        el.className = 'n4status' + (kind ? ' ' + kind : '');
    }

    function renderHits(hits) {
        const box = document.querySelector('#' + PANEL_ID + ' .n4hits');
        if (!box) return;
        box.innerHTML = '';
        lastHits = hits || [];
        for (const hit of lastHits) {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'n4hit';
            btn.innerHTML =
                '<strong>' + escapeHtml(hit.name) + '</strong>' +
                '<span>Seite ' + hit.page + ' · ' + hit.count + ' Artikel · ' + escapeHtml(formatDate(hit.created)) + '</span>';
            btn.addEventListener('click', () => openCart(hit));
            box.appendChild(btn);
        }
    }

    function escapeHtml(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function renderOrderSection() {
        const host = document.querySelector('#' + PANEL_ID + ' .n4order');
        if (!host) return;
        const flow = getFlow();
        let html = '<div class="n4order-title">Bestellung</div>';

        if (flow.bestellnummer) {
            html +=
                '<div class="n4bnr"><label>Bestellnummer</label>' +
                '<div class="n4bnr-row"><strong data-bnr>' +
                escapeHtml(flow.bestellnummer) +
                '</strong>' +
                '<button type="button" class="n4copy" data-copy-bnr>Kopieren</button></div>' +
                '<div class="n4copy-hint" data-copy-hint></div></div>';
        }

        if (flow.stage === 'idle') {
            html += '<button type="button" class="n4obtn n4obtn-primary" data-order-start>Bestellung Jetzt auslösen</button>';
            if (flow.error) html += '<div class="n4status err" style="margin-top:8px">' + escapeHtml(flow.error) + '</div>';
        } else if (flow.stage === 'confirm') {
            html +=
                '<div class="n4confirm">Bist du Dir sicher?</div>' +
                '<button type="button" class="n4obtn n4obtn-yes" data-order-yes>JA jetzt bestellen</button>' +
                '<button type="button" class="n4obtn n4obtn-no" data-order-no>Nein, Abbrechen</button>';
        } else if (flow.stage === 'running') {
            html +=
                '<div class="n4status ok">Läuft: ' + escapeHtml(flow.step || '…') + '</div>' +
                '<button type="button" class="n4obtn n4obtn-no" data-order-abort>Abbrechen</button>';
        } else if (flow.stage === 'done') {
            if (!flow.bestellnummer && flow.error) {
                html += '<div class="n4status err">' + escapeHtml(flow.error) + '</div>';
            } else if (flow.bestellnummer) {
                html += '<div class="n4status ok">Bestellung durch</div>';
            }
            html += '<button type="button" class="n4obtn n4obtn-primary" data-order-next>Nächste Bestellung bearbeiten</button>';
            html += '<button type="button" class="n4obtn n4obtn-no" data-order-reset>Zurücksetzen</button>';
        }

        host.innerHTML = html;
        syncPanelTone();

        const copyBtn = host.querySelector('[data-copy-bnr]');
        if (copyBtn) {
            copyBtn.addEventListener('click', async () => {
                const nr = getFlow().bestellnummer || '';
                const hint = host.querySelector('[data-copy-hint]');
                try {
                    if (navigator.clipboard && navigator.clipboard.writeText) {
                        await navigator.clipboard.writeText(nr);
                    } else {
                        const ta = document.createElement('textarea');
                        ta.value = nr;
                        ta.style.position = 'fixed';
                        ta.style.left = '-9999px';
                        document.body.appendChild(ta);
                        ta.select();
                        document.execCommand('copy');
                        ta.remove();
                    }
                    copyBtn.textContent = 'Kopiert!';
                    copyBtn.classList.add('ok');
                    if (hint) {
                        hint.textContent = nr + ' in Zwischenablage';
                        hint.style.color = '';
                    }
                    setTimeout(() => {
                        copyBtn.textContent = 'Kopieren';
                        copyBtn.classList.remove('ok');
                        if (hint) hint.textContent = '';
                    }, 1800);
                } catch (err) {
                    if (hint) {
                        hint.textContent = 'Kopieren fehlgeschlagen';
                        hint.style.color = '#f85149';
                    }
                }
            });
        }

        const start = host.querySelector('[data-order-start]');
        if (start) start.addEventListener('click', () => setFlow({ stage: 'confirm', error: '', step: '' }));

        const yes = host.querySelector('[data-order-yes]');
        if (yes) yes.addEventListener('click', () => runOrderFlow());

        const no = host.querySelector('[data-order-no]');
        if (no) no.addEventListener('click', () => setFlow({ stage: 'idle', step: '', error: '' }));

        const abort = host.querySelector('[data-order-abort]');
        if (abort) abort.addEventListener('click', () => {
            flowAbort = true;
            setFlow({ stage: 'idle', step: '', error: 'Abgebrochen' });
        });

        const next = host.querySelector('[data-order-next]');
        if (next) {
            next.addEventListener('click', () => {
                sessionStorage.removeItem(FLOW_KEY);
                location.hash = '#/cart';
                location.reload();
            });
        }

        const reset = host.querySelector('[data-order-reset]');
        if (reset) reset.addEventListener('click', () => {
            sessionStorage.removeItem(FLOW_KEY);
            renderOrderSection();
        });
    }

    async function openCart(hit) {
        setStatus('Öffne ' + hit.name + ' …');
        try {
            await selectCart(hit.id);
            sessionStorage.setItem('n4_stockid_opened', hit.id);
            location.hash = '#/cart';
            location.reload();
        } catch (err) {
            setStatus('Select fehlgeschlagen: ' + err.message, 'err');
        }
    }

    async function runSearch() {
        if (searching) return;
        const input = document.querySelector('#' + PANEL_ID + ' input');
        const btn = document.querySelector('#' + PANEL_ID + ' .n4go');
        const q = (input && input.value || '').trim();
        if (!q) {
            setStatus('StockID eingeben', 'err');
            return;
        }
        localStorage.setItem(LAST_KEY, q);
        searching = true;
        if (btn) btn.disabled = true;
        renderHits([]);
        setStatus('Suche…');
        try {
            const hits = await findStock(q, (cur, total) => {
                setStatus('Suche… Seite ' + cur + '/' + total);
            });
            if (!hits.length) {
                setStatus('Nicht gefunden: ' + q, 'err');
            } else if (hits.length === 1) {
                setStatus('Gefunden auf Seite ' + hits[0].page + ' – öffne…', 'ok');
                renderHits(hits);
                await openCart(hits[0]);
            } else {
                setStatus(hits.length + ' Treffer – bitte wählen', 'ok');
                renderHits(hits);
            }
        } catch (err) {
            setStatus('Fehler: ' + err.message, 'err');
        } finally {
            searching = false;
            if (btn) btn.disabled = false;
        }
    }

    function clampPos(left, top, panel) {
        const w = panel.offsetWidth || 280;
        const h = panel.offsetHeight || 120;
        const maxL = Math.max(0, window.innerWidth - w);
        const maxT = Math.max(0, window.innerHeight - h);
        return {
            left: Math.min(Math.max(0, left), maxL),
            top: Math.min(Math.max(0, top), maxT)
        };
    }

    function applyPos(panel, left, top) {
        const p = clampPos(left, top, panel);
        panel.style.left = p.left + 'px';
        panel.style.top = p.top + 'px';
        panel.style.right = 'auto';
        panel.style.bottom = 'auto';
        return p;
    }

    function loadPos(panel) {
        try {
            const raw = localStorage.getItem(POS_KEY);
            if (!raw) return;
            const pos = JSON.parse(raw);
            if (typeof pos.left === 'number' && typeof pos.top === 'number') {
                applyPos(panel, pos.left, pos.top);
            }
        } catch (e) {}
    }

    function savePos(panel) {
        const left = parseFloat(panel.style.left);
        const top = parseFloat(panel.style.top);
        if (Number.isNaN(left) || Number.isNaN(top)) return;
        localStorage.setItem(POS_KEY, JSON.stringify({ left, top }));
    }

    function enableDrag(panel) {
        const handle = panel.querySelector('.n4h');
        if (!handle) return;

        function onPointerDown(e) {
            if (e.button != null && e.button !== 0) return;
            if (e.target && e.target.closest('[data-close]')) return;
            const rect = panel.getBoundingClientRect();
            dragState = {
                panel,
                offsetX: e.clientX - rect.left,
                offsetY: e.clientY - rect.top,
                pointerId: e.pointerId
            };
            panel.classList.add('n4dragging');
            try {
                handle.setPointerCapture(e.pointerId);
            } catch (err) {}
            e.preventDefault();
        }

        function onPointerMove(e) {
            if (!dragState || dragState.panel !== panel) return;
            applyPos(panel, e.clientX - dragState.offsetX, e.clientY - dragState.offsetY);
        }

        function onPointerUp(e) {
            if (!dragState || dragState.panel !== panel) return;
            applyPos(panel, e.clientX - dragState.offsetX, e.clientY - dragState.offsetY);
            savePos(panel);
            panel.classList.remove('n4dragging');
            dragState = null;
        }

        handle.addEventListener('pointerdown', onPointerDown);
        handle.addEventListener('pointermove', onPointerMove);
        handle.addEventListener('pointerup', onPointerUp);
        handle.addEventListener('pointercancel', onPointerUp);

        window.addEventListener('resize', () => {
            const left = parseFloat(panel.style.left);
            const top = parseFloat(panel.style.top);
            if (Number.isNaN(left) || Number.isNaN(top)) return;
            const p = applyPos(panel, left, top);
            localStorage.setItem(POS_KEY, JSON.stringify(p));
        });
    }

    function createPanel() {
        if (document.getElementById(PANEL_ID)) {
            renderOrderSection();
            return;
        }
        ensureStyles();
        const panel = document.createElement('div');
        panel.id = PANEL_ID;
        panel.setAttribute('data-tone', 'idle');
        panel.innerHTML =
            '<div class="n4h">' +
            '  <div>' +
            '    <div class="n4h-eyebrow">N4PARTS · STOCKID</div>' +
            '    <div class="n4h-title">Warenkorb Suche</div>' +
            '  </div>' +
            '  <div class="n4h-actions">' +
            '    <button type="button" class="n4h-close" title="Schließen" data-close>X</button>' +
            '  </div>' +
            '</div>' +
            '<div class="n4b">' +
            '  <div class="n4badge" data-tone-badge>BEREIT</div>' +
            '  <div class="n4row">' +
            '    <input type="text" placeholder="z.B. RC28374" autocomplete="off" spellcheck="false">' +
            '    <button type="button" class="n4go">Suchen</button>' +
            '  </div>' +
            '  <div class="n4status"></div>' +
            '  <div class="n4hits"></div>' +
            '  <div class="n4order"></div>' +
            '  <div class="n4foot"><span>Bot v' + BOT_VERSION + '</span><span>Stock · Order</span></div>' +
            '  <div class="n4credit">by Arasaka</div>' +
            '</div>';
        document.body.appendChild(panel);
        loadPos(panel);
        enableDrag(panel);
        renderOrderSection();
        syncPanelTone();

        const input = panel.querySelector('input');
        const last = localStorage.getItem(LAST_KEY);
        if (last) input.value = last;

        panel.querySelector('.n4go').addEventListener('click', runSearch);
        panel.querySelector('[data-close]').addEventListener('click', () => panel.remove());
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                runSearch();
            }
        });
    }

    function syncPanel() {
        if (shouldShowPanel()) {
            createPanel();
        } else {
            const el = document.getElementById(PANEL_ID);
            if (el) el.remove();
        }
    }

    if (getFlow().stage === 'running') {
        installPdfHook();
        setTimeout(() => {
            if (getFlow().stage === 'running' && !flowBusy) runOrderFlow();
        }, 800);
    }

    let lastHash = location.hash;
    let lastFlowSig = sessionStorage.getItem(FLOW_KEY) || '';
    setInterval(() => {
        if (location.hash !== lastHash) {
            lastHash = location.hash;
            syncPanel();
        } else if (shouldShowPanel() && !document.getElementById(PANEL_ID)) {
            syncPanel();
        }
        const sig = sessionStorage.getItem(FLOW_KEY) || '';
        if (sig !== lastFlowSig) {
            lastFlowSig = sig;
            if (document.getElementById(PANEL_ID)) renderOrderSection();
        }
    }, 500);

    syncPanel();
})();
