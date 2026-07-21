// ==UserScript==
// @name         N4Parts StockID Warenkorb Suche
// @namespace    http://tampermonkey.net/
// @version      1.2
// @description  Sucht Warenkörbe nach StockID über alle Seiten via /api/cart/list
// @author       ARASAKA
// @match        https://www.n4parts.net/*
// @match        https://n4parts.net/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    const PANEL_ID = 'n4-stockid-search-panel';
    const LAST_KEY = 'n4_stockid_last_query';
    const POS_KEY = 'n4_stockid_panel_pos';
    const PAGE_SIZE = 20;
    const ACCENT = '#5CB8B2';

    let searching = false;
    let lastHits = [];
    let dragState = null;

    function isWarenkoerbePage() {
        const h = (location.hash || '').split('?')[0];
        return h === '#/cart' || h === '#/cart/';
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

    function ensureStyles() {
        if (document.getElementById(PANEL_ID + '-style')) return;
        const style = document.createElement('style');
        style.id = PANEL_ID + '-style';
        style.textContent = `
#${PANEL_ID}{
  position:fixed;top:72px;left:12px;z-index:2147483646;
  width:280px;max-width:calc(100vw - 24px);
  background:#fff;color:#3c3c3b;
  border:1px solid #bcbcba;border-radius:4px;
  box-shadow:0 4px 18px rgba(0,0,0,.18);
  font-family:Roboto,Arial,sans-serif;font-size:12px;line-height:1.35;
  user-select:none;
}
#${PANEL_ID} .n4h{
  display:flex;align-items:center;justify-content:space-between;
  background:#3c3c3b;color:#fff;padding:8px 10px;font-weight:500;
  cursor:move;touch-action:none;
}
#${PANEL_ID} .n4h button{
  background:transparent;border:0;color:#fff;cursor:pointer;font-size:14px;padding:0 2px;
}
#${PANEL_ID}.n4dragging{opacity:.92;box-shadow:0 8px 28px rgba(0,0,0,.28)}
#${PANEL_ID} .n4b,#${PANEL_ID} input{user-select:text}
#${PANEL_ID} .n4b{padding:10px}
#${PANEL_ID} .n4row{display:flex;gap:6px}
#${PANEL_ID} input[type="text"]{
  flex:1;min-width:0;border:1px solid #bcbcba;border-radius:2px;
  padding:6px 8px;font:inherit;outline:none;
}
#${PANEL_ID} input[type="text"]:focus{border-color:${ACCENT}}
#${PANEL_ID} .n4go{
  border:0;border-radius:2px;background:${ACCENT};color:#fff;
  padding:6px 10px;cursor:pointer;font:inherit;font-weight:500;white-space:nowrap;
}
#${PANEL_ID} .n4go:disabled{opacity:.55;cursor:default}
#${PANEL_ID} .n4status{margin-top:8px;color:#666;min-height:16px}
#${PANEL_ID} .n4status.ok{color:#1f7a4d}
#${PANEL_ID} .n4status.err{color:#b00020}
#${PANEL_ID} .n4hits{margin-top:8px;max-height:220px;overflow:auto;border-top:1px solid #e6e6e6}
#${PANEL_ID} .n4hit{
  display:block;width:100%;text-align:left;border:0;border-bottom:1px solid #eee;
  background:#fff;padding:8px 6px;cursor:pointer;font:inherit;color:inherit;
}
#${PANEL_ID} .n4hit:hover{background:#f3fafa}
#${PANEL_ID} .n4hit strong{display:block;color:#222}
#${PANEL_ID} .n4hit span{color:#777;font-size:11px}
`;
        document.head.appendChild(style);
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
        if (document.getElementById(PANEL_ID)) return;
        ensureStyles();
        const panel = document.createElement('div');
        panel.id = PANEL_ID;
        panel.innerHTML =
            '<div class="n4h"><span>StockID suchen</span><button type="button" title="Schließen" data-close>×</button></div>' +
            '<div class="n4b">' +
            '<div class="n4row">' +
            '<input type="text" placeholder="z.B. RC28374" autocomplete="off" spellcheck="false">' +
            '<button type="button" class="n4go">Suchen</button>' +
            '</div>' +
            '<div class="n4status"></div>' +
            '<div class="n4hits"></div>' +
            '</div>';
        document.body.appendChild(panel);
        loadPos(panel);
        enableDrag(panel);

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
        if (isWarenkoerbePage()) {
            createPanel();
        } else {
            const el = document.getElementById(PANEL_ID);
            if (el) el.remove();
        }
    }

    let lastHash = location.hash;
    setInterval(() => {
        if (location.hash !== lastHash) {
            lastHash = location.hash;
            syncPanel();
        } else if (isWarenkoerbePage() && !document.getElementById(PANEL_ID)) {
            syncPanel();
        }
    }, 500);

    syncPanel();
})();
