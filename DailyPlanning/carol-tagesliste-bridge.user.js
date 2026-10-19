// ==UserScript==
// @name         Carol Tagesliste Herausgegeben Bridge
// @namespace    http://tampermonkey.net/
// @version      1.3
// @description  Carol Herausgegeben Automatisierung // Tagesliste
// @author       ARASAKA
// @match        *://carol.autohero.com/*
// @match        *://docs.google.com/spreadsheets/*
// @match        *://script.google.com/*
// @match        *://*.googleusercontent.com/*
// @updateURL    https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/DailyPlanning/carol-tagesliste-bridge.user.js
// @downloadURL  https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/DailyPlanning/carol-tagesliste-bridge.user.js
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addValueChangeListener
// ==/UserScript==

(function () {
    'use strict';

    const KEY_QUEUE = 'dpl_tl_carol_queue_v4';
    const KEY_STATE = 'dpl_tl_carol_state_v4';
    const KEY_EVENTS = 'dpl_tl_carol_events_v4';
    const KEY_ABORT = 'dpl_tl_carol_abort_v4';
    const KEY_HEARTBEAT = 'dpl_tl_carol_hb_v4';

    try {
        if (typeof GM_setValue === 'function') {
            GM_setValue('dpl_tl_carol_queue_v1', '[]');
            GM_setValue('dpl_tl_carol_queue_v2', '[]');
            GM_setValue('dpl_tl_carol_queue_v3', '[]');
        }
    } catch (e0) {}

    const isCarolHost = location.hostname.includes('carol.autohero.com');
    const isWmsSingleAuto = isCarolHost && location.href.includes('arasaka_auto=1');

    function readJson(key, fallback) {
        try {
            if (typeof GM_getValue === 'function') {
                const rawGm = GM_getValue(key, '');
                if (rawGm) return JSON.parse(rawGm);
            }
        } catch (e1) {}
        return fallback;
    }

    function writeJson(key, val) {
        try {
            if (typeof GM_setValue === 'function') {
                GM_setValue(key, JSON.stringify(val));
            }
        } catch (e1) {}
    }

    if (!isCarolHost) {
        let lastEventCursor = 0;
        let lastStateTs = 0;
        const SIDEBAR_WIDTH_KEY = 'dpl_tl_sidebar_width_px';

        function applySidebarWidthOnSheets(widthPx) {
            if (!location.hostname.includes('docs.google.com')) return;
            const sidebar = document.querySelector('.script-application-sidebar');
            if (!sidebar) return;
            const clamped = Math.max(260, Math.min(Math.floor(window.innerWidth * 0.75), Number(widthPx) || 360));
            sidebar.style.setProperty('width', clamped + 'px', 'important');
            sidebar.style.setProperty('min-width', clamped + 'px', 'important');
            sidebar.style.setProperty('max-width', clamped + 'px', 'important');
            const inner = sidebar.querySelectorAll('.script-application-sidebar-content, iframe');
            for (let i = 0; i < inner.length; i++) {
                inner[i].style.setProperty('width', '100%', 'important');
            }
            try { localStorage.setItem(SIDEBAR_WIDTH_KEY, String(clamped)); } catch (e) {}
        }

        function ensureSheetsSidebarResizer() {
            if (!location.hostname.includes('docs.google.com')) return;
            const sidebar = document.querySelector('.script-application-sidebar');
            if (!sidebar) return;

            if (!document.getElementById('dpl-sheets-sidebar-resizer-style')) {
                const st = document.createElement('style');
                st.id = 'dpl-sheets-sidebar-resizer-style';
                st.textContent = `
                    .script-application-sidebar {
                        overflow: visible !important;
                    }
                    #dpl-sheets-sidebar-resizer {
                        position: absolute;
                        left: -5px;
                        top: 0;
                        bottom: 0;
                        width: 10px;
                        cursor: ew-resize;
                        z-index: 2147483647;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        background: transparent;
                        transition: background 0.15s ease;
                    }
                    #dpl-sheets-sidebar-resizer::after {
                        content: "";
                        width: 4px;
                        height: 44px;
                        border-radius: 999px;
                        background: rgba(0, 255, 204, 0.35);
                        box-shadow: 0 0 6px rgba(0, 255, 204, 0.25);
                        transition: background 0.15s ease, height 0.15s ease, transform 0.15s ease;
                    }
                    #dpl-sheets-sidebar-resizer:hover,
                    #dpl-sheets-sidebar-resizer.dragging {
                        background: rgba(0, 255, 204, 0.12);
                    }
                    #dpl-sheets-sidebar-resizer:hover::after,
                    #dpl-sheets-sidebar-resizer.dragging::after {
                        background: #00ffcc;
                        height: 64px;
                        transform: scaleX(1.2);
                    }
                `;
                document.head.appendChild(st);
            }

            if (!sidebar.dataset.dplWidthInitialized) {
                sidebar.dataset.dplWidthInitialized = '1';
                let savedW = 360;
                try {
                    const rawW = parseInt(localStorage.getItem(SIDEBAR_WIDTH_KEY) || '360', 10);
                    if (rawW >= 260 && rawW <= 1200) savedW = rawW;
                } catch (e) {}
                applySidebarWidthOnSheets(savedW);
            }

            if (sidebar.querySelector('#dpl-sheets-sidebar-resizer')) return;

            const handle = document.createElement('div');
            handle.id = 'dpl-sheets-sidebar-resizer';
            handle.title = 'Ziehen, um die Seitenleiste größer oder kleiner zu machen';
            sidebar.appendChild(handle);

            handle.addEventListener('mousedown', (e) => {
                if (e.button !== 0) return;
                e.preventDefault();
                e.stopPropagation();
                handle.classList.add('dragging');

                const shield = document.createElement('div');
                shield.style.cssText = 'position:fixed;inset:0;z-index:2147483646;cursor:ew-resize;background:transparent;';
                document.body.appendChild(shield);

                function onMove(ev) {
                    const newWidth = window.innerWidth - ev.clientX;
                    applySidebarWidthOnSheets(newWidth);
                }

                function onUp() {
                    handle.classList.remove('dragging');
                    shield.remove();
                    window.removeEventListener('mousemove', onMove, true);
                    window.removeEventListener('mouseup', onUp, true);
                }

                window.addEventListener('mousemove', onMove, true);
                window.addEventListener('mouseup', onUp, true);
            });
        }

        function postToFrameTree(win, payload, depth) {
            if (!win || depth > 4) return;
            try { win.postMessage(payload, '*'); } catch (e0) {}
            try {
                for (let i = 0; i < win.frames.length; i++) {
                    postToFrameTree(win.frames[i], payload, depth + 1);
                }
            } catch (e1) {}
        }

        function forwardToHud() {
            const hb = readJson(KEY_HEARTBEAT, null);
            if (hb && (Date.now() - Number(hb.ts || 0)) < 25000) {
                postToFrameTree(window, {
                    source: 'dpl-tagesliste-bridge',
                    type: 'heartbeat',
                    idle: !!hb.idle,
                    ts: hb.ts
                }, 0);
            }

            const st = readJson(KEY_STATE, null);
            if (st && Number(st.ts || 0) > lastStateTs && (Date.now() - Number(st.ts || 0)) < 25000) {
                lastStateTs = Number(st.ts || 0);
                postToFrameTree(window, {
                    source: 'dpl-tagesliste-bridge',
                    type: 'state',
                    busy: !!st.busy,
                    currentSid: st.currentSid || '',
                    stepText: st.stepText || '',
                    ts: st.ts
                }, 0);
            }

            const evs = readJson(KEY_EVENTS, []);
            if (Array.isArray(evs) && evs.length) {
                for (let i = 0; i < evs.length; i++) {
                    const ev = evs[i];
                    if (!ev || !ev.stockId) continue;
                    if (Number(ev.ts || 0) <= lastEventCursor) continue;
                    postToFrameTree(window, {
                        source: 'dpl-tagesliste-bridge',
                        type: 'item_result',
                        eventId: ev.eventId,
                        stockId: ev.stockId,
                        dateStr: ev.dateStr || '',
                        rows: ev.rows || [],
                        status: ev.status || 'ok',
                        message: ev.message || '',
                        ts: ev.ts
                    }, 0);
                }
                const maxTs = evs.reduce((m, x) => Math.max(m, Number((x && x.ts) || 0)), lastEventCursor);
                lastEventCursor = maxTs;
            }
        }

        window.addEventListener('message', (ev) => {
            const d = ev && ev.data;
            if (!d || d.source !== 'dpl-tagesliste-hud') return;

            if (d.cmd === 'ping') {
                ensureSheetsSidebarResizer();
                forwardToHud();
                return;
            }

            if (d.cmd === 'resize_sidebar' && d.payload) {
                ensureSheetsSidebarResizer();
                const curSidebar = document.querySelector('.script-application-sidebar');
                if (curSidebar) {
                    const curW = curSidebar.getBoundingClientRect().width || 360;
                    if (typeof d.payload.delta === 'number') {
                        applySidebarWidthOnSheets(curW + d.payload.delta);
                    } else if (typeof d.payload.width === 'number') {
                        applySidebarWidthOnSheets(d.payload.width);
                    }
                }
                return;
            }

            if (d.cmd === 'stop_all') {
                writeJson(KEY_ABORT, { ts: Date.now() });
                writeJson(KEY_QUEUE, []);
                writeJson(KEY_STATE, { busy: false, currentSid: '', stepText: '', ts: Date.now() });
                return;
            }

            if (d.cmd === 'enqueue_jobs') {
                const incoming = (d.payload && d.payload.jobs) || [];
                if (!incoming.length) return;
                const bySid = {};
                const freshQueue = [];
                for (let j = 0; j < incoming.length; j++) {
                    const inc = incoming[j];
                    if (!inc || !inc.stockId) continue;
                    const sid = String(inc.stockId).trim().toUpperCase();
                    if (!bySid[sid]) {
                        const job = {
                            stockId: sid,
                            dateStr: inc.dateStr || '',
                            carolUrl: inc.carolUrl || ('https://carol.autohero.com/en-GB/refurbishment?rsv=' + encodeURIComponent(sid)),
                            model: inc.model || '',
                            rows: inc.rows || [],
                            row: inc.row || 0,
                            status: 'pending',
                            phase: 'init',
                            addedAt: Date.now()
                        };
                        bySid[sid] = job;
                        freshQueue.push(job);
                    }
                }
                writeJson(KEY_ABORT, { ts: 0 });
                writeJson(KEY_QUEUE, freshQueue);
            }
        });

        if (typeof GM_addValueChangeListener === 'function') {
            GM_addValueChangeListener(KEY_STATE, () => forwardToHud());
            GM_addValueChangeListener(KEY_EVENTS, () => forwardToHud());
            GM_addValueChangeListener(KEY_HEARTBEAT, () => forwardToHud());
        }
        setInterval(() => {
            ensureSheetsSidebarResizer();
            forwardToHud();
        }, 350);
        ensureSheetsSidebarResizer();
        return;
    }

    if (isWmsSingleAuto) return;

    if (location.href.includes('tl_standby=1') || location.href.includes('tl_job=1')) {
        try { sessionStorage.setItem('dpl_tl_worker_tab', '1'); } catch (e) {}
    }

    let workerBusy = false;
    let isNavigating = false;
    let abortFlag = false;
    let currentSid = '';
    let currentStepText = '';
    let hudEl = null;

    const initialQ = readJson(KEY_QUEUE, []);
    if (Array.isArray(initialQ) && initialQ.length > 0 && initialQ[0] && initialQ[0].stockId) {
        currentSid = initialQ[0].stockId;
        currentStepText = 'Lade Auftrag in Carol…';
        try { sessionStorage.setItem('dpl_tl_worker_tab', '1'); } catch (e) {}
    }

    function isWorkerTab() {
        try {
            if (sessionStorage.getItem('dpl_tl_worker_tab') === '1') return true;
        } catch (e) {}
        const q = readJson(KEY_QUEUE, []);
        return Array.isArray(q) && q.length > 0;
    }

    function postToOpenerTree(win, payload, depth) {
        if (!win || depth > 4) return;
        try { win.postMessage(payload, '*'); } catch (e0) {}
        try {
            for (let i = 0; i < win.frames.length; i++) {
                postToOpenerTree(win.frames[i], payload, depth + 1);
            }
        } catch (e1) {}
    }

    function sendToOpenerAllFrames(payload) {
        try {
            if (window.opener && !window.opener.closed) {
                let root = window.opener;
                try { if (root.top) root = root.top; } catch (e) {}
                postToOpenerTree(root, payload, 0);
            }
        } catch (e) {}
    }

    function sendHeartbeat() {
        const ts = Date.now();
        const q = readJson(KEY_QUEUE, []);
        const isIdle = !workerBusy && !isNavigating && (!Array.isArray(q) || q.length === 0);
        writeJson(KEY_HEARTBEAT, { ts: ts, idle: isIdle });
        sendToOpenerAllFrames({
            source: 'dpl-tagesliste-bridge',
            type: 'heartbeat',
            idle: isIdle,
            ts: ts
        });
    }

    function publishActiveState(busy, sid, stepText) {
        const payload = {
            busy: !!busy,
            currentSid: sid || '',
            stepText: stepText || '',
            ts: Date.now()
        };
        writeJson(KEY_STATE, payload);
        writeJson(KEY_HEARTBEAT, { ts: payload.ts, idle: !payload.busy });
        sendToOpenerAllFrames({
            source: 'dpl-tagesliste-bridge',
            type: 'state',
            busy: payload.busy,
            currentSid: payload.currentSid,
            stepText: payload.stepText,
            ts: payload.ts
        });
    }

    function pushResultEvent(stockId, status, message, jobMeta) {
        const ev = {
            eventId: stockId + '_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
            stockId: stockId,
            dateStr: (jobMeta && jobMeta.dateStr) || '',
            rows: (jobMeta && jobMeta.rows) || [],
            status: status,
            message: message || '',
            ts: Date.now()
        };
        const list = readJson(KEY_EVENTS, []);
        const trimmed = Array.isArray(list) ? list.slice(-40) : [];
        trimmed.push(ev);
        writeJson(KEY_EVENTS, trimmed);

        const msgPayload = {
            source: 'dpl-tagesliste-bridge',
            type: 'item_result',
            eventId: ev.eventId,
            stockId: ev.stockId,
            dateStr: ev.dateStr,
            rows: ev.rows,
            status: ev.status,
            message: ev.message,
            ts: ev.ts
        };
        sendToOpenerAllFrames(msgPayload);
        let retrySend = 0;
        const t = setInterval(() => {
            retrySend++;
            sendToOpenerAllFrames(msgPayload);
            if (retrySend >= 6) clearInterval(t);
        }, 250);
    }

    function ensureHud() {
        if (!document.body || !isWorkerTab()) return;
        if (!document.getElementById('dpl-tl-bridge-style')) {
            const st = document.createElement('style');
            st.id = 'dpl-tl-bridge-style';
            st.textContent = `
                #dpl-tl-bridge-hud {
                    position: fixed;
                    bottom: 20px;
                    left: 20px;
                    width: 350px;
                    background: rgba(8, 12, 18, 0.96);
                    border: 1.5px solid #00ffcc;
                    border-radius: 10px;
                    color: #e6edf3;
                    font-family: "Segoe UI", Consolas, monospace;
                    font-size: 13px;
                    z-index: 999998;
                    box-shadow: 0 10px 30px rgba(0, 0, 0, 0.55), 0 0 18px rgba(0, 255, 204, 0.2);
                    overflow: hidden;
                    user-select: none;
                }
                .dpl-tl-head {
                    display: flex;
                    align-items: center;
                    justify-content: space-between;
                    padding: 9px 12px;
                    background: rgba(0, 255, 204, 0.1);
                    border-bottom: 1px solid rgba(0, 255, 204, 0.25);
                    font-weight: 800;
                    font-size: 12px;
                    letter-spacing: 0.8px;
                    color: #00ffcc;
                }
                .dpl-tl-body {
                    padding: 10px 12px 6px;
                }
                .dpl-tl-sid {
                    font-size: 15px;
                    font-weight: 800;
                    color: #ffffff;
                    margin-bottom: 4px;
                }
                .dpl-tl-step {
                    font-size: 12.5px;
                    color: #8bcef7;
                    white-space: nowrap;
                    overflow: hidden;
                    text-overflow: ellipsis;
                }
                .dpl-tl-foot {
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    padding: 4px 12px 8px;
                    font-size: 10.5px;
                    color: #6e7681;
                    font-weight: 600;
                }
                .dpl-tl-stop {
                    background: rgba(248, 81, 73, 0.18);
                    border: 1px solid #f85149;
                    color: #ff7b72;
                    border-radius: 5px;
                    padding: 2px 8px;
                    font-size: 11px;
                    font-weight: 700;
                    cursor: pointer;
                }
            `;
            document.head.appendChild(st);
        }
        if (!hudEl || !document.body.contains(hudEl)) {
            hudEl = document.createElement('div');
            hudEl.id = 'dpl-tl-bridge-hud';
            hudEl.innerHTML = `
                <div class="dpl-tl-head">
                    <span>TAGESLISTE // CAROL BRIDGE</span>
                    <button type="button" class="dpl-tl-stop" id="dpl-tl-stop-btn">STOP</button>
                </div>
                <div class="dpl-tl-body">
                    <div class="dpl-tl-sid" id="dpl-tl-sid-txt">${currentSid || 'BEREIT'}</div>
                    <div class="dpl-tl-step" id="dpl-tl-step-txt">${currentStepText || 'Bereit für Stock-IDs vom HUD'}</div>
                </div>
                <div class="dpl-tl-foot">
                    <span>Carol Bridge v1.3 · [ESC] Stop</span>
                    <span>By Arasaka</span>
                </div>
            `;
            document.body.appendChild(hudEl);
            const stopBtn = document.getElementById('dpl-tl-stop-btn');
            if (stopBtn) {
                stopBtn.addEventListener('click', () => {
                    abortFlag = true;
                    workerBusy = false;
                    isNavigating = false;
                    writeJson(KEY_QUEUE, []);
                    setHudStatus('GESTOPPT', 'Warteschlange geleert', '#f85149', false);
                });
            }
        }
    }

    function setHudStatus(sidLabel, stepText, borderColor, isBusyState = true) {
        currentStepText = stepText || '';
        ensureHud();
        if (hudEl) {
            if (borderColor) hudEl.style.borderColor = borderColor;
            const sidEl = document.getElementById('dpl-tl-sid-txt');
            const stEl = document.getElementById('dpl-tl-step-txt');
            if (sidEl) sidEl.textContent = sidLabel || 'BEREIT';
            if (stEl) stEl.textContent = stepText || '';
        }
        if (isBusyState && currentSid) {
            publishActiveState(true, currentSid, stepText);
        } else if (!isBusyState) {
            publishActiveState(false, '', '');
        }
    }

    function playTone(ok) {
        try {
            const ctx = new (window.AudioContext || window.webkitAudioContext)();
            const osc = ctx.createOscillator();
            const gain = ctx.createGain();
            osc.connect(gain);
            gain.connect(ctx.destination);
            osc.type = ok ? 'sine' : 'sawtooth';
            osc.frequency.setValueAtTime(ok ? 880 : 180, ctx.currentTime);
            if (ok) osc.frequency.exponentialRampToValueAtTime(1560, ctx.currentTime + 0.12);
            gain.gain.setValueAtTime(0.18, ctx.currentTime);
            gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.35);
            osc.start();
            osc.stop(ctx.currentTime + 0.35);
        } catch (e) {}
    }

    async function sleep(ms) {
        let elapsed = 0;
        const step = Math.min(50, ms);
        while (elapsed < ms) {
            if (abortFlag) throw new Error('Abort');
            await new Promise(r => setTimeout(r, step));
            elapsed += step;
        }
    }

    const AUTH_STORAGE_KEY = 'authData';
    const TOKEN_URL = '/api/v1/auth/oauth/token';
    let workshopIdsCache = null;
    let refreshInFlight = null;

    function readAuth() {
        try {
            const raw = localStorage.getItem(AUTH_STORAGE_KEY);
            if (!raw) return null;
            const parsed = JSON.parse(raw);
            if (!parsed || typeof parsed !== 'object') return null;
            return parsed;
        } catch (e) {
            return null;
        }
    }

    function writeAuth(auth) {
        localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify(auth));
    }

    function authPauseReason() {
        try { return sessionStorage.getItem('dpl_tl_auth_pause') || ''; } catch (e) { return ''; }
    }

    function setAuthPause(msg) {
        try { sessionStorage.setItem('dpl_tl_auth_pause', msg || '1'); } catch (e) {}
    }

    function clearAuthPause() {
        try { sessionStorage.removeItem('dpl_tl_auth_pause'); } catch (e) {}
    }

    function decodeJwt(token) {
        const part = String(token || '').split('.')[1];
        if (!part) return null;
        try {
            const b64 = part.replace(/-/g, '+').replace(/_/g, '/');
            const pad = b64 + '==='.slice((b64.length + 3) % 4);
            const json = atob(pad);
            try {
                return JSON.parse(decodeURIComponent(escape(json)));
            } catch (e2) {
                return JSON.parse(json);
            }
        } catch (e) {
            return null;
        }
    }

    function tokenStillValid(token, skewMs) {
        const payload = decodeJwt(token);
        if (!payload || !payload.exp) return !!token;
        return (Number(payload.exp) * 1000) > (Date.now() + (skewMs || 0));
    }

    async function refreshAuth() {
        if (refreshInFlight) return refreshInFlight;
        refreshInFlight = (async () => {
            const cur = readAuth();
            if (!cur || !cur.refresh_token) {
                throw new Error('Carol Login abgelaufen, bitte neu anmelden');
            }
            const body = new URLSearchParams();
            body.set('client_id', 'refurbishment');
            body.set('grant_type', 'refresh_token');
            body.set('refresh_token', cur.refresh_token);
            const headers = {
                'Content-Type': 'application/x-www-form-urlencoded',
                'Accept': 'application/json'
            };
            if (cur.access_token) headers.authorization = 'Bearer ' + cur.access_token;
            const res = await fetch(TOKEN_URL, { method: 'POST', headers: headers, body: body.toString() });
            let data = null;
            try { data = await res.json(); } catch (e) { data = null; }
            const tokenObj = (data && data.access_token) ? data : (data && data.result && data.result.access_token ? data.result : null);
            if (!res.ok || !tokenObj) {
                throw new Error('Carol Token konnte nicht erneuert werden, bitte neu anmelden');
            }
            const next = Object.assign({}, cur, tokenObj);
            if (!next.refresh_token) next.refresh_token = cur.refresh_token;
            writeAuth(next);
            try { sessionStorage.setItem('dpl_tl_auth_refresh_ts', String(Date.now())); } catch (e2) {}
            return next;
        })();
        try {
            return await refreshInFlight;
        } finally {
            refreshInFlight = null;
        }
    }

    async function ensureFreshLogin() {
        const auth = readAuth();
        if (!auth || !auth.access_token) {
            throw new Error('Nicht in Carol eingeloggt');
        }
        if (tokenStillValid(auth.access_token, 90000)) return auth.access_token;
        let recent = 0;
        try { recent = Number(sessionStorage.getItem('dpl_tl_auth_refresh_ts') || 0); } catch (e) {}
        if (recent && (Date.now() - recent) < 20000) {
            if (tokenStillValid(auth.access_token, 0)) return auth.access_token;
            throw new Error('Carol Login abgelaufen, bitte neu anmelden');
        }
        setHudStatus(currentSid || 'LOGIN', 'Token erneuert, Carol lädt neu…', '#ffaa00', true);
        await refreshAuth();
        isNavigating = true;
        location.reload();
        throw new Error('Reload');
    }

    function assertGql(arr) {
        if (!Array.isArray(arr)) throw new Error('Carol API Antwort ungültig');
        const msgs = [];
        for (let i = 0; i < arr.length; i++) {
            const row = arr[i];
            if (row && Array.isArray(row.errors) && row.errors.length) {
                for (let j = 0; j < row.errors.length; j++) {
                    msgs.push((row.errors[j] && row.errors[j].message) || 'GraphQL Fehler');
                }
            }
        }
        if (msgs.length) throw new Error(msgs.join('; '));
        return arr;
    }

    async function gqlBatch(ops) {
        const token = await ensureFreshLogin();
        const version = (window.REACT_CFG && window.REACT_CFG.VITE_INFO_VERSION) || 'v4.32.1';
        const res = await fetch('/api/v1/refurbishment-aggregation/graphql', {
            method: 'POST',
            headers: {
                'content-type': 'application/json',
                'accept-language': 'en-GB',
                'x-carol-frontend-version': version,
                'authorization': 'Bearer ' + token
            },
            body: JSON.stringify(ops)
        });
        if (res.status === 401) {
            let recent = 0;
            try { recent = Number(sessionStorage.getItem('dpl_tl_auth_refresh_ts') || 0); } catch (e) {}
            if (recent && (Date.now() - recent) < 20000) {
                throw new Error('Carol Login abgelaufen, bitte neu anmelden');
            }
            setHudStatus(currentSid || 'LOGIN', 'Token erneuert, Carol lädt neu…', '#ffaa00', true);
            await refreshAuth();
            isNavigating = true;
            location.reload();
            throw new Error('Reload');
        }
        if (!res.ok) throw new Error('Carol API HTTP ' + res.status);
        return assertGql(await res.json());
    }

    async function gqlOne(operationName, query, variables) {
        const arr = await gqlBatch([{ operationName: operationName, variables: variables, query: query }]);
        return arr[0] && arr[0].data;
    }

    async function loadWorkshopIds() {
        if (workshopIdsCache && workshopIdsCache.length) return workshopIdsCache;
        const auth = readAuth();
        const payload = decodeJwt(auth && auth.access_token);
        const userId = payload && (payload.user_uuid || payload.userUuid || payload.sub);
        if (!userId) throw new Error('Carol Benutzer konnte nicht gelesen werden');
        const data = await gqlOne(
            'user',
            'query user($id: UUID!) { user(id: $id) { userPartnerWorkshopConnections { mainWorkshop { branchId } } } }',
            { id: userId }
        );
        const conns = (data && data.user && data.user.userPartnerWorkshopConnections) || [];
        const ids = [];
        for (let i = 0; i < conns.length; i++) {
            const branchId = conns[i] && conns[i].mainWorkshop && conns[i].mainWorkshop.branchId;
            const branchNum = Number(branchId);
            if (Number.isFinite(branchNum) && branchNum > 0 && ids.indexOf(branchNum) === -1) ids.push(branchNum);
        }
        if (!ids.length) throw new Error('Keine Carol Werkstatt am Login gefunden');
        workshopIdsCache = ids;
        return ids;
    }

    function getEffectiveVal(item) {
        const vals = item.values || [];
        const pa = vals.find(v => v.sourceType === 'POST_AUTHORIZATION' && v.proposedValue);
        if (pa) return { payload: pa.proposedValue, state: pa.state };
        const auth = vals.find(v => v.sourceType === 'AUTHORIZATION' && v.proposedValue);
        if (auth) return { payload: auth.proposedValue, state: auth.state };
        const ecVal = vals.find(v => v.sourceType === 'ENTRY_CHECK');
        if (ecVal) {
            if (item.verified && ecVal.proposedValue) return { payload: ecVal.proposedValue, state: ecVal.state };
            if (ecVal.originalValue) return { payload: ecVal.originalValue, state: ecVal.state };
            if (ecVal.proposedValue) return { payload: ecVal.proposedValue, state: ecVal.state };
        }
        return null;
    }

    function buildCleanProposed(payload, spareStatus) {
        if (!payload) return null;
        const pt = payload.payloadType;
        if (pt === 'VEHICLE_SERVICE_NAME' || pt === 'DAMAGE_SERVICE_NAME') {
            return { payloadType: pt, name: payload.name, type: payload.type };
        }
        if (pt === 'SERVICE_BUDGET') {
            const budget = { payloadType: pt, budgetMinorUnits: payload.budgetMinorUnits };
            if (payload.consumableCostMinorUnits !== undefined && payload.consumableCostMinorUnits !== null) {
                budget.consumableCostMinorUnits = payload.consumableCostMinorUnits;
                budget.sparePartCostMinorUnits = payload.sparePartCostMinorUnits;
                budget.laborTimeHours = payload.laborTimeHours;
                budget.laborRateMinorUnits = payload.laborRateMinorUnits;
            }
            return budget;
        }
        if (pt === 'SERVICE_COMMENT') return { payloadType: pt, comment: payload.comment ?? null };
        if (pt === 'SERVICE_STATE') return { payloadType: pt, serviceState: payload.serviceState };
        if (pt === 'PARTNER_WORKSHOP_CONNECTION') return { payloadType: pt, id: payload.id };
        if (pt === 'DAMAGE_TYPE') {
            const base = {
                payloadType: pt,
                area: payload.area,
                subArea: payload.subArea,
                partName: payload.partName,
                typeName: payload.typeName,
                isSecondaryWheel: payload.isSecondaryWheel
            };
            if (payload.isPA3Damage) {
                base.isPA3Damage = true;
                base.quantity = payload.quantity;
                if (payload.attribute1Name) {
                    base.attribute1Name = payload.attribute1Name;
                    base.attribute1Value = payload.attribute1Value;
                }
                if (payload.attribute2Name) {
                    base.attribute2Name = payload.attribute2Name;
                    base.attribute2Value = payload.attribute2Value;
                }
            }
            return base;
        }
        if (pt === 'DAMAGE_IMAGE') return { payloadType: pt, imageUrl: payload.imageUrl ?? null };
        if (pt === 'DAMAGE_CANCELLATION') {
            return { payloadType: pt, cancelComment: payload.cancelComment ?? null, cancelReason: payload.cancelReason ?? null };
        }
        if (pt === 'EC_COMMON_PAYLOAD') {
            return {
                payloadType: pt,
                sparePart: payload.sparePart ? {
                    status: spareStatus || payload.sparePart.status,
                    comment: payload.sparePart.comment ?? null,
                    lastStatusUpdateOn: null
                } : null
            };
        }
        return null;
    }

    const ENTRY_CHECK_QUERY = `query entryCheckPostAuthorization($refurbishmentId: UUID!) {
      entryCheckPostAuthorization(refurbishmentId: $refurbishmentId) {
        id state
        items {
          id entityId entityType verified
          values {
            id sourceType state
            originalValue {
              payloadType
              ... on DamageTypePayload { area subArea partName typeName isSecondaryWheel isImportedFromCat isPA3Damage attribute1Name attribute1Value attribute2Name attribute2Value quantity }
              ... on DamageImagePayload { imageUrl }
              ... on DamageServiceNamePayload { name type }
              ... on VehicleServiceNamePayload { name type }
              ... on ServiceBudgetPayload { budgetMinorUnits consumableCostMinorUnits sparePartCostMinorUnits laborRateMinorUnits laborTimeHours }
              ... on ServiceCommentPayload { comment }
              ... on DamageCancellationPayload { cancelReason cancelComment }
              ... on ServiceStatePayload { serviceState }
              ... on PartnerWorkshopConnectionPayload { id }
              ... on EcCommonPayload { sparePart { status comment lastStatusUpdatedOn } }
            }
            proposedValue {
              payloadType
              ... on DamageTypePayload { area subArea partName typeName isSecondaryWheel isImportedFromCat isPA3Damage attribute1Name attribute1Value attribute2Name attribute2Value quantity }
              ... on DamageImagePayload { imageUrl }
              ... on DamageServiceNamePayload { name type }
              ... on VehicleServiceNamePayload { name type }
              ... on ServiceBudgetPayload { budgetMinorUnits consumableCostMinorUnits sparePartCostMinorUnits laborRateMinorUnits laborTimeHours }
              ... on ServiceCommentPayload { comment }
              ... on DamageCancellationPayload { cancelReason cancelComment }
              ... on ServiceStatePayload { serviceState }
              ... on PartnerWorkshopConnectionPayload { id }
              ... on EcCommonPayload { sparePart { status comment lastStatusUpdatedOn } }
            }
          }
        }
      }
    }`;

    function refurbishmentIdFromUrl(url) {
        const m = String(url || '').match(/\/refurbishment\/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i);
        return m ? m[1] : '';
    }

    function isHandedSpareStatus(status) {
        const s = String(status || '').toUpperCase();
        return s === 'HANDED_OUT' || s === 'SPARE_PARTS_HANDED_OUT';
    }

    async function spareSectionsVerdict(refurbishmentId) {
        const data = await gqlOne(
            'refurbishmentDetails',
            `query refurbishmentDetails($id: UUID!) {
              refurbishmentDetails(id: $id) {
                id
                state
                spareParts {
                  serviceWorkSpareParts { status serviceName }
                  vehicleServiceSpareParts { status serviceName }
                }
              }
            }`,
            { id: refurbishmentId }
        );
        const details = data && data.refurbishmentDetails;
        const spare = details && details.spareParts;
        const groups = [
            { name: 'Maintenance parts', rows: (spare && spare.serviceWorkSpareParts) || [] },
            { name: 'Vehicle services parts', rows: (spare && spare.vehicleServiceSpareParts) || [] }
        ];
        let seen = 0;
        let open = 0;
        for (let g = 0; g < groups.length; g++) {
            const rows = groups[g].rows || [];
            for (let i = 0; i < rows.length; i++) {
                if (!rows[i]) continue;
                seen++;
                if (!isHandedSpareStatus(rows[i].status)) open++;
            }
        }
        return {
            state: String((details && details.state) || ''),
            seen: seen,
            handedOut: seen > 0 && open === 0
        };
    }

    async function handedOutViaApi(stockNumber, onStep, opts) {
        const sid = String(stockNumber || '').trim().toUpperCase();
        const workshops = await loadWorkshopIds();
        if (abortFlag) throw new Error('Abort');
        if (onStep) onStep('Suche Auftrag…');
        const ov = await gqlOne(
            'refurbishmentOverview',
            `query refurbishmentOverview($page: Int!, $pageSize: Int!, $filter: RefurbishmentOverviewFilter!) {
              refurbishmentOverview(page: $page, pageSize: $pageSize, filter: $filter) {
                items { id state refurbishmentSparePartStatus vehicle { stockNumber } }
              }
            }`,
            { page: 0, pageSize: 20, filter: { stockNumbers: [sid], mainWorkshopIds: workshops } }
        );
        const refItem = ov && ov.refurbishmentOverview && ov.refurbishmentOverview.items && ov.refurbishmentOverview.items[0];
        const urlId = opts && opts.refurbishmentId;
        if (!refItem && !urlId) return { status: 'error', message: 'Auftrag in Carol nicht gefunden' };
        const refId = (refItem && refItem.id) || urlId;
        let refState = String((refItem && refItem.state) || '');
        if (refState.indexOf('COMPLETED') !== -1 || !refItem) {
            if (onStep) onStep('Prüfe Spare Parts auf der Übersicht…');
            const verdict = await spareSectionsVerdict(refId);
            if (!refState) refState = verdict.state;
            if (refState.indexOf('COMPLETED') !== -1) {
                if (verdict.handedOut) {
                    return { status: 'ok', message: 'Completed, Maintenance parts und Vehicle services parts sind Herausgegeben' };
                }
                return { status: 'keep', message: 'Completed ohne Edit — Spare Parts nicht vollständig Herausgegeben, Farbe bleibt' };
            }
        }
        if (abortFlag) throw new Error('Abort');
        if (onStep) onStep('Lade Spare Parts…');
        const pa = await gqlOne('entryCheckPostAuthorization', ENTRY_CHECK_QUERY, { refurbishmentId: refId });
        const ec = pa && pa.entryCheckPostAuthorization;
        if (!ec || !ec.id) {
            if (onStep) onStep('Prüfe Spare Parts auf der Übersicht…');
            const verdict = await spareSectionsVerdict(refId);
            if (verdict.handedOut) {
                return { status: 'ok', message: 'Kein Edit-Button, Maintenance parts und Vehicle services parts sind Herausgegeben' };
            }
            return { status: 'keep', message: 'Kein Edit-Button — Farbe in Spalte F bleibt unverändert' };
        }

        const byEntity = new Map();
        const items = ec.items || [];
        for (let i = 0; i < items.length; i++) {
            const item = items[i];
            if (!item || !item.entityId) continue;
            if (item.entityType !== 'SERVICE' && item.entityType !== 'DAMAGE') continue;
            if (!byEntity.has(item.entityId)) byEntity.set(item.entityId, []);
            byEntity.get(item.entityId).push(item);
        }

        const updateBatches = [];
        const createBatches = [];
        byEntity.forEach((entityItems) => {
            const effMap = new Map();
            for (let i = 0; i < entityItems.length; i++) {
                const eff = getEffectiveVal(entityItems[i]);
                if (eff && eff.payload && eff.payload.payloadType) {
                    effMap.set(eff.payload.payloadType, { item: entityItems[i], payload: eff.payload, state: eff.state });
                }
            }
            const ecCommon = effMap.get('EC_COMMON_PAYLOAD');
            if (!ecCommon || !ecCommon.payload.sparePart) return;
            let declined = false;
            effMap.forEach((info) => {
                if (info.state === 'DECLINED' && info.payload.payloadType !== 'DAMAGE_CANCELLATION') declined = true;
            });
            if (declined) return;
            if (ecCommon.payload.sparePart.status === 'SPARE_PARTS_HANDED_OUT') return;

            let hasPostAuth = false;
            for (let i = 0; i < entityItems.length; i++) {
                const vals = entityItems[i].values || [];
                for (let v = 0; v < vals.length; v++) {
                    const pv = vals[v].proposedValue;
                    if (vals[v].sourceType === 'POST_AUTHORIZATION' && pv && pv.payloadType !== 'DAMAGE_CUSTOMER_DISPLAY') {
                        hasPostAuth = true;
                    }
                }
            }

            if (hasPostAuth) {
                const inputs = [];
                effMap.forEach((info, pt) => {
                    if (pt === 'DAMAGE_CUSTOMER_DISPLAY') return;
                    const paVal = (info.item.values || []).find(v => v.sourceType === 'POST_AUTHORIZATION');
                    const proposedValue = buildCleanProposed(info.payload, 'SPARE_PARTS_HANDED_OUT');
                    if (paVal && proposedValue) inputs.push({ id: paVal.id, value: { state: 'CONFIRMED', proposedValue: proposedValue } });
                });
                if (inputs.length) updateBatches.push(inputs);
            } else {
                const inputs = [];
                effMap.forEach((info, pt) => {
                    if (pt === 'DAMAGE_CUSTOMER_DISPLAY') return;
                    const proposedValue = buildCleanProposed(info.payload, 'SPARE_PARTS_HANDED_OUT');
                    if (proposedValue) {
                        inputs.push({
                            entryCheckItemId: info.item.id,
                            sourceType: 'POST_AUTHORIZATION',
                            state: 'CONFIRMED',
                            proposedValue: proposedValue
                        });
                    }
                });
                if (inputs.length) createBatches.push(inputs);
            }
        });

        const forceSubmit = !!(opts && opts.forceSubmit);
        if (!updateBatches.length && !createBatches.length && !forceSubmit) {
            return { status: 'ok', message: 'War bereits auf Herausgegeben, setze Spalte trotzdem auf grün' };
        }
        if (abortFlag) throw new Error('Abort');
        if ((updateBatches.length || createBatches.length) && onStep) onStep('Setze Teile auf Handed out…');

        if (updateBatches.length) {
            await gqlBatch(updateBatches.map(inputs => ({
                operationName: 'updateEntryCheckItemValues',
                variables: { inputs: inputs },
                query: 'mutation updateEntryCheckItemValues($inputs: [EntryCheckItemValueUpdateEntryInput!]!) { updateEntryCheckItemValues(inputs: $inputs) { id state } }'
            })));
        }
        if (createBatches.length) {
            await gqlBatch(createBatches.map(inputs => ({
                operationName: 'createEntryCheckItemValues',
                variables: { inputs: inputs },
                query: 'mutation createEntryCheckItemValues($inputs: [EntryCheckItemValueCreateInput!]!) { createEntryCheckItemValues(inputs: $inputs) { id state } }'
            })));
        }
        if (abortFlag) throw new Error('Abort');
        if (opts && opts.onBeforeSubmit) opts.onBeforeSubmit();
        if (onStep) onStep('Speichere (Submit)…');
        await gqlOne(
            'postAuthorizeEntryCheck',
            'mutation postAuthorizeEntryCheck($id: UUID!) { postAuthorizeEntryCheck(id: $id) { id state refurbishment { id state } } }',
            { id: ec.id }
        );
        return { status: 'ok', message: 'Herausgegeben' };
    }

    function updateJobInQueue(stockId, patch) {
        const q = readJson(KEY_QUEUE, []);
        let changed = false;
        for (let i = 0; i < q.length; i++) {
            if (q[i] && q[i].stockId === stockId) {
                Object.assign(q[i], patch);
                changed = true;
            }
        }
        if (changed) writeJson(KEY_QUEUE, q);
    }

    function removeJobFromQueue(stockId) {
        const q = readJson(KEY_QUEUE, []);
        const next = q.filter(j => j && j.stockId !== stockId);
        writeJson(KEY_QUEUE, next);
        return next;
    }

    async function finishJobAndAdvance(job, status, message) {
        const sid = job.stockId;
        const remaining = removeJobFromQueue(sid);
        pushResultEvent(sid, status, message, job);
        playTone(status === 'ok');

        if (remaining.length > 0 && !abortFlag) {
            const nextJob = remaining[0];
            currentSid = nextJob.stockId;
            setHudStatus(`offen: ${remaining.length} · ${nextJob.stockId}`, 'Nächster Auftrag…', '#00ffcc', true);
            return 'done';
        }

        currentSid = '';
        if (status === 'ok') {
            setHudStatus(`FERTIG · ${sid}`, `✓ ${message}`, '#00ff00', false);
        } else if (status === 'keep') {
            setHudStatus(`FERTIG · ${sid}`, message, '#ffaa00', false);
        } else {
            setHudStatus(`FEHLER · ${sid}`, `✗ ${message}`, '#f85149', false);
        }
        return 'all_done';
    }

    function isLoginStopError(err) {
        const msg = String((err && err.message) || err || '');
        return msg.indexOf('neu anmelden') !== -1 || msg.indexOf('Nicht in Carol eingeloggt') !== -1;
    }

    async function processCurrentJob(job, indexLabel) {
        const sid = job.stockId;
        currentSid = sid;
        const keepSubmit = job.phase === 'needs_submit';
        updateJobInQueue(sid, { status: 'running', phase: keepSubmit ? 'needs_submit' : 'api' });
        try {
            const result = await handedOutViaApi(sid, (step) => {
                const color = step.indexOf('Handed out') !== -1 ? '#ffff00' : '#00ffcc';
                setHudStatus(`${indexLabel} · ${sid}`, step, color, true);
            }, {
                forceSubmit: keepSubmit,
                refurbishmentId: refurbishmentIdFromUrl(job.carolUrl),
                onBeforeSubmit: () => updateJobInQueue(sid, { phase: 'needs_submit' })
            });
            return await finishJobAndAdvance(job, result.status, result.message, indexLabel);
        } catch (err) {
            if (err && (err.message === 'Abort' || err.message === 'Reload')) throw err;
            if (isLoginStopError(err)) throw err;
            return await finishJobAndAdvance(job, 'error', String((err && err.message) || err), indexLabel);
        }
    }

    async function checkAndRunQueue() {
        if (workerBusy || isNavigating) return;
        const abortObj = readJson(KEY_ABORT, null);
        if (abortObj && Number(abortObj.ts || 0) > (Date.now() - 4000)) {
            writeJson(KEY_QUEUE, []);
            return;
        }

        const currentQ = readJson(KEY_QUEUE, []);
        if (!Array.isArray(currentQ) || !currentQ.length) return;

        if (authPauseReason()) {
            const pausedAuth = readAuth();
            if (pausedAuth && tokenStillValid(pausedAuth.access_token, 0)) {
                clearAuthPause();
            } else {
                return;
            }
        }

        try { sessionStorage.setItem('dpl_tl_worker_tab', '1'); } catch (e) {}
        workerBusy = true;
        abortFlag = false;

        try {
            const nextJob = currentQ[0];
            if (!nextJob || !nextJob.stockId) {
                writeJson(KEY_QUEUE, currentQ.slice(1));
                workerBusy = false;
                return;
            }

            const result = await processCurrentJob(nextJob, `offen: ${currentQ.length}`);
            if (result === 'navigating') {
                return;
            }
        } catch (err) {
            if (err && err.message === 'Reload') {
                return;
            }
            if (err && err.message === 'Abort') {
                currentSid = '';
                setHudStatus('GESTOPPT', 'Durchlauf abgebrochen', '#f85149', false);
            } else if (isLoginStopError(err)) {
                const loginMsg = String((err && err.message) || 'Carol Login abgelaufen, bitte neu anmelden');
                setAuthPause(loginMsg);
                setHudStatus('LOGIN', loginMsg, '#f85149', false);
            } else {
                setHudStatus('FEHLER', String((err && err.message) || err), '#f85149', false);
            }
        } finally {
            if (!isNavigating) {
                workerBusy = false;
            }
        }
    }

    window.addEventListener('message', (ev) => {
        const d = ev && ev.data;
        if (!d || d.source !== 'dpl-tagesliste-hud') return;
        try { sessionStorage.setItem('dpl_tl_worker_tab', '1'); } catch (e) {}

        if (d.cmd === 'enqueue_jobs' && d.payload && Array.isArray(d.payload.jobs)) {
            const incoming = d.payload.jobs;
            const bySid = {};
            const freshQueue = [];
            for (let j = 0; j < incoming.length; j++) {
                const inc = incoming[j];
                if (!inc || !inc.stockId) continue;
                const sid = String(inc.stockId).trim().toUpperCase();
                if (!bySid[sid]) {
                    const job = {
                        stockId: sid,
                        dateStr: inc.dateStr || '',
                        carolUrl: inc.carolUrl || ('https://carol.autohero.com/en-GB/refurbishment?rsv=' + encodeURIComponent(sid)),
                        model: inc.model || '',
                        rows: inc.rows || [],
                        row: inc.row || 0,
                        status: 'pending',
                        phase: 'init',
                        addedAt: Date.now()
                    };
                    bySid[sid] = job;
                    freshQueue.push(job);
                }
            }
            writeJson(KEY_ABORT, { ts: 0 });
            clearAuthPause();
            writeJson(KEY_QUEUE, freshQueue);
            checkAndRunQueue();
        }

        if (d.cmd === 'stop_all') {
            abortFlag = true;
            workerBusy = false;
            isNavigating = false;
            writeJson(KEY_QUEUE, []);
            setHudStatus('GESTOPPT', 'Durchlauf durch HUD gestoppt', '#f85149', false);
        }
    });

    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && isWorkerTab()) {
            abortFlag = true;
            workerBusy = false;
            isNavigating = false;
            writeJson(KEY_QUEUE, []);
            setHudStatus('GESTOPPT', 'Durchlauf mit ESC abgebrochen', '#f85149', false);
        }
    });

    if (typeof GM_addValueChangeListener === 'function') {
        GM_addValueChangeListener(KEY_QUEUE, () => {
            const qNow = readJson(KEY_QUEUE, []);
            const fresh = qNow[0] && qNow[0].addedAt && (Date.now() - Number(qNow[0].addedAt)) < 8000;
            if (fresh) clearAuthPause();
            if (!workerBusy && !isNavigating) checkAndRunQueue();
        });
    }

    setInterval(() => {
        if (isWorkerTab()) {
            ensureHud();
            sendHeartbeat();
        }
        if (!workerBusy && !isNavigating) {
            const q = readJson(KEY_QUEUE, []);
            if (Array.isArray(q) && q.length > 0) {
                checkAndRunQueue();
            }
        }
    }, 1000);

    setTimeout(() => {
        if (isWorkerTab()) {
            ensureHud();
            sendHeartbeat();
        }
        checkAndRunQueue();
    }, 300);
})();
