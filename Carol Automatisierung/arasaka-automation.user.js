// ==UserScript==
// @name Carol-Automation
// @namespace http://tampermonkey.net/
// @version 5.1
// @description ARASAKA v5.1 - Spare Parts per API auf Delivered, Werkstattauftrag wird weiter von der offenen Seite gedruckt
// @author ARASAKA
// @match        *://carol.autohero.com/*
// @updateURL https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/Carol%20Automatisierung/arasaka-automation.user.js
// @downloadURL https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/Carol%20Automatisierung/arasaka-automation.user.js
// @grant GM_setValue
// @grant GM_getValue
// @grant GM_registerMenuCommand
// @grant GM_xmlhttpRequest
// @grant window.close
// @grant window.focus
// @connect localhost
// ==/UserScript==

(function () {
    'use strict';

    const launchedAsAuto = window.location.href.includes('arasaka_auto=1');
    let abortMission = false;
    let hudElement = null;
    let isLocked = false;
    const speedMultiplier = GM_getValue('arasaka_speed', 1);
    const TOTAL_STEPS = 4;
    let hudLog = [];
    let currentStep = 0;
    let startTime = null;
    let timerInterval = null;

    GM_registerMenuCommand(
        `⚙️ ARASAKA Speed: ${speedMultiplier === 1 ? 'NORMAL' : speedMultiplier === 1.5 ? 'LANGSAM' : 'SEHR LANGSAM'}`,
        () => {
            const next = speedMultiplier === 1 ? 1.5 : speedMultiplier === 1.5 ? 2 : 1;
            const name = next === 1 ? 'NORMAL' : next === 1.5 ? 'LANGSAM' : 'SEHR LANGSAM';
            GM_setValue('arasaka_speed', next);
            alert(`ARASAKA Geschwindigkeit → ${name}\nSeite neu laden, damit die Änderung aktiv wird.`);
        }
    );

    function injectStyles() {
        if (document.getElementById('arasaka-styles')) return;
        const style = document.createElement('style');
        style.id = 'arasaka-styles';
        style.innerHTML = `
            @keyframes arasaka-fadein {
                from { opacity: 0; transform: translateY(16px); }
                to { opacity: 1; transform: translateY(0); }
            }
            @keyframes arasaka-pulse {
                0%, 100% { box-shadow: 0 0 14px var(--a-col, #00ffcc88); }
                50% { box-shadow: 0 0 28px var(--a-col, #00ffcccc); }
            }
            @keyframes arasaka-blink {
                0%, 100% { opacity: 1; } 50% { opacity: 0; }
            }
            @keyframes arasaka-scanline {
                from { background-position: 0 0; }
                to { background-position: 0 100%; }
            }
            #arasaka-hud {
                position: fixed;
                bottom: 28px;
                right: 28px;
                width: 360px;
                background: rgba(5, 8, 12, 0.96);
                color: #00ffcc;
                border: 1.5px solid #00ffcc;
                border-radius: 6px;
                font-family: 'Courier New', monospace;
                font-size: 14px;
                z-index: 999999;
                pointer-events: none;
                animation: arasaka-fadein 0.3s ease, arasaka-pulse 2.5s ease-in-out infinite;
                --a-col: #00ffcc88;
                overflow: hidden;
                user-select: none;
            }
            #arasaka-hud::after {
                content: '';
                position: absolute;
                inset: 0;
                background: repeating-linear-gradient(
                    0deg,
                    transparent,
                    transparent 3px,
                    rgba(0,255,204,0.02) 3px,
                    rgba(0,255,204,0.02) 4px
                );
                pointer-events: none;
            }
            .a-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 12px 16px 10px;
                border-bottom: 1px solid rgba(0,255,204,0.25);
                font-size: 14px;
                letter-spacing: 2px;
                font-weight: bold;
            }
            .a-timer { font-size: 13px; opacity: 0.7; letter-spacing: 1px; }
            .a-progress-wrap {
                padding: 10px 16px 6px;
            }
            .a-step-label {
                display: flex;
                justify-content: space-between;
                font-size: 12px;
                opacity: 0.6;
                margin-bottom: 4px;
                letter-spacing: 1px;
            }
            .a-bar-bg {
                height: 5px;
                background: rgba(0,255,204,0.1);
                border-radius: 3px;
                overflow: hidden;
            }
            .a-bar-fill {
                height: 100%;
                background: #00ffcc;
                border-radius: 3px;
                transition: width 0.4s ease;
            }
            .a-log {
                padding: 8px 16px 0;
                min-height: 54px;
            }
            .a-log-entry {
                font-size: 12px;
                opacity: 0.55;
                margin-bottom: 4px;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            }
            .a-log-entry::before { content: '✓ '; }
            .a-current {
                padding: 10px 16px 8px;
                font-size: 15px;
                border-top: 1px solid rgba(0,255,204,0.15);
                margin-top: 6px;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            }
            .a-cursor { animation: arasaka-blink 0.9s infinite; }
            .a-footer {
                padding: 6px 16px 10px;
                font-size: 12px;
                opacity: 0.35;
                letter-spacing: 1px;
            }
            #arasaka-confirm {
                position: fixed;
                inset: 0;
                background: rgba(0,0,0,0.7);
                z-index: 1000000;
                display: flex;
                align-items: center;
                justify-content: center;
                animation: arasaka-fadein 0.2s ease;
                pointer-events: all;
            }
            .a-dialog {
                background: rgba(5, 8, 12, 0.98);
                border: 1.5px solid #ffaa00;
                border-radius: 6px;
                padding: 24px 28px;
                font-family: 'Courier New', monospace;
                color: #ffaa00;
                max-width: 420px;
                width: 90%;
                box-shadow: 0 0 30px rgba(255,170,0,0.4);
            }
            .a-dialog-title {
                font-size: 13px;
                font-weight: bold;
                letter-spacing: 2px;
                margin-bottom: 12px;
            }
            .a-dialog-msg {
                font-size: 12px;
                line-height: 1.6;
                opacity: 0.85;
                margin-bottom: 20px;
                white-space: pre-wrap;
            }
            .a-dialog-btns { display: flex; gap: 10px; justify-content: flex-end; }
            .a-btn {
                padding: 7px 18px;
                font-family: 'Courier New', monospace;
                font-size: 12px;
                font-weight: bold;
                letter-spacing: 1px;
                border-radius: 3px;
                cursor: pointer;
                border: 1.5px solid currentColor;
                background: transparent;
                transition: background 0.15s;
            }
            .a-btn-ok { color: #00ffcc; }
            .a-btn-ok:hover { background: rgba(0,255,204,0.15); }
            .a-btn-cancel { color: #ff4444; }
            .a-btn-cancel:hover { background: rgba(255,68,68,0.15); }
        `;
        document.head.appendChild(style);
    }

    function createHUD() {
        if (hudElement) return;
        injectStyles();
        hudElement = document.createElement('div');
        hudElement.id = 'arasaka-hud';
        hudElement.innerHTML = buildHUDHTML('Initialisiere...', '#00ffcc', currentStep);
        document.body.appendChild(hudElement);

        startTime = Date.now();
        timerInterval = setInterval(() => {
            const elapsed = Math.floor((Date.now() - startTime) / 1000);
            const m = String(Math.floor(elapsed / 60)).padStart(2, '0');
            const s = String(elapsed % 60).padStart(2, '0');
            const el = document.querySelector('.a-timer');
            if (el) el.textContent = `${m}:${s}`;
        }, 1000);
    }

    function buildHUDHTML(action, color, step) {
        const displayStep = Math.min(step, TOTAL_STEPS);
        const pct = Math.min(100, Math.round((displayStep / TOTAL_STEPS) * 100));
        const logHTML = hudLog.slice(-3).map(t =>
            `<div class="a-log-entry">${t}</div>`
        ).join('');

        return `
            <div class="a-header" style="color:${color}; border-color: ${color}40;">
                <span>ARASAKA SYS //</span>
                <span class="a-timer">00:00</span>
            </div>
            <div class="a-progress-wrap">
                <div class="a-step-label" style="color:${color}">
                    <span>FORTSCHRITT</span>
                    <span>SCHRITT ${displayStep}/${TOTAL_STEPS}</span>
                </div>
                <div class="a-bar-bg">
                    <div class="a-bar-fill" style="width:${pct}%; background:${color};"></div>
                </div>
            </div>
            <div class="a-log" style="color:${color}">${logHTML}</div>
            <div class="a-current" style="color:${color}">
                ► ${action}<span class="a-cursor">_</span>
            </div>
            <div class="a-footer">[ALT+Y] START  ·  [ESC] ABBRUCH</div>
        `;
    }

    function updateHUD(text, color = '#00ffcc', advanceStep = false) {
        if (advanceStep && hudLog[hudLog.length - 1] !== text) {
            const prev = document.querySelector('.a-current');
            if (prev) {
                const prevText = prev.textContent.replace('► ', '').replace('_', '').trim();
                if (prevText && prevText !== text) hudLog.push(prevText);
            }
            currentStep++;
        }

        if (!hudElement) createHUD();

        hudElement.style.borderColor = color;
        hudElement.style.setProperty('--a-col', color + '88');

        const elapsed = document.querySelector('.a-timer')?.textContent || '00:00';

        hudElement.innerHTML = buildHUDHTML(text, color, currentStep);
        hudElement.style.borderColor = color;

        const timerEl = hudElement.querySelector('.a-timer');
        if (timerEl) timerEl.textContent = elapsed;
    }

    function removeHUD() {
        clearInterval(timerInterval);
        timerInterval = null;
        if (hudElement) {
            hudElement.style.animation = 'none';
            hudElement.style.transition = 'opacity 0.4s ease';
            hudElement.style.opacity = '0';
            setTimeout(() => { hudElement?.remove(); hudElement = null; }, 400);
        }
        hudLog = [];
        currentStep = 0;
        startTime = null;
    }

    function arasakaConfirm(message) {
        return new Promise((resolve) => {
            injectStyles();
            const overlay = document.createElement('div');
            overlay.id = 'arasaka-confirm';
            overlay.innerHTML = `
                <div class="a-dialog">
                    <div class="a-dialog-title">⚠ ARASAKA WARNUNG</div>
                    <div class="a-dialog-msg">${message}</div>
                    <div class="a-dialog-btns">
                        <button class="a-btn a-btn-cancel" id="a-btn-no">ABBRECHEN</button>
                        <button class="a-btn a-btn-ok"    id="a-btn-yes">RETTUNG STARTEN</button>
                    </div>
                </div>
            `;
            document.body.appendChild(overlay);

            document.getElementById('a-btn-yes').onclick = () => { overlay.remove(); resolve(true); };
            document.getElementById('a-btn-no').onclick  = () => { overlay.remove(); resolve(false); };
        });
    }

    function playDing(type = 'success') {
        try {
            const ctx = new (window.AudioContext || window.webkitAudioContext)();
            const osc  = ctx.createOscillator();
            const gain = ctx.createGain();
            osc.connect(gain);
            gain.connect(ctx.destination);
            if (type === 'success') {
                osc.type = 'sine';
                osc.frequency.setValueAtTime(880, ctx.currentTime);
                osc.frequency.exponentialRampToValueAtTime(1760, ctx.currentTime + 0.1);
                gain.gain.setValueAtTime(0.2, ctx.currentTime);
                gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.5);
                osc.start(); osc.stop(ctx.currentTime + 0.5);
            } else {
                osc.type = 'sawtooth';
                osc.frequency.setValueAtTime(150, ctx.currentTime);
                gain.gain.setValueAtTime(0.3, ctx.currentTime);
                gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.4);
                osc.start(); osc.stop(ctx.currentTime + 0.4);
            }
        } catch (e) {}
    }

    async function sleep(ms) {
        let t = 0;
        const target = ms * speedMultiplier;
        while (t < target) {
            if (abortMission) throw new Error('Abort');
            await new Promise(r => setTimeout(r, 100));
            t += 100;
        }
    }

    function getDeepestElementsByText(text) {
        const results = [];
        for (const el of document.querySelectorAll('*')) {
            if (['SCRIPT', 'STYLE', 'HTML', 'HEAD', 'BODY'].includes(el.tagName)) continue;
            if (el.textContent?.includes(text)) {
                const childHasText = Array.from(el.children).some(c => c.textContent?.includes(text));
                if (!childHasText) results.push(el);
            }
        }
        return results;
    }

    async function waitForElement(text, timeout = 10000, pickLast = false) {
        let passed = 0;
        const target = timeout * speedMultiplier;
        while (passed < target) {
            if (abortMission) throw new Error('Abort');
            const els = getDeepestElementsByText(text);
            if (els.length > 0) return pickLast ? els[els.length - 1] : els[0];
            await sleep(500);
            passed += 500;
        }
        return null;
    }

    const AUTH_STORAGE_KEY = 'authData';
    const TOKEN_URL = '/api/v1/auth/oauth/token';
    const DELIVERED_STATUS = 'SPARE_PARTS_DELIVERED';
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
            if (!cur || !cur.refresh_token) throw new Error('Carol Login abgelaufen, bitte neu anmelden');
            const body = new URLSearchParams();
            body.set('client_id', 'refurbishment');
            body.set('grant_type', 'refresh_token');
            body.set('refresh_token', cur.refresh_token);
            const headers = { 'Content-Type': 'application/x-www-form-urlencoded', 'Accept': 'application/json' };
            if (cur.access_token) headers.authorization = 'Bearer ' + cur.access_token;
            const res = await fetch(TOKEN_URL, { method: 'POST', headers: headers, body: body.toString() });
            let data = null;
            try { data = await res.json(); } catch (e) { data = null; }
            const tokenObj = (data && data.access_token) ? data : (data && data.result && data.result.access_token ? data.result : null);
            if (!res.ok || !tokenObj) throw new Error('Carol Token konnte nicht erneuert werden, bitte neu anmelden');
            const next = Object.assign({}, cur, tokenObj);
            if (!next.refresh_token) next.refresh_token = cur.refresh_token;
            writeAuth(next);
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
        if (!auth || !auth.access_token) throw new Error('Nicht in Carol eingeloggt');
        if (tokenStillValid(auth.access_token, 90000)) return auth.access_token;
        const next = await refreshAuth();
        return next.access_token;
    }

    function assertGql(arr) {
        if (!Array.isArray(arr)) throw new Error('Carol API Antwort ungültig');
        const msgs = [];
        for (let i = 0; i < arr.length; i++) {
            const row = arr[i];
            if (row && Array.isArray(row.errors) && row.errors.length) {
                for (let j = 0; j < row.errors.length; j++) msgs.push((row.errors[j] && row.errors[j].message) || 'GraphQL Fehler');
            }
        }
        if (msgs.length) throw new Error(msgs.join('; '));
        return arr;
    }

    async function gqlBatch(ops, retried) {
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
        if (res.status === 401 && !retried) {
            await refreshAuth();
            return gqlBatch(ops, true);
        }
        if (!res.ok) throw new Error('Carol API HTTP ' + res.status);
        return assertGql(await res.json());
    }

    async function gqlOne(operationName, query, variables) {
        const arr = await gqlBatch([{ operationName: operationName, variables: variables, query: query }], false);
        return arr[0] && arr[0].data;
    }

    function currentRefurbishmentId() {
        const m = location.pathname.match(/\/refurbishment\/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i);
        return m ? m[1] : '';
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
            const branchNum = Number(conns[i] && conns[i].mainWorkshop && conns[i].mainWorkshop.branchId);
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
        if (pt === 'VEHICLE_SERVICE_NAME' || pt === 'DAMAGE_SERVICE_NAME') return { payloadType: pt, name: payload.name, type: payload.type };
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
            const base = { payloadType: pt, area: payload.area, subArea: payload.subArea, partName: payload.partName, typeName: payload.typeName, isSecondaryWheel: payload.isSecondaryWheel };
            if (payload.isPA3Damage) {
                base.isPA3Damage = true;
                base.quantity = payload.quantity;
                if (payload.attribute1Name) { base.attribute1Name = payload.attribute1Name; base.attribute1Value = payload.attribute1Value; }
                if (payload.attribute2Name) { base.attribute2Name = payload.attribute2Name; base.attribute2Value = payload.attribute2Value; }
            }
            return base;
        }
        if (pt === 'DAMAGE_IMAGE') return { payloadType: pt, imageUrl: payload.imageUrl ?? null };
        if (pt === 'DAMAGE_CANCELLATION') return { payloadType: pt, cancelComment: payload.cancelComment ?? null, cancelReason: payload.cancelReason ?? null };
        if (pt === 'EC_COMMON_PAYLOAD') {
            return { payloadType: pt, sparePart: payload.sparePart ? { status: spareStatus || payload.sparePart.status, comment: payload.sparePart.comment ?? null, lastStatusUpdateOn: null } : null };
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

    async function resolveRefurbishmentId() {
        const direct = currentRefurbishmentId();
        if (direct) return direct;
        const sid = (new URLSearchParams(location.search).get('rsv') || '').trim().toUpperCase();
        if (!sid) throw new Error('Carol Auftrag auf dieser Seite nicht erkannt');
        const workshops = await loadWorkshopIds();
        const ov = await gqlOne(
            'refurbishmentOverview',
            `query refurbishmentOverview($page: Int!, $pageSize: Int!, $filter: RefurbishmentOverviewFilter!) {
              refurbishmentOverview(page: $page, pageSize: $pageSize, filter: $filter) { items { id } }
            }`,
            { page: 0, pageSize: 5, filter: { stockNumbers: [sid], mainWorkshopIds: workshops } }
        );
        const item = ov && ov.refurbishmentOverview && ov.refurbishmentOverview.items && ov.refurbishmentOverview.items[0];
        if (!item || !item.id) throw new Error('Auftrag in Carol nicht gefunden');
        return item.id;
    }

    async function deliveredViaApi(onStep) {
        if (abortMission) throw new Error('Abort');
        if (onStep) onStep('Lade Spare Parts…');
        const refId = await resolveRefurbishmentId();
        const pa = await gqlOne('entryCheckPostAuthorization', ENTRY_CHECK_QUERY, { refurbishmentId: refId });
        const ec = pa && pa.entryCheckPostAuthorization;
        if (!ec || !ec.id) return { status: 'error', message: 'Post Authorization nicht verfügbar' };
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
        let spareCount = 0;
        byEntity.forEach((entityItems) => {
            const effMap = new Map();
            for (let i = 0; i < entityItems.length; i++) {
                const eff = getEffectiveVal(entityItems[i]);
                if (eff && eff.payload && eff.payload.payloadType) effMap.set(eff.payload.payloadType, { item: entityItems[i], payload: eff.payload, state: eff.state });
            }
            const ecCommon = effMap.get('EC_COMMON_PAYLOAD');
            if (!ecCommon || !ecCommon.payload.sparePart) return;
            let declined = false;
            effMap.forEach((info) => {
                if (info.state === 'DECLINED' && info.payload.payloadType !== 'DAMAGE_CANCELLATION') declined = true;
            });
            if (declined) return;
            spareCount++;
            if (ecCommon.payload.sparePart.status === DELIVERED_STATUS) return;
            let hasPostAuth = false;
            for (let i = 0; i < entityItems.length; i++) {
                const vals = entityItems[i].values || [];
                for (let v = 0; v < vals.length; v++) {
                    const pv = vals[v].proposedValue;
                    if (vals[v].sourceType === 'POST_AUTHORIZATION' && pv && pv.payloadType !== 'DAMAGE_CUSTOMER_DISPLAY') hasPostAuth = true;
                }
            }
            if (hasPostAuth) {
                const inputs = [];
                effMap.forEach((info, pt) => {
                    if (pt === 'DAMAGE_CUSTOMER_DISPLAY') return;
                    const paVal = (info.item.values || []).find(v => v.sourceType === 'POST_AUTHORIZATION');
                    const proposedValue = buildCleanProposed(info.payload, DELIVERED_STATUS);
                    if (paVal && proposedValue) inputs.push({ id: paVal.id, value: { state: 'CONFIRMED', proposedValue: proposedValue } });
                });
                if (inputs.length) updateBatches.push(inputs);
            } else {
                const inputs = [];
                effMap.forEach((info, pt) => {
                    if (pt === 'DAMAGE_CUSTOMER_DISPLAY') return;
                    const proposedValue = buildCleanProposed(info.payload, DELIVERED_STATUS);
                    if (proposedValue) inputs.push({ entryCheckItemId: info.item.id, sourceType: 'POST_AUTHORIZATION', state: 'CONFIRMED', proposedValue: proposedValue });
                });
                if (inputs.length) createBatches.push(inputs);
            }
        });
        if (!spareCount) return { status: 'error', message: 'Keine Spare Parts auf diesem Auftrag' };
        if (abortMission) throw new Error('Abort');
        if (updateBatches.length || createBatches.length) {
            if (onStep) onStep('Setze Teile auf Delivered…');
            if (updateBatches.length) {
                await gqlBatch(updateBatches.map(inputs => ({
                    operationName: 'updateEntryCheckItemValues',
                    variables: { inputs: inputs },
                    query: 'mutation updateEntryCheckItemValues($inputs: [EntryCheckItemValueUpdateEntryInput!]!) { updateEntryCheckItemValues(inputs: $inputs) { id state } }'
                })), false);
            }
            if (createBatches.length) {
                await gqlBatch(createBatches.map(inputs => ({
                    operationName: 'createEntryCheckItemValues',
                    variables: { inputs: inputs },
                    query: 'mutation createEntryCheckItemValues($inputs: [EntryCheckItemValueCreateInput!]!) { createEntryCheckItemValues(inputs: $inputs) { id state } }'
                })), false);
            }
        }
        if (abortMission) throw new Error('Abort');
        if (onStep) onStep('Speichere (Submit)…');
        await gqlOne(
            'postAuthorizeEntryCheck',
            'mutation postAuthorizeEntryCheck($id: UUID!) { postAuthorizeEntryCheck(id: $id) { id state refurbishment { id state } } }',
            { id: ec.id }
        );
        return { status: 'ok', message: 'Delivered' };
    }

    function forceClick(el) {
        if (!el) return;
        try {
            const orig = { outline: el.style.outline, shadow: el.style.boxShadow, trans: el.style.transition };
            el.style.transition = 'all 0.1s';
            el.style.outline = '2px solid #ff0000';
            el.style.boxShadow = '0 0 15px #ff0000';
            setTimeout(() => {
                if (el) { el.style.outline = orig.outline; el.style.boxShadow = orig.shadow; el.style.transition = orig.trans; }
            }, 400);
        } catch (e) {}
        el.click();
        if (el.parentElement) setTimeout(() => el.parentElement.click(), 50);
    }

    function extractPdfUrl(el) {
        if (el.tagName === 'A' && el.href) return el.href;
        const link = el.closest('a');
        if (link && link.href) return link.href;
        if (el.href) return el.href;
        const child = el.querySelector('a[href]');
        if (child) return child.href;
        return null;
    }

    function findWerkstattPdf() {
        for (const el of getDeepestElementsByText('Werkstattauftrag')) {
            const row = el.closest('tr') || el.closest('[role="row"]');
            if (!row) continue;
            const link = row.querySelector('a[href]');
            if (link?.href) return { url: link.href, element: link };
        }

        for (const link of document.querySelectorAll('a[href]')) {
            const row = link.closest('tr') || link.closest('[role="row"]');
            if (row?.textContent?.includes('Werkstattauftrag') && link.href.toLowerCase().includes('.pdf')) {
                return { url: link.href, element: link };
            }
        }

        for (const link of document.querySelectorAll('a[href*=".pdf"], a[href*="pdf"]')) {
            const row = link.closest('tr') || link.closest('[role="row"]');
            if (row?.textContent?.toLowerCase().includes('werkstatt')) {
                return { url: link.href, element: link };
            }
        }

        return null;
    }

    async function waitForWerkstattPdf(timeout = 2000) {
        let passed = 0;
        while (passed < timeout) {
            if (abortMission) throw new Error('Abort');
            const found = findWerkstattPdf();
            if (found) return found;
            await sleep(100);
            passed += 100;
        }
        return null;
    }

    function isBridgeUp() {
        return new Promise((resolve) => {
            GM_xmlhttpRequest({
                method: 'GET',
                url: 'http://localhost:9150/status',
                timeout: 1500,
                onload: () => resolve(true),
                onerror: () => resolve(false),
                ontimeout: () => resolve(false)
            });
        });
    }

    function sendToBridge(url) {
        return new Promise((resolve) => {
            GM_xmlhttpRequest({
                method: 'POST',
                url: 'http://localhost:9150/print',
                headers: { 'Content-Type': 'application/json' },
                data: JSON.stringify({ url: url, copies: 1 }),
                timeout: 60000,
                onload: r => {
                    try {
                        const res = JSON.parse(r.responseText);
                        resolve(res.success ? { ok: true, printer: res.printer } : { ok: false, error: res.message });
                    } catch { resolve({ ok: false, error: 'Parse error' }); }
                },
                onerror: () => resolve({ ok: false, error: 'Bridge nicht erreichbar' }),
                ontimeout: () => resolve({ ok: false, error: 'Timeout' })
            });
        });
    }

    function queuePrint(url) {
        sendToBridge(url).then((result) => {
            if (!result.ok && hudElement) {
                const cur = document.querySelector('.a-current');
                if (cur && !cur.textContent.includes('ABGESCHLOSSEN') && !cur.textContent.includes('ABBRUCH')) {
                    updateHUD(`Druck-Hinweis: ${result.error}`, '#ffaa00');
                }
            }
        });
    }

    async function sucheUndOeffnePdf({ removeHudOnComplete = true, fireAndForget = false } = {}) {
        updateHUD('Suche Werkstattauftrag...', '#ffff00', true);

        const pdf = await waitForWerkstattPdf(2000);

        if (pdf) {
            const pdfUrl = pdf.url || extractPdfUrl(pdf.element);
            if (pdfUrl) {
                if (fireAndForget) {
                    updateHUD('Druck gestartet...', '#00ff00', true);
                    queuePrint(pdfUrl);
                    playDing('success');
                } else {
                    const bridgeOk = await isBridgeUp();
                    if (bridgeOk) {
                        const result = await sendToBridge(pdfUrl);
                        if (result.ok) {
                            updateHUD(`PDF GEDRUCKT ─ ${result.printer}`, '#00ff00', true);
                            playDing('success');
                        } else {
                            updateHUD(`Druck fehlgeschlagen: ${result.error}`, '#ff4444', true);
                            playDing('error');
                            forceClick(pdf.element);
                            updateHUD('PDF GEÖFFNET ─ Bereit für STRG+P (Fallback)', '#ffaa00', true);
                        }
                    } else {
                        forceClick(pdf.element);
                        updateHUD('PDF GEÖFFNET ─ Bridge offline (STRG+P)', '#ffaa00', true);
                        playDing('error');
                    }
                }
            } else {
                forceClick(pdf.element);
                updateHUD('PDF GEÖFFNET ─ Bereit für STRG+P (kein Link)', '#ffaa00', true);
                playDing('success');
            }

            if (removeHudOnComplete) {
                setTimeout(removeHUD, 6000);
            }
        } else {
            updateHUD('Kein PDF/Auftrag gefunden!', '#ff4444', true);
            playDing('error');
            if (removeHudOnComplete) {
                setTimeout(removeHUD, 5000);
            }
        }
    }

    function postDoneToWindowTree(win, depth) {
        if (!win || depth > 3) return;
        try { win.postMessage({ type: 'arasaka_carol_done' }, '*'); } catch (e) {}
        try {
            for (let i = 0; i < win.frames.length; i++) {
                postDoneToWindowTree(win.frames[i], depth + 1);
            }
        } catch (e) {}
    }

    function notifyWmsDone() {
        let sent = 0;
        const timer = setInterval(() => {
            sent++;
            try {
                if (window.opener && !window.opener.closed) {
                    let root = window.opener;
                    try { if (root.top) root = root.top; } catch (e) {}
                    postDoneToWindowTree(root, 0);
                }
            } catch (e) {}
            if (sent >= 8) clearInterval(timer);
        }, 300);
    }

    function closeCarolTabLikeCtrlW() {
        try { window.close(); } catch (e) {}
        setTimeout(() => {
            try {
                const self = window.open('', '_self');
                if (self) self.close();
            } catch (e) {}
        }, 200);
        setTimeout(() => {
            try { window.close(); } catch (e) {}
        }, 500);
    }

    function showCloseTabPrompt() {
        const old = document.getElementById('arasaka-close-prompt');
        if (old) old.remove();

        const box = document.createElement('div');
        box.id = 'arasaka-close-prompt';
        box.style.cssText = [
            'position:fixed', 'top:50%', 'left:50%', 'transform:translate(-50%,-50%)',
            'background:rgba(10,10,10,0.96)', 'border:4px solid #00ff00',
            'box-shadow:0 0 40px rgba(0,255,0,0.7), inset 0 0 20px rgba(0,255,0,0.2)',
            'padding:50px 90px', 'border-radius:18px', 'z-index:9999999',
            'color:#00ff00', 'font-family:monospace', 'text-align:center', 'cursor:pointer'
        ].join(';');
        box.innerHTML = `
            <div style="font-size:32px;font-weight:bold;text-shadow:0 0 12px #00ff00;">AUFTRAG ABGESCHLOSSEN</div>
            <div style="font-size:18px;margin-top:24px;opacity:0.9;">ENTER drücken oder klicken</div>
            <div style="font-size:15px;margin-top:6px;opacity:0.6;">─ Tab wird geschlossen ─</div>
        `;
        document.body.appendChild(box);

        const doClose = (e) => {
            if (e) e.preventDefault();
            document.removeEventListener('keydown', onKey, true);
            box.remove();
            closeCarolTabLikeCtrlW();
        };
        const onKey = (e) => {
            if (e.key === 'Enter' || e.key === ' ') doClose(e);
        };
        box.addEventListener('click', doClose);
        document.addEventListener('keydown', onKey, true);
    }

    function findExactCloseButton() {
        const candidates = [];
        for (const el of document.querySelectorAll('button, [role="button"], a')) {
            const txt = (el.textContent || '').trim();
            if (txt === 'Close') candidates.push(el);
        }
        return candidates.length ? candidates[candidates.length - 1] : null;
    }

    async function clickSavedDialogClose(timeout = 10000) {
        let passed = 0;
        while (passed < timeout) {
            if (abortMission) throw new Error('Abort');
            const btn = findExactCloseButton();
            if (btn) {
                try {
                    btn.style.outline = '2px solid #00ff00';
                    btn.style.boxShadow = '0 0 15px #00ff00';
                } catch (e) {}
                btn.click();
                return true;
            }
            await sleep(250);
            passed += 250;
        }
        return false;
    }

    function finishAndCloseTab() {
        sessionStorage.removeItem('arasaka_close_after_done');
        notifyWmsDone();
        setTimeout(removeHUD, 1200);
        showCloseTabPrompt();
        setTimeout(closeCarolTabLikeCtrlW, 800);
    }

    async function startMacro(skipPrint = false) {
        try {
            hudLog = []; currentStep = 0;
            createHUD();
            updateHUD('Initialisiere...');

            if (!skipPrint) {
                await sucheUndOeffnePdf({ removeHudOnComplete: false, fireAndForget: true });
                if (abortMission) return;
            }

            const result = await deliveredViaApi((step) => updateHUD(step, '#00ffcc', true));
            if (!result || result.status !== 'ok') {
                updateHUD((result && result.message) || 'Delivered fehlgeschlagen', '#ff4444', true);
                playDing('error');
                setTimeout(removeHUD, 5000);
                return;
            }

            if (skipPrint) {
                updateHUD('NACHBESTELLUNG ABGESCHLOSSEN ─ Kein Druck', '#00ff00', true);
            } else {
                updateHUD('AUFTRAG ABGESCHLOSSEN', '#00ff00', true);
            }
            playDing('success');
            if (launchedAsAuto) {
                finishAndCloseTab();
            } else {
                setTimeout(removeHUD, 5000);
            }

        } catch (e) {
            sessionStorage.removeItem('arasaka_close_after_done');
            if (e.message === 'Abort') {
                updateHUD('ABBRUCH DURCH USER', '#ff4444');
            } else {
                updateHUD(`SYSTEMFEHLER: ${e.message}`, '#ff4444');
            }
            playDing('error');
            setTimeout(removeHUD, 4000);
        }
    }

    if (sessionStorage.getItem('arasaka_close_after_done') === '1') {
        isLocked = true;
        playDing('success');
        finishAndCloseTab();
    } else if (launchedAsAuto) {
        const noPrint = window.location.href.includes('arasaka_noprint=1');
        setTimeout(() => {
            if (isLocked) return;
            isLocked = true;
            abortMission = false;
            startMacro(noPrint).finally(() => { isLocked = false; });
        }, 1500);
    }

    document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') {
            abortMission = true;
            isLocked = false;
        }
        if (event.altKey && event.key.toLowerCase() === 'y') {
            event.preventDefault();
            if (isLocked) return;
            isLocked = true;
            abortMission = false;
            startMacro().finally(() => { isLocked = false; });
        }
    });

})();
