// ==UserScript==
// @name         AAG Lieferschein Sync
// @namespace    http://tampermonkey.net/
// @version      1.2.2
// @description  Synchronisiert AAG Lieferscheine ins WMS
// @author       ARASAKA
// @match        https://portal.allianceautomotive.de/*
// @match        https://sso.aag-cp.de/*
// @run-at       document-start
// @grant        unsafeWindow
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @connect      portal.allianceautomotive.de
// @connect      script.google.com
// @connect      script.googleusercontent.com
// ==/UserScript==

(function() {
  'use strict';

  const SYNC_WEB_APP_URL = 'PASTE_AAG_SYNC_WEB_APP_EXEC_URL_HERE';
  const SYNC_KEY = 'PASTE_AAG_SYNC_KEY_HERE';
  const API_BASE = 'https://portal.allianceautomotive.de/api/v1/customer/receipt';
  const POLL_MS = 2 * 60 * 1000;
  const SEEN_KEY = 'aag_lieferschein_seen_v1';
  const BACKFILL_KEY = 'aag_lieferschein_backfill_4d_v1';
  const RELOAD_KEY = 'aag_lieferschein_reload_v1';
  const MAX_RECEIPTS_PER_RUN = 150;
  const UPLOAD_BATCH_SIZE = 25;
  const AUTO_RELOAD_MS = 5 * 60 * 1000;
  const RECOVER_KEY = 'aag_session_recover_v1';
  const CLEAN_KEY = 'aag_sheet_clean_v1';
  let bearerToken = '';
  let syncRunning = false;
  let syncTimer = 0;
  let statusEl = null;

  function configured() {
    return /^https:\/\/script\.google\.com\/.+\/exec$/i.test(SYNC_WEB_APP_URL)
      && SYNC_KEY.length >= 16
      && SYNC_KEY.indexOf('PASTE_') !== 0;
  }

  function ensureStatus() {
    if (statusEl && document.documentElement.contains(statusEl)) return statusEl;
    if (!document.body) return null;
    statusEl = document.createElement('div');
    statusEl.id = 'aag-lieferschein-sync-status';
    statusEl.style.cssText = 'position:fixed;right:14px;bottom:14px;z-index:2147483647;padding:9px 13px;border-radius:9px;background:#00244e;color:#fff;font:700 12px Arial,sans-serif;box-shadow:0 4px 18px rgba(0,0,0,.28);border:1px solid rgba(255,255,255,.25);cursor:default';
    statusEl.textContent = 'AAG Sync startet…';
    document.body.appendChild(statusEl);
    return statusEl;
  }

  function setStatus(text, color) {
    const el = ensureStatus();
    if (!el) {
      setTimeout(function() { setStatus(text, color); }, 300);
      return;
    }
    el.textContent = text;
    el.style.background = color || '#00244e';
  }

  function captureToken(value) {
    const raw = String(value || '');
    const match = raw.match(/^Bearer\s+(.+)$/i);
    if (!match || match[1].length < 40) return;
    bearerToken = match[1];
    setStatus(configured() ? 'AAG Sync bereit' : 'AAG Sync nicht eingerichtet', configured() ? '#176b45' : '#a45b00');
    scheduleSync(800);
  }

  function patchPortalRequests() {
    const w = unsafeWindow;
    if (!w || !w.XMLHttpRequest) return;
    const xhrProto = w.XMLHttpRequest.prototype;
    if (!xhrProto.__aagSyncPatched) {
      const originalSetRequestHeader = xhrProto.setRequestHeader;
      xhrProto.setRequestHeader = function(name, value) {
        if (String(name || '').toLowerCase() === 'authorization') captureToken(value);
        return originalSetRequestHeader.apply(this, arguments);
      };
      xhrProto.__aagSyncPatched = true;
    }
    if (w.fetch && !w.fetch.__aagSyncPatched) {
      const originalFetch = w.fetch;
      const wrappedFetch = function(input, init) {
        try {
          const headers = new w.Headers(init && init.headers || input && input.headers || {});
          captureToken(headers.get('authorization'));
        } catch (e) {}
        return originalFetch.apply(this, arguments);
      };
      wrappedFetch.__aagSyncPatched = true;
      w.fetch = wrappedFetch;
    }
  }

  function gmRequest(options) {
    return new Promise(function(resolve) {
      GM_xmlhttpRequest({
        method: options.method || 'GET',
        url: options.url,
        headers: options.headers || {},
        data: options.data,
        timeout: options.timeout || 30000,
        onload: function(r) { resolve(r); },
        onerror: function(r) { resolve(r || { status: 0, responseText: '' }); },
        ontimeout: function(r) { resolve(r || { status: 0, responseText: '' }); }
      });
    });
  }

  async function apiJson(url) {
    const response = await gmRequest({
      url: url,
      headers: {
        Accept: 'application/json, text/plain, */*',
        Authorization: 'Bearer ' + bearerToken
      }
    });
    if (response.status === 401 || response.status === 403) {
      bearerToken = '';
      throw new Error('LOGIN_REQUIRED');
    }
    if (response.status < 200 || response.status >= 300) {
      const detail = String(response.responseText || '')
        .replace(/\s+/g, ' ')
        .replace(/Bearer\s+[A-Za-z0-9._~-]+/gi, 'Bearer [hidden]')
        .slice(0, 180);
      throw new Error('AAG_HTTP_' + response.status + (detail ? ': ' + detail : ''));
    }
    return JSON.parse(String(response.responseText || '{}'));
  }

  function dateWindow() {
    const from = new Date();
    from.setDate(from.getDate() - 3);
    from.setHours(0, 0, 0, 0);
    const to = new Date();
    to.setHours(23, 59, 59, 999);
    return {
      from: Math.floor(from.getTime() / 1000),
      to: Math.floor(to.getTime() / 1000)
    };
  }

  async function fetchHeaders(seen, backfillPending) {
    const dates = dateWindow();
    const all = [];
    async function fetchPage(page) {
      const url = API_BASE + '/header?startDate=' + dates.from
        + '&endDate=' + dates.to
        + '&page=' + page
        + '&size=10&type=DELIVERY';
      const payload = await apiJson(url);
      const data = payload && Array.isArray(payload.data) ? payload.data : [];
      let unknownOnPage = 0;
      for (let i = 0; i < data.length; i++) {
        const ls = String(data[i].receiptNumber || '').toUpperCase();
        if (/^LS[A-Z0-9-]{5,}$/i.test(ls)) {
          all.push(data[i]);
          if (!seen[ls]) unknownOnPage++;
        }
      }
      return {
        pages: Math.max(1, Number(payload.totalPages) || 1),
        unknown: unknownOnPage
      };
    }
    const first = await fetchPage(1);
    const pages = first.pages;
    if (backfillPending) {
      for (let page = 2; page <= pages && page <= 100; page++) {
        await fetchPage(page);
      }
    } else if (pages > 1) {
      let knownPageStreak = 0;
      let checkedPages = 0;
      for (let page = pages; page >= 2 && checkedPages < 100; page--) {
        const result = await fetchPage(page);
        checkedPages++;
        knownPageStreak = result.unknown === 0 ? knownPageStreak + 1 : 0;
        if (knownPageStreak >= 2) break;
      }
    }
    return all;
  }

  async function fetchReceipt(receiptNumber) {
    const url = API_BASE + '/positions?page=0&size=100000000&type=DELIVERY&number=' + encodeURIComponent(receiptNumber);
    const payload = await apiJson(url);
    return {
      receiptNumber: receiptNumber,
      detailInfo: payload.detailInfo || {},
      positions: Array.isArray(payload.data) ? payload.data : []
    };
  }

  function loadSeen() {
    const raw = GM_getValue(SEEN_KEY, {});
    return raw && typeof raw === 'object' ? raw : {};
  }

  function trimSeen(seen) {
    const cutoff = Date.now() - 120 * 24 * 60 * 60 * 1000;
    const out = {};
    Object.keys(seen || {}).forEach(function(key) {
      if (Number(seen[key]) >= cutoff) out[key] = Number(seen[key]);
    });
    return out;
  }

  async function sendToSheet(receipts) {
    const response = await gmRequest({
      method: 'POST',
      url: SYNC_WEB_APP_URL,
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      data: JSON.stringify({
        action: receipts ? 'sync' : 'cleanup',
        key: SYNC_KEY,
        receipts: receipts || []
      }),
      timeout: 120000
    });
    if (response.status < 200 || response.status >= 400) throw new Error('SHEET_HTTP_' + response.status);
    const result = JSON.parse(String(response.responseText || '{}'));
    if (!result.success) {
      const detail = String(result.error || response.responseText || 'SHEET_REJECTED')
        .replace(/\s+/g, ' ')
        .slice(0, 180);
      throw new Error(detail);
    }
    return result;
  }

  async function cleanupOldReceipts() {
    const last = Number(GM_getValue(CLEAN_KEY, 0)) || 0;
    if (Date.now() - last < 6 * 60 * 60 * 1000) return;
    try {
      const result = await sendToSheet(null);
      if (result && result.success) GM_setValue(CLEAN_KEY, Date.now());
    } catch (e) {}
  }

  function scheduleSync(delay) {
    if (syncTimer) clearTimeout(syncTimer);
    syncTimer = setTimeout(runSync, delay == null ? POLL_MS : delay);
  }

  async function runSync() {
    if (navigator.locks && navigator.locks.request) {
      await navigator.locks.request('aag-lieferschein-sync-v1', { ifAvailable: true }, async function(lock) {
        if (!lock) {
          scheduleSync(POLL_MS);
          return;
        }
        await runSyncExclusive();
      });
      return;
    }
    await runSyncExclusive();
  }

  async function runSyncExclusive() {
    if (syncRunning) return;
    if (!configured()) {
      setStatus('AAG Sync: URL/Key eintragen', '#a45b00');
      scheduleSync(POLL_MS);
      return;
    }
    if (!bearerToken) {
      setStatus('AAG Sync: Portal Login nötig', '#a11d2f');
      scheduleSync(15000);
      return;
    }
    syncRunning = true;
    let nextSyncDelay = POLL_MS;
    setStatus('AAG Sync läuft…', '#005f8f');
    try {
      await cleanupOldReceipts();
      const seen = trimSeen(loadSeen());
      const backfillPending = !GM_getValue(BACKFILL_KEY, false);
      const headers = await fetchHeaders(seen, backfillPending);
      const allFresh = [];
      const freshSet = {};
      for (let i = 0; i < headers.length; i++) {
        const ls = String(headers[i].receiptNumber || '').toUpperCase();
        if (!seen[ls] && !freshSet[ls]) {
          freshSet[ls] = true;
          allFresh.push(ls);
        }
      }
      const fresh = allFresh.slice(0, MAX_RECEIPTS_PER_RUN);
      const syncHasMore = allFresh.length > MAX_RECEIPTS_PER_RUN;
      if (syncHasMore) GM_setValue(BACKFILL_KEY, false);
      if (!fresh.length) {
        GM_setValue(SEEN_KEY, seen);
        if (backfillPending) GM_setValue(BACKFILL_KEY, true);
        setStatus('AAG Sync aktuell · ' + new Date().toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' }), '#176b45');
      } else {
        const receipts = [];
        for (let r = 0; r < fresh.length; r++) {
          setStatus('AAG Sync ' + (r + 1) + '/' + fresh.length, '#005f8f');
          receipts.push(await fetchReceipt(fresh[r]));
        }
        let synced = 0;
        for (let offset = 0; offset < receipts.length; offset += UPLOAD_BATCH_SIZE) {
          const batch = receipts.slice(offset, offset + UPLOAD_BATCH_SIZE);
          setStatus('AAG speichert ' + (offset + 1) + '–' + Math.min(offset + batch.length, receipts.length) + '/' + receipts.length, '#005f8f');
          await sendToSheet(batch);
          const now = Date.now();
          for (let s = offset; s < offset + batch.length; s++) seen[fresh[s]] = now;
          synced += batch.length;
          GM_setValue(SEEN_KEY, seen);
        }
        if (backfillPending && !syncHasMore) GM_setValue(BACKFILL_KEY, true);
        if (syncHasMore) nextSyncDelay = 5000;
        setStatus('AAG Sync: ' + synced + ' neu' + (syncHasMore ? ' · Rückstand folgt' : '') + ' · ' + new Date().toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' }), '#176b45');
      }
    } catch (err) {
      const msg = String(err && err.message || err);
      if (msg === 'LOGIN_REQUIRED') {
        setStatus('AAG Sync: Sitzung wird erneuert…', '#a45b00');
        GM_setValue(RECOVER_KEY, Date.now());
        GM_setValue(RELOAD_KEY, Date.now());
        setTimeout(function() {
          unsafeWindow.location.reload();
        }, 1200);
      } else {
        setStatus('AAG Sync Fehler: ' + msg, '#a11d2f');
      }
    } finally {
      syncRunning = false;
      scheduleSync(nextSyncDelay);
    }
  }

  function controlParts(el) {
    if (!el) return [];
    return [el.innerText, el.textContent, el.value, el.getAttribute('aria-label'), el.getAttribute('title')]
      .map(function(value) { return String(value || '').replace(/\s+/g, ' ').trim(); })
      .filter(Boolean);
  }

  function findControl(pattern) {
    const nodes = document.querySelectorAll('a, button, input[type="submit"], [role="button"]');
    for (let i = 0; i < nodes.length; i++) {
      if (controlParts(nodes[i]).some(function(text) { return pattern.test(text); })) return nodes[i];
    }
    return null;
  }

  function deliveryTableVisible() {
    const text = document.body ? document.body.innerText : '';
    return /Belegnummer/i.test(text) && /Belegdatum/i.test(text) && /Kundenreferenz/i.test(text);
  }

  function findLieferscheineCard() {
    const nodes = document.querySelectorAll('a, button, div');
    let best = null;
    let bestArea = 0;
    for (let i = 0; i < nodes.length; i++) {
      const text = String(nodes[i].innerText || '').replace(/\s+/g, ' ').trim();
      if (!/lieferscheine/i.test(text) || text.length > 80) continue;
      const rect = nodes[i].getBoundingClientRect();
      const area = rect.width * rect.height;
      if (rect.width < 80 || rect.height < 70 || area <= bestArea) continue;
      best = nodes[i];
      bestArea = area;
    }
    return best;
  }

  function enableStaySignedIn() {
    const labels = document.querySelectorAll('label');
    for (let i = 0; i < labels.length; i++) {
      if (!controlParts(labels[i]).some(function(text) { return /angemeldet bleiben/i.test(text); })) continue;
      const box = labels[i].querySelector('input[type="checkbox"]')
        || document.getElementById(labels[i].getAttribute('for') || '');
      if (box && box.type === 'checkbox' && !box.checked) box.click();
      return;
    }
  }

  let deliveryNavAt = 0;

  function openDeliveries() {
    if (deliveryTableVisible() || syncRunning || Date.now() - deliveryNavAt < 4000) return;
    const card = findLieferscheineCard();
    if (card) {
      setStatus('AAG Sync: öffne Lieferscheine…', '#005f8f');
      deliveryNavAt = Date.now();
      card.click();
      return;
    }
    const portalButton = findControl(/zum kundenportal/i);
    if (portalButton && location.pathname.indexOf('/ordering') === 0) {
      if (portalButton.tagName === 'A') portalButton.target = '_self';
      setStatus('AAG Sync: öffne Kundenportal…', '#005f8f');
      deliveryNavAt = Date.now();
      portalButton.click();
      return;
    }
    setStatus('AAG Sync: öffne Kundenportal…', '#005f8f');
    deliveryNavAt = Date.now();
    location.href = 'https://portal.allianceautomotive.de/dashboard';
  }

  function recoverPortalSession() {
    const onLogin = location.hostname === 'sso.aag-cp.de' || /openid-connect\/auth/i.test(location.pathname);
    if (onLogin) {
      enableStaySignedIn();
      setStatus('AAG Sync: einmal auf Anmelden klicken', '#a45b00');
      return;
    }
    if (location.hostname !== 'portal.allianceautomotive.de') return;
    const active = document.activeElement;
    const userTyping = !document.hidden && active && /^(INPUT|TEXTAREA|SELECT)$/.test(active.tagName);
    if (deliveryTableVisible()) {
      GM_setValue(RECOVER_KEY, 0);
      return;
    }
    if (userTyping) return;
    const recovering = Date.now() - (Number(GM_getValue(RECOVER_KEY, 0)) || 0) < 3 * 60 * 1000;
    const onOrdering = location.pathname.indexOf('/ordering') === 0;
    if (findLieferscheineCard() || onOrdering || recovering || document.hidden || !bearerToken) openDeliveries();
  }

  patchPortalRequests();
  document.addEventListener('DOMContentLoaded', function() {
    ensureStatus();
    if (!GM_getValue(RELOAD_KEY, 0)) GM_setValue(RELOAD_KEY, Date.now());
    if (!bearerToken) setStatus(configured() ? 'AAG Sync wartet auf Portal…' : 'AAG Sync: URL/Key eintragen', configured() ? '#005f8f' : '#a45b00');
  });
  setInterval(function() {
    patchPortalRequests();
  }, 30000);
  setInterval(recoverPortalSession, 2000);
  setInterval(function() {
    if (location.hostname === 'sso.aag-cp.de' || syncRunning) return;
    const lastReload = Number(GM_getValue(RELOAD_KEY, 0)) || 0;
    if (Date.now() - lastReload < AUTO_RELOAD_MS) return;
    GM_setValue(RELOAD_KEY, Date.now());
    unsafeWindow.location.reload();
  }, 30000);
})();
