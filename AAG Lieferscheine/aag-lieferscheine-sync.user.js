// ==UserScript==
// @name         AAG Lieferschein Sync
// @namespace    http://tampermonkey.net/
// @version      1.0.0
// @description  Synchronisiert AAG Lieferscheine ins WMS
// @author       ARASAKA
// @match        https://portal.allianceautomotive.de/*
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
  const MAX_RECEIPTS_PER_RUN = 150;
  let bearerToken = '';
  let syncRunning = false;
  let syncTimer = 0;
  let statusEl = null;
  let lastTokenAt = 0;

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
    lastTokenAt = Date.now();
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
    if (response.status < 200 || response.status >= 300) throw new Error('AAG_HTTP_' + response.status);
    return JSON.parse(String(response.responseText || '{}'));
  }

  function dateWindow() {
    const from = new Date();
    from.setDate(from.getDate() - 1);
    from.setHours(0, 0, 0, 0);
    const to = new Date();
    to.setDate(to.getDate() + 1);
    to.setHours(23, 59, 59, 999);
    return {
      from: Math.floor(from.getTime() / 1000),
      to: Math.floor(to.getTime() / 1000)
    };
  }

  async function fetchHeaders() {
    const dates = dateWindow();
    let page = 0;
    let pages = 1;
    const all = [];
    while (page < pages && page < 20) {
      const url = API_BASE + '/header?startDate=' + dates.from
        + '&endDate=' + dates.to
        + '&page=' + page
        + '&size=100&type=DELIVERY&orderColumn=documentDate&order=DESC';
      const payload = await apiJson(url);
      const data = payload && Array.isArray(payload.data) ? payload.data : [];
      for (let i = 0; i < data.length; i++) {
        if (/^LS[A-Z0-9-]{5,}$/i.test(String(data[i].receiptNumber || ''))) all.push(data[i]);
      }
      pages = Math.max(1, Number(payload.totalPages) || 1);
      page++;
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
        action: 'sync',
        key: SYNC_KEY,
        receipts: receipts
      }),
      timeout: 60000
    });
    if (response.status < 200 || response.status >= 400) throw new Error('SHEET_HTTP_' + response.status);
    const result = JSON.parse(String(response.responseText || '{}'));
    if (!result.success) throw new Error(result.error || 'SHEET_REJECTED');
    return result;
  }

  function scheduleSync(delay) {
    if (syncTimer) clearTimeout(syncTimer);
    syncTimer = setTimeout(runSync, delay == null ? POLL_MS : delay);
  }

  async function runSync() {
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
    setStatus('AAG Sync läuft…', '#005f8f');
    try {
      const headers = await fetchHeaders();
      const seen = trimSeen(loadSeen());
      const fresh = [];
      for (let i = 0; i < headers.length && fresh.length < MAX_RECEIPTS_PER_RUN; i++) {
        const ls = String(headers[i].receiptNumber || '').toUpperCase();
        if (!seen[ls]) fresh.push(ls);
      }
      if (!fresh.length) {
        GM_setValue(SEEN_KEY, seen);
        setStatus('AAG Sync aktuell · ' + new Date().toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' }), '#176b45');
      } else {
        const receipts = [];
        for (let r = 0; r < fresh.length; r++) {
          setStatus('AAG Sync ' + (r + 1) + '/' + fresh.length, '#005f8f');
          receipts.push(await fetchReceipt(fresh[r]));
        }
        const result = await sendToSheet(receipts);
        const now = Date.now();
        for (let s = 0; s < fresh.length; s++) seen[fresh[s]] = now;
        GM_setValue(SEEN_KEY, seen);
        setStatus('AAG Sync: ' + result.receipts + ' neu · ' + new Date().toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' }), '#176b45');
      }
    } catch (err) {
      const msg = String(err && err.message || err);
      if (msg === 'LOGIN_REQUIRED') {
        setStatus('AAG Sync: neu einloggen', '#a11d2f');
      } else {
        setStatus('AAG Sync Fehler: ' + msg, '#a11d2f');
      }
    } finally {
      syncRunning = false;
      scheduleSync(POLL_MS);
    }
  }

  patchPortalRequests();
  document.addEventListener('DOMContentLoaded', function() {
    ensureStatus();
    if (!bearerToken) setStatus(configured() ? 'AAG Sync wartet auf Portal…' : 'AAG Sync: URL/Key eintragen', configured() ? '#005f8f' : '#a45b00');
  });
  setInterval(function() {
    patchPortalRequests();
  }, 30000);
})();
