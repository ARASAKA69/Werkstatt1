// ==UserScript==
// @name         ALFAH Retoure Bot
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Legt ALFAH Retouren an, zeigt die R-Nummer und schreibt sie ins Sheet
// @author       ARASAKA
// @match        https://alfah.de/*
// @match        https://mail.google.com/*
// @run-at       document-idle
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addValueChangeListener
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      alfah.de
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @connect      127.0.0.1
// @connect      localhost
// @updateURL    https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/Warenr%C3%BCckgabe/alfah-retoure-automation.user.js
// @downloadURL  https://github.com/ARASAKA69/Werkstatt1/raw/refs/heads/main/Warenr%C3%BCckgabe/alfah-retoure-automation.user.js
// ==/UserScript==

(function () {
    'use strict';

    var BOT_VERSION = '1.0';
    var BASE = 'https://alfah.de/';
    var QUEUE_KEY = 'alfah_retoure_queue';
    var ORDER_CACHE_KEY = 'alfah_retoure_orders';
    var WORKER_KEY = 'alfah_retoure_worker';
    var SHEET_URL_KEY = 'alfah_retoure_sheet_url';
    var SHEET_WEB_APP_URL = '';
    var HISTORY_KEY = 'alfah_retoure_history';
    var ADMIN_PASS = '1208';
    var REASON_WANTED = 'ITEM_IS_NOT_NEEDED';
    var PROBLEM_DEFAULT = 'Teil wird nicht benötigt';

    function pageWindow() {
        return typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
    }

    function token() {
        try {
            return pageWindow().sessionStorage.getItem('token') || '';
        } catch (e) {
            return '';
        }
    }

    function tokenInfo() {
        try {
            return JSON.parse(pageWindow().sessionStorage.getItem('alfahTokenInformation') || 'null') || {};
        } catch (e) {
            return {};
        }
    }

    function customerId() {
        var info = tokenInfo();
        return info.principalId || info.customerId || info.partyId || '';
    }

    function customerName() {
        return tokenInfo().displayName || '';
    }

    function saveQueue(data) {
        GM_setValue(QUEUE_KEY, JSON.stringify(data || null));
    }

    function loadQueue() {
        try {
            var raw = GM_getValue(QUEUE_KEY, '');
            if (!raw) return null;
            var data = JSON.parse(raw);
            if (!data || !data.orderId) return null;
            if (data.savedAt && Date.now() - data.savedAt > 36e5 * 18) return null;
            return data;
        } catch (e) {
            return null;
        }
    }

    function explain(body) {
        if (!body) return '';
        if (typeof body === 'string') return body.slice(0, 600);
        return body.message || body.Message || body.title || body.detail || body.Detail || JSON.stringify(body).slice(0, 600);
    }

    function request(method, url, opts) {
        opts = opts || {};
        return new Promise(function (resolve, reject) {
            var headers = {
                Auth: token(),
                'Accept-Language': 'de'
            };
            if (opts.headers) {
                Object.keys(opts.headers).forEach(function (key) {
                    headers[key] = opts.headers[key];
                });
            }
            if (opts.json) headers['Content-Type'] = 'application/json';
            GM_xmlhttpRequest({
                method: method,
                url: url,
                headers: headers,
                data: opts.json ? JSON.stringify(opts.json) : opts.data,
                responseType: opts.responseType || 'text',
                onload: function (res) {
                    var body = res.response;
                    if (opts.responseType !== 'arraybuffer' && opts.responseType !== 'blob') {
                        var text = typeof res.responseText === 'string' ? res.responseText : '';
                        var parsed = null;
                        if (text && text !== '[object Object]') {
                            try {
                                parsed = JSON.parse(text);
                            } catch (e) {
                                parsed = text;
                            }
                        }
                        if (parsed && typeof parsed === 'object') body = parsed;
                        else if (body && typeof body === 'object') body = body;
                        else if (parsed != null) body = parsed;
                    }
                    if (res.status >= 200 && res.status < 300) {
                        resolve({ status: res.status, body: body, headers: res.responseHeaders || '' });
                    } else {
                        reject(new Error(res.status + ' ' + explain(body || res.responseText)));
                    }
                },
                onerror: function () {
                    reject(new Error('Netzwerkfehler'));
                }
            });
        });
    }

    function api(method, path, opts) {
        return request(method, BASE + path.replace(/^\//, ''), opts).then(function (res) {
            return res.body;
        });
    }

    function asList(data) {
        if (Array.isArray(data)) return data;
        if (!data || typeof data !== 'object') return [];
        if (Array.isArray(data.value)) return data.value;
        if (Array.isArray(data.Items)) return data.Items;
        if (Array.isArray(data.items)) return data.items;
        if (Array.isArray(data.data)) return data.data;
        return [];
    }

    function parseMail(text) {
        var raw = String(text || '').replace(/\u00a0/g, ' ').replace(/[\u200b\u200c\u200d\ufeff]/g, '').replace(/\uFF1A/g, ':').replace(/\uFF08/g, '(').replace(/\uFF09/g, ')');
        var orderMatch = raw.match(/Bestellung\s*Nr\.?\s*#?\s*(\d{5,})/i) || raw.match(/Auftragsbestätigung\s*#\s*(\d{5,})/i) || raw.match(/#(\d{6,})/);
        var komMatch = raw.match(/Kommission:\s*([^\n]+)/i);
        var re = /\[([A-Za-z0-9]{4,})\]\s*([^:\n]{2,40}):\s*([^(\n]{2,}?)\s*\(([^)\n]{2,})\)/g;
        var found = [];
        var match;
        while ((match = re.exec(raw))) found.push({ match: match, index: match.index });
        var items = found.map(function (entry, index) {
            var next = found[index + 1] ? found[index + 1].index : entry.index + 900;
            var chunk = raw.slice(entry.index, next);
            var vinMatch = chunk.match(/VIN\s*:\s*([A-HJ-NPR-Z0-9]{17})/i);
            var qtyMatch = chunk.match(/(\d+)\s*Stk/i);
            var hit = entry.match;
            return {
                article: hit[1],
                brand: hit[2].trim(),
                name: hit[3].trim(),
                supplier: hit[4].trim(),
                vin: vinMatch ? vinMatch[1].toUpperCase() : '',
                qty: qtyMatch ? qtyMatch[1] : '1'
            };
        });
        var kommission = komMatch ? komMatch[1].trim() : '';
        return {
            orderId: orderMatch ? orderMatch[1] : '',
            kommission: kommission,
            stockId: stockIdFromCommission(kommission),
            items: items
        };
    }

    function stockIdFromCommission(value) {
        var token = String(value || '').trim().split(/[\s\/]+/)[0] || '';
        if (!/^[A-Za-z][A-Za-z0-9]{2,}$/.test(token)) return '';
        return token.toUpperCase();
    }

    function strictStockId(value) {
        var match = String(value || '').toUpperCase().match(/[A-Z]{2}\d{5}/);
        return match ? match[0] : '';
    }

    function workerName() {
        try {
            return String(GM_getValue(WORKER_KEY, '') || '').trim();
        } catch (e) {
            return '';
        }
    }

    function sheetUrl() {
        var saved = '';
        try {
            saved = String(GM_getValue(SHEET_URL_KEY, '') || '').trim();
        } catch (e) {}
        if (/^https:\/\/script\.google\.com\/.+\/exec\/?$/i.test(saved)) return saved;
        var built = String(SHEET_WEB_APP_URL || '').trim();
        if (/^https:\/\/script\.google\.com\/.+\/exec\/?$/i.test(built)) return built;
        return '';
    }

    function formatBerlin(date) {
        var fmt = new Intl.DateTimeFormat('de-DE', {
            timeZone: 'Europe/Berlin',
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            hourCycle: 'h23'
        });
        var map = {};
        fmt.formatToParts(date).forEach(function (part) { map[part.type] = part.value; });
        var hour = map.hour === '24' ? '00' : map.hour;
        return {
            day: map.year + '-' + map.month + '-' + map.day,
            date: map.day + '.' + map.month + '.' + map.year,
            time: hour + ':' + map.minute
        };
    }

    function berlinStamp(value) {
        if (value instanceof Date) return formatBerlin(value);
        if (typeof value === 'number') {
            var fromNum = new Date(value);
            if (!isNaN(fromNum.getTime())) return formatBerlin(fromNum);
            return { day: '', date: '', time: '' };
        }
        if (value && typeof value === 'object') return berlinStamp(value.Date || value.date || value.Value || value.value || '');
        var text = String(value || '').trim();
        if (!text) return { day: '', date: '', time: '' };
        var ms = text.match(/\/Date\((\d+)\)\//);
        if (ms) return berlinStamp(Number(ms[1]));
        var iso = text.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T\s](\d{2}):(\d{2}))?/);
        if (iso) {
            if (iso[4]) {
                var parsed = new Date(text);
                if (!isNaN(parsed.getTime())) return formatBerlin(parsed);
            }
            return { day: iso[1] + '-' + iso[2] + '-' + iso[3], date: iso[3] + '.' + iso[2] + '.' + iso[1], time: iso[4] ? (iso[4] + ':' + iso[5]) : '' };
        }
        var de = text.match(/^(\d{2})[.\-](\d{2})[.\-](\d{4})(?:[ T](\d{2}):(\d{2}))?/);
        if (de) {
            if (de[4]) {
                var local = new Date(de[3] + '-' + de[2] + '-' + de[1] + 'T' + de[4] + ':' + de[5] + ':00');
                if (!isNaN(local.getTime())) return formatBerlin(local);
            }
            return { day: de[3] + '-' + de[2] + '-' + de[1], date: de[1] + '.' + de[2] + '.' + de[3], time: de[4] ? (de[4] + ':' + de[5]) : '' };
        }
        var loose = Date.parse(text);
        if (!isNaN(loose)) return formatBerlin(new Date(loose));
        return { day: '', date: '', time: '' };
    }

    function berlinNow() {
        return formatBerlin(new Date());
    }

    function stampFromCase(row) {
        if (!row || typeof row !== 'object') return { day: '', date: '', time: '' };
        var preferred = [];
        var fallback = [];
        Object.keys(row).forEach(function (key) {
            if (!/date|datum|time|created|erstell|insert|anleg/i.test(key)) return;
            if (/update|modif|änder|aender|action|letzt|change/i.test(key)) fallback.push(key);
            else preferred.push(key);
        });
        var lists = [preferred, fallback];
        var g;
        for (g = 0; g < lists.length; g++) {
            var i;
            for (i = 0; i < lists[g].length; i++) {
                var hit = berlinStamp(row[lists[g][i]]);
                if (hit.day) return hit;
            }
        }
        return { day: '', date: '', time: '' };
    }

    function casePerson(row) {
        if (!row || typeof row !== 'object') return '';
        var keys = Object.keys(row);
        var i;
        for (i = 0; i < keys.length; i++) {
            if (!/createdby|author|anmelder|username|editor/i.test(keys[i])) continue;
            var value = row[keys[i]];
            var name = '';
            if (typeof value === 'string') name = value.trim();
            else if (value && typeof value === 'object') name = String(value.Name || value.DisplayName || value.UserName || '').trim();
            if (!name || name.length > 40 || /^\d+$/.test(name) || /@|gmbh|\bag\b/i.test(name)) continue;
            return name;
        }
        return '';
    }

    function caseStock(row) {
        if (!row || typeof row !== 'object') return '';
        var bits = [];
        function push(value) {
            if (value == null) return;
            if (typeof value === 'string') bits.push(value);
            else if (typeof value === 'object') bits.push(value.Content || value.Text || value.Note || value.Value || '');
        }
        push(row.Commission);
        push(row.Kommission);
        push(row.Comment);
        var notes = row.Notes || row.notes || row.Comments || [];
        if (!Array.isArray(notes)) notes = notes ? [notes] : [];
        notes.forEach(push);
        return strictStockId(bits.join(' '));
    }

    function caseSupplier(row) {
        if (!row) return '';
        if (typeof row.Supplier === 'string') return row.Supplier;
        if (row.SupplierName) return String(row.SupplierName);
        if (row.Supplier && typeof row.Supplier === 'object') return row.Supplier.Name || row.Supplier.CompanyName || '';
        return '';
    }

    function sheetEcho(txt) {
        var match = String(txt || '').replace(/&amp;/g, '&').match(/https:\/\/script\.googleusercontent\.com\/macros\/echo[^"'\\\s<>]*/);
        return match ? match[0] : '';
    }

    function sheetParse(txt) {
        var raw = String(txt || '').trim();
        if (raw.charAt(0) === '{') {
            try { return JSON.parse(raw); } catch (e) {}
        }
        var pre = raw.match(/<pre[^>]*>([\s\S]*?)<\/pre>/i);
        if (pre) {
            try {
                return JSON.parse(pre[1].replace(/&quot;/g, '"').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>'));
            } catch (e2) {}
        }
        return null;
    }

    function sheetReq(method, url, body, hop) {
        return new Promise(function (resolve, reject) {
            GM_xmlhttpRequest({
                method: method,
                url: url,
                data: body || undefined,
                headers: method === 'POST' ? { 'Content-Type': 'text/plain;charset=utf-8' } : {},
                anonymous: false,
                timeout: 30000,
                onload: function (res) {
                    var txt = res && res.responseText != null ? res.responseText : '';
                    var data = sheetParse(txt);
                    if (data && (data.ok === true || data.ok === false)) {
                        if (data.ok === false) reject(new Error(data.error || 'Sheet hat abgelehnt'));
                        else resolve(data);
                        return;
                    }
                    var jump = hop < 2 ? sheetEcho(txt) : '';
                    if (jump) {
                        sheetReq('GET', jump, '', hop + 1).then(resolve, reject);
                        return;
                    }
                    var low = String(txt || '').toLowerCase();
                    if (/accounts\.google|servicelogin|anmelden/.test(low)) reject(new Error('Google-Login fehlt fürs Sheet'));
                    else reject(new Error('Sheet Antwort ' + (res && res.status)));
                },
                onerror: function () { reject(new Error('Sheet nicht erreichbar')); },
                ontimeout: function () { reject(new Error('Sheet Timeout')); }
            });
        });
    }

    function packSheet(rows) {
        var groups = [];
        var map = {};
        (rows || []).forEach(function (row) {
            if (!row || !/^R-\d+$/.test(row.caseId || '')) return;
            var key = row.caseId;
            if (!map[key]) {
                map[key] = { lines: [] };
                groups.push(map[key]);
            }
            map[key].lines.push(row);
        });
        return groups;
    }

    function postSheet(rows) {
        var url = sheetUrl();
        if (!url) return Promise.reject(new Error('Sheet-Web-App fehlt noch'));
        var entries = Array.isArray(rows) && rows.length && rows[0].lines ? rows : packSheet(rows);
        if (!entries.length) return Promise.resolve({ ok: true, added: 0, skipped: 0 });
        return sheetReq('POST', url, JSON.stringify({ action: 'append', entries: entries }), 0);
    }

    function repairSheetWho() {
        var name = workerName();
        var url = sheetUrl();
        if (!name || !url) return Promise.resolve({ ok: true, updated: 0 });
        return sheetReq('POST', url, JSON.stringify({ action: 'fixWho', who: name }), 0);
    }

    function orderIdOf(order, fallback) {
        var candidates = [order && order.OrderId, order && order.OrderNumber, order && order.OrderProcessId, order && order.Number, order && order.Id, order && order.id, fallback];
        var i;
        for (i = 0; i < candidates.length; i++) {
            if (/^\d{5,}$/.test(String(candidates[i] || ''))) return String(candidates[i]);
        }
        for (i = 0; i < candidates.length; i++) {
            if (candidates[i]) return String(candidates[i]);
        }
        return '';
    }

    function attrName(row) {
        return String((row && (row.Name || row.name || row.Key || row.AttributeName || row.Code || row.Id)) || '');
    }

    function attrValue(row) {
        if (!row) return '';
        var value = row.Value != null ? row.Value : (row.value != null ? row.value : row.AttributeValue);
        if (value && typeof value === 'object') return String(value.Value || value.value || '');
        return value == null ? '' : String(value);
    }

    function commissionOf(order) {
        if (!order) return '';
        var attrs = order.OrderAttributes || order.orderAttributes || order.Attributes || [];
        if (attrs && !Array.isArray(attrs) && typeof attrs === 'object') {
            var direct = attrs.ORDER_COMMISSION || attrs.Kommission || attrs.Commission;
            if (direct && typeof direct === 'object') return attrValue(direct);
            if (direct) return String(direct);
            attrs = [];
        }
        var i;
        for (i = 0; i < attrs.length; i++) {
            var name = attrName(attrs[i]);
            if (name === 'ORDER_COMMISSION' || /commission|kommission/i.test(name)) return attrValue(attrs[i]);
        }
        return String(order.Commission || order.Kommission || '');
    }

    function looksLikeLine(row) {
        return !!(row && typeof row === 'object' && !Array.isArray(row) && (row.ArticleNumber || row.CatalogNumber || row.ArticleId || row.ExternalArticleId || row.article || row.SearchedNumber || row.PartNumber));
    }

    function lineScore(key, rows) {
        var score = rows.length * 10;
        if (/^orderitems$/i.test(key)) score += 80;
        else if (/items|positions|lines|articles/i.test(key)) score += 30;
        rows.forEach(function (row) {
            if (row.ArticleNumber || row.CatalogNumber) score += 8;
            if (row.Quantity) score += 2;
        });
        return score;
    }

    function lineOf(row) {
        if (looksLikeLine(row)) return row;
        if (!row || typeof row !== 'object' || Array.isArray(row)) return null;
        var keys = ['Article', 'article', 'Product', 'Item', 'OrderItem'];
        var i;
        for (i = 0; i < keys.length; i++) {
            var nested = row[keys[i]];
            if (!looksLikeLine(nested)) continue;
            return {
                ArticleNumber: nested.CatalogNumber || nested.ArticleNumber || nested.ExternalArticleId || '',
                CatalogNumber: nested.CatalogNumber || '',
                ArticleId: nested.ArticleId || nested.Id || null,
                Description: nested.Description || nested.Name || row.Description || row.Name || '',
                Name: nested.Name || '',
                Manufacturer: nested.Manufacturer || nested.ManufacturerName || row.Manufacturer || row.ManufacturerName || '',
                ManufacturerName: nested.ManufacturerName || row.ManufacturerName || '',
                Supplier: row.Supplier || row.SupplierName || nested.Supplier || nested.SupplierName || '',
                SupplierId: row.SupplierId || nested.SupplierId || '',
                Quantity: row.Quantity || nested.Quantity || 1,
                Attributes: nested.Attributes || row.Attributes || row.OrderItemAttributes || [],
                ItemId: row.ItemId || row.OrderItemId || row.Id || null,
                OrderItemId: row.OrderItemId || row.ItemId || null
            };
        }
        return null;
    }

    function collectLineSets(value, depth, keyName, sets) {
        if (!value || depth > 8 || sets.length > 40) return;
        if (Array.isArray(value)) {
            var lines = value.map(lineOf).filter(Boolean);
            if (lines.length) sets.push({ key: keyName, rows: lines });
            value.forEach(function (row) { collectLineSets(row, depth + 1, keyName, sets); });
            return;
        }
        if (typeof value !== 'object') return;
        Object.keys(value).forEach(function (key) {
            collectLineSets(value[key], depth + 1, key, sets);
        });
    }

    function bestLines(order) {
        var sets = [];
        collectLineSets(order, 0, '', sets);
        sets.sort(function (a, b) { return lineScore(b.key, b.rows) - lineScore(a.key, a.rows); });
        return sets.length ? sets[0].rows : [];
    }

    function describedParts(order) {
        var texts = [];
        function walk(value, depth) {
            if (!value || depth > 7 || texts.length > 30) return;
            if (typeof value === 'string') {
                if (value.indexOf('[') !== -1 && value.indexOf(']') !== -1) texts.push(value);
                return;
            }
            if (typeof value !== 'object') return;
            Object.keys(value).forEach(function (key) { walk(value[key], depth + 1); });
        }
        walk(order, 0);
        var items = [];
        texts.forEach(function (text) {
            parseMail(text).items.forEach(function (item) {
                if (!items.some(function (row) { return row.article === item.article; })) items.push(item);
            });
        });
        return items;
    }

    function rowsFromMail(parts, apiRows) {
        var used = {};
        return (parts || []).map(function (part) {
            var articleIndex = -1;
            var nameIndex = -1;
            var supplierIndex = -1;
            (apiRows || []).forEach(function (row, rowIndex) {
                if (articleIndex < 0 && normArticle(articleNo(row)) && normArticle(articleNo(row)) === normArticle(part.article)) articleIndex = rowIndex;
                if (nameIndex < 0 && !used[rowIndex] && sameName(partName(row), part.name)) nameIndex = rowIndex;
                if (supplierIndex < 0 && sameName(supplierNameOf(row), part.supplier)) supplierIndex = rowIndex;
            });
            var hitIndex = articleIndex >= 0 ? articleIndex : nameIndex;
            if (hitIndex >= 0) used[hitIndex] = true;
            var api = hitIndex >= 0 ? apiRows[hitIndex] : {};
            var supplierRow = supplierIndex >= 0 ? apiRows[supplierIndex] : api;
            return {
                ArticleNumber: part.article,
                Description: part.name,
                Manufacturer: part.brand,
                Supplier: part.supplier || supplierNameOf(supplierRow),
                SupplierId: supplierIdOf(supplierRow) || supplierIdOf(api),
                Quantity: Number(part.qty) || api.Quantity || 1,
                Vin: part.vin || findVin(api, 0) || '',
                ArticleId: api.ArticleId || null,
                ItemId: api.ItemId || api.OrderItemId || null,
                OrderItemId: api.OrderItemId || api.ItemId || null,
                Attributes: api.Attributes || []
            };
        });
    }

    function orderItems(order) {
        if (order && Array.isArray(order._alfahMailRows) && order._alfahMailRows.length) return order._alfahMailRows;
        var rows = bestLines(order);
        var described = describedParts(order);
        if (described.length > rows.length) {
            return described.map(function (part, index) {
                var api = rows.filter(function (row) {
                    return normArticle(articleNo(row)) === normArticle(part.article) || sameName(row.Description || row.Name, part.name);
                })[0] || rows[index] || rows[0] || {};
                return {
                    ArticleNumber: part.article,
                    Description: part.name,
                    Manufacturer: part.brand,
                    Supplier: part.supplier || supplierNameOf(api),
                    SupplierId: supplierIdOf(api),
                    Quantity: Number(part.qty) || api.Quantity || 1,
                    Vin: part.vin,
                    ArticleId: api.ArticleId || null,
                    ItemId: api.ItemId || api.OrderItemId || null,
                    OrderItemId: api.OrderItemId || api.ItemId || null
                };
            });
        }
        return rows;
    }

    function bracketArticle(item) {
        var text = String((item && (item.Description || item.Name)) || '');
        var match = text.match(/\[([A-Za-z0-9]{4,})\]/);
        return match ? match[1] : '';
    }

    function articleNo(item) {
        var bracket = bracketArticle(item);
        if (bracket) return bracket;
        return String(item.CatalogNumber || item.ExternalArticleId || item.SearchedNumber || item.PartNumber || item.ArticleNumber || item.article || '').replace(/\s+/g, '');
    }

    function brandOf(item) {
        if (!item) return '';
        if (typeof item.Manufacturer === 'string' && item.Manufacturer) return item.Manufacturer;
        if (item.Manufacturer && typeof item.Manufacturer === 'object') return item.Manufacturer.Name || item.Manufacturer.ManufacturerName || '';
        if (item.ManufacturerName) return String(item.ManufacturerName);
        var attrs = item.Attributes || item.OrderItemAttributes || [];
        if (!Array.isArray(attrs)) attrs = [];
        var i;
        for (i = 0; i < attrs.length; i++) {
            if (/CAR_MANUFACTURER|HERSTELLER|MANUFACTURER/i.test(attrName(attrs[i]))) {
                var named = attrValue(attrs[i]);
                if (named) return named;
            }
        }
        var parsed = parseMail(item.Description || item.Name || '').items[0];
        return parsed ? parsed.brand : '';
    }

    function partName(item) {
        var parsed = parseMail(item.Description || item.Name || '').items[0];
        if (parsed && parsed.name) return parsed.name;
        return item.Description || item.Name || '';
    }

    function supplierIdOf(item) {
        if (!item) return '';
        var id = item.SupplierId || item.supplierId;
        if (id && typeof id === 'object') return id.Id || id.id || '';
        if (id) return String(id);
        if (item.Supplier && typeof item.Supplier === 'object') return item.Supplier.Id || item.Supplier.id || '';
        return '';
    }

    function supplierLabel(row) {
        if (!row) return '';
        var direct = row.Name || row.CompanyName || row.SupplierName || row.DisplayName || row.Title || '';
        if (direct) return String(direct);
        var keys = Object.keys(row);
        var i;
        for (i = 0; i < keys.length; i++) {
            var value = row[keys[i]];
            if (typeof value === 'string' && value.length > 2 && value.length < 80 && /name|company|title/i.test(keys[i])) return value;
        }
        return '';
    }

    function supplierRowId(row) {
        if (!row) return '';
        return row.Id || row.SupplierId || row.id || row.SupplierID || '';
    }

    function matchSupplier(list, name) {
        var exact = null;
        var loose = null;
        (list || []).forEach(function (row) {
            [row.Name, row.CompanyName, row.SupplierName].filter(Boolean).forEach(function (label) {
                if (normName(label) === normName(name)) exact = row;
                else if (!exact && !loose && sameName(label, name)) loose = row;
            });
        });
        return exact || loose;
    }

    function supplierNameOf(item) {
        if (typeof item.Supplier === 'string') return item.Supplier;
        if (item.Supplier && typeof item.Supplier === 'object') return item.Supplier.CompanyName || item.Supplier.Name || '';
        return item.SupplierName || item.supplier || '';
    }

    function normArticle(value) {
        return String(value || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
    }

    function normName(value) {
        return String(value || '').toLowerCase().replace(/\b(gmbh|mbh|ag|kg|ohg|co)\b/g, ' ').replace(/[^a-z0-9äöüß]+/g, ' ').replace(/\s+/g, ' ').trim();
    }

    function sameName(a, b) {
        var left = normName(a);
        var right = normName(b);
        if (!left || !right) return false;
        if (left === right) return true;
        return left.length >= 4 && right.length >= 4 && (left.indexOf(right) !== -1 || right.indexOf(left) !== -1);
    }

    function matchByName(list, name, read) {
        var exact = null;
        var loose = null;
        var wanted = normName(name);
        list.forEach(function (row) {
            var label = read(row);
            if (normName(label) === wanted) exact = row;
            else if (!loose && sameName(label, name)) loose = row;
        });
        return exact || loose;
    }

    function mailItemFor(queue, article) {
        var wanted = normArticle(article);
        var found = null;
        ((queue && queue.items) || []).forEach(function (item) {
            if (normArticle(item.article) === wanted) found = item;
        });
        return found;
    }

    function vinToken(value) {
        var text = String(value || '').toUpperCase();
        var exact = text.replace(/[^A-Z0-9]/g, '');
        if (/^[A-HJ-NPR-Z0-9]{17}$/.test(exact)) return exact;
        var loose = text.match(/[A-HJ-NPR-Z0-9]{17}/);
        return loose ? loose[0] : '';
    }

    function findVin(value, depth) {
        if (!value || depth > 8) return '';
        if (typeof value === 'string') {
            if (/vin|fin|fahrgestell/i.test(value)) return vinToken(value);
            return '';
        }
        if (typeof value !== 'object') return '';
        var named = attrName(value);
        if (/vin|fin|fahrgestell/i.test(named)) {
            var fromAttr = vinToken(attrValue(value));
            if (fromAttr) return fromAttr;
        }
        var keys = Object.keys(value);
        var i;
        for (i = 0; i < keys.length; i++) {
            if (!/vin|fin|fahrgestell/i.test(keys[i])) continue;
            var direct = vinToken(value[keys[i]]);
            if (direct) return direct;
            var nested = findVin(value[keys[i]], depth + 1);
            if (nested) return nested;
        }
        var list = Array.isArray(value) ? value : keys.map(function (key) { return value[key]; });
        for (i = 0; i < list.length; i++) {
            if (value[keys[i]] && /vin|fin|fahrgestell/i.test(keys[i])) continue;
            var deep = findVin(list[i], depth + 1);
            if (deep) return deep;
        }
        return '';
    }

    function emailOf(row) {
        if (!row) return '';
        var value = row.Email || row.email || row.EMail;
        if (!value) return '';
        if (typeof value === 'string') return value;
        if (Array.isArray(value)) return emailOf({ Email: value[0] });
        return value.Address || value.Value || value.Email || '';
    }

    function phoneOf(row) {
        if (!row) return '';
        var value = row.Phone || row.phone || row.Telephone;
        if (!value) return '';
        if (typeof value === 'string') return value;
        if (Array.isArray(value)) return phoneOf({ Phone: value[0] });
        return String(value.CountryCode || '') + String(value.AreaCode || '') + String(value.ContactNumber || value.Number || '');
    }

    function addressRow(row) {
        return (row && (row.Address || row.address)) || row || {};
    }

    function addressLabel(row) {
        var address = addressRow(row);
        return [address.PostalCode, address.City, address.Address1, address.ToName].filter(Boolean).join(' · ');
    }

    function addressId(row) {
        var address = addressRow(row);
        return address.Id || address.AddressId || address.id || '';
    }

    function pickAddress(list) {
        var rows = asList(list);
        if (!rows.length && list && (list.Address || list.Id)) rows = [list];
        var preferred = rows.find(function (row) {
            var blob = addressLabel(row).toLowerCase();
            return blob.indexOf('93155') !== -1 || blob.indexOf('hemau') !== -1;
        });
        return preferred || rows[0] || null;
    }

    function returnMethod(settings) {
        var ret = settings && settings.Return;
        if (ret && ret.SupplierReturn) return { id: 'PICK_UP_BY_SUPPLIER_RETURN', bulky: null };
        var customer = ret && ret.CustomerReturn;
        var map = {
            PickupMethod: 'PICK_UP_RETURN',
            HimselfMethod: 'SEND_BACK_RETURN',
            ReturnMethod: 'BRING_IN_RETURN'
        };
        if (customer && typeof customer === 'object') {
            var order = ['PickupMethod', 'HimselfMethod', 'ReturnMethod'];
            var i;
            for (i = 0; i < order.length; i++) {
                if (customer[order[i]]) return { id: map[order[i]], bulky: false };
            }
        }
        return { id: null, bulky: false };
    }

    function caseIdOf(body, headers) {
        var fromHeader = String(headers || '').match(/R-\d+/);
        if (fromHeader) return fromHeader[0];
        if (!body) return '';
        if (typeof body === 'string') {
            var plain = body.match(/R-\d+/);
            return plain ? plain[0] : '';
        }
        if (body.CaseId) return String(body.CaseId);
        if (body.Id && /^R-\d+/.test(String(body.Id))) return String(body.Id);
        return '';
    }

    function caseHasArticle(row, article) {
        var blob = JSON.stringify(row.Items || row.items || []) + ' ' + String(row.ItemsForWRTable || '');
        return blob.indexOf(article) !== -1;
    }

    function openCase(cases, orderId, article) {
        var i;
        for (i = 0; i < cases.length; i++) {
            var row = cases[i];
            if (String(row.OrderProcessId || '') !== String(orderId)) continue;
            if (!caseHasArticle(row, article)) continue;
            var status = String(row.StatusId || row.Status || '');
            if (/REJECT|ABGELEHNT|CLOSED|STORNI/i.test(status)) continue;
            return row;
        }
        return null;
    }

    function sheetBase() {
        return String(GM_getValue(SHEET_URL_KEY, '') || '').trim();
    }

    function sheetLooksLikeAuth(status, text) {
        var body = String(text || '');
        if (status === 401 || status === 403) return true;
        if (/accounts\.google\.com|ServiceLogin|signin|oauth|Zugriffsanfrage|Request access|Melden Sie sich/i.test(body)) return true;
        if (/^\s*</.test(body) && body.indexOf('"ok"') === -1) return true;
        return false;
    }

    function pushSheet(entry) {
        var base = sheetBase();
        if (!base) return Promise.resolve({ ok: false, missing: true, error: 'Sheet-Link fehlt' });
        var params = ['action=add'];
        ['r', 'order', 'stock', 'kommission', 'article', 'name', 'qty', 'supplier', 'brand', 'vin', 'reason', 'who'].forEach(function (key) {
            var value = entry[key];
            if (key === 'who' && !value) value = 'Berger';
            params.push(key + '=' + encodeURIComponent(value == null ? '' : value));
        });
        var url = base + (base.indexOf('?') === -1 ? '?' : '&') + params.join('&');
        return new Promise(function (resolve) {
            GM_xmlhttpRequest({
                method: 'GET',
                url: url,
                timeout: 30000,
                onload: function (res) {
                    var raw = String(res.responseText || '').replace(/^\)\]\}'\n?/, '').trim();
                    var parsed = null;
                    if (raw && raw.charAt(0) !== '<') {
                        try { parsed = JSON.parse(raw); } catch (e) { parsed = null; }
                    }
                    if (sheetLooksLikeAuth(res.status, raw)) {
                        resolve({ ok: false, auth: true, error: 'Zugriff' });
                        return;
                    }
                    if (parsed && parsed.ok) resolve({ ok: true, duplicate: !!parsed.duplicate, row: parsed.row || '' });
                    else resolve({ ok: false, error: (parsed && parsed.error) || ('Antwort ' + res.status) });
                },
                onerror: function () { resolve({ ok: false, error: 'Sheet nicht erreichbar' }); },
                ontimeout: function () { resolve({ ok: false, error: 'Sheet Timeout' }); }
            });
        });
    }

    function rememberReturn(entry) {
        var list = readHistory().filter(function (row) { return row && row.r !== entry.r; });
        list.unshift(entry);
        GM_setValue(HISTORY_KEY, JSON.stringify(list.slice(0, 200)));
    }

    function readHistory() {
        try {
            var list = JSON.parse(GM_getValue(HISTORY_KEY, '[]'));
            return Array.isArray(list) ? list : [];
        } catch (e) {
            return [];
        }
    }

    function commissionFromCase(row) {
        var notes = (row && (row.Notes || row.notes)) || [];
        var text = '';
        if (Array.isArray(notes)) {
            notes.forEach(function (note) { text += ' ' + (note.Content || note.Text || note.Note || ''); });
        }
        var match = text.match(/Kommission:\s*([^\n]+)/i);
        return match ? match[1].trim() : '';
    }

    function dymoCall(method, path, data) {
        var hosts = ['https://127.0.0.1:41951', 'https://localhost:41951'];
        function once(index) {
            return new Promise(function (resolve, reject) {
                GM_xmlhttpRequest({
                    method: method,
                    url: hosts[index] + '/DYMO/DLS/Printing/' + path,
                    data: data || null,
                    headers: data ? { 'Content-Type': 'application/x-www-form-urlencoded' } : {},
                    onload: function (res) {
                        if (res.status >= 200 && res.status < 300) resolve(res.responseText || '');
                        else if (index + 1 < hosts.length) once(index + 1).then(resolve, reject);
                        else reject(new Error('DYMO ' + res.status));
                    },
                    onerror: function () {
                        if (index + 1 < hosts.length) once(index + 1).then(resolve, reject);
                        else reject(new Error('DYMO Connect nicht erreichbar'));
                    }
                });
            });
        }
        return once(0);
    }

    function dymoPrinters(xml) {
        var blocks = String(xml || '').match(/<LabelWriterPrinter>[\s\S]*?<\/LabelWriterPrinter>/g) || [];
        return blocks.map(function (block) {
            return {
                name: ((block.match(/<Name>([^<]*)<\/Name>/) || [])[1] || '').trim(),
                connected: !/<IsConnected>False<\/IsConnected>/i.test(block)
            };
        }).filter(function (printer) { return printer.name && !/\btape\b/i.test(printer.name); });
    }

    function buildRLabel(text, variant) {
        var is450 = variant === '450';
        var safe = String(text || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        var obj = '<TextObject><Name>RNumber</Name><Brushes><BackgroundBrush><SolidColorBrush><Color A="0" R="0" G="0" B="0"></Color></SolidColorBrush></BackgroundBrush><BorderBrush><SolidColorBrush><Color A="1" R="0" G="0" B="0"></Color></SolidColorBrush></BorderBrush><StrokeBrush><SolidColorBrush><Color A="1" R="0" G="0" B="0"></Color></SolidColorBrush></StrokeBrush><FillBrush><SolidColorBrush><Color A="0" R="0" G="0" B="0"></Color></SolidColorBrush></FillBrush></Brushes><Rotation>Rotation0</Rotation><OutlineThickness>1</OutlineThickness><IsOutlined>False</IsOutlined><BorderStyle>SolidLine</BorderStyle><Margin><DYMOThickness Left="0" Top="0" Right="0" Bottom="0" /></Margin><HorizontalAlignment>Center</HorizontalAlignment><VerticalAlignment>Middle</VerticalAlignment><FitMode>AlwaysFit</FitMode><IsVertical>False</IsVertical><FormattedText><FitMode>AlwaysFit</FitMode><HorizontalAlignment>Center</HorizontalAlignment><VerticalAlignment>Middle</VerticalAlignment><IsVertical>False</IsVertical><LineTextSpan><TextSpan><Text>' + safe + '</Text><FontInfo><FontName>Arial</FontName><FontSize>48</FontSize><IsBold>True</IsBold><IsItalic>False</IsItalic><IsUnderline>False</IsUnderline><FontBrush><SolidColorBrush><Color A="1" R="0" G="0" B="0"></Color></SolidColorBrush></FontBrush></FontInfo></TextSpan></LineTextSpan></FormattedText><ObjectLayout><DYMOPoint><X>0.28</X><Y>' + (is450 ? '0.16' : '0.26') + '</Y></DYMOPoint><Size><Width>3.05</Width><Height>' + (is450 ? '0.66' : '0.82') + '</Height></Size></ObjectLayout></TextObject>';
        if (is450) return '<?xml version="1.0" encoding="utf-8"?><DesktopLabel Version="1"><DYMOLabel Version="3"><Description>DYMO Label</Description><Orientation>Landscape</Orientation><LabelName>Address1982991</LabelName><InitialLength>0</InitialLength><BorderStyle>SolidLine</BorderStyle><DYMORect><DYMOPoint><X>0.23</X><Y>0.06</Y></DYMOPoint><Size><Width>3.21</Width><Height>0.9966667</Height></Size></DYMORect><BorderColor><SolidColorBrush><Color A="1" R="0" G="0" B="0"></Color></SolidColorBrush></BorderColor><BorderThickness>1</BorderThickness><Show_Border>False</Show_Border><DynamicLayoutManager><RotationBehavior>ClearObjects</RotationBehavior><LabelObjects>' + obj + '</LabelObjects></DynamicLayoutManager></DYMOLabel><LabelApplication>Blank</LabelApplication><DataTable><Columns></Columns><Rows></Rows></DataTable></DesktopLabel>';
        return '<?xml version="1.0" encoding="utf-8"?><DesktopLabel Version="1"><DYMOLabel Version="4"><Description>DYMO Label</Description><Orientation>Landscape</Orientation><LabelName>LargeAddressS0722400</LabelName><InitialLength>0</InitialLength><BorderStyle>SolidLine</BorderStyle><DYMORect><DYMOPoint><X>0.22333334</X><Y>0.060000002</Y></DYMOPoint><Size><Width>3.2033331</Width><Height>1.3066665</Height></Size></DYMORect><BorderColor><SolidColorBrush><Color A="1" R="0" G="0" B="0"></Color></SolidColorBrush></BorderColor><BorderThickness>1</BorderThickness><Show_Border>False</Show_Border><DynamicLayoutManager><RotationBehavior>ClearObjects</RotationBehavior><LabelObjects>' + obj + '</LabelObjects></DynamicLayoutManager></DYMOLabel><LabelApplication>Blank</LabelApplication><DataTable><Columns></Columns><Rows></Rows></DataTable></DesktopLabel>';
    }

    function printRLabel(caseId, copies) {
        var only = String(caseId || '').match(/R-\d+/);
        var count = Math.max(1, Math.min(30, parseInt(copies, 10) || 1));
        if (!only) return Promise.reject(new Error('Keine R-Nummer'));
        return dymoCall('GET', 'GetPrinters').then(function (xml) {
            var list = dymoPrinters(xml);
            var printer = list.filter(function (item) { return item.connected; })[0] || list[0];
            if (!printer) throw new Error('Kein DYMO Etikettendrucker');
            var variant = /450|duo/i.test(printer.name) && !/550/i.test(printer.name) ? '450' : '550';
            var params = '<?xml version="1.0" encoding="utf-8"?><LabelWriterPrintParams><Copies>' + count + '</Copies><PrintQuality>Text</PrintQuality></LabelWriterPrintParams>';
            var body = 'printerName=' + encodeURIComponent(printer.name) + '&printParamsXml=' + encodeURIComponent(params) + '&labelXml=' + encodeURIComponent(buildRLabel(only[0], variant)) + '&labelSetXml=';
            return dymoCall('POST', 'PrintLabel', body);
        });
    }

    function ensureStyle() {
        if (document.getElementById('alfah-retoure-style')) return;
        var style = document.createElement('style');
        style.id = 'alfah-retoure-style';
        style.textContent = ''
            + '#alfah-retoure-bot{position:fixed;top:36px;right:16px;z-index:2147483646;width:720px;height:90vh;min-width:480px;min-height:520px;max-width:96vw;max-height:96vh;resize:both;overflow:hidden;'
            + 'display:flex;flex-direction:column;background:#0b0f14;color:#e6edf3;border:1px solid #2b3545;border-radius:16px;'
            + 'box-shadow:0 18px 50px rgba(0,0,0,.45);font:16px/1.45 Segoe UI,Arial,sans-serif}'
            + '#alfah-retoure-bot *{box-sizing:border-box}'
            + '#alfah-retoure-bot .head{display:flex;align-items:center;justify-content:space-between;gap:8px;padding:12px 14px;cursor:move;'
            + 'background:linear-gradient(180deg,#161d27 0%,#10151d 100%);border-bottom:1px solid #2b3545;border-radius:16px 16px 0 0}'
            + '#alfah-retoure-bot .head b{font-size:18px;letter-spacing:.4px;text-transform:uppercase}'
            + '#alfah-retoure-bot .head b span{color:#00ffcc}'
            + '#alfah-retoure-bot .head button{width:32px;height:32px;border:0;border-radius:8px;background:transparent;color:#8b949e;font-weight:700;cursor:pointer;font-size:18px;line-height:32px;padding:0}'
            + '#alfah-retoure-bot .head button:hover{background:#1c2430;color:#e6edf3}'
            + '#alfah-retoure-bot .tools{display:flex;gap:4px;flex:0 0 auto}'
            + '#alfah-retoure-bot.min{height:auto !important;width:260px !important;min-width:0 !important;min-height:0 !important;resize:none}'
            + '#alfah-retoure-bot.min .body{display:none !important}'
            + '#alfah-retoure-bot.min .head{border-radius:16px}'
            + '#alfah-retoure-bot .body{padding:12px;display:flex;flex-direction:column;gap:8px;overflow-x:hidden;overflow-y:auto;flex:1;min-width:0}'
            + '#alfah-retoure-bot label{font-size:14px;font-weight:700;color:#8b949e;letter-spacing:.3px}'
            + '#alfah-retoure-bot input[type=text],#alfah-retoure-bot input[type=number],#alfah-retoure-bot input[type=file],#alfah-retoure-bot select,#alfah-retoure-bot textarea{'
            + 'width:100%;max-width:100%;min-width:0;border:1px solid #2b3545;border-radius:10px;padding:8px 10px;font:inherit;background:#141a23;color:#e6edf3}'
            + '#alfah-retoure-bot select.missing{border-color:#f85149}'
            + '#alfah-retoure-bot textarea{min-height:72px;resize:vertical}'
            + '#alfah-retoure-bot button.act{border:0;border-radius:10px;padding:8px 10px;font-weight:800;cursor:pointer;background:#00ffcc;color:#06221c}'
            + '#alfah-retoure-bot button.act:disabled{opacity:.55;cursor:wait}'
            + '#alfah-retoure-bot button.ghost{border:1px solid #2b3545;border-radius:10px;padding:7px 10px;background:#141a23;color:#e6edf3;cursor:pointer;font-weight:700}'
            + '#alfah-retoure-bot button.reset{border:0;border-radius:10px;padding:8px 10px;font-weight:800;cursor:pointer;background:#f0883e;color:#2a1404}'
            + '#alfah-retoure-bot button.reset:hover{background:#ff9d4d}'
            + '#alfah-retoure-bot .row{display:flex;gap:8px;min-width:0}'
            + '#alfah-retoure-bot .row>*{flex:1;min-width:0}'
            + '#alfah-retoure-bot .items{display:flex;flex-direction:column;gap:10px}'
            + '#alfah-retoure-bot .item{border:1px solid #2b3545;border-radius:12px;padding:10px;background:#141a23;cursor:pointer;min-width:0;overflow-wrap:anywhere}'
            + '#alfah-retoure-bot .item.on{border-color:#00ffcc;box-shadow:inset 0 0 0 1px rgba(0,255,204,.45)}'
            + '#alfah-retoure-bot .item input[type=checkbox]{accent-color:#00ffcc;margin-top:3px}'
            + '#alfah-retoure-bot .item .top{display:flex;gap:8px;align-items:flex-start}'
            + '#alfah-retoure-bot .item .meta{color:#8b949e;font-size:15px}'
            + '#alfah-retoure-bot .log{white-space:pre-wrap;background:#141a23;border:1px solid #2b3545;border-radius:10px;padding:12px;min-height:48px;font-size:16px}'
            + '#alfah-retoure-bot .result{display:flex;flex-direction:column;gap:6px;padding:8px 0;border-top:1px solid #2b3545}'
            + '#alfah-retoure-bot .result-id{color:#00ffcc;font-weight:800;font-size:22px}'
            + '#alfah-retoure-bot .ok{color:#2ea043;font-weight:700}'
            + '#alfah-retoure-bot .bad{color:#f85149;font-weight:700}'
            + '#alfah-retoure-bot .hint{color:#8b949e;font-size:15px}'
            + '#alfah-retoure-bot .badge{display:inline-flex;align-items:center;width:fit-content;max-width:100%;padding:8px 12px;border-radius:999px;background:rgba(46,160,67,.16);border:1px solid rgba(63,185,80,.7);color:#3fb950;font-size:15px;font-weight:800}'
            + '#alfah-retoure-bot .badge[hidden]{display:none}'
            + '#alfah-retoure-bot .badge.bad{background:rgba(248,81,73,.16);border-color:rgba(248,81,73,.7);color:#f85149}'
            + '#alfah-retoure-bot .foot{display:flex;justify-content:space-between;gap:10px;margin-top:8px;padding-top:10px;border-top:1px solid rgba(255,255,255,.07);color:#8b949e;font-size:12px;font-weight:800}'
            + '#alfah-retoure-bot .credit{text-align:right;margin-top:4px;color:#6e7681;font-size:10px;font-weight:600;letter-spacing:.3px}'
            + '#alfah-retoure-bot .qty{display:flex;gap:8px;align-items:center;min-width:0}'
            + '#alfah-retoure-bot .qty input{width:72px;max-width:72px;flex:0 0 72px;text-align:center;padding:8px 4px}'
            + '#alfah-retoure-bot .qty button.step{flex:0 0 40px;width:40px;border:1px solid #2b3545;border-radius:10px;background:#141a23;color:#e6edf3;font-weight:800;font-size:18px;cursor:pointer}'
            + '#alfah-retoure-bot .qty button.act{flex:1;min-width:0}'
            + '#alfah-retoure-bot .status-row{position:sticky;bottom:0;display:flex;gap:8px;align-items:stretch;background:#0b0f14;padding-top:8px;z-index:2}'
            + '#alfah-retoure-bot .status-row > button.ghost{flex:0 0 auto;padding:14px 18px;font-size:16px}'
            + '#alfah-retoure-bot .search-badge{display:none;flex:1;padding:14px 16px;border-radius:14px;background:#14342c;border:2px solid #00ffcc;color:#00ffcc;font-size:20px;font-weight:800;line-height:1.25}'
            + '#alfah-retoure-bot .search-badge.on{display:block}'
            + '#alfah-retoure-bot button.stop{flex:0 0 auto;background:#3a2216;border:2px solid #ffb020;color:#ffb020;font-size:16px;font-weight:800;padding:14px 16px;border-radius:14px;cursor:pointer}'
            + '#alfah-retoure-bot button.pick{width:100%;text-align:left;border:1px solid #2b3545;border-radius:12px;padding:14px;background:#141a23;color:#e6edf3;font-size:16px;font-weight:700;cursor:pointer}'
            + '#alfah-retoure-bot button.pick:hover{border-color:#00ffcc}'
            + '#alfah-retoure-bot .admin-pop{position:absolute;inset:0;z-index:6;display:flex;flex-direction:column;gap:10px;padding:16px;background:#0b0f14;overflow:auto}'
            + '#alfah-retoure-bot .admin-pop[hidden]{display:none !important}'
            + '#alfah-retoure-bot .admin-pop h3{margin:0;font-size:22px;color:#00ffcc}'
            + '#alfah-retoure-bot .case-list{display:flex;flex-direction:column;gap:8px;min-height:0}'
            + '#alfah-retoure-bot .case{display:flex;gap:10px;align-items:flex-start;border:1px solid #2b3545;border-radius:12px;padding:12px;background:#141a23;font-size:16px}'
            + '#alfah-retoure-bot .foot{display:flex;gap:8px}'
            + '#alfah-retoure-bot .foot>*{flex:1}';
        document.documentElement.appendChild(style);
    }

    function drag(panel, handle) {
        var state = null;
        handle.addEventListener('pointerdown', function (event) {
            if (event.button !== 0 || event.target.closest('button')) return;
            var rect = panel.getBoundingClientRect();
            state = { dx: event.clientX - rect.left, dy: event.clientY - rect.top };
            panel.style.left = rect.left + 'px';
            panel.style.top = rect.top + 'px';
            panel.style.right = 'auto';
            function move(e) {
                if (!state) return;
                panel.style.left = (e.clientX - state.dx) + 'px';
                panel.style.top = (e.clientY - state.dy) + 'px';
            }
            function up() {
                state = null;
                document.removeEventListener('pointermove', move);
            }
            document.addEventListener('pointermove', move);
            document.addEventListener('pointerup', up, { once: true });
        });
    }

    function el(tag, className, text) {
        var node = document.createElement(tag);
        if (className) node.className = className;
        if (text != null) node.textContent = text;
        return node;
    }

    function startGmail() {
        ensureStyle();
        var panel = el('div');
        panel.id = 'alfah-retoure-bot';
        panel.style.display = 'none';
        var head = el('div', 'head');
        head.appendChild(el('b', '', 'ALFAH Retoure'));
        var tools = el('div', 'tools');
        var min = el('button', '', '–');
        min.type = 'button';
        var close = el('button', '', '×');
        close.type = 'button';
        tools.appendChild(min);
        tools.appendChild(close);
        head.appendChild(tools);
        var body = el('div', 'body');
        var info = el('div', 'hint');
        var list = el('div', 'items');
        var save = el('button', 'act', 'An ALFAH übergeben');
        var status = el('div', 'log');
        body.appendChild(info);
        body.appendChild(list);
        body.appendChild(save);
        body.appendChild(status);
        var gmailFoot = el('div', 'foot');
        gmailFoot.appendChild(el('span', '', 'Bot v' + BOT_VERSION));
        body.appendChild(gmailFoot);
        body.appendChild(el('div', 'credit', 'by Arasaka'));
        panel.appendChild(head);
        panel.appendChild(body);
        document.documentElement.appendChild(panel);
        drag(panel, head);
        min.addEventListener('click', function (event) {
            event.preventDefault();
            event.stopPropagation();
            var closed = panel.classList.toggle('min');
            min.textContent = closed ? '+' : '–';
        });
        close.addEventListener('click', function () {
            panel.style.display = 'none';
        });

        var currentKey = '';

        function mailText() {
            var subject = document.querySelector('h2.hP');
            var message = document.querySelector('div.a3s.aiL') || document.querySelector('div.a3s');
            return ((subject && subject.innerText) || '') + '\n' + ((message && message.innerText) || '');
        }

        function render() {
            var parsed = parseMail(mailText());
            if (!parsed.orderId || parsed.items.length === 0) {
                panel.style.display = 'none';
                currentKey = '';
                return;
            }
            var key = parsed.orderId + ':' + parsed.items.map(function (item) { return item.article; }).join(',');
            panel.style.display = 'block';
            if (key === currentKey) return;
            currentKey = key;
            info.textContent = 'Bestellung ' + parsed.orderId + (parsed.stockId ? ' · ' + parsed.stockId : '');
            list.textContent = '';
            parsed.items.forEach(function (item) {
                var row = el('div', 'item on');
                var top = el('div', 'top');
                var box = document.createElement('input');
                box.type = 'checkbox';
                box.checked = true;
                box.dataset.article = item.article;
                var text = el('div');
                text.appendChild(el('div', '', item.article + ' · ' + item.name));
                text.appendChild(el('div', 'meta', 'Lieferant: ' + item.supplier + ' · Hersteller: ' + item.brand + (item.vin ? ' · FIN ' + item.vin : '')));
                top.appendChild(box);
                top.appendChild(text);
                row.appendChild(top);
                row.addEventListener('click', function (event) {
                    if (event.target === box) return;
                    box.checked = !box.checked;
                    row.classList.toggle('on', box.checked);
                });
                box.addEventListener('change', function () { row.classList.toggle('on', box.checked); });
                list.appendChild(row);
            });
            status.textContent = 'Haken bei den Teilen, die retour gehen. Danach auf alfah.de einloggen und anlegen.';
            save.onclick = function () {
                var picked = [];
                list.querySelectorAll('input[type=checkbox]:checked').forEach(function (box) {
                    var item = parsed.items.find(function (row) { return row.article === box.dataset.article; });
                    if (item) picked.push(item);
                });
                if (!picked.length) {
                    status.textContent = 'Kein Teil ausgewählt.';
                    return;
                }
                saveQueue({
                    orderId: parsed.orderId,
                    kommission: parsed.kommission,
                    stockId: parsed.stockId,
                    items: picked,
                    savedAt: Date.now()
                });
                status.textContent = picked.length + ' Teil(e) gespeichert. Jetzt alfah.de öffnen, eingeloggt bleiben, und dort auf Retouren anlegen.';
            };
        }

        render();
        var timer = null;
        new MutationObserver(function () {
            clearTimeout(timer);
            timer = setTimeout(render, 400);
        }).observe(document.documentElement, { childList: true, subtree: true });
    }

    function startAlfah() {
        ensureStyle();
        var state = {
            order: null,
            queue: loadQueue(),
            addresses: [],
            reasons: [],
            settings: [],
            purposes: [],
            cases: [],
            makerCache: {},
            supplierList: null,
            busy: false,
            stopSearch: false,
            pickFinish: null,
            sheetAuthOpened: false,
            adminOn: false
        };

        var panel = el('div');
        panel.id = 'alfah-retoure-bot';
        var head = el('div', 'head');
        var title = el('b');
        title.appendChild(document.createTextNode('ALFAH '));
        title.appendChild(el('span', '', 'Retoure v' + BOT_VERSION));
        head.appendChild(title);
        var min = el('button', '', '–');
        min.type = 'button';
        head.appendChild(min);
        var body = el('div', 'body');
        var login = el('div', 'hint');
        var orderInput = document.createElement('input');
        orderInput.type = 'text';
        orderInput.placeholder = '1333502 oder BG50729';
        var loadBtn = el('button', 'act', 'Auftrag laden');
        var paste = document.createElement('textarea');
        paste.placeholder = 'Oder die Auftragsmail hier reinpasten';
        var readBtn = el('button', 'ghost', 'Mail lesen');
        var meta = el('div', 'hint');
        var items = el('div', 'items');
        var addressSelect = document.createElement('select');
        var reasonSelect = document.createElement('select');
        var problem = document.createElement('input');
        problem.type = 'text';
        problem.value = PROBLEM_DEFAULT;
        var file = document.createElement('input');
        file.type = 'file';
        file.accept = 'application/pdf,.pdf';
        var again = document.createElement('input');
        again.type = 'checkbox';
        var againLabel = el('label', 'hint');
        againLabel.appendChild(again);
        againLabel.appendChild(document.createTextNode(' Neu anlegen, auch wenn die Retoure schon existiert'));
        var sendBtn = el('button', 'act', 'Retouren anlegen');
        var startLog = 'Bestellnummer laden. Pro Retoure entsteht eine eigene Vorgangsnummer.';
        var log = el('div', 'log', startLog);
        var resetBtn = el('button', 'reset', 'Zurücksetzen');
        resetBtn.type = 'button';
        body.appendChild(resetBtn);
        body.appendChild(login);
        body.appendChild(el('label', '', 'Dein Name'));
        var workerInput = document.createElement('input');
        workerInput.type = 'text';
        workerInput.placeholder = 'z.B. Francesco';
        workerInput.value = workerName();
        body.appendChild(workerInput);
        var sheetHint = el('div', 'hint', '');
        body.appendChild(sheetHint);
        var sheetUrlLabel = el('label', '', 'Sheet Web-App');
        var sheetUrlInput = document.createElement('input');
        sheetUrlInput.type = 'text';
        sheetUrlInput.placeholder = 'https://script.google.com/.../exec';
        sheetUrlInput.value = sheetUrl();
        sheetUrlLabel.hidden = !!sheetUrl();
        sheetUrlInput.hidden = !!sheetUrl();
        body.appendChild(sheetUrlLabel);
        body.appendChild(sheetUrlInput);
        body.appendChild(el('label', '', 'Bestellung oder Kommission'));
        var orderRow = el('div', 'row');
        orderRow.appendChild(orderInput);
        orderRow.appendChild(loadBtn);
        body.appendChild(orderRow);
        body.appendChild(paste);
        body.appendChild(readBtn);
        body.appendChild(meta);
        body.appendChild(items);
        body.appendChild(el('label', '', 'Adresse'));
        body.appendChild(addressSelect);
        body.appendChild(el('label', '', 'Grund'));
        body.appendChild(reasonSelect);
        body.appendChild(el('label', '', 'Problem'));
        body.appendChild(problem);
        body.appendChild(el('label', '', 'Lieferschein'));
        var slipBadge = el('div', 'badge', 'Lieferschein gefunden, wird automatisch hinzugefügt');
        slipBadge.hidden = true;
        body.appendChild(slipBadge);
        body.appendChild(el('div', 'hint', 'Leer lassen. Das PDF kommt aus der ALFAH-Bestellung, dieselbe Datei wie der Anhang in der Mail. Nur eine Datei wählen, wenn ein anderes PDF rauf soll.'));
        body.appendChild(file);
        body.appendChild(againLabel);
        body.appendChild(sendBtn);
        body.appendChild(log);
        var statusRow = el('div', 'status-row');
        var searchBadge = el('div', 'search-badge');
        var stopBtn = el('button', 'stop', 'Suche stoppen');
        stopBtn.type = 'button';
        stopBtn.hidden = true;
        var adminBtn = el('button', 'ghost', 'Admin');
        adminBtn.type = 'button';
        statusRow.appendChild(searchBadge);
        statusRow.appendChild(stopBtn);
        statusRow.appendChild(adminBtn);
        body.appendChild(statusRow);
        var adminPop = el('div', 'admin-pop');
        adminPop.hidden = true;
        var foot = el('div', 'foot');
        foot.appendChild(el('span', '', 'Bot v' + BOT_VERSION));
        body.appendChild(foot);
        body.appendChild(el('div', 'credit', 'by Arasaka'));
        panel.appendChild(head);
        panel.appendChild(body);
        panel.appendChild(adminPop);
        document.documentElement.appendChild(panel);
        drag(panel, head);
        min.addEventListener('click', function (event) {
            event.preventDefault();
            event.stopPropagation();
            var closed = panel.classList.toggle('min');
            min.textContent = closed ? '+' : '–';
        });

        function setLog(text) {
            log.textContent = text;
        }

        function showProgress(text) {
            searchBadge.textContent = text;
            searchBadge.classList.add('on');
            stopBtn.hidden = false;
        }

        function hideProgress() {
            searchBadge.textContent = '';
            searchBadge.classList.remove('on');
            stopBtn.hidden = true;
        }

        function labelCount(input) {
            var count = parseInt(input.value, 10);
            if (!count || count < 1) count = 1;
            if (count > 30) count = 30;
            input.value = String(count);
            return count;
        }

        function showResults(entries, notes, sheetIds) {
            log.textContent = '';
            entries.forEach(function (entry) {
                var row = el('div', 'result');
                row.appendChild(el('div', 'result-id', entry.caseId || 'ohne R-Nummer'));
                row.appendChild(el('div', 'meta', entry.article + ' · ' + entry.supplier));
                if (entry.sheet) row.appendChild(el('div', entry.sheetOk ? 'ok' : 'bad', entry.sheet));
                if (entry.sheet && !entry.sheetOk && /Zugriff/.test(entry.sheet)) {
                    var allow = el('button', 'act', 'Sheet-Zugriff erlauben');
                    allow.type = 'button';
                    allow.addEventListener('click', function () {
                        var url = sheetBase();
                        if (!url) {
                            setLog('Sheet-Link fehlt. Admin öffnen und den Web-App-Link eintragen.');
                            return;
                        }
                        window.open(url, '_blank');
                    });
                    row.appendChild(allow);
                }
                if (/^R-\d+$/.test(entry.caseId || '')) {
                    var sheetState = sheetIds && sheetIds[entry.caseId];
                    if (sheetState === 'ok') row.appendChild(el('div', 'badge', 'Erfolgreich ins Sheet übertragen'));
                    else if (sheetState === 'bad') row.appendChild(el('div', 'badge bad', 'Nicht ins Sheet gekommen'));
                    var qty = el('div', 'qty');
                    var minus = el('button', 'step', '−');
                    minus.type = 'button';
                    var count = document.createElement('input');
                    count.type = 'number';
                    count.min = '1';
                    count.max = '30';
                    count.value = '1';
                    var plus = el('button', 'step', '+');
                    plus.type = 'button';
                    var btn = el('button', 'act', '1 Etikett drucken');
                    function paint() {
                        var n = labelCount(count);
                        btn.textContent = n === 1 ? '1 Etikett drucken' : n + ' Etiketten drucken';
                    }
                    minus.addEventListener('click', function () {
                        count.value = String(Math.max(1, labelCount(count) - 1));
                        paint();
                    });
                    plus.addEventListener('click', function () {
                        count.value = String(Math.min(30, labelCount(count) + 1));
                        paint();
                    });
                    count.addEventListener('input', paint);
                    btn.addEventListener('click', function () {
                        var n = labelCount(count);
                        btn.disabled = true;
                        printRLabel(entry.caseId, n).then(function () {
                            btn.textContent = n === 1 ? entry.caseId + ' gedruckt' : entry.caseId + ' × ' + n + ' gedruckt';
                            btn.disabled = false;
                        }).catch(function (error) {
                            btn.textContent = error.message || 'DYMO Fehler';
                            btn.disabled = false;
                        });
                    });
                    qty.appendChild(minus);
                    qty.appendChild(count);
                    qty.appendChild(plus);
                    qty.appendChild(btn);
                    row.appendChild(qty);
                    var copy = el('button', 'ghost', 'R-Nummer kopieren');
                    copy.type = 'button';
                    copy.addEventListener('click', function () {
                        var id = entry.caseId;
                        if (!navigator.clipboard || !navigator.clipboard.writeText) return;
                        navigator.clipboard.writeText(id).then(function () {
                            copy.textContent = id + ' kopiert';
                        });
                    });
                    row.appendChild(copy);
                }
                log.appendChild(row);
            });
            (notes || []).forEach(function (line) {
                log.appendChild(el('div', 'bad', line));
            });
        }

        function refreshSlipBadge() {
            var custom = file.files && file.files[0];
            slipBadge.hidden = !(state.order && orderInput.value.trim() && !custom);
        }

        function clearTop() {
            orderInput.value = '';
            paste.value = '';
            file.value = '';
            items.textContent = '';
            meta.textContent = '';
            state.order = null;
            state.queue = null;
            saveQueue(null);
            refreshSlipBadge();
        }

        function resetHud() {
            clearTop();
            again.checked = false;
            problem.value = PROBLEM_DEFAULT;
            fillAddresses();
            fillReasons();
            setLog(startLog);
        }

        function selectedRows() {
            return Array.prototype.slice.call(items.querySelectorAll('.item')).filter(function (row) {
                return row.querySelector('input[type=checkbox]').checked;
            });
        }

        function renderItems(order) {
            items.textContent = '';
            var rows = orderItems(order);
            if (!rows.length) {
                meta.textContent = 'Auftrag geladen, aber ohne Positionen.';
                return;
            }
            var supplier = supplierNameOf(rows[0]);
            meta.textContent = 'Bestellung ' + orderInput.value.trim() + (supplier ? ' · ' + supplier : '') + ' · ' + rows.length + ' Position(en)';
            rows.forEach(function (item) {
                var number = articleNo(item);
                var fromMail = mailItemFor(state.queue, number);
                var supplierLabel = (fromMail && fromMail.supplier) || supplierNameOf(item);
                var brandLabel = (fromMail && fromMail.brand) || brandOf(item);
                var card = el('div', 'item');
                card.dataset.article = number;
                card._item = item;
                card._wantedSupplier = supplierLabel;
                var top = el('div', 'top');
                var box = document.createElement('input');
                box.type = 'checkbox';
                box.checked = fromMail ? true : rows.length === 1;
                var text = el('div');
                text.appendChild(el('div', '', (number || '—') + ' · ' + (partName(item) || (fromMail && fromMail.name) || '')));
                text.appendChild(el('div', 'meta', 'Hersteller: ' + (brandLabel || '—')));
                top.appendChild(box);
                top.appendChild(text);
                var supplierSelect = document.createElement('select');
                var brand = document.createElement('input');
                brand.type = 'text';
                brand.placeholder = 'Hersteller';
                brand.value = brandLabel;
                var vin = document.createElement('input');
                vin.type = 'text';
                vin.placeholder = 'FIN aus der Mail';
                vin.value = (fromMail && fromMail.vin) || item.Vin || findVin(item, 0) || findVin(order, 0) || '';
                var qty = document.createElement('input');
                qty.type = 'number';
                qty.min = '1';
                qty.value = String((fromMail && fromMail.qty) || item.Quantity || 1);
                var replace = document.createElement('input');
                replace.type = 'checkbox';
                var replaceLabel = el('label', 'hint');
                replaceLabel.appendChild(replace);
                replaceLabel.appendChild(document.createTextNode(' Ersatzlieferung'));
                card.appendChild(top);
                card.appendChild(supplierSelect);
                card.appendChild(brand);
                card.appendChild(vin);
                card.appendChild(qty);
                card.appendChild(replaceLabel);
                card._brand = brand;
                card._vin = vin;
                card._qty = qty;
                card._replace = replace;
                card._supplier = supplierSelect;
                supplierSelect.addEventListener('change', function () {
                    supplierSelect.classList.toggle('missing', !supplierSelect.value);
                });
                card.classList.toggle('on', box.checked);
                card.addEventListener('click', function (event) {
                    if (event.target.closest('input, select, button, label')) return;
                    box.checked = !box.checked;
                    card.classList.toggle('on', box.checked);
                });
                box.addEventListener('change', function () { card.classList.toggle('on', box.checked); });
                items.appendChild(card);
            });
            suppliers().then(function (list) {
                items.querySelectorAll('.item').forEach(function (card) {
                    fillSupplierSelect(card._supplier, card._wantedSupplier, supplierIdOf(card._item), list);
                });
            }).catch(function (error) {
                setLog('Lieferantenliste nicht geladen: ' + (error.message || error));
            });
        }

        function fillSupplierSelect(select, wantedName, knownId, list) {
            if (!select) return;
            var current = select.value;
            select.textContent = '';
            var blank = el('option', '', 'Lieferant wählen');
            blank.value = '';
            select.appendChild(blank);
            var rows = (list || []).slice().sort(function (a, b) {
                return supplierLabel(a).localeCompare(supplierLabel(b), 'de');
            });
            rows.forEach(function (row) {
                var id = supplierRowId(row);
                var name = supplierLabel(row);
                if (!id || !name) return;
                var option = el('option', '', name);
                option.value = String(id);
                select.appendChild(option);
            });
            var picked = '';
            if (knownId) {
                Array.prototype.forEach.call(select.options, function (option) {
                    if (option.value === String(knownId)) picked = option.value;
                });
            }
            if (!picked && wantedName) {
                var hit = matchSupplier(rows, wantedName);
                if (hit) picked = String(supplierRowId(hit));
            }
            if (!picked && current) picked = current;
            select.value = picked || '';
            select.classList.toggle('missing', !select.value);
        }

        function fillReasons() {
            reasonSelect.textContent = '';
            state.reasons.forEach(function (reason) {
                var option = el('option', '', reason.Name || reason.Id);
                option.value = reason.Id;
                reasonSelect.appendChild(option);
            });
            var wanted = state.reasons.find(function (reason) {
                return reason.Id === REASON_WANTED || /nicht ben[oö]tigt/i.test(reason.Name || '');
            });
            if (wanted) reasonSelect.value = wanted.Id;
        }

        function fillAddresses() {
            addressSelect.textContent = '';
            state.addresses.forEach(function (row, index) {
                var option = el('option', '', addressLabel(row) || ('Adresse ' + (index + 1)));
                option.value = String(index);
                addressSelect.appendChild(option);
            });
            var preferred = pickAddress(state.addresses);
            var index = state.addresses.indexOf(preferred);
            if (index >= 0) addressSelect.value = String(index);
        }

        async function prepare() {
            if (!token()) {
                login.textContent = 'Bitte bei ALFAH einloggen und die Seite neu laden.';
                login.className = 'bad';
                return;
            }
            login.textContent = 'Eingeloggt als ' + (customerName() || customerId() || 'Kunde');
            login.className = 'hint';
            try {
                var packed = await Promise.all([
                    api('GET', 'api/user/details/address?purpose_type_id=SHIPPING'),
                    api('GET', 'api/wr/cases/details/reasons?case-type-id=RETURN'),
                    api('GET', 'api/wr/supplier-settings'),
                    api('GET', 'api/wr/cases/details/attachment-purposes'),
                    api('GET', 'api/wr/cases?include-items=true')
                ]);
                state.addresses = asList(packed[0]);
                if (!state.addresses.length && packed[0]) state.addresses = [packed[0]];
                state.reasons = asList(packed[1]);
                state.settings = asList(packed[2]);
                state.purposes = asList(packed[3]);
                state.cases = asList(packed[4]);
                fillAddresses();
                fillReasons();
            } catch (error) {
                setLog(error.message || String(error));
            }
        }

        function orderQuery(raw) {
            return String(raw || '').trim().split(/[\s\/,]+/)[0];
        }

        function orderStamp(order) {
            var raw = order && (order.OrderDate || order.CreatedAt || order.Date || '');
            var time = Date.parse(raw);
            return isNaN(time) ? 0 : time;
        }

        function dayStamp(date) {
            var y = date.getFullYear();
            var m = String(date.getMonth() + 1).padStart(2, '0');
            var d = String(date.getDate()).padStart(2, '0');
            return y + '-' + m + '-' + d;
        }

        function orderMatches(order, query) {
            var up = String(query || '').toUpperCase();
            if (!up || !order) return false;
            if (String(orderIdOf(order)).toUpperCase() === up) return true;
            var stock = stockIdFromCommission(commissionOf(order));
            if (stock && stock === up) return true;
            var blob = '';
            try {
                blob = JSON.stringify(order);
            } catch (e) {
                return false;
            }
            return new RegExp('(^|[^A-Za-z0-9])' + up.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '([^A-Za-z0-9]|$)', 'i').test(blob);
        }

        function findOrder(list, query) {
            var hits = (list || []).filter(function (order) { return orderMatches(order, query); });
            hits.sort(function (a, b) {
                var items = (orderItems(b).length ? 1 : 0) - (orderItems(a).length ? 1 : 0);
                if (items) return items;
                return orderStamp(b) - orderStamp(a);
            });
            return hits[0] || null;
        }

        async function ordersBetween(from, thru) {
            var path = 'api/ordering/orders?include-order-attributes=true&limit=400&from=' + dayStamp(from) + '&thru=' + dayStamp(thru);
            var data = await api('GET', path);
            var list = asList(data);
            if (!list.length && data && typeof data === 'object') list = data.Orders || data.orders || data.Data || data.Results || [];
            return Array.isArray(list) ? list : [];
        }

        function emptyOrderCache() {
            return { ranges: {}, orders: {} };
        }

        function readOrderCache() {
            if (state.orderCache) return state.orderCache;
            try {
                var data = JSON.parse(GM_getValue(ORDER_CACHE_KEY, '') || 'null');
                state.orderCache = data && data.orders && data.ranges ? data : emptyOrderCache();
            } catch (e) {
                state.orderCache = emptyOrderCache();
            }
            return state.orderCache;
        }

        function writeOrderCache(cache) {
            state.orderCache = cache;
            try {
                GM_setValue(ORDER_CACHE_KEY, JSON.stringify(cache));
            } catch (e) {}
        }

        function rangeKey(from, thru) {
            return dayStamp(from) + ':' + dayStamp(thru);
        }

        function shiftDays(date, days) {
            var next = new Date(date.getFullYear(), date.getMonth(), date.getDate());
            next.setDate(next.getDate() + days);
            return next;
        }

        function bucketStart(date) {
            var epoch = new Date(2024, 0, 1);
            var day = new Date(date.getFullYear(), date.getMonth(), date.getDate());
            var days = Math.round((day.getTime() - epoch.getTime()) / 86400000);
            return shiftDays(epoch, Math.floor(days / 14) * 14);
        }

        function orderSearchText(order) {
            var parts = [orderIdOf(order), commissionOf(order)];
            var attrs = order.OrderAttributes || order.orderAttributes || order.Attributes || [];
            if (Array.isArray(attrs)) attrs.forEach(function (row) { parts.push(attrValue(row)); });
            orderItems(order).forEach(function (item) {
                parts.push(articleNo(item), item.Description || '', item.Name || '');
                var nested = item.Attributes || item.OrderItemAttributes || [];
                if (Array.isArray(nested)) nested.forEach(function (row) { parts.push(attrName(row), attrValue(row)); });
            });
            return parts.filter(Boolean).join(' ').slice(0, 2000);
        }

        function rememberOrders(cache, from, thru, list) {
            var key = rangeKey(from, thru);
            cache.ranges[key] = { savedAt: Date.now(), full: list.length >= 400 };
            list.forEach(function (order) {
                var id = orderIdOf(order);
                if (!id) return;
                var prev = cache.orders[id];
                var stock = stockIdFromCommission(commissionOf(order));
                if (prev && prev.enriched && !stock) return;
                cache.orders[id] = {
                    id: id,
                    date: order.OrderDate || order.CreatedAt || order.Date || (prev && prev.date) || '',
                    stock: stock || (prev && prev.stock) || '',
                    text: stock ? orderSearchText(order) : ((prev && prev.text) || orderSearchText(order)),
                    enriched: !!stock || !!(prev && prev.enriched)
                };
            });
            writeOrderCache(cache);
        }

        function rangeCovered(cache, from, thru) {
            var entry = cache.ranges[rangeKey(from, thru)];
            if (!entry || entry.full) return false;
            var today = dayStamp(new Date());
            if (dayStamp(thru) >= today && Date.now() - entry.savedAt > 3 * 3600 * 1000) return false;
            return true;
        }

        function cachedHits(cache, query) {
            var up = String(query || '').toUpperCase();
            var tokenRe = new RegExp('(^|[^A-Za-z0-9])' + up.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '([^A-Za-z0-9]|$)', 'i');
            var exact = [];
            var loose = [];
            Object.keys(cache.orders || {}).forEach(function (id) {
                var row = cache.orders[id];
                if (!row) return;
                if (String(row.id).toUpperCase() === up || String(row.stock || '').toUpperCase() === up) exact.push(row);
                else if (tokenRe.test(row.text || '')) loose.push(row);
            });
            loose.sort(function (a, b) { return Date.parse(b.date || 0) - Date.parse(a.date || 0); });
            return exact.length ? exact : loose.slice(0, 12);
        }

        async function orderById(orderId) {
            return api('GET', 'api/ordering/orders/' + encodeURIComponent(orderId) + '?include-order-items=true&include-invoice-ids=true&include-shipment-info=true&include-order-attributes=true&include-order-item-attributes=true');
        }

        async function orderAttrs(orderId) {
            return api('GET', 'api/ordering/orders/' + encodeURIComponent(orderId) + '?include-order-attributes=true');
        }

        async function loadRange(cache, from, thru) {
            if (rangeCovered(cache, from, thru)) return [];
            if (state.stopSearch) return [];
            showProgress('Suche ' + dayStamp(from) + ' bis ' + dayStamp(thru));
            var list = await ordersBetween(from, thru);
            if (list.length >= 400) {
                var mid = new Date(from.getTime() + (thru.getTime() - from.getTime()) / 2);
                var newerFrom = shiftDays(mid, 1);
                var older = await loadRange(cache, from, mid);
                var newer = await loadRange(cache, newerFrom, thru);
                list = older.concat(newer);
            }
            rememberOrders(cache, from, thru, list);
            var saved = cache.ranges[rangeKey(from, thru)];
            if (saved) saved.full = false;
            writeOrderCache(cache);
            return list;
        }

        function listCarriesCommission(list) {
            var i;
            for (i = 0; i < list.length; i++) {
                if (commissionOf(list[i])) return true;
            }
            return false;
        }

        function ordersInSpan(cache, from, thru) {
            var start = dayStamp(from);
            var end = dayStamp(thru);
            return Object.keys(cache.orders).map(function (id) {
                return cache.orders[id];
            }).filter(function (row) {
                var day = String(row.date || '').slice(0, 10);
                return day && day >= start && day <= end;
            });
        }

        async function openCached(cache, hit) {
            setLog('Kommission ist gespeichert. Lade Bestellung ' + hit.id + ' …');
            try {
                return await orderById(hit.id);
            } catch (error) {
                delete cache.orders[hit.id];
                writeOrderCache(cache);
                return null;
            }
        }

        async function enrichForQuery(cache, query, rows) {
            var up = String(query || '').toUpperCase();
            var matches = [];
            var pending = [];
            (rows || []).forEach(function (row) {
                if (!row || !row.id) return;
                if (String(row.stock || '').toUpperCase() === up) matches.push(row.id);
                else if (!row.enriched) pending.push(row);
            });
            pending.sort(function (a, b) { return Date.parse(b.date || 0) - Date.parse(a.date || 0); });
            var size = 12;
            var done = 0;
            var index = 0;
            while (index < pending.length && !state.stopSearch) {
                var chunk = pending.slice(index, index + size);
                index += chunk.length;
                showProgress('Suche ' + query + ' · ' + (done + 1) + '–' + (done + chunk.length) + ' von ' + pending.length);
                var details = await Promise.all(chunk.map(function (row) {
                    return orderAttrs(row.id).then(function (order) {
                        return { row: row, order: order };
                    }).catch(function () {
                        return { row: row, order: null };
                    });
                }));
                var i;
                for (i = 0; i < details.length; i++) {
                    done += 1;
                    var order = details[i].order;
                    var id = order ? (orderIdOf(order) || details[i].row.id) : details[i].row.id;
                    var stock = order ? stockIdFromCommission(commissionOf(order)) : '';
                    cache.orders[id] = {
                        id: id,
                        date: (order && (order.OrderDate || order.CreatedAt || order.Date)) || details[i].row.date || '',
                        stock: stock,
                        text: order ? orderSearchText(order) : (details[i].row.text || ''),
                        enriched: true
                    };
                    if (String(stock || '').toUpperCase() === up || (order && orderMatches(order, query))) matches.push(id);
                }
                writeOrderCache(cache);
            }
            return matches;
        }

        async function scanSpan(cache, query, from, thru) {
            if (state.stopSearch) return [];
            var list = [];
            if (!rangeCovered(cache, from, thru)) {
                list = await loadRange(cache, from, thru);
                var directs = (list || []).filter(function (order) { return orderMatches(order, query); });
                if (directs.length) {
                    return directs.map(function (order) { return orderIdOf(order); }).filter(Boolean);
                }
                if (listCarriesCommission(list)) return [];
            }
            var rows = [];
            if (list.length) {
                list.forEach(function (order) {
                    var id = orderIdOf(order);
                    if (id && cache.orders[id]) rows.push(cache.orders[id]);
                });
            } else {
                rows = ordersInSpan(cache, from, thru);
            }
            return enrichForQuery(cache, query, rows);
        }

        async function openMatch(cache, id) {
            showProgress('Lade Bestellung ' + id + ' …');
            try {
                return await orderById(id);
            } catch (error) {
                delete cache.orders[id];
                writeOrderCache(cache);
                return null;
            }
        }

        async function findStockOrders(query) {
            var cache = readOrderCache();
            var known = Object.keys(cache.orders).length;
            if (known) showProgress('Prüfe ' + known + ' gespeicherte Bestellungen · ' + query);
            var ids = [];
            cachedHits(cache, query).forEach(function (row) {
                if (row.id && ids.indexOf(row.id) === -1) ids.push(row.id);
            });
            if (!ids.length) {
                var today = new Date();
                var monthFrom = new Date(today.getFullYear(), today.getMonth(), 1);
                showProgress('Suche ' + query + ' im aktuellen Monat …');
                ids = await scanSpan(cache, query, monthFrom, today);
                var thru = shiftDays(monthFrom, -1);
                var guard = new Date();
                guard.setMonth(guard.getMonth() - 18);
                var steps = 0;
                while (!ids.length && !state.stopSearch && thru > guard && steps < 40) {
                    var from = bucketStart(thru);
                    if (from < guard) from = new Date(guard.getFullYear(), guard.getMonth(), guard.getDate());
                    steps += 1;
                    showProgress('Suche ' + query + ' · ' + dayStamp(from) + ' bis ' + dayStamp(thru));
                    ids = await scanSpan(cache, query, from, thru);
                    thru = shiftDays(from, -1);
                }
            }
            var unique = [];
            ids.forEach(function (id) {
                if (id && unique.indexOf(String(id)) === -1) unique.push(String(id));
            });
            var orders = [];
            var i;
            for (i = 0; i < unique.length && !state.stopSearch; i++) {
                var opened = await openMatch(cache, unique[i]);
                if (opened) orders.push(opened);
            }
            return orders;
        }

        function mailSource(pasted, orderId, query) {
            var sources = [];
            if (pasted && pasted.items && pasted.items.length) sources.push(pasted);
            if (state.queue && state.queue.items && state.queue.items.length) sources.push(state.queue);
            var i;
            for (i = 0; i < sources.length; i++) {
                var row = sources[i];
                var sameOrder = row.orderId && orderId && String(row.orderId) === String(orderId);
                var sameStock = row.stockId && query && String(row.stockId).toUpperCase() === String(query).toUpperCase();
                if (sameOrder || sameStock) return row.items;
            }
            return [];
        }

        async function loadOrder() {
            var query = orderQuery(orderInput.value);
            if (!query) {
                setLog('Bestellung oder Kommission eingeben, zum Beispiel 1333502 oder BG50729.');
                return;
            }
            if (!token()) {
                setLog('Nicht eingeloggt.');
                return;
            }
            state.stopSearch = false;
            state.pickFinish = null;
            loadBtn.disabled = true;
            showProgress('Lade ' + query + ' …');
            try {
                var pasted = parseMail(paste.value);
                if (pasted.orderId || pasted.items.length) {
                    state.queue = {
                        orderId: pasted.orderId,
                        kommission: pasted.kommission,
                        stockId: pasted.stockId,
                        items: pasted.items,
                        savedAt: Date.now()
                    };
                    if (!/^\d{5,}$/.test(query) && pasted.orderId) query = pasted.orderId;
                    if (/^[A-Za-z]/.test(query) && pasted.stockId && query.toUpperCase() !== pasted.stockId && pasted.orderId) query = pasted.orderId;
                }
                var order = null;
                if (/^\d{5,}$/.test(query)) {
                    try {
                        order = await orderById(query);
                    } catch (error) {
                        order = null;
                    }
                }
                if (/^[A-Za-z]/.test(query)) {
                    var foundOrders = await findStockOrders(query);
                    if (foundOrders.length > 1) {
                        hideProgress();
                        order = await pickOrder(foundOrders);
                        if (!order) {
                            setLog(state.stopSearch ? 'Suche gestoppt.' : 'Keine Bestellung ausgewählt.');
                            meta.textContent = '';
                            return;
                        }
                    } else if (foundOrders.length === 1) order = foundOrders[0];
                    else if (state.stopSearch) {
                        setLog('Suche gestoppt.');
                        meta.textContent = '';
                        return;
                    }
                } else if (!order || !orderItems(order).length) {
                    showProgress('Suche ' + query + ' in den letzten Bestellungen …');
                    var thru = new Date();
                    var from = new Date();
                    from.setDate(from.getDate() - 120);
                    var found = findOrder(await ordersBetween(from, thru), query);
                    if (found) {
                        try {
                            order = await orderById(orderIdOf(found, query));
                        } catch (e) {
                            order = found;
                        }
                    }
                }
                if (!order) {
                    setLog('Kein Auftrag zu ' + query + '. Die Mailbox wird nicht gelesen. Bestellnummer aus der Auftragsbestätigung eintragen oder die Mail hier reinkopieren.');
                    meta.textContent = '';
                    return;
                }
                var orderId = orderIdOf(order, /^\d{5,}$/.test(query) ? query : '');
                if (orderId) orderInput.value = orderId;
                var parts = mailSource(pasted, orderId, query);
                if (!parts.length && pasted.items.length) {
                    state.queue.orderId = state.queue.orderId || orderId;
                    parts = mailSource(pasted, orderId || pasted.orderId, query);
                }
                if (parts.length) order._alfahMailRows = rowsFromMail(parts, orderItems(order));
                state.order = order;
                renderItems(order);
                refreshSlipBadge();
                var stock = stockIdFromCommission(commissionOf(order)) || (state.queue && state.queue.stockId) || '';
                var count = orderItems(order).length;
                if (!count) {
                    var keys = typeof order === 'object' ? Object.keys(order).slice(0, 25).join(', ') : typeof order;
                    console.log('[ALFAH Retoure] order without items', order);
                    setLog('Auftrag ' + (orderId || query) + ' hat keine Positionen. Felder: ' + keys);
                    return;
                }
                var missingVin = Array.prototype.some.call(items.querySelectorAll('.item'), function (card) {
                    return card._vin && !card._vin.value.trim();
                });
                setLog('Bestellung ' + (orderId || query) + (stock ? ' · Kommission ' + stock : '') + ' · ' + count + ' Position(en).' + (missingVin ? ' FIN fehlt bei mindestens einem Teil.' : ' FIN steht an jedem Teil.'));
            } catch (error) {
                setLog(error.message || String(error));
            } finally {
                hideProgress();
                loadBtn.disabled = false;
            }
        }

        function orderChoiceLabel(order) {
            var id = orderIdOf(order);
            var date = String(order.OrderDate || order.CreatedAt || order.Date || '').replace('T', ' ').slice(0, 16);
            var stock = stockIdFromCommission(commissionOf(order));
            var rows = orderItems(order);
            var supplier = rows.length ? supplierNameOf(rows[0]) : '';
            return [id, date, stock, supplier, rows.length ? rows.length + ' Positionen' : ''].filter(Boolean).join(' · ');
        }

        function pickOrder(list) {
            return new Promise(function (resolve) {
                items.textContent = '';
                meta.textContent = list.length + ' Bestellungen zu dieser Kommission. Welche soll geladen werden?';
                var settled = false;
                function finish(order) {
                    if (settled) return;
                    settled = true;
                    state.pickFinish = null;
                    resolve(order || null);
                }
                state.pickFinish = finish;
                list.forEach(function (order) {
                    var btn = el('button', 'pick', orderChoiceLabel(order));
                    btn.type = 'button';
                    btn.addEventListener('click', function () { finish(order); });
                    items.appendChild(btn);
                });
                var cancel = el('button', 'ghost', 'Abbrechen');
                cancel.type = 'button';
                cancel.addEventListener('click', function () { finish(null); });
                items.appendChild(cancel);
            });
        }

        function chosenAddress() {
            return state.addresses[Number(addressSelect.value)] || pickAddress(state.addresses);
        }

        function purposeName() {
            var hit = state.purposes.find(function (row) { return row.Id === 'SHIPMENT_DOC'; });
            return (hit && hit.Name) || 'Lieferschein';
        }

        function settingsFor(supplierId) {
            return state.settings.find(function (row) {
                return String(row.SupplierId || '') === String(supplierId);
            }) || null;
        }

        async function suppliers() {
            if (state.supplierList) return state.supplierList;
            var raw = await api('GET', 'api/catalog/suppliers?classification-id=SUPPLIER_WR');
            var list = asList(raw);
            if (!list.length && raw && typeof raw === 'object') list = raw.Suppliers || raw.suppliers || raw.Results || raw.results || [];
            if (!Array.isArray(list)) list = [];
            var seen = {};
            var out = [];
            function add(id, name) {
                id = String(id || '');
                name = String(name || '').trim();
                if (!id || !name || seen[id]) return;
                seen[id] = true;
                out.push({ Id: id, Name: name });
            }
            list.forEach(function (row) { add(supplierRowId(row), supplierLabel(row)); });
            (state.settings || []).forEach(function (row) {
                add(row.SupplierId || supplierRowId(row), row.SupplierName || supplierLabel(row));
            });
            state.supplierList = out;
            return state.supplierList;
        }

        async function makersFor(supplierId) {
            var key = String(supplierId || '');
            if (!key) return [];
            if (state.makerCache[key]) return state.makerCache[key];
            var list = asList(await api('GET', 'api/catalog/suppliers/' + encodeURIComponent(key) + '/manufacturers'));
            state.makerCache[key] = list;
            return list;
        }

        function makerName(row) {
            return row.Name || row.ManufacturerName || row.CompanyName || '';
        }

        function makerId(row) {
            return row.Id || row.ManufacturerId || row.id || '';
        }

        async function pdfBlob(orderId) {
            if (file.files && file.files[0]) return file.files[0];
            var res = await request('GET', BASE + 'api/ordering/orders/' + encodeURIComponent(orderId) + '/pdf?Auth=' + encodeURIComponent(token()), { responseType: 'arraybuffer' });
            var bytes = new Uint8Array(res.body);
            var head = '';
            var i;
            for (i = 0; i < Math.min(5, bytes.length); i++) head += String.fromCharCode(bytes[i]);
            if (head.indexOf('%PDF') !== 0) throw new Error('Auftrags-PDF ist kein PDF. Bitte die Datei aus der Mail auswählen.');
            return new Blob([res.body], { type: 'application/pdf' });
        }

        async function uploadPdf(blob) {
            var data = new FormData();
            var name = blob.name || 'Details.pdf';
            data.append('file', blob, name);
            var res = await request('POST', BASE + 'api/wr/storage/files', { data: data });
            var body = res.body || {};
            return {
                Id: body.Id || body.id,
                Name: body.Name || body.name || name,
                Size: body.Size || body.size || blob.size || 0,
                ContentType: body.ContentType || body.contentType || 'application/pdf'
            };
        }

        async function refreshCases() {
            state.cases = asList(await api('GET', 'api/wr/cases?include-items=true'));
            return state.cases;
        }

        async function deliverSheet(entry) {
            if (entry && entry.r) rememberReturn(entry);
            if (!sheetUrl()) return { ok: false, text: 'Sheet-Link fehlt. Admin öffnen und den Web-App-Link eintragen.' };
            var stamp = berlinNow();
            var row = {
                date: stamp.date,
                time: stamp.time,
                who: workerName() || entry.who || 'Berger',
                orderId: entry.order || '',
                stockId: strictStockId(entry.stock || entry.kommission || ''),
                caseId: entry.r,
                article: entry.article || '',
                name: entry.name || '',
                supplier: entry.supplier || '',
                brand: entry.brand || ''
            };
            try {
                var saved = await postSheet([row]);
                if (saved && (saved.skipped || 0) > 0 && !(saved.added || 0)) return { ok: true, text: 'Steht schon im Sheet' };
                return { ok: true, text: 'Im Sheet' };
            } catch (error) {
                var message = error && error.message ? error.message : String(error);
                if (/login|anmeld|zugriff|auth/i.test(message)) {
                    if (!state.sheetAuthOpened) {
                        state.sheetAuthOpened = true;
                        window.open(sheetUrl(), '_blank');
                    }
                    return { ok: false, text: 'Sheet-Zugriff fehlt. Der Tab ist offen — dort erlauben, danach nochmal senden.' };
                }
                return { ok: false, text: 'Nicht im Sheet: ' + message };
            }
        }

        function entryFromCaseFast(row) {
            if (row && row.r && !row.CaseId) return {
                r: row.r,
                order: row.order || '',
                stock: row.stock || '',
                kommission: row.kommission || '',
                article: row.article || '',
                name: row.name || '',
                qty: row.qty || 1,
                supplier: row.supplier || '',
                brand: row.brand || '',
                vin: row.vin || '',
                reason: row.reason || PROBLEM_DEFAULT
            };
            var history = readHistory().filter(function (item) { return item.r === String(row.CaseId || ''); })[0] || {};
            var caseItems = row.Items || row.items || [];
            var first = (Array.isArray(caseItems) && caseItems[0]) || {};
            var kommission = commissionFromCase(row) || history.kommission || '';
            return {
                r: String(row.CaseId || ''),
                order: String(row.OrderProcessId || history.order || ''),
                stock: stockIdFromCommission(kommission) || history.stock || '',
                kommission: kommission,
                article: String(first.CatalogNumber || first.ArticleNumber || history.article || ''),
                name: String(first.Name || first.Description || history.name || ''),
                qty: first.Quantity || history.qty || 1,
                supplier: String(row.SupplierName || history.supplier || ''),
                brand: String(first.ManufacturerName || history.brand || ''),
                vin: String(row.Vin || history.vin || ''),
                reason: history.reason || PROBLEM_DEFAULT
            };
        }

        function sheetRowFromCard(card, caseId, orderId, who, stamp) {
            var item = card._item || {};
            var article = card.dataset.article || articleNo(item);
            var fromMail = mailItemFor(state.queue, article);
            var supplierPick = card._supplier && card._supplier.selectedOptions && card._supplier.selectedOptions[0];
            var supplier = (supplierPick && supplierPick.value && supplierPick.textContent) || (fromMail && fromMail.supplier) || supplierNameOf(item);
            var brand = (card._brand && card._brand.value.trim()) || (fromMail && fromMail.brand) || brandOf(item);
            var stock = strictStockId([
                commissionOf(state.order),
                state.queue && state.queue.stockId,
                state.queue && state.queue.kommission
            ].filter(Boolean).join(' '));
            return {
                date: (stamp && stamp.date) || '',
                time: (stamp && stamp.time) || '',
                who: who || '',
                orderId: orderId || '',
                stockId: stock,
                caseId: caseId,
                article: article,
                name: partName(item) || (fromMail && fromMail.name) || '',
                brand: brand || '',
                supplier: supplier || ''
            };
        }

        async function sendAll() {
            if (state.busy) return;
            var orderId = orderInput.value.trim();
            var selected = selectedRows();
            var address = chosenAddress();
            if (!token()) return setLog('Nicht eingeloggt.');
            if (!workerName()) return setLog('Erst deinen Namen eintragen, sonst weiß später keiner wer die Retoure angemeldet hat.');
            if (!orderId) return setLog('Bestellung Nr fehlt.');
            if (!selected.length) return setLog('Kein Teil angehakt.');
            if (!address || !addressId(address)) return setLog('Keine Lieferadresse gefunden.');
            if (!reasonSelect.value) return setLog('Kein Retourengrund geladen.');
            state.busy = true;
            state.sheetAuthOpened = false;
            sendBtn.disabled = true;
            var lines = [];
            var done = [];
            var sheetRows = [];
            try {
                var blob = await pdfBlob(orderId);
                var known = new Set(state.cases.map(function (row) { return row.CaseId; }));
                var index;
                for (index = 0; index < selected.length; index++) {
                    var card = selected[index];
                    var item = card._item;
                    var article = card.dataset.article;
                    var fromMail = mailItemFor(state.queue, article);
                    var supplierPick = card._supplier && card._supplier.selectedOptions && card._supplier.selectedOptions[0];
                    var supplierId = (card._supplier && card._supplier.value) || supplierIdOf(item);
                    var supplierName = (supplierPick && supplierPick.value && supplierPick.textContent) || (fromMail && fromMail.supplier) || supplierNameOf(item);
                    if (!supplierId) {
                        var supplierHit = matchSupplier(await suppliers(), supplierName);
                        if (supplierHit) {
                            supplierId = supplierRowId(supplierHit);
                            supplierName = supplierLabel(supplierHit);
                            if (card._supplier) {
                                card._supplier.value = String(supplierId);
                                card._supplier.classList.remove('missing');
                            }
                        }
                    }
                    if (!supplierId) {
                        if (card._supplier) card._supplier.classList.add('missing');
                        lines.push(article + '  Lieferant wählen: ' + supplierName);
                        setLog(lines.join('\n'));
                        continue;
                    }
                    var brandName = (card._brand && card._brand.value.trim()) || (fromMail && fromMail.brand) || item.Manufacturer || '';
                    var maker = matchByName(await makersFor(supplierId), brandName, makerName);
                    if (!maker && item.ManufacturerId && sameName(item.Manufacturer || item.ManufacturerName, brandName)) {
                        maker = { Id: item.ManufacturerId, Name: item.Manufacturer || brandName };
                    }
                    if (!maker) {
                        lines.push(article + '  Hersteller nicht in der Liste: ' + (brandName || 'leer') + ' · ' + supplierName);
                        setLog(lines.join('\n'));
                        continue;
                    }
                    if (!again.checked) {
                        var existing = openCase(state.cases, orderId, article);
                        if (existing && existing.CaseId) {
                            done.push({ caseId: existing.CaseId, article: article, supplier: supplierName });
                            var oldStamp = stampFromCase(existing);
                            if (!oldStamp.day) oldStamp = berlinNow();
                            sheetRows.push(sheetRowFromCard(card, existing.CaseId, orderId, workerName() || casePerson(existing) || 'nachgetragen', oldStamp));
                            lines.push(existing.CaseId + '  ' + article + '  ' + supplierName + '  schon vorhanden');
                            setLog(lines.join('\n'));
                            continue;
                        }
                    }
                    setLog(lines.concat(['Lege an ' + article + ' …']).join('\n'));
                    var uploaded = await uploadPdf(blob);
                    var method = returnMethod(settingsFor(supplierId));
                    var notes = [{ NoteTypeId: 'PROBLEM_DESCRIPTION', Content: problem.value.trim() || PROBLEM_DEFAULT }];
                    if (state.queue && state.queue.kommission && state.queue.orderId === orderId) {
                        notes.push({ NoteTypeId: 'COMMENT', Content: 'Kommission: ' + state.queue.kommission });
                    }
                    var payload = {
                        CustomerName: customerName(),
                        CustomerPhone: phoneOf(address),
                        CustomerEmail: emailOf(address),
                        CustomerContact: addressRow(address).ToName || customerName(),
                        CustomerAddressId: addressId(address),
                        CaseTypeId: 'RETURN',
                        CaseReasonId: reasonSelect.value,
                        Kba: null,
                        Vin: card._vin.value.trim().toUpperCase() || null,
                        OrderProcessId: orderId,
                        SupplierId: supplierId,
                        CustomerId: customerId(),
                        SupplierName: supplierName,
                        Items: [{
                            ArticleId: item.ArticleId || null,
                            ManufacturerId: makerId(maker),
                            ManufacturerName: makerName(maker) || brandName,
                            CatalogNumber: article,
                            Name: item.Description || item.Name || (fromMail && fromMail.name) || '',
                            Quantity: Number(card._qty.value) || 1,
                            ReplacementDeliveryDesired: card._replace.checked,
                            OrderItemId: item.ItemId || item.OrderItemId || null,
                            OrderId: orderId
                        }],
                        Notes: notes,
                        Attachments: [{
                            PurposeId: 'SHIPMENT_DOC',
                            PurposeDescription: purposeName(),
                            ContentType: uploaded.ContentType,
                            Size: uploaded.Size,
                            FileName: uploaded.Name,
                            Id: uploaded.Id
                        }],
                        ItemsReturnMethodId: method.id,
                        PickUpReturnDate: null,
                        BulkyGoods: method.bulky
                    };
                    console.log('[ALFAH Retoure]', payload);
                    var created = await request('POST', BASE + 'api/wr/cases', { json: payload });
                    var caseId = caseIdOf(created.body, created.headers);
                    if (!caseId) {
                        await new Promise(function (resolve) { setTimeout(resolve, 800); });
                        var fresh = await refreshCases();
                        var match = fresh.find(function (row) {
                            return row.CaseId && !known.has(row.CaseId) && String(row.OrderProcessId || '') === orderId && caseHasArticle(row, article);
                        });
                        caseId = match ? match.CaseId : '';
                        fresh.forEach(function (row) { if (row.CaseId) known.add(row.CaseId); });
                    } else {
                        known.add(caseId);
                    }
                    var entry = {
                        caseId: caseId || '',
                        article: article,
                        supplier: supplierName
                    };
                    done.push(entry);
                    if (/^R-\d+$/.test(caseId || '')) sheetRows.push(sheetRowFromCard(card, caseId, orderId, workerName(), berlinNow()));
                    lines.push((caseId || 'angelegt, R-Nummer nicht im Antworttext') + '  ' + article + '  ' + supplierName);
                    setLog(lines.join('\n'));
                }
                var sheetIds = {};
                if (sheetRows.length) {
                    var went = 'bad';
                    try {
                        var saved = await postSheet(sheetRows);
                        if (saved && saved.ok !== false && ((saved.added || 0) + (saved.skipped || 0) > 0)) went = 'ok';
                    } catch (sheetError) {
                        went = 'bad';
                    }
                    sheetRows.forEach(function (row) { sheetIds[row.caseId] = went; });
                }
                await refreshCases();
                var notes = lines.filter(function (line) { return !/^R-\d+/.test(line); });
                if (done.length) showResults(done, notes, sheetIds);
                else if (notes.length) setLog(notes.join('\n'));
                else setLog('Nichts angelegt.');
                var complete = selected.length && done.length === selected.length && done.every(function (entry) { return /^R-\d+/.test(entry.caseId); });
                if (complete) clearTop();
            } catch (error) {
                console.log('[ALFAH Retoure]', error);
                setLog(lines.concat([error.message || String(error)]).join('\n'));
            } finally {
                state.busy = false;
                sendBtn.disabled = false;
            }
        }

        var adminPass = document.createElement('input');
        adminPass.type = 'password';
        adminPass.placeholder = 'Passwort';
        var adminUnlock = el('button', 'act', 'Öffnen');
        adminUnlock.type = 'button';
        var adminUrl = document.createElement('input');
        adminUrl.type = 'text';
        adminUrl.placeholder = 'https://script.google.com/.../exec';
        var adminSaveUrl = el('button', 'ghost', 'Link speichern');
        adminSaveUrl.type = 'button';
        var adminAllow = el('button', 'act', 'Sheet-Zugriff erlauben');
        adminAllow.type = 'button';
        var adminFilter = document.createElement('input');
        adminFilter.type = 'text';
        adminFilter.placeholder = 'R-Nummer filtern';
        var adminList = el('div', 'case-list');
        var adminSend = el('button', 'act', 'Ausgewählte an Sheet senden');
        adminSend.type = 'button';
        var adminClose = el('button', 'ghost', 'Schließen');
        adminClose.type = 'button';
        var adminStatus = el('div', 'log');

        function paintAdmin(unlocked) {
            adminPop.textContent = '';
            adminPop.appendChild(el('h3', '', 'Admin'));
            if (!unlocked) {
                adminPop.appendChild(adminPass);
                adminPop.appendChild(adminUnlock);
                adminPop.appendChild(adminClose);
                return;
            }
            adminPop.appendChild(el('label', '', 'Sheet Web-App'));
            adminPop.appendChild(adminUrl);
            var urlRow = el('div', 'foot');
            urlRow.appendChild(adminSaveUrl);
            urlRow.appendChild(adminAllow);
            adminPop.appendChild(urlRow);
            adminPop.appendChild(adminFilter);
            adminPop.appendChild(adminList);
            var foot = el('div', 'foot');
            foot.appendChild(adminSend);
            foot.appendChild(adminClose);
            adminPop.appendChild(foot);
            adminPop.appendChild(adminStatus);
        }

        function renderAdminRows() {
            var needle = adminFilter.value.trim().toUpperCase();
            adminList.textContent = '';
            (state.adminRows || []).forEach(function (row) {
                var label = [row.r, row.order, row.article, row.name, row.supplier, row.stock].filter(Boolean).join(' · ');
                if (needle && label.toUpperCase().indexOf(needle) === -1) return;
                var line = el('label', 'case');
                var box = document.createElement('input');
                box.type = 'checkbox';
                line.appendChild(box);
                line.appendChild(document.createTextNode(label));
                line._entry = row;
                adminList.appendChild(line);
            });
        }

        async function fillAdminList() {
            adminStatus.textContent = 'Lade Retouren …';
            adminList.textContent = '';
            var rows = [];
            try {
                var fresh = await refreshCases();
                var i;
                for (i = 0; i < fresh.length; i++) {
                    if (!/^R-\d+$/.test(String(fresh[i].CaseId || ''))) continue;
                    rows.push(entryFromCaseFast(fresh[i]));
                }
            } catch (error) {
                adminStatus.textContent = error.message || String(error);
            }
            readHistory().forEach(function (item) {
                if (!item || !/^R-\d+$/.test(item.r || '')) return;
                if (!rows.some(function (row) { return row.r === item.r; })) rows.push(entryFromCaseFast(item));
            });
            rows.sort(function (a, b) {
                var left = Number((String(b.r).match(/(\d+)/) || [])[1] || 0);
                var right = Number((String(a.r).match(/(\d+)/) || [])[1] || 0);
                return left - right;
            });
            state.adminRows = rows;
            renderAdminRows();
            adminStatus.textContent = rows.length ? rows.length + ' Retouren. Haken setzen und an das Sheet senden.' : 'Keine R-Nummer gefunden.';
        }

        async function sendAdminSelection() {
            var picked = [];
            adminList.querySelectorAll('.case').forEach(function (line) {
                var box = line.querySelector('input[type=checkbox]');
                if (box && box.checked && line._entry) picked.push(line._entry);
            });
            if (!picked.length) {
                adminStatus.textContent = 'Keine Retoure ausgewählt.';
                return;
            }
            if (!sheetBase()) {
                adminStatus.textContent = 'Zuerst den Web-App-Link speichern.';
                return;
            }
            adminSend.disabled = true;
            state.sheetAuthOpened = false;
            var notes = [];
            var i;
            for (i = 0; i < picked.length; i++) {
                var entry = picked[i];
                if (!entry.stock && entry.order) {
                    try {
                        var order = await orderById(entry.order);
                        entry.kommission = entry.kommission || commissionOf(order);
                        entry.stock = stockIdFromCommission(entry.kommission) || '';
                    } catch (e) {}
                }
                adminStatus.textContent = 'Sende ' + entry.r + ' …';
                var result = await deliverSheet(entry);
                notes.push(entry.r + '  ' + result.text);
                adminStatus.textContent = notes.join('\n');
            }
            adminSend.disabled = false;
        }

        adminBtn.addEventListener('click', function () {
            state.adminOn = false;
            adminPass.value = '';
            adminPass.placeholder = 'Passwort';
            adminPop.hidden = false;
            paintAdmin(false);
            adminPass.focus();
        });
        adminUnlock.addEventListener('click', function () {
            if (adminPass.value.trim() !== ADMIN_PASS) {
                adminPass.value = '';
                adminPass.placeholder = 'Falsch';
                return;
            }
            state.adminOn = true;
            adminUrl.value = sheetBase();
            paintAdmin(true);
            fillAdminList();
        });
        adminPass.addEventListener('keydown', function (event) {
            if (event.key === 'Enter') adminUnlock.click();
        });
        adminClose.addEventListener('click', function () {
            adminPop.hidden = true;
        });
        adminSaveUrl.addEventListener('click', function () {
            GM_setValue(SHEET_URL_KEY, adminUrl.value.trim());
            adminStatus.textContent = adminUrl.value.trim() ? 'Link gespeichert.' : 'Link gelöscht.';
        });
        adminAllow.addEventListener('click', function () {
            GM_setValue(SHEET_URL_KEY, adminUrl.value.trim());
            var url = sheetBase();
            if (!url) {
                adminStatus.textContent = 'Zuerst den Web-App-Link eintragen.';
                return;
            }
            window.open(url, '_blank');
            adminStatus.textContent = 'Tab ist offen. Zugriff erlauben, danach die Retoure nochmal senden.';
        });
        adminFilter.addEventListener('input', renderAdminRows);
        adminSend.addEventListener('click', sendAdminSelection);
        stopBtn.addEventListener('click', function () {
            state.stopSearch = true;
            showProgress('Suche wird gestoppt …');
            if (state.pickFinish) state.pickFinish(null);
        });
        workerInput.addEventListener('input', function () {
            GM_setValue(WORKER_KEY, workerInput.value.trim());
        });
        workerInput.addEventListener('change', function () {
            GM_setValue(WORKER_KEY, workerInput.value.trim());
        });
        sheetUrlInput.addEventListener('change', function () {
            var value = sheetUrlInput.value.trim();
            if (value && !/^https:\/\/script\.google\.com\/.+\/exec\/?$/i.test(value)) {
                sheetHint.textContent = 'Das ist keine exec-URL. Die endet auf /exec.';
                return;
            }
            GM_setValue(SHEET_URL_KEY, value);
            sheetUrlLabel.hidden = !!sheetUrl();
            sheetUrlInput.hidden = !!sheetUrl();
        });
        resetBtn.addEventListener('click', resetHud);
        loadBtn.addEventListener('click', loadOrder);
        file.addEventListener('change', refreshSlipBadge);
        sendBtn.addEventListener('click', sendAll);
        readBtn.addEventListener('click', function () {
            var parsed = parseMail(paste.value);
            if (!parsed.orderId) {
                setLog('In dem Text steht keine Bestellung Nr.');
                return;
            }
            state.queue = {
                orderId: parsed.orderId,
                kommission: parsed.kommission,
                stockId: parsed.stockId,
                items: parsed.items,
                savedAt: Date.now()
            };
            saveQueue(state.queue);
            orderInput.value = parsed.orderId;
            loadOrder();
        });
        orderInput.addEventListener('keydown', function (event) {
            if (event.key === 'Enter') loadOrder();
        });

        if (state.queue && state.queue.orderId) orderInput.value = state.queue.orderId;
        prepare().then(function () {
            if (orderInput.value.trim()) return loadOrder();
        });
        if (typeof GM_addValueChangeListener === 'function') {
            GM_addValueChangeListener(QUEUE_KEY, function () {
                state.queue = loadQueue();
                if (state.queue && state.queue.orderId) {
                    orderInput.value = state.queue.orderId;
                    loadOrder();
                }
            });
        }
    }

    if (location.hostname.indexOf('mail.google.com') !== -1) startGmail();
    else if (location.hostname === 'alfah.de') startAlfah();
})();
