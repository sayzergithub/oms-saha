/*
 * OMS field mode - stock catalog from the PC (spec V6 §12, P1).
 *   request : GET {lan}/saha/api/catalog?dev=<dev>&ts=<unix ms>&sig=<b64url HMAC-SHA256(macKey, "catalog|dev|ts")>
 *   response: 200 text/plain = armor (spec §2 envelope, same device secret, deflate) of
 *             {"v":1,"type":"catalog","pc","at","count","parts":[{id,m,f,s,q,l,c,val,pkg,r:[rawHash…]}]}
 *             else a short code: 404 DISABLED, 400 BAD_REQUEST, 403 UNKNOWN_DEVICE | REVOKED | BAD_SIG | "CLOCK <server ms>"
 *   rawHash : drop every UTF-16 unit <= 0x20, UTF-8, SHA-256, base64url, first 16 chars (twin of C# RemoteConst.RawHash).
 * The decrypted catalog is kept in IndexedDB 'oms-catalog' (localStorage fallback, memory only when neither works).
 * UMD, depends on envelope.js (global OmsEnvelope; Node: require('./envelope.js')). Browser: global OmsCatalog.
 */
(function (root, factory) {
  'use strict';
  if (typeof module === 'object' && module.exports) module.exports = factory(root, require('./envelope.js'));
  else root.OmsCatalog = factory(root, root.OmsEnvelope);
})(typeof globalThis !== 'undefined' ? globalThis : (typeof self !== 'undefined' ? self : this), function (root, Env) {
  'use strict';

  var DB_NAME = 'oms-catalog', DB_STORE = 'kv', DB_KEY = 'current';
  var LS_KEY = 'oms-catalog.v1', LS_SKEW = 'oms-catalog.skew';
  var API_PATH = '/saha/api/catalog';
  var DEFAULT_TIMEOUT = 8000;
  var MAX_PARTS = 200000;

  // ------------------------------------------------------------------ helpers
  function getCrypto() {
    var c = root && root.crypto;
    if (c && c.subtle) return c;
    if (typeof module === 'object' && module.exports && typeof require === 'function') {
      try { var w = require('crypto').webcrypto; if (w && w.subtle) return w; } catch (e) { /* not node */ }
    }
    return null;
  }

  function catError(code, message) {
    var e = new Error(message || code);
    e.code = code;
    return e;
  }

  function ascii(s) {
    var out = new Uint8Array(s.length);
    for (var i = 0; i < s.length; i++) out[i] = s.charCodeAt(i) & 0x7f;
    return out;
  }

  function str(x) { return typeof x === 'string' ? x : (x == null ? '' : String(x)); }

  function secBytes(sec) {
    if (typeof sec === 'string') return Env.b64urlDecode(sec);
    if (sec instanceof Uint8Array) return sec;
    if (sec instanceof ArrayBuffer) return new Uint8Array(sec);
    if (sec && ArrayBuffer.isView(sec)) return new Uint8Array(sec.buffer, sec.byteOffset, sec.byteLength);
    if (Array.isArray(sec)) return new Uint8Array(sec);
    return null;
  }

  /** Thousands with '.' (Turkish), e.g. 12000 -> "12.000". */
  function fmtQty(q) {
    if (typeof q !== 'number' || !isFinite(q)) return '';
    var neg = q < 0, s = String(Math.round(Math.abs(q)));
    for (var i = s.length - 3; i > 0; i -= 3) s = s.substring(0, i) + '.' + s.substring(i);
    return (neg ? '-' : '') + s;
  }

  function fmtMinutes(ms) {
    var m = Math.round(Math.abs(ms) / 60000);
    if (m < 60) return (m < 1 ? 1 : m) + ' dk';
    var h = Math.floor(m / 60), r = m % 60;
    if (h < 48) return h + ' sa' + (r ? ' ' + r + ' dk' : '');
    return Math.round(h / 24) + ' gün';
  }

  // ------------------------------------------------------------------ rawHash + request signature
  /** RemoteConst.RawHash twin. Resolves with the 16-char hash, or null for an empty / whitespace-only raw. */
  function rawHash(raw) {
    return Promise.resolve().then(function () {
      if (typeof raw !== 'string' || raw.length === 0) return null;
      var kept = [];
      for (var i = 0; i < raw.length; i++) if (raw.charCodeAt(i) > 0x20) kept.push(raw[i]);
      if (kept.length === 0) return null;
      var c = getCrypto();
      if (!c) throw catError('NO_CRYPTO', 'Bu tarayıcıda şifreleme (WebCrypto) yok. Sayfayı HTTPS adresinden açın.');
      // Lone surrogates become U+FFFD in both TextEncoder and C# Encoding.UTF8.
      return c.subtle.digest('SHA-256', Env.utf8(kept.join(''))).then(function (d) {
        return Env.b64urlEncode(new Uint8Array(d)).substring(0, 16);
      });
    });
  }

  /** b64url(HMAC-SHA256(macKey(sec), ASCII("catalog|" + dev + "|" + tsMs))). */
  function signRequest(dev, sec, tsMs) {
    return Promise.resolve().then(function () {
      if (!Env.isDeviceId(dev)) throw catError('BAD_ARGUMENT', 'Geçersiz cihaz kimliği');
      var s = secBytes(sec);
      if (!s || s.length === 0) throw catError('BAD_ARGUMENT', 'Cihaz sırrı geçersiz');
      var ts = typeof tsMs === 'number' ? String(Math.floor(tsMs)) : str(tsMs).trim();
      if (!/^\d{1,16}$/.test(ts)) throw catError('BAD_ARGUMENT', 'Geçersiz zaman damgası');
      var c = getCrypto();
      if (!c) throw catError('NO_CRYPTO', 'Bu tarayıcıda şifreleme (WebCrypto) yok. Sayfayı HTTPS adresinden açın.');
      return Env.deriveKeys(s).then(function (k) {
        return c.subtle.importKey('raw', k.mac, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
      }).then(function (key) {
        return c.subtle.sign('HMAC', key, ascii('catalog|' + dev + '|' + ts));
      }).then(function (sig) { return Env.b64urlEncode(new Uint8Array(sig)); });
    });
  }

  // ------------------------------------------------------------------ in-memory catalog + indexes
  var cur = null;          // {parts, pc, builtAt, syncedAt, url, count}
  var byHash = null;       // Map rawHash -> part
  var byMpn = null;        // Map normKey -> [parts]
  var bySpn = null;
  var listeners = [];

  /** Case / whitespace insensitive key: "  sn74hc00 dr" -> "SN74HC00DR". toUpperCase is locale independent ('i' -> 'I'). */
  function normKey(s) {
    if (typeof s !== 'string') return '';
    return s.replace(/[\s ​﻿]+/g, '').toUpperCase();
  }

  function cleanPart(p) {
    if (!p || typeof p !== 'object') return null;
    var r = [];
    if (Array.isArray(p.r)) for (var i = 0; i < p.r.length; i++) if (typeof p.r[i] === 'string' && p.r[i]) r.push(p.r[i]);
    var q = typeof p.q === 'number' && isFinite(p.q) ? p.q : null;
    return {
      id: str(p.id), m: str(p.m), f: str(p.f), s: str(p.s), q: q, l: str(p.l),
      c: str(p.c), val: str(p.val), pkg: str(p.pkg), r: r, d: str(p.d)
    };
  }

  function buildIndexes(parts) {
    var h = new Map(), m = new Map(), s = new Map(), i, j, p, k;
    for (i = 0; i < parts.length; i++) {
      p = parts[i];
      for (j = 0; j < p.r.length; j++) if (!h.has(p.r[j])) h.set(p.r[j], p);
      k = normKey(p.m);
      if (k) { if (!m.has(k)) m.set(k, []); m.get(k).push(p); }
      k = normKey(p.s);
      if (k) { if (!s.has(k)) s.set(k, []); s.get(k).push(p); }
    }
    byHash = h; byMpn = m; bySpn = s;
  }

  function setCurrent(rec) {
    cur = rec;
    if (rec) buildIndexes(rec.parts); else { byHash = byMpn = bySpn = null; }
    var inf = info();
    for (var i = 0; i < listeners.length; i++) { try { listeners[i](inf); } catch (e) { /* listener bug must not break sync */ } }
  }

  /** Validates the decrypted catalog JSON and turns it into a stored record. Throws BAD_CATALOG. */
  function recordFromJson(json, url, nowMs) {
    var o;
    try { o = JSON.parse(json); } catch (e) { throw catError('BAD_CATALOG', 'Stok listesi okunamadı (bozuk veri)'); }
    if (!o || typeof o !== 'object' || o.type !== 'catalog' || !Array.isArray(o.parts))
      throw catError('BAD_CATALOG', 'PC\'den gelen yanıt bir stok listesi değil');
    if (typeof o.v === 'number' && o.v > 1) throw catError('BAD_CATALOG', 'Stok listesinin sürümü bu sayfadan yeni. Sayfayı yenileyin.');
    if (o.parts.length > MAX_PARTS) throw catError('BAD_CATALOG', 'Stok listesi çok büyük');
    var parts = [];
    for (var i = 0; i < o.parts.length; i++) { var p = cleanPart(o.parts[i]); if (p) parts.push(p); }
    return {
      v: 1, parts: parts, count: parts.length, pc: str(o.pc), builtAt: str(o.at),
      syncedAt: new Date(nowMs).toISOString(), url: str(url)
    };
  }

  function isRecord(r) {
    return !!(r && typeof r === 'object' && r.v === 1 && Array.isArray(r.parts) && typeof r.syncedAt === 'string');
  }

  // ------------------------------------------------------------------ persistence (IndexedDB -> localStorage -> memory)
  var dbPromise = null;

  function idb() {
    try { return root.indexedDB || null; } catch (e) { return null; }
  }

  function openDb() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise(function (resolve, reject) {
      var f = idb();
      if (!f) { reject(catError('NO_IDB', 'IndexedDB yok')); return; }
      var rq;
      try { rq = f.open(DB_NAME, 1); } catch (e) { reject(e); return; }
      rq.onupgradeneeded = function () {
        var db = rq.result;
        if (!db.objectStoreNames.contains(DB_STORE)) db.createObjectStore(DB_STORE);
      };
      rq.onsuccess = function () {
        var db = rq.result;
        // Another tab upgrading the schema: let it, reopen lazily next time.
        db.onversionchange = function () { try { db.close(); } catch (e) { /* ignore */ } dbPromise = null; };
        resolve(db);
      };
      rq.onerror = function () { reject(rq.error || catError('NO_IDB', 'IndexedDB açılamadı')); };
      rq.onblocked = function () { reject(catError('NO_IDB', 'IndexedDB meşgul')); };
    });
    dbPromise.catch(function () { dbPromise = null; });
    return dbPromise;
  }

  function idbDo(mode, fn) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(DB_STORE, mode), st = tx.objectStore(DB_STORE), out;
        var rq = fn(st);
        if (rq) rq.onsuccess = function () { out = rq.result; };
        tx.oncomplete = function () { resolve(out); };
        tx.onerror = function () { reject(tx.error || catError('IDB', 'IndexedDB hatası')); };
        tx.onabort = function () { reject(tx.error || catError('IDB', 'IndexedDB işlemi iptal edildi')); };
      });
    });
  }

  function lsGet(key) { try { return root.localStorage ? root.localStorage.getItem(key) : null; } catch (e) { return null; } }
  function lsSet(key, v) { try { if (!root.localStorage) return false; root.localStorage.setItem(key, v); return true; } catch (e) { return false; } }
  function lsDel(key) { try { if (root.localStorage) root.localStorage.removeItem(key); } catch (e) { /* ignore */ } }

  function persist(rec) {
    return idbDo('readwrite', function (st) { return st.put(rec, DB_KEY); })
      .then(function () { lsDel(LS_KEY); return 'idb'; })
      .catch(function () { return lsSet(LS_KEY, JSON.stringify(rec)) ? 'ls' : 'memory'; });
  }

  function readStored() {
    return idbDo('readonly', function (st) { return st.get(DB_KEY); })
      .catch(function () { return null; })
      .then(function (r) {
        if (isRecord(r)) return r;
        var t = lsGet(LS_KEY);
        if (!t) return null;
        try { r = JSON.parse(t); } catch (e) { return null; }
        return isRecord(r) ? r : null;
      });
  }

  var loadPromise = null;

  /** Loads the stored catalog (call at page start). Resolves with info() (null when none). Never rejects. */
  function load() {
    if (loadPromise) return loadPromise;
    loadPromise = readStored().then(function (r) {
      // A sync that finished while we were reading wins over the stored copy.
      if (r && !cur) {
        var parts = [];
        for (var i = 0; i < r.parts.length; i++) { var p = cleanPart(r.parts[i]); if (p) parts.push(p); }
        r.parts = parts; r.count = parts.length;
        setCurrent(r);
      }
      return info();
    }).catch(function () { return info(); });
    return loadPromise;
  }

  /** Removes the stored catalog (e.g. after unpairing). */
  function clear() {
    setCurrent(null);
    loadPromise = null;
    lsDel(LS_KEY);
    return idbDo('readwrite', function (st) { return st.delete(DB_KEY); }).then(function () { return true; }, function () { return true; });
  }

  // ------------------------------------------------------------------ info + lookups
  /** {count, at (local sync time, ISO), pc, url, builtAt (PC time of the list)} or null when no catalog is loaded. */
  function info() {
    if (!cur) return null;
    return { count: cur.count, at: cur.syncedAt, pc: cur.pc, url: cur.url, builtAt: cur.builtAt };
  }

  /** Milliseconds since the last successful sync (Infinity when none). */
  function ageMs(nowMs) {
    if (!cur) return Infinity;
    var t = Date.parse(cur.syncedAt);
    return isFinite(t) ? Math.max(0, (nowMs != null ? nowMs : Date.now()) - t) : Infinity;
  }

  /** True when there is no catalog or the last sync is older than maxAgeMs (default 12 h, spec §12 silent retry). */
  function isStale(maxAgeMs, nowMs) {
    return ageMs(nowMs) > (maxAgeMs != null ? maxAgeMs : 12 * 3600 * 1000);
  }

  function lookupRawHash(hash) {
    if (!byHash || typeof hash !== 'string' || !hash) return null;
    return byHash.get(hash) || null;
  }

  /** Resolves with the part whose reel (or any of its lots) has this raw label code, or null. */
  function lookupRaw(raw) {
    return rawHash(raw).then(function (h) { return h ? lookupRawHash(h) : null; }, function () { return null; });
  }

  function locationsOf(parts) {
    var seen = {}, out = [];
    for (var i = 0; i < parts.length; i++) {
      var l = parts[i].l.trim();
      if (l && !seen[l.toLocaleUpperCase('tr')]) { seen[l.toLocaleUpperCase('tr')] = 1; out.push(l); }
    }
    out.sort(function (a, b) { return a.localeCompare(b, 'tr', { numeric: true, sensitivity: 'base' }); });
    return out.join(', ');
  }

  /** First part's fields + q = summed quantity (null when unknown everywhere), l = all locations, parts = every match. */
  function aggregate(list) {
    if (!list || list.length === 0) return null;
    var total = null;
    for (var i = 0; i < list.length; i++) if (list[i].q != null) total = (total || 0) + list[i].q;
    var a = {};
    for (var k in list[0]) if (Object.prototype.hasOwnProperty.call(list[0], k)) a[k] = list[0][k];
    a.q = total;
    a.l = locationsOf(list);
    a.parts = list.slice();
    return a;
  }

  /** MPN lookup (case / whitespace insensitive). Returns {…first part, q:total, l:"A3, B1", parts:[…]} or null. */
  function lookupMpn(mpn) {
    var k = normKey(mpn);
    return k && byMpn ? aggregate(byMpn.get(k)) : null;
  }

  /** Same as lookupMpn for the supplier part number (e.g. LCSC C916399). */
  function lookupSpn(spn) {
    var k = normKey(spn);
    return k && bySpn ? aggregate(bySpn.get(k)) : null;
  }

  /** Turkish one-liner for a lookupMpn result: "Stokta: 12.000 adet · A3, B1". */
  function describe(hit) {
    if (!hit) return '';
    var t = 'Stokta' + (hit.q != null ? ': ' + fmtQty(hit.q) + ' adet' : ' var');
    if (hit.l) t += ' · ' + hit.l;
    if (hit.parts && hit.parts.length > 1) t += ' (' + hit.parts.length + ' kayıt)';
    return t;
  }

  /** cb(info) after every load / sync / clear. Returns an unsubscribe function. */
  function onChange(cb) {
    if (typeof cb !== 'function') return function () { };
    listeners.push(cb);
    return function () { var i = listeners.indexOf(cb); if (i >= 0) listeners.splice(i, 1); };
  }

  // ------------------------------------------------------------------ sync
  /** "192.168.1.20:47803/" -> "https://192.168.1.20:47803"; strips a trailing /saha or /saha/... Returns '' when unusable. */
  function normRoot(u) {
    var s = str(u).trim();
    if (!s) return '';
    if (!/^[a-z][a-z0-9+.-]*:\/\//i.test(s)) s = 'https://' + s;
    var m = /^(https?:\/\/[^\/?#\s]+)/i.exec(s);
    return m ? m[1].replace(/^HTTPS?:/i, function (x) { return x.toLowerCase(); }) : '';
  }

  function rootList(lan) {
    var list = [], seen = {}, i, r;
    var arr = Array.isArray(lan) ? lan : (typeof lan === 'string' ? lan.split(/[\s,;]+/) : []);
    // The page served by the PC itself (/saha/) can always reach its own origin: try it first.
    try {
      var loc = root.location;
      if (loc && /^https?:$/.test(loc.protocol) && /^\/saha\//.test(loc.pathname)) arr = [loc.origin].concat(arr);
    } catch (e) { /* no location */ }
    for (i = 0; i < arr.length; i++) {
      r = normRoot(arr[i]);
      if (r && !seen[r.toLowerCase()]) { seen[r.toLowerCase()] = 1; list.push(r); }
    }
    return list;
  }

  function pageIsHttps() {
    try { return !!(root.location && root.location.protocol === 'https:'); } catch (e) { return false; }
  }

  function getSkew() { var v = Number(lsGet(LS_SKEW)); return isFinite(v) ? v : 0; }
  var memSkew = null;
  function skewMs() { return memSkew != null ? memSkew : getSkew(); }
  function setSkew(v) { memSkew = v; lsSet(LS_SKEW, String(v)); }

  /** One GET with a timeout. Resolves {status, text} or rejects with code NETWORK | TIMEOUT. */
  function fetchOnce(fetchFn, url, timeoutMs) {
    var ctl = typeof root.AbortController === 'function' ? new root.AbortController() : null;
    var timer = null, timedOut = false;
    var opts = { method: 'GET', mode: 'cors', cache: 'no-store', credentials: 'omit', referrerPolicy: 'no-referrer' };
    if (ctl) opts.signal = ctl.signal;
    var timeout = new Promise(function (_, reject) {
      timer = setTimeout(function () {
        timedOut = true;
        if (ctl) { try { ctl.abort(); } catch (e) { /* ignore */ } }
        reject(catError('TIMEOUT', 'zaman aşımı'));
      }, timeoutMs);
    });
    var req = Promise.resolve().then(function () { return fetchFn(url, opts); }).then(function (res) {
      return res.text().then(function (t) { return { status: res.status, text: t }; });
    }).catch(function (e) {
      if (timedOut || (e && e.name === 'AbortError')) throw catError('TIMEOUT', 'zaman aşımı');
      // fetch rejects with TypeError for DNS / refused / TLS (self-signed certificate not yet accepted) / CORS / PNA failures.
      var err = catError('NETWORK', (e && e.message) || 'ağ hatası');
      err.cause = e;
      throw err;
    });
    return Promise.race([req, timeout]).then(function (r) { clearTimeout(timer); return r; }, function (e) { clearTimeout(timer); throw e; });
  }

  var HTTP_MESSAGES = {
    DISABLED: 'PC\'de stok listesi paylaşımı kapalı (programda Saha ayarlarından açın).',
    UNKNOWN_DEVICE: 'PC bu cihazı tanımıyor. Cihazı PC\'de yeniden eşleştirin.',
    REVOKED: 'Bu cihazın erişimi PC\'de iptal edilmiş. Yeniden eşleştirin.',
    BAD_SIG: 'PC isteği reddetti (imza tutmadı). Eşleştirme kodu eski olabilir; yeniden eşleştirin.',
    CLOCK: 'Telefonun saati PC ile uyuşmuyor. Telefonda otomatik tarih ve saati açın, sonra tekrar deneyin.',
    BAD_REQUEST: 'PC isteği anlamadı. Sayfayı yenileyip tekrar deneyin.',
    NOT_FOUND: 'PC programı stok listesi özelliğini desteklemiyor. Programı güncelleyin.',
    METHOD: 'PC programı stok listesi özelliğini desteklemiyor. Programı güncelleyin.',
    ERROR: 'PC programında bir hata oluştu. Biraz sonra tekrar deneyin.',
    SERVER: 'PC programında bir hata oluştu. Biraz sonra tekrar deneyin.',
    HTTP: 'PC beklenmeyen bir yanıt verdi.'
  };

  /** HTTP error -> {code, serverMs?}. Bodies are "CODE" or "CLOCK <ms>"; HttpServer JSON errors look like {"err":"NOT_FOUND"}. */
  function parseErrorBody(status, text) {
    var t = str(text).trim(), code = '', m;
    if ((m = /^\{\s*"err"\s*:\s*"([A-Z_]+)"/.exec(t))) code = m[1];
    else if ((m = /^([A-Z_]{2,40})(?:\s+(-?\d{1,16}))?\s*$/.exec(t))) {
      code = m[1];
      if (code === 'CLOCK' && m[2]) return { code: 'CLOCK', serverMs: Number(m[2]) };
    }
    if (!code) code = status >= 500 ? 'SERVER' : (status === 404 ? 'NOT_FOUND' : 'HTTP');
    return { code: code };
  }

  function result(fields) {
    var r = { ok: false, count: 0, at: null, pc: '', url: '', error: '', code: '', needCertUrl: '' };
    for (var k in fields) if (Object.prototype.hasOwnProperty.call(fields, k)) r[k] = fields[k];
    return r;
  }

  var syncPromise = null;

  /**
   * Fetches the catalog from the first reachable PC root in opts.lan (tried in order).
   * opts: {dev, sec (Uint8Array | b64url), lan:[roots], timeoutMs=8000}. Test hooks: opts.fetch, opts.now().
   * Resolves (never rejects) with {ok, count, at, pc, url, error, code, needCertUrl}; error is Turkish.
   * A concurrent call while a sync runs gets the same promise.
   */
  function sync(opts) {
    if (syncPromise) return syncPromise;
    syncPromise = doSync(opts || {}).catch(function (e) {
      return result({ code: (e && e.code) || 'ERROR', error: 'Stok listesi alınamadı: ' + ((e && e.message) || e) });
    }).then(function (r) { syncPromise = null; return r; });
    return syncPromise;
  }

  function doSync(o) {
    var now = typeof o.now === 'function' ? o.now : function () { return Date.now(); };
    var fetchFn = typeof o.fetch === 'function' ? o.fetch : (typeof root.fetch === 'function' ? root.fetch.bind(root) : null);
    var timeoutMs = o.timeoutMs > 0 ? o.timeoutMs : DEFAULT_TIMEOUT;
    var dev = str(o.dev).trim(), sec = secBytes(o.sec);
    if (!Env.isDeviceId(dev) || !sec || sec.length !== 32)
      return Promise.resolve(result({ code: 'NOT_PAIRED', error: 'Cihaz eşleştirilmemiş. Önce PC\'deki eşleştirme kodunu okutun.' }));
    if (!getCrypto() || !Env.supported())
      return Promise.resolve(result({ code: 'NO_CRYPTO', error: 'Bu tarayıcıda şifreleme (WebCrypto) yok. Sayfayı HTTPS adresinden açın.' }));
    if (!fetchFn) return Promise.resolve(result({ code: 'NO_FETCH', error: 'Bu tarayıcı ağ isteği (fetch) desteklemiyor.' }));
    var roots = rootList(o.lan);
    if (roots.length === 0)
      return Promise.resolve(result({ code: 'NO_LAN', error: 'PC adresi tanımlı değil. Ayarlar\'dan PC adresini ekleyin (ör. https://192.168.1.20:47803).' }));

    var https = pageIsHttps();
    var firstNet = null, firstTimeout = null, mixed = null, httpFail = null;

    function attempt(rootUrl, retriedClock) {
      var ts = Math.floor(now() + skewMs());
      return signRequest(dev, sec, ts).then(function (sig) {
        var url = rootUrl + API_PATH + '?dev=' + encodeURIComponent(dev) + '&ts=' + ts + '&sig=' + encodeURIComponent(sig);
        return fetchOnce(fetchFn, url, timeoutMs);
      }).then(function (res) {
        if (res.status === 200) return accept(rootUrl, res.text);
        var pe = parseErrorBody(res.status, res.text);
        if (pe.code === 'CLOCK' && isFinite(pe.serverMs)) {
          var skew = pe.serverMs - now();
          // The phone clock is off: sign once more with the PC's time (the HMAC still proves the device).
          if (!retriedClock) { setSkew(Math.round(skew)); return attempt(rootUrl, true); }
          return result({
            code: 'CLOCK', url: rootUrl,
            error: 'Telefonun saati PC\'den ' + fmtMinutes(skew) + ' ' + (skew > 0 ? 'geri' : 'ileri') +
              '. Telefonda otomatik tarih ve saati açın, sonra tekrar deneyin.'
          });
        }
        return result({ code: pe.code, url: rootUrl, error: HTTP_MESSAGES[pe.code] || (HTTP_MESSAGES.HTTP + ' (' + res.status + ' ' + pe.code + ')') });
      });
    }

    function accept(rootUrl, text) {
      var armor = str(text).trim();
      return Env.openFull(armor, sec).catch(function (e) {
        var c = e && e.code;
        throw catError('BAD_ENVELOPE', c === 'BadTag'
          ? 'PC\'den gelen stok listesi doğrulanamadı. Eşleştirme eski olabilir; yeniden eşleştirin.'
          : 'PC\'den gelen stok listesi açılamadı (' + (c || 'hata') + ').');
      }).then(function (r) {
        if (r.dev !== dev) throw catError('BAD_ENVELOPE', 'Stok listesi başka bir cihaz için hazırlanmış.');
        var rec = recordFromJson(r.json, rootUrl, now());
        setCurrent(rec);
        loadPromise = Promise.resolve(info());   // a later load() must not replace the fresh list with the stored one
        return persist(rec).then(function (where) {
          return result({ ok: true, count: rec.count, at: rec.syncedAt, pc: rec.pc, url: rootUrl, storage: where });
        });
      }).catch(function (e) {
        return result({ code: (e && e.code) || 'BAD_ENVELOPE', url: rootUrl, error: (e && e.message) || 'Stok listesi açılamadı' });
      });
    }

    var i = 0;
    function next() {
      if (i >= roots.length) return Promise.resolve(finish());
      var r = roots[i++];
      if (https && /^http:/i.test(r)) { if (!mixed) mixed = r; return next(); }
      return attempt(r, false).then(function (res) {
        if (res.ok) return res;
        // A PC that answered decides (it is the paired PC); only unreachable roots fall through to the next one.
        if (!httpFail) httpFail = res;
        return finish();
      }, function (e) {
        if (e && e.code === 'TIMEOUT') { if (!firstTimeout) firstTimeout = r; }
        else if (e && e.code === 'NETWORK') { if (!firstNet) firstNet = r; }
        else throw e;
        return next();
      });
    }

    function finish() {
      if (httpFail) return httpFail;
      if (firstNet) return result({
        code: 'NETWORK', url: firstNet, needCertUrl: firstNet + '/saha/',
        error: 'PC\'ye bağlanılamadı. Telefon PC ile aynı ağda olmalı. İlk bağlantıda PC sertifikasını onaylayın ' +
          '("PC sertifikasını onayla", açılan sayfada uyarıyı kabul edin), sonra tekrar deneyin.'
      });
      if (firstTimeout) return result({
        code: 'TIMEOUT', url: firstTimeout,
        error: 'PC yanıt vermedi (' + Math.round(timeoutMs / 1000) + ' sn). PC açık mı, telefon aynı ağda mı?'
      });
      if (mixed) return result({
        code: 'MIXED', url: mixed,
        error: 'Güvenli (https) sayfadan http:// adresine bağlanılamaz. PC adresini https:// ile yazın.'
      });
      return result({ code: 'NO_LAN', error: 'Denenecek PC adresi yok.' });
    }

    return next();
  }

  // ------------------------------------------------------------------ V8: whole list + catalog from the cloud (spec §18)
  /** Every part of the loaded catalog ([] when none). The array is shared: treat it as read-only. A new array after each load / sync. */
  function all() {
    return cur ? cur.parts : [];
  }

  /**
   * Opens a catalog armor that came another way (OneDrive file oms_katalog_<dev>.omsk) with the device secret and stores it
   * exactly like sync() does. secBytes: Uint8Array | b64url string. opts (optional): {dev: expected device id, source: label kept
   * as info().url (default 'onedrive'), now: () => ms}. Resolves with info(); rejects with an Error whose .code is NOT_PAIRED |
   * NO_CRYPTO | BAD_ENVELOPE | BAD_CATALOG and whose message is Turkish. A list older than the one already loaded from the same PC
   * is not taken: the promise then resolves with a copy of the current info() carrying older: true.
   */
  function loadArmor(armor, secBytesArg, opts) {
    var o = opts || {};
    var now = typeof o.now === 'function' ? o.now : function () { return Date.now(); };
    return Promise.resolve().then(function () {
      var sec = secBytes(secBytesArg);
      if (!sec || sec.length !== 32) throw catError('NOT_PAIRED', 'Cihaz eşleştirilmemiş. Önce PC\'deki eşleştirme kodunu okutun.');
      if (!getCrypto() || !Env.supported()) throw catError('NO_CRYPTO', 'Bu tarayıcıda şifreleme (WebCrypto) yok. Sayfayı HTTPS adresinden açın.');
      var text = str(armor).replace(/^\uFEFF/, '').trim();
      if (!text) throw catError('BAD_ENVELOPE', 'Stok listesi dosyası boş.');
      return Env.openFull(text, sec).catch(function (e) {
        var c = e && e.code;
        throw catError('BAD_ENVELOPE', c === 'BadTag'
          ? 'Stok listesi doğrulanamadı. Eşleştirme eski olabilir; yeniden eşleştirin.'
          : 'Stok listesi açılamadı (' + (c || 'hata') + ').');
      }).then(function (r) {
        var dev = str(o.dev).trim();
        if (dev && r.dev !== dev) throw catError('BAD_ENVELOPE', 'Stok listesi başka bir cihaz için hazırlanmış.');
        var rec = recordFromJson(r.json, typeof o.source === 'string' ? o.source : 'onedrive', now());
        var a = cur ? Date.parse(cur.builtAt) : NaN, b = Date.parse(rec.builtAt);
        if (cur && cur.pc === rec.pc && isFinite(a) && isFinite(b) && b < a) {
          var keep = info();
          keep.older = true;
          return keep;
        }
        setCurrent(rec);
        loadPromise = Promise.resolve(info());   // a later load() must not replace the fresh list with the stored one
        return persist(rec).then(function () { return info(); });
      });
    });
  }

  return {
    API_PATH: API_PATH,
    all: all,
    loadArmor: loadArmor,
    rawHash: rawHash,
    signRequest: signRequest,
    sync: sync,
    load: load,
    info: info,
    ageMs: ageMs,
    isStale: isStale,
    lookupRawHash: lookupRawHash,
    lookupRaw: lookupRaw,
    lookupMpn: lookupMpn,
    lookupSpn: lookupSpn,
    describe: describe,
    formatQty: fmtQty,
    onChange: onChange,
    clear: clear,
    normRoot: normRoot,
    normKey: normKey
  };
});
