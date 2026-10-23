/*
 * OMS field mode - direct send through Microsoft Graph + outbox (spec V6 §13, P2).
 *   MSAL.js v5 (Field/msal-browser.min.js: globals `msal` + `msalRedirectBridge`), authority consumers, scope Mail.Send only,
 *   token cache localStorage. POST https://graph.microsoft.com/v1.0/me/sendMail -> 202 = sent.
 *   Outbox: IndexedDB 'oms-outbox' (fallback localStorage, then memory); retried on `online`, page open, every 60 s, tab visible.
 *
 * MSAL v5 note: every interactive/iframe flow returns to `redirectUri`, and that page must run the redirect bridge
 * (msalRedirectBridge.broadcastResponseToMainFrame). Our redirectUri is the page itself, so this file runs the bridge at load time
 * when the URL carries an auth response; `OmsMail.isAuthCallback` is then true and the app should not boot (popup/iframe closes).
 *
 * UMD, no dependencies besides MSAL (and OmsEnvelope if present). Browser: global OmsMail. Node: module.exports (tests inject deps).
 */
(function (root, factory) {
  'use strict';
  if (typeof module === 'object' && module.exports) module.exports = factory(root);
  else root.OmsMail = factory(root);
})(typeof globalThis !== 'undefined' ? globalThis : (typeof self !== 'undefined' ? self : this), function (root) {
  'use strict';

  // ------------------------------------------------------------------ constants
  var AUTHORITY = 'https://login.microsoftonline.com/consumers';
  var SCOPES = ['Mail.Send'];
  var SEND_URL = 'https://graph.microsoft.com/v1.0/me/sendMail';
  var SUBJECT = 'oms_inv_sepet';
  var BEGIN_MARKER = '-----BEGIN OMS SEPET-----';
  var END_MARKER = '-----END OMS SEPET-----';
  var MAIL_COLS = 76;
  var MAX_REQUEST = 3500000;          // Graph JSON request limit is 4 MB; keep a margin
  var SEND_TIMEOUT_MS = 30000;
  var FLUSH_EVERY_MS = 60000;
  var OUTBOX_DB = 'oms-outbox', OUTBOX_STORE = 'items', OUTBOX_LS = 'oms-outbox.v1';
  var GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  var MAIL_RE = /^[^\s@<>()",;:]+@[^\s@<>()",;:]+\.[^\s@<>()",;:]+$/;

  var MSG = {
    CONFIG: 'Microsoft girişi ayarlanmamış (uygulama kimliği / Client ID yok).',
    NO_LIB: 'Microsoft giriş kitaplığı yüklenemedi. Sayfayı yenileyin.',
    LOGIN: 'Microsoft hesabıyla yeniden giriş gerekli.',
    OFFLINE: 'İnternet bağlantısı yok. Sepet giden kutusunda bekleyecek.',
    THROTTLED: 'Microsoft şu an çok fazla istek alıyor. Biraz sonra yeniden denenecek.',
    FORBIDDEN: 'Bu hesabın e-posta gönderme izni (Mail.Send) ya da Outlook posta kutusu yok.',
    BAD_REQUEST: 'Microsoft isteği reddetti.',
    TOO_LARGE: 'Sepet doğrudan gönderim için çok büyük. E-posta uygulamasıyla gönderin.',
    NO_TO: 'Alıcı e-posta adresi yok ya da geçersiz.',
    BAD_ARMOR: 'Gönderilecek sepet verisi geçersiz.',
    SERVER: 'Microsoft sunucusu hata verdi. Yeniden denenecek.',
    BUSY: 'Başka bir Microsoft giriş işlemi sürüyor. Bitince yeniden deneyin.',
    CANCELLED: 'Giriş iptal edildi.',
    UNKNOWN: 'Gönderilemedi.'
  };
  // Codes that describe the whole channel (not the item): flush stops at the first one.
  var GLOBAL_CODES = { CONFIG: 1, NO_LIB: 1, LOGIN: 1, OFFLINE: 1, THROTTLED: 1 };

  // ------------------------------------------------------------------ environment (overridable for Node tests)
  var deps = {
    msal: null, bridge: null,                         // null = take root.msal / root.msalRedirectBridge
    fetch: null, indexedDB: undefined, storage: undefined,
    location: null, navigator: null, now: null
  };
  function lib() { return deps.msal || root.msal || null; }
  function bridgeLib() { return deps.bridge || root.msalRedirectBridge || null; }
  function doFetch(url, opts) { var f = deps.fetch || root.fetch; return f.call(root, url, opts); }
  function nav() { return deps.navigator || root.navigator || null; }
  function loc() { return deps.location || root.location || null; }
  function now() { return deps.now ? deps.now() : Date.now(); }
  function idbFactory() { return deps.indexedDB !== undefined ? deps.indexedDB : (root.indexedDB || null); }
  function lsObj() {
    if (deps.storage !== undefined) return deps.storage;
    try { return root.localStorage || null; } catch (e) { return null; }
  }
  function isOffline() { var n = nav(); return !!(n && n.onLine === false); }

  function result(ok, status, code, error, extra) {
    var r = { ok: !!ok, status: status || 0, code: code || '', error: error || '' };
    if (extra) for (var k in extra) if (Object.prototype.hasOwnProperty.call(extra, k)) r[k] = extra[k];
    return r;
  }
  function fail(code, status, detail, extra) {
    return result(false, status, code, (MSG[code] || MSG.UNKNOWN) + (detail ? ' (' + detail + ')' : ''), extra);
  }

  // ------------------------------------------------------------------ redirect bridge (runs at load)
  /** True when the current URL carries an MSAL auth response (fragment or query) - i.e. we are the popup / iframe / redirect target. */
  function hasAuthResponse(l) {
    if (!l) return false;
    function look(s) {
      if (!s || s.length < 2) return false;
      var p = '&' + s.substring(1) + '&';
      if (p.indexOf('&state=omsd') >= 0) return false;   // OmsDrive's own sign-in response (onedrive.js), not MSAL's
      return p.indexOf('&state=') >= 0 && (p.indexOf('&code=') >= 0 || p.indexOf('&error=') >= 0);
    }
    return look(l.hash || '') || look(l.search || '');
  }

  var authCallback = false, bridgePromise = null;
  function runBridge() {
    if (bridgePromise) return bridgePromise;
    var b = bridgeLib();
    if (!b || typeof b.broadcastResponseToMainFrame !== 'function') return null;
    authCallback = true;
    bridgePromise = Promise.resolve().then(function () { return b.broadcastResponseToMainFrame(); })
      .catch(function (e) {
        // Stale/foreign response (e.g. bookmarked URL with #code=): forget it; the page boots normally next time.
        authCallback = false;
        try { if (root.console) root.console.warn('OmsMail bridge:', (e && (e.errorCode || e.message)) || e); } catch (x) { }
      });
    return bridgePromise;
  }
  if (typeof document !== 'undefined' && hasAuthResponse(loc())) runBridge();

  // ------------------------------------------------------------------ configuration + MSAL client
  var cfg = { clientId: '', redirectUri: '' };
  var pca = null, initPromise = null, activeAcc = null, needLogin = false;

  function defaultRedirectUri() {
    var l = loc();
    if (!l || !l.origin) return '';
    var p = String(l.pathname || '/');
    return l.origin + p.substring(0, p.lastIndexOf('/') + 1);   // directory of the page, no index.html
  }

  /** Sets the Entra app (client) id and redirect URI. Changing them drops the current MSAL client (init() builds a new one). */
  function configure(o) {
    o = o || {};
    var id = String(o.clientId || '').trim();
    var ru = String(o.redirectUri || '').trim() || defaultRedirectUri();
    if (id === cfg.clientId && ru === cfg.redirectUri) return isConfigured();
    cfg.clientId = id; cfg.redirectUri = ru;
    pca = null; initPromise = null; activeAcc = null; needLogin = false;
    return isConfigured();
  }

  function isConfigured() { return GUID_RE.test(cfg.clientId) && !!cfg.redirectUri; }

  function isIos() {
    var n = nav();
    if (!n) return false;
    var ua = String(n.userAgent || '');
    return /iPad|iPhone|iPod/.test(ua) || (n.platform === 'MacIntel' && n.maxTouchPoints > 1);
  }
  function isStandalone() {
    var n = nav();
    if (n && n.standalone === true) return true;
    try { return !!(root.matchMedia && root.matchMedia('(display-mode: standalone)').matches); } catch (e) { return false; }
  }

  function msalConfig() {
    return {
      auth: {
        clientId: cfg.clientId,
        authority: AUTHORITY,
        knownAuthorities: [],
        redirectUri: cfg.redirectUri,
        postLogoutRedirectUri: cfg.redirectUri
      },
      cache: { cacheLocation: 'localStorage' },
      // Microsoft: navigatePopups=false for installed PWAs (no about:blank there).
      system: { navigatePopups: !isStandalone() }
    };
  }

  /**
   * Builds the MSAL client and finishes a pending redirect sign-in. Call once on page open (after configure()).
   * Resolves to account() (or null); never rejects.
   */
  function init() {
    if (authCallback) return Promise.resolve(null);
    if (!isConfigured()) return Promise.resolve(null);
    if (initPromise) return initPromise;
    var m = lib();
    if (!m || typeof m.PublicClientApplication !== 'function') return Promise.resolve(null);
    var p = Promise.resolve().then(function () {
      if (typeof m.createStandardPublicClientApplication === 'function') return m.createStandardPublicClientApplication(msalConfig());
      var c = new m.PublicClientApplication(msalConfig());
      return c.initialize().then(function () { return c; });
    }).then(function (c) {
      if (initPromise !== p) return null;   // reconfigured meanwhile
      pca = c;
      return c.handleRedirectPromise().catch(function (e) {
        warn('handleRedirectPromise', e);
        return null;
      }).then(function (res) {
        var acc = (res && res.account) || c.getActiveAccount() || firstAccount(c);
        setActive(acc);
        if (res && res.account) { needLogin = false; emit({ type: 'account', basketId: '' }); }
        return account();
      });
    }).catch(function (e) {
      warn('init', e);
      if (initPromise === p) initPromise = null;   // allow a retry later
      return null;
    });
    initPromise = p;
    return p;
  }

  function firstAccount(c) {
    try { var all = c.getAllAccounts() || []; return all.length ? all[0] : null; } catch (e) { return null; }
  }
  function setActive(acc) {
    activeAcc = acc || null;
    if (pca) { try { pca.setActiveAccount(activeAcc); } catch (e) { } }
  }
  function warn(where, e) {
    try { if (root.console) root.console.warn('OmsMail ' + where + ':', (e && (e.errorCode || e.message)) || e); } catch (x) { }
  }

  /** Signed-in account {username, name} or null. */
  function account() {
    var a = activeAcc;
    if (!a && pca) { try { a = pca.getActiveAccount() || firstAccount(pca); activeAcc = a; } catch (e) { a = null; } }
    return a ? { username: a.username || '', name: a.name || a.username || '' } : null;
  }

  /** True after a send/flush found that a fresh interactive sign-in is required. */
  function loginNeeded() { return needLogin; }

  function errCode(e) { return String((e && (e.errorCode || e.code)) || ''); }
  function errName(e) { return String((e && e.name) || ''); }

  function isInteractionRequired(e) {
    var m = lib(), code = errCode(e);
    if (m && m.InteractionRequiredAuthError && e instanceof m.InteractionRequiredAuthError) return true;
    if (errName(e) === 'InteractionRequiredAuthError') return true;
    var codes = (m && m.InteractionRequiredAuthErrorCodes) || {};
    for (var k in codes) if (codes[k] === code) return true;
    // no account in cache / iframe renewal blocked (third-party cookies) / token cache unreadable after browser restart
    return /^(interaction_required|login_required|consent_required|no_tokens_found|refresh_token_expired|bad_token|no_account_error|monitor_window_timeout|timed_out|iframe_closed_prematurely|block_iframe_reload|unable_to_load_token|crypto_key_not_found)$/.test(code);
  }
  function isNetworkError(e) {
    var code = errCode(e);
    return /^(no_network_connectivity|post_request_failed|get_request_failed|network_error|endpoints_resolution_error)$/.test(code) ||
      (errName(e) === 'TypeError') || isOffline();
  }

  /**
   * Interactive sign-in. Popup on desktop/Android; full-page redirect on iOS (WebKit) or when the popup is blocked.
   * Call straight from a click handler (the popup must open inside the user gesture).
   * Resolves {ok, account, redirecting, code, error}; never rejects.
   */
  function signIn() {
    if (!isConfigured()) return Promise.resolve(fail('CONFIG'));
    if (!lib()) return Promise.resolve(fail('NO_LIB'));
    if (isOffline()) return Promise.resolve(fail('OFFLINE'));
    var req = { scopes: SCOPES.slice(), prompt: 'select_account' };
    // No await before loginPopup: keeps the click's user activation for window.open.
    if (pca && !isIos()) return popupLogin(req);
    // Client not built yet (or iOS): after the await the activation may be gone, so a blocked popup falls back to redirect.
    return init().then(function () {
      if (!pca) return fail('NO_LIB');
      return isIos() ? redirectLogin(req) : popupLogin(req);
    });
  }

  function popupLogin(req) {
    var p;
    try { p = pca.loginPopup(req); } catch (e) { p = Promise.reject(e); }
    return Promise.resolve(p).then(signedIn, function (e) {
      var c = errCode(e);
      if (c === 'popup_window_error' || c === 'empty_window_error' || c === 'block_nested_popups') return redirectLogin(req);
      return loginError(e);
    });
  }

  function redirectLogin(req) {
    var p;
    try { p = pca.loginRedirect(req); } catch (e) { return Promise.resolve(loginError(e)); }
    // loginRedirect navigates away; its promise normally never settles before unload.
    return Promise.resolve(p).then(function () { return result(true, 0, '', '', { account: null, redirecting: true }); }, loginError);
  }

  function signedIn(res) {
    setActive(res && res.account);
    needLogin = false;
    emit({ type: 'account', basketId: '' });
    if (started) setTimeout(function () { flush(); }, 0);
    return result(true, 0, '', '', { account: account(), redirecting: false });
  }

  function loginError(e) {
    var c = errCode(e);
    if (c === 'user_cancelled') return fail('CANCELLED');
    if (c === 'interaction_in_progress') return fail('BUSY');
    if (isNetworkError(e)) return fail('OFFLINE');
    warn('signIn', e);
    return fail('UNKNOWN', 0, c || (e && e.message) || '');
  }

  /** Forgets the account on this device (local token cache only; does not sign out of Microsoft everywhere). */
  function signOut() {
    var acc = activeAcc;
    var p = Promise.resolve();
    if (pca && acc) {
      p = Promise.resolve().then(function () { return pca.clearCache({ account: acc }); }).catch(function (e) { warn('signOut', e); });
    }
    return p.then(function () { setActive(null); needLogin = false; emit({ type: 'account', basketId: '' }); });
  }

  // ------------------------------------------------------------------ token + Graph sendMail
  function getToken(forceRefresh) {
    if (!pca) return Promise.reject(Object.assign(new Error('no client'), { errorCode: 'no_account_error' }));
    var acc = activeAcc || (function () { try { return pca.getActiveAccount() || firstAccount(pca); } catch (e) { return null; } })();
    if (!acc) return Promise.reject(Object.assign(new Error('no account'), { errorCode: 'no_account_error' }));
    return pca.acquireTokenSilent({ scopes: SCOPES.slice(), account: acc, forceRefresh: !!forceRefresh })
      .then(function (r) {
        if (!r || !r.accessToken) throw Object.assign(new Error('empty token'), { errorCode: 'no_tokens_found' });
        if (r.account) activeAcc = r.account;
        return r.accessToken;
      });
  }

  function tokenFailure(e) {
    if (isNetworkError(e)) return fail('OFFLINE');
    if (isInteractionRequired(e)) { needLogin = true; return fail('LOGIN'); }
    warn('token', e);
    return fail('UNKNOWN', 0, errCode(e) || (e && e.message) || '');
  }

  function wrapForMail(armor) {
    var E = root.OmsEnvelope;
    if (E && typeof E.wrapForMail === 'function') return E.wrapForMail(armor);
    var a = String(armor).replace(/\s+/g, ''), lines = [BEGIN_MARKER];
    for (var i = 0; i < a.length; i += MAIL_COLS) lines.push(a.substring(i, i + MAIL_COLS));
    lines.push(END_MARKER);
    return lines.join('\r\n') + '\r\n';
  }

  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function defaultFileName() {
    var E = root.OmsEnvelope;
    if (E && typeof E.fileName === 'function') return E.fileName();
    var d = new Date(now());
    return 'oms_sepet_' + d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate()) + '_' +
      pad2(d.getHours()) + pad2(d.getMinutes()) + pad2(d.getSeconds()) + '.omsb';
  }

  function b64Ascii(s) {
    if (typeof root.btoa === 'function') return root.btoa(s);
    return Buffer.from(s, 'latin1').toString('base64');   // Node without btoa
  }

  /** Graph request body (spec §13). Attachment is dropped when the request would pass the size limit (the body still has the armor). */
  function buildMessage(to, armor, fileName, subject) {
    var content = wrapForMail(armor);
    var msg = {
      subject: subject || SUBJECT,
      body: { contentType: 'Text', content: content },
      toRecipients: [{ emailAddress: { address: to } }]
    };
    var attB64 = Math.ceil(armor.length / 3) * 4;
    if (content.length + attB64 + 2048 <= MAX_REQUEST) {
      msg.attachments = [{
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: fileName || defaultFileName(),
        contentType: 'text/plain',
        contentBytes: b64Ascii(armor)
      }];
    }
    return { message: msg, saveToSentItems: false };
  }

  function retryAfterMs(resp) {
    var h = resp && resp.headers && typeof resp.headers.get === 'function' ? resp.headers.get('Retry-After') : null;
    var s = parseInt(h, 10);
    return (isFinite(s) && s > 0 ? Math.min(s, 3600) : 60) * 1000;
  }

  function graphDetail(resp) {
    return Promise.resolve().then(function () { return resp.text(); }).then(function (t) {
      try { var j = JSON.parse(t); return (j && j.error && (j.error.code + ': ' + j.error.message)) || ''; } catch (e) { return ''; }
    }).catch(function () { return ''; }).then(function (d) { return String(d).substring(0, 200); });
  }

  function postMail(token, body) {
    var ctl = typeof AbortController === 'function' ? new AbortController() : null;
    var t = ctl ? setTimeout(function () { try { ctl.abort(); } catch (e) { } }, SEND_TIMEOUT_MS) : null;
    var opts = {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: body, cache: 'no-store', credentials: 'omit', mode: 'cors'
    };
    if (ctl) opts.signal = ctl.signal;
    return Promise.resolve().then(function () { return doFetch(SEND_URL, opts); })
      .then(function (r) { if (t) clearTimeout(t); return r; }, function (e) { if (t) clearTimeout(t); throw e; });
  }

  /**
   * Sends one basket armor to `to` (subject oms_inv_sepet). Resolves {ok, status, error, code}; never rejects.
   * code: '' (sent) | CONFIG | NO_LIB | LOGIN | OFFLINE | THROTTLED (+retryAfter ms) | FORBIDDEN | BAD_REQUEST | TOO_LARGE |
   *       NO_TO | BAD_ARMOR | SERVER | UNKNOWN.
   */
  function send(o) {
    o = o || {};
    var to = String(o.to || '').trim();
    var armor = String(o.armor || '').replace(/\s+/g, '');
    if (!MAIL_RE.test(to)) return Promise.resolve(fail('NO_TO'));
    if (armor.indexOf('OMSB1:') !== 0 || /[^\x21-\x7e]/.test(armor)) return Promise.resolve(fail('BAD_ARMOR'));
    if (wrapForMail(armor).length + 2048 > MAX_REQUEST) return Promise.resolve(fail('TOO_LARGE'));
    if (!isConfigured()) return Promise.resolve(fail('CONFIG'));
    if (!lib()) return Promise.resolve(fail('NO_LIB'));
    if (isOffline()) return Promise.resolve(fail('OFFLINE'));
    var body = JSON.stringify(buildMessage(to, armor, o.fileName, o.subject));

    function attempt(forceRefresh) {
      return getToken(forceRefresh).then(function (token) {
        return postMail(token, body).then(function (resp) {
          var st = resp.status;
          if (st === 202 || st === 200) { needLogin = false; return result(true, st); }
          if (st === 401 && !forceRefresh) return attempt(true);
          if (st === 401) { needLogin = true; return fail('LOGIN', st); }
          if (st === 429 || st === 503 || st === 504) return fail('THROTTLED', st, '', { retryAfter: retryAfterMs(resp) });
          if (st === 413) return fail('TOO_LARGE', st);
          return graphDetail(resp).then(function (d) {
            if (st === 403) return fail('FORBIDDEN', st, d);
            if (st >= 400 && st < 500) return fail('BAD_REQUEST', st, d || ('HTTP ' + st));
            return fail('SERVER', st, 'HTTP ' + st);
          });
        }, function (e) {
          // fetch rejects with TypeError (no network / CORS / TLS) or AbortError (timeout)
          if (errName(e) === 'AbortError') return fail('OFFLINE', 0, 'zaman aşımı');
          return fail('OFFLINE');
        });
      }, tokenFailure);
    }
    return init().then(function () { return attempt(false); }).catch(function (e) {
      warn('send', e);
      return fail('UNKNOWN', 0, (e && e.message) || '');
    });
  }

  // ------------------------------------------------------------------ outbox storage
  var store = { mode: '', db: null, mem: {}, opening: null };

  function openStore() {
    if (store.opening) return store.opening;
    store.opening = new Promise(function (res) {
      var done = false;
      function fin(mode) { if (done) return; done = true; store.mode = mode; res(mode); }
      function fallback() { fin(lsUsable() ? 'ls' : 'mem'); }
      var f = idbFactory();
      if (!f) { fallback(); return; }
      var rq;
      try { rq = f.open(OUTBOX_DB, 1); } catch (e) { fallback(); return; }
      rq.onupgradeneeded = function () {
        try { if (!rq.result.objectStoreNames.contains(OUTBOX_STORE)) rq.result.createObjectStore(OUTBOX_STORE, { keyPath: 'basketId' }); } catch (e) { }
      };
      rq.onsuccess = function () {
        store.db = rq.result;
        store.db.onversionchange = function () { try { store.db.close(); } catch (e) { } store.db = null; store.opening = null; };
        fin('idb');
      };
      rq.onerror = fallback;
      rq.onblocked = fallback;
      setTimeout(fallback, 3000);   // some private modes never answer
    });
    return store.opening;
  }

  function lsUsable() {
    var ls = lsObj();
    if (!ls) return false;
    try { ls.setItem('__oms_ob', '1'); ls.removeItem('__oms_ob'); return true; } catch (e) { return false; }
  }
  function lsAll() {
    var ls = lsObj(), v = null;
    try { v = ls ? ls.getItem(OUTBOX_LS) : null; } catch (e) { v = null; }
    if (!v) return {};
    try { var o = JSON.parse(v); return o && typeof o === 'object' ? o : {}; } catch (e) { return {}; }
  }
  function lsSave(map) {
    var ls = lsObj();
    try { ls.setItem(OUTBOX_LS, JSON.stringify(map)); return true; } catch (e) { return false; }   // quota: caller keeps a memory copy
  }

  function idbReq(mode, fn) {
    return new Promise(function (res, rej) {
      if (!store.db) { rej(new Error('db closed')); return; }
      var tx, out;
      try { tx = store.db.transaction(OUTBOX_STORE, mode); out = fn(tx.objectStore(OUTBOX_STORE)); } catch (e) { rej(e); return; }
      tx.oncomplete = function () { res(out && 'result' in out ? out.result : undefined); };
      tx.onerror = tx.onabort = function () { rej(tx.error || new Error('tx failed')); };
    });
  }

  function dbAll() {
    return openStore().then(function (mode) {
      if (mode === 'idb') return idbReq('readonly', function (s) { return s.getAll(); }).catch(function () { return memList(); });
      if (mode === 'ls') { var m = lsAll(), out = []; for (var k in m) if (Object.prototype.hasOwnProperty.call(m, k)) out.push(m[k]); return mergeMem(out); }
      return memList();
    });
  }
  function dbPut(item) {
    return openStore().then(function (mode) {
      if (mode === 'idb') return idbReq('readwrite', function (s) { return s.put(item); }).then(function () { delete store.mem[item.basketId]; },
        function () { store.mem[item.basketId] = item; });
      if (mode === 'ls') { var m = lsAll(); m[item.basketId] = item; if (lsSave(m)) delete store.mem[item.basketId]; else store.mem[item.basketId] = item; return; }
      store.mem[item.basketId] = item;
    });
  }
  function dbDel(id) {
    return openStore().then(function (mode) {
      delete store.mem[id];
      if (mode === 'idb') return idbReq('readwrite', function (s) { return s.delete(id); }).catch(function () { });
      if (mode === 'ls') { var m = lsAll(); if (m[id]) { delete m[id]; lsSave(m); } }
    });
  }
  function memList() { var out = []; for (var k in store.mem) if (Object.prototype.hasOwnProperty.call(store.mem, k)) out.push(store.mem[k]); return out; }
  function mergeMem(list) {
    var seen = {}, i;
    for (i = 0; i < list.length; i++) seen[list[i].basketId] = 1;
    var mem = memList();
    for (i = 0; i < mem.length; i++) if (!seen[mem[i].basketId]) list.push(mem[i]);
    return list;
  }
  function sortItems(list) {
    return list.sort(function (a, b) { return (a.queuedAt || 0) - (b.queuedAt || 0); });
  }

  // ------------------------------------------------------------------ outbox API
  var listeners = [];
  function emit(ev) {
    for (var i = 0; i < listeners.length; i++) { try { listeners[i](ev); } catch (e) { warn('onChange', e); } }
  }
  /** Subscribes to outbox / send events: {type:'queued'|'sent'|'failed'|'login'|'removed'|'account', basketId[, code, error]}. Returns an unsubscribe function. */
  function onChange(cb) {
    if (typeof cb !== 'function') return function () { };
    listeners.push(cb);
    return function () { var i = listeners.indexOf(cb); if (i >= 0) listeners.splice(i, 1); };
  }

  /** Queues a basket (same basketId replaces the queued copy). Resolves to the stored item. */
  function enqueue(o) {
    o = o || {};
    var id = String(o.basketId || '').trim();
    var armor = String(o.armor || '').replace(/\s+/g, '');
    if (!id) return Promise.reject(new Error('basketId gerekli'));
    if (armor.indexOf('OMSB1:') !== 0) return Promise.reject(new Error(MSG.BAD_ARMOR));
    return dbAll().then(function (list) {
      var prev = null;
      for (var i = 0; i < list.length; i++) if (list[i].basketId === id) prev = list[i];
      var item = {
        basketId: id, to: String(o.to || '').trim(), armor: armor, fileName: String(o.fileName || '') || (prev && prev.fileName) || defaultFileName(),
        queuedAt: prev ? prev.queuedAt : now(), tries: prev ? prev.tries : 0, lastTry: prev ? prev.lastTry : 0, lastCode: '', lastError: ''
      };
      return dbPut(item).then(function () { emit({ type: 'queued', basketId: id }); return publicItem(item); });
    });
  }

  function publicItem(it) {
    return {
      basketId: it.basketId, to: it.to, fileName: it.fileName, queuedAt: it.queuedAt, tries: it.tries || 0,
      lastTry: it.lastTry || 0, lastCode: it.lastCode || '', lastError: it.lastError || '', armor: it.armor
    };
  }

  /** Queued items, oldest first: [{basketId, to, fileName, queuedAt, tries, lastTry, lastCode, lastError, armor}]. */
  function pending() {
    return dbAll().then(function (list) { return sortItems(list).map(publicItem); });
  }

  /** Drops a queued basket (e.g. the user sent it another way). */
  function remove(basketId) {
    var id = String(basketId || '');
    return dbDel(id).then(function () { emit({ type: 'removed', basketId: id }); });
  }

  var flushing = null, holdUntil = 0;

  /**
   * Sends every queued basket (single flight: a call during a running flush gets the same promise).
   * Stops at the first channel-wide problem (offline, sign-in needed, throttled). Resolves {sent:[basketId], failed:[{basketId, code, error}]}.
   */
  function flush() {
    if (flushing) return flushing;
    flushing = dbAll().then(function (list) {
      var items = sortItems(list), out = { sent: [], failed: [] }, i = 0;
      if (!items.length) return out;
      var gate = !isConfigured() ? 'CONFIG' : !lib() ? 'NO_LIB' : isOffline() ? 'OFFLINE' : now() < holdUntil ? 'THROTTLED' : '';
      if (gate) return skipAll(items, out, fail(gate), false);
      return init().then(function () {
        if (!account()) { needLogin = true; return skipAll(items, out, fail('LOGIN'), true); }
        function next() {
          if (i >= items.length) return out;
          var it = items[i++];
          return send({ to: it.to, armor: it.armor, fileName: it.fileName }).then(function (r) {
            if (r.ok) {
              out.sent.push(it.basketId);
              return dbDel(it.basketId).then(function () { emit({ type: 'sent', basketId: it.basketId }); return next(); });
            }
            it.tries = (it.tries || 0) + 1; it.lastTry = now(); it.lastCode = r.code; it.lastError = r.error;
            out.failed.push({ basketId: it.basketId, code: r.code, error: r.error });
            if (r.code === 'THROTTLED') holdUntil = now() + (r.retryAfter || 60000);
            return dbPut(it).then(function () {
              emit({ type: r.code === 'LOGIN' ? 'login' : 'failed', basketId: it.basketId, code: r.code, error: r.error });
              if (GLOBAL_CODES[r.code]) return skipAll(items.slice(i), out, r, false);
              return next();
            });
          });
        }
        return next();
      });
    }).catch(function (e) {
      warn('flush', e);
      return { sent: [], failed: [] };
    }).then(function (out) { flushing = null; return out; });
    return flushing;
  }

  // Items not attempted because the channel is down: reported, not counted as tries. `emitLogin` raises one 'login' event.
  function skipAll(items, out, r, emitLogin) {
    for (var k = 0; k < items.length; k++) out.failed.push({ basketId: items[k].basketId, code: r.code, error: r.error });
    if (emitLogin && r.code === 'LOGIN' && items.length) emit({ type: 'login', basketId: items[0].basketId, code: r.code, error: r.error });
    return out;
  }

  // ------------------------------------------------------------------ automatic retry
  var started = false, timer = null, handlers = null;

  /** Starts automatic outbox delivery: now, on `online`, when the tab becomes visible, and every 60 s. Idempotent. */
  function start() {
    if (started || authCallback) return;
    started = true;
    function kick() { if (!isOffline()) flush(); }
    handlers = {
      online: function () { holdUntil = 0; kick(); },
      vis: function () { if (typeof document !== 'undefined' && document.visibilityState === 'visible') kick(); }
    };
    if (typeof root.addEventListener === 'function') root.addEventListener('online', handlers.online);
    if (typeof document !== 'undefined' && document.addEventListener) document.addEventListener('visibilitychange', handlers.vis);
    timer = setInterval(function () {
      if (typeof document !== 'undefined' && document.visibilityState === 'hidden') return;
      kick();
    }, FLUSH_EVERY_MS);
    setTimeout(kick, 0);
  }

  /** Stops automatic delivery (tests / sign-out). */
  function stop() {
    if (!started) return;
    started = false;
    if (timer) clearInterval(timer);
    timer = null;
    if (handlers) {
      if (typeof root.removeEventListener === 'function') root.removeEventListener('online', handlers.online);
      if (typeof document !== 'undefined' && document.removeEventListener) document.removeEventListener('visibilitychange', handlers.vis);
    }
    handlers = null;
  }

  // ------------------------------------------------------------------ test hooks
  /** Node tests: inject {msal, bridge, fetch, indexedDB, storage, location, navigator, now}; resets all state. */
  function _setDeps(d) {
    stop();
    d = d || {};
    for (var k in deps) if (Object.prototype.hasOwnProperty.call(deps, k)) deps[k] = Object.prototype.hasOwnProperty.call(d, k) ? d[k] : (k === 'indexedDB' || k === 'storage' ? undefined : null);
    cfg = { clientId: '', redirectUri: '' };
    pca = null; initPromise = null; activeAcc = null; needLogin = false;
    authCallback = false; bridgePromise = null;
    store = { mode: '', db: null, mem: {}, opening: null };
    flushing = null; holdUntil = 0; listeners = [];
  }

  var api = {
    SCOPES: SCOPES.slice(),
    AUTHORITY: AUTHORITY,
    SEND_URL: SEND_URL,
    configure: configure,
    isConfigured: isConfigured,
    init: init,
    account: account,
    loginNeeded: loginNeeded,
    signIn: signIn,
    signOut: signOut,
    send: send,
    enqueue: enqueue,
    pending: pending,
    flush: flush,
    start: start,
    stop: stop,
    onChange: onChange,
    remove: remove,
    buildMessage: buildMessage,
    hasAuthResponse: hasAuthResponse,
    handleAuthCallback: function () { return hasAuthResponse(loc()) ? runBridge() : null; },
    _setDeps: _setDeps
  };
  Object.defineProperty(api, 'isAuthCallback', { enumerable: true, get: function () { return authCallback; } });
  Object.defineProperty(api, 'storageMode', { enumerable: true, get: function () { return store.mode; } });
  return api;
});
