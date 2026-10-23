/*
 * OMS field mode - direct OneDrive transport (spec V8 §17 P2, §18 P3). No share sheet, no folder picking.
 *   MSAL.js v5 (Field/msal-browser.min.js: globals `msal` + `msalRedirectBridge`), authority common (personal + work/school),
 *   scope Files.ReadWrite (+ openid/profile/offline_access added by MSAL), token cache localStorage.
 *   Upload:   PUT  https://graph.microsoft.com/v1.0/me/drive/root:/OMS Saha/<name>:/content?@microsoft.graph.conflictBehavior=rename
 *             Content-Type text/plain, body = armor text; 200/201 = sent. The first PUT creates the "OMS Saha" folder.
 *   Download: GET  .../root:/OMS Saha/<name>  -> @microsoft.graph.downloadUrl, then GET that URL without Authorization.
 *             Microsoft (driveitem-get-content, "Downloading files in JavaScript apps"): /content answers 302, and a 302 after a
 *             CORS preflight is prohibited, so browser apps must use downloadUrl (pre-authenticated, no preflight).
 *   Outbox:   IndexedDB 'oms-drive-outbox' (fallback localStorage, then memory); retried on `online`, page open, every 60 s, tab visible.
 *
 * MSAL v5: popup / iframe / redirect all come back to `redirectUri` (the page directory), and that page must run the redirect
 * bridge exactly once. Field/sendmail.js (OmsMail) runs it too; whichever file loads first owns it, the other one follows
 * (load sendmail.js before onedrive.js). When `OmsDrive.isAuthCallback` (or OmsMail.isAuthCallback) is true the app must not boot.
 *
 * UMD. Browser: global OmsDrive. Node: module.exports (tests inject deps with _setDeps).
 */
(function (root, factory) {
  'use strict';
  if (typeof module === 'object' && module.exports) module.exports = factory(root);
  else root.OmsDrive = factory(root);
})(typeof globalThis !== 'undefined' ? globalThis : (typeof self !== 'undefined' ? self : this), function (root) {
  'use strict';

  // ------------------------------------------------------------------ app registration (filled by the main session)
  // Entra app (client) id. Registration: platform "Single-page application", redirect = the page directory,
  // supported accounts = "any organizational directory and personal Microsoft accounts", delegated Files.ReadWrite.
  // Empty -> available() is false and the page hides the OneDrive feature.
  var CLIENT_ID = '7869b178-008b-462b-a983-7e3f9c58f79d';   // Entra app "OMS Saha" (tenant sayzerhotmail, personal+work, SPA redirects: Pages + PC /saha/)

  // ------------------------------------------------------------------ constants
  var AUTHORITY = 'https://login.microsoftonline.com/common';
  var SCOPES = ['Files.ReadWrite'];
  var GRAPH = 'https://graph.microsoft.com/v1.0';
  var FOLDER = 'OMS Saha';
  var REQUEST_TIMEOUT_MS = 30000;
  var FLUSH_EVERY_MS = 60000;
  var MAX_UPLOAD = 200 * 1024 * 1024;   // simple PUT supports up to 250 MB
  var OUTBOX_DB = 'oms-drive-outbox', OUTBOX_STORE = 'items', OUTBOX_LS = 'oms-drive-outbox.v1';
  var GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  // OneDrive forbids " * : < > ? / \ | and control chars; also no leading/trailing space or dot-only names.
  var NAME_BAD_RE = /["*:<>?\/\\|\x00-\x1f\x7f]/;

  // User-visible texts: plain Turkish, no technical words (details go to `detail`).
  var MSG = {
    CONFIG: 'OneDrive bağlantısı bu sayfada hazır değil.',
    NO_LIB: 'OneDrive bağlantısı yüklenemedi. Sayfayı yenileyin.',
    LOGIN: 'OneDrive\'a yeniden bağlanın.',
    OFFLINE: 'İnternet yok. İnternet gelince kendiliğinden gider.',
    THROTTLED: 'OneDrive şu an meşgul. Biraz sonra kendiliğinden yeniden denenecek.',
    QUOTA: 'OneDrive\'da yer kalmadı. Biraz yer açın; sonra kendiliğinden gider.',
    FORBIDDEN: 'Bu hesap OneDrive\'a dosya yazamıyor. Başka bir hesapla bağlanın.',
    NO_DRIVE: 'Bu hesabın OneDrive alanı yok. Başka bir hesapla bağlanın.',
    NOT_FOUND: 'Dosya OneDrive\'da bulunamadı.',
    BAD_REQUEST: 'OneDrive bu dosyayı kabul etmedi.',
    BAD_NAME: 'Dosya adı geçersiz.',
    BAD_DATA: 'Gönderilecek veri boş ya da çok büyük.',
    SERVER: 'OneDrive hata verdi. Kendiliğinden yeniden denenecek.',
    BUSY: 'Başka bir bağlanma işlemi sürüyor. Bitince yeniden deneyin.',
    CANCELLED: 'Bağlanma iptal edildi.',
    UNKNOWN: 'İşlem tamamlanamadı. Yeniden deneyin.'
  };
  // Codes that describe the whole channel (not one item): flush stops at the first one.
  var GLOBAL_CODES = { CONFIG: 1, NO_LIB: 1, LOGIN: 1, OFFLINE: 1, THROTTLED: 1, QUOTA: 1, FORBIDDEN: 1, NO_DRIVE: 1, SERVER: 1 };
  // Codes worth keeping in the outbox (send() queues on these).
  var RETRY_CODES = { CONFIG: 1, NO_LIB: 1, LOGIN: 1, OFFLINE: 1, THROTTLED: 1, QUOTA: 1, SERVER: 1, BUSY: 1, UNKNOWN: 1, FORBIDDEN: 1, NO_DRIVE: 1 };

  // ------------------------------------------------------------------ environment (overridable for Node tests)
  var deps = {
    msal: null, bridge: null,                         // null = take root.msal / root.msalRedirectBridge
    fetch: null, indexedDB: undefined, storage: undefined,
    location: null, navigator: null, now: null, clientId: null
  };
  function lib() { return deps.msal || root.msal || null; }
  function bridgeLib() { return deps.bridge || root.msalRedirectBridge || null; }
  function doFetch(url, opts) { var f = deps.fetch || root.fetch; return f.call(root, url, opts); }
  function nav() { return deps.navigator || root.navigator || null; }
  function loc() { return deps.location || root.location || null; }
  function now() { return deps.now ? deps.now() : Date.now(); }
  function clientId() { return String(deps.clientId != null ? deps.clientId : CLIENT_ID).trim(); }
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
    var r = result(false, status, code, MSG[code] || MSG.UNKNOWN, extra);
    if (detail) r.detail = String(detail).substring(0, 300);
    return r;
  }
  function warn(where, e) {
    try { if (root.console) root.console.warn('OmsDrive ' + where + ':', (e && (e.errorCode || e.message)) || e); } catch (x) { }
  }

  // ------------------------------------------------------------------ redirect bridge (runs at load, shared with OmsMail)
  /** True when the URL carries an MSAL auth response (fragment or query): we are the popup / iframe / redirect target. */
  function hasAuthResponse(l) {
    if (!l) return false;
    function look(s) {
      if (!s || s.length < 2) return false;
      var p = '&' + s.substring(1) + '&';
      return p.indexOf('&state=') >= 0 && (p.indexOf('&code=') >= 0 || p.indexOf('&error=') >= 0);
    }
    return look(l.hash || '') || look(l.search || '');
  }

  var authCallback = false, bridgePromise = null, mailOwnsBridge = false;
  function mailCallback() { var m = root.OmsMail; return !!(m && m !== api && m.isAuthCallback); }
  function runBridge() {
    if (bridgePromise) return bridgePromise;
    // OmsMail already runs the bridge for this page load: never call it twice (the second call fails, URL is already cleaned).
    if (mailCallback()) { mailOwnsBridge = true; bridgePromise = Promise.resolve(); return bridgePromise; }
    var b = bridgeLib();
    if (!b || typeof b.broadcastResponseToMainFrame !== 'function') return null;
    authCallback = true;
    bridgePromise = Promise.resolve().then(function () { return b.broadcastResponseToMainFrame(); })
      .catch(function (e) {
        // Stale/foreign response (e.g. bookmarked URL with #code=): forget it; the page boots normally next time.
        authCallback = false;
        warn('bridge', e);
      });
    return bridgePromise;
  }
  function inCallback() { return authCallback || (mailOwnsBridge && mailCallback()); }
  if (typeof document !== 'undefined') {
    if (mailCallback()) { mailOwnsBridge = true; bridgePromise = Promise.resolve(); }
    else if (hasAuthResponse(loc())) runBridge();
  }

  // ------------------------------------------------------------------ MSAL client
  var pca = null, initPromise = null, activeAcc = null, needLogin = false;

  function redirectUri() {
    var l = loc();
    if (!l || !l.origin) return '';
    var p = String(l.pathname || '/');
    return l.origin + p.substring(0, p.lastIndexOf('/') + 1);   // page directory, no index.html (same as OmsMail)
  }

  /** True when this page can use OneDrive at all (app id set and the sign-in library loaded). */
  function available() {
    var m = lib();
    return GUID_RE.test(clientId()) && !!redirectUri() && !!(m && typeof m.PublicClientApplication === 'function');
  }

  function isIos() {
    var n = nav();
    if (!n) return false;
    var ua = String(n.userAgent || '');
    return /iPad|iPhone|iPod/.test(ua) || (n.platform === 'MacIntel' && n.maxTouchPoints > 1);
  }
  // Field report 2026-10-23: in the installed (home-screen) app on Android the sign-in popup opened a separate Chrome tab the
  // page could not drive — Chrome just showed its last page. Phones and installed apps sign in by full-page redirect.
  function preferRedirect() {
    var n = nav(), ua = String((n && n.userAgent) || '');
    return isIos() || isStandalone() || /Android|Mobile/i.test(ua);
  }
  function isStandalone() {
    var n = nav();
    if (n && n.standalone === true) return true;
    try { return !!(root.matchMedia && root.matchMedia('(display-mode: standalone)').matches); } catch (e) { return false; }
  }

  function msalConfig() {
    var ru = redirectUri();
    return {
      auth: {
        clientId: clientId(),
        authority: AUTHORITY,
        knownAuthorities: [],
        redirectUri: ru,
        postLogoutRedirectUri: ru
      },
      cache: { cacheLocation: 'localStorage' },
      // Microsoft: navigatePopups=false for installed PWAs (no about:blank there).
      system: { navigatePopups: !isStandalone() }
    };
  }

  /**
   * Builds the MSAL client and finishes a pending redirect sign-in. Call once on page open.
   * Resolves to account() (or null); never rejects.
   */
  function init() {
    if (inCallback() || !available()) return Promise.resolve(null);
    if (initPromise) return initPromise;
    var m = lib();
    var p = Promise.resolve().then(function () {
      if (typeof m.createStandardPublicClientApplication === 'function') return m.createStandardPublicClientApplication(msalConfig());
      var c = new m.PublicClientApplication(msalConfig());
      return c.initialize().then(function () { return c; });
    }).then(function (c) {
      if (initPromise !== p) return null;   // reset meanwhile
      pca = c;
      return Promise.resolve().then(function () { return c.handleRedirectPromise(); }).catch(function (e) {
        warn('handleRedirectPromise', e);
        return null;
      }).then(function (res) {
        var acc = (res && res.account) || safeActive(c);
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

  function safeActive(c) {
    try {
      var a = c.getActiveAccount();
      if (a) return a;
      var all = c.getAllAccounts() || [];
      return all.length ? all[0] : null;
    } catch (e) { return null; }
  }
  function setActive(acc) {
    activeAcc = acc || null;
    if (pca) { try { pca.setActiveAccount(activeAcc); } catch (e) { } }
  }

  /** Connected account {username, name} or null. */
  function account() {
    var a = activeAcc;
    if (!a && pca) { a = safeActive(pca); activeAcc = a; }
    return a ? { username: a.username || '', name: a.name || a.username || '' } : null;
  }

  /** True after an upload/flush found that the user must connect again (e.g. the 24 h sign-in limit of browser apps). */
  function loginNeeded() { return needLogin; }

  function errCode(e) { return String((e && (e.errorCode || e.code)) || ''); }
  function errName(e) { return String((e && e.name) || ''); }

  function isInteractionRequired(e) {
    var m = lib(), code = errCode(e);
    if (m && m.InteractionRequiredAuthError && e instanceof m.InteractionRequiredAuthError) return true;
    if (errName(e) === 'InteractionRequiredAuthError') return true;
    var codes = (m && m.InteractionRequiredAuthErrorCodes) || {};
    for (var k in codes) if (codes[k] === code) return true;
    // no account / iframe renewal blocked (third-party cookies, iOS) / token cache unreadable after a browser restart
    return /^(interaction_required|login_required|consent_required|no_tokens_found|refresh_token_expired|bad_token|no_account_error|monitor_window_timeout|timed_out|iframe_closed_prematurely|block_iframe_reload|unable_to_load_token|crypto_key_not_found)$/.test(code);
  }
  function isNetworkError(e) {
    var code = errCode(e);
    return /^(no_network_connectivity|post_request_failed|get_request_failed|network_error|endpoints_resolution_error)$/.test(code) ||
      errName(e) === 'TypeError' || isOffline();
  }

  /**
   * Interactive connect. Popup on desktop/Android; full-page redirect on iOS (WebKit) or when the popup is blocked.
   * Call straight from a click handler (the popup must open inside the user gesture).
   * Resolves {ok, account, redirecting, code, error}; never rejects.
   */
  function signIn() {
    if (!GUID_RE.test(clientId())) return Promise.resolve(fail('CONFIG'));
    if (!lib()) return Promise.resolve(fail('NO_LIB'));
    if (isOffline()) return Promise.resolve(fail('OFFLINE'));
    var req = { scopes: SCOPES.slice(), prompt: 'select_account' };
    // No await before loginPopup: keeps the click's user activation for window.open.
    if (pca && !preferRedirect()) return popupLogin(req);
    // Client not built yet (or iOS): after the await the activation may be gone, so a blocked popup falls back to redirect.
    return init().then(function () {
      if (!pca) return fail('NO_LIB');
      return preferRedirect() ? redirectLogin(req) : popupLogin(req);
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

  /** Disconnects this device (clears the local token cache only; the outbox is kept and waits for a new connection). */
  function signOut() {
    var acc = activeAcc;
    var p = Promise.resolve();
    if (pca && acc) {
      p = Promise.resolve().then(function () { return pca.clearCache({ account: acc }); }).catch(function (e) { warn('signOut', e); });
    }
    return p.then(function () { setActive(null); needLogin = false; emit({ type: 'account', basketId: '' }); });
  }

  // ------------------------------------------------------------------ token + Graph requests
  function getToken(forceRefresh) {
    var acc = activeAcc || (pca ? safeActive(pca) : null);
    if (!pca || !acc) return Promise.reject(Object.assign(new Error('no account'), { errorCode: 'no_account_error' }));
    return Promise.resolve().then(function () {
      return pca.acquireTokenSilent({ scopes: SCOPES.slice(), account: acc, forceRefresh: !!forceRefresh });
    }).then(function (r) {
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

  function validName(n) {
    n = String(n == null ? '' : n);
    return n.length > 0 && n.length <= 200 && !NAME_BAD_RE.test(n) && n.trim() === n && !/^\.+$/.test(n) && !/\.$/.test(n);
  }
  /** Graph path of a file in the "OMS Saha" folder: .../me/drive/root:/OMS%20Saha/<name> (append ':/content' etc.). */
  function itemPath(name) {
    return GRAPH + '/me/drive/root:/' + encodeURIComponent(FOLDER) + '/' + encodeURIComponent(name);
  }

  function timedFetch(url, opts) {
    var ctl = typeof AbortController === 'function' ? new AbortController() : null;
    var t = ctl ? setTimeout(function () { try { ctl.abort(); } catch (e) { } }, REQUEST_TIMEOUT_MS) : null;
    opts.cache = 'no-store'; opts.credentials = 'omit'; opts.mode = 'cors';
    if (ctl) opts.signal = ctl.signal;
    return Promise.resolve().then(function () { return doFetch(url, opts); })
      .then(function (r) { if (t) clearTimeout(t); return r; }, function (e) { if (t) clearTimeout(t); throw e; });
  }

  function retryAfterMs(resp) {
    var h = resp && resp.headers && typeof resp.headers.get === 'function' ? resp.headers.get('Retry-After') : null;
    var s = parseInt(h, 10);
    return (isFinite(s) && s > 0 ? Math.min(s, 3600) : 60) * 1000;
  }

  function readJson(resp) {
    return Promise.resolve().then(function () { return resp.text(); })
      .then(function (t) { try { return JSON.parse(t); } catch (e) { return null; } }, function () { return null; });
  }

  // Graph error -> our code. `forRead`: 404 means "no such file" (download) instead of "no drive" (upload).
  function mapError(resp, forRead) {
    var st = resp.status;
    if (st === 429 || st === 503 || st === 504) return Promise.resolve(fail('THROTTLED', st, 'HTTP ' + st, { retryAfter: retryAfterMs(resp) }));
    return readJson(resp).then(function (j) {
      var gc = (j && j.error && String(j.error.code || '')) || '';
      var detail = 'HTTP ' + st + (gc ? ' ' + gc : '') + (j && j.error && j.error.message ? ': ' + j.error.message : '');
      if (st === 507 || /quotaLimitReached|insufficientStorage/i.test(gc)) return fail('QUOTA', st, detail);
      // Upload by path creates missing folders, so a 404 there means the account has no OneDrive (e.g. work account without a license).
      if (st === 404) return fail(forRead ? 'NOT_FOUND' : 'NO_DRIVE', st, detail);
      if (st === 403) return fail('FORBIDDEN', st, detail);
      if (st >= 400 && st < 500) return fail('BAD_REQUEST', st, detail);
      return fail('SERVER', st, detail);
    });
  }

  function fetchFailure(e) {
    if (errName(e) === 'AbortError') return fail('OFFLINE', 0, 'timeout');
    return fail('OFFLINE', 0, (e && e.message) || '');
  }

  // Runs `doReq(token)` with one forced token refresh on 401.
  function withToken(doReq) {
    function attempt(forceRefresh) {
      return getToken(forceRefresh).then(function (token) {
        return doReq(token).then(function (resp) {
          if (resp && resp.status === 401) {
            if (!forceRefresh) return attempt(true);
            needLogin = true;
            return { failed: fail('LOGIN', 401) };
          }
          return { resp: resp };
        }, function (e) { return { failed: fetchFailure(e) }; });
      }, function (e) { return { failed: tokenFailure(e) }; });
    }
    return attempt(false);
  }

  function preflight() {
    if (!GUID_RE.test(clientId())) return fail('CONFIG');
    if (!lib()) return fail('NO_LIB');
    if (isOffline()) return fail('OFFLINE');
    return null;
  }

  /**
   * Writes `text` as "OMS Saha/<name>" in the connected OneDrive (folder made on first use; a same-named file gets a new name).
   * Resolves {ok, status, code, error[, detail, retryAfter, item:{id,name,size,webUrl}]}; never rejects.
   * code: '' | CONFIG | NO_LIB | LOGIN | OFFLINE | THROTTLED | QUOTA | FORBIDDEN | NO_DRIVE | BAD_REQUEST | BAD_NAME | BAD_DATA | SERVER | UNKNOWN
   */
  function upload(o) { return uploadRaw(o).then(loginEvent); }
  // A public call that needs a new connection raises one 'login' event (flush raises its own, per basket).
  function loginEvent(r) {
    if (r && r.code === 'LOGIN') emit({ type: 'login', basketId: '', code: r.code, error: r.error });
    return r;
  }

  function uploadRaw(o) {
    o = o || {};
    var name = String(o.name == null ? '' : o.name);
    var text = o.text == null ? '' : String(o.text);
    if (!validName(name)) return Promise.resolve(fail('BAD_NAME'));
    if (!text || text.length > MAX_UPLOAD) return Promise.resolve(fail('BAD_DATA'));
    var pre = preflight();
    if (pre) return Promise.resolve(pre);
    var url = itemPath(name) + ':/content?@microsoft.graph.conflictBehavior=rename';
    return init().then(function () {
      return withToken(function (token) {
        return timedFetch(url, { method: 'PUT', headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'text/plain' }, body: text });
      });
    }).then(function (x) {
      if (x.failed) return x.failed;
      var resp = x.resp, st = resp.status;
      if (st === 200 || st === 201) {
        needLogin = false;
        return readJson(resp).then(function (j) {
          return result(true, st, '', '', { item: j ? { id: j.id || '', name: j.name || name, size: j.size || 0, webUrl: j.webUrl || '' } : { id: '', name: name, size: 0, webUrl: '' } });
        });
      }
      return mapError(resp, false);
    }).catch(function (e) {
      warn('upload', e);
      return fail('UNKNOWN', 0, (e && e.message) || '');
    });
  }

  /**
   * Reads "OMS Saha/<name>" from the connected OneDrive.
   * opt.etag: when the file still has this eTag, resolves {ok:true, notModified:true} without downloading.
   * Resolves {ok, text, code, error, status, etag, lastModified (ms), size[, notModified]}; 404 -> code NOT_FOUND. Never rejects.
   */
  function download(name, opt) { return downloadRaw(name, opt).then(loginEvent); }

  function downloadRaw(name, opt) {
    opt = opt || {};
    name = String(name == null ? '' : name);
    function dfail(r) { r.text = ''; return r; }
    if (!validName(name)) return Promise.resolve(dfail(fail('BAD_NAME')));
    var pre = preflight();
    if (pre) return Promise.resolve(dfail(pre));
    var meta = null;
    return init().then(function () {
      return withToken(function (token) {
        return timedFetch(itemPath(name), { method: 'GET', headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' } });
      });
    }).then(function (x) {
      if (x.failed) return dfail(x.failed);
      var resp = x.resp;
      if (resp.status !== 200) return mapError(resp, true).then(dfail);
      return readJson(resp).then(function (j) {
        if (!j || !j.file) return dfail(fail('NOT_FOUND', 404, 'not a file'));
        var lm = Date.parse(j.lastModifiedDateTime || '');
        meta = { etag: String(j.eTag || ''), lastModified: isFinite(lm) ? lm : 0, size: j.size || 0 };
        if (opt.etag && meta.etag && opt.etag === meta.etag) return result(true, 200, '', '', { text: '', notModified: true, etag: meta.etag, lastModified: meta.lastModified, size: meta.size });
        var dl = j['@microsoft.graph.downloadUrl'];
        if (dl) {
          // Pre-authenticated URL: no Authorization header -> simple request, no CORS preflight.
          return timedFetch(dl, { method: 'GET' }).then(gotContent, function (e) { return dfail(fetchFailure(e)); });
        }
        // No download URL in the answer (should not happen): try /content directly.
        return withToken(function (token) {
          return timedFetch(itemPath(name) + ':/content', { method: 'GET', headers: { 'Authorization': 'Bearer ' + token } });
        }).then(function (y) { return y.failed ? dfail(y.failed) : gotContent(y.resp); });
      });
    }).catch(function (e) {
      warn('download', e);
      return dfail(fail('UNKNOWN', 0, (e && e.message) || ''));
    });

    function gotContent(resp) {
      if (resp.status !== 200) return mapError(resp, true).then(dfail);
      return Promise.resolve().then(function () { return resp.text(); }).then(function (t) {
        needLogin = false;
        return result(true, 200, '', '', { text: String(t), etag: meta.etag, lastModified: meta.lastModified, size: meta.size });
      });
    }
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
        if (done) { try { rq.result.close(); } catch (e) { } return; }
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
    try { ls.setItem('__oms_dob', '1'); ls.removeItem('__oms_dob'); return true; } catch (e) { return false; }
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

  function memList() { var out = []; for (var k in store.mem) if (Object.prototype.hasOwnProperty.call(store.mem, k)) out.push(store.mem[k]); return out; }
  function mergeMem(list) {
    var seen = {}, i;
    for (i = 0; i < list.length; i++) seen[list[i].basketId] = 1;
    var mem = memList();
    for (i = 0; i < mem.length; i++) if (!seen[mem[i].basketId]) list.push(mem[i]);
    return list;
  }
  function dbAll() {
    return openStore().then(function (mode) {
      if (mode === 'idb') return idbReq('readonly', function (s) { return s.getAll(); }).then(function (l) { return mergeMem(l || []); }, function () { return memList(); });
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
  function sortItems(list) { return list.sort(function (a, b) { return (a.queuedAt || 0) - (b.queuedAt || 0); }); }

  // ------------------------------------------------------------------ outbox API
  var listeners = [];
  function emit(ev) {
    for (var i = 0; i < listeners.length; i++) { try { listeners[i](ev); } catch (e) { warn('onChange', e); } }
  }
  /**
   * Subscribes to events {type, basketId[, code, error]}: 'queued' | 'sent' | 'failed' | 'login' (connect again) |
   * 'removed' | 'account' (connected / disconnected). Returns an unsubscribe function.
   */
  function onChange(cb) {
    if (typeof cb !== 'function') return function () { };
    listeners.push(cb);
    return function () { var i = listeners.indexOf(cb); if (i >= 0) listeners.splice(i, 1); };
  }

  /** Queues a basket file {basketId, name, text} (same basketId replaces the queued copy). Resolves to the stored item; rejects on bad input. */
  function enqueue(o) {
    o = o || {};
    var id = String(o.basketId || '').trim();
    var name = String(o.name == null ? '' : o.name);
    var text = o.text == null ? '' : String(o.text);
    if (!id) return Promise.reject(new Error('basketId gerekli'));
    if (!validName(name)) return Promise.reject(new Error(MSG.BAD_NAME));
    if (!text || text.length > MAX_UPLOAD) return Promise.reject(new Error(MSG.BAD_DATA));
    return dbAll().then(function (list) {
      var prev = null;
      for (var i = 0; i < list.length; i++) if (list[i].basketId === id) prev = list[i];
      var item = {
        basketId: id, name: name, text: text,
        queuedAt: prev ? prev.queuedAt : now(), tries: prev ? prev.tries : 0, lastTry: prev ? prev.lastTry : 0, lastCode: '', lastError: ''
      };
      return dbPut(item).then(function () { emit({ type: 'queued', basketId: id }); return publicItem(item); });
    });
  }

  function publicItem(it) {
    return {
      basketId: it.basketId, name: it.name, text: it.text, queuedAt: it.queuedAt || 0, tries: it.tries || 0,
      lastTry: it.lastTry || 0, lastCode: it.lastCode || '', lastError: it.lastError || ''
    };
  }

  /** Queued items, oldest first: [{basketId, name, text, queuedAt, tries, lastTry, lastCode, lastError}]. */
  function pending() { return dbAll().then(function (list) { return sortItems(list).map(publicItem); }); }

  /** Drops a queued basket (e.g. the user sent it another way). */
  function remove(basketId) {
    var id = String(basketId || '');
    return dbDel(id).then(function () { emit({ type: 'removed', basketId: id }); });
  }

  /**
   * One tap send: upload now; when that fails for a reason that can pass (no internet, connect again, busy, full...)
   * the basket goes to the outbox. Resolves the upload result plus {queued: bool}. Emits 'sent' / 'queued'.
   */
  function send(o) {
    o = o || {};
    var id = String(o.basketId || '').trim();
    return upload(o).then(function (r) {
      r.queued = false;
      if (r.ok) {
        if (!id) return r;
        return dbDel(id).then(function () { emit({ type: 'sent', basketId: id }); return r; });
      }
      if (r.code === 'THROTTLED') holdUntil = now() + (r.retryAfter || 60000);
      if (!id || !RETRY_CODES[r.code]) return r;
      return enqueue({ basketId: id, name: o.name, text: o.text }).then(function () { r.queued = true; return r; }, function () { return r; });
    });
  }

  var flushing = null, holdUntil = 0;

  /**
   * Uploads every queued basket (single flight: a call during a running flush gets the same promise).
   * Stops at the first channel-wide problem (offline, connect again, busy, full). Resolves {sent:[basketId], failed:[{basketId, code, error}]}.
   */
  function flush() {
    if (flushing) return flushing;
    flushing = dbAll().then(function (list) {
      var items = sortItems(list), out = { sent: [], failed: [] }, i = 0;
      if (!items.length) return out;
      var gate = !GUID_RE.test(clientId()) ? 'CONFIG' : !lib() ? 'NO_LIB' : isOffline() ? 'OFFLINE' : now() < holdUntil ? 'THROTTLED' : '';
      if (gate) return skipAll(items, out, fail(gate), false);
      return init().then(function () {
        if (!account()) { needLogin = true; return skipAll(items, out, fail('LOGIN'), true); }
        function next() {
          if (i >= items.length) return out;
          var it = items[i++];
          return uploadRaw({ name: it.name, text: it.text }).then(function (r) {
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

  // Items not tried because the channel is down: reported, not counted as tries. `emitLogin` raises one 'login' event.
  function skipAll(items, out, r, emitLogin) {
    for (var k = 0; k < items.length; k++) out.failed.push({ basketId: items[k].basketId, code: r.code, error: r.error });
    if (emitLogin && r.code === 'LOGIN' && items.length) emit({ type: 'login', basketId: items[0].basketId, code: r.code, error: r.error });
    return out;
  }

  // ------------------------------------------------------------------ automatic retry
  var started = false, timer = null, handlers = null;

  /** Starts automatic outbox delivery: now, on `online`, when the tab becomes visible, and every 60 s. Idempotent. */
  function start() {
    if (started || inCallback() || !available()) return;
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

  /** Stops automatic delivery (tests). */
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

  // ------------------------------------------------------------------ helpers for the page
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  /** Spec §17 basket file name: oms_sepet_yyyyMMdd_HHmmss_<dev>.omsb.txt (local time). */
  function basketFileName(dev, date) {
    var d = date instanceof Date ? date : new Date(date != null ? date : now());
    var tag = String(dev || '').replace(/[^A-Za-z0-9_-]/g, '');
    return 'oms_sepet_' + d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate()) + '_' +
      pad2(d.getHours()) + pad2(d.getMinutes()) + pad2(d.getSeconds()) + (tag ? '_' + tag : '') + '.omsb.txt';
  }
  /** Spec §18 catalog file name for this device: oms_katalog_<dev>.omsk */
  function catalogFileName(dev) { return 'oms_katalog_' + String(dev || '').replace(/[^A-Za-z0-9_-]/g, '') + '.omsk'; }

  // ------------------------------------------------------------------ test hooks
  /** Node tests: inject {msal, bridge, fetch, indexedDB, storage, location, navigator, now, clientId}; resets all state. */
  function _setDeps(d) {
    stop();
    d = d || {};
    for (var k in deps) if (Object.prototype.hasOwnProperty.call(deps, k)) deps[k] = Object.prototype.hasOwnProperty.call(d, k) ? d[k] : (k === 'indexedDB' || k === 'storage' ? undefined : null);
    pca = null; initPromise = null; activeAcc = null; needLogin = false;
    authCallback = false; bridgePromise = null; mailOwnsBridge = false;
    store = { mode: '', db: null, mem: {}, opening: null };
    flushing = null; holdUntil = 0; listeners = [];
  }

  var api = {
    SCOPES: SCOPES.slice(),
    AUTHORITY: AUTHORITY,
    FOLDER: FOLDER,
    MESSAGES: MSG,
    available: available,
    init: init,
    account: account,
    loginNeeded: loginNeeded,
    signIn: signIn,
    signOut: signOut,
    upload: upload,
    download: download,
    send: send,
    enqueue: enqueue,
    pending: pending,
    flush: flush,
    start: start,
    stop: stop,
    onChange: onChange,
    remove: remove,
    basketFileName: basketFileName,
    catalogFileName: catalogFileName,
    hasAuthResponse: hasAuthResponse,
    handleAuthCallback: function () { return hasAuthResponse(loc()) ? runBridge() : null; },
    _setDeps: _setDeps
  };
  Object.defineProperty(api, 'CLIENT_ID', { enumerable: true, get: function () { return clientId(); } });
  Object.defineProperty(api, 'isAuthCallback', { enumerable: true, get: function () { return inCallback(); } });
  Object.defineProperty(api, 'storageMode', { enumerable: true, get: function () { return store.mode; } });
  return api;
});
