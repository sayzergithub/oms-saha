/*
 * OMS field mode - direct send over the local network (spec V10 §20).
 *   ping : GET  {root}/saha/api/ping   -> 200 {"ok":true,"pc":"WIN-..","lan":["https://192.168.1.20:47803",…],"v":1}
 *   send : POST {root}/saha/api/basket  body = basket armor (text/plain, spec §2)
 *          -> 200 {"ok":true,"outcome":"Applied|Duplicate|Rejected","basket","applied","skipped","issues","message","pc","lan"}
 *          -> 503 same JSON with outcome "Retry";  400 {"ok":false,"code":"BAD_ARMOR"}
 * Browsers cannot listen to UDP / mDNS, so the PC is found by pinging the roots from the pairing code (+ the page's own
 * origin when the PC serves it under /saha/, + https://oms-inv.local:<port>) in parallel; the first good answer wins.
 * No dependencies. UMD: Node module / browser global OmsLan. Test hooks: opts.fetch, opts.location, opts.now.
 */
(function (root, factory) {
  'use strict';
  if (typeof module === 'object' && module.exports) module.exports = factory(root);
  else root.OmsLan = factory(root);
})(typeof globalThis !== 'undefined' ? globalThis : (typeof self !== 'undefined' ? self : this), function (root) {
  'use strict';

  var PING_PATH = '/saha/api/ping', BASKET_PATH = '/saha/api/basket';
  var DEFAULT_PORT = 47803, MDNS_HOST = 'oms-inv.local';
  var FIND_TIMEOUT = 1500, SEND_TIMEOUT = 15000, CACHE_MS = 60000, MAX_LAN = 8;

  var cache = null;          // {root, pc, lan, at}
  var okRoots = {};          // roots that answered a ping in this page session (certificate already accepted)
  var listeners = [];

  // ------------------------------------------------------------------ helpers
  function str(v) { return v == null ? '' : String(v); }
  /** Same rules as OmsCatalog.normRoot: scheme://host[:port], scheme defaults to https, path dropped. */
  function normRoot(u) {
    var s = str(u).trim();
    if (!s) return '';
    if (!/^[a-z][a-z0-9+.-]*:\/\//i.test(s)) s = 'https://' + s;
    var m = /^(https?:\/\/[^\/?#\s]+)/i.exec(s);
    return m ? m[1].replace(/^HTTPS?:/i, function (x) { return x.toLowerCase(); }) : '';
  }
  function portOf(r) {
    var m = /^https?:\/\/(\[[^\]]+\]|[^:\/]+)(?::(\d+))?$/i.exec(r || '');
    if (!m) return 0;
    if (m[2]) return +m[2];
    return /^http:/i.test(r) ? 80 : 443;
  }
  function isMdns(r) { return /^https?:\/\/[^\/:]+\.local(:\d+)?$/i.test(r || ''); }
  function nowOf(o) { return o && typeof o.now === 'function' ? o.now() : Date.now(); }
  function fetchOf(o) {
    if (o && typeof o.fetch === 'function') return o.fetch;
    if (typeof root.fetch === 'function') return function (u, opt) { return root.fetch(u, opt); };
    return null;
  }
  function locOf(o) {
    if (o && o.location !== undefined) return o.location;
    try { return root.location || null; } catch (e) { return null; }
  }
  /** Normalized, de-duplicated root list (≤ 8); accepts an array or a whitespace / comma separated string. */
  function normList(lan) {
    var arr = Array.isArray(lan) ? lan : (typeof lan === 'string' ? lan.split(/[\s,;]+/) : []);
    var out = [], seen = {};
    for (var i = 0; i < arr.length && out.length < MAX_LAN; i++) {
      var r = normRoot(arr[i]);
      if (r && !seen[r.toLowerCase()]) { seen[r.toLowerCase()] = 1; out.push(r); }
    }
    return out;
  }

  /**
   * Roots to try, in order: the page's own origin (when served by the PC under /saha/), the pairing's lan roots,
   * https://oms-inv.local:<port of the first lan root, else 47803>.
   */
  function candidates(pairing, opts) {
    var list = [], seen = {};
    function add(r) { r = normRoot(r); if (r && !seen[r.toLowerCase()]) { seen[r.toLowerCase()] = 1; list.push(r); } }
    var loc = locOf(opts);
    try { if (loc && /^https?:$/.test(loc.protocol) && /^\/saha\//.test(loc.pathname || '')) add(loc.origin); } catch (e) { /* no location */ }
    var lan = normList(pairing && pairing.lan);
    lan.forEach(add);
    var port = lan.length ? portOf(lan[0]) : 0;
    add('https://' + MDNS_HOST + ':' + (port || DEFAULT_PORT));
    return list;
  }

  /** fetch with a timeout. Resolves the Response or rejects with {code:'NETWORK'|'TIMEOUT', cause}. */
  function fetchT(fetchFn, url, init, timeoutMs, outerSignal) {
    var ctl = typeof root.AbortController === 'function' ? new root.AbortController() : null;
    var timer = null, timedOut = false, aborted = false;
    if (ctl) init.signal = ctl.signal;
    var onOuter = function () { aborted = true; if (ctl) { try { ctl.abort(); } catch (e) { /* ignore */ } } };
    if (outerSignal) outerSignal.push(onOuter);
    return new Promise(function (resolve, reject) {
      timer = setTimeout(function () {
        timedOut = true;
        if (ctl) { try { ctl.abort(); } catch (e) { /* ignore */ } }
        reject({ code: 'TIMEOUT' });
      }, timeoutMs);
      Promise.resolve().then(function () { return fetchFn(url, init); }).then(function (res) {
        clearTimeout(timer); resolve(res);
      }, function (e) {
        clearTimeout(timer);
        if (timedOut) return;
        // TypeError = DNS / refused / TLS (self-signed certificate not accepted yet) / CORS / private-network block.
        reject({ code: aborted ? 'ABORTED' : (e && e.name === 'AbortError' ? 'TIMEOUT' : 'NETWORK'), cause: e });
      });
    });
  }
  function readJson(res) {
    return Promise.resolve().then(function () { return res.text(); }).then(function (t) {
      try { return JSON.parse(t); } catch (e) { return null; }
    }, function () { return null; });
  }

  function emitLan(info) {
    if (!info || !info.lan || !info.lan.length) return;
    listeners.slice().forEach(function (cb) { try { cb(info); } catch (e) { /* listener errors stay local */ } });
  }
  /** cb({root, pc, lan}) is called whenever a ping or send answer carries the PC's current address list. Returns an unsubscribe. */
  function onLan(cb) {
    if (typeof cb !== 'function') return function () { };
    listeners.push(cb);
    return function () { var i = listeners.indexOf(cb); if (i >= 0) listeners.splice(i, 1); };
  }

  /** The cached winner if it is younger than 60 s, else null. */
  function active(opts) {
    if (!cache) return null;
    if (nowOf(opts) - cache.at > CACHE_MS) return null;
    return { ok: true, root: cache.root, pc: cache.pc, lan: cache.lan.slice(), cached: true };
  }
  function forget() { cache = null; }

  /**
   * opts: {pairing, timeoutMs=1500, force (skip the 60 s cache), roots (override the candidate list)}.
   * Resolves (never rejects) {ok, root, pc, lan, needCertUrl, code, tried}.
   *   code '' (found) | 'NONE' (no roots / no fetch) | 'NOT_FOUND' (nobody answered; needCertUrl when every IP root failed with a
   *   network/TLS error rather than a timeout and none of them ever answered before = probably the self-signed certificate).
   */
  function find(opts) {
    var o = opts || {};
    var roots = Array.isArray(o.roots) ? normList(o.roots) : candidates(o.pairing, o);
    var fetchFn = fetchOf(o);
    var timeoutMs = o.timeoutMs > 0 ? o.timeoutMs : FIND_TIMEOUT;
    if (!o.force && cache && nowOf(o) - cache.at <= CACHE_MS) {
      for (var i = 0; i < roots.length; i++) if (roots[i].toLowerCase() === cache.root.toLowerCase()) return Promise.resolve(active(o));
    }
    if (!roots.length || !fetchFn) return Promise.resolve({ ok: false, root: '', pc: '', lan: [], needCertUrl: '', code: 'NONE', tried: [] });
    var aborts = [], done = false;
    return new Promise(function (resolve) {
      var left = roots.length, fails = {};
      function finishNotFound() {
        var ip = roots.filter(function (r) { return !isMdns(r); });
        var allNet = ip.length > 0 && ip.every(function (r) { return fails[r] === 'NETWORK'; });
        var fresh = ip.filter(function (r) { return !okRoots[r.toLowerCase()]; });
        resolve({ ok: false, root: '', pc: '', lan: [], code: 'NOT_FOUND', tried: roots.slice(),
          needCertUrl: allNet && fresh.length ? fresh[0] + '/saha/' : '' });
      }
      roots.forEach(function (r) {
        var init = { method: 'GET', mode: 'cors', cache: 'no-store', credentials: 'omit', referrerPolicy: 'no-referrer' };
        fetchT(fetchFn, r + PING_PATH, init, timeoutMs, aborts).then(function (res) {
          if (!res || res.status !== 200) throw { code: 'HTTP' };
          return readJson(res).then(function (j) {
            if (!j || j.ok !== true) throw { code: 'HTTP' };
            return j;
          });
        }).then(function (j) {
          if (done) return;
          done = true;
          aborts.forEach(function (f) { f(); });
          okRoots[r.toLowerCase()] = 1;
          var lan = normList(j.lan);
          cache = { root: r, pc: str(j.pc), lan: lan, at: nowOf(o) };
          emitLan({ root: r, pc: cache.pc, lan: lan });
          resolve({ ok: true, root: r, pc: cache.pc, lan: lan.slice(), needCertUrl: '', code: '', tried: roots.slice() });
        }, function (e) {
          fails[r] = (e && e.code) || 'NETWORK';
          if (--left === 0 && !done) { done = true; finishNotFound(); }
        });
      });
    });
  }

  function num(v) { var n = Number(v); return isFinite(n) && n >= 0 ? Math.floor(n) : 0; }

  /**
   * opts: {root, armor, timeoutMs=15000}. Resolves (never rejects)
   *   {ok, outcome, applied, skipped, issues, message, basket, pc, lan, code, error, needCertUrl, status}
   * ok = the PC took the basket (outcome Applied or Duplicate). code: '' | 'REJECTED' | 'RETRY' | 'BAD_ARMOR' | 'SERVER' |
   * 'OFFLINE' (no answer / timeout) | 'NEED_CERT' (network/TLS error on a root that never answered a ping) | 'NO_ROOT'.
   * error is short Turkish text for the user.
   */
  function send(opts) {
    var o = opts || {};
    var r = normRoot(o.root), fetchFn = fetchOf(o), timeoutMs = o.timeoutMs > 0 ? o.timeoutMs : SEND_TIMEOUT;
    var out = { ok: false, outcome: '', applied: 0, skipped: 0, issues: 0, message: '', basket: '', pc: '', lan: [], code: '', error: '', needCertUrl: '', status: 0 };
    if (!r || !fetchFn) { out.code = 'NO_ROOT'; out.error = 'Yerel ağda PC bulunamadı.'; return Promise.resolve(out); }
    var init = { method: 'POST', mode: 'cors', cache: 'no-store', credentials: 'omit', referrerPolicy: 'no-referrer',
      headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: str(o.armor) };
    return fetchT(fetchFn, r + BASKET_PATH, init, timeoutMs, null).then(function (res) {
      out.status = res.status;
      return readJson(res).then(function (j) {
        j = j || {};
        out.outcome = str(j.outcome); out.basket = str(j.basket); out.message = str(j.message); out.pc = str(j.pc);
        out.applied = num(j.applied); out.skipped = num(j.skipped); out.issues = num(j.issues);
        out.lan = normList(j.lan);
        if (out.lan.length || out.pc) {
          okRoots[r.toLowerCase()] = 1;
          if (cache && cache.root.toLowerCase() === r.toLowerCase()) { if (out.lan.length) cache.lan = out.lan.slice(); cache.at = nowOf(o); }
          emitLan({ root: r, pc: out.pc, lan: out.lan.slice() });
        }
        if (res.status === 400 || j.code === 'BAD_ARMOR') { out.code = 'BAD_ARMOR'; out.error = 'PC sepeti okuyamadı (bozuk veri).'; return out; }
        if (res.status === 503 || out.outcome === 'Retry') { out.outcome = 'Retry'; out.code = 'RETRY'; out.error = out.message || 'PC şu an sepeti işleyemedi; biraz sonra yeniden deneyin.'; return out; }
        if (res.status !== 200) { out.code = 'SERVER'; out.error = 'PC beklenmeyen bir yanıt verdi (' + res.status + ').'; return out; }
        if (out.outcome === 'Applied' || out.outcome === 'Duplicate') { out.ok = true; return out; }
        if (out.outcome === 'Rejected') { out.code = 'REJECTED'; out.error = out.message || 'PC sepeti kabul etmedi.'; return out; }
        out.code = 'SERVER'; out.error = 'PC beklenmeyen bir yanıt verdi.'; return out;
      });
    }, function (e) {
      if (cache && cache.root.toLowerCase() === r.toLowerCase()) cache = null;   // re-search next time
      if (e && e.code === 'NETWORK' && !okRoots[r.toLowerCase()] && !isMdns(r)) {
        out.code = 'NEED_CERT'; out.needCertUrl = r + '/saha/';
        out.error = 'PC\'ye bağlanılamadı. Tarayıcının PC\'ye bağlanmayı bir kez onaylaması gerekebilir.';
        return out;
      }
      out.code = 'OFFLINE';
      out.error = e && e.code === 'TIMEOUT' ? 'PC zamanında yanıt vermedi. Sepet PC\'ye ulaşmış olabilir; yeniden gönderirseniz PC aynı sepeti iki kez işlemez.'
        : 'PC\'ye ulaşılamadı. Telefon PC ile aynı Wi-Fi ağında mı?';
      return out;
    });
  }

  /** New pairing.lan after the PC reported its addresses: the PC's list first, then old non-IP names (hand-typed), ≤ 8. */
  function mergeLan(oldList, newList) {
    var n = normList(newList);
    if (!n.length) return normList(oldList);
    normList(oldList).forEach(function (r) {
      if (n.length >= MAX_LAN) return;
      if (/^https?:\/\/(\d{1,3}\.){3}\d{1,3}(:\d+)?$/i.test(r) || /^https?:\/\/\[/.test(r)) return;   // stale DHCP address: drop
      if (n.map(function (x) { return x.toLowerCase(); }).indexOf(r.toLowerCase()) < 0) n.push(r);
    });
    return n;
  }

  /** Test hook: clear cache and session state. */
  function _reset() { cache = null; okRoots = {}; listeners = []; }

  return {
    candidates: candidates, find: find, send: send, onLan: onLan, active: active, forget: forget,
    mergeLan: mergeLan, normRoot: normRoot, _reset: _reset,
    DEFAULT_PORT: DEFAULT_PORT, CACHE_MS: CACHE_MS
  };
});
