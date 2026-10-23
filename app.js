/* OMS Saha: field page logic (offline, serverless).
   Flow: pair once (OMSP1 code) -> scan reels into a basket (IndexedDB, localStorage backup) -> seal (envelope.js) ->
   mail it (mailto / Web Share / copy / download) -> keep in history for 30 days.
   V6: camera settings tab (OmsCam, camsettings.js), stock catalog from the PC on the LAN (OmsCatalog, catalog.js),
   direct send through Microsoft Graph with an outbox (OmsMail, sendmail.js + msal-browser.min.js; both load before app.js).
   V7: "Buluta gönder" is the primary way (OmsCloud, cloudsend.js): share sheet -> the cloud app's "OMS Saha" folder (phones),
   File System Access write into that folder (desktop Chrome/Edge), download as the last resort. The PC watches the cloud's
   sync folder. E-mail / direct send / copy / download stay under "Diğer yollar"; their settings under Ayarlar > Gelişmiş.
   V8: direct OneDrive (OmsDrive, onedrive.js): once connected, Gönder is one tap (no dialog), with an outbox; the share dialog stays
   behind the "⋯" button. Stok tab (OmsStock, stockview.js) over the catalog, which now also comes from OneDrive (oms_katalog_<dev>.omsk ->
   OmsCatalog.loadArmor). Scan mode switch Seri | Manuel: Seri = OmsSerial (serialscan.js) decides per frame, every new reel is added
   without a form, big counter + last 3 + Bitti; Manuel = the line editor opens after every add, the camera waits until it closes.
   All modules are optional: when one is missing the page keeps its earlier behaviour.
   Contract: reviews/V5_SAHA_SPEC.md sections 1, 3, 8, 10, 12-19. UI text is Turkish; code comments are English.
   Language level: ES2017 (async/await), no optional chaining / nullish operators (older Safari / Samsung Internet). */
(function () {
'use strict';

var SUBJ_DEFAULT = 'oms_inv_sepet';
var MAX_LINES = 2000, MAX_RAW = 8192, HIST_DAYS = 30, MAILTO_MAX = 60000, DUP_MS = 3000, HOLD_1D_MS = 700;
var K = { pair: 'oms.saha.pair', user: 'oms.saha.user', loc: 'oms.saha.loc', snd: 'oms.saha.snd', vib: 'oms.saha.vib', cam: 'oms.saha.cam',
          basket: 'oms.saha.basket', hist: 'oms.saha.hist', catTry: 'oms.saha.cattry',
          cloud: 'oms.saha.cloud', guided: 'oms.saha.guided', adv: 'oms.saha.adv', welcomed: 'oms.saha.welcomed',
          smode: 'oms.saha.smode', drvTry: 'oms.saha.drvtry', drvSend: 'oms.saha.drvsend', catEtag: 'oms.saha.catetag',
          lanOk: 'oms.saha.lanok', lanCertHide: 'oms.saha.lancerthide' };
var CAT_STALE_MS = 12 * 3600000, CAT_RETRY_MS = 30 * 60000, DRV_CAT_STALE_MS = 30 * 60000, DRV_CAT_RETRY_MS = 10 * 60000, SER_GAP_MS = 120;
var IS_ANDROID = /Android/i.test(navigator.userAgent || '');
var IS_IOS = /iP(hone|ad|od)/.test(navigator.userAgent || '') || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
var DEMO = (function () { var m = /[?&](selftest=1|demo=([a-z]+))/.exec(location.search); return m ? (m[2] || 'selftest') : ''; })();

/* ===================== small helpers ===================== */
function $(id) { return document.getElementById(id); }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function ico(id, cls) { return '<svg class="ic' + (cls ? ' ' + cls : '') + '" aria-hidden="true"><use href="#' + id + '"/></svg>'; }
function pad(n) { return (n < 10 ? '0' : '') + n; }
function iso(ms) { return new Date(ms == null ? Date.now() : ms).toISOString().replace(/\.\d{3}Z$/, 'Z'); }
function fmtTime(ms) { var d = new Date(ms); return pad(d.getDate()) + '.' + pad(d.getMonth() + 1) + '.' + d.getFullYear() + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes()); }
function stamp(ms) { var d = new Date(ms == null ? Date.now() : ms); return '' + d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + '_' + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds()); }
function fmtQty(n) { if (n == null || n === '') return ''; var x = Number(n); if (!isFinite(x)) return String(n); try { return x.toLocaleString('tr-TR'); } catch (e) { return String(x); } }
function tick(ms) { return new Promise(function (r) { setTimeout(r, ms || 0); }); }
function isEmail(s) { return /^[^\s@<>()",;:\\]+@[^\s@<>()",;:\\]+\.[^\s@<>()",;:\\]+$/.test(String(s || '')); }
/* control characters made visible: RS/GS/EOT are part of ISO 15434 data, show them as symbols */
function showRaw(s) {
  return esc(s).replace(/[\x00-\x1F\x7F]/g, function (c) {
    var m = { '\x1e': 'RS', '\x1d': 'GS', '\x04': 'EOT', '\x1c': 'FS', '\x1f': 'US', '\n': 'LF', '\r': 'CR', '\t': 'TAB' }[c];
    return '<span class="cc">\u2039' + (m || ('x' + c.charCodeAt(0).toString(16))) + '\u203A</span>';
  });
}
function shortRaw(s) { s = String(s || '').replace(/[\x00-\x1F\x7F]/g, '\u00B7'); return s.length > 40 ? s.slice(0, 40) + '\u2026' : s; }
function b64uEnc(bytes) { var s = ''; for (var i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]); return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, ''); }
function b64uDec(s) { s = String(s || '').replace(/-/g, '+').replace(/_/g, '/').replace(/[^A-Za-z0-9+/]/g, ''); while (s.length % 4) s += '='; var b = atob(s), a = new Uint8Array(b.length); for (var i = 0; i < b.length; i++) a[i] = b.charCodeAt(i); return a; }
function uuid() {
  try { if (window.OmsEnvelope && OmsEnvelope.newUuid) return OmsEnvelope.newUuid(); } catch (e) { }
  try { if (window.crypto && crypto.randomUUID) return crypto.randomUUID(); } catch (e) { }
  var b = new Uint8Array(16);
  if (window.crypto && crypto.getRandomValues) crypto.getRandomValues(b); else for (var i = 0; i < 16; i++) b[i] = Math.random() * 256 | 0;
  b[6] = (b[6] & 15) | 64; b[8] = (b[8] & 63) | 128;
  var h = ''; for (var j = 0; j < 16; j++) h += (b[j] + 256).toString(16).slice(1);
  return h.slice(0, 8) + '-' + h.slice(8, 12) + '-' + h.slice(12, 16) + '-' + h.slice(16, 20) + '-' + h.slice(20);
}
function busy(btn, on) { if (!btn) return; btn.disabled = !!on; btn.setAttribute('aria-busy', on ? 'true' : 'false'); }
function setMsg(id, txt, kind) { var e = $(id); if (!e) return; e.className = 'msg' + (kind ? ' ' + kind : ''); e.setAttribute('role', kind === 'err' ? 'alert' : 'status'); e.textContent = txt || ''; }
function mql(q) { return !!(window.matchMedia && matchMedia(q).matches); }

/* localStorage with try/catch (private mode, blocked site data) and an in-memory fallback */
var LS = {
  mem: {},
  get: function (k) { try { var v = localStorage.getItem(k); if (v != null) return v; } catch (e) { } return Object.prototype.hasOwnProperty.call(this.mem, k) ? this.mem[k] : null; },
  set: function (k, v) { if (DEMO) { this.mem[k] = String(v); return true; } try { localStorage.setItem(k, v); delete this.mem[k]; return true; } catch (e) { this.mem[k] = String(v); return false; } },
  del: function (k) { delete this.mem[k]; if (DEMO) return; try { localStorage.removeItem(k); } catch (e) { } },
  ok: function () { try { var t = '__oms_t'; localStorage.setItem(t, '1'); localStorage.removeItem(t); return true; } catch (e) { return false; } }
};
function jget(k, def) { var v = LS.get(k); if (!v) return def; try { return JSON.parse(v); } catch (e) { return def; } }

/* ===================== storage: IndexedDB (primary) + localStorage mirror ===================== */
var DB = {
  db: null, mode: 'mem',
  open: function () {
    var self = this;
    return new Promise(function (res) {
      var done = false; function fin(m) { if (done) return; done = true; self.mode = m; res(); }
      if (DEMO) { fin('mem'); return; }
      var lsOk = LS.ok();
      if (!window.indexedDB) { fin(lsOk ? 'ls' : 'mem'); return; }
      var rq;
      try { rq = indexedDB.open('oms-saha', 1); } catch (e) { fin(lsOk ? 'ls' : 'mem'); return; }
      rq.onupgradeneeded = function () { try { rq.result.createObjectStore('kv'); } catch (e) { } };
      rq.onsuccess = function () { self.db = rq.result; self.db.onversionchange = function () { try { self.db.close(); } catch (e) { } self.db = null; }; fin('idb'); };
      rq.onerror = function () { fin(lsOk ? 'ls' : 'mem'); };
      rq.onblocked = function () { fin(lsOk ? 'ls' : 'mem'); };
      setTimeout(function () { fin(lsOk ? 'ls' : 'mem'); }, 3000);   /* some private modes never answer */
    });
  },
  get: function (k) {
    var self = this;
    return new Promise(function (res) {
      if (!self.db) { res(jget(k, null)); return; }
      try {
        var rq = self.db.transaction('kv', 'readonly').objectStore('kv').get(k);
        rq.onsuccess = function () { res(rq.result != null ? rq.result : jget(k, null)); };
        rq.onerror = function () { res(jget(k, null)); };
      } catch (e) { res(jget(k, null)); }
    });
  },
  set: function (k, v) {
    if (DEMO) return Promise.resolve();
    var self = this, txt = null;
    try { txt = JSON.stringify(v); } catch (e) { }
    if (txt != null) { if (txt.length < 2500000) LS.set(k, txt); else LS.del(k); }   /* mirror; skip when it would blow the 5 MB quota */
    return new Promise(function (res) {
      if (!self.db) { res(); return; }
      try {
        var tx = self.db.transaction('kv', 'readwrite'); tx.objectStore('kv').put(v, k);
        tx.oncomplete = function () { res(); }; tx.onerror = tx.onabort = function () { res(); };
      } catch (e) { res(); }
    });
  }
};

/* ===================== state ===================== */
var S = {
  pair: null,          /* {v,dev,sec(b64url),name,to,pc,subj,at,lan:[origin],cid} */
  user: '', defLoc: '',
  cloud: '',           /* 'onedrive' | 'gdrive' | 'icloud' | '' (unknown): only picks the guide text */
  fb: { snd: true, vib: true },
  basket: null,        /* {basket,created,lines:[{id,raw,fmt,qty,loc,note,ts}]} */
  smode: 'manual',     /* scan mode: 'serial' | 'manual' (spec section 19; remembered, manual by default) */
  hist: []             /* [{basket,created,lines,devName,user,sends:[ms],via?:'cloud'|'direct'|'drive',out?:'queued'|'failed'|'login',ch?:'mail'|'drive',q?:ms,outErr?}] newest first */
};
var PV = new Map();    /* raw -> OmsLabel.preview() cache */
function pv(raw, fmt) {
  var key = (fmt || '') + '\u0000' + raw;
  if (PV.has(key)) return PV.get(key);
  var p = null;
  try { if (window.OmsLabel && OmsLabel.preview) p = OmsLabel.preview(raw, fmt || undefined) || null; } catch (e) { p = null; }
  if (PV.size > 5000) PV.clear();
  PV.set(key, p); return p;
}
function lines() { return S.basket ? S.basket.lines : []; }
function ensureBasket() { if (!S.basket) S.basket = { basket: uuid(), created: iso(), lines: [] }; return S.basket; }
var saveT = null;
function saveBasket(now) {
  clearTimeout(saveT);
  if (now) { saveT = null; return DB.set(K.basket, S.basket); }
  saveT = setTimeout(function () { saveT = null; DB.set(K.basket, S.basket); }, 120);
  return Promise.resolve();
}
function saveHist() { return DB.set(K.hist, S.hist); }
function flushSave() { if (saveT) { clearTimeout(saveT); saveT = null; DB.set(K.basket, S.basket); } }
function lastAt(h) { return h.sends && h.sends.length ? h.sends[h.sends.length - 1] : (h.q || 0); }
function pruneHist() {
  /* entries still waiting in the outbox are kept regardless of age */
  var lim = Date.now() - HIST_DAYS * 86400000, n = S.hist.length;
  S.hist = S.hist.filter(function (h) { return h.out || lastAt(h) >= lim; });
  return S.hist.length !== n;
}
function savePair() { if (S.pair) LS.set(K.pair, JSON.stringify(S.pair)); else LS.del(K.pair); }
function subj() { return (S.pair && S.pair.subj) || SUBJ_DEFAULT; }
function secBytes() { return b64uDec(S.pair.sec); }

/* ===================== snackbar ===================== */
var SN = { t: null, act: null };
function toast(txt, err, action) {
  var s = $('snack'), a = s.querySelector('.act');
  s.querySelector('.tx').textContent = txt;
  s.classList.toggle('err', !!err);
  s.setAttribute('role', err ? 'alert' : 'status');
  s.querySelector('.st use').setAttribute('href', err ? '#i-err' : '#i-ok');
  SN.act = action || null;
  if (action) { a.textContent = action.t; a.hidden = false; } else { a.hidden = true; a.textContent = ''; }
  s.classList.add('show');
  clearTimeout(SN.t); SN.t = setTimeout(hideToast, err ? 7000 : (action ? 6000 : 3500));
}
function hideToast() { $('snack').classList.remove('show'); SN.act = null; }
$('snack').querySelector('.x').addEventListener('click', hideToast);
$('snack').querySelector('.act').addEventListener('click', function () { var a = SN.act; hideToast(); if (a && a.fn) a.fn(); });

/* ===================== layers (modal / camera) + Android back button ===================== */
var LAYER = {
  stack: [],
  push: function (name, close) { this.stack.push({ n: name, close: close }); try { history.pushState({ omsLayer: name }, ''); } catch (e) { } },
  /* closed from the UI: drop our history entry too */
  pop: function (name) {
    for (var i = this.stack.length - 1; i >= 0; i--) if (this.stack[i].n === name) {
      this.stack.splice(i, 1);
      if (history.state && history.state.omsLayer === name) { this.skip = true; try { history.back(); } catch (e) { this.skip = false; } }
      return;
    }
  }
};
window.addEventListener('popstate', function () {
  if (LAYER.skip) { LAYER.skip = false; return; }
  var top = LAYER.stack.pop(); if (top && top.close) top.close(true);
});

/* ===================== modal ===================== */
var MOD = { ret: null, onclose: null };
function openModal(title, html, onclose) {
  var m = $('modal'), wasOpen = !m.hidden;
  if (wasOpen && MOD.onclose) { var f = MOD.onclose; MOD.onclose = null; try { f(); } catch (e) { } }
  $('mt').textContent = title; $('mb').onclick = null; $('mb').innerHTML = html;
  MOD.onclose = onclose || null;
  if (!wasOpen) { MOD.ret = document.activeElement; m.hidden = false; LAYER.push('modal', function () { closeModal(true); }); }
  m.scrollTop = 0;
  setTimeout(function () { var f = m.querySelector('[autofocus]') || m.querySelector('.box'); try { f.focus({ preventScroll: true }); } catch (e) { f.focus(); } }, 30);
}
function closeModal(fromBack) {
  var m = $('modal'); if (m.hidden) return;
  m.hidden = true; $('mb').innerHTML = '';
  if (!fromBack) LAYER.pop('modal'); else LAYER.stack = LAYER.stack.filter(function (x) { return x.n !== 'modal'; });
  var f = MOD.onclose; MOD.onclose = null; if (f) try { f(); } catch (e) { }
  var r = MOD.ret; MOD.ret = null; if (r && r.focus && document.contains(r)) try { r.focus({ preventScroll: true }); } catch (e) { }
}
$('modal').addEventListener('click', function (e) { if (e.target === this) closeModal(); });
/* focus trap */
$('modal').addEventListener('keydown', function (e) {
  if (e.key !== 'Tab') return;
  var f = [].filter.call(this.querySelectorAll('a[href],button:not([disabled]),input:not([disabled]),textarea:not([disabled]),select:not([disabled]),summary'), function (x) { return x.offsetParent !== null; });
  if (!f.length) return;
  if (e.shiftKey && document.activeElement === f[0]) { e.preventDefault(); f[f.length - 1].focus(); }
  else if (!e.shiftKey && document.activeElement === f[f.length - 1]) { e.preventDefault(); f[0].focus(); }
});
/* in-page confirm (no window.confirm: captive/in-app browsers may suppress it) */
function ask(title, text, okText, danger) {
  return new Promise(function (res) {
    var done = false;
    openModal(title, '<p>' + text + '</p><div class="acts r"><button type="button" class="btn btn-sec" data-a="0">Vazgeç</button>' +
      '<button type="button" class="btn ' + (danger ? 'btn-danger' : 'btn-pri') + '" data-a="1" autofocus>' + esc(okText) + '</button></div>',
      function () { if (!done) { done = true; res(false); } });
    $('mb').onclick = function (e) { var b = e.target.closest('button[data-a]'); if (!b) return; var v = b.getAttribute('data-a') === '1'; done = true; closeModal(); res(v); };
  });
}

/* ===================== feedback (sound / vibration) ===================== */
var FB = {
  ax: null,
  unlock: function () { try { if (!this.ax) { var AC = window.AudioContext || window.webkitAudioContext; if (AC) this.ax = new AC(); } if (this.ax && this.ax.state === 'suspended') this.ax.resume(); } catch (e) { } },
  tone: function (f, ms, at) {
    try {
      var a = this.ax; if (!a || !S.fb.snd) return;
      var t = a.currentTime + (at || 0), o = a.createOscillator(), g = a.createGain(); o.type = 'sine'; o.frequency.value = f;
      g.gain.setValueAtTime(0.0001, t); g.gain.exponentialRampToValueAtTime(0.3, t + 0.01); g.gain.exponentialRampToValueAtTime(0.0001, t + ms / 1000);
      o.connect(g); g.connect(a.destination); o.start(t); o.stop(t + ms / 1000 + 0.02);
    } catch (e) { }
  },
  /* success != warning != error (sound + vibration + visual; iPhone has no vibration API, visual channel carries it) */
  go: function (kind) {
    var vib = S.fb.vib && ('vibrate' in navigator);
    try {
      if (kind === 'ok') { this.tone(1760, 120); if (vib) navigator.vibrate(60); }
      else if (kind === 'dup') { this.tone(880, 70); this.tone(880, 70, 0.13); if (vib) navigator.vibrate([40, 60, 40]); }
      else if (kind === 'err') { this.tone(440, 200); if (vib) navigator.vibrate([90, 60, 90]); }
      else if (kind === 'ser') { this.tone(1760, 60); if (vib) navigator.vibrate(30); }
      else if (kind === 'soft') { this.tone(988, 40); if (vib) navigator.vibrate(15); }
      else if (kind === 'inbasket') { this.tone(1319, 90); this.tone(659, 220, 0.11); if (vib) navigator.vibrate(350); }   /* V13: "zaten sepette" */
      else if (kind === 'stock') { this.tone(660, 90); this.tone(660, 90, 0.15); this.tone(660, 90, 0.3); if (vib) navigator.vibrate([120, 80, 120, 80, 120]); }
    } catch (e) { }
  }
};

/* ===================== copy / download ===================== */
function copyText(t, okMsg) {
  /* clipboard API first (secure context); execCommand fallback inside the same click */
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(t).then(function () { toast(okMsg || 'Kopyalandı'); }, function () { copyLegacy(t, okMsg); });
    return;
  }
  copyLegacy(t, okMsg);
}
function copyLegacy(t, okMsg) {
  var ta = document.createElement('textarea'); ta.value = t; ta.readOnly = true;
  ta.style.position = 'fixed'; ta.style.top = '0'; ta.style.left = '0'; ta.style.width = '1px'; ta.style.height = '1px'; ta.style.opacity = '0'; ta.style.fontSize = '16px';
  document.body.appendChild(ta); ta.focus(); ta.select();
  try { ta.setSelectionRange(0, t.length); } catch (e) { }
  var ok = false; try { ok = document.execCommand('copy'); } catch (e) { }
  document.body.removeChild(ta);
  toast(ok ? (okMsg || 'Kopyalandı') : 'Kopyalanamadı. Metni seçip basılı tutun → Kopyala.', !ok);
  return ok;
}
function download(name, text) {
  var b = new Blob([text], { type: 'application/octet-stream' }), u = URL.createObjectURL(b), a = document.createElement('a');
  a.href = u; a.download = name; a.rel = 'noopener'; document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(function () { URL.revokeObjectURL(u); }, 40000);   /* iOS downloads asynchronously */
}

/* ===================== views ===================== */
var VIEW = { cur: null };
var VIEWS = ['pair', 'who', 'scan', 'basket', 'stock', 'hist', 'cam', 'set'];
function showView(v) {
  if (VIEW.cur === v) { render(v); return; }
  if (CAM.stream || CAM.want) CAM.stop();
  if (VIEW.cur === 'cam') dropCamPanel();
  VIEW.cur = v;
  VIEWS.forEach(function (x) { $('v-' + x).hidden = x !== v; });
  document.body.classList.toggle('v-pair', v === 'pair' || v === 'who');
  var navV = v === 'cam' ? 'set' : v;   /* camera settings live under Ayarlar (no tab of their own since V8) */
  [].forEach.call(document.querySelectorAll('.nav button'), function (b) { if (b.getAttribute('data-v') === navV) b.setAttribute('aria-current', 'page'); else b.removeAttribute('aria-current'); });
  var t = { pair: 'Eşleştir', who: 'Eşleşme', scan: 'Tara', basket: 'Sepet', stock: 'Stok', hist: 'Geçmiş', cam: 'Kamera ayarları', set: 'Ayarlar' }[v];
  document.title = DEMO === 'selftest' ? document.title : (t ? t + ' · OMS Saha' : 'OMS Saha');
  try { window.scrollTo(0, 0); } catch (e) { }
  render(v);
}
function render(v) {
  if (v === 'scan') renderScan(); else if (v === 'basket') renderBasket(); else if (v === 'hist') renderHist(); else if (v === 'set') renderSet(); else if (v === 'who') renderWho(); else if (v === 'cam') renderCam();
  else if (v === 'stock') renderStock();
  renderCounts();
}
[].forEach.call(document.querySelectorAll('.nav button'), function (b) { b.addEventListener('click', function () { showView(b.getAttribute('data-v')); }); });
function renderCounts() {
  var n = lines().length, bd = $('nbadge');
  bd.hidden = n === 0; bd.textContent = n > 999 ? '999+' : String(n);
  $('stagecnt').textContent = String(n);
  $('bstagebasket').setAttribute('aria-label', 'Sepeti aç (' + n + ' satır)');
  $('scount').textContent = fmtQty(n);
  $('hname').textContent = S.pair ? '· ' + (S.pair.name || S.pair.dev) : '';
}

/* ----- line presentation ----- */
function lineInfo(l) {
  var p = pv(l.raw, l.fmt) || {};
  var title = p.mpn || p.supplierPn || shortRaw(l.raw);
  var sub = [];
  if (p.mfr) sub.push(p.mfr);
  if (p.mpn && p.supplierPn && p.supplierPn !== p.mpn) sub.push(p.supplierPn);
  if (p.lot) sub.push('Lot ' + p.lot);
  if (l.loc) sub.push('Raf ' + l.loc);
  if (l.note) sub.push(l.note);
  var qty = l.qty != null ? { t: fmtQty(l.qty), src: 'girildi', n: l.qty } : (p.qty != null && p.qty !== '' ? { t: fmtQty(p.qty), src: 'etiket', n: +String(p.qty).replace(/[^0-9]/g, '') } : null);
  var rn = reelsOf(l);
  if (qty) qty.total = rn > 1 && isFinite(qty.n) && qty.n > 0 ? fmtQty(qty.n * rn) : qty.t;
  if (rn > 1) { if (qty) qty.t = rn + ' × ' + qty.t; sub.unshift(rn + ' makara' + (qty ? ' · toplam ' + qty.total : '')); }
  var tag = '';
  if (!qty) tag = '<span class="tag warn">' + ico('i-warn') + 'miktar yok</span>';
  else if (!p.mpn && !p.supplierPn) tag = '<span class="tag neutral">önizleme yok</span>';
  else if (qty.src === 'etiket') tag = '<span class="tag neutral">etiketten</span>';
  else tag = '<span class="tag info">' + ico('i-edit') + 'elle</span>';
  return { p: p, title: title, sub: sub.join(' · ') || shortRaw(l.raw), qty: qty, tag: tag, stk: stockOf(l, p) };
}
function stockHtml(st) {
  if (!st) return '';
  if (st.same) return '<span class="t3"><span class="tag warn">' + ico('i-warn') + 'STOKTA · aynı makara</span>' + esc(stockWhere(st.same)) + '</span>';
  if (st.mpn) return '<span class="t3">' + ico('i-box') + esc(stockLine(st.mpn)) + '</span>';
  return '';
}
function lineRow(l, isNew) {
  var i = lineInfo(l);
  return '<li><button type="button" class="li' + (isNew ? ' new' : '') + (i.stk && i.stk.same ? ' same' : '') + '" data-id="' + esc(l.id) + '">' +
    '<span class="t1">' + esc(i.title) + '</span><span class="q">' + (i.qty ? esc(i.qty.t) : '\u2014') + '</span>' +
    '<span class="t2">' + esc(i.sub) + '</span><span class="s">' + i.tag + '</span>' + stockHtml(i.stk) + '</button></li>';
}

/* ----- scan view ----- */
var LASTADD = null;
function renderScan() {
  var ls = lines(), r = $('recent');
  if (document.activeElement !== $('defloc')) $('defloc').value = S.defLoc;
  smodeUi();
  if (!ls.length) { r.hidden = true; r.innerHTML = ''; return; }
  var h = '';
  for (var i = ls.length - 1; i >= 0 && i >= ls.length - 3; i--) h += lineRow(ls[i], LASTADD === ls[i].id);
  r.innerHTML = h; r.hidden = false; LASTADD = null;
  catFill(ls.slice(-3));
}
$('recent').addEventListener('click', function (e) { var b = e.target.closest('button.li'); if (b) editLine(b.getAttribute('data-id')); });
/* ----- scan mode: Seri | Manuel (spec section 19) ----- */
function serOk() { return !!(window.OmsSerial && typeof OmsSerial.create === 'function'); }
function smode() { return S.smode === 'serial' && serOk() ? 'serial' : 'manual'; }
function smodeUi() {
  var m = smode(), ok = serOk(), rs = document.querySelectorAll('input[name=smode]');
  for (var i = 0; i < rs.length; i++) {
    rs[i].checked = rs[i].value === m;
    rs[i].disabled = rs[i].value === 'serial' && !ok;
    rs[i].parentNode.classList.toggle('on', rs[i].checked);
  }
  $('smodehelp').textContent = !ok ? 'Seri tarama modülü (serialscan.js) yüklenemedi; internete bağlıyken sayfayı yenileyin. Şimdilik Manuel.' :
    m === 'serial' ? 'Kamerayı makaraların önünden sırayla geçirin. Her yeni makara kendiliğinden sepete eklenir, hiçbir şey sorulmaz. Bitince Bitti\'ye basın.' :
    'Her makara okununca miktar, raf no ve not sorulur. Kaydedince tarama kaldığı yerden sürer.';
  $('bcam').lastChild.textContent = m === 'serial' ? 'Seri taramayı başlat' : 'Kamerayı aç ve tara';
}
$('smode').addEventListener('change', function (e) {
  if (e.target.name !== 'smode') return;
  S.smode = e.target.value === 'serial' ? 'serial' : 'manual'; LS.set(K.smode, S.smode); smodeUi();
});
$('bgobasket').addEventListener('click', function () { showView('basket'); });
$('defloc').addEventListener('change', function () { S.defLoc = this.value.trim(); LS.set(K.loc, S.defLoc); });
$('defloc').addEventListener('keydown', function (e) { if (e.key === 'Enter') { e.preventDefault(); this.blur(); } });
function manualAdd() {
  var el = $('code'), v = el.value;
  if (!v.replace(/\s/g, '')) { el.setAttribute('aria-invalid', 'true'); setMsg('cammsg', 'Önce kodu yazın veya okutun.', 'err'); el.focus(); return; }
  var r = addLine(v, '', 'manual');
  if (r.ok) { el.value = ''; el.removeAttribute('aria-invalid'); setMsg('cammsg', ''); }
  if (r.ok && smode() === 'manual') { editLine(r.line.id, { after: true, onDone: function () { try { $('code').focus({ preventScroll: true }); } catch (e) { } } }); return; }
  el.focus();
}
$('badd').addEventListener('click', manualAdd);
$('code').addEventListener('keydown', function (e) { if (e.key === 'Enter') { e.preventDefault(); manualAdd(); } });
$('code').addEventListener('input', function () { this.removeAttribute('aria-invalid'); });
$('bkbd').addEventListener('click', function () { var c = $('code'); c.focus(); try { c.scrollIntoView({ block: 'center' }); } catch (e) { } });

/* add one scanned/typed code to the basket. src: camera | manual | photo | wedge */
function addLine(raw, fmt, src) {
  raw = String(raw == null ? '' : raw).replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, '');
  if (!raw) return { err: 'empty' };
  if (raw.length > MAX_RAW) { FB.go('err'); toast('Kod çok uzun (' + raw.length + ' karakter, en çok ' + MAX_RAW + ').', true); return { err: 'long' }; }
  var b = ensureBasket();
  for (var i = 0; i < b.lines.length; i++) if (b.lines[i].raw === raw) {
    if (src !== 'camera' && src !== 'serial') {
      var dl = b.lines[i];
      FB.go('inbasket');
      toast('Bu etiket zaten sepette (' + reelsOf(dl) + ' makara): ' + lineInfo(dl).title + '. Aynı etiketli başka bir makaraysa "+1 makara"ya dokunun.', true, { t: '+1 makara', fn: function () { addReel(dl.id); } });
    }
    return { dup: true, line: b.lines[i] };
  }
  if (b.lines.length >= MAX_LINES) { FB.go('err'); toast('Sepet dolu (' + MAX_LINES + ' satır). Önce bu sepeti gönderin.', true); return { err: 'full' }; }
  var l = { id: uuid(), raw: raw, fmt: fmt || '', qty: null, loc: S.defLoc || '', note: '', ts: iso() };
  b.lines.push(l); LASTADD = l.id;
  saveBasket();
  askPersist();
  if (src !== 'camera' && src !== 'serial') {
    stockCheck(l).then(function (st) {
      var title = lineInfo(l).title, undo = { t: 'Geri al', fn: function () { removeLine(l.id, true); } };
      if (st.same) { FB.go('stock'); toast('STOKTA · aynı makara: ' + title + ' (' + stockWhere(st.same) + '). Yine de sepete eklendi.', true, undo); }
      else { FB.go('ok'); toast('Sepete eklendi: ' + title + (st.mpn ? ' · ' + stockLine(st.mpn) : ''), false, undo); }
      if (st.same) render(VIEW.cur);
    });
  }
  render(VIEW.cur);
  return { ok: true, line: l };
}
/* V13: identical labels (same part, same lot, same quantity: many labels carry no reel serial) are ONE line with a reel count;
   the PC adds quantity x reels. Absent / 1 = one reel (older baskets stay valid). */
var MAX_REELS = 999;
function reelsOf(l) { var n = l && l.reels; return (typeof n === 'number' && n >= 1) ? Math.min(MAX_REELS, Math.floor(n)) : 1; }
function addReel(id) {
  var ls = lines(), l = null;
  for (var i = 0; i < ls.length; i++) if (ls[i].id === id) l = ls[i];
  if (!l) return 0;
  var n = reelsOf(l);
  if (n >= MAX_REELS) { FB.go('err'); toast('En çok ' + MAX_REELS + ' makara.', true); return n; }
  l.reels = n + 1; saveBasket(); render(VIEW.cur); FB.go('ok');
  if (CAM.serial && SER.s) { SER.n++; SER.paint(); }
  var inf = lineInfo(l), msg = 'Aynı etiketten ' + l.reels + '. makara · ' + inf.title + (inf.qty ? ' · toplam ' + inf.qty.total : '');
  if (CAM.stream && CAM.mode === 'scan') { CAM.state(msg, 'ok'); CAM.flash('hit'); CAM.armIdle(); }
  else toast(msg, false, { t: 'Geri al', fn: function () { l.reels = Math.max(1, reelsOf(l) - 1); if (l.reels === 1) delete l.reels; saveBasket(); render(VIEW.cur); } });
  return l.reels;
}
/* "+1 makara" on the camera stage, offered for 8 s after a "zaten sepette" read (never automatic: re-reading the same reel by
   accident must not count it twice) */
var PLUS = {
  id: null, t: null,
  show: function (line) {
    if (!line) return;
    this.id = line.id; clearTimeout(this.t);
    $('bplus').querySelector('span').textContent = '+1 makara (şu an ' + reelsOf(line) + ')';
    $('cplus').hidden = false; $('chelp').hidden = true;
    var self = this; this.t = setTimeout(function () { self.hide(); }, 8000);
  },
  hide: function () { clearTimeout(this.t); this.id = null; $('cplus').hidden = true; }
};
$('bplus').addEventListener('click', function () { var id = PLUS.id; if (!id) return; var n = addReel(id); var l = null; lines().forEach(function (x) { if (x.id === id) l = x; }); if (l && n) PLUS.show(l); });

function removeLine(id, silent) {
  var ls = lines();
  for (var i = 0; i < ls.length; i++) if (ls[i].id === id) {
    var l = ls.splice(i, 1)[0], at = i;
    SER.forget(l.raw);
    saveBasket(); render(VIEW.cur);
    if (!silent) toast('Satır silindi', false, { t: 'Geri al', fn: function () { var b = ensureBasket(); b.lines.splice(Math.min(at, b.lines.length), 0, l); saveBasket(); render(VIEW.cur); } });
    return;
  }
}
var persistAsked = false;
function askPersist() {
  if (persistAsked || DEMO) return; persistAsked = true;
  try { if (navigator.storage && navigator.storage.persist) navigator.storage.persisted().then(function (p) { if (!p) navigator.storage.persist(); }).catch(function () { }); } catch (e) { }
}

/* keyboard-wedge scanner while no field has focus: fast keystrokes ending with Enter */
var WEDGE = { buf: '', last: 0, slow: 0 };
document.addEventListener('keydown', function (e) {
  if (VIEW.cur !== 'scan' || !$('modal').hidden) return;
  var a = document.activeElement, tag = a && a.tagName;
  if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || (a && a.isContentEditable)) return;
  if (e.ctrlKey || e.altKey || e.metaKey) return;
  var now = Date.now();
  if (now - WEDGE.last > 120) { WEDGE.buf = ''; WEDGE.slow = 0; }
  WEDGE.last = now;
  if (e.key === 'Enter') {
    if (WEDGE.buf.length >= 4 && WEDGE.slow <= 1) { e.preventDefault(); var wr = addLine(WEDGE.buf, '', 'wedge'); if (wr.ok && smode() === 'manual') editLine(wr.line.id, { after: true, noFocus: true }); }
    WEDGE.buf = ''; return;
  }
  if (e.key && e.key.length === 1) { if (WEDGE.buf && now - WEDGE.prev > 60) WEDGE.slow++; WEDGE.buf += e.key; WEDGE.prev = now; }
});

/* ----- basket view ----- */
function renderBasket() {
  var ls = lines(), n = ls.length, on = driveOn(), lan = lanOn();
  if (!DRV.busy && !LANS.busy) $('bsendtx').textContent = lan ? 'Yerel ağdan gönder' : on ? 'OneDrive\'a gönder' : 'Gönder';
  $('bsend').querySelector('use').setAttribute('href', lan ? '#i-lan' : on ? '#i-cloudup' : '#i-send');
  $('bsendmore').hidden = !(on || lan);
  renderLanLine(); lanKick();
  $('bempty').hidden = n > 0; $('blist').hidden = n === 0; $('sendbar').hidden = n === 0; $('bdz').hidden = n === 0; $('bhint').hidden = n === 0;
  $('bcnt').textContent = n ? fmtQty(n) + ' satır' : '';
  $('bcnt').hidden = !n;
  var h = '';
  for (var i = n - 1; i >= 0; i--) h += lineRow(ls[i], LASTADD === ls[i].id);
  $('blist').innerHTML = h; LASTADD = null;
  catFill(ls);
}
$('blist').addEventListener('click', function (e) { var b = e.target.closest('button.li'); if (b) editLine(b.getAttribute('data-id')); });
$('bgoscan').addEventListener('click', function () { showView('scan'); });
$('bclear').addEventListener('click', async function () {
  var n = lines().length; if (!n) return;
  if (!(await ask('Sepet boşaltılsın mı?', 'Sepetteki <b>' + fmtQty(n) + ' satır</b> silinecek. Bu işlem geri alınamaz.', 'Sepeti boşalt', true))) return;
  S.basket = null; await saveBasket(true); render(VIEW.cur); toast('Sepet boşaltıldı');
});
$('bsend').addEventListener('click', function () {
  if (!S.basket || !S.basket.lines.length) return;
  if (lanOn()) lanSend(S.basket, false, this); else if (driveOn()) driveSend(S.basket, false, this); else startSend(S.basket, false);
});
$('bsendmore').addEventListener('click', function () { if (S.basket && S.basket.lines.length) startSend(S.basket, false); });

/* ----- line editor ----- */
/* opt (Manuel mode, opened right after an add): {after, onDone (any close), noFocus (keyboard-wedge source)} */
function editLine(id, opt) {
  opt = opt || {};
  var ls = lines(), l = null;
  for (var i = 0; i < ls.length; i++) if (ls[i].id === id) l = ls[i];
  if (!l) { if (opt.onDone) opt.onDone(); return; }
  var camWait = !!(opt.after && CAM.stream && CAM.mode === 'scan');
  var inf = lineInfo(l), p = inf.p, kv = '';
  [['Üretici', p.mfr], ['<span lang="en">MPN</span>', p.mpn], ['Tedarikçi kodu', p.supplierPn], ['Tedarikçi', p.supplier], ['Lot', p.lot], ['Tarih kodu', p.dc],
   ['Etiketteki miktar', p.qty != null && p.qty !== '' ? fmtQty(p.qty) : null], ['Kod türü', l.fmt ? fmtName(l.fmt) : null]].forEach(function (r) {
    if (r[1] != null && r[1] !== '') kv += '<div class="kv"><dt>' + r[0] + '</dt><dd>' + esc(r[1]) + '</dd></div>';
  });
  if (inf.stk && inf.stk.mpn && !inf.stk.same) kv += '<div class="kv"><dt>Stokta</dt><dd>' + esc(stockWhereMpn(inf.stk.mpn)) + '</dd></div>';
  var ph = p.qty != null && p.qty !== '' ? 'Boş = etiketten (' + fmtQty(p.qty) + ')' : 'Etikette miktar yok: girin';
  var warn = (inf.stk && inf.stk.same ? '<div class="note warn" style="margin-bottom:12px">' + ico('i-warn') + '<div><b>Bu makara zaten stokta</b> (' + esc(stockWhere(inf.stk.same)) +
      '). Aynı etiket PC\'de kayıtlı; gönderilirse PC bunu "aynı makara" durumu olarak işaretler. Yeni bir makara değilse satırı silin.</div></div>' : '');
  var info = (kv ? '<dl class="kvl">' + kv + '</dl>' : '<div class="note warn">' + ico('i-warn') + '<div>Bu kod sayfada önizlenemedi. PC yine de çözmeyi dener; gerekirse miktarı girin.</div></div>') +
    '<p class="small" style="margin:8px 0 0">Önizlemedir; kesin çözümleme PC\'de yapılır. Eklenme: ' + esc(fmtTime(Date.parse(l.ts) || Date.now())) + '</p>';
  var raw = '<div class="lbl" style="margin-top:16px">Ham kod</div><div class="rawbox">' + showRaw(l.raw) + '</div>';
  var flds = '<div class="fld" style="margin-top:16px"><label for="e_qty">Miktar</label>' +
    '<input id="e_qty" type="number" inputmode="numeric" pattern="[0-9]*" min="1" step="1" enterkeyhint="next" placeholder="' + esc(ph) + '" value="' + (l.qty != null ? esc(l.qty) : '') + '"' +
      (opt.after && !opt.noFocus && !(p.qty != null && p.qty !== '') ? ' autofocus' : '') + '>' +
    '<p class="help">Makara başına miktar. Kısmen kullanılmış makarada kalan miktarı yazın; etiketteki miktarın yerine geçer.</p></div>' +
    '<div class="fld"><label for="e_reels">Makara sayısı <small>(aynı etiketli)</small></label><div class="stepper">' +
      '<button type="button" class="btn btn-sec" id="e_rm" aria-label="Bir makara azalt">−</button>' +
      '<input id="e_reels" type="number" inputmode="numeric" pattern="[0-9]*" min="1" max="' + MAX_REELS + '" step="1" value="' + reelsOf(l) + '">' +
      '<button type="button" class="btn btn-sec" id="e_rp" aria-label="Bir makara artır">+</button></div>' +
    '<p class="help" id="e_rhelp">Elinizde aynı etiketli birden çok makara varsa hepsini tek tek okutmanız gerekmez: sayıyı yazın. PC miktar × makara sayısı kadar ekler.</p></div>' +
    '<div class="fld"><label for="e_loc">Raf No <small>(boş bırakılabilir)</small></label><input id="e_loc" type="text" autocapitalize="characters" autocorrect="off" spellcheck="false" autocomplete="off" maxlength="64" enterkeyhint="next" value="' + esc(l.loc) + '"></div>' +
    '<div class="fld"><label for="e_note">Not</label><input id="e_note" type="text" autocomplete="off" maxlength="200" enterkeyhint="done" value="' + esc(l.note) + '"></div>';
  var acts = '<div id="e_msg" class="msg" role="status" aria-live="polite"></div>' +
    '<div class="acts r"><button type="button" class="btn btn-sec" id="e_copy">' + ico('i-copy') + 'Kodu kopyala</button>' +
    '<button type="button" class="btn btn-sec" id="e_cancel">Kapat</button><button type="button" class="btn btn-pri" id="e_save">' + ico('i-save') + (camWait ? 'Kaydet, taramaya dön' : 'Kaydet') + '</button></div>';
  var del = '<div class="danger-row"><button type="button" class="btn btn-danger-ghost" id="e_del">' + ico('i-del') + 'Satırı sil</button></div>';
  /* right after an add (Manuel): the three fields and Kaydet first, label details folded below */
  var h = opt.after ? '<div class="note ok" style="margin-bottom:12px">' + ico('i-check') + '<div><b>Sepete eklendi.</b> ' + (camWait ? 'Bilgileri girin; kaydedince ya da kapatınca tarama sürer.' : 'Gerekirse miktar, raf no ve not girin (hepsi boş bırakılabilir).') + '</div></div>' + warn + flds + acts +
      '<details class="more"><summary>' + ico('i-info') + 'Etiket bilgileri' + ico('i-chev', 'chev') + '</summary><div class="inner">' + info + raw + '</div></details>' + del :
    warn + info + flds + raw + acts + del;
  openModal(inf.title, h, opt.onDone || null);
  function save() {
    var q = $('e_qty').value.trim(), qn = null;
    if (q !== '') {
      qn = Number(q.replace(/[.\s]/g, '').replace(',', '.'));
      if (!isFinite(qn) || qn <= 0 || Math.floor(qn) !== qn || qn > 1e9) { $('e_qty').setAttribute('aria-invalid', 'true'); setMsg('e_msg', 'Miktar 1 veya daha büyük bir tam sayı olmalı (ya da boş).', 'err'); $('e_qty').focus(); return; }
    }
    var rv = Number($('e_reels').value);
    if (!isFinite(rv) || rv < 1 || Math.floor(rv) !== rv || rv > MAX_REELS) { $('e_reels').setAttribute('aria-invalid', 'true'); setMsg('e_msg', 'Makara sayısı 1 ile ' + MAX_REELS + ' arasında bir tam sayı olmalı.', 'err'); $('e_reels').focus(); return; }
    if (rv > 1) l.reels = rv; else delete l.reels;
    l.qty = qn; l.loc = $('e_loc').value.trim().slice(0, 64); l.note = $('e_note').value.trim().slice(0, 200);
    saveBasket(); closeModal(); render(VIEW.cur); toast('Satır kaydedildi');
  }
  $('e_save').onclick = save;
  $('e_cancel').onclick = function () { closeModal(); };
  $('e_copy').onclick = function () { copyText(l.raw, 'Ham kod kopyalandı'); };
  $('e_del').onclick = function () { closeModal(); removeLine(l.id); };
  $('e_qty').oninput = function () { this.removeAttribute('aria-invalid'); setMsg('e_msg', ''); };
  function stepReels(d) { var v = Math.floor(Number($('e_reels').value) || 1) + d; $('e_reels').value = String(Math.max(1, Math.min(MAX_REELS, v))); $('e_reels').removeAttribute('aria-invalid'); setMsg('e_msg', ''); }
  $('e_rm').onclick = function () { stepReels(-1); };
  $('e_rp').onclick = function () { stepReels(1); };
  $('e_reels').oninput = function () { this.removeAttribute('aria-invalid'); setMsg('e_msg', ''); };
  ['e_qty', 'e_loc'].forEach(function (id, i) { $(id).onkeydown = function (e) { if (e.key === 'Enter') { e.preventDefault(); $(i ? 'e_note' : 'e_loc').focus(); } }; });
  $('e_note').onkeydown = function (e) { if (e.key === 'Enter') { e.preventDefault(); save(); } };
}

/* ----- history ----- */
function renderHist() {
  if (pruneHist()) saveHist();
  var hs = S.hist;
  $('hempty').hidden = hs.length > 0;
  var h = '';
  hs.forEach(function (x, i) {
    var last = lastAt(x), n = x.lines.length, items = '';
    x.lines.slice(0, 200).forEach(function (l) { var inf = lineInfo(l); items += '<li>' + esc(inf.title) + (inf.qty ? ' · ' + esc(inf.qty.t) : '') + (l.loc ? ' · ' + esc(l.loc) : '') + '</li>'; });
    if (n > 200) items += '<li>… ve ' + fmtQty(n - 200) + ' satır daha</li>';
    var ns = x.sends.length, st = histState(x);
    h += '<div class="card hcard"><div class="hh"><div><b>' + esc(fmtTime(last)) + '</b><span class="small">' + fmtQty(n) + ' satır · ' +
      (ns > 1 ? ns + ' kez gönderildi' : ns ? 'gönderildi' : 'henüz gönderilmedi') + (x.via === 'direct' && ns ? ' (doğrudan)' : x.via === 'drive' && ns ? ' (OneDrive)' : x.via === 'lan' && ns ? ' (yerel ağ)' : '') + (x.user ? ' · ' + esc(x.user) : '') + '</span></div>' +
      st.tag + '</div>' + st.more +
      '<details><summary>Satırları göster' + ico('i-chev') + '</summary><ol>' + items + '</ol></details>' +
      '<div class="acts"><button type="button" class="btn btn-sec" data-a="re" data-i="' + i + '">' + ico('i-send') + 'Yeniden gönder</button>' +
      '<button type="button" class="btn btn-danger-ghost" data-a="del" data-i="' + i + '" style="flex:0 0 auto">' + ico('i-del') + 'Sil</button></div></div>';
  });
  $('hlist').innerHTML = h;
}
$('hlist').addEventListener('click', async function (e) {
  var b = e.target.closest('button[data-a]'); if (!b) return;
  var x = S.hist[+b.getAttribute('data-i')]; if (!x) return;
  var act = b.getAttribute('data-a');
  if (act === 're') { if (lanOn()) lanSend(x, true, b); else if (driveOn()) driveSend(x, true, b); else startSend(x, true); return; }
  if (act === 'flush') { if (x.ch === 'drive') driveFlush(b); else outboxFlush(b); return; }
  if (act === 'dlogin') { driveSignIn(b).then(function (ok) { if (ok) driveFlush(null, true); }); return; }
  if (act === 'login') { mailSignIn(b).then(function (ok) { if (ok) outboxFlush(null); }); return; }
  if (!(await ask('Geçmişten silinsin mi?', esc(fmtTime(lastAt(x))) + ' tarihli ' + fmtQty(x.lines.length) + ' satırlık sepet bu cihazdan silinecek. ' +
    (x.out ? '<b>Giden kutusundaki gönderimi de iptal edilir.</b> ' : '') + 'PC\'deki kayıtlar etkilenmez.', 'Sil', true))) return;
  if (x.out) { if (x.ch === 'drive') driveRemove(x.basket); else outboxRemove(x.basket); }
  var k = S.hist.indexOf(x); if (k >= 0) S.hist.splice(k, 1);
  await saveHist(); renderHist(); toast('Silindi');
});

/* ----- settings ----- */
function renderSet() {
  var p = S.pair || {}, kv = '';
  function kvs(rows) { return rows.filter(function (r) { return r[1]; }).map(function (r) { return '<div class="kv"><dt>' + r[0] + '</dt><dd' + (r[2] ? ' class="' + r[2] + '"' : '') + '>' + esc(r[1]) + '</dd></div>'; }).join(''); }
  kv = kvs([['Cihaz adı', p.name], ['Cihaz kimliği', p.dev, 'mono'], ['Bilgisayar', p.pc], ['Eşleşme', p.at ? fmtTime(p.at) : '']]);
  $('pairkv').innerHTML = kv;
  $('mailkv').innerHTML = kvs([['E-posta konusu', subj(), 'mono']]);
  if (document.activeElement !== $('sto')) $('sto').value = p.to || '';
  if (document.activeElement !== $('suser')) $('suser').value = S.user;
  $('c_mail').hidden = !S.pair;
  if (!ADV.init) { ADV.init = true; $('advset').open = LS.get(K.adv) === '1'; }
  renderCloudCard();
  $('o_snd').checked = S.fb.snd; $('o_vib').checked = S.fb.vib; $('o_vibw').hidden = !('vibrate' in navigator);
  themeUi();
  var dec = DEC.kind === 'native' ? 'Tarayıcının yerleşik okuyucusu' : DEC.kind ? 'ZXing (sayfa içi)' : (('BarcodeDetector' in window) ? 'Yerleşik okuyucu (denenecek)' : 'ZXing (sayfa içi)');
  var store = DB.mode === 'idb' ? 'Tarayıcı veritabanı (IndexedDB)' : DB.mode === 'ls' ? 'Yerel depolama (localStorage)' : 'Yalnız bellek: sayfa kapanınca sepet kaybolur!';
  var defl = false; try { defl = !!(window.OmsEnvelope && OmsEnvelope.supportsDeflate()); } catch (e) { }
  var a = [['Sürüm', APP.ver || 'bilinmiyor'], ['Çevrimdışı kullanım', APP.sw], ['Sepet deposu', store], ['Sıkıştırma', defl ? 'açık (deflate)' : 'yok (tarayıcı desteklemiyor)'], ['Kod okuyucu', dec]];
  $('appkv').innerHTML = a.map(function (r) { return '<div class="kv"><dt>' + r[0] + '</dt><dd>' + esc(r[1]) + '</dd></div>'; }).join('');
  var standalone = mql('(display-mode: standalone)') || navigator.standalone === true;
  var ios = /iP(hone|ad|od)/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  $('installhint').textContent = standalone ? 'Ana ekrandan uygulama olarak çalışıyor.' :
    ios ? 'Ana ekrana eklemek için Safari\'de Paylaş düğmesine, sonra "Ana Ekrana Ekle"ye dokunun. Böylece sayfa internetsiz de açılır.' :
    APP.bip ? 'Bu sayfayı uygulama gibi ana ekrana ekleyebilirsiniz.' : 'Tarayıcı menüsünden "Ana ekrana ekle" / "Uygulamayı yükle" ile kurabilirsiniz.';
  $('binstall').hidden = !APP.bip || standalone;
  renderCatCard(); renderLan(); renderMs();
}
$('bsetsave').addEventListener('click', function () {
  var u = $('suser').value.trim();
  S.user = u; LS.set(K.user, u);
  setMsg('setmsg', 'Kaydedildi.', 'ok'); renderCounts();
});
$('suser').addEventListener('keydown', function (e) { if (e.key === 'Enter') { e.preventDefault(); $('bsetsave').click(); } });
$('bmailsave').addEventListener('click', function () {
  var to = $('sto').value.trim();
  if (to && !isEmail(to)) { $('sto').setAttribute('aria-invalid', 'true'); setMsg('mailmsg', 'E-posta adresi geçersiz görünüyor (ör. stok@hotmail.com).', 'err'); $('sto').focus(); return; }
  $('sto').removeAttribute('aria-invalid');
  if (S.pair) { S.pair.to = to; savePair(); }
  setMsg('mailmsg', to ? 'Kaydedildi.' : 'Kaydedildi: adres yok (e-posta seçenekleri adres sorar).', 'ok');
});
$('sto').addEventListener('keydown', function (e) { if (e.key === 'Enter') { e.preventDefault(); $('bmailsave').click(); } });
$('sto').addEventListener('input', function () { this.removeAttribute('aria-invalid'); });
var ADV = { init: false };
$('advset').addEventListener('toggle', function () { LS.set(K.adv, this.open ? '1' : '0'); });
$('o_snd').addEventListener('change', function () { S.fb.snd = this.checked; LS.set(K.snd, this.checked ? '1' : '0'); if (this.checked) { FB.unlock(); FB.go('ok'); } });
$('o_vib').addEventListener('change', function () { S.fb.vib = this.checked; LS.set(K.vib, this.checked ? '1' : '0'); });
function themeUi() { var cur = OMSTHEME.get(), rs = document.querySelectorAll('input[name=theme]'); for (var i = 0; i < rs.length; i++) { rs[i].checked = rs[i].value === cur; rs[i].parentNode.classList.toggle('on', rs[i].checked); } }
$('themeseg').addEventListener('change', function (e) { if (e.target.name === 'theme') { OMSTHEME.set(e.target.value); themeUi(); } });
$('bunpair').addEventListener('click', async function () {
  var n = lines().length;
  if (!(await ask('Eşleştirme kaldırılsın mı?', 'Bu cihaz artık sepet gönderemez. Yeniden eşleştirmek için PC\'de yeni kod üretmeniz gerekir.' +
    (n ? ' Sepetteki ' + fmtQty(n) + ' satır korunur.' : ''), 'Eşleştirmeyi kaldır', true))) return;
  S.pair = null; savePair(); showView('pair'); toast('Eşleştirme kaldırıldı');
});
$('binstall').addEventListener('click', function () {
  var e = APP.bip; if (!e) return; APP.bip = null;
  try { e.prompt(); e.userChoice.then(function () { renderSet(); }, function () { }); } catch (x) { }
});

/* ===================== pairing ===================== */
function envReady() {
  if (!window.OmsEnvelope) return 'Şifreleme modülü (envelope.js) yüklenemedi. İnternete bağlıyken sayfayı yenileyin.';
  if (!(window.crypto && crypto.subtle)) return 'Bu tarayıcı veya bağlantı şifrelemeyi desteklemiyor. Sayfayı https:// adresinden, güncel Chrome, Safari, Edge veya Firefox ile açın.';
  return '';
}
async function doPair(text, msgId) {
  var er = envReady(); if (er) { setMsg(msgId, er, 'err'); FB.go('err'); return false; }
  var p = null;
  try { p = OmsEnvelope.parsePairing(text); } catch (e) { p = null; }
  if (!p || !p.dev || !p.sec) { setMsg(msgId, 'Bu bir OMS eşleştirme kodu değil ya da kod eksik kopyalanmış. Kod "OMSP1:" ile başlar.', 'err'); FB.go('err'); return false; }
  var sec = typeof p.sec === 'string' ? p.sec : b64uEnc(p.sec);
  var ex = pairExtras(text, p);
  if (S.pair && S.pair.dev === p.dev && S.pair.sec === sec) {
    /* same device again: a newer code may carry new PC addresses / Client ID */
    var chg = false;
    if (ex.lan.length && ex.lan.join(' ') !== (S.pair.lan || []).join(' ')) { S.pair.lan = ex.lan; chg = true; }
    if (ex.cid && ex.cid !== S.pair.cid) { S.pair.cid = ex.cid; chg = true; }
    if (ex.cloud && ex.cloud !== S.cloud) setCloud(ex.cloud);
    if (chg) { savePair(); afterPairChange(); }
    toast(chg ? 'Eşleşme güncellendi (PC adresi / hesap ayarı)' : 'Bu cihaz zaten bu kodla eşleşmiş'); if (VIEW.cur === 'pair') showView(nextAfterPair()); return true;
  }
  if (S.pair && S.pair.dev !== p.dev) {
    if (!(await ask('Eşleşme değiştirilsin mi?', 'Bu cihaz şu anda "<b>' + esc(S.pair.name || S.pair.dev) + '</b>" olarak eşli. Yeni kod ("' + esc(p.name || p.dev) + '") ile değiştirilsin mi? Sepet ve geçmiş korunur.', 'Değiştir', false))) return false;
  }
  var same = S.pair && S.pair.dev === p.dev;
  S.pair = { v: 1, dev: p.dev, sec: sec, name: p.name || '', to: p.to || (S.pair ? S.pair.to : '') || '', pc: p.pc || '', subj: p.subj || SUBJ_DEFAULT, at: Date.now(),
    lan: ex.lan.length ? ex.lan : (same && S.pair.lan) || [], cid: ex.cid || (S.pair && S.pair.cid) || '' };
  savePair();
  if (ex.cloud) setCloud(ex.cloud);
  if (!same) { CATR.clear(); try { if (window.OmsCatalog && OmsCatalog.clear) OmsCatalog.clear(); } catch (e) { } }
  afterPairChange();
  FB.go('ok');
  if (CAM.stream) CAM.stop();
  showView(nextAfterPair());
  toast('Eşleşti: ' + (S.pair.name || S.pair.dev));
  return true;
}
function nextAfterPair() { return 'who'; }   /* "Hazır!" welcome after every new pairing */
function needWelcome() { return !S.user || LS.get(K.welcomed) !== '1'; }
$('bpair').addEventListener('click', function () {
  var t = $('pcode').value;
  if (!t.trim()) { $('pcode').setAttribute('aria-invalid', 'true'); setMsg('pmsg', 'Önce kodu yapıştırın.', 'err'); $('pcode').focus(); return; }
  doPair(t, 'pmsg');
});
$('pcode').addEventListener('input', function () { this.removeAttribute('aria-invalid'); setMsg('pmsg', ''); });
$('pcode').addEventListener('keydown', function (e) { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); $('bpair').click(); } });
$('bpairqr').addEventListener('click', function () { FB.unlock(); CAM.start('pair'); });
function renderWho() {
  var p = S.pair || {};
  $('whoinfo').textContent = 'Bu cihaz "' + (p.name || p.dev || '') + '" adıyla ' + (p.pc ? p.pc + ' bilgisayarına' : 'stok bilgisayarına') + ' bağlandı.';
  if (document.activeElement !== $('wname')) $('wname').value = S.user;
  cloudSeg('wcloud', 'wc');
  var dok = driveOk(), don = driveOn();
  $('whodrv').hidden = !dok;
  if (dok) { $('bwdrvtx').textContent = don ? 'OneDrive\'a bağlı ✓' : 'OneDrive\'a bağlan'; $('bwdrv').disabled = don; }
  var g3 = document.querySelector('#v-who .gsteps li:nth-child(3) .gt');
  if (g3) g3.textContent = don ? 'Sepet sekmesinde "OneDrive\'a gönder"e basın. Dosya PC\'ye kendiliğinden gider.' : (g3.getAttribute('data-t') || g3.textContent);
  if (g3 && !g3.getAttribute('data-t') && !don) g3.setAttribute('data-t', g3.textContent);
}
$('fwho').addEventListener('submit', function (e) {
  e.preventDefault();
  var u = $('wname').value.trim();
  if (!u) { $('wname').setAttribute('aria-invalid', 'true'); setMsg('wmsg', 'Adınızı yazın.', 'err'); $('wname').focus(); return; }
  S.user = u; LS.set(K.user, u); LS.set(K.welcomed, '1');
  showView('scan');
});
$('wname').addEventListener('input', function () { this.removeAttribute('aria-invalid'); setMsg('wmsg', ''); });
/* link opened from the PC's QR with the phone's own camera app: ...#OMSP1:xxxx */
function pairFromHash() {
  var h = location.hash || '';
  if (h.indexOf('OMSP1') < 0) return false;
  var code = h.replace(/^#/, '');
  try { code = decodeURIComponent(code); } catch (e) { }
  try { history.replaceState(history.state, '', location.pathname + location.search); } catch (e) { try { location.hash = ''; } catch (x) { } }
  doPair(code, 'pmsg');
  return true;
}
window.addEventListener('hashchange', pairFromHash);

/* ===================== send ===================== */
/* basket -> sealed armor + mail body (both send paths) */
async function sealCtx(b, fromHist) {
    var now = Date.now();
    var obj = {
      v: 1, type: 'basket', basket: b.basket, dev: S.pair.dev, devName: S.pair.name || '', user: S.user || '', created: b.created, sent: iso(now),
      lines: b.lines.map(function (l) { var o = { id: l.id, raw: l.raw, fmt: l.fmt || '', qty: l.qty == null ? null : l.qty, loc: l.loc || '', note: l.note || '', ts: l.ts }; if (reelsOf(l) > 1) o.reels = reelsOf(l); return o; })
    };
    var json = JSON.stringify(obj), deflate = false;
    try { deflate = !!OmsEnvelope.supportsDeflate(); } catch (e) { }
    var armor = await OmsEnvelope.seal(json, S.pair.dev, secBytes(), { deflate: deflate });
    var wrapped = String(OmsEnvelope.wrapForMail(armor)).replace(/\r?\n/g, '\r\n');
    var body = 'OMS saha sepeti: ' + b.lines.length + ' satır, ' + (S.pair.name || S.pair.dev) + (S.user ? ' / ' + S.user : '') + ', ' + fmtTime(now) + '\r\n' +
      'Bu e-postayı değiştirmeden gönderin. Konu satırı "' + subj() + '" kalmalı.\r\n\r\n' + wrapped + '\r\n';
    return { b: b, hist: fromHist, armor: armor, body: body, name: 'oms_sepet_' + stamp(now) + '.omsb', date: new Date(now), used: false, pv: '', n: b.lines.length };
}
async function startSend(b, fromHist) {
  if (!S.pair) { toast('Önce cihazı eşleştirin.', true); showView('pair'); return; }
  var er = envReady(); if (er) { toast(er, true); return; }
  var sendBtn = $('bsend'); busy(sendBtn, true);
  var ctx;
  try { ctx = await sealCtx(b, fromHist); }
  catch (e) {
    busy(sendBtn, false);
    toast('Sepet hazırlanamadı: ' + ((e && e.message) || e), true); return;
  }
  busy(sendBtn, false);
  renderSendModal(ctx);
}
/* ===================== V7: cloud folder send (OmsCloud, cloudsend.js; spec section 16) ===================== */
var CLOUDS = { onedrive: 'OneDrive', gdrive: 'Google Drive', icloud: 'iCloud Drive', '': '' };
var CLOUD_OPTS = [['onedrive', 'OneDrive'], ['gdrive', 'Google Drive'], ['icloud', 'iCloud'], ['', 'Bilmiyorum']];
var SEND = { ctx: null };
function cloudNorm(v) { v = String(v == null ? '' : v); return CLOUDS.hasOwnProperty(v) ? v : ''; }
function cloudLabel(c) { return CLOUDS[cloudNorm(c)] || ''; }
function setCloud(v) { S.cloud = cloudNorm(v); LS.set(K.cloud, S.cloud); }
function cloudOk() { return !!(window.OmsCloud && typeof OmsCloud.send === 'function'); }
function cloudPlat() {
  try { if (window.OmsCloud && typeof OmsCloud.platform === 'function') { var p = OmsCloud.platform(); if (p === 'ios' || p === 'android' || p === 'desktop') return p; } } catch (e) { }
  return IS_IOS ? 'ios' : IS_ANDROID ? 'android' : 'desktop';
}
/* desktop Chrome / Edge: write straight into the chosen "OMS Saha" folder */
function folderOk() {
  try { return cloudOk() && cloudPlat() === 'desktop' && typeof OmsCloud.pickFolder === 'function' && !!(OmsCloud.folderSupported && OmsCloud.folderSupported()); } catch (e) { return false; }
}
function folderName() { try { return folderOk() && OmsCloud.folderName ? String(OmsCloud.folderName() || '') : ''; } catch (e) { return ''; } }
function isSahaName(n) { try { if (window.OmsCloud && OmsCloud.isSahaName) return !!OmsCloud.isSahaName(n); } catch (e) { } return /^\s*oms[\s_-]*saha\s*$/i.test(String(n || '')); }
/* 'folder' | 'share' | 'download' (OmsCloud) or 'local' (module missing: plain download) */
function cloudPrimary() {
  if (!cloudOk()) return 'local';
  if (folderOk()) return 'folder';
  var m = []; try { m = (OmsCloud.methods && OmsCloud.methods()) || []; } catch (e) { m = []; }
  m = m.filter(function (x) { return x === 'share' || x === 'download'; });
  return m[0] || 'download';
}
/* the module may restore the saved folder handle from IndexedDB asynchronously */
function cloudInit() {
  if (!window.OmsCloud) return;
  var r = null;
  try { r = typeof OmsCloud.ready === 'function' ? OmsCloud.ready() : null; } catch (e) { r = null; }   /* restores the saved folder handle */
  if (r && typeof r.then === 'function') r.then(cloudRefresh, function () { });
  setTimeout(cloudRefresh, 600);
}
function cloudRefresh() {
  if (VIEW.cur === 'set') renderCloudCard();
  if (SEND.ctx && !$('modal').hidden && $('s_cloud') && !SEND.ctx.busy) renderSendModal(SEND.ctx);
}
/* radio segment: OneDrive / Google Drive / iCloud / Bilmiyorum */
function cloudSeg(boxId, name) {
  var box = $(boxId);
  box.innerHTML = CLOUD_OPTS.map(function (o) {
    var on = S.cloud === o[0];
    return '<label' + (on ? ' class="on"' : '') + '><input type="radio" name="' + name + '" value="' + o[0] + '"' + (on ? ' checked' : '') + '>' + esc(o[1]) + '</label>';
  }).join('');
  if (!box.getAttribute('data-b')) {
    box.setAttribute('data-b', '1');
    box.addEventListener('change', function (e) {
      if (!e.target || e.target.name !== name) return;
      setCloud(e.target.value);
      [].forEach.call(box.querySelectorAll('label'), function (l) { l.classList.toggle('on', l.querySelector('input').checked); });
      if (box.getAttribute('data-guide') === '1') renderGuideBody();
    });
  }
}

/* ----- 3-step guide ----- */
function localGuide(pl, cloud) {
  var app = cloudLabel(cloud) || 'bulut uygulamanız';
  if (pl === 'ios') return [
    { title: 'Buluta gönder\'e basın', text: 'Paylaşma menüsü açılır. Gerekirse aşağı kaydırın.' },
    { title: 'Dosyalar\'a Kaydet\'i seçin', text: (cloud === 'icloud' || !cloud ? 'iCloud Drive' : app) + ' altındaki OMS Saha klasörünü açın.' },
    { title: 'Kaydet\'e basın', text: 'Dosya buluta gider; PC onu kendiliğinden işler.' }];
  if (pl === 'android') return [
    { title: 'Buluta gönder\'e basın', text: 'Paylaşma menüsü açılır.' },
    { title: (cloud ? app : 'Drive veya OneDrive') + ' uygulamasını seçin', text: 'Klasör olarak OMS Saha\'yı seçin.' },
    { title: 'Kaydet\'e (Yükle\'ye) basın', text: 'Dosya buluta gider; PC onu kendiliğinden işler.' }];
  return [
    { title: 'Buluta gönder\'e basın', text: 'İlk seferde OMS Saha klasörünü seçmeniz istenir.' },
    { title: 'OMS Saha klasörünü seçin', text: (cloud ? app + ' içindeki' : 'Bulut klasörünüzdeki') + ' OMS Saha klasörü. Tarayıcı izin isterse onaylayın.' },
    { title: 'Bitti', text: 'Sonraki gönderimler tek tıkla aynı klasöre gider. Klasör seçilemiyorsa dosya iner; onu OMS Saha klasörüne taşıyın.' }];
}
function guideSteps(pl, cloud) {
  var g = null;
  try { if (window.OmsCloud && typeof OmsCloud.guide === 'function') g = OmsCloud.guide(pl, cloudNorm(cloud)); } catch (e) { g = null; }
  if (Array.isArray(g)) g = g.filter(function (s) { return s && (s.title || s.text); });
  return Array.isArray(g) && g.length ? g : localGuide(pl, cloudNorm(cloud));
}
function guideList(pl) {
  var icons = ['i-cloudup', 'i-folder', 'i-check'], map = { send: 'i-cloudup', save: 'i-save', share: 'i-share', folder: 'i-folder', key: 'i-key', check: 'i-check', download: 'i-dl', move: 'i-folderopen' };
  return '<ol class="gsteps">' + guideSteps(pl, S.cloud).map(function (s, i) {
    return '<li><span class="gn" aria-hidden="true">' + (i + 1) + '</span><span class="gb"><b>' + ico(map[s.icon] || icons[Math.min(i, 2)]) + esc(s.title || '') + '</b>' +
      (s.text ? '<span class="gt">' + esc(s.text) + '</span>' : '') + (s.note ? '<span class="gt"><i>' + esc(s.note) + '</i></span>' : '') + '</span></li>';
  }).join('') + '</ol>';
}
var GUIDE = { pl: '', ctx: null };
function renderGuideBody() { var b = $('g_steps'); if (b) b.innerHTML = guideList(GUIDE.pl); }
/* ctx: opened from the send dialog on the first cloud send; its go button sends at once (a fresh tap = user activation) */
function openGuide(ctx, pl) {
  GUIDE.pl = pl || cloudPlat(); GUIDE.ctx = ctx || null;
  var dev = { ios: 'iPhone / iPad', android: 'Android telefon', desktop: 'Bilgisayar' }[GUIDE.pl];
  var h = '<p>' + esc(dev) + ' · 3 adım. Sepet dosyası bulutunuzdaki <b>OMS Saha</b> klasörüne gider; PC onu oradan alır.</p>' +
    '<div id="g_steps">' + guideList(GUIDE.pl) + '</div>' +
    '<div class="fld" style="margin-top:12px"><span class="lbl" id="gcloudl">Bulutunuz</span><div class="seg g2" role="radiogroup" aria-labelledby="gcloudl" id="gcloud" data-guide="1"></div></div>';
  if (ctx) h += '<div class="acts r"><button type="button" class="btn btn-sec" id="g_back">Geri</button><button type="button" class="btn btn-pri" id="g_go" autofocus>' + ico('i-cloudup') + 'Anladım, gönder</button></div>';
  else h += '<div class="acts r"><button type="button" class="btn btn-pri" id="g_ok" autofocus>Tamam</button></div>';
  openModal('Nasıl gönderilir?', h, ctx ? function () { if (SEND.ctx === ctx) SEND.ctx = null; } : null);
  cloudSeg('gcloud', 'gc');
  LS.set(K.guided, '1');
  if (ctx) {
    SEND.ctx = ctx;
    $('g_back').addEventListener('click', function () { renderSendModal(ctx); });
    $('g_go').addEventListener('click', function () { renderSendModal(ctx); runCloud(ctx); });
  } else $('g_ok').addEventListener('click', function () { closeModal(); });
}

/* ----- send dialog ----- */
function cloudBtn(ctx) {
  var how = cloudPrimary(), app = cloudLabel(S.cloud), fn = how === 'folder' && !ctx.needPick ? folderName() : '';
  var t, sub, icon = 'i-cloudup';
  if (how === 'folder' && fn) { t = 'OMS Saha klasörüne kaydet'; sub = 'Klasör: ' + fn + ' · tek tık'; icon = 'i-folder'; }
  else if (how === 'folder') { t = 'Klasörü seç ve gönder'; sub = 'İlk seferde ' + (app ? app + ' içindeki ' : '') + 'OMS Saha klasörünü seçin; sonra tek tıkla gider.'; icon = 'i-folderopen'; }
  else if (how === 'share') {
    t = 'Buluta gönder';
    var pl = cloudPlat();
    sub = pl === 'ios' ? 'Paylaşma menüsünde “Dosyalar\'a Kaydet” → ' + (app || 'bulutunuz') + ' → OMS Saha → Kaydet.' :
      pl === 'android' ? 'Paylaşma menüsünde ' + (app || 'Drive / OneDrive') + ' → OMS Saha klasörü → Kaydet.' :
      'Paylaşma menüsünde ' + (app || 'bulut uygulamanızı') + ' → OMS Saha klasörünü seçin.';
  }
  else { t = 'Buluta gönder'; sub = 'Dosya iner; onu ' + (app ? app + ' içindeki ' : 'bulut klasörünüzdeki ') + 'OMS Saha klasörüne taşıyın.'; }
  return '<button type="button" class="btn btn-pri cloudbtn" id="s_cloud">' + ico(icon) + '<span><b>' + esc(t) + '</b><small>' + esc(sub) + '</small></span></button>';
}
function renderSendModal(ctx) {
  var to = (S.pair && S.pair.to) || '', okTo = isEmail(to);
  var href = okTo ? 'mailto:' + to + '?subject=' + encodeURIComponent(subj()) + '&body=' + encodeURIComponent(ctx.body) : '';
  var tooLong = href.length > MAILTO_MAX;
  var acc = mailAccount(), direct = !!(acc && okTo), loginOpt = !acc && okTo && MAIL.st === 'ready';
  var h = '<p>' + (ctx.hist ? 'Bu sepet yeniden gönderilecek: ' : 'Sepet hazır: ') + fmtQty(ctx.n) + ' satır, şifreli ve imzalı.</p>';
  var lanTop = lanOn();
  if (lanTop) h += '<button type="button" class="btn btn-pri cloudbtn" id="s_lan">' + ico('i-lan') + '<span><b>Yerel ağdan gönder (anında)</b><small>' + esc(LANS.pc ? 'PC: ' + LANS.pc + ' · ' : '') + 'Sonuç hemen burada görünür.</small></span></button>';
  if (driveOn()) h += '<button type="button" class="btn ' + (lanTop ? 'btn-sec' : 'btn-pri') + ' cloudbtn" id="s_drvsend">' + ico('i-cloudup') + '<span><b>OneDrive\'a gönder</b><small>Tek dokunuş; PC birkaç dakika içinde işler.</small></span></button>';
  h += lanTop || driveOn() ? cloudBtn(ctx).replace('btn btn-pri cloudbtn', 'btn btn-sec cloudbtn') : cloudBtn(ctx);
  h += '<div class="howto"><button type="button" class="btn btn-ghost" id="s_how">' + ico('i-help') + 'Nasıl gönderilir?</button></div>';
  if (driveOk() && !driveOn() && !ctx.used) h += '<div class="note info" style="margin-top:8px">' + ico('i-cloudlink') + '<div><b>Daha kolayı:</b> OneDrive\'a bir kez bağlanın; sepetler paylaşma menüsü olmadan tek dokunuşla gider.' +
    '<div class="acts row-keep"><button type="button" class="btn btn-sec" id="s_drv">' + ico('i-cloudlink') + 'OneDrive\'a bağlan ve gönder</button></div><div id="s_drvmsg" class="msg" role="status" aria-live="polite"></div></div></div>';
  h += '<div id="s_res" class="sres"></div>';
  /* confirmation after share / download / mail: the page cannot see where the file went */
  var cq = ctx.pv === 'cloud' ? (ctx.dl ? 'Dosya indirildi (' + esc(ctx.dl) + '). Onu OMS Saha klasörüne taşıdıysanız Evet\'e basın' : 'Dosyayı OMS Saha klasörüne kaydettiyseniz Evet\'e basın')
    : 'E-postayı gönderdiyseniz Evet\'e basın';
  h += '<div id="s_conf" class="confirm-sent"' + (ctx.used ? '' : ' hidden') + '><p><b>Gönderdiniz mi?</b> ' + cq + '; sepet Geçmiş\'e taşınır' + (ctx.hist ? '' : ' ve yeni, boş bir sepet başlar') + '.</p>' +
    '<div class="acts row-keep"><button type="button" class="btn btn-sec" id="s_no">Hayır</button><button type="button" class="btn btn-pri" id="s_yes">' + ico('i-check') + 'Evet, gönderildi</button></div></div>';
  /* other ways: e-mail app, direct send (Microsoft account), copy, download */
  var o = '';
  if (!okTo) o += '<div class="fld" style="margin:4px 0 12px"><label for="s_to">E-posta için alıcı adresi</label><div class="row"><input id="s_to" type="email" inputmode="email" autocomplete="email" autocapitalize="off" autocorrect="off" spellcheck="false" enterkeyhint="done" value="' + esc(to) + '">' +
    '<button type="button" class="btn btn-sec" id="s_tosave">Tamam</button></div><p class="help">PC\'deki OMS programının okuduğu adres. E-posta yolları bu adresle açılır.</p></div>';
  if (direct) o += '<button type="button" class="btn btn-sec" id="s_direct">' + ico('i-send') + '<span>Doğrudan e-postayla gönder<small>' + esc(acc.username || acc.name || 'Microsoft hesabı') + ' hesabından; e-posta uygulaması açılmaz. İnternet yoksa giden kutusunda bekler.</small></span></button>' +
    '<div id="s_dmsg" class="msg" role="status" aria-live="polite"></div>';
  else if (loginOpt) o += '<button type="button" class="btn btn-sec" id="s_login">' + ico('i-key') + '<span>Microsoft ile giriş yap<small>Bir kez giriş yapınca sepetler bu sayfadan doğrudan e-postayla gönderilir.</small></span></button>' +
    '<div id="s_dmsg" class="msg" role="status" aria-live="polite"></div>';
  if (okTo && !tooLong) o += '<a class="btn btn-sec" id="s_mail" href="' + esc(href) + '">' + ico('i-mail') + '<span>E-posta uygulamasıyla gönder<small>Alıcı, konu ve metin hazır gelir; yalnız Gönder\'e basın.</small></span></a>';
  else if (okTo) o += '<div class="note warn">' + ico('i-warn') + '<div>Sepet e-posta bağlantısı için çok büyük. Kopyala ya da İndir ile gönderin.</div></div>';
  o += '<button type="button" class="btn btn-sec" id="s_copy">' + ico('i-copy') + '<span>E-posta metnini kopyala<small>Yeni e-postanın gövdesine yapıştırın.</small></span></button>';
  o += '<button type="button" class="btn btn-sec" id="s_dl">' + ico('i-dl') + '<span>Dosyayı indir (.omsb)<small>E-postaya ek olarak ekleyin ya da OMS Saha klasörüne kendiniz taşıyın.</small></span></button>';
  o += '<div class="mini"><button type="button" class="btn btn-ghost" id="s_cto"' + (okTo ? '' : ' disabled') + '>' + ico('i-copy') + 'Alıcıyı kopyala</button>' +
    '<button type="button" class="btn btn-ghost" id="s_csub">' + ico('i-copy') + 'Konuyu kopyala</button></div>';
  h += '<details class="more" id="s_more"' + (ctx.moreOpen ? ' open' : '') + '><summary>' + ico('i-more') + 'Diğer yollar (e-posta, kopyala, indir)' + ico('i-chev', 'chev') + '</summary><div class="inner sendopts">' + o + '</div></details>';
  h += '<div class="acts r" id="s_close"' + (ctx.used ? ' hidden' : '') + '><button type="button" class="btn btn-sec" id="s_cancel">Vazgeç</button></div>';
  openModal(ctx.hist ? 'Sepeti yeniden gönder' : 'Sepeti gönder', h, function () { if (SEND.ctx === ctx) SEND.ctx = null; });
  SEND.ctx = ctx; if (ctx.busy) busy($('s_cloud'), true);
  var sl = $('s_lan'); if (sl) sl.addEventListener('click', function () { lanSend(ctx.b, ctx.hist, sl, ctx); });
  var sdv = $('s_drvsend'); if (sdv) sdv.addEventListener('click', function () { closeModal(); driveSend(ctx.b, ctx.hist, null); });
  if (ctx.res) showRes(ctx, ctx.res);
  function used(pv) {
    ctx.used = true; ctx.pv = pv || ''; if (!pv) ctx.dl = '';
    renderSendModal(ctx);
    setTimeout(function () { try { $('s_conf').scrollIntoView({ block: 'nearest' }); } catch (e) { } }, 50);
  }
  $('s_cloud').addEventListener('click', function () {
    if (LS.get(K.guided) !== '1' && !DEMO) { openGuide(ctx); return; }   /* first time: show the 3 steps, "Anladım, gönder" sends */
    runCloud(ctx);
  });
  $('s_how').addEventListener('click', function () { openGuide(ctx); });
  var sd = $('s_drv');
  if (sd) sd.addEventListener('click', function () {
    /* phones sign in by full-page redirect: remember the intent so the basket goes out by itself after coming back */
    if (!ctx.hist) LS.set(K.drvSend, ctx.b && ctx.b.basket ? ctx.b.basket : '1');
    driveSignIn(sd, 's_drvmsg').then(function (ok) { if (ok && SEND.ctx === ctx) { LS.del(K.drvSend); closeModal(); driveSend(ctx.b, ctx.hist, null); } });
  });
  $('s_more').addEventListener('toggle', function () { ctx.moreOpen = this.open; });
  var m = $('s_mail'); if (m) m.addEventListener('click', function () { setTimeout(function () { used(''); }, 400); });
  var dv = $('s_direct'); if (dv) dv.addEventListener('click', function () { directSend(ctx, to, dv); });
  var lg = $('s_login'); if (lg) lg.addEventListener('click', function () { mailSignIn(lg, 's_dmsg').then(function (ok) { if (ok && !$('modal').hidden) renderSendModal(ctx); }); });
  $('s_copy').addEventListener('click', function () { copyText(ctx.body, 'E-posta metni kopyalandı. Alıcı ve konuyu da kopyalayın.'); used(''); });
  $('s_dl').addEventListener('click', function () { download(ctx.name, ctx.armor); toast('Dosya indirildi: ' + ctx.name); used(''); });
  $('s_cto').addEventListener('click', function () { copyText(to, 'Alıcı kopyalandı'); });
  $('s_csub').addEventListener('click', function () { copyText(subj(), 'Konu kopyalandı'); });
  $('s_cancel').addEventListener('click', function () { closeModal(); });
  $('s_no').addEventListener('click', function () { closeModal(); toast('Sepet yerinde kaldı; istediğiniz zaman yeniden gönderin.'); });
  $('s_yes').addEventListener('click', function () { markSent(ctx, ctx.pv === 'cloud' ? 'cloud' : ''); });
  var ti = $('s_to');
  if (ti) {
    var saveTo = function () {
      var v = ti.value.trim();
      if (!isEmail(v)) { ti.setAttribute('aria-invalid', 'true'); ti.focus(); toast('Geçerli bir e-posta adresi yazın.', true); return; }
      S.pair.to = v; savePair(); ctx.moreOpen = true; renderSendModal(ctx);
    };
    $('s_tosave').addEventListener('click', saveTo);
    ti.addEventListener('keydown', function (e) { if (e.key === 'Enter') { e.preventDefault(); saveTo(); } });
    ti.addEventListener('input', function () { ti.removeAttribute('aria-invalid'); });
  }
}
/* result line under the big button: {k:'err'|'warn'|'info', t, act:'pick'|'force'} */
function showRes(ctx, r) {
  ctx.res = r;
  var box = $('s_res'); if (!box || SEND.ctx !== ctx) return;
  if (!r) { box.innerHTML = ''; return; }
  var h = '<div class="note ' + (r.k || 'info') + '" role="' + (r.k === 'err' ? 'alert' : 'status') + '">' + ico(r.k === 'err' ? 'i-err' : r.k === 'warn' ? 'i-warn' : 'i-info') + '<div>' + esc(r.t);
  if (r.act === 'force') h += '<div class="acts row-keep"><button type="button" class="btn btn-sec" id="s_repick">' + ico('i-folderopen') + 'Başka klasör seç</button><button type="button" class="btn btn-pri" id="s_force">Yine de kaydet</button></div>';
  h += '</div></div>';
  box.innerHTML = h;
  var f = $('s_force'); if (f) f.addEventListener('click', function () { ctx.res = null; runCloud(ctx, 'force'); });
  var rp = $('s_repick'); if (rp) rp.addEventListener('click', function () { ctx.res = null; ctx.needPick = true; runCloud(ctx); });
}
/* the big button. Called straight from a click: OmsCloud.send / pickFolder need the user activation of that tap */
function runCloud(ctx, mode) {
  if (ctx.busy) return;
  var how = cloudPrimary(), btn = $('s_cloud');
  showRes(ctx, null);
  if (how === 'local') {   /* cloudsend.js missing: plain download, the user moves the file */
    download(ctx.name, ctx.armor); ctx.dl = ctx.name; ctx.used = true; ctx.pv = 'cloud'; renderSendModal(ctx); return;
  }
  var opt = { armor: ctx.armor, date: ctx.date, prefer: how, fileName: ctx.name, cloud: S.cloud };
  var p;
  /* called synchronously inside the tap: WebKit / Chromium only allow share / pickers with a live user activation */
  function call(f) { try { return Promise.resolve(f()); } catch (e) { return Promise.reject(e); } }
  if (how === 'folder' && mode !== 'force' && (ctx.needPick || !folderName())) {
    p = call(function () { return OmsCloud.pickFolder(); }).then(function (r) {
      if (!r || !r.ok) return { ok: false, method: 'folder', cancelled: !r || !!r.cancelled || r.code === 'CANCELLED', error: r && r.error, code: (r && r.code) || 'PICK' };
      ctx.needPick = false; if (VIEW.cur === 'set') renderCloudCard();
      if (r.looksRight === false || (r.looksRight == null && !isSahaName(r.name))) return { ok: false, method: 'folder', code: 'NOT_SAHA', name: r.name };
      return OmsCloud.send(opt);
    });
  } else p = call(function () { return OmsCloud.send(opt); });
  ctx.busy = true; busy(btn, true);
  p.then(function (r) { cloudDone(ctx, r || { ok: false, code: 'ERROR' }); }, function (e) { cloudDone(ctx, { ok: false, code: 'ERROR', error: (e && (e.message || e.name)) || String(e) }); });
}
function cloudDone(ctx, r) {
  ctx.busy = false;
  var open = SEND.ctx === ctx && !$('modal').hidden;
  if (open) busy($('s_cloud'), false);
  if (r.ok && r.method === 'folder') { markSent(ctx, 'cloud'); return; }   /* written into the folder: nothing to confirm */
  if (r.ok) {
    ctx.used = true; ctx.pv = 'cloud'; ctx.dl = r.method === 'download' ? (r.name || ctx.name) : '';
    if (open) { renderSendModal(ctx); setTimeout(function () { try { $('s_conf').scrollIntoView({ block: 'nearest' }); } catch (e) { } }, 50); }
    return;
  }
  if (!open) return;
  var why = r.error || '';
  if (r.cancelled || r.code === 'CANCELLED') { showRes(ctx, { k: 'info', t: 'Gönderilmedi: pencere kapatıldı. Sepet yerinde; yeniden deneyebilirsiniz.' }); return; }
  if (r.code === 'NOT_SAHA') { showRes(ctx, { k: 'warn', t: 'Seçilen klasörün adı "' + (r.name || '?') + '", "OMS Saha" değil. Dosya PC\'ye ulaşmayabilir.', act: 'force' }); return; }
  if (r.code === 'NEED_TAP' || r.code === 'BUSY') { showRes(ctx, { k: 'info', t: why || 'Düğmeye yeniden basın.' }); return; }
  if (r.code === 'NEED_FOLDER' || r.code === 'FOLDER_GONE' || r.code === 'NO_PERMISSION') {
    ctx.needPick = r.code !== 'NO_PERMISSION'; renderSendModal(ctx);
    showRes(ctx, { k: 'warn', t: why || 'OMS Saha klasörüne erişilemedi. Büyük düğmeyle klasörü yeniden seçin.' });
    if (VIEW.cur === 'set') renderCloudCard();
    return;
  }
  ctx.moreOpen = true; renderSendModal(ctx);
  showRes(ctx, { k: 'err', t: (why || 'Gönderilemedi.') + ' Olmazsa aşağıdaki "Diğer yollar"ı kullanın.' });
}

/* ----- settings: cloud card ----- */
function renderCloudCard() {
  renderDrive();
  cloudSeg('scloud', 'sc');
  var fo = folderOk(); $('c_folder').hidden = !fo;
  if (!fo) return;
  var n = folderName();
  $('fldname').textContent = n ? n + (isSahaName(n) ? '' : ' (adı "OMS Saha" değil)') : 'Seçilmedi: ilk gönderimde sorulur.';
  $('bfldpick').innerHTML = ico('i-folderopen') + (n ? 'Değiştir' : 'Klasörü seç');
  $('bfldforget').hidden = !n;
}
$('bfldpick').addEventListener('click', function () {
  if (!folderOk()) return;
  var btn = this; busy(btn, true); setMsg('fldmsg', '');
  var pp; try { pp = Promise.resolve(OmsCloud.pickFolder()); } catch (e) { pp = Promise.reject(e); }   /* picker needs the live tap */
  pp.then(function (r) {
    busy(btn, false);
    var good = r && (r.looksRight != null ? r.looksRight : isSahaName(r.name));
    if (r && r.ok) setMsg('fldmsg', good ? 'Seçildi: ' + r.name + '. Gönderimler buraya gider.' : 'Seçildi: ' + r.name + '. Bu klasörün adı "OMS Saha" değil; dosyalar PC\'ye ulaşmayabilir.', good ? 'ok' : 'err');
    else if (r && !r.cancelled && r.code !== 'CANCELLED') setMsg('fldmsg', r.error || 'Klasör seçilemedi.', 'err');
    renderCloudCard();
  }, function (e) { busy(btn, false); setMsg('fldmsg', 'Klasör seçilemedi: ' + ((e && e.message) || e), 'err'); });
});
$('bfldforget').addEventListener('click', function () {
  Promise.resolve().then(function () { return OmsCloud.forgetFolder(); }).then(function () { setMsg('fldmsg', 'Klasör unutuldu. Sonraki gönderimde yeniden sorulur.', 'ok'); renderCloudCard(); }, function () { renderCloudCard(); });
});
$('bhowto').addEventListener('click', function () { openGuide(null); });
/* via: '' (mail app / copy / download, user confirmed), 'cloud' (cloud folder: written, or shared/downloaded + user confirmed),
   'direct' (Graph 202), 'drive' (OneDrive PUT 200/201), 'queued' (in an outbox, not sent yet; ch 'mail' = OmsMail, 'drive' = OmsDrive,
   st = 'queued' | 'login') */
function markSent(ctx, via, err, ch, st) {
  var now = Date.now(), x;
  if (ctx.hist) {
    x = ctx.b;
    var i = S.hist.indexOf(x); if (i > 0) { S.hist.splice(i, 1); S.hist.unshift(x); }
  } else {
    x = { basket: ctx.b.basket, created: ctx.b.created, lines: ctx.b.lines, devName: S.pair ? S.pair.name : '', user: S.user, sends: [] };
    S.hist.unshift(x);
    if (S.basket === ctx.b) S.basket = null;
    saveBasket(true);
  }
  if (via === 'queued') {
    if (x.out && x.ch !== (ch || 'mail')) { if (x.ch === 'drive') driveRemove(x.basket); else outboxRemove(x.basket); }
    x.out = st || 'queued'; x.ch = ch || 'mail'; x.q = now; x.outErr = err || '';
  } else {
    x.sends.push(now); x.via = via || 'mail';
    if (via === 'lan' && ctx.lanRes) x.lanRes = ctx.lanRes; else delete x.lanRes;
    if (x.out) { if (x.ch === 'drive') { if (via !== 'drive') driveRemove(x.basket); } else outboxRemove(x.basket); delete x.out; delete x.outErr; delete x.ch; }
  }
  pruneHist(); saveHist();
  if (SEND.ctx === ctx) closeModal();
  render(VIEW.cur);
  var go = { t: 'Geçmiş', fn: function () { showView('hist'); } };
  if (via === 'lan') return;   /* lanDone shows the PC's answer */
  if (via === 'queued' && st === 'login') toast('Sepet giden kutusunda. OneDrive\'a yeniden bağlanın; sonra kendiliğinden gider.', true, { t: 'Bağlan', fn: function () { driveSignIn(null).then(function (ok) { if (ok) driveFlush(null, true); }); } });
  else if (via === 'queued') toast((ch === 'drive' && err ? err + ' ' : '') + 'Sepet giden kutusunda: internet gelince kendiliğinden gider.', false, go);
  else if (via === 'drive') toast(ctx.hist ? 'OneDrive\'a yeniden gönderildi ✓' : 'OneDrive\'a gönderildi ✓ PC birkaç dakika içinde işler.', false, ctx.hist ? null : go);
  else if (via === 'cloud') toast(ctx.hist ? 'Buluta yeniden gönderildi ✓' : 'Buluta gönderildi ✓ PC birkaç dakika içinde işler.', false, ctx.hist ? null : go);
  else if (via === 'direct') toast('Gönderildi ✓ PC birkaç dakika içinde işler.', false, ctx.hist ? null : go);
  else toast(ctx.hist ? 'Yeniden gönderim kaydedildi' : 'Sepet Geçmiş\'e taşındı', false, ctx.hist ? null : go);
}

/* ===================== V10: direct send over the local network (OmsLan, lansend.js; spec section 20) ===================== */
/* st: 'off' (no module / pairing / addresses) | 'looking' | 'found' | 'none'. cert = "{root}/saha/" when the browser probably has
   not accepted the PC's own certificate yet (asked once; a root that ever answered is never asked again). */
var LANS = { st: 'off', root: '', pc: '', cert: '', certHidden: false, busy: false, p: null, at: 0 };
function lanOk() { return !!(window.OmsLan && typeof OmsLan.find === 'function' && typeof OmsLan.send === 'function'); }
function lanOn() { return lanOk() && LANS.st === 'found' && !!LANS.root; }
function lanSeen() { var o = jget(K.lanOk, {}); return o && typeof o === 'object' ? o : {}; }
function lanUi() {
  if (VIEW.cur === 'basket') renderBasket();
  var c = SEND.ctx, on = lanOn();
  if (c && !$('modal').hidden && $('s_cloud') && !c.busy && !c.lanBusy && on !== !!$('s_lan')) renderSendModal(c);
}
/* background search (Sepet opened, page visible again, back online). force = skip the 60 s cache: the network may have changed */
function lanProbe(force) {
  if (!lanOk() || !S.pair || !hasLan() || DEMO === 'selftest') {
    if (LANS.st !== 'off') { LANS.st = 'off'; LANS.root = ''; LANS.pc = ''; LANS.cert = ''; lanUi(); }
    return Promise.resolve(null);
  }
  if (LANS.p) return LANS.p;
  if (LANS.st !== 'found') LANS.st = 'looking';
  LANS.p = Promise.resolve().then(function () { return OmsLan.find({ pairing: S.pair, force: !!force }); }).then(function (r) {
    r = r || {};
    LANS.p = null; LANS.at = Date.now();
    if (r.ok) {
      LANS.st = 'found'; LANS.root = r.root; LANS.pc = r.pc || ''; LANS.cert = '';
      var o = lanSeen(); o[String(r.root).toLowerCase()] = Date.now(); LS.set(K.lanOk, JSON.stringify(o));
    } else {
      LANS.st = 'none'; LANS.root = ''; LANS.pc = '';
      var cu = r.needCertUrl || '';
      if (cu && lanSeen()[cu.replace(/\/saha\/$/, '').toLowerCase()]) cu = '';   /* accepted before: just not on this network now */
      LANS.cert = cu;
    }
    lanUi();
    return r;
  }, function () { LANS.p = null; LANS.at = Date.now(); LANS.st = 'none'; LANS.root = ''; lanUi(); return null; });
  return LANS.p;
}
function lanKick() { if (!LANS.p && !LANS.busy && Date.now() - LANS.at > 10000) lanProbe(false); }
/* the PC reports its current addresses with every answer (DHCP may move it): keep the pairing up to date */
function lanLearn(i) {
  if (!S.pair || !i || !Array.isArray(i.lan) || !i.lan.length) return;
  var old = S.pair.lan || [], m = OmsLan.mergeLan(old, i.lan);
  if (m.join(' ') === old.join(' ')) return;
  S.pair.lan = m; savePair();
  if (VIEW.cur === 'set' && !LANED) renderLan();
}
if (lanOk()) { try { OmsLan.onLan(lanLearn); } catch (e) { } }
function lanSummary(r) {
  if (!r) return '';
  if (r.outcome === 'Duplicate') return 'Bu sepet zaten işlenmişti';
  return 'PC aldı: ' + fmtQty(r.applied || 0) + ' satır stoğa işlendi' + (r.issues ? ' · ' + fmtQty(r.issues) + ' durum' : '') + (r.skipped ? ' · ' + fmtQty(r.skipped) + ' atlandı' : '');
}
/* small status line at the top of Sepet: PC found, or the one-time certificate step */
function renderLanLine() {
  var el = $('blan'); if (!el) return;
  var h = '';
  if (lanOn()) h = '<div class="lanok">' + ico('i-lan') + '<span>Yerel ağda PC bulundu: <b>' + esc(LANS.pc || 'PC') + '</b></span></div>';
  else if (LANS.cert && !LANS.certHidden && Date.now() - (+LS.get(K.lanCertHide) || 0) > 86400000)
    h = '<div class="note info">' + ico('i-shield') + '<div><b>PC aynı ağda olabilir.</b> Tarayıcının PC\'ye bir kez bağlanma izni vermesi gerekiyor.' +
      '<div class="acts row-keep"><a class="btn btn-pri" id="blancert" href="' + esc(LANS.cert) + '" target="_blank" rel="noopener">' + ico('i-shield') + 'PC\'ye bağlanmayı onayla</a>' +
      '<button type="button" class="btn btn-sec" id="blanlater">Şimdi değil</button></div>' +
      '<p class="small" style="margin:8px 0 0">Açılan sayfada <b>Gelişmiş</b> → <b>Devam et</b> deyin, sonra bu sayfaya dönün.</p></div></div>';
  el.innerHTML = h; el.hidden = !h;
}
$('blan').addEventListener('click', function (e) {
  if (e.target.closest('#blancert')) { LANS.certHidden = true; setTimeout(renderLanLine, 300); return; }   /* back on this tab: visibilitychange searches again */
  if (e.target.closest('#blanlater')) { LS.set(K.lanCertHide, String(Date.now())); renderLanLine(); }
});
document.addEventListener('visibilitychange', function () { if (!document.hidden && VIEW.cur === 'basket') lanProbe(true); });
window.addEventListener('online', function () { if (VIEW.cur === 'basket') lanProbe(true); });
/* seal exactly like the other paths (sealCtx) -> POST to the PC -> the PC's real answer */
async function lanSend(b, fromHist, btn, ctx0) {
  if (LANS.busy) return;
  if (!S.pair) { toast('Önce cihazı eşleştirin.', true); showView('pair'); return; }
  var er = envReady(); if (er) { toast(er, true); return; }
  var root = LANS.root;
  if (!root) { startSend(b, fromHist); return; }
  LANS.busy = true; busy(btn, true);
  var main = btn && btn.id === 'bsend';
  if (main) $('bsendtx').textContent = 'Gönderiliyor…';
  var ctx = ctx0 || null, r = null;
  if (ctx) ctx.lanBusy = true;
  try {
    if (!ctx) ctx = await sealCtx(b, fromHist);
    r = await OmsLan.send({ root: root, armor: ctx.armor });
  } catch (e) { r = { ok: false, code: ctx ? 'ERROR' : 'SEAL', error: (e && e.message) || String(e) }; }
  LANS.busy = false; if (ctx) ctx.lanBusy = false; busy(btn, false);
  r = r || { ok: false, code: 'ERROR' };
  if (main) renderBasket();
  if (r.code === 'SEAL') { toast('Sepet hazırlanamadı: ' + r.error, true); return; }
  lanDone(ctx, r);
}
function lanDone(ctx, r) {
  if (r.ok) {
    FB.go('ok');
    ctx.lanRes = { outcome: r.outcome, applied: r.applied || 0, skipped: r.skipped || 0, issues: r.issues || 0, pc: r.pc || LANS.pc || '', at: Date.now() };
    markSent(ctx, 'lan');
    var dup = r.outcome === 'Duplicate';
    var h = '<div class="note ' + (dup ? 'info' : 'ok') + ' lanres" role="status">' + ico(dup ? 'i-info' : 'i-check') + '<div><b>' + esc(lanSummary(ctx.lanRes)) + '</b>' +
      (dup ? '<br>PC aynı sepeti ikinci kez stoğa eklemedi.' : '') +
      (!dup && r.issues ? '<br>Durumlu satırları PC\'de Saha sekmesinden çözün.' : '') + '</div></div>' +
      '<p class="small" style="margin:12px 0 0">' + esc(ctx.lanRes.pc ? 'PC: ' + ctx.lanRes.pc + ' · ' : '') + 'Sepet Geçmiş\'e taşındı' + (ctx.hist ? '.' : '; yeni, boş bir sepet başladı.') + '</p>' +
      '<div class="acts r"><button type="button" class="btn btn-sec" id="l_hist">' + ico('i-hist') + 'Geçmiş</button><button type="button" class="btn btn-pri" id="l_ok" autofocus>Tamam</button></div>';
    openModal(dup ? 'Zaten işlenmiş' : 'PC aldı ✓', h, null);
    $('l_ok').addEventListener('click', function () { closeModal(); });
    $('l_hist').addEventListener('click', function () { closeModal(); showView('hist'); });
    return;
  }
  FB.go('err');
  if (r.code === 'OFFLINE' || r.code === 'NEED_CERT') {
    LANS.st = 'none'; LANS.root = ''; LANS.pc = ''; LANS.at = Date.now();
    if (r.needCertUrl) { LANS.cert = r.needCertUrl; LANS.certHidden = false; }
  }
  var t = r.code === 'REJECTED' ? 'PC sepeti kabul etmedi: ' + (r.error || '') : (r.error || 'Gönderilemedi.');
  if (!/[.!?]$/.test(t)) t += '.';
  var h2 = '<div class="note ' + (r.code === 'RETRY' ? 'warn' : 'err') + '" role="alert">' + ico(r.code === 'RETRY' ? 'i-warn' : 'i-err') + '<div>' + esc(t) +
    '<br>Sepet yerinde duruyor.</div></div>';
  if (r.code === 'NEED_CERT' && r.needCertUrl) h2 += '<div class="acts"><a class="btn btn-pri" id="f_cert" href="' + esc(r.needCertUrl) + '" target="_blank" rel="noopener">' + ico('i-shield') + 'PC\'ye bağlanmayı onayla</a></div>' +
    '<p class="small" style="margin:8px 0 0">Açılan sayfada <b>Gelişmiş</b> → <b>Devam et</b> deyin, sonra bu sayfaya dönüp yeniden gönderin.</p>';
  h2 += '<div class="acts r">';
  if (r.code === 'RETRY' && lanOn()) h2 += '<button type="button" class="btn btn-sec" id="f_again">' + ico('i-resync') + 'Yeniden dene</button>';
  if (driveOn()) h2 += '<button type="button" class="btn btn-pri" id="f_drive">' + ico('i-cloudup') + 'OneDrive ile gönder</button>';
  h2 += '<button type="button" class="btn ' + (driveOn() ? 'btn-sec' : 'btn-pri') + '" id="f_other">' + ico('i-send') + 'Başka yolla gönder</button>';
  h2 += '<button type="button" class="btn btn-sec" id="f_close">Kapat</button></div>';
  openModal('Yerel ağdan gönderilemedi', h2, null);
  var a = $('f_again'); if (a) a.addEventListener('click', function () { closeModal(); lanSend(ctx.b, ctx.hist, null, ctx); });
  var d = $('f_drive'); if (d) d.addEventListener('click', function () { closeModal(); driveSend(ctx.b, ctx.hist, null); });
  var c = $('f_cert'); if (c) c.addEventListener('click', function () { LANS.certHidden = true; });
  $('f_other').addEventListener('click', function () { renderSendModal(ctx); });
  $('f_close').addEventListener('click', function () { closeModal(); });
  if (VIEW.cur === 'basket') renderBasket();
}

/* ===================== barcode decoding ===================== */
var FMT_ALL = ['DATA_MATRIX', 'QR_CODE', 'CODE_128', 'CODE_39', 'PDF_417'], FMT_2D = ['DATA_MATRIX', 'QR_CODE', 'PDF_417'];
var NATIVE_MAP = { data_matrix: 'DATA_MATRIX', qr_code: 'QR_CODE', code_128: 'CODE_128', code_39: 'CODE_39', code_93: 'CODE_93', pdf417: 'PDF_417', aztec: 'AZTEC',
  ean_13: 'EAN_13', ean_8: 'EAN_8', upc_a: 'UPC_A', upc_e: 'UPC_E', itf: 'ITF', codabar: 'CODABAR' };
function zxUrl() { return new URL('./zxing.min.js', location.href).href; }
/* worker body: luminance in, {text, format} out. ZXing source is prepended to the blob, so no importScripts / network inside the worker */
var WSRC = 'var R={};function mk(l){var h=new Map();h.set(ZXing.DecodeHintType.TRY_HARDER,true);h.set(ZXing.DecodeHintType.POSSIBLE_FORMATS,l.map(function(n){return ZXing.BarcodeFormat[n]}));var r=new ZXing.MultiFormatReader();r.setHints(h);return r}\n' +
  'self.onmessage=function(e){var d=e.data;if(d.t==="init"){try{R.all=mk(d.all);R.d2=mk(d.d2);R.qr=mk(["QR_CODE"]);self.postMessage({t:"ready"})}catch(x){self.postMessage({t:"fail",m:String(x&&x.message||x)})}return}\n' +
  'if(d.t==="dec"){var o={t:"res",id:d.id};try{var r=R[d.set]||R.all;var src=new ZXing.RGBLuminanceSource(new Uint8ClampedArray(d.buf),d.w,d.h);var bb=new ZXing.BinaryBitmap(new ZXing.HybridBinarizer(src));var res;try{res=r.decodeWithState(bb)}finally{r.reset()}o.text=res.getText();o.format=ZXing.BarcodeFormat[res.getBarcodeFormat()]}catch(x){}self.postMessage(o)}};';
var DEC = {
  kind: null, det: null, worker: null, seq: 0, pend: {}, mainR: null, nerr: 0, initP: null,
  init: function () { if (!this.initP) this.initP = this._init(); return this.initP; },
  _init: async function () {
    if ('BarcodeDetector' in window) {
      try {
        var f = await window.BarcodeDetector.getSupportedFormats();
        if (f.indexOf('data_matrix') >= 0 && f.indexOf('qr_code') >= 0) {
          var want = f.filter(function (x) { return ['data_matrix', 'qr_code', 'code_128', 'code_39', 'pdf417'].indexOf(x) >= 0; });
          this.det = new window.BarcodeDetector({ formats: want }); this.kind = 'native'; return this.kind;
        }
      } catch (e) { }
    }
    return this.zx();
  },
  /* ZXing fallback: blob worker (keeps the camera preview smooth), else main thread */
  zx: async function () {
    try { await this.startWorker(); this.kind = 'worker'; return this.kind; } catch (e) { }
    await loadZxing(); this.mainR = mkReaders(); this.kind = 'main'; return this.kind;
  },
  startWorker: function () {
    var self = this;
    return fetch(zxUrl()).then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.text(); }).then(function (src) {
      return new Promise(function (res, rej) {
        var u = URL.createObjectURL(new Blob([src, '\n;\n', WSRC], { type: 'text/javascript' })), w;
        try { w = new Worker(u); } catch (e) { URL.revokeObjectURL(u); rej(e); return; }
        var to = setTimeout(function () { try { w.terminate(); } catch (e) { } rej(new Error('worker timeout')); }, 15000);
        w.onmessage = function (e) {
          var d = e.data;
          if (d.t === 'ready') { clearTimeout(to); URL.revokeObjectURL(u); self.worker = w; w.onmessage = function (ev) { self.onRes(ev.data); }; res(); }
          else if (d.t === 'fail') { clearTimeout(to); URL.revokeObjectURL(u); try { w.terminate(); } catch (x) { } rej(new Error(d.m)); }
        };
        w.onerror = function (e) { clearTimeout(to); URL.revokeObjectURL(u); try { w.terminate(); } catch (x) { } rej(new Error((e && e.message) || 'worker error')); };
        w.postMessage({ t: 'init', all: FMT_ALL, d2: FMT_2D });
      });
    });
  },
  onRes: function (d) { var p = this.pend[d.id]; if (!p) return; delete this.pend[d.id]; clearTimeout(p.to); p.res(d.text ? { text: d.text, format: d.format } : null); },
  /* luminance (Uint8ClampedArray w*h) -> {text, format} | null */
  decodeLum: function (lum, w, h, set) {
    var self = this;
    if (this.kind === 'worker' && this.worker) {
      return new Promise(function (res) {
        var id = ++self.seq;
        self.pend[id] = { res: res, to: setTimeout(function () { if (self.pend[id]) { delete self.pend[id]; res(null); } }, 6000) };
        self.worker.postMessage({ t: 'dec', id: id, buf: lum.buffer, w: w, h: h, set: set }, [lum.buffer]);
      });
    }
    var r = (this.mainR || {})[set] || (this.mainR || {}).all;
    if (!r) return Promise.resolve(null);
    try {
      var bb = new ZXing.BinaryBitmap(new ZXing.HybridBinarizer(new ZXing.RGBLuminanceSource(lum, w, h))), res;
      try { res = r.decodeWithState(bb); } finally { r.reset(); }
      return Promise.resolve({ text: res.getText(), format: ZXing.BarcodeFormat[res.getBarcodeFormat()] });
    } catch (e) { return Promise.resolve(null); }
  },
  /* native detector -> [{text, format}] */
  detect: function (src) {
    var self = this;
    return this.det.detect(src).then(function (list) {
      self.nerr = 0;
      return (list || []).map(function (b) { return { text: b.rawValue, format: NATIVE_MAP[b.format] || String(b.format || '').toUpperCase() }; });
    }, function (e) {
      /* e.g. Android without Play services: the API exists but the service does not -> switch to ZXing */
      if (++self.nerr >= 3) { self.kind = null; self.det = null; self.initP = self.zx(); }
      throw e;
    });
  }
};
var ZX = { p: null };
function loadZxing() {
  if (window.ZXing) return Promise.resolve();
  if (ZX.p) return ZX.p;
  ZX.p = new Promise(function (res, rej) {
    var s = document.createElement('script'); s.src = './zxing.min.js';
    s.onload = function () { window.ZXing ? res() : rej(new Error('zxing yok')); };
    s.onerror = function () { ZX.p = null; rej(new Error('Kod çözücü (zxing.min.js) yüklenemedi')); };
    document.head.appendChild(s);
  });
  return ZX.p;
}
function mkReaders() {
  function mk(l) { var h = new Map(); h.set(ZXing.DecodeHintType.TRY_HARDER, true); h.set(ZXing.DecodeHintType.POSSIBLE_FORMATS, l.map(function (n) { return ZXing.BarcodeFormat[n]; })); var r = new ZXing.MultiFormatReader(); r.setHints(h); return r; }
  return { all: mk(FMT_ALL), d2: mk(FMT_2D), qr: mk(['QR_CODE']) };
}
function toLum(data, n) {
  var lum = new Uint8ClampedArray(n);
  for (var i = 0, j = 0; i < n; i++, j += 4) lum[i] = (data[j] * 77 + data[j + 1] * 150 + data[j + 2] * 29) >> 8;
  return lum;
}
function fmtName(f) {
  var m = { DATA_MATRIX: 'DataMatrix', QR_CODE: 'QR', PDF_417: 'PDF417', AZTEC: 'Aztec', CODE_128: 'Code 128', CODE_39: 'Code 39', CODE_93: 'Code 93',
    EAN_13: 'EAN-13', EAN_8: 'EAN-8', UPC_A: 'UPC-A', UPC_E: 'UPC-E', ITF: 'ITF', CODABAR: 'Codabar' };
  return m[f] || String(f || 'kod');
}
function is2D(f) { return /MATRIX|QR|PDF|AZTEC|MAXI/.test(String(f || '')); }
/* reel labels carry one 2D code (full ISO 15434 record) and several 1D codes (one field each): prefer the 2D one */
function pickPrimary(hits) {
  var best = null, bs = -1;
  hits.forEach(function (h) { var t = String(h.text || ''), s = t.length + (is2D(h.format) ? 1000 : 0) + (/\x1D|^\[\)>|^\{/.test(t) ? 500 : 0); if (t && s > bs) { bs = s; best = h; } });
  return best;
}

/* ===================== camera ===================== */
var CAM = {
  stream: null, track: null, running: false, paused: false, inflight: false, timer: null, want: false, mode: 'scan',
  canvas: null, ctx: null, k: 0, hold: null, last: { raw: '', t: 0 }, warnT: 0, torch: false, idle: { t0: 0, hint: 0 }, wl: null, vids: [], vcur: '',
  secure: function () { return window.isSecureContext === true || location.hostname === 'localhost' || location.hostname === '127.0.0.1'; },
  supported: function () { return !!(navigator.mediaDevices && navigator.mediaDevices.getUserMedia); },
  msgId: function () { return this.mode === 'pair' ? 'pcammsg' : this.mode === 'cfg' ? 'cfgmsg' : 'cammsg'; },
  btnId: function () { return this.mode === 'pair' ? 'bpairqr' : this.mode === 'cfg' ? 'bcfgcam' : 'bcam'; },
  state: function (txt, kind) {
    var p = $('cstate'), s = p.querySelector('span'); if (s.textContent !== txt) s.textContent = txt;
    var c = 'spill' + (kind ? ' ' + kind : ''); if (p.className !== c) p.className = c;
    var d = p.firstElementChild, want = kind === 'ok' ? 'i-check' : kind === 'warn' ? 'i-warn' : kind === 'err' ? 'i-err' : '';
    if (want) { var n = document.createElement('span'); n.innerHTML = ico(want); p.replaceChild(n.firstChild, d); }
    else if (d.tagName !== 'I') { var i = document.createElement('i'); i.className = 'dot'; p.replaceChild(i, d); }
  },
  flash: function (cls) { var st = $('stage'); st.classList.remove('hit', 'dup'); void st.offsetWidth; st.classList.add(cls); clearTimeout(this.flT); this.flT = setTimeout(function () { st.classList.remove('hit', 'dup'); }, 450); },
  stageOn: function (on) {
    var st = $('stage');
    var cfg = this.mode === 'cfg';
    if (on) { $(this.mode === 'pair' ? 'pairstage' : cfg ? 'camstage' : 'scanstage').appendChild(st); st.classList.toggle('qr', this.mode === 'pair'); st.classList.toggle('cfg', cfg); }
    st.hidden = !on; document.body.classList.toggle('scanning', on && !cfg); document.body.classList.toggle('camcfg', on && cfg);
    $('bstagebasket').hidden = this.mode === 'pair';
    var ser = on && this.serial;
    st.classList.toggle('serial', ser); $('sercnt').hidden = !ser; $('bdone').hidden = !ser; $('bmanual').hidden = ser; if (!ser) $('serlist').hidden = true;
    $('cfghint').hidden = !(on && cfg && hasCam());
    $('bmanual').querySelector('span').textContent = this.mode === 'pair' ? 'Kodu yapıştır' : 'Elle';
    if (!on) { $('cpause').hidden = true; $('chelp').hidden = true; PLUS.hide(); }
  },
  start: async function (mode) {
    this.mode = mode || 'scan';
    var mid = this.msgId();
    if (!this.secure()) { setMsg(mid, 'Kamera yalnız güvenli (https) adreste açılır. Sayfayı https:// adresinden açın ya da kodu elle girin.', 'err'); return; }
    if (!this.supported()) { setMsg(mid, 'Bu tarayıcı sayfaya kamera açmıyor (uygulama içi tarayıcı olabilir). Chrome, Safari, Edge veya Firefox ile açın' + (this.mode === 'pair' ? ' ya da kodu yapıştırın.' : ' ya da "Fotoğraftan oku"yu kullanın.'), 'err'); return; }
    if (this.stream) { this.resume(); return; }
    FB.unlock(); hideToast();
    var btn = $(this.btnId()); busy(btn, true); setMsg(mid, 'Kamera açılıyor…');
    this.serial = this.mode === 'scan' && smode() === 'serial';
    var decP = this.mode === 'cfg' ? Promise.resolve('none') : DEC.init().catch(function (e) { return e; });
    try { await this.open(LS.get(K.cam)); }
    catch (e) {
      busy(btn, false); this.want = false; var n = e && e.name;
      setMsg(mid, n === 'NotAllowedError' || n === 'SecurityError' ? 'Kamera izni verilmedi. Adres çubuğundaki kilit / kamera simgesinden bu siteye kamera izni verip tekrar deneyin.' :
        n === 'NotReadableError' || n === 'TrackStartError' ? 'Kamera başka bir uygulamada açık. Onu kapatıp tekrar deneyin.' :
        n === 'NotFoundError' || n === 'OverconstrainedError' ? 'Kamera bulunamadı.' : 'Kamera açılamadı: ' + ((e && e.message) || n || 'bilinmeyen hata'), 'err');
      return;
    }
    var dk = await decP;
    busy(btn, false);
    if (typeof dk !== 'string') { this.stop(); setMsg(mid, 'Kod çözücü yüklenemedi: ' + ((dk && dk.message) || dk) + '. İnternete bağlıyken sayfayı bir kez yenileyin.', 'err'); return; }
    this.want = true; this.running = this.mode !== 'cfg'; this.paused = false; this.hold = null;
    this.stageOn(true);
    LAYER.push('cam', function () { CAM.stop(false, true); });
    this.caps(); this.devices(); setMsg(mid, '');
    if (this.mode === 'cfg') { this.state('Önizleme · ayarlar canlı uygulanır'); renderCamPanel(); try { $('stage').scrollIntoView({ block: 'start' }); } catch (e) { } return; }
    if (this.serial) SER.begin();
    this.state(this.mode === 'pair' ? 'PC ekranındaki QR kodu çerçeveye alın' : this.serial ? 'Makaraları sırayla çerçeveden geçirin' : 'Etiketi çerçeveye alın');
    var iss = hasCam() && OmsCam.openIssue ? OmsCam.openIssue() : '';   /* V13b: say it when the chosen camera / settings did not fully apply */
    if (iss) this.state(iss, 'warn');
    this.armIdle(); this.lock(true); this.schedule(300);
    try { $('stage').scrollIntoView({ block: 'start' }); } catch (e) { }
  },
  open: async function (deviceId) {
    var vid0 = $('cvid');
    if (hasCam()) {
      /* camera choice, resolution, fps and the saved advanced settings come from the OmsCam module (Kamera tab) */
      vid0.muted = true; vid0.setAttribute('playsinline', '');
      var r = await OmsCam.open(vid0);
      this.stream = r.stream; this.track = r.track || (r.stream && r.stream.getVideoTracks()[0]) || null;
      if (vid0.srcObject !== r.stream) { vid0.srcObject = r.stream; try { await vid0.play(); } catch (e) { } }
      return;
    }
    var v = { width: { ideal: 1920 }, height: { ideal: 1080 } };
    if (deviceId) v.deviceId = { exact: deviceId }; else v.facingMode = { ideal: 'environment' };
    var s;
    try { s = await navigator.mediaDevices.getUserMedia({ audio: false, video: v }); }
    catch (e) {
      if (deviceId && e && (e.name === 'OverconstrainedError' || e.name === 'NotFoundError')) { LS.del(K.cam); return this.open(null); }
      if (e && e.name === 'OverconstrainedError') s = await navigator.mediaDevices.getUserMedia({ audio: false, video: true }); else throw e;
    }
    this.stream = s; this.track = s.getVideoTracks()[0] || null;
    var vid = $('cvid'); vid.muted = true; vid.setAttribute('playsinline', ''); vid.srcObject = s;
    try { await vid.play(); } catch (e) { }
  },
  caps: function () {
    var t = this.track, c = {}, st = {}; if (!t) return;
    try { c = t.getCapabilities ? t.getCapabilities() : {}; } catch (e) { }
    try { st = t.getSettings ? t.getSettings() : {}; } catch (e) { }
    /* OmsCam.open already applied the saved focus/exposure settings; only the legacy path forces continuous focus */
    if (!hasCam() && c.focusMode && c.focusMode.indexOf('continuous') >= 0) { try { t.applyConstraints({ advanced: [{ focusMode: 'continuous' }] }).catch(function () { }); } catch (e) { } }
    $('stage').classList.toggle('mirror', st.facingMode === 'user');
    this.torch = false; var bt = $('btorch'); bt.hidden = !c.torch; bt.setAttribute('aria-pressed', 'false'); bt.classList.remove('nudge'); bt.querySelector('use').setAttribute('href', '#i-torch');
    buildZoom(c, st.zoom || (c.zoom && c.zoom.min) || 1);
    $('bsnap').querySelector('span').textContent = this.hdOk() ? 'HD fotoğraf' : 'Fotoğraf çek';   /* V12b */
  },
  devices: function () {
    if (!navigator.mediaDevices.enumerateDevices) return; var self = this;
    navigator.mediaDevices.enumerateDevices().then(function (ds) {
      var cur = '', vids = []; try { cur = self.track.getSettings().deviceId || ''; } catch (e) { }
      for (var i = 0; i < ds.length; i++) if (ds[i].kind === 'videoinput') vids.push(ds[i]);
      self.vids = vids; self.vcur = cur; var b = $('bcamsw'); b.hidden = vids.length < 2 || hasCam();
      $('bcamset').hidden = !hasCam() || self.mode !== 'scan';
      if (vids.length > 1) { var ix = 0; vids.forEach(function (d, i) { if (d.deviceId === cur) ix = i; }); b.setAttribute('aria-label', 'Kamera değiştir (' + (ix + 1) + '/' + vids.length + (vids[ix].label ? ': ' + vids[ix].label : '') + ')'); }
    }).catch(function () { });
  },
  nextCam: async function () {
    var vs = this.vids || []; if (vs.length < 2) return; var ix = 0;
    for (var i = 0; i < vs.length; i++) if (vs[i].deviceId === this.vcur) ix = i;
    var id = vs[(ix + 1) % vs.length].deviceId; LS.set(K.cam, id);
    this.release();
    try { await this.open(id); } catch (e) { LS.del(K.cam); try { await this.open(null); } catch (e2) { this.stop(); setMsg(this.msgId(), 'Kamera açılamadı.', 'err'); return; } }
    this.running = true; this.paused = false; this.caps(); this.devices(); this.armIdle(); this.schedule(300);
  },
  release: function () {
    clearTimeout(this.timer); this.timer = null; this.running = false;
    if (this.stream) this.stream.getTracks().forEach(function (t) { try { t.stop(); } catch (e) { } });
    this.stream = null; this.track = null;
    var vid = $('cvid'); try { vid.pause(); } catch (e) { } vid.srcObject = null;
  },
  stop: function (keepWant, fromBack) {
    if (this.serial && !keepWant) SER.end(false);
    this.release(); this.paused = false; this.lock(false); this.hold = null;
    if (!keepWant) this.want = false;
    if (!fromBack) LAYER.pop('cam'); else LAYER.stack = LAYER.stack.filter(function (x) { return x.n !== 'cam'; });
    this.stageOn(false); $('zseg').hidden = true; busy($('bcam'), false); busy($('bpairqr'), false); busy($('bcfgcam'), false);
    if (VIEW.cur === 'cam') renderCamPanel();
  },
  /* reopen after the tab came back / bfcache restore: same mode, no new layer */
  reopen: function () {
    var self = this;
    return this.open(LS.get(K.cam)).then(function () {
      self.caps(); self.devices();
      if (self.mode === 'cfg') { renderCamPanel(); return; }
      self.running = true; self.paused = false; self.armIdle(); self.lock(true); self.schedule(300);
    }, function () { self.stop(); });
  },
  /* OmsCam panel changed camera / resolution / fps: reopen the stream with the new constraints */
  restart: async function () {
    if (!this.stream) return;
    var wasRunning = this.mode !== 'cfg';
    this.release();
    try { await this.open(LS.get(K.cam)); }
    catch (e) { this.stop(); setMsg(this.msgId(), 'Kamera bu ayarla açılamadı: ' + ((e && (e.message || e.name)) || 'bilinmeyen hata'), 'err'); return; }
    this.caps(); this.devices();
    if (wasRunning) { this.running = true; this.paused = false; this.armIdle(); this.schedule(300); }
  },
  resume: function () { this.paused = false; this.running = true; $('cpause').hidden = true; this.armIdle(); this.lock(true); this.schedule(100); },
  /* V12b: the live stream cannot exceed the camera driver's video sizes (often 1920x1080 even on 12 / 50 MP phones); a still from
     ImageCapture.takePhoto has the full photo size (camsettings.js explains, with sources). Decoded like a photo, committed like a hit. */
  hdOk: function () { return this.mode !== 'cfg' && hasCam() && !!OmsCam.canPhoto && OmsCam.canPhoto(this.track); },
  hd: async function () {
    var t = this.track, m = this.mode; if (!t || this.hdBusy) return;
    this.hdBusy = true; this.running = false; clearTimeout(this.timer); $('chelp').hidden = true;
    var b = $('bsnap'); busy(b, true);
    this.state('HD fotoğraf çekiliyor… telefonu sabit tutun');
    try {
      var blob = await OmsCam.takePhoto(t);
      if (!this.stream) return;
      var dims = ''; try { var bm = await createImageBitmap(blob); dims = ' (' + bm.width + '×' + bm.height + ')'; if (bm.close) bm.close(); } catch (e) { }
      this.state('Çözülüyor' + dims + '…');
      var hits = (await decodePhoto(blob, m, true)).filter(function (h) { return h && h.text; });
      if (!this.stream) return;
      this.running = true; this.paused = false; this.armIdle();
      if (!hits.length) { FB.go('err'); this.state('HD fotoğrafta' + dims + ' kod bulunamadı. Yaklaşıp tekrar deneyin.', 'warn'); this.schedule(1500); return; }
      if (m === 'pair' || this.serial) { this.onHits(hits); return; }
      this.commit(pickPrimary(hits), Date.now());
    } catch (e) {
      if (this.stream) { this.running = true; this.paused = false; this.armIdle(); this.schedule(800); }
      this.state('HD fotoğraf çekilemedi: ' + ((e && e.message) || e), 'warn');
    } finally { this.hdBusy = false; busy(b, false); }
  },
  schedule: function (ms) { clearTimeout(this.timer); var self = this; this.timer = setTimeout(function () { self.tick(); }, ms); },
  /* hints at 6 s / 12 s, pause after 60 s idle (battery, heat) */
  armIdle: function () { this.idle.t0 = Date.now(); this.idle.hint = 0; $('chelp').hidden = true; $('btorch').classList.remove('nudge'); },
  checkIdle: function () {
    var d = Date.now() - this.idle.t0;
    if (d > 60000) { this.running = false; clearTimeout(this.timer); $('chelp').hidden = true; $('cpause').hidden = false; this.lock(false); this.state('Tarama duraklatıldı'); return true; }
    if (d > 12000 && this.idle.hint < 2) { this.idle.hint = 2; this.state(this.serial ? 'Sıradaki makarayı çerçeveye alın' : 'Okunamıyor mu?', this.serial ? '' : 'warn'); $('chelp').hidden = !!this.serial; }
    else if (d > 6000 && this.idle.hint < 1) { this.idle.hint = 1; this.state('Yaklaştırın, etiketi düz tutun'); if (!$('btorch').hidden) $('btorch').classList.add('nudge'); if (!this.serial && this.hdOk()) $('chelp').hidden = false; }
    return false;
  },
  /* keep the screen on while scanning (secure context; Chrome, Safari 16.4+) */
  lock: function (on) {
    var self = this;
    try {
      if (on && ('wakeLock' in navigator) && !this.wl) {
        this.wl = 'p';
        navigator.wakeLock.request('screen').then(function (l) { if (self.wl !== 'p') { try { l.release(); } catch (e) { } return; } self.wl = l; l.addEventListener('release', function () { if (self.wl === l) self.wl = null; }); }, function () { self.wl = null; });
      } else if (!on && this.wl) { if (this.wl !== 'p') { try { this.wl.release(); } catch (e) { } } this.wl = null; }
    } catch (e) { }
  },
  /* crop of the video under the aim box (object-fit: cover mapping), scaled to at most `max` px; full frame when `full` */
  grab: function (max, full) {
    var vid = $('cvid'), vw = vid.videoWidth, vh = vid.videoHeight; if (!vw || !vh) return null;
    var st = $('stage').getBoundingClientRect(), am = $('stage').querySelector('.aim').getBoundingClientRect();
    var W = st.width || vw, H = st.height || vh, s = Math.max(W / vw, H / vh), dx = (W - vw * s) / 2, dy = (H - vh * s) / 2;
    var mx = am.width * 0.15, my = am.height * 0.15;
    var x0 = Math.max(0, (am.left - st.left - mx - dx) / s), y0 = Math.max(0, (am.top - st.top - my - dy) / s);
    var x1 = Math.min(vw, (am.right - st.left + mx - dx) / s), y1 = Math.min(vh, (am.bottom - st.top + my - dy) / s);
    if (full || !(x1 - x0 > 32 && y1 - y0 > 32)) { x0 = 0; y0 = 0; x1 = vw; y1 = vh; }
    if ($('stage').classList.contains('mirror')) { var w0 = x1 - x0; x0 = vw - x1; x1 = x0 + w0; }
    var sw = x1 - x0, sh = y1 - y0, sc = Math.min(1, max / Math.max(sw, sh)), cw = Math.max(1, Math.round(sw * sc)), ch = Math.max(1, Math.round(sh * sc));
    if (!this.canvas) { this.canvas = document.createElement('canvas'); }
    var c = this.canvas; if (c.width !== cw) c.width = cw; if (c.height !== ch) c.height = ch;
    if (!this.ctx) { try { this.ctx = c.getContext('2d', { willReadFrequently: true }); } catch (e) { } if (!this.ctx) this.ctx = c.getContext('2d'); }
    this.ctx.drawImage(vid, x0, y0, sw, sh, 0, 0, cw, ch);
    var d = this.ctx.getImageData(0, 0, cw, ch);
    return { lum: toLum(d.data, cw * ch), w: cw, h: ch };
  },
  tick: function () {
    if (!this.running || this.paused || this.inflight) return;
    if (document.hidden) { this.schedule(500); return; }
    if (this.checkIdle()) return;
    var vid = $('cvid'); if (!vid.videoWidth || vid.readyState < 2) { this.schedule(150); return; }
    var self = this, p;
    this.inflight = true;
    try {
      if (DEC.kind === 'native' && DEC.det) p = DEC.detect(vid);
      else if (DEC.kind === 'worker' || DEC.kind === 'main') {
        this.k++;
        /* scan: aim-box crop, alternating near (fast) / far (small modules), every 4th frame the whole frame;
           pair: whole frame (the QR is often larger than the aim box) */
        var g = this.mode === 'pair' ? this.grab(this.k % 2 ? 1280 : 960, true) : this.grab(this.k % 2 ? 1280 : 800, this.k % 4 === 3);
        var set = this.mode === 'pair' ? 'qr' : (this.hold || (this.serial && SER.waiting() && this.k % 2)) ? 'd2' : 'all';
        p = g ? DEC.decodeLum(g.lum, g.w, g.h, set).then(function (r) { return r ? [r] : []; }) : Promise.resolve([]);
      } else p = DEC.init().then(function () { return []; });
    } catch (e) { p = Promise.reject(e); }
    p.then(function (hits) {
      self.inflight = false; if (!self.running || self.paused) return;
      self.onHits(hits || []);
    }, function () {
      self.inflight = false; if (self.running && !self.paused) self.schedule(250);
    });
  },
  onHits: function (hits) {
    var now = Date.now(), gap = DEC.kind === 'main' ? 120 : 40;
    hits = hits.filter(function (h) { return h && h.text; });
    if (this.mode === 'pair') {
      var pc = null; hits.forEach(function (h) { if (String(h.text).indexOf('OMSP1') >= 0) pc = h; });
      if (pc) { this.armIdle(); this.paused = true; this.state('Kod okundu', 'ok'); this.flash('hit'); doPair(pc.text, 'pcammsg').then(function (ok) { if (!ok && CAM.stream) { CAM.paused = false; CAM.schedule(1500); } }); return; }
      if (hits.length && now - this.warnT > 3000) { this.warnT = now; this.state('Bu bir OMS eşleştirme kodu değil', 'warn'); FB.go('err'); }
      this.schedule(gap); return;
    }
    if (hits.length) this.armIdle();
    if (this.serial) { this.serialHits(hits, now, gap); return; }
    var prim = pickPrimary(hits);
    if (prim && !is2D(prim.format)) {
      /* only a 1D code: wait briefly for the label's 2D code before accepting it */
      if (!this.hold) { this.hold = { hit: prim, t: now }; this.state('Okunuyor…'); this.schedule(gap); return; }
      this.hold.hit = prim;
      if (now - this.hold.t < HOLD_1D_MS) { this.schedule(gap); return; }
    } else if (!prim && this.hold) {
      if (now - this.hold.t < HOLD_1D_MS) { this.schedule(gap); return; }
      prim = this.hold.hit;
    }
    this.hold = null;
    if (!prim) { this.schedule(gap); return; }
    this.commit(prim, now);
  },
  commit: function (hit, now) {
    var raw = String(hit.text);
    var same = this.last.raw === raw && now - this.last.t < DUP_MS;
    this.last = { raw: raw, t: now };
    if (same) { this.schedule(250); return; }   /* same code within 3 s: not counted again */
    var r = addLine(raw, hit.format, 'camera'), self = this;
    var wait = r.ok && smode() === 'manual';
    if (wait) { this.paused = true; clearTimeout(this.timer); }   /* the editor opens once the stock lookup is done */
    if (r.ok) {
      PLUS.hide();
      /* scanning stays paused until the (fast, local) stock lookup is done, so the feedback is one clear signal */
      stockCheck(r.line).then(function (st) {
        var title = lineInfo(r.line).title;
        if (st.same) {
          FB.go('stock'); self.flash('dup');
          self.state('STOKTA · aynı makara · ' + stockWhere(st.same), 'warn');
          render(VIEW.cur); if (self.running && !self.paused && !wait) self.schedule(1600);
        } else {
          FB.go('ok'); self.flash('hit');
          self.state('Eklendi · ' + title + (st.mpn ? ' · ' + stockLine(st.mpn) : ' · sepette ' + lines().length), 'ok');
          if (self.running && !self.paused && !wait) self.schedule(900);
        }
        if (wait) self.editAfter(r.line);
      });
    } else if (r.dup) {
      FB.go('inbasket'); this.flash('dup');
      this.state('Zaten sepette · ' + lineInfo(r.line).title + ' · ' + reelsOf(r.line) + ' makara', 'warn');
      PLUS.show(r.line);
      this.schedule(600);
    } else { this.state(r.err === 'full' ? 'Sepet dolu' : 'Kod eklenemedi', 'err'); this.schedule(1200); }
  },
  /* Manuel: line editor over the paused camera; any close (save / Kapat / back / delete) resumes scanning */
  editAfter: function (line) {
    var self = this;
    if (!this.stream) { this.paused = false; return; }
    editLine(line.id, { after: true, onDone: function () {
      if (!self.stream || self.mode !== 'scan') return;
      self.paused = false; self.last = { raw: line.raw, t: Date.now() };
      self.armIdle(); self.state('Sıradaki etiketi çerçeveye alın'); self.schedule(500);
    } });
  },
  /* Seri: OmsSerial decides on every frame (also empty ones, so a waiting 1D-only reel gets committed) */
  serialHits: function (hits, now, gap) {
    var d = SER.offer(hits);
    if (d && d.action === 'add' && d.hit) { if (SER.add(d.hit, now)) { this.schedule(SER_GAP_MS); return; } this.schedule(1200); return; }
    if (d && d.action === 'ignore' && d.reason === 'inBasket' && now - SER.dupT > 2500) {
      SER.dupT = now; FB.go('inbasket'); this.flash('dup');
      var ex = null, raw = d.hit ? String(d.hit.text || d.hit.rawValue || '') : '';
      lines().forEach(function (l) { if (l.raw === raw.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, '')) ex = l; });
      this.state('Zaten sepette' + (ex ? ' · ' + lineInfo(ex).title + ' · ' + reelsOf(ex) + ' makara' : ''), 'warn');
      if (ex) PLUS.show(ex);
    } else if (d && d.action === 'wait') this.state('Okunuyor…');
    this.schedule(gap);
  }
};
/* ===================== V8: serial scan session (OmsSerial, serialscan.js; spec section 19) ===================== */
var SER = {
  s: null, n: 0, recent: [], dupT: 0,
  begin: function () {
    var self = this; this.n = 0; this.recent = []; this.dupT = 0; this.s = null;
    try {
      this.s = OmsSerial.create({ inBasket: function (raw) { return self.has(raw); }, now: function () { return Date.now(); } });
    } catch (e) { this.s = null; }
    this.paint();
  },
  has: function (raw) {
    raw = String(raw == null ? '' : raw).replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, '');
    var ls = lines(); for (var i = 0; i < ls.length; i++) if (ls[i].raw === raw) return true;
    return false;
  },
  waiting: function () { try { return !!(this.s && this.s.waiting); } catch (e) { return false; } },
  forget: function (raw) { try { if (this.s && this.s.forget) this.s.forget(raw); } catch (e) { } },
  offer: function (hits) {
    if (!this.s) {   /* module failed: plain behaviour, the first code of a frame that is not in the basket yet */
      var h = pickPrimary(hits); return h && !this.has(h.text) ? { action: 'add', hit: h, reason: 'fallback' } : null;
    }
    try { return this.s.offer(hits); } catch (e) { return null; }
  },
  count: function () { var c = 0; try { c = this.s ? +this.s.count : 0; } catch (e) { c = 0; } return Math.max(c || 0, this.n); },
  /* one decided reel -> basket line, short beep, counter + last 3; stock lookup marks the row afterwards */
  add: function (hit, now) {
    var raw = String(hit.text != null ? hit.text : hit.rawValue || ''), fmt = hit.format || '';
    var r = addLine(raw, fmt, 'serial');
    if (!r.ok) {
      if (r.dup) return true;
      FB.go('err'); CAM.state(r.err === 'full' ? 'Sepet dolu (' + MAX_LINES + ' satır): önce bu sepeti gönderin' : 'Kod eklenemedi', 'err'); return false;
    }
    this.n++;
    var inf = lineInfo(r.line), row = { id: r.line.id, t: inf.title, q: inf.qty ? inf.qty.t : '', w: false };
    this.recent.unshift(row); if (this.recent.length > 3) this.recent.length = 3;
    FB.go('ser'); CAM.flash('hit');
    CAM.state('Eklendi · ' + inf.title, 'ok');
    this.paint();
    var self = this;
    stockCheck(r.line).then(function (st) {
      if (!st.same) return;
      row.w = true; row.q = 'STOKTA · ' + stockWhere(st.same); self.paint();
      if (CAM.serial && CAM.stream) { FB.go('dup'); CAM.state('STOKTA · aynı makara · ' + inf.title, 'warn'); }
    });
    return true;
  },
  paint: function () {
    var c = this.count();
    $('sercntn').textContent = fmtQty(c);
    $('sercnt').setAttribute('aria-label', c + ' makara eklendi');
    var ul = $('serlist');
    ul.innerHTML = this.recent.map(function (x) {
      return '<li' + (x.w ? ' class="w"' : '') + '>' + ico(x.w ? 'i-warn' : 'i-check') + '<span class="t">' + esc(x.t) + '</span>' + (x.q ? '<span class="q">' + esc(x.q) + '</span>' : '') + '</li>';
    }).join('');
    ul.hidden = !this.recent.length || !CAM.serial;
  },
  /* Bitti / camera closed: commit what the session still holds (a waiting 1D-only reel), then summary */
  end: function (toBasket) {
    if (!this.s && !this.n) { CAM.serial = false; return; }
    var self = this, out = [];
    try { if (this.s && this.s.flush) out = this.s.flush() || []; } catch (e) { out = []; }
    out.forEach(function (d) { if (d && d.action === 'add' && d.hit) self.add(d.hit, Date.now()); });
    var n = this.count();
    try { if (this.s && this.s.reset) this.s.reset(); } catch (e) { }
    this.s = null; this.recent = []; this.n = 0; CAM.serial = false;
    if (toBasket) showView('basket');
    toast(n ? fmtQty(n) + ' makara sepete eklendi' : 'Yeni makara eklenmedi', false, n && !toBasket ? { t: 'Sepete git', fn: function () { showView('basket'); } } : null);
  }
};
$('bdone').addEventListener('click', function () { SER.end(true); CAM.stop(); });
function buildZoom(c, cur) {
  var box = $('zseg');
  if (!(c && c.zoom && c.zoom.max > c.zoom.min)) { box.hidden = true; box.innerHTML = ''; return; }
  var vals = [c.zoom.min, 2, Math.min(c.zoom.max, 4)].filter(function (x, i, a) { return x >= c.zoom.min && x <= c.zoom.max && a.indexOf(x) === i; });
  box.innerHTML = vals.map(function (x) { var t = (+x.toFixed(1)) + '×'; return '<button type="button" role="radio" aria-checked="' + (Math.abs(x - cur) < 0.05) + '" aria-label="Yakınlaştırma ' + t + '" data-z="' + x + '">' + t + '</button>'; }).join('');
  box.hidden = vals.length < 2;
}
$('zseg').addEventListener('click', function (e) {
  var b = e.target.closest('button'); if (!b || !CAM.track) return; var z = +b.getAttribute('data-z');
  (hasCam() && OmsCam.zoom ? Promise.resolve(OmsCam.zoom(CAM.track, z)) : CAM.track.applyConstraints({ advanced: [{ zoom: z }] })).then(function () { [].forEach.call($('zseg').children, function (x) { x.setAttribute('aria-checked', x === b ? 'true' : 'false'); }); }, function () { });
});
$('bcam').addEventListener('click', function () { FB.unlock(); CAM.start('scan'); });
$('bcamoff').addEventListener('click', function () { CAM.stop(); });
$('bcamsw').addEventListener('click', function () { CAM.nextCam(); });
$('bstagebasket').addEventListener('click', function () { showView('basket'); });
function camToManual() { var m = CAM.mode; CAM.stop(); setTimeout(function () { var el = $(m === 'pair' ? 'pcode' : 'code'); el.focus(); try { el.scrollIntoView({ block: 'center' }); } catch (e) { } }, 60); }
$('bmanual').addEventListener('click', camToManual);
$('bman3').addEventListener('click', camToManual);
$('bsnap').addEventListener('click', function () {
  if (CAM.hdOk()) { CAM.hd(); return; }   /* V12b: full-resolution still in the page */
  var m = CAM.mode; CAM.stop(); PH.mode = m; $('ph').click();
});
$('bresume').addEventListener('click', function () { CAM.resume(); CAM.state('Etiketi çerçeveye alın'); });
$('btorch').addEventListener('click', function () {
  var t = CAM.track; if (!t) return; var on = !CAM.torch, self = this; self.classList.remove('nudge');
  (hasCam() && OmsCam.torch ? Promise.resolve(OmsCam.torch(t, on)) : t.applyConstraints({ advanced: [{ torch: on }] })).then(function () { CAM.torch = on; self.setAttribute('aria-pressed', on ? 'true' : 'false'); self.querySelector('use').setAttribute('href', on ? '#i-torch-f' : '#i-torch'); }, function () { CAM.state('Işık bu kamerada kullanılamıyor', 'warn'); });
});
document.addEventListener('visibilitychange', function () {
  if (document.hidden) { flushSave(); if (CAM.stream) { CAM.release(); CAM.lock(false); } }
  else if (CAM.want && !CAM.stream && (VIEW.cur === 'scan' || VIEW.cur === 'pair' || VIEW.cur === 'cam')) CAM.reopen();
});
window.addEventListener('pagehide', function () { flushSave(); if (CAM.stream) { CAM.release(); CAM.lock(false); } });
window.addEventListener('pageshow', function (e) {
  if (e.persisted && CAM.want && !CAM.stream && !document.hidden) CAM.reopen();
});
document.addEventListener('keydown', function (e) {
  if (e.key !== 'Escape') return;
  if (!$('modal').hidden) { closeModal(); return; }
  if (CAM.stream || CAM.want) CAM.stop();
});

/* ----- photo fallback (iOS without good live decoding, blurry labels) ----- */
var PH = { mode: 'scan' };
$('bphoto').addEventListener('click', function () { PH.mode = 'scan'; $('ph').click(); });
function loadImage(file) {
  if (window.createImageBitmap) return createImageBitmap(file, { imageOrientation: 'from-image' }).catch(function () { return createImageBitmap(file); }).catch(function () { return imgFallback(file); });
  return imgFallback(file);
}
function imgFallback(file) { return new Promise(function (res, rej) { var u = URL.createObjectURL(file), im = new Image(); im.onload = function () { URL.revokeObjectURL(u); res(im); }; im.onerror = function () { URL.revokeObjectURL(u); rej(new Error('görsel okunamadı')); }; im.src = u; }); }
async function decodePhoto(file, mode, big) {
  await DEC.init();
  var bmp = await loadImage(file), W = bmp.width || bmp.naturalWidth, Hh = bmp.height || bmp.naturalHeight, hits = [];
  try {
    if (DEC.kind === 'native' && DEC.det) { try { hits = await DEC.detect(bmp); } catch (e) { hits = []; } if (hits.length) return hits; await DEC.initP; }
    if (DEC.kind !== 'worker' && DEC.kind !== 'main') await DEC.zx();
    var sets = mode === 'pair' ? ['qr'] : ['d2', 'all'], edges = big ? [1600, 2600, 4000] : [1200, 2000, 3200], prev = 0;   /* V12b: big stills get a full-size pass */
    for (var k = 0; k < edges.length; k++) {
      var s = Math.min(1, edges[k] / Math.max(W, Hh)), w = Math.max(1, W * s | 0), h = Math.max(1, Hh * s | 0);
      if (w === prev) break; prev = w;
      var c = document.createElement('canvas'); c.width = w; c.height = h; var cx = c.getContext('2d'); cx.drawImage(bmp, 0, 0, w, h);
      var data = cx.getImageData(0, 0, w, h).data;
      for (var j = 0; j < sets.length; j++) { var r = await DEC.decodeLum(toLum(data, w * h), w, h, sets[j]); if (r) return [r]; }
      await tick(0);
    }
  } finally { if (bmp.close) try { bmp.close(); } catch (e) { } }
  return [];
}
$('ph').addEventListener('change', async function () {
  var f = this.files && this.files[0]; this.value = ''; if (!f) return;
  var mode = PH.mode, mid = mode === 'pair' ? 'pcammsg' : 'cammsg', btn = $(mode === 'pair' ? 'bpairqr' : 'bphoto');
  busy(btn, true); setMsg(mid, 'Çözülüyor…');
  try {
    var hits = await decodePhoto(f, mode);
    if (mode === 'pair') {
      var pc = null; hits.forEach(function (h) { if (String(h.text).indexOf('OMSP1') >= 0) pc = h; });
      if (!pc) { setMsg(mid, 'Fotoğrafta eşleştirme QR\'ı bulunamadı. Kodu yapıştırmayı deneyin.', 'err'); return; }
      setMsg(mid, ''); await doPair(pc.text, 'pcammsg'); return;
    }
    var prim = pickPrimary(hits);
    if (!prim) { FB.go('err'); setMsg(mid, 'Kod okunamadı. Etiketi yakından, düz ve ışıklı çekin ya da kodu elle yazın.', 'err'); return; }
    setMsg(mid, 'Okundu (' + fmtName(prim.format) + ')', 'ok');
    var pr = addLine(prim.text, prim.format, 'photo');
    if (pr.ok && smode() === 'manual') editLine(pr.line.id, { after: true });
  } catch (e) { setMsg(mid, 'Çözücü hatası: ' + ((e && e.message) || e), 'err'); }
  finally { busy(btn, false); }
});

/* ===================== V6: camera settings tab (OmsCam, camsettings.js) ===================== */
function hasCam() { return !!(window.OmsCam && OmsCam.open && OmsCam.renderPanel); }
function camInit() { if (hasCam() && OmsCam.load) { try { OmsCam.load(); } catch (e) { } } bindTap(); }
var PANEL = null;   /* OmsCam.renderPanel handle {refresh, destroy, root}; destroyed when the tab is left (it runs a timer) */
function renderCam() {
  $('camnomod').hidden = hasCam();
  renderCamPanel();
}
function renderCamPanel() {
  var box = $('campanel');
  if (!hasCam()) { box.innerHTML = ''; return; }
  if (VIEW.cur !== 'cam') return;
  try {
    if (PANEL && PANEL.root && box.contains(PANEL.root)) { PANEL.refresh(); return; }
    PANEL = OmsCam.renderPanel(box, { getTrack: function () { return CAM.track; }, onRestart: function () { return CAM.restart(); },
      onDone: function () { showView('scan'); CAM.start('scan'); } });   /* V12b: 'Kaydet ve taramaya dön' */
  } catch (e) { PANEL = null; box.innerHTML = '<div class="note err">' + ico('i-err') + '<div>Kamera ayar paneli açılamadı: ' + esc((e && e.message) || e) + '</div></div>'; }
}
function dropCamPanel() { if (PANEL) { try { PANEL.destroy(); } catch (e) { } PANEL = null; } $('campanel').innerHTML = ''; }
$('bcfgcam').addEventListener('click', function () { CAM.start('cfg'); });
$('bgocam').addEventListener('click', function () { showView('cam'); });
$('bcamback').addEventListener('click', function () { showView('set'); });
$('bcamset').addEventListener('click', function () { showView('cam'); CAM.start('cfg'); });
/* tap to focus on the preview (scan and Kamera tab): OmsCam honours the panel's "Dokunarak odakla" switch and draws the ring */
var TAP = { bound: false };
function bindTap() {
  if (TAP.bound || !hasCam()) return;
  TAP.bound = true;
  var mirrored = function () { return $('stage').classList.contains('mirror'); };
  if (OmsCam.bindTapFocus) {
    OmsCam.bindTapFocus($('cvid'), function () { return CAM.track; }, { host: $('stage'), mirrored: mirrored,
      onFocus: function (ok) { if (CAM.mode === 'scan') CAM.armIdle(); if (!ok) CAM.state('Bu kamera bu noktaya odaklanamadı', 'warn'); } });
    return;
  }
  /* older module without bindTapFocus: map the tap (object-fit: cover) back to the video frame ourselves */
  $('cvid').addEventListener('click', function (e) {
    var t = CAM.track, vid = this; if (!t || !OmsCam.focusAt) return;
    var p = null; try { p = OmsCam.prefs ? OmsCam.prefs(t) : null; } catch (x) { }
    if (p && p.tapFocus === false) return;
    var vw = vid.videoWidth, vh = vid.videoHeight; if (!vw || !vh) return;
    var st = $('stage').getBoundingClientRect(), cx = e.clientX - st.left, cy = e.clientY - st.top;
    var s = Math.max(st.width / vw, st.height / vh), dx = (st.width - vw * s) / 2, dy = (st.height - vh * s) / 2;
    var x = (cx - dx) / s / vw, y = (cy - dy) / s / vh;
    if (mirrored()) x = 1 - x;
    x = Math.min(1, Math.max(0, x)); y = Math.min(1, Math.max(0, y));
    var fr = $('fring'); fr.style.left = cx + 'px'; fr.style.top = cy + 'px'; fr.classList.remove('on'); void fr.offsetWidth; fr.classList.add('on');
    clearTimeout(TAP.t); TAP.t = setTimeout(function () { fr.classList.remove('on'); }, 900);
    if (CAM.mode === 'scan') CAM.armIdle();
    Promise.resolve().then(function () { return OmsCam.focusAt(t, x, y); }).then(function (ok) {
      if (ok === false) CAM.state('Bu kamera dokunarak odaklamayı desteklemiyor', 'warn');
    }, function () { CAM.state('Bu kamera dokunarak odaklamayı desteklemiyor', 'warn'); });
  });
}

/* ===================== V6: stock catalog (OmsCatalog, catalog.js; spec section 12) ===================== */
var CATR = new Map();   /* raw -> part | null (lookupRaw result cache) */
var CATQ = { busy: false, t: null };
function hasCat() { return !!(window.OmsCatalog && OmsCatalog.info); }
function catOn() { try { return hasCat() && !!OmsCatalog.info(); } catch (e) { return false; } }
function stockOf(l, p) {
  if (!catOn()) return null;
  var same = CATR.has(l.raw) ? CATR.get(l.raw) : null, mpn = null;
  try {
    if (p && p.mpn) mpn = OmsCatalog.lookupMpn(p.mpn);
    if (!mpn && p && p.supplierPn && OmsCatalog.lookupSpn) mpn = OmsCatalog.lookupSpn(p.supplierPn);
  } catch (e) { mpn = null; }
  return same || mpn ? { same: same, mpn: mpn } : null;
}
function stockWhere(part) {
  var a = []; if (part && part.l) a.push(part.l); if (part && part.q != null) a.push(fmtQty(part.q) + ' adet');
  return a.join(', ') || 'raf no yok';
}
function stockLine(hit) {
  try { if (OmsCatalog.describe) return OmsCatalog.describe(hit); } catch (e) { }
  return 'Stokta' + (hit.q != null ? ': ' + fmtQty(hit.q) + ' adet' : ' var') + (hit.l ? ' · ' + hit.l : '');
}
function stockWhereMpn(hit) { return stockLine(hit).replace(/^Stokta:?\s*/, ''); }
/* lookup for one new line (rawHash is a local SHA-256, a few ms); never waits more than 400 ms */
function stockCheck(l) {
  var p = pv(l.raw, l.fmt) || {};
  if (!catOn()) return Promise.resolve({ same: null, mpn: null });
  var look = OmsCatalog.lookupRaw(l.raw).then(function (part) { CATR.set(l.raw, part || null); return part || null; }, function () { return null; });
  var cap = new Promise(function (r) { setTimeout(function () { r(null); }, 400); });
  return Promise.race([look, cap]).then(function (same) { var st = stockOf(l, p) || {}; return { same: same || st.same || null, mpn: st.mpn || null }; });
}
/* fill the raw-hash cache for rendered lines in the background, then re-render once */
function catFill(ls) {
  if (!catOn() || CATQ.busy) return;
  var todo = []; ls.forEach(function (l) { if (!CATR.has(l.raw) && todo.indexOf(l.raw) < 0) todo.push(l.raw); });
  if (!todo.length) return;
  CATQ.busy = true;
  (async function () {
    var hit = false;
    for (var i = 0; i < todo.length; i++) {
      var part = null; try { part = await OmsCatalog.lookupRaw(todo[i]); } catch (e) { }
      CATR.set(todo[i], part || null); if (part) hit = true;
      if (i % 50 === 49) await tick(0);
    }
    CATQ.busy = false;
    if (hit && (VIEW.cur === 'scan' || VIEW.cur === 'basket')) render(VIEW.cur);
  })();
}
function catChanged() { CATR.clear(); if (VIEW.cur === 'scan' || VIEW.cur === 'basket' || VIEW.cur === 'stock') render(VIEW.cur); else if (VIEW.cur === 'set') renderCatCard(); }
function catInit() {
  if (!hasCat()) return;
  try { if (OmsCatalog.onChange && !CATQ.sub) { CATQ.sub = true; OmsCatalog.onChange(catChanged); } } catch (e) { }
  if (DEMO) return;
  Promise.resolve().then(function () { return OmsCatalog.load(); }).then(function () { catChanged(); catSilent(); }, function () { catSilent(); });
}
function catStale() {
  try { if (OmsCatalog.isStale) return OmsCatalog.isStale(CAT_STALE_MS); } catch (e) { }
  var i = null; try { i = OmsCatalog.info(); } catch (e) { }
  var t = i ? (typeof i.at === 'number' ? i.at : Date.parse(i.at)) : 0;
  return !t || Date.now() - t > CAT_STALE_MS;
}
/* silent sync on page open: catalog older than 12 h, PC addresses known, at most every 30 min, no UI on failure */
function catSilent() {
  if (!hasCat() || !hasLan() || navigator.onLine === false || !catStale()) return;
  var last = +(LS.get(K.catTry) || 0); if (Date.now() - last < CAT_RETRY_MS) return;
  LS.set(K.catTry, String(Date.now()));
  OmsCatalog.sync({ dev: S.pair.dev, sec: secBytes(), lan: S.pair.lan }).then(function (r) {
    if (r && r.ok) { catChanged(); if (VIEW.cur === 'set') renderCatCard(); }
  }, function () { });
}
var CATUI = { cert: '', note: '' };
function renderCatCard() {
  var card = $('c_cat'); card.hidden = !S.pair;
  if (!S.pair) return;
  if (!hasCat()) {
    $('catkv').innerHTML = ''; $('bcatsync').disabled = true; $('bcatcert').hidden = true;
    $('catnote').innerHTML = '<div class="note warn">' + ico('i-warn') + '<div>Stok listesi modülü (catalog.js) yüklenemedi. İnternete bağlıyken sayfayı yenileyin.</div></div>'; $('catnote').hidden = false;
    return;
  }
  var i = null; try { i = OmsCatalog.info(); } catch (e) { }
  var kv = '';
  if (i) {
    var at = typeof i.at === 'number' ? i.at : Date.parse(i.at), stale = catStale();
    [['Son alınma', (isFinite(at) && at ? fmtTime(at) : '?') + (stale ? ' (12 saatten eski)' : '')], ['Parça sayısı', fmtQty(i.count)], ['Bilgisayar', i.pc]].forEach(function (r) {
      if (r[1] != null && r[1] !== '') kv += '<div class="kv"><dt>' + r[0] + '</dt><dd>' + esc(r[1]) + '</dd></div>';
    });
  } else kv = '<div class="kv"><dt>Durum</dt><dd>Henüz alınmadı</dd></div>';
  $('catkv').innerHTML = kv;
  var noLan = !hasLan(), canRef = catCanRefresh();
  if (!$('bcatsync').getAttribute('aria-busy') || $('bcatsync').getAttribute('aria-busy') === 'false') $('bcatsync').disabled = !canRef;
  var note = CATUI.note;
  if (!canRef) note = '<div class="note info">' + ico('i-info') + '<div>' + (driveOk() ? 'Yukarıdan <b>OneDrive\'a bağlanın</b> ya da telefonu PC ile aynı Wi-Fi ağına bağlayın. ' : '') +
    'PC adresi yok: eşleştirme kodunu PC\'den yeniden okutun ya da <b>Gelişmiş › PC adresleri</b> bölümüne ekleyin.</div></div>';
  $('catnote').innerHTML = note; $('catnote').hidden = !note;
  var c = $('bcatcert'); c.hidden = !CATUI.cert || noLan; if (CATUI.cert) c.href = CATUI.cert;
}
$('bcatsync').addEventListener('click', function () { catManual(this, 'catmsg'); });
$('bstkref').addEventListener('click', function () { catManual(this, 'stkmsg'); });
/* one refresh for both buttons (Ayarlar card, Stok tab): OneDrive first when connected, then the PC on the LAN */
async function catManual(btn, msgId) {
  if (!S.pair || !hasCat()) return;
  busy(btn, true); setMsg(msgId, 'Stok listesi alınıyor…'); CATUI.note = ''; CATUI.cert = ''; renderCatCard(); stkNote();
  var r = await catRefresh(false);
  busy(btn, false);
  if (r.ok) setMsg(msgId, (r.same ? 'Stok listesi zaten güncel: ' : 'Stok listesi alındı: ') + fmtQty(r.count) + ' parça' + (r.src === 'drive' ? ' (OneDrive)' : r.pc ? ' (' + r.pc + ')' : '') + '.', 'ok');
  else setMsg(msgId, r.error, 'err');
  renderCatCard(); stkNote();
  if (msgId === 'catmsg' && btn) btn.disabled = !catCanRefresh();
}
function lanCertNote(btnName) {
  return '<div class="note warn">' + ico('i-shield') + '<div><b>PC\'ye bağlanılamadı.</b> Telefon PC ile aynı Wi-Fi ağında olmalı. PC kendi imzaladığı bir sertifika kullanır; tarayıcının bir kez onaylaması gerekir: ' +
    '<b>PC sertifikasını onayla</b> düğmesi PC adresini yeni sekmede açar. Uyarı sayfasında <b>Gelişmiş</b> → <b>Devam et</b> deyin, sonra bu sekmeye dönüp yeniden <b>' + btnName + '</b>\'e basın.</div></div>';
}
function catCanRefresh() { return !!(S.pair && hasCat() && ((driveOn() && catDriveOk()) || hasLan())); }
function catDriveOk() { return !!(window.OmsDrive && OmsDrive.download && window.OmsCatalog && OmsCatalog.loadArmor); }
function driveWhy(d) {
  d = d || {};
  if (d.code === 'NOT_FOUND') return 'OneDrive\'da stok listesi henüz yok. PC\'deki OMS programı açık ve bulut klasörü kurulu olmalı; birkaç dakika sonra yeniden deneyin.';
  if (d.code === 'LOGIN') return 'OneDrive bağlantısının süresi doldu: Ayarlar\'dan OneDrive\'a yeniden bağlanın.';
  return d.error || 'Stok listesi OneDrive\'dan alınamadı.';
}
/* -> {ok, count, pc, src:'drive'|'lan', same, error}; never rejects. quiet: no cert note, no UI */
var CATREF = { p: null };
function catRefresh(quiet) {
  if (CATREF.p) return CATREF.p;
  CATREF.p = (async function () {
    if (!S.pair || !hasCat()) return { ok: false, error: 'Stok listesi modülü yüklenemedi. İnternete bağlıyken sayfayı yenileyin.' };
    var errs = [], lanR = null;
    if (driveOn() && catDriveOk()) {
      LS.set(K.drvTry, String(Date.now()));
      var d = null, et = '';
      try { var ci = OmsCatalog.info(); if (ci && ci.count != null) et = LS.get(K.catEtag) || ''; } catch (e) { }
      try { d = await OmsDrive.download(OmsDrive.catalogFileName ? OmsDrive.catalogFileName(S.pair.dev) : 'oms_katalog_' + S.pair.dev + '.omsk', et ? { etag: et } : undefined); }
      catch (e) { d = { ok: false, code: 'ERROR', error: (e && e.message) || String(e) }; }
      if (d && d.ok && d.notModified) { var i0 = null; try { i0 = OmsCatalog.info(); } catch (e) { } return { ok: true, same: true, count: i0 ? i0.count : 0, src: 'drive' }; }
      if (d && d.ok && d.text) {
        try {
          var info = await OmsCatalog.loadArmor(d.text, secBytes(), { dev: S.pair.dev, source: 'onedrive' });
          if (info && info.ok === false) throw new Error(info.error || info.code || 'açılamadı');
          if (d.etag) LS.set(K.catEtag, d.etag); else LS.del(K.catEtag);
          catChanged();
          return { ok: true, count: info && info.count != null ? info.count : 0, pc: info && info.pc, src: 'drive' };
        } catch (e) { errs.push('OneDrive\'daki stok listesi açılamadı: ' + ((e && e.message) || e)); }
      } else errs.push(driveWhy(d));
    }
    if (hasLan()) {
      LS.set(K.catTry, String(Date.now()));
      try { lanR = await OmsCatalog.sync({ dev: S.pair.dev, sec: secBytes(), lan: S.pair.lan || [] }); }
      catch (e) { lanR = { ok: false, code: 'ERROR', error: 'Stok listesi alınamadı: ' + ((e && e.message) || e) }; }
      lanR = lanR || { ok: false };
      if (lanR.ok) { LS.del(K.catEtag); catChanged(); return { ok: true, count: lanR.count, pc: lanR.pc, src: 'lan' }; }
      if (lanR.needCertUrl && !quiet) { CATUI.cert = lanR.needCertUrl; CATUI.note = lanCertNote(VIEW.cur === 'stock' ? 'Yenile' : 'Stok listesini al'); }
      errs.push(lanR.error || ('PC\'den stok listesi alınamadı' + (lanR.code ? ' (' + lanR.code + ')' : '') + '.'));
    }
    if (!errs.length) errs.push(driveOk() ? 'Stok listesi için Ayarlar\'dan OneDrive\'a bağlanın ya da telefonu PC ile aynı Wi-Fi ağına bağlayın.' :
      'Stok listesi telefon PC ile aynı Wi-Fi ağındayken alınır. PC adresi yok: eşleştirme kodunu PC\'den yeniden okutun.');
    return { ok: false, error: errs.join(' '), needCertUrl: lanR && lanR.needCertUrl || '' };
  })().then(function (r) { CATREF.p = null; return r; }, function (e) { CATREF.p = null; return { ok: false, error: String((e && e.message) || e) }; });
  return CATREF.p;
}
/* OneDrive catalog on page open / when the Stok tab opens: older than 30 min, at most every 10 min, silent */
function catDriveSilent() {
  if (!driveOn() || !catDriveOk() || !S.pair || navigator.onLine === false) return;
  var i = null; try { i = OmsCatalog.info(); } catch (e) { }
  var at = i ? (typeof i.at === 'number' ? i.at : Date.parse(i.at)) : 0;
  if (at && Date.now() - at < DRV_CAT_STALE_MS) return;
  if (Date.now() - (+(LS.get(K.drvTry) || 0)) < DRV_CAT_RETRY_MS) return;
  catRefresh(true).then(function (r) { if (VIEW.cur === 'stock' && !r.ok && !(i && i.count)) setMsg('stkmsg', r.error, 'err'); });
}

/* ===================== V8: Stok tab (OmsStock, stockview.js; spec section 18) ===================== */
var STK = { h: null };
function stockOk() { return !!(window.OmsStock && typeof OmsStock.render === 'function'); }
function stockInfo() {
  var i = null; try { i = hasCat() ? OmsCatalog.info() : null; } catch (e) { i = null; }
  return i;
}
function renderStock() {
  var box = $('stockbox');
  $('bstkref').disabled = !catCanRefresh(); $('bstkref').hidden = stockOk();   /* the module has its own Yenile */
  stkNote();
  if (!stockOk()) { box.innerHTML = '<div class="note warn">' + ico('i-warn') + '<div>Stok listesi modülü (stockview.js) yüklenemedi. İnternete bağlıyken sayfayı yenileyin.</div></div>'; STK.h = null; return; }
  try {
    if (STK.h && box.firstChild && typeof STK.h.update === 'function') STK.h.update();
    else {
      box.innerHTML = '';
      STK.h = OmsStock.render(box, {
        getInfo: stockInfo,
        onRefresh: function () { setMsg('stkmsg', ''); CATUI.note = ''; CATUI.cert = ''; return catRefresh(true).then(function (r) { stkNote(); return r; }); }   /* the module shows the result and the certificate link */
      }) || {};
    }
  } catch (e) { STK.h = null; box.innerHTML = '<div class="note err">' + ico('i-err') + '<div>Stok listesi gösterilemedi: ' + esc((e && e.message) || e) + '</div></div>'; }
  catDriveSilent();
}
/* how the list gets here: one line, plus the PC certificate hint after a failed LAN attempt */
function stkNote() {
  var h = '';
  if (CATUI.note) h = CATUI.note + (CATUI.cert ? '<div class="acts"><a class="btn btn-sec" href="' + esc(CATUI.cert) + '" target="_blank" rel="noopener">' + ico('i-shield') + 'PC sertifikasını onayla</a></div>' : '');
  else if (S.pair && !catCanRefresh()) h = '<div class="note info">' + ico('i-info') + '<div>' + (driveOk() ?
    'Stok listesini her yerden almak için <b>Ayarlar</b>\'dan <b>OneDrive\'a bağlanın</b>. Ya da telefonu PC ile aynı Wi-Fi ağına bağlayın.' :
    'Stok listesi telefon PC ile aynı Wi-Fi ağındayken alınır. PC adresi yok: eşleştirme kodunu PC\'den yeniden okutun.') + '</div></div>';
  $('stknote').innerHTML = h;
}

/* ===================== V8: direct OneDrive (OmsDrive, onedrive.js; spec sections 17-18) ===================== */
var DRV = { st: 'off', p: null, started: false, busy: false, loginToastAt: 0, err: '' };
function driveOk() { try { return !!(window.OmsDrive && typeof OmsDrive.available === 'function' && OmsDrive.available() && typeof OmsDrive.send === 'function'); } catch (e) { return false; } }
function driveAcc() { if (!driveOk() || DRV.st !== 'ready') return null; try { return OmsDrive.account() || null; } catch (e) { return null; } }
function driveOn() { return !!driveAcc(); }
function driveName(a) { return a ? (a.name && a.username && a.name !== a.username ? a.name + ' (' + a.username + ')' : (a.username || a.name || 'Microsoft hesabı')) : ''; }
function driveInit() {
  if (!driveOk() || DEMO) return Promise.resolve(false);
  if (DRV.p) return DRV.p;
  DRV.st = 'loading';
  DRV.p = Promise.resolve().then(function () { return OmsDrive.init(); }).then(function () {
    DRV.st = 'ready';
    if (!DRV.started) {
      DRV.started = true;
      try { OmsDrive.onChange(onDriveEvent); } catch (e) { }
      try { OmsDrive.start(); } catch (e) { }
    }
    driveSync(); driveUi(); catDriveSilent();
    resumeDriveSend();
    return true;
  }, function (e) { DRV.st = 'error'; DRV.err = (e && e.message) || String(e); driveUi(); return false; });
  return DRV.p;
}
/* "OneDrive'a bağlan ve gönder" survived the Microsoft redirect: send the open basket now (once) */
function resumeDriveSend() {
  var want = LS.get(K.drvSend);
  if (!want) return;
  LS.del(K.drvSend);
  if (!driveAcc()) return;   /* sign-in cancelled: nothing to resume */
  var b = S.basket;
  if (!b || !b.lines || !b.lines.length || (want !== '1' && b.basket !== want)) return;
  toast('OneDrive\'a bağlandı: ' + driveName(driveAcc()) + ' · sepet gönderiliyor…');
  setTimeout(function () { driveSend(b, false, null); }, 300);
}
/* every place that shows the connection */
function driveUi() {
  if (VIEW.cur === 'set') { renderDrive(); renderCatCard(); renderMs(); }
  else if (VIEW.cur === 'basket') renderBasket();
  else if (VIEW.cur === 'who') renderWho();
  else if (VIEW.cur === 'stock') { $('bstkref').disabled = !catCanRefresh(); stkNote(); }
}
/* Settings > cloud send card: connect button on top; when connected the share-sheet settings step aside */
function renderDrive() {
  var ok = driveOk(), a = driveAcc();
  $('c_drive').hidden = !ok;
  $('sharehd').hidden = !ok || !!a;
  $('c_share').hidden = !!a;
  $('c_drive').classList.toggle('last', !!a);
  if (!ok) return;
  $('drv_off').hidden = !!a; $('drv_on').hidden = !a;
  if (a) {
    $('drvacct').textContent = 'OneDrive\'a bağlı: ' + driveName(a);
    var q = S.hist.filter(function (x) { return x.ch === 'drive' && x.out; }).length;
    $('drvq').hidden = !q; $('drvq').textContent = q ? 'Giden kutusunda ' + q + ' sepet bekliyor.' : '';
    $('bdrvflush').hidden = !q;
  }
}
/* popup inside the tap (desktop / Android); iOS or a blocked popup -> the whole page goes to Microsoft and comes back */
async function driveSignIn(btn, msgId) {
  if (!driveOk()) return false;
  busy(btn, true); if (msgId) setMsg(msgId, '');
  var r;
  try { r = await OmsDrive.signIn(); }
  catch (e) { r = { ok: false, code: 'ERROR', error: (e && e.message) || String(e) }; }
  busy(btn, false);
  r = r || { ok: false };
  if (r.redirecting) { if (msgId) setMsg(msgId, 'Microsoft giriş sayfasına gidiliyor…'); return false; }
  if (DRV.st !== 'ready') await driveInit();
  var a = driveAcc();
  if (r.ok && a) {
    toast('OneDrive\'a bağlandı: ' + driveName(a));
    driveFlush(null, true);
    LS.del(K.drvTry); catDriveSilent();
    driveUi(); if (VIEW.cur === 'hist') renderHist();
    return true;
  }
  var m = r.code === 'CANCELLED' ? 'Bağlanma iptal edildi.' : (r.error || 'OneDrive\'a bağlanılamadı.');
  if (msgId) setMsg(msgId, m, r.code === 'CANCELLED' ? '' : 'err'); else toast(m, r.code !== 'CANCELLED');
  return false;
}
$('bdrvin').addEventListener('click', function () { driveSignIn(this, 'drvmsg'); });
$('bwdrv').addEventListener('click', function () { driveSignIn(this, 'wmsg'); });
$('bdrvflush').addEventListener('click', function () { driveFlush(this); });
$('bdrvout').addEventListener('click', async function () {
  var q = S.hist.filter(function (x) { return x.ch === 'drive' && x.out; }).length;
  if (!(await ask('OneDrive bağlantısı kesilsin mi?', 'Sepetler bundan sonra paylaşma menüsüyle gönderilir.' +
    (q ? ' Giden kutusundaki ' + q + ' sepet yeniden bağlanana kadar bekler.' : ''), 'Bağlantıyı kes', false))) return;
  var btn = this; busy(btn, true);
  try { await OmsDrive.signOut(); toast('OneDrive bağlantısı kesildi'); } catch (e) { toast('Bağlantı kesilemedi: ' + ((e && e.message) || e), true); }
  busy(btn, false); renderSet();
});
/* one tap: seal -> OmsDrive.send (uploads; a passing problem puts it in the outbox itself) -> history */
async function driveSend(b, fromHist, btn) {
  if (DRV.busy) return;
  if (!S.pair) { toast('Önce cihazı eşleştirin.', true); showView('pair'); return; }
  var er = envReady(); if (er) { toast(er, true); return; }
  DRV.busy = true; busy(btn, true);
  var main = btn && btn.id === 'bsend';
  if (main) { $('bsendtx').textContent = 'Gönderiliyor…'; }
  var ctx = null, r = null;
  try {
    ctx = await sealCtx(b, fromHist);
    ctx.dname = OmsDrive.basketFileName ? OmsDrive.basketFileName(S.pair.dev, ctx.date) : 'oms_sepet_' + stamp(+ctx.date) + '_' + S.pair.dev + '.omsb.txt';
    r = await OmsDrive.send({ basketId: b.basket, name: ctx.dname, text: ctx.armor });
  } catch (e) { r = { ok: false, code: ctx ? 'ERROR' : 'SEAL', error: (e && e.message) || String(e) }; }
  DRV.busy = false; busy(btn, false);
  r = r || { ok: false, code: 'ERROR' };
  if (main) renderBasket();
  if (r.code === 'SEAL') { toast('Sepet hazırlanamadı: ' + r.error, true); return; }
  if (r.ok) { FB.go('ok'); markSent(ctx, 'drive'); return; }
  if (r.queued) { markSent(ctx, 'queued', r.code === 'OFFLINE' || r.code === 'LOGIN' ? '' : (r.error || ''), 'drive', r.code === 'LOGIN' ? 'login' : 'queued'); return; }
  /* not sendable this way (e.g. a bad file): keep the basket, offer the other ways */
  FB.go('err');
  toast((r.error || 'OneDrive\'a gönderilemedi.') + ' Başka yolla göndermek için ⋯ düğmesine basın.', true, fromHist ? null : { t: 'Başka yol', fn: function () { startSend(b, fromHist); } });
}
function driveRemove(id) { try { if (window.OmsDrive && OmsDrive.remove) Promise.resolve(OmsDrive.remove(id)).catch(function () { }); } catch (e) { } }
async function driveFlush(btn, quiet) {
  if (!driveOk()) return;
  if (DRV.st !== 'ready') await driveInit();
  busy(btn, true);
  var r = null;
  try { r = await OmsDrive.flush(); } catch (e) { r = null; }
  busy(btn, false);
  await driveSync();
  if (quiet || !r) return;
  var sN = (r.sent || []).length, f = (r.failed || [])[0];
  if (sN && !f) toast(sN + ' sepet OneDrive\'a gönderildi ✓');
  else if (f) toast(f.code === 'LOGIN' ? 'OneDrive\'a yeniden bağlanın.' : (f.error || 'Gönderilemedi.'), f.code !== 'OFFLINE');
  else toast('Bekleyen sepet yok');
}
/* history flags <-> the real OneDrive outbox (page open, after flush) */
function driveSync() {
  if (!driveOk() || DRV.st !== 'ready' || !OmsDrive.pending) return Promise.resolve();
  return Promise.resolve().then(function () { return OmsDrive.pending(); }).then(function (list) {
    var ids = {}, chg = false, login = false;
    try { login = !!(OmsDrive.loginNeeded && OmsDrive.loginNeeded()); } catch (e) { }
    (list || []).forEach(function (it) { if (it && it.basketId) ids[it.basketId] = it; });
    S.hist.forEach(function (h) {
      var it = ids[h.basket];
      if (it && (!h.out || h.ch !== 'drive')) { h.out = 'queued'; h.ch = 'drive'; h.q = h.q || it.queuedAt || Date.now(); chg = true; }
      else if (!it && h.ch === 'drive' && h.out) { delete h.out; delete h.outErr; delete h.ch; chg = true; }   /* left the outbox: sent (event missed) or removed */
      if (!it) return;
      var st = it.lastCode === 'LOGIN' || (login && !driveOn()) ? 'login' : it.lastCode && it.lastCode !== 'OFFLINE' && it.lastCode !== 'THROTTLED' && it.lastCode !== 'LOGIN' ? 'failed' : 'queued';
      var e = it.lastCode === 'OFFLINE' || it.lastCode === 'LOGIN' ? '' : (it.lastError || '');
      if (h.out !== st || (h.outErr || '') !== e) { h.out = st; h.outErr = e; chg = true; }
    });
    if (chg) { saveHist(); if (VIEW.cur === 'hist') renderHist(); }
    if (VIEW.cur === 'set') renderDrive();
  }, function () { });
}
function onDriveEvent(ev) {
  if (!ev) return;
  if (ev.type === 'account') { driveUi(); if (driveOn()) S.hist.forEach(function (h) { if (h.ch === 'drive' && h.out === 'login') { h.out = 'queued'; } }); if (VIEW.cur === 'hist') renderHist(); return; }
  var x = ev.basketId ? histByBasket(ev.basketId) : null, now = Date.now(), chg = false;
  if (x && x.ch !== 'drive' && ev.type !== 'sent') x = null;
  if (ev.type === 'sent' && x && x.out && x.ch === 'drive') {
    x.sends.push(now); x.via = 'drive'; delete x.out; delete x.outErr; delete x.ch; chg = true;
    toast('Giden kutusundaki sepet OneDrive\'a gönderildi ✓ (' + fmtQty(x.lines.length) + ' satır)');
  } else if (ev.type === 'queued' && x) { if (!x.out) { x.out = 'queued'; x.q = x.q || now; chg = true; } }
  else if (ev.type === 'failed' && x) {
    var soft = ev.code === 'OFFLINE' || ev.code === 'THROTTLED' || ev.code === 'SERVER';
    x.out = soft ? 'queued' : 'failed'; x.outErr = ev.code === 'OFFLINE' ? '' : (ev.error || ev.code || ''); chg = true;
  } else if (ev.type === 'removed' && x && x.out) { delete x.out; delete x.outErr; delete x.ch; chg = true; }
  else if (ev.type === 'login') {
    S.hist.forEach(function (h) { if (h.ch === 'drive' && h.out) { h.out = 'login'; chg = true; } });
    if (now - DRV.loginToastAt > 600000 && S.hist.some(function (h) { return h.ch === 'drive' && h.out; })) {
      DRV.loginToastAt = now;
      toast('Giden kutusundaki sepetler için OneDrive\'a yeniden bağlanın.', true, { t: 'Bağlan', fn: function () { driveSignIn(null).then(function (ok) { if (ok) driveFlush(null, true); }); } });
    }
  }
  if (chg) { saveHist(); if (VIEW.cur === 'hist') renderHist(); if (VIEW.cur === 'set') renderDrive(); }
}

/* ===================== V6: pairing extras + PC addresses (spec section 10) ===================== */
function normLan(u) {
  u = String(u || '').trim(); if (!u) return '';
  if (!/^[a-z][a-z0-9+.-]*:\/\//i.test(u)) u = 'https://' + u;
  var x; try { x = new URL(u); } catch (e) { return ''; }
  if (x.protocol !== 'https:' && x.protocol !== 'http:') return '';
  if (!x.hostname || /\s/.test(u)) return '';
  return x.origin;
}
/* the PC itself serves this page under /saha/: catalog.js then also tries the page's own origin */
function servedByPc() { return /^https?:$/.test(location.protocol) && /^\/saha\//.test(location.pathname); }
function hasLan() { return !!(S.pair && ((S.pair.lan || []).length || servedByPc())); }
/* lan / cid from the pairing code: envelope.parsePairing returns them (V6); older envelope.js -> decode the JSON here */
function pairExtras(text, p) {
  var lan = p && Array.isArray(p.lan) ? p.lan : null, cid = p && typeof p.cid === 'string' ? p.cid : null, cloud = p && typeof p.cloud === 'string' ? p.cloud : null;
  if (lan == null || cid == null || cloud == null) {
    try {
      var t = String(text || ''); if (t.indexOf('OMSP1:') < 0 && /%3A/i.test(t)) t = decodeURIComponent(t);
      var i = t.indexOf('OMSP1:'), m = i < 0 ? null : /^[A-Za-z0-9\-_+\/=\s]*/.exec(t.slice(i + 6));
      var o = m ? JSON.parse(new TextDecoder().decode(b64uDec(m[0]))) : null;
      if (o) { if (lan == null && Array.isArray(o.lan)) lan = o.lan; if (cid == null && typeof o.cid === 'string') cid = o.cid; if (cloud == null && typeof o.cloud === 'string') cloud = o.cloud; }
    } catch (e) { }
  }
  var out = [];
  (lan || []).forEach(function (u) { var n = typeof u === 'string' ? normLan(u) : ''; if (n && out.indexOf(n) < 0) out.push(n); });
  return { lan: out.slice(0, 8), cid: String(cid || '').trim(), cloud: CLOUDS.hasOwnProperty(cloud) && cloud ? cloud : '' };
}
function afterPairChange() {
  mailSetup();
  setTimeout(catSilent, 500);
  if (VIEW.cur === 'set') renderSet();
}
var LANED = null;   /* working copy while editing */
function renderLan() {
  $('c_lan').hidden = !S.pair; if (!S.pair) return;
  if (!LANED) LANED = (S.pair.lan || []).slice();
  if (!LANED.length) LANED.push('');
  $('lanlist').innerHTML = LANED.map(function (u, i) {
    return '<div class="lanrow"><input type="url" inputmode="url" class="mono" data-i="' + i + '" value="' + esc(u) + '" aria-label="PC adresi ' + (i + 1) + '" placeholder="https://192.168.1.20:47803" ' +
      'autocapitalize="off" autocorrect="off" spellcheck="false" autocomplete="off" enterkeyhint="done" maxlength="200">' +
      '<button type="button" class="btn btn-icon" data-del="' + i + '" aria-label="Adres ' + (i + 1) + ' sil">' + ico('i-del') + '</button></div>';
  }).join('');
}
function lanRead() { LANED = [].map.call($('lanlist').querySelectorAll('input'), function (x) { return x.value; }); }
$('lanlist').addEventListener('click', function (e) {
  var b = e.target.closest('button[data-del]'); if (!b) return;
  lanRead(); LANED.splice(+b.getAttribute('data-del'), 1); renderLan(); setMsg('lanmsg', 'Kaydet\'e basınca geçerli olur.');
});
$('lanlist').addEventListener('input', function (e) { if (e.target.tagName === 'INPUT') e.target.removeAttribute('aria-invalid'); });
$('lanlist').addEventListener('keydown', function (e) { if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); $('blansave').click(); } });
$('blanadd').addEventListener('click', function () {
  lanRead(); if (LANED.length >= 8) { setMsg('lanmsg', 'En çok 8 adres girilebilir.', 'err'); return; }
  LANED.push(''); renderLan(); var ins = $('lanlist').querySelectorAll('input'); ins[ins.length - 1].focus();
});
$('blansave').addEventListener('click', function () {
  if (!S.pair) return;
  lanRead();
  var out = [], bad = -1, ins = $('lanlist').querySelectorAll('input');
  LANED.forEach(function (u, i) { if (!u.trim()) return; var n = normLan(u); if (!n) { if (bad < 0) bad = i; ins[i].setAttribute('aria-invalid', 'true'); return; } if (out.indexOf(n) < 0) out.push(n); });
  if (bad >= 0) { setMsg('lanmsg', 'Adres geçersiz: https://192.168.1.20:47803 biçiminde yazın.', 'err'); ins[bad].focus(); return; }
  S.pair.lan = out; savePair(); LANED = null; CATUI.cert = ''; CATUI.note = '';
  renderLan(); renderCatCard();
  setMsg('lanmsg', out.length ? 'Kaydedildi: ' + out.length + ' adres.' : 'Kaydedildi: adres yok (stok listesi alınamaz).', 'ok');
});

/* ===================== V6: direct send (OmsMail, sendmail.js + MSAL; spec section 13) ===================== */
var MAIL = { st: 'off', err: '', p: null, cid: '', started: false, loginToastAt: 0 };
function mailCid() { return (S.pair && S.pair.cid) || ''; }
function isGuid(s) { return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(s || '')); }
function redirectUri() { return location.origin + location.pathname.replace(/[^\/]*$/, ''); }
var MSALP = null;
/* msal-browser.min.js is a static script in index.html; this only covers an old cached page without it */
function loadMsal() {
  if (window.msal) return Promise.resolve();
  if (MSALP) return MSALP;
  MSALP = new Promise(function (res, rej) {
    var s = document.createElement('script'); s.src = './msal-browser.min.js';
    s.onload = function () { window.msal ? res() : (MSALP = null, rej(new Error('Microsoft giriş kitaplığı yüklenemedi'))); };
    s.onerror = function () { MSALP = null; rej(new Error('Microsoft giriş kitaplığı (msal-browser.min.js) yüklenemedi; internete bağlıyken sayfayı yenileyin')); };
    document.head.appendChild(s);
  });
  return MSALP;
}
/* configure + init once per Client ID; a changed ID reloads the page (MSAL keeps one app instance per page) */
function mailSetup() {
  var cid = mailCid();
  if (!window.OmsMail || !cid || DEMO || driveOk()) { if (!MAIL.p) MAIL.st = 'off'; return Promise.resolve(false); }
  if (MAIL.p && MAIL.cid === cid) return MAIL.p;
  if (MAIL.p && MAIL.cid !== cid) { flushSave(); location.reload(); return MAIL.p; }
  MAIL.cid = cid; MAIL.st = 'loading'; MAIL.err = '';
  MAIL.p = loadMsal().then(function () {
    OmsMail.configure({ clientId: cid, redirectUri: redirectUri() });
    return OmsMail.init();
  }).then(function () {
    MAIL.st = 'ready';
    if (!MAIL.started) {
      MAIL.started = true;
      try { OmsMail.onChange(onMailEvent); } catch (e) { }
      try { OmsMail.start(); } catch (e) { }
    }
    outboxSync();
    if (VIEW.cur === 'set') renderMs();
    return true;
  }, function (e) {
    MAIL.st = 'error'; MAIL.err = (e && (e.message || e.errorCode || e.name)) || String(e);
    if (VIEW.cur === 'set') renderMs();
    return false;
  });
  return MAIL.p;
}
function mailAccount() {
  if (MAIL.st !== 'ready' || !window.OmsMail) return null;
  try { return OmsMail.account() || null; } catch (e) { return null; }
}
async function mailSignIn(btn, msgId) {
  if (!mailCid()) { toast('Önce Ayarlar > Gelişmiş\'te uygulama (Client) kimliğini girin.', true); return false; }
  busy(btn, true);
  try {
    /* the MSAL client is built at page open, so signIn() opens the popup inside this click (user activation) */
    var ok = await mailSetup(); if (!ok) throw new Error(MAIL.err || 'Microsoft girişi hazırlanamadı');
    var r = await OmsMail.signIn();   /* never rejects: {ok, account, redirecting, code, error}; iOS / blocked popup -> full-page redirect */
    if (r && r.redirecting) { if (msgId) setMsg(msgId, 'Microsoft giriş sayfasına yönlendiriliyor…'); return false; }
    var a = mailAccount();
    if (r && r.ok && a) { toast('Giriş yapıldı: ' + (a.username || a.name)); outboxFlush(null, true); }
    else { var m0 = (r && r.code === 'CANCELLED') ? 'Giriş iptal edildi.' : 'Giriş yapılamadı: ' + ((r && (r.error || r.code)) || 'bilinmeyen hata'); if (msgId) setMsg(msgId, m0, 'err'); else toast(m0, true); }
    if (VIEW.cur === 'set') renderMs();
    if (VIEW.cur === 'hist') renderHist();
    return !!(r && r.ok && a);
  } catch (e) {
    var m = 'Giriş yapılamadı: ' + ((e && (e.message || e.errorCode || e.name)) || String(e));
    if (msgId) setMsg(msgId, m, 'err'); else toast(m, true);
    return false;
  } finally { busy(btn, false); }
}
function renderMs() {
  var card = $('c_ms'); card.hidden = !S.pair || driveOk(); if (card.hidden) return;   /* sub-section of Ayarlar > Gelişmiş; OneDrive replaces it */
  var cid = mailCid(), stat = $('msstat'), a = mailAccount();
  if (document.activeElement !== $('scid')) $('scid').value = cid;
  $('redir').textContent = redirectUri();
  var mod = !!window.OmsMail;
  $('bmsin').hidden = !!a; $('bmsout').hidden = !a;
  $('bmsin').disabled = !mod || !cid || MAIL.st === 'loading';
  if (!mod) stat.textContent = 'Doğrudan gönderim modülü (sendmail.js) yüklenemedi.';
  else if (!cid) stat.textContent = 'Kapalı: uygulama (Client) kimliği girilmedi.';
  else if (MAIL.st === 'loading') stat.textContent = 'Hazırlanıyor…';
  else if (MAIL.st === 'error') stat.textContent = 'Microsoft girişi hazırlanamadı: ' + MAIL.err;
  else if (a) stat.textContent = 'Giriş yapıldı: ' + (a.name && a.username && a.name !== a.username ? a.name + ' (' + a.username + ')' : (a.username || a.name));
  else stat.textContent = 'Giriş yapılmadı.';
}
$('bmsin').addEventListener('click', function () { setMsg('msmsg', ''); mailSignIn(this, 'msmsg'); });
$('bmsout').addEventListener('click', async function () {
  var n = S.hist.filter(function (x) { return x.out; }).length;
  if (!(await ask('Microsoft hesabından çıkılsın mı?', 'Sepetler artık doğrudan gönderilemez; e-posta uygulaması yolu çalışmaya devam eder.' +
    (n ? ' Giden kutusundaki ' + n + ' sepet yeniden giriş yapılana kadar bekler.' : ''), 'Çıkış yap', false))) return;
  var btn = this; busy(btn, true);
  try { await OmsMail.signOut(); toast('Çıkış yapıldı'); } catch (e) { toast('Çıkış yapılamadı: ' + ((e && e.message) || e), true); }
  busy(btn, false); renderMs();
});
$('bcidsave').addEventListener('click', function () {
  if (!S.pair) return;
  var v = $('scid').value.trim();
  if (v && !isGuid(v)) { $('scid').setAttribute('aria-invalid', 'true'); setMsg('msmsg', 'Kimlik 00000000-0000-0000-0000-000000000000 biçiminde olmalı (ya da boş).', 'err'); $('scid').focus(); return; }
  $('scid').removeAttribute('aria-invalid');
  if (v === mailCid()) { setMsg('msmsg', 'Değişiklik yok.', 'ok'); return; }
  var reload = !!MAIL.p;
  S.pair.cid = v; savePair();
  if (reload) { setMsg('msmsg', 'Kaydedildi. Sayfa yeniden yükleniyor…', 'ok'); flushSave(); setTimeout(function () { location.reload(); }, 600); return; }
  setMsg('msmsg', v ? 'Kaydedildi.' : 'Kaydedildi: doğrudan gönderim kapalı.', 'ok');
  mailSetup().then(function () { renderMs(); }); renderMs();
});
$('scid').addEventListener('input', function () { this.removeAttribute('aria-invalid'); });
$('scid').addEventListener('keydown', function (e) { if (e.key === 'Enter') { e.preventDefault(); $('bcidsave').click(); } });
$('bredircopy').addEventListener('click', function () { copyText(redirectUri(), 'Yönlendirme adresi kopyalandı'); });

/* send dialog: direct send button */
async function directSend(ctx, to, btn) {
  busy(btn, true); setMsg('s_dmsg', 'Gönderiliyor…');
  var r;
  try { r = await OmsMail.send({ to: to, armor: ctx.armor, fileName: ctx.name }); }
  catch (e) { r = { ok: false, code: 'ERROR', error: (e && e.message) || String(e) }; }
  r = r || { ok: false, code: 'ERROR' };
  busy(btn, false);
  if (r.ok) { markSent(ctx, 'direct'); return; }
  if (r.code === 'LOGIN') {
    setMsg('s_dmsg', 'Microsoft oturumunun süresi dolmuş: yeniden giriş yapın, sonra tekrar Doğrudan gönder\'e basın.', 'err');
    if (!$('s_relogin')) {
      var b = document.createElement('button'); b.type = 'button'; b.id = 's_relogin'; b.className = 'btn btn-sec'; b.innerHTML = ico('i-key') + '<span>Yeniden giriş yap</span>';
      $('s_dmsg').parentNode.insertBefore(b, $('s_dmsg').nextSibling);
      b.addEventListener('click', function () { mailSignIn(b, 's_dmsg').then(function (ok) { if (ok && !$('modal').hidden) renderSendModal(ctx); }); });
    }
    return;
  }
  /* offline or a send error: keep it in the outbox, OmsMail retries (online event, every 60 s, page open) */
  try { await OmsMail.enqueue({ basketId: ctx.b.basket, to: to, armor: ctx.armor, fileName: ctx.name }); }
  catch (e) { setMsg('s_dmsg', 'Gönderilemedi (' + (r.error || r.code) + ') ve giden kutusuna da alınamadı. E-posta uygulamasıyla gönderin.', 'err'); return; }
  markSent(ctx, 'queued', r.code === 'OFFLINE' ? '' : (r.error || r.code || ''));
}

/* outbox state <-> history entries */
function histByBasket(id) { for (var i = 0; i < S.hist.length; i++) if (S.hist[i].basket === id) return S.hist[i]; return null; }
function histState(x) {
  var ix = S.hist.indexOf(x);
  if (x.ch === 'drive' && x.out) {
    if (x.out === 'login') return { tag: '<span class="tag err">' + ico('i-key') + 'Bağlantı gerekli</span>',
      more: '<p class="small hst">Giden kutusunda bekliyor. OneDrive\'a yeniden bağlanınca kendiliğinden gider.</p>' +
        '<div class="acts"><button type="button" class="btn btn-pri" data-a="dlogin" data-i="' + ix + '">' + ico('i-cloudlink') + 'OneDrive\'a yeniden bağlan</button></div>' };
    var bad = x.out === 'failed';
    return { tag: bad ? '<span class="tag err">' + ico('i-err') + 'Gönderilemedi</span>' : '<span class="tag warn">' + ico('i-sync') + 'Giden kutusunda</span>',
      more: '<p class="small hst">' + (x.outErr ? esc(x.outErr) + ' ' : '') + (bad ? 'Giden kutusunda bekliyor; yeniden denenecek.' : 'İnternet gelince kendiliğinden OneDrive\'a gider.') + '</p>' +
        '<div class="acts"><button type="button" class="btn btn-sec" data-a="flush" data-i="' + ix + '">' + ico('i-resync') + 'Şimdi dene</button></div>' };
  }
  if (x.via === 'drive' && !x.out) return { tag: '<span class="tag ok">' + ico('i-cloudok') + 'OneDrive\'a gönderildi</span>', more: '' };
  if (x.out === 'queued') return { tag: '<span class="tag warn">' + ico('i-sync') + 'Giden kutusunda</span>',
    more: '<p class="small hst">İnternet gelince kendiliğinden gönderilir.' + (x.outErr ? ' Son hata: ' + esc(x.outErr) : '') + '</p>' +
      '<div class="acts"><button type="button" class="btn btn-sec" data-a="flush" data-i="' + S.hist.indexOf(x) + '">' + ico('i-resync') + 'Şimdi dene</button></div>' };
  if (x.out === 'failed') return { tag: '<span class="tag err">' + ico('i-err') + 'Gönderilemedi</span>',
    more: '<p class="small hst">' + esc(x.outErr || 'Gönderim başarısız.') + ' Giden kutusunda bekliyor; yeniden denenecek.</p>' +
      '<div class="acts"><button type="button" class="btn btn-sec" data-a="flush" data-i="' + S.hist.indexOf(x) + '">' + ico('i-resync') + 'Şimdi dene</button></div>' };
  if (x.out === 'login') return { tag: '<span class="tag err">' + ico('i-key') + 'Giriş gerekli</span>',
    more: '<p class="small hst">Giden kutusunda bekliyor: Microsoft hesabına yeniden giriş yapınca gider.</p>' +
      '<div class="acts"><button type="button" class="btn btn-pri" data-a="login" data-i="' + S.hist.indexOf(x) + '">' + ico('i-key') + 'Giriş yap</button></div>' };
  if (x.via === 'lan') return { tag: '<span class="tag ok">' + ico('i-lan') + 'PC aldı</span>', more: x.lanRes ? '<p class="small hst">' + esc(lanSummary(x.lanRes)) + '</p>' : '' };
  if (x.via === 'cloud') return { tag: '<span class="tag ok">' + ico('i-cloudok') + 'Buluta gönderildi</span>', more: '' };
  return { tag: '<span class="tag ok">' + ico('i-check') + 'Gönderildi</span>', more: '' };
}
function onMailEvent(ev) {
  if (!ev) return;
  var x = ev.basketId ? histByBasket(ev.basketId) : null, now = Date.now(), chg = false;
  if (x && x.ch === 'drive') x = null;
  if (ev.type === 'sent' && x) {
    x.sends.push(now); x.via = 'direct'; delete x.out; delete x.outErr; chg = true;
    toast('Giden kutusundaki sepet gönderildi ✓ (' + fmtQty(x.lines.length) + ' satır)');
  } else if (ev.type === 'queued' && x) { if (x.out !== 'queued') { x.out = 'queued'; x.q = x.q || now; chg = true; } }
  else if (ev.type === 'failed' && x) {
    /* channel-wide conditions (no internet, throttled) keep the item 'queued'; item errors show as 'failed' */
    x.out = (ev.code === 'OFFLINE' || ev.code === 'THROTTLED') ? 'queued' : 'failed'; x.outErr = ev.code === 'OFFLINE' ? '' : (ev.error || ev.code || ''); chg = true;
  }
  else if (ev.type === 'account') { if (VIEW.cur === 'set') renderMs(); if (mailAccount()) S.hist.forEach(function (h) { if (h.out === 'login' && h.ch !== 'drive') { h.out = 'queued'; chg = true; } }); }
  else if (ev.type === 'login') {
    S.hist.forEach(function (h) { if (h.out && h.ch !== 'drive' && (!ev.basketId || h.basket === ev.basketId)) { h.out = 'login'; chg = true; } });
    if (now - MAIL.loginToastAt > 600000) {
      MAIL.loginToastAt = now;
      toast('Giden kutusundaki sepetler için Microsoft hesabına yeniden giriş yapın.', true, { t: 'Giriş yap', fn: function () { mailSignIn(null).then(function (ok) { if (ok) outboxFlush(null, true); }); } });
    }
  }
  if (chg) { saveHist(); if (VIEW.cur === 'hist') renderHist(); }
}
/* reconcile history flags with the real outbox (after init / page open) */
function outboxSync() {
  if (MAIL.st !== 'ready' || !OmsMail.pending) return;
  Promise.resolve().then(function () { return OmsMail.pending(); }).then(function (list) {
    var ids = {}, chg = false;
    (list || []).forEach(function (it) { if (it && it.basketId) ids[it.basketId] = it; });
    S.hist.forEach(function (h) {
      if (h.ch === 'drive') return;   /* OneDrive outbox: driveSync() */
      var it = ids[h.basket];
      if (it && !h.out) { h.out = 'queued'; h.q = h.q || Date.now(); chg = true; }
      else if (!it && h.out) { delete h.out; delete h.outErr; chg = true; }   /* left the outbox: sent (event missed) or removed */
      var e = it && it.lastCode !== 'OFFLINE' ? (it.lastError || it.lastCode || '') : '';
      if (it && e && h.out !== 'login' && h.outErr !== e) { h.outErr = e; chg = true; }
    });
    if (chg) { saveHist(); if (VIEW.cur === 'hist') renderHist(); }
  }, function () { });
}
async function outboxFlush(btn, quiet) {
  if (MAIL.st !== 'ready') { var ok = await mailSetup(); if (!ok) { if (!quiet) toast('Doğrudan gönderim hazır değil: ' + (MAIL.err || 'Ayarlar\'da Microsoft hesabını kontrol edin.'), true); return; } }
  busy(btn, true);
  var r = null;
  try { r = await OmsMail.flush(); } catch (e) { if (!quiet) toast('Gönderilemedi: ' + ((e && e.message) || e), true); }
  busy(btn, false);
  if (r && !quiet) {
    var s = (r.sent || []).length, f = (r.failed || []).length;
    if (s && !f) toast(s + ' sepet gönderildi ✓');
    else if (f) { var f0 = r.failed[0]; toast(f0.code === 'OFFLINE' ? 'İnternet yok: sepetler giden kutusunda bekliyor.' : f0.code === 'LOGIN' ? 'Microsoft hesabına yeniden giriş gerekli.' : 'Gönderilemedi: ' + (f0.error || f0.code), true); }
    else toast('Giden kutusu boş');
  }
  outboxSync();
}
function outboxRemove(id) { try { if (window.OmsMail && OmsMail.remove) Promise.resolve(OmsMail.remove(id)).catch(function () { }); } catch (e) { } }

/* MSAL v5: the popup / iframe / redirect returns to this page with #code=&state=; sendmail.js already ran the redirect bridge.
   That window must not boot the app. A stale / foreign response clears the flag -> clean the URL and boot normally. */
function authCallbackPage() {
  document.body.classList.add('v-pair');
  $('main').innerHTML = '<div class="hero"><div class="card"><p id="authmsg" role="status" aria-live="polite"><span class="spin" aria-hidden="true"></span>Microsoft girişi tamamlanıyor…</p></div></div>';
  var p = null;
  try { p = window.OmsMail && OmsMail.isAuthCallback ? OmsMail.handleAuthCallback() : null; } catch (e) { }
  try { if (!p && window.OmsDrive && OmsDrive.handleAuthCallback) p = OmsDrive.handleAuthCallback(); } catch (e) { }
  Promise.resolve(p).then(function () {
    if (!(window.OmsMail && OmsMail.isAuthCallback) && !(window.OmsDrive && OmsDrive.isAuthCallback)) { try { history.replaceState(null, '', location.pathname); } catch (e) { } location.reload(); return; }
    setTimeout(function () {
      var m = $('authmsg'); if (!m) return;
      m.innerHTML = 'Giriş tamamlandı. Bu pencere kendiliğinden kapanmadıysa kapatın ya da <a href="' + esc(redirectUri()) + '">sayfaya dönün</a>.';
    }, 4000);
  });
}

/* ===================== app shell: service worker, install, network ===================== */
var APP = { ver: '', sw: 'hazırlanıyor…', bip: null, reg: null, reloading: false };
function swInit() {
  if (DEMO) { APP.sw = 'deneme modu'; return; }
  if (!('serviceWorker' in navigator)) { APP.sw = 'bu tarayıcıda yok'; return; }
  if (!window.isSecureContext || location.protocol === 'file:') { APP.sw = 'yok (https gerekir)'; return; }
  navigator.serviceWorker.register('./sw.js').then(function (reg) {
    APP.reg = reg; APP.sw = 'hazır';
    function watch(w) { if (!w) return; w.addEventListener('statechange', function () { if (w.state === 'installed' && navigator.serviceWorker.controller) $('upd').hidden = false; }); }
    if (reg.waiting && navigator.serviceWorker.controller) $('upd').hidden = false;
    watch(reg.installing);
    reg.addEventListener('updatefound', function () { watch(reg.installing); });
    readVer();
    if (VIEW.cur === 'set') renderSet();
  }, function (e) { APP.sw = 'kurulamadı (' + ((e && e.message) || e) + ')'; });
  navigator.serviceWorker.addEventListener('controllerchange', function () { if (APP.reloading) return; APP.reloading = true; location.reload(); });
  /* look for a new version when the page comes back to the foreground (at most hourly) */
  var lastChk = Date.now();
  document.addEventListener('visibilitychange', function () { if (!document.hidden && APP.reg && Date.now() - lastChk > 3600000) { lastChk = Date.now(); try { APP.reg.update(); } catch (e) { } } });
}
function readVer() {
  try { caches.keys().then(function (ks) { ks.forEach(function (k) { if (k.indexOf('oms-saha-') === 0) APP.ver = k.slice(9); }); if (VIEW.cur === 'set') renderSet(); }); } catch (e) { }
}
$('bupd').addEventListener('click', function () {
  var w = APP.reg && APP.reg.waiting;
  if (!w) { location.reload(); return; }
  flushSave(); APP.reloading = false; w.postMessage({ t: 'skip' });
  setTimeout(function () { if (!APP.reloading) location.reload(); }, 3000);
});
window.addEventListener('beforeinstallprompt', function (e) { e.preventDefault(); APP.bip = e; if (VIEW.cur === 'set') renderSet(); });
function netUi() { $('net').hidden = navigator.onLine !== false; }
window.addEventListener('online', netUi); window.addEventListener('offline', netUi);
function banners() {
  var b = $('ban'), m = '';
  if (location.protocol !== 'file:' && window.isSecureContext === false) m = 'Sayfa güvenli olmayan (http) bir adresten açıldı: şifreleme, kamera ve çevrimdışı çalışma kapalı. Sayfayı https:// ile başlayan adresten açın.';
  else if (!window.OmsEnvelope) m = 'Şifreleme modülü (envelope.js) yüklenemedi: sepet gönderilemez. İnternete bağlıyken sayfayı yenileyin.';
  else if (!(window.crypto && crypto.subtle)) m = 'Bu tarayıcı şifrelemeyi (WebCrypto) desteklemiyor: sepet gönderilemez. Güncel bir tarayıcı kullanın.';
  if (m) { b.innerHTML = ico('i-err') + '<div>' + esc(m) + '</div>'; b.hidden = false; } else b.hidden = true;
}

/* ===================== boot ===================== */
async function boot() {
  if ((window.OmsMail && OmsMail.isAuthCallback) || (window.OmsDrive && OmsDrive.isAuthCallback)) { authCallbackPage(); return; }   /* MSAL popup / redirect target only */
  var drvP = driveInit();   /* finishes a OneDrive redirect sign-in (iOS) before the URL is cleaned below */
  if (window.OmsMail && OmsMail.hasAuthResponse && OmsMail.hasAuthResponse(location)) await Promise.race([drvP, tick(3000)]);
  if (window.OmsMail && OmsMail.hasAuthResponse && OmsMail.hasAuthResponse(location)) { try { history.replaceState(null, '', location.pathname + location.search.replace(/[?&](code|state|error|error_description|client_info|session_state)=[^&]*/g, '').replace(/^&/, '?')); } catch (e) { } }   /* stale MSAL response in the URL */
  S.pair = jget(K.pair, null);
  if (S.pair && !(S.pair.dev && S.pair.sec)) S.pair = null;
  S.user = LS.get(K.user) || ''; S.defLoc = LS.get(K.loc) || '';
  S.fb.snd = LS.get(K.snd) !== '0'; S.fb.vib = LS.get(K.vib) !== '0';
  S.cloud = cloudNorm(LS.get(K.cloud));
  S.smode = LS.get(K.smode) === 'serial' ? 'serial' : 'manual';
  banners(); cloudInit(); netUi(); swInit();
  await DB.open();
  var b = await DB.get(K.basket), h = await DB.get(K.hist);
  if (b && b.basket && Array.isArray(b.lines)) S.basket = b.lines.length ? b : null;
  if (Array.isArray(h)) S.hist = h.filter(function (x) { return x && x.basket && Array.isArray(x.lines) && Array.isArray(x.sends) && (x.sends.length || x.out); });
  if (S.pair && !Array.isArray(S.pair.lan)) S.pair.lan = [];
  if (S.pair && typeof S.pair.cid !== 'string') S.pair.cid = '';
  camInit(); catInit(); mailSetup();
  if (pruneHist()) saveHist();
  if (DB.mode === 'mem' && !DEMO) toast('Bu tarayıcı veri saklamaya izin vermiyor (gizli sekme?). Sayfa kapanınca sepet kaybolur.', true);
  showView(!S.pair ? 'pair' : needWelcome() ? 'who' : 'scan');
  pairFromHash();
}

/* ===================== demo / self-test (headless render checks; never persists) ===================== */
function demoData() {
  var now = Date.now();
  S.pair = { v: 1, dev: 'd7k3m9qa', sec: 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA', name: 'Ahmet telefon', to: 'stok.depo@hotmail.com', pc: 'WIN-DEPO01', subj: SUBJ_DEFAULT, at: now - 86400000,
    lan: ['https://192.168.1.20:47803', 'https://10.0.0.5:47803'], cid: '11111111-2222-3333-4444-555555555555' };
  S.user = 'Ahmet'; S.defLoc = 'A3'; S.cloud = 'onedrive';
  var raws = ['[)>\x1e06\x1dPC1608X7R104K\x1d1PCL10B104KB8NNNC\x1dQ4000\x1d1T2245889\x1d9D2241\x1e\x04',
    '{pbn:PICK2210050012,on:SO2210050011,pc:C25804,pm:0603WAF1002T5E,qty:5000,mc:,cc:1,pdi:69432561,hp:0,wc:JS}',
    '[)>\x1e06\x1dP296-1395-1-ND\x1d1PSN74HC595DR\x1dKPO-88812\x1d1K74002211\x1d10K88213344\x1d11K1\x1d4LMY\x1dQ250\x1d11ZPICK\x1d12Z1123344\x1d13Z554433\x1d20Z0000000000\x1e\x04',
    '1PGRM188R71H104KA93D', 'XYZ-UNKNOWN-0001'];
  S.basket = { basket: uuid(), created: iso(now - 3600000), lines: raws.map(function (r, i) { return { id: uuid(), raw: r, fmt: i < 3 ? 'DATA_MATRIX' : 'CODE_128', qty: i === 2 ? 180 : null, loc: i % 2 ? 'B12' : 'A3', note: i === 2 ? 'açık makara' : '', ts: iso(now - 600000 + i * 1000) }; }) };
  S.hist = [{ basket: uuid(), created: iso(now - 90000000), lines: S.basket.lines.slice(0, 3), devName: 'Ahmet telefon', user: 'Ahmet', sends: [now - 86000000, now - 3600000] },
            { basket: uuid(), created: iso(now - 190000000), lines: S.basket.lines.slice(1, 2), devName: 'Ahmet telefon', user: 'Ahmet', sends: [now - 180000000], via: 'direct' },
            { basket: uuid(), created: iso(now - 50000000), lines: S.basket.lines.slice(2, 4), devName: 'Ahmet telefon', user: 'Ahmet', sends: [now - 40000000], via: 'cloud' },
            { basket: uuid(), created: iso(now - 7200000), lines: S.basket.lines.slice(3, 5), devName: 'Ahmet telefon', user: 'Ahmet', sends: [], out: 'queued', q: now - 1800000 },
            { basket: uuid(), created: iso(now - 9200000), lines: S.basket.lines.slice(0, 1), devName: 'Ahmet telefon', user: 'Ahmet', sends: [], out: 'failed', q: now - 5400000, outErr: 'Sunucu yanıtı 403 (izin yok)' }];
}
function demoCtx(used, more) {
  var now = Date.now();
  return { b: S.basket, hist: false, armor: 'OMSB1:d7k3m9qa:' + new Array(400).join('AbCd'), body: 'x', name: 'oms_sepet_' + stamp(now) + '.omsb', date: new Date(now),
    used: !!used, pv: used ? 'cloud' : '', moreOpen: !!more, n: S.basket.lines.length };
}
function overflow() { return document.documentElement.scrollWidth > window.innerWidth + 1 || document.body.scrollWidth > window.innerWidth + 1; }
async function selfTest() {
  demoData(); banners(); cloudInit(); await DB.open(); camInit(); catInit();
  var bad = [], fails = [];
  VIEWS.forEach(function (v) { showView(v); if (overflow()) bad.push(v); });
  showView('basket'); editLine(S.basket.lines[2].id); if (overflow()) bad.push('edit'); closeModal();
  editLine(S.basket.lines[0].id, { after: true }); if (overflow()) bad.push('edit-after'); closeModal();
  if (smode() !== 'manual' && !serOk()) fails.push('smode');
  if (!lanOk() || lanSummary({ outcome: 'Applied', applied: 12, issues: 1 }) !== 'PC aldı: 12 satır stoğa işlendi · 1 durum') fails.push('lan');
  renderSendModal(demoCtx(true));
  if (overflow()) bad.push('send'); closeModal();
  renderSendModal(demoCtx(false, true)); if (overflow()) bad.push('send-more'); closeModal();
  ['ios', 'android', 'desktop'].forEach(function (pl) { openGuide(null, pl); if (overflow()) bad.push('guide-' + pl); closeModal(); });
  showView('who'); if (overflow()) bad.push('welcome');
  ['ios', 'android', 'desktop'].forEach(function (pl) { ['onedrive', 'gdrive', 'icloud', ''].forEach(function (c) { var g = guideSteps(pl, c); if (!g || g.length < 3 || !g[0].title) fails.push('guide-' + pl + '-' + c); }); });
  if (localGuide('ios', 'icloud').length !== 3 || localGuide('desktop', '').length !== 3) fails.push('localGuide');
  if (cloudNorm('gdrive') !== 'gdrive' || cloudNorm('x') !== '' || cloudNorm(null) !== '') fails.push('cloudNorm');
  /* pure checks */
  if (b64uEnc(b64uDec('AQIDBP8')) !== 'AQIDBP8') fails.push('b64u');
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/.test(uuid())) fails.push('uuid');
  if (showRaw('a\x1db').indexOf('GS') < 0) fails.push('showRaw');
  if (!/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ$/.test(iso())) fails.push('iso');
  if (normLan('192.168.1.20:47803/saha/') !== 'https://192.168.1.20:47803' || normLan('ftp://x') !== '' || normLan('https://pc.local:47803/') !== 'https://pc.local:47803') fails.push('normLan');
  var ex = pairExtras('OMSP1:' + b64uEnc(new TextEncoder().encode(JSON.stringify({ v: 1, lan: ['https://192.168.1.20:47803/', 'x y', 'https://192.168.1.20:47803'], cid: '11111111-2222-3333-4444-555555555555' }))), {});
  if (ex.lan.length !== 1 || ex.cid !== '11111111-2222-3333-4444-555555555555' || ex.cloud !== '') fails.push('pairExtras');
  if (pairExtras('OMSP1:' + b64uEnc(new TextEncoder().encode(JSON.stringify({ v: 1, cloud: 'icloud' }))), {}).cloud !== 'icloud') fails.push('pairExtras-cloud');
  showView('scan');
  document.title = (fails.length ? 'SELFTEST FAIL ' + fails.join(',') : 'SELFTEST OK') + (bad.length ? ' OVERFLOW ' + bad.join(',') : '');
}
async function demo(v) {
  demoData(); banners(); cloudInit(); await DB.open(); camInit(); catInit();
  if (v === 'edit') { showView('basket'); editLine(S.basket.lines[0].id); return; }
  if (v === 'send') { showView('basket'); renderSendModal(demoCtx(false)); return; }
  if (v === 'sendmore') { showView('basket'); renderSendModal(demoCtx(false, true)); return; }
  if (v === 'sent') { showView('basket'); renderSendModal(demoCtx(true)); return; }
  if (v === 'guide') { showView('set'); openGuide(null); return; }
  if (v === 'sguide') { showView('basket'); openGuide(demoCtx(false)); return; }
  if (v === 'setadv') { showView('set'); $('advset').open = true; return; }
  if (v === 'pair') { S.pair = null; showView('pair'); return; }
  if (v === 'empty') { S.basket = null; showView('basket'); return; }
  if (v === 'who') { S.user = ''; S.cloud = ''; showView('who'); return; }
  showView(VIEWS.indexOf(v) >= 0 ? v : 'scan');
}

/* debugging / test hook */
window.OMSSAHA = { S: S, LANS: LANS, lanProbe: lanProbe, lanSend: lanSend, SER: SER, DRV: DRV, driveSend: driveSend, driveSync: driveSync, catRefresh: catRefresh, editLine: editLine, addLine: addLine, showView: showView, doPair: doPair, DEC: DEC, CAM: CAM, MAIL: MAIL, CATR: CATR, startSend: startSend,
  renderSendModal: renderSendModal, openGuide: openGuide, guideSteps: guideSteps, cloudPrimary: cloudPrimary, SEND: SEND, mailSetup: mailSetup, catInit: catInit, render: render, version: function () { return APP.ver; } };

if (DEMO === 'selftest') selfTest(); else if (DEMO) demo(DEMO); else boot();
})();
