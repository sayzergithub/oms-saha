/* OmsSerial — serial (continuous) reel scanning logic. Pure logic, no DOM. Spec V5_SAHA_SPEC §19, contract _PROMPTS_V8 D4.
 *
 *   var s = OmsSerial.create({ inBasket: function (raw) { return bool; }, now: function () { return ms; } });
 *   var d = s.offer(hits);   // hits = codes decoded from ONE camera frame: [{text, format}] (call it for every frame, also with [])
 *   d = { action: 'add' | 'ignore' | 'wait', hit: {text, format} | null, reason: string, parts: [text…] }
 *   s.count; s.reset(); s.flush() -> [decisions]  (on "Bitti": commits a waiting 1D reel / queued reels)
 *
 * A reel label usually carries one 2D code (Data Matrix / QR / PDF417 / Aztec) holding everything, plus several 1D codes
 * (MPN "1P…", quantity "Q…", lot "1T…") that repeat parts of it. Rules:
 *   - 2D in the frame -> take the 2D, swallow the 1D codes of the same frame (they belong to that label).
 *   - Only 1D -> wait wait2dMs for a 2D; if none comes, add the best 1D (MPN-like first) — all 1D seen meanwhile = same label.
 *     If no part-number-like 1D was seen yet (only qty / lot), keep waiting up to waitIdMs so the MPN code gets a chance.
 *   - 1D that is a substring of the last added / current 2D raw -> part of that label (swallowed).
 *   - a DIFFERENT 1D within sameLabelMs of the current label -> part of it; a DIFFERENT 2D is always a new reel.
 *   - a code once assigned to a label stays assigned (sticky) — holding the camera on a reel never creates a second entry.
 *   - raw already in the basket -> ignore 'inBasket' (once; then 'repeat' while it stays in view); same raw within repeatMs -> 'repeat'.
 * Reasons: add: 'new2d' | 'new1d' | 'queued';  ignore: 'empty' | 'repeat' | 'inBasket' | 'sameLabel' | 'substring';  wait: 'waiting2d'.
 */
(function (root, factory) {
  'use strict';
  var api = factory();
  if (typeof module === 'object' && module && module.exports) module.exports = api;
  if (root) root.OmsSerial = api;
})(typeof self !== 'undefined' ? self : (typeof globalThis !== 'undefined' ? globalThis : this), function () {
  'use strict';

  var DEFAULTS = {
    wait2dMs: 600,      // only-1D frame: how long to wait for the label's 2D code
    waitIdMs: 1500,     // ...extended up to this while the waiting 1D codes hold no part number yet (only qty / lot seen)
    sameLabelMs: 1500,  // a different 1D this soon after the current label (add / 2D sighting) = same label
    repeatMs: 3000,     // same raw again within this time = repeat (sliding: every sighting restarts it)
    minStripLen: 6      // substring test also tries the 1D value without its data identifier ("1P", "Q", "1T"…) if this long
  };

  // ZXing enum numbers (ZXing.BarcodeFormat) for callers passing raw numbers.
  var ZX_NUM = { 0: 'AZTEC', 1: 'CODABAR', 2: 'CODE39', 3: 'CODE93', 4: 'CODE128', 5: 'DATAMATRIX', 6: 'EAN8', 7: 'EAN13', 8: 'ITF',
    9: 'MAXICODE', 10: 'PDF417', 11: 'QRCODE', 12: 'RSS14', 13: 'RSSEXPANDED', 14: 'UPCA', 15: 'UPCE', 16: 'UPCEANEXTENSION' };
  var TWO_D = { DATAMATRIX: 1, QRCODE: 1, MICROQRCODE: 1, RMQRCODE: 1, PDF417: 1, MICROPDF417: 1, AZTEC: 1, MAXICODE: 1, HANXIN: 1, DOTCODE: 1 };
  var ONE_D = { CODE128: 1, CODE39: 1, CODE93: 1, CODABAR: 1, EAN8: 1, EAN13: 1, ITF: 1, UPCA: 1, UPCE: 1, UPCEANEXTENSION: 1,
    RSS14: 1, RSSEXPANDED: 1, GS1DATABAR: 1, GS1DATABAREXPANDED: 1, GS1DATABARLIMITED: 1, CODE11: 1, MSI: 1, PHARMACODE: 1 };

  // 'DATA_MATRIX' (ZXing), 'data_matrix' (BarcodeDetector), 5 (ZXing enum) -> 'DATAMATRIX'
  function normFormat(f) {
    if (typeof f === 'number') return ZX_NUM[f] || '';
    return String(f == null ? '' : f).toUpperCase().replace(/[^A-Z0-9]/g, '');
  }
  function looks2D(text) { return /^\[\)>|[\x1D\x1E]/.test(text) || /^\{[^{}]*:[^{}]*\}$/.test(text); }
  function is2D(format, text) {
    var f = normFormat(format);
    if (TWO_D[f]) return true;
    if (ONE_D[f]) return false;
    return looks2D(String(text == null ? '' : text));   // unknown format: structured payload = 2D
  }
  function keyOf(text) { return String(text == null ? '' : text).replace(/^\s+|\s+$/g, ''); }
  function cmpOf(key) { return key.replace(/[\x00-\x1F\x7F]/g, '').toUpperCase(); }

  // Is the 1D (cmp form) a part of the 2D payload (cmp form)?
  function partOf(c1, c2, minStrip) {
    if (!c1 || !c2) return false;
    if (c2.indexOf(c1) >= 0) return true;
    var bare = c1.replace(/^\d{0,2}[A-Z]/, '');
    return bare.length >= minStrip && bare !== c1 && c2.indexOf(bare) >= 0;
  }

  // Which 1D of an only-1D label identifies the reel best: MPN > plain part number > lot/date > quantity/number.
  function score1D(c) {
    if (/^1?P[A-Z0-9]/.test(c) && /[A-Z]/.test(c.slice(2)) && c.length >= 5) return 4;     // 1P<MPN> / P<part>
    if (/^Q\d+$/.test(c) || /^\d+$/.test(c)) return 0;                                    // quantity / bare number
    if (/^(1T|10D|9D|1D|4L|D)\w/.test(c)) return 1;                                       // lot / date code / origin
    if (/[A-Z]/.test(c) && /\d/.test(c) && c.length >= 5) return 3;                       // plain MPN (Code39 / Code128)
    return 2;
  }

  function create(opts) {
    opts = opts || {};
    var P = {};
    var k;
    for (k in DEFAULTS) if (Object.prototype.hasOwnProperty.call(DEFAULTS, k)) P[k] = DEFAULTS[k];
    if (opts.params) for (k in opts.params) if (Object.prototype.hasOwnProperty.call(opts.params, k)) P[k] = opts.params[k];
    var now = typeof opts.now === 'function' ? opts.now : function () { return Date.now(); };
    var inBasketFn = typeof opts.inBasket === 'function' ? opts.inBasket : null;

    var st;
    function fresh() {
      st = {
        count: 0,
        added: {},        // key -> time added in this session
        seen: {},         // key -> last sighting time
        assoc: {},        // key -> label id (1D codes swallowed as part of a label; sticky)
        labelSeq: 0,
        cur: null,        // current label {id, key, cmp, is2d, anchor}
        lastAdded2d: '',  // cmp of the last added 2D payload
        pending: null,    // only-1D wait {since, items:[{hit,key,cmp,order}]}
        backlog: []       // further new 2D codes from a multi-reel frame [{hit,key,cmp,parts}]
      };
    }
    fresh();

    function inBasket(hit, key) {
      if (st.added[key] !== undefined) return true;
      if (!inBasketFn) return false;
      try { return !!inBasketFn(hit.text); } catch (e) { return false; }
    }
    function recent(key, t) { var s = st.seen[key]; return s !== undefined && t - s < P.repeatMs; }
    function newLabel(key, cmp, two, t) { st.cur = { id: ++st.labelSeq, key: key, cmp: cmp, is2d: two, anchor: t }; return st.cur; }
    function focus(key, cmp, two, t) {          // a label's main code is in view: make it current / refresh its time anchor
      if (st.cur && st.cur.key === key) st.cur.anchor = t; else newLabel(key, cmp, two, t);
    }
    function assoc(key, label) { if (label && st.assoc[key] === undefined && st.added[key] === undefined) st.assoc[key] = label.id; }
    function result(action, hit, reason, parts) { return { action: action, hit: hit || null, reason: reason, parts: parts || [] }; }

    function normHits(hits) {
      var out = [], dup = {};
      if (!hits) return out;
      if (!Array.isArray(hits)) hits = [hits];
      for (var i = 0; i < hits.length; i++) {
        var h = hits[i];
        if (h == null) continue;
        if (typeof h !== 'object') h = { text: String(h), format: '' };
        var text = h.text != null ? h.text : h.rawValue;
        var key = keyOf(text);
        if (!key || dup[key]) continue;
        dup[key] = 1;
        var hit = { text: String(text), format: h.format == null ? '' : h.format };
        out.push({ hit: hit, key: key, cmp: cmpOf(key), two: is2D(hit.format, key) });
      }
      return out;
    }

    function addHit(x, reason, parts, t) {
      st.count++;
      st.added[x.key] = t;
      delete st.assoc[x.key];
      if (x.two) st.lastAdded2d = x.cmp;
      newLabel(x.key, x.cmp, x.two, t);
      return result('add', x.hit, reason, parts);
    }

    function pendingDue(t) {
      var p = st.pending;
      if (!p) return false;
      var age = t - p.since;
      if (age < P.wait2dMs) return false;
      if (age >= P.waitIdMs) return true;
      return p.items.some(function (it) { return score1D(it.cmp) >= 3; });
    }

    function commitPending(t) {
      var p = st.pending; st.pending = null;
      if (!p || !p.items.length) return null;
      var best = p.items[0], bs = -1;
      p.items.forEach(function (it) {
        var s = score1D(it.cmp) * 1000 + Math.min(it.cmp.length, 999);
        if (s > bs) { bs = s; best = it; }
      });
      var others = p.items.filter(function (it) { return it !== best; });
      if (inBasket(best.hit, best.key)) {        // basket changed while waiting
        var lab = newLabel(best.key, best.cmp, false, t);
        others.forEach(function (it) { assoc(it.key, lab); });
        return result('ignore', best.hit, 'inBasket');
      }
      var r = addHit(best, 'new1d', others.map(function (it) { return it.hit.text; }), t);
      others.forEach(function (it) { assoc(it.key, st.cur); });
      return r;
    }

    function popBacklog(t) {
      while (st.backlog.length) {
        var b = st.backlog.shift();
        if (inBasket(b.hit, b.key)) continue;
        var r = addHit(b, 'queued', b.parts, t);
        st.seen[b.key] = t;
        return r;
      }
      return null;
    }
    function inBacklog(key) {
      for (var i = 0; i < st.backlog.length; i++) if (st.backlog[i].key === key) return i;
      return -1;
    }

    function offer2D(twos, ones, t) {
      var addX = null, firstKnown = null, reasons = [];
      twos.forEach(function (x) {
        var bi = inBacklog(x.key);
        if (bi >= 0) { st.backlog.splice(bi, 1); x.isNew = true; }
        else if (recent(x.key, t)) { reasons.push('repeat'); x.known = true; }
        else if (inBasket(x.hit, x.key)) { reasons.push('inBasket'); x.known = true; }
        else x.isNew = true;
        if (x.isNew) {
          if (!addX) addX = x;
          else st.backlog.push({ hit: x.hit, key: x.key, cmp: x.cmp, two: true, parts: [] });
        } else if (!firstKnown) firstKnown = x;
      });
      var parts = ones.map(function (o) { return o.hit.text; });
      var pend = st.pending; st.pending = null;   // 1D codes waiting for a 2D belong to this frame's label
      if (pend) pend.items.forEach(function (it) { if (parts.indexOf(it.hit.text) < 0) parts.push(it.hit.text); });
      var r;
      if (addX) r = addHit(addX, 'new2d', parts, t);
      else focus(firstKnown.key, firstKnown.cmp, true, t);
      ones.forEach(function (o) { assoc(o.key, st.cur); });
      if (pend) pend.items.forEach(function (it) { assoc(it.key, st.cur); });
      twos.forEach(function (x) { st.seen[x.key] = t; });
      ones.forEach(function (o) { st.seen[o.key] = t; });
      if (r) return r;
      r = popBacklog(t);
      if (r) return r;
      return result('ignore', firstKnown.hit, reasons.indexOf('inBasket') >= 0 ? 'inBasket' : 'repeat');
    }

    function offer1D(ones, t) {
      var reasons = [], firstHit = null, inBasketHit = null, cur = st.cur;
      ones.forEach(function (o) {
        var why;
        if (st.assoc[o.key] !== undefined) why = 'sameLabel';
        else if (st.pending && st.pending.items.some(function (it) { return it.key === o.key; })) why = 'waiting2d';
        else if (recent(o.key, t)) { why = 'repeat'; if (st.added[o.key] !== undefined || (cur && cur.key === o.key)) focus(o.key, o.cmp, false, t); }
        else if (inBasket(o.hit, o.key)) { why = 'inBasket'; focus(o.key, o.cmp, false, t); if (!inBasketHit) inBasketHit = o.hit; }
        else if (partOf(o.cmp, cur && cur.is2d ? cur.cmp : '', P.minStripLen) || partOf(o.cmp, st.lastAdded2d, P.minStripLen)) {
          why = 'substring'; assoc(o.key, cur || { id: 0 });
        }
        else if (cur && t - cur.anchor <= P.sameLabelMs) { why = 'sameLabel'; assoc(o.key, cur); }
        else {
          why = 'waiting2d';
          if (!st.pending) st.pending = { since: t, items: [] };
          st.pending.items.push({ hit: o.hit, key: o.key, cmp: o.cmp, order: st.pending.items.length });
        }
        st.seen[o.key] = t;
        reasons.push(why);
        if (!firstHit) firstHit = o.hit;
      });
      if (pendingDue(t)) { var c = commitPending(t); if (c) return c; }
      var q = popBacklog(t);
      if (q) return q;
      if (inBasketHit) return result('ignore', inBasketHit, 'inBasket');
      if (st.pending) return result('wait', st.pending.items[0].hit, 'waiting2d');
      var order = ['repeat', 'substring', 'sameLabel'];
      for (var i = 0; i < order.length; i++) if (reasons.indexOf(order[i]) >= 0) return result('ignore', firstHit, order[i]);
      return result('ignore', firstHit, reasons[0] || 'empty');
    }

    var session = {
      params: P,
      offer: function (hits) {
        var t = +now();
        var list = normHits(hits);
        if (!list.length) {
          if (pendingDue(t)) { var c = commitPending(t); if (c) return c; }
          var q = popBacklog(t);
          if (q) return q;
          return st.pending ? result('wait', st.pending.items[0].hit, 'waiting2d') : result('ignore', null, 'empty');
        }
        var twos = list.filter(function (x) { return x.two; }), ones = list.filter(function (x) { return !x.two; });
        return twos.length ? offer2D(twos, ones, t) : offer1D(ones, t);
      },
      // "Bitti": commit everything still undecided (a waiting only-1D reel, queued reels). Returns the 'add' decisions.
      flush: function () {
        var t = +now(), out = [], d;
        if (st.pending) { d = commitPending(t); if (d && d.action === 'add') out.push(d); }
        while ((d = popBacklog(t))) out.push(d);
        return out;
      },
      reset: function () { fresh(); },
      // Forget a raw (e.g. the user deleted it from the list) so the same reel can be scanned again.
      forget: function (raw) {
        var key = keyOf(raw);
        if (st.added[key] !== undefined) { delete st.added[key]; st.count = Math.max(0, st.count - 1); }
        delete st.seen[key];
        if (st.cur && st.cur.key === key) st.cur = null;
      }
    };
    Object.defineProperty(session, 'count', { enumerable: true, get: function () { return st.count; } });
    Object.defineProperty(session, 'waiting', { enumerable: true, get: function () { return !!st.pending; } });
    return session;
  }

  return { DEFAULTS: DEFAULTS, create: create, is2D: is2D, normFormat: normFormat, version: 1 };
});
