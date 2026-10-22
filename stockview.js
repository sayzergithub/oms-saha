/*
 * OMS field mode - "Stok" tab: the PC's stock list on the phone (spec V8 §18).
 *
 *   var view = OmsStock.render(container, {
 *     getInfo:   () => OmsCatalog.info(),       // {count, at, pc, url, builtAt} | null   (default: OmsCatalog.info)
 *     getParts:  () => OmsCatalog.all(),        // parts[] {id,m,f,s,q,l,c,val,pkg,d}       (default: OmsCatalog.all)
 *     onRefresh: () => Promise<{ok, error, needCertUrl}>   // "Yenile" button; hidden when missing
 *     onPick:    (part) => …   // optional. With opts.pickLabel: a button with that label on the detail card calls it.
 *                                // Without pickLabel: called on every row tap; returning false skips the built-in detail card.
 *     stickyTop: CSS top of the sticky search bar (default: below the page's top bar), staleMs: "old list" warning age (24 h)
 *   });
 *   view.update()   re-read parts / info (also automatic on OmsCatalog.onChange)
 *   view.refresh()  same as tapping "Yenile";  view.setQuery(text);  view.destroy()
 *
 * Built for 20 000 parts on a slow phone: search strings and sort orders are computed once per catalog (lazily), filtering walks a
 * pre-sorted index (no sort per keystroke) and narrows the previous result while the user keeps typing, and only the rows on screen
 * exist in the DOM (fixed row height, a small recycled pool, positioned with transform). The page (window) scrolls, not an inner box.
 * The pure core (fold, natKey, buildIndex, filter, relTime) is exported for tests. UMD: browser global OmsStock, Node module.
 */
(function (root, factory) {
  'use strict';
  if (typeof module === 'object' && module.exports) module.exports = factory(root);
  else root.OmsStock = factory(root);
})(typeof globalThis !== 'undefined' ? globalThis : (typeof self !== 'undefined' ? self : this), function (root) {
  'use strict';

  var SORTS = ['mpn', 'qty', 'loc'];
  var PREF_KEY = 'oms.stock.view';
  var DAY = 24 * 3600 * 1000;

  // ================================================================== pure core
  // Turkish + English, case and accent insensitive: dotted / dotless i, cedilla, breve and umlaut letters, the micro sign.
  var FOLD_MAP = {
    'İ': 'i', 'I': 'i', 'ı': 'i', 'Ş': 's', 'ş': 's', 'Ç': 'c', 'ç': 'c', 'Ğ': 'g', 'ğ': 'g', 'Ö': 'o', 'ö': 'o',
    'Ü': 'u', 'ü': 'u', 'Â': 'a', 'â': 'a', 'Î': 'i', 'î': 'i', 'Û': 'u', 'û': 'u', 'µ': 'u', 'μ': 'u', 'Μ': 'u'
  };
  var FOLD_RE = /[İIıŞşÇçĞğÖöÜüÂâÎîÛûµμΜ]/g;
  function foldChar(c) { return FOLD_MAP[c]; }
  var NON_ASCII = /[^\x00-\x7f]/, ODD_SPACE = /\s\s|^\s|\s$|[^\S ]/;

  /** Search form of a text: folded, lower case, single spaces. */
  function fold(s) {
    if (typeof s !== 'string' || !s) return '';
    // Plain ASCII (most MPNs / packages) skips the Turkish map.
    var t = NON_ASCII.test(s) ? s.replace(FOLD_RE, foldChar).toLowerCase().replace(/\u0307/g, '') : s.toLowerCase();
    return ODD_SPACE.test(t) ? t.replace(/\s+/g, ' ').trim() : t;
  }

  /** "SN74-HC00 DR" -> "sn74hc00dr" (MPN typed without separators still matches). */
  function compact(folded) { return folded.replace(/[\s\-_.\/,]+/g, ''); }

  /** Natural sort key: digit runs are zero padded so "A2" < "A10" with a plain string compare (no slow localeCompare). */
  function natKey(s) { return natFolded(fold(s)); }
  function natFolded(f) {
    return f.replace(/\d+/g, function (d) {
      d = d.replace(/^0+(?=\d)/, '');
      return d.length >= 12 ? d : '000000000000'.substring(d.length) + d;
    });
  }

  function str(x) { return typeof x === 'string' ? x : (x == null ? '' : String(x)); }

  /**
   * Builds the search / sort index of a parts array. Cheap per part; the three sort orders are built on first use.
   * Returns {parts, n, hay[], cat[] (category number per part), cats:[{name, key, count}], order(sort) -> Int32Array}.
   */
  function buildIndex(parts) {
    parts = Array.isArray(parts) ? parts : [];
    var n = parts.length, hay = new Array(n), cat = new Int32Array(n), catByKey = {}, catByRaw = {}, cats = [], mk = new Array(n), i, p;
    for (i = 0; i < n; i++) {
      p = parts[i] || {};
      var m = fold(str(p.m)), s = fold(str(p.s));
      mk[i] = natFolded(m);
      // \u0001 separates fields so a query never matches across two of them. One fold for the rest of the record, and join()
      // instead of + so each string is flat: a rope would be flattened by the first indexOf, i.e. on the user's first key.
      hay[i] = [m, compact(m), s, compact(s),
        fold([str(p.val), str(p.pkg), str(p.f), str(p.l), str(p.d), str(p.c)].join('\u0001'))].join('\u0001');
      var cname = str(p.c), ci = catByRaw[cname];
      if (ci === undefined) {
        var ckey = fold(cname);
        ci = ckey ? catByKey[ckey] : -1;
        if (ci === undefined) { ci = catByKey[ckey] = cats.length; cats.push({ name: cname.trim(), key: ckey, count: 0 }); }
        catByRaw[cname] = ci;
      }
      cat[i] = ci;
      if (ci >= 0) cats[ci].count++;
    }
    // Chip order: most parts first, then name. Category numbers stay as assigned (index into cats).
    var chipOrder = cats.map(function (c, k) { return k; });
    chipOrder.sort(function (a, b) { return cats[b].count - cats[a].count || (cats[a].key < cats[b].key ? -1 : cats[a].key > cats[b].key ? 1 : 0); });

    var orders = {};
    function byMpn(a, b) { return mk[a] < mk[b] ? -1 : mk[a] > mk[b] ? 1 : a - b; }
    function order(sort) {
      if (SORTS.indexOf(sort) < 0) sort = 'mpn';
      if (orders[sort]) return orders[sort];
      var idx = new Array(n), k;
      for (k = 0; k < n; k++) idx[k] = k;
      if (sort === 'mpn') idx.sort(byMpn);
      else if (sort === 'qty') {
        var q = new Float64Array(n);
        for (k = 0; k < n; k++) { var v = parts[k] && parts[k].q; q[k] = typeof v === 'number' && isFinite(v) ? v : -Infinity; }
        idx.sort(function (a, b) { return q[b] - q[a] || byMpn(a, b); });   // most first, unknown quantity last
      } else {
        var lk = new Array(n);
        for (k = 0; k < n; k++) lk[k] = natKey(str(parts[k] && parts[k].l)) || '\uffff';   // no location last
        idx.sort(function (a, b) { return lk[a] < lk[b] ? -1 : lk[a] > lk[b] ? 1 : byMpn(a, b); });
      }
      return (orders[sort] = Int32Array.from ? Int32Array.from(idx) : new Int32Array(idx));
    }
    return { parts: parts, n: n, hay: hay, cat: cat, cats: cats, chipOrder: chipOrder, order: order };
  }

  /** Query -> folded tokens (all must match, any order). */
  function tokens(q) {
    var f = fold(str(q));
    return f ? f.split(' ') : [];
  }

  /**
   * Indexes (into ix.parts) matching query q and category number cat (-2 = all), in the order of sort.
   * prev (optional) = an earlier result {q, cat, sort, list}: when the new query only adds letters / words to it, only that
   * result is scanned (typing gets faster as the list shrinks).
   */
  function filter(ix, q, cat, sort, prev) {
    var t = tokens(q), all = cat === -2 || cat == null;
    if (!all && (cat < 0 || cat >= ix.cats.length)) all = true;
    var src = ix.order(sort), hay = ix.hay, cats = ix.cat;
    if (t.length === 0 && all) return { q: str(q), cat: -2, sort: sort, list: src, total: ix.n };
    if (prev && prev.sort === sort && prev.cat === (all ? -2 : cat) && narrows(prev.q, q)) src = prev.list;
    var out = new Int32Array(src.length), n = 0, i, j, k, h, ok;
    for (i = 0; i < src.length; i++) {
      k = src[i];
      if (!all && cats[k] !== cat) continue;
      h = hay[k]; ok = true;
      for (j = 0; j < t.length; j++) if (h.indexOf(t[j]) < 0) { ok = false; break; }
      if (ok) out[n++] = k;
    }
    return { q: str(q), cat: all ? -2 : cat, sort: sort, list: out.subarray(0, n), total: ix.n };
  }

  /** True when every match of newQ is also a match of oldQ (each old token is inside some new token). */
  function narrows(oldQ, newQ) {
    var a = tokens(oldQ), b = tokens(newQ);
    if (a.length === 0) return true;
    for (var i = 0; i < a.length; i++) {
      var hit = false;
      for (var j = 0; j < b.length; j++) if (b[j].indexOf(a[i]) >= 0) { hit = true; break; }
      if (!hit) return false;
    }
    return true;
  }

  /** Relative age in Turkish (exact texts: see tests/stockview.test.js). */
  function relTime(ms, nowMs) {
    if (!isFinite(ms)) return '';
    var d = (nowMs != null ? nowMs : Date.now()) - ms;
    if (d < 60000) return 'az önce';
    var m = Math.floor(d / 60000);
    if (m < 60) return m + ' dk önce';
    var h = Math.floor(m / 60);
    if (h < 24) return h + ' saat önce';
    var days = Math.floor(h / 24);
    if (days === 1) return 'dün';
    if (days < 45) return days + ' gün önce';
    return Math.round(days / 30) + ' ay önce';
  }

  /** Time the list was made: the PC time in the list (builtAt) when sane, else the local sync time. */
  function listTime(inf, nowMs) {
    if (!inf) return NaN;
    var now = nowMs != null ? nowMs : Date.now();
    var b = Date.parse(str(inf.builtAt)), a = Date.parse(str(inf.at));
    if (isFinite(b) && b <= now + 10 * 60000) return Math.min(b, now);
    return isFinite(a) ? Math.min(a, now) : NaN;
  }

  /** 12000 -> "12.000" (Turkish thousands, no Intl needed). */
  function fmtQty(q) {
    if (typeof q !== 'number' || !isFinite(q)) return '';
    var neg = q < 0, s = String(Math.round(Math.abs(q)));
    for (var i = s.length - 3; i > 0; i -= 3) s = s.substring(0, i) + '.' + s.substring(i);
    return (neg ? '-' : '') + s;
  }

  // ================================================================== DOM
  var CSS = [
    '.osv{--osv-row:64px;position:relative}',
    '.osv [hidden],.osv-modal [hidden]{display:none!important}',
    '.osv-bar{margin:0 0 4px}',
    '.osv-srch{position:sticky;z-index:6;background:var(--bg,#f9f9fb);padding:8px 0}',
    '.osv-srch svg{position:absolute;left:12px;top:20px;width:20px;height:20px;fill:var(--text-2,#60646c);pointer-events:none}',
    '.osv .osv-srch .osv-q.osv-q{display:block;width:100%;box-sizing:border-box;min-height:var(--h-ctl,44px);padding:10px 48px 10px 42px;font:inherit;font-size:16px;line-height:22px;',
    '  color:var(--text,#1c2024);background:var(--surface,#fff);border:1px solid var(--line-ctl,#8b8d98);border-radius:var(--r-md,8px);-webkit-appearance:none;appearance:none}',
    '.osv-q::-webkit-search-cancel-button{-webkit-appearance:none;appearance:none;display:none}',
    '.osv-q::placeholder{color:var(--text-3,#6b6f77);opacity:1}',
    '.osv-clr{position:absolute;right:0;top:8px;width:44px;height:44px;border:0;background:none;color:var(--text-2,#60646c);display:flex;align-items:center;justify-content:center;border-radius:var(--r-md,8px)}',
    '.osv-clr svg{width:20px;height:20px;fill:currentColor}',
    '.osv-chips{display:flex;overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none;margin:0 -4px;padding:0 4px;',
    '  -webkit-mask-image:linear-gradient(90deg,#000 calc(100% - 24px),transparent);mask-image:linear-gradient(90deg,#000 calc(100% - 24px),transparent)}',
    '.osv-chips::-webkit-scrollbar{display:none}',
    '.osv-chip{flex:none;display:inline-flex;align-items:center;min-height:44px;margin:0 6px 0 0;padding:0 14px;border-radius:999px;font:inherit;font-size:14px;font-weight:500;',
    '  white-space:nowrap;color:var(--text,#1c2024);background:var(--surface,#fff);border:1px solid var(--line,#d9d9e0)}',
    '.osv-chip .n{margin-left:6px;color:var(--text-2,#60646c);font-variant-numeric:tabular-nums;font-weight:400}',
    '.osv-chip[aria-pressed=true]{background:var(--accent-soft,#e6f4fe);border-color:var(--accent,#0b68c1);color:var(--text,#1c2024);font-weight:600}',
    '.osv-chip[aria-pressed=true]::before{content:"";display:inline-block;width:6px;height:10px;margin:-3px 8px 0 0;border:solid var(--accent-text,#0b68c1);border-width:0 2px 2px 0;transform:rotate(45deg)}',
    '.osv-tools{display:flex;align-items:center;margin-top:8px;min-height:44px}',
    '.osv-tools .osv-ref{flex:none;margin-left:12px}',
    '.osv-sort{display:flex;align-items:center;margin:8px 0 0;padding:2px;border:1px solid var(--line,#d9d9e0);border-radius:var(--r-md,8px);background:var(--surface-2,#f0f0f3)}',
    '.osv-sort>span{font-size:13px;color:var(--text-2,#60646c);padding:0 4px 0 8px}',
    '.osv-sort button{flex:1 1 0;min-height:44px;min-width:44px;padding:0 8px;border:1px solid transparent;border-radius:6px;background:none;font:inherit;font-size:15px;font-weight:500;color:var(--text-2,#60646c)}',
    '.osv-sort button[aria-pressed=true]{background:var(--surface,#fff);color:var(--text,#1c2024);border-color:var(--line,#d9d9e0);font-weight:600}',
    '.osv-meta{flex:1;min-width:0;margin:0;font-size:14px;line-height:20px;color:var(--text-2,#60646c);font-variant-numeric:tabular-nums}',
    '.osv-btn{display:inline-flex;align-items:center;justify-content:center;min-height:44px;min-width:44px;padding:0 14px;border-radius:var(--r-md,8px);font:inherit;font-size:15px;font-weight:600;',
    '  color:var(--text,#1c2024);background:var(--surface,#fff);border:1px solid var(--line-ctl,#8b8d98);margin-left:8px;text-decoration:none}',
    '.osv-btn svg{width:18px;height:18px;fill:currentColor;margin-right:6px}',
    '.osv-btn.pri{background:var(--accent,#0b68c1);border-color:var(--accent,#0b68c1);color:var(--on-accent,#fff)}',
    '.osv-btn[disabled]{opacity:.6;cursor:default}',
    '.osv-btn.busy svg{animation:osv-spin 1s linear infinite}',
    '@keyframes osv-spin{to{transform:rotate(360deg)}}',
    '.osv-note{margin:8px 0 0;padding:10px 12px;border-radius:var(--r-md,8px);font-size:15px;line-height:22px;border:1px solid var(--line,#d9d9e0);background:var(--surface,#fff);color:var(--text,#1c2024)}',
    '.osv-note.warn{background:var(--warn-soft,#fff7c2);border-color:var(--warn-line,#e9c162)}',
    '.osv-note.err{background:var(--err-soft,#feebec);border-color:var(--err-line,#f4a9aa)}',
    '.osv-note.ok{background:var(--ok-soft,#e6f6eb);border-color:var(--ok-line,#8eceaa)}',
    '.osv-note b{font-weight:600}',
    '.osv-note a{display:inline-block;margin-top:4px;min-height:24px;color:var(--accent-text,#0b68c1);font-weight:600}',
    '.osv-head{display:none}',
    '.osv.wide .osv-sort{display:inline-flex}.osv.wide .osv-sort button{flex:none;padding:0 12px}',
    '.osv-list{position:relative;contain:strict;overflow:hidden;background:var(--surface,#fff);border:1px solid var(--line,#d9d9e0);border-radius:var(--r-lg,12px)}',
    '.osv-row{position:absolute;left:0;right:0;top:0;height:var(--osv-row);box-sizing:border-box;display:grid;grid-template-columns:minmax(0,1fr) auto;',
    '  grid-template-rows:24px 20px;align-content:center;column-gap:12px;width:100%;margin:0;padding:0 14px;border:0;border-bottom:1px solid var(--line,#d9d9e0);',
    '  background:var(--surface,#fff);color:var(--text,#1c2024);font:inherit;text-align:left;will-change:transform;-webkit-tap-highlight-color:transparent}',
    '.osv-row>span{overflow:hidden;white-space:nowrap;text-overflow:ellipsis;min-width:0}',
    '.osv-row .m{font-weight:600;font-size:16px;line-height:24px}',
    '.osv-row .q{font-weight:600;font-size:16px;line-height:24px;text-align:right;font-variant-numeric:tabular-nums}',
    '.osv-row .q.z{color:var(--err,#ce2c31)}',
    '.osv-row .d{grid-column:1;font-size:14px;line-height:20px;color:var(--text-2,#60646c)}',
    '.osv-row .l{font-size:14px;line-height:20px;text-align:right;color:var(--text-2,#60646c);max-width:40vw}',
    '.osv-row .v,.osv-row .k,.osv-row .f{display:none}',
    '.osv-row:active{background:var(--pressed,#e0e1e6)}',
    '@media(hover:hover){.osv-row:hover{background:var(--hover,#e8e8ec)}}',
    '.osv-row:focus-visible{outline:2px solid var(--focus,#0b68c1);outline-offset:-3px}',
    /* wide: one line per part, column headers */
    '.osv.wide{--osv-row:48px}',
    '.osv.wide .osv-head{display:grid;grid-template-columns:minmax(0,2.7fr) minmax(0,1fr) minmax(0,1fr) minmax(0,1.4fr) minmax(0,1fr) 88px;column-gap:12px;',
    '  padding:8px 15px 6px;font-size:13px;line-height:18px;font-weight:600;color:var(--text-2,#60646c)}',
    '.osv.wide .osv-head span:last-child{text-align:right}',
    '.osv.wide .osv-row{grid-template-columns:minmax(0,2.7fr) minmax(0,1fr) minmax(0,1fr) minmax(0,1.4fr) minmax(0,1fr) 88px;grid-template-rows:auto;align-items:center}',
    '.osv.wide .osv-row .d{display:none}',
    '.osv.wide .osv-row .v,.osv.wide .osv-row .k,.osv.wide .osv-row .f{display:block;font-size:15px;line-height:22px}',
    '.osv.wide .osv-row .f{color:var(--text-2,#60646c)}',
    '.osv.wide .osv-row .m{order:1}.osv.wide .osv-row .v{order:2}.osv.wide .osv-row .k{order:3}.osv.wide .osv-row .f{order:4}',
    '.osv.wide .osv-row .l{order:5;text-align:left;font-size:15px;color:var(--text,#1c2024);max-width:none}.osv.wide .osv-row .q{order:6}',
    '.osv-empty{text-align:center;padding:40px 16px;color:var(--text-2,#60646c);font-size:15px;line-height:22px}',
    '.osv-empty b{display:block;color:var(--text,#1c2024);font-size:17px;margin:12px 0 4px}',
    '.osv-empty>svg{width:40px;height:40px;fill:var(--text-3,#6b6f77)}',
    '.osv.none .osv-srch,.osv.none .osv-bar,.osv.none .osv-head{display:none!important}',
    '.osv-empty .osv-btn{margin:16px 0 0}',
    /* detail card */
    '.osv-modal{position:fixed;top:0;right:0;bottom:0;left:0;z-index:95;background:var(--scrim,rgba(0,0,0,.32));display:flex;align-items:flex-start;justify-content:center;',
    '  overflow-y:auto;-webkit-overflow-scrolling:touch;overscroll-behavior:contain;padding:calc(16px + env(safe-area-inset-top,0px)) 16px calc(16px + env(safe-area-inset-bottom,0px))}',
    '.osv-box{margin:auto;width:100%;max-width:480px;box-sizing:border-box;background:var(--surface,#fff);color:var(--text,#1c2024);border:1px solid var(--line,#d9d9e0);',
    '  border-radius:var(--r-lg,12px);padding:20px;box-shadow:0 12px 32px rgba(0,0,0,.24)}',
    '.osv-box h3{margin:0 0 4px;font-size:20px;line-height:26px;font-weight:700;overflow-wrap:anywhere}',
    '.osv-box .sub{margin:0 0 12px;color:var(--text-2,#60646c);font-size:15px;line-height:22px;overflow-wrap:anywhere}',
    '.osv-big{display:flex;flex-wrap:wrap;margin:0 0 12px}',
    '.osv-big div{flex:1 1 120px;margin:0 8px 8px 0;padding:10px 12px;border-radius:var(--r-md,8px);background:var(--surface-2,#f0f0f3)}',
    '.osv-big small{display:block;font-size:13px;line-height:18px;color:var(--text-2,#60646c)}',
    '.osv-big strong{display:block;font-size:20px;line-height:28px;font-variant-numeric:tabular-nums;overflow-wrap:anywhere}',
    '.osv-kv{display:grid;grid-template-columns:auto minmax(0,1fr);column-gap:16px;row-gap:6px;margin:0 0 8px;font-size:15px;line-height:22px}',
    '.osv-kv dt{color:var(--text-2,#60646c)}.osv-kv dd{margin:0;overflow-wrap:anywhere;font-weight:500}',
    '.osv-acts{display:flex;flex-wrap:wrap;margin-top:16px}',
    '.osv-acts .osv-btn{flex:1 1 140px;margin:8px 8px 0 0}',
    '.osv-said{min-height:22px;font-size:14px;line-height:22px;color:var(--text-2,#60646c)}',
    '@media(prefers-reduced-motion:reduce){.osv-btn.busy svg{animation:none}}',
    '@media(prefers-contrast:more){.osv-row .d,.osv-row .l,.osv-meta{color:var(--text,#1c2024)}}'
  ].join('\n');

  var ICON = {
    search: 'M15.5 14h-.79l-.28-.27A6.47 6.47 0 0 0 16 9.5 6.5 6.5 0 1 0 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z',
    clear: 'M19 6.41 17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z',
    refresh: 'M17.65 6.35A7.96 7.96 0 0 0 12 4a8 8 0 1 0 7.73 10h-2.08A6 6 0 1 1 12 6c1.66 0 3.14.69 4.22 1.78L13 11h7V4l-2.35 2.35z',
    box: 'M20 2H4c-1 0-2 .9-2 2v3.01c0 .72.43 1.34 1 1.69V20c0 1.1 1.1 2 2 2h14c.9 0 2-.9 2-2V8.7c.57-.35 1-.97 1-1.69V4c0-1.1-1-2-2-2zm-1 18H5V9h14v11zm1-13H4V4h16v3zM9 12h6v2H9z'
  };

  function svg(name) {
    return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="' + ICON[name] + '"/></svg>';
  }

  function esc(s) {
    return str(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; });
  }

  function injectCss(doc) {
    if (doc.getElementById('oms-stock-css')) return;
    var st = doc.createElement('style');
    st.id = 'oms-stock-css';
    st.textContent = CSS;
    (doc.head || doc.documentElement).appendChild(st);
  }

  function loadPrefs() {
    try { var o = JSON.parse(root.localStorage.getItem(PREF_KEY) || '{}'); return o && typeof o === 'object' ? o : {}; } catch (e) { return {}; }
  }
  function savePrefs(o) {
    try { root.localStorage.setItem(PREF_KEY, JSON.stringify(o)); } catch (e) { /* private mode: not remembered */ }
  }

  function catalog() { return root.OmsCatalog || null; }

  /** Clipboard with the execCommand fallback (plain http pages have no navigator.clipboard). Resolves true/false. */
  function copyText(doc, text) {
    function legacy() {
      var ta = doc.createElement('textarea');
      ta.value = text; ta.setAttribute('readonly', ''); ta.style.cssText = 'position:fixed;top:0;left:0;opacity:0;font-size:16px';
      doc.body.appendChild(ta);
      var ok = false;
      try { ta.select(); ta.setSelectionRange(0, text.length); ok = doc.execCommand('copy'); } catch (e) { ok = false; }
      doc.body.removeChild(ta);
      return ok;
    }
    // The legacy path must run inside the tap, so try it first when the async API is missing.
    var nav = root.navigator;
    if (!nav || !nav.clipboard || !root.isSecureContext) return Promise.resolve(legacy());
    return nav.clipboard.writeText(text).then(function () { return true; }, function () { return legacy(); });
  }

  /**
   * Mounts the stock list into container (an empty element in the "Stok" view). Calling it again for the same container
   * returns the existing view after update().
   */
  function render(container, opts) {
    if (!container || !container.ownerDocument) throw new Error('OmsStock.render: container gerekli');
    if (container.__omsStock && !container.__omsStock.destroyed) { container.__omsStock.setOptions(opts); container.__omsStock.update(); return container.__omsStock; }
    var view = new View(container, opts || {});
    container.__omsStock = view;
    return view;
  }

  function View(container, opts) {
    var self = this, doc = container.ownerDocument;
    this.doc = doc; this.el = container; this.destroyed = false;
    this.o = {};
    this.setOptions(opts);
    var pr = loadPrefs();
    this.sort = SORTS.indexOf(pr.sort) >= 0 ? pr.sort : 'mpn';
    this.catKey = typeof pr.cat === 'string' ? pr.cat : '';
    this.q = '';
    this.ix = null; this.parts = null; this.res = null;
    this.pool = []; this.rowH = 64; this.wide = false; this.first = -1; this.last = -1;
    this.raf = 0; this.qTimer = 0; this.msgTimer = 0; this.busy = false;
    injectCss(doc);

    container.innerHTML =
      '<div class="osv">' +
      // The search row is a direct child of .osv so it stays stuck while the whole list scrolls by.
      '<div class="osv-srch">' + svg('search') +
      '<input class="osv-q" type="search" enterkeyhint="search" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false"' +
      ' aria-label="Stokta ara" placeholder="Ara: parça kodu, değer, üretici, konum">' +
      '<button type="button" class="osv-clr" aria-label="Aramayı temizle" hidden>' + svg('clear') + '</button></div>' +
      '<div class="osv-bar">' +
      '<div class="osv-chips" role="group" aria-label="Kategori"></div>' +
      '<div class="osv-tools"><p class="osv-meta"><span class="c" role="status" aria-live="polite"></span></p>' +
      '<button type="button" class="osv-btn osv-ref" hidden>' + svg('refresh') + '<span>Yenile</span></button></div>' +
      '<div class="osv-sort" role="group" aria-label="Sıralama"><span>Sırala</span>' +
      '<button type="button" data-s="mpn">Kod</button><button type="button" data-s="qty">Miktar</button><button type="button" data-s="loc">Konum</button></div>' +
      '</div>' +
      '<div class="osv-msg"></div>' +
      '<div class="osv-head" aria-hidden="true"><span>Parça kodu</span><span>Değer</span><span>Kılıf</span><span>Üretici</span><span>Konum</span><span>Miktar</span></div>' +
      '<div class="osv-empty" hidden></div>' +
      '<div class="osv-list" role="group" aria-label="Stok listesi"></div>' +
      '</div>';
    var $ = function (s) { return container.querySelector(s); };
    this.root = $('.osv'); this.stick = $('.osv-srch'); this.inp = $('.osv-q'); this.clr = $('.osv-clr'); this.chips = $('.osv-chips');
    this.sortEl = $('.osv-sort'); this.count = $('.osv-meta .c'); this.refBtn = $('.osv-ref'); this.msg = $('.osv-msg');
    this.empty = $('.osv-empty'); this.list = $('.osv-list');
    this.stick.style.top = typeof this.o.stickyTop === 'string' ? this.o.stickyTop : 'calc(var(--h-top, 56px) + var(--safe-t, 0px))';

    // ---- events (all removed in destroy)
    this.h = {
      input: function () {
        self.clr.hidden = !self.inp.value;
        // Big lists: let a burst of keys settle (a slow phone filters 20 000 rows in ~20 ms; this keeps typing smooth).
        clearTimeout(self.qTimer);
        self.qTimer = setTimeout(function () { self.setQuery(self.inp.value, true); }, self.ix && self.ix.n > 4000 ? 90 : 0);
      },
      key: function (e) {
        if (e.key === 'Enter') { e.preventDefault(); self.setQuery(self.inp.value, true); self.inp.blur(); }
        else if (e.key === 'Escape' && self.inp.value) { e.preventDefault(); self.setQuery('', false); }
      },
      clear: function () { self.setQuery('', false); self.inp.focus(); },
      chip: function (e) {
        var b = e.target.closest ? e.target.closest('.osv-chip') : null;
        if (!b) return;
        self.catKey = b.getAttribute('data-k') || '';
        self.save(); self.paintChips(); self.apply(true);
      },
      sort: function (e) {
        var b = e.target.closest ? e.target.closest('button[data-s]') : null;
        if (!b) return;
        self.sort = b.getAttribute('data-s');
        self.save(); self.paintSort(); self.apply(true);
      },
      ref: function () { self.refresh(); },
      empty: function (e) {
        var b = e.target.closest ? e.target.closest('button') : null;
        if (!b) return;
        if (b.getAttribute('data-a') === 'ref') self.refresh();
        else if (b.getAttribute('data-a') === 'clr') { self.catKey = ''; self.save(); self.paintChips(); self.setQuery('', false); }
      },
      row: function (e) {
        var b = e.target.closest ? e.target.closest('.osv-row') : null;
        if (b && b._i >= 0) self.open(b._i, b);
      },
      scroll: function () { self.schedule(); },
      resize: function () { self.measure(); self.schedule(); }
    };
    this.inp.addEventListener('input', this.h.input);
    this.inp.addEventListener('keydown', this.h.key);
    this.clr.addEventListener('click', this.h.clear);
    this.chips.addEventListener('click', this.h.chip);
    this.sortEl.addEventListener('click', this.h.sort);
    this.refBtn.addEventListener('click', this.h.ref);
    this.empty.addEventListener('click', this.h.empty);
    this.list.addEventListener('click', this.h.row);
    // Capture: any scrolling ancestor (window or an inner pane) moves the list.
    doc.addEventListener('scroll', this.h.scroll, { passive: true, capture: true });
    root.addEventListener('resize', this.h.resize);
    if (typeof root.ResizeObserver === 'function') {
      // Fires when the tab becomes visible (display:none -> shown) and on width changes.
      this.ro = new root.ResizeObserver(function () { self.measure(); self.schedule(); });
      this.ro.observe(this.root);
    }
    var C = catalog();
    if (!this.o.getParts && C && typeof C.onChange === 'function') this.unsub = C.onChange(function () { self.update(); });
    this.ageTimer = setInterval(function () { self.paintMeta(); }, 60000);

    this.paintSort();
    this.measure();
    this.update();
  }

  View.prototype.setOptions = function (opts) {
    var o = opts || {};
    for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) this.o[k] = o[k];
    if (this.refBtn) this.refBtn.hidden = typeof this.o.onRefresh !== 'function';
  };

  View.prototype.save = function () { savePrefs({ sort: this.sort, cat: this.catKey }); };

  View.prototype.getParts = function () {
    try {
      if (typeof this.o.getParts === 'function') return this.o.getParts() || [];
      var C = catalog();
      return C && typeof C.all === 'function' ? C.all() : [];
    } catch (e) { return []; }
  };

  View.prototype.getInfo = function () {
    try {
      if (typeof this.o.getInfo === 'function') return this.o.getInfo() || null;
      var C = catalog();
      return C && typeof C.info === 'function' ? C.info() : null;
    } catch (e) { return null; }
  };

  /** Re-reads the catalog (cheap when it did not change). */
  View.prototype.update = function () {
    if (this.destroyed) return;
    var parts = this.getParts();
    if (parts !== this.parts || !this.ix || this.ix.n !== parts.length) {
      this.parts = parts;
      this.ix = buildIndex(parts);
      this.res = null;
      this.paintChips();
    }
    this.refBtn.hidden = typeof this.o.onRefresh !== 'function';
    this.apply(false);
  };

  View.prototype.catNo = function () {
    if (!this.catKey || !this.ix) return -2;
    for (var i = 0; i < this.ix.cats.length; i++) if (this.ix.cats[i].key === this.catKey) return i;
    return -2;
  };

  View.prototype.setQuery = function (q, keepInput) {
    if (this.destroyed) return;
    clearTimeout(this.qTimer);
    this.q = str(q);
    if (!keepInput) this.inp.value = this.q;
    this.clr.hidden = !this.inp.value;
    this.apply(true);
  };

  /** Filters + repaints. toTop: bring the list start into view when the user changed the query / filter / sort. */
  View.prototype.apply = function (toTop) {
    if (!this.ix) return;
    var cat = this.catNo();
    if (cat === -2 && this.catKey && this.ix.n > 0) { this.catKey = ''; this.paintChips(); }
    this.res = filter(this.ix, this.q, cat, this.sort, this.res);
    var n = this.res.list.length;
    this.list.style.height = (n * this.rowH) + 'px';
    this.list.hidden = n === 0;
    this.paintEmpty();
    this.paintMeta();
    this.first = this.last = -1;
    for (var i = 0; i < this.pool.length; i++) this.pool[i]._i = -2;   // force repaint of every visible row
    if (toTop) {
      var r = this.list.getBoundingClientRect(), barH = this.stick.getBoundingClientRect().height, top = this.stickyPx();
      // Only scroll back when the list start is above the sticky search bar (the user had scrolled down).
      if (r.top < top + barH - 1) {
        var y = (root.pageYOffset || 0) + r.top - top - barH;
        try { root.scrollTo(0, Math.max(0, y)); } catch (e) { /* old browser */ }
      }
    }
    this.draw();
  };

  View.prototype.stickyPx = function () {
    var t = parseFloat(root.getComputedStyle ? root.getComputedStyle(this.stick).top : '0');
    return isFinite(t) ? t : 0;
  };

  View.prototype.measure = function () {
    var w = this.root.getBoundingClientRect().width;
    if (!w) return;
    var wide = w >= 600;
    if (wide !== this.wide) {
      this.wide = wide;
      this.root.classList.toggle('wide', wide);
    }
    var h = wide ? 48 : 64;
    if (h !== this.rowH) {
      this.rowH = h;
      if (this.res) this.list.style.height = (this.res.list.length * h) + 'px';
      this.first = this.last = -1;
    }
  };

  View.prototype.schedule = function () {
    if (this.raf || this.destroyed) return;
    var self = this;
    var raf = root.requestAnimationFrame || function (f) { return setTimeout(f, 16); };
    this.raf = raf(function () { self.raf = 0; self.draw(); });
  };

  /** Renders the rows inside the viewport (+ a few above / below). */
  View.prototype.draw = function () {
    if (this.destroyed || !this.res) return;
    var list = this.res.list, n = list.length;
    if (!n) { this.hidePool(0); return; }
    var r = this.list.getBoundingClientRect();
    if (!r.width && !r.height) return;   // hidden tab
    if ((r.width >= 600) !== this.wide) this.measure();
    var vh = root.innerHeight || this.doc.documentElement.clientHeight || 800, H = this.rowH;
    var over = 6;
    var first = Math.max(0, Math.floor(-r.top / H) - over);
    var last = Math.min(n - 1, Math.ceil((vh - r.top) / H) + over);
    if (last < first) { this.hidePool(0); return; }
    if (first === this.first && last === this.last) return;
    this.first = first; this.last = last;
    var need = last - first + 1, i;
    while (this.pool.length < need) this.pool.push(this.makeRow());
    // Keep a row on the same item when possible (no text rewrite while scrolling).
    var byItem = {}, free = [];
    for (i = 0; i < this.pool.length; i++) {
      var row = this.pool[i];
      if (row._pos >= first && row._pos <= last && row._i === list[row._pos]) byItem[row._pos] = row; else free.push(row);
    }
    for (var pos = first; pos <= last; pos++) {
      var el = byItem[pos];
      if (!el) { el = free.pop(); this.fill(el, list[pos], pos); }
      if (el._y !== pos * H) { el._y = pos * H; el.style.transform = 'translateY(' + el._y + 'px)'; }
      if (el.hidden) el.hidden = false;
    }
    for (i = 0; i < free.length; i++) { free[i].hidden = true; free[i]._pos = -1; }
  };

  View.prototype.hidePool = function () {
    for (var i = 0; i < this.pool.length; i++) { this.pool[i].hidden = true; this.pool[i]._pos = -1; }
    this.first = this.last = -1;
  };

  View.prototype.makeRow = function () {
    var d = this.doc, b = d.createElement('button');
    b.type = 'button'; b.className = 'osv-row';
    b.innerHTML = '<span class="m"></span><span class="q"></span><span class="d"></span><span class="l"></span>' +
      '<span class="v"></span><span class="k"></span><span class="f"></span>';
    var s = b.children;
    b._m = s[0]; b._q = s[1]; b._d = s[2]; b._l = s[3]; b._v = s[4]; b._k = s[5]; b._f = s[6];
    b._i = -2; b._pos = -1; b._y = -1; b.hidden = true;
    this.list.appendChild(b);
    return b;
  };

  View.prototype.fill = function (b, i, pos) {
    var p = this.ix.parts[i] || {};
    b._i = i; b._pos = pos;
    var m = str(p.m) || '(kodsuz parça)', val = str(p.val), pkg = str(p.pkg), f = str(p.f), l = str(p.l);
    var hasQ = typeof p.q === 'number' && isFinite(p.q);
    b._m.textContent = m;
    b._q.textContent = hasQ ? fmtQty(p.q) : '–';
    b._q.className = 'q' + (hasQ && p.q <= 0 ? ' z' : '');
    var det = [];
    if (val) det.push(val);
    if (pkg) det.push(pkg);
    if (f) det.push(f);
    b._d.textContent = det.length ? det.join(' · ') : (str(p.c) || ' ');
    b._l.textContent = l || 'konum yok';
    b._v.textContent = val; b._k.textContent = pkg; b._f.textContent = f;
    b.setAttribute('aria-label', m + (val ? ', ' + val : '') + (pkg ? ', ' + pkg : '') + (f ? ', ' + f : '') +
      ', ' + (hasQ ? fmtQty(p.q) + ' adet' : 'miktar bilinmiyor') + (l ? ', konum ' + l : ', konum yok'));
  };

  View.prototype.paintChips = function () {
    var ix = this.ix, html = '', cur = this.catKey;
    if (!ix || ix.cats.length < 2) { this.chips.innerHTML = ''; this.chips.hidden = true; return; }
    this.chips.hidden = false;
    var known = false;
    for (var i = 0; i < ix.cats.length; i++) if (ix.cats[i].key === cur) known = true;
    if (!known) cur = this.catKey = '';
    html += '<button type="button" class="osv-chip" data-k="" aria-pressed="' + (!cur) + '">Tümü<span class="n">' + fmtQty(ix.n) + '</span></button>';
    for (var j = 0; j < ix.chipOrder.length; j++) {
      var c = ix.cats[ix.chipOrder[j]];
      html += '<button type="button" class="osv-chip" data-k="' + esc(c.key) + '" aria-pressed="' + (c.key === cur) + '">' + esc(c.name) +
        '<span class="n">' + fmtQty(c.count) + '</span></button>';
    }
    this.chips.innerHTML = html;
    var sel = this.chips.querySelector('[aria-pressed=true]');
    if (sel && sel.offsetLeft + sel.offsetWidth > this.chips.clientWidth) this.chips.scrollLeft = sel.offsetLeft - 8;
  };

  View.prototype.paintSort = function () {
    var bs = this.sortEl.querySelectorAll('button');
    for (var i = 0; i < bs.length; i++) bs[i].setAttribute('aria-pressed', String(bs[i].getAttribute('data-s') === this.sort));
  };

  View.prototype.paintMeta = function () {
    if (this.destroyed || !this.ix) return;
    var total = this.ix.n, shown = this.res ? this.res.list.length : total, inf = this.getInfo();
    var t = total ? (shown === total ? fmtQty(total) + ' parça' : fmtQty(shown) + ' / ' + fmtQty(total) + ' parça') : '';
    var when = listTime(inf);
    if (total && isFinite(when)) t += ' · son güncelleme ' + relTime(when);
    if (this.count.textContent !== t) this.count.textContent = t;
    // Old list: say so above the list (not only by colour).
    var stale = total > 0 && isFinite(when) && Date.now() - when > (this.o.staleMs > 0 ? this.o.staleMs : DAY);
    var old = this.msg.querySelector('[data-stale]');
    if (stale && !this.busy && !this.msg.querySelector('[data-res]')) {
      var txt = '<b>Stok listesi eski</b> (' + esc(relTime(when)) + ' güncellendi).' +
        (typeof this.o.onRefresh === 'function' ? ' Güncel miktarlar için Yenile\'ye dokunun.' : '');
      if (!old) this.msg.innerHTML = '<p class="osv-note warn" data-stale role="status">' + txt + '</p>';
      else if (old.innerHTML !== txt) old.innerHTML = txt;
    } else if (old && !stale) this.msg.innerHTML = '';
  };

  View.prototype.paintEmpty = function () {
    var e = this.empty, total = this.ix.n, n = this.res.list.length, canRef = typeof this.o.onRefresh === 'function';
    this.root.classList.toggle('none', total === 0);
    if (n > 0) { if (!e.hidden) { e.hidden = true; e.innerHTML = ''; } return; }
    var html;
    if (total === 0) {
      html = svg('box') + '<b>Stok listesi yok</b><span>' +
        (canRef ? 'Listeyi almak için Yenile\'ye dokunun.' : 'Liste PC\'den alınınca burada görünür.') + '</span>' +
        (canRef ? '<br><button type="button" class="osv-btn pri" data-a="ref"' + (this.busy ? ' disabled' : '') + '>' + svg('refresh') + '<span>Yenile</span></button>' : '');
    } else {
      var q = this.q.trim();
      html = svg('search') + '<b>Sonuç yok</b><span>' +
        (q ? '“' + esc(q.length > 40 ? q.substring(0, 40) + '…' : q) + '” ' : 'Bu kategoride ') + 'eşleşen parça bulunamadı.</span>' +
        '<br><button type="button" class="osv-btn" data-a="clr">Aramayı temizle</button>';
    }
    e.innerHTML = html;
    e.hidden = false;
  };

  /** Runs opts.onRefresh with a busy button; the button always comes back (timeout guard). */
  View.prototype.refresh = function () {
    var self = this, fn = this.o.onRefresh;
    if (typeof fn !== 'function' || this.busy) return Promise.resolve(null);
    this.setBusy(true);
    this.say('', '');
    var guard;
    var timeout = new Promise(function (resolve) {
      guard = setTimeout(function () { resolve({ ok: false, error: 'Yanıt gelmedi. İnternet bağlantınızı kontrol edip tekrar deneyin.' }); }, 45000);
    });
    var run;
    try { run = Promise.resolve(fn()); } catch (e) { run = Promise.reject(e); }
    return Promise.race([run, timeout]).then(function (r) { return r; }, function (e) {
      return { ok: false, error: (e && e.message) || 'Stok listesi alınamadı.' };
    }).then(function (r) {
      clearTimeout(guard);
      if (self.destroyed) return r;
      self.setBusy(false);
      self.update();
      if (r && r.ok === false) {
        var link = r.needCertUrl ? '<br><a href="' + esc(r.needCertUrl) + '" target="_blank" rel="noopener">PC sertifikasını onayla</a>' : '';
        self.say('err', esc(r.error || 'Stok listesi alınamadı.') + link);
      } else if (r && r.older) {
        self.say('ok', 'Elinizdeki liste zaten en yenisi.');
      } else {
        var n = self.ix ? self.ix.n : 0;
        self.say('ok', 'Stok listesi güncellendi: ' + fmtQty(n) + ' parça.', 5000);
      }
      return r;
    });
  };

  View.prototype.setBusy = function (on) {
    this.busy = !!on;
    var bs = this.root.querySelectorAll('.osv-ref,[data-a=ref]');
    for (var i = 0; i < bs.length; i++) {
      bs[i].disabled = this.busy;
      bs[i].classList.toggle('busy', this.busy);
      var sp = bs[i].querySelector('span');
      if (sp) sp.textContent = this.busy ? 'Yenileniyor…' : 'Yenile';
    }
  };

  /** Message above the list. kind: ok | err | warn | '' (clear). html is already escaped. */
  View.prototype.say = function (kind, html, ms) {
    var self = this;
    clearTimeout(this.msgTimer);
    if (!kind) { this.msg.innerHTML = ''; this.paintMeta(); return; }
    this.msg.innerHTML = '<p class="osv-note ' + kind + '" data-res role="' + (kind === 'err' ? 'alert' : 'status') + '">' + html + '</p>';
    if (ms) this.msgTimer = setTimeout(function () { if (!self.destroyed) self.say('', ''); }, ms);
  };

  // ---------------------------------------------------------------- detail card
  View.prototype.open = function (i, fromEl) {
    var p = this.ix && this.ix.parts[i];
    if (!p) return;
    var pick = this.o.onPick;
    // onPick returning false = the page shows its own detail.
    if (typeof pick === 'function' && !this.o.pickLabel) {
      var r;
      try { r = pick(p); } catch (e) { r = undefined; }
      if (r === false) return;
    }
    this.closeCard();
    var self = this, d = this.doc, C = catalog();
    var hasQ = typeof p.q === 'number' && isFinite(p.q);
    var same = null;
    try { same = C && typeof C.lookupMpn === 'function' && !this.o.getParts ? C.lookupMpn(p.m) : null; } catch (e) { same = null; }
    var rows = [['Değer', p.val], ['Kılıf', p.pkg], ['Tedarikçi kodu', p.s], ['Kategori', p.c], ['Açıklama', p.d]];
    var kv = '';
    for (var k = 0; k < rows.length; k++) if (str(rows[k][1]).trim()) kv += '<dt>' + rows[k][0] + '</dt><dd>' + esc(rows[k][1]) + '</dd>';
    var other = '';
    if (same && same.parts && same.parts.length > 1) {
      other = '<p class="osv-note">Bu parça kodu stokta <b>' + same.parts.length + ' kayıtta</b> var: toplam ' +
        (same.q != null ? '<b>' + fmtQty(same.q) + ' adet</b>' : 'miktar bilinmiyor') + (same.l ? ' · ' + esc(same.l) : '') + '</p>';
    }
    var m = d.createElement('div');
    m.className = 'osv-modal';
    m.innerHTML = '<div class="osv-box" role="dialog" aria-modal="true" aria-labelledby="osv-dt">' +
      '<h3 id="osv-dt">' + esc(p.m || '(kodsuz parça)') + '</h3>' +
      (p.f ? '<p class="sub">' + esc(p.f) + '</p>' : '') +
      '<div class="osv-big"><div><small>Miktar</small><strong>' + (hasQ ? fmtQty(p.q) + ' adet' : 'bilinmiyor') + '</strong></div>' +
      '<div><small>Konum</small><strong>' + (str(p.l).trim() ? esc(p.l) : 'yok') + '</strong></div></div>' +
      (kv ? '<dl class="osv-kv">' + kv + '</dl>' : '') + other +
      '<div class="osv-acts">' +
      (p.m ? '<button type="button" class="osv-btn" data-a="copy">Parça kodunu kopyala</button>' : '') +
      (typeof pick === 'function' && this.o.pickLabel ? '<button type="button" class="osv-btn" data-a="pick">' + esc(this.o.pickLabel) + '</button>' : '') +
      '<button type="button" class="osv-btn pri" data-a="close">Kapat</button></div>' +
      '<p class="osv-said" role="status" aria-live="polite"></p></div>';
    d.body.appendChild(m);
    this.card = { el: m, from: fromEl, item: i };
    this.cardKey = function (e) {
      if (e.key === 'Escape') { e.preventDefault(); self.closeCard(); return; }
      if (e.key !== 'Tab') return;
      var f = m.querySelectorAll('button'), a = f[0], z = f[f.length - 1];   // keep focus inside the card
      if (e.shiftKey && d.activeElement === a) { e.preventDefault(); z.focus(); }
      else if (!e.shiftKey && d.activeElement === z) { e.preventDefault(); a.focus(); }
    };
    d.addEventListener('keydown', this.cardKey, true);
    m.addEventListener('click', function (e) {
      if (e.target === m) { self.closeCard(); return; }
      var b = e.target.closest ? e.target.closest('button[data-a]') : null;
      if (!b) return;
      var a = b.getAttribute('data-a');
      if (a === 'close') self.closeCard();
      else if (a === 'pick') { self.closeCard(); try { pick(p); } catch (err) { /* page bug */ } }
      else if (a === 'copy') {
        var said = m.querySelector('.osv-said');
        copyText(d, str(p.m)).then(function (ok) {
          said.textContent = ok ? 'Kopyalandı: ' + p.m : 'Kopyalanamadı. Kodu basılı tutup seçin.';
        });
      }
    });
    var close = m.querySelector('[data-a=close]');
    try { close.focus({ preventScroll: true }); } catch (e) { close.focus(); }
  };

  View.prototype.closeCard = function () {
    if (!this.card) return;
    var c = this.card;
    this.card = null;
    this.doc.removeEventListener('keydown', this.cardKey, true);
    if (c.el.parentNode) c.el.parentNode.removeChild(c.el);
    // Back to the row it came from when that row still shows the same part, else to the search box.
    var to = c.from && c.from._i === c.item && !c.from.hidden ? c.from : this.inp;
    try { to.focus({ preventScroll: true }); } catch (e) { /* ignore */ }
  };

  View.prototype.destroy = function () {
    if (this.destroyed) return;
    this.destroyed = true;
    this.closeCard();
    clearTimeout(this.qTimer); clearTimeout(this.msgTimer); clearInterval(this.ageTimer);
    if (this.raf && root.cancelAnimationFrame) root.cancelAnimationFrame(this.raf);
    this.doc.removeEventListener('scroll', this.h.scroll, { passive: true, capture: true });
    root.removeEventListener('resize', this.h.resize);
    if (this.ro) this.ro.disconnect();
    if (this.unsub) this.unsub();
    this.el.innerHTML = '';
    if (this.el.__omsStock === this) this.el.__omsStock = null;
  };

  return {
    render: render,
    // pure core (tests, other pages)
    fold: fold,
    natKey: natKey,
    buildIndex: buildIndex,
    filter: filter,
    narrows: narrows,
    relTime: relTime,
    listTime: listTime,
    formatQty: fmtQty
  };
});
