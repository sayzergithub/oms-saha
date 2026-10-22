/*
 * OMS field mode - lightweight label preview (spec V5 §8). Port of the essentials of Labels/*.cs
 * (LabelNormalizer, Iso15434, DiSingle, Lcsc, Tme, Gs1, BraceKv, GenericKv, Universal, Raw fallback).
 * Purpose: show the operator what the reel is while scanning offline. The authoritative parse happens on the PC
 * (learned templates, component engine and date normalisation are PC-only).
 * UMD, no dependencies. Browser: global OmsLabel. Node: module.exports.
 *
 *   OmsLabel.preview(raw [, fmt]) -> { mpn, qty, mfr, supplierPn, lot, dc, supplier, parser, reelId }   (any field may be null)
 */
(function (root, factory) {
  'use strict';
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.OmsLabel = factory();
})(typeof globalThis !== 'undefined' ? globalThis : (typeof self !== 'undefined' ? self : this), function () {
  'use strict';

  var GS = '\x1D', RS = '\x1E', EOT = '\x04';

  // ------------------------------------------------------------------ helpers
  function isBlank(s) { return s == null || String(s).trim().length === 0; }
  function trimChars(s, chars) {
    var a = 0, b = s.length;
    while (a < b && chars.indexOf(s[a]) >= 0) a++;
    while (b > a && chars.indexOf(s[b - 1]) >= 0) b--;
    return s.substring(a, b);
  }
  function trimEndChars(s, chars) {
    var b = s.length;
    while (b > 0 && chars.indexOf(s[b - 1]) >= 0) b--;
    return s.substring(0, b);
  }
  function splitNonEmpty(s, sepChars) {
    var out = [], cur = '';
    for (var i = 0; i < s.length; i++) {
      if (sepChars.indexOf(s[i]) >= 0) { if (cur.length) out.push(cur); cur = ''; } else cur += s[i];
    }
    if (cur.length) out.push(cur);
    return out;
  }
  /** First argument that is not null/undefined (C# ??). */
  function coalesce() {
    for (var i = 0; i < arguments.length; i++) if (arguments[i] != null) return arguments[i];
    return null;
  }
  function normKey(k) { return String(k || '').toUpperCase().replace(/[ _\-.]/g, ''); }

  /** Unicode-aware regex with an ASCII + CJK fallback for engines without \p{L} (pre-2018). */
  function uni(src, fallback, flags) {
    try { return new RegExp(src, 'u' + (flags || '')); } catch (e) { return new RegExp(fallback, flags || ''); }
  }

  // ------------------------------------------------------------------ draft (PartDraft subset)
  function Draft(aim) {
    this.fields = {};          // case-insensitive like the C# Fields dictionary (keys upper-cased)
    this.aim = aim;
    this.mpn = null; this.qty = null; this.mfr = null; this.supplierPn = null;
    this.lot = null; this.dc = null; this.supplier = null; this.reelId = null;
  }
  Draft.prototype.set = function (k, v) { this.fields[String(k).toUpperCase()] = v; };
  Draft.prototype.get = function (k) { var v = this.fields[String(k).toUpperCase()]; return v === undefined ? null : v; };
  Draft.prototype.has = function (k) { return this.fields.hasOwnProperty(String(k).toUpperCase()); };
  Draft.prototype.setLot = function (lot) { this.lot = isBlank(lot) ? null : String(lot).trim(); };
  Draft.prototype.setDateCode = function (raw) {
    if (raw == null) return;
    var s = String(raw).trim();
    if (!s.length) return;
    var u = s.toUpperCase();
    this.dc = (u === 'N/T' || u === 'NT' || u === 'N/A') ? null : s;   // "not traceable" = no date
  };
  Draft.prototype.setQty = function (raw) {
    if (isLongNumber(raw)) return;   // 7+ digits: internal number (Digi-Key oLPN / pick ID), never a reel quantity
    var q = tryParseQty(raw);
    if (q !== null) this.qty = q;
  };

  // ------------------------------------------------------------------ LabelNormalizer
  var RX_AIM = /^\]([A-Za-z])(\d)/;
  var RX_GS = /␝|\{GS\}|<GS>|\\x1D|\\u001D|\^\]/gi;
  var RX_RS = /␞|\{RS\}|<RS>|\\x1E|\\u001E|\^\^/gi;
  var RX_EOT = /␄|\{EOT\}|<EOT>|\\x04|\\u0004|\^D/gi;

  /** Canonical control characters (RS/GS/EOT as real bytes), AIM prefix split off. Returns {text, aim}. */
  function normalize(raw) {
    if (raw == null) return { text: '', aim: null };
    var s = trimChars(String(raw), '\r\n\t ﻿'), aim = null;
    var m = RX_AIM.exec(s);
    if (m) { aim = s.substring(0, 3); s = s.substring(3); }
    if (/[{<\\^␝␞␄]/.test(s)) s = s.replace(RX_RS, RS).replace(RX_GS, GS).replace(RX_EOT, EOT);
    var out = '';
    for (var i = 0; i < s.length; i++) {
      var c = s[i];
      // keep RS/GS/EOT, printable text and line breaks; other C0 controls are scanner framing
      if (c === GS || c === RS || c === EOT || c >= ' ' || c === '\n' || c === '\r' || c === '\t') out += c;
    }
    return { text: trimChars(out, '\r\n\t '), aim: aim };
  }

  function symbologyFromAim(aim) {
    if (!aim || aim.length < 2) return null;
    switch (aim[1].toUpperCase()) {
      case 'D': return 'DATA_MATRIX';
      case 'Q': return 'QR_CODE';
      case 'C': return 'CODE_128';
      case 'A': return 'CODE_39';
      case 'E': return 'EAN_13';
      case 'L': return 'PDF_417';
      default: return null;
    }
  }

  /** Digits only ('1.000' / '1,000' groups and a leading '+' tolerated); null when not a quantity. */
  function tryParseQty(raw) {
    if (isBlank(raw)) return null;
    var s = String(raw).trim();
    if (s[0] === '+') s = s.substring(1);
    if (/^\d{1,3}(?:[.,]\d{3})+$/.test(s)) s = s.replace(/[.,]/g, '');
    if (!/^\d{1,12}$/.test(s)) return null;
    return parseInt(s, 10);
  }

  function isLongNumber(s) {
    if (isBlank(s)) return false;
    s = String(s).trim();
    return /^\d+$/.test(s) && s.length > 6;
  }

  // ------------------------------------------------------------------ MpnHints (recognition + manufacturer only)
  var MPN_RULES = [
    [/^(RC|AC|RT|RL|RE|AF|RN|AA)(0075|0100|0201|0402|0603|0805|1206|1210|1218|2010|2512)([BCDFGJKM])([RKPS])([A-Z]?)-?(\d{2}|7W)([0-9RKM]+?)L?$/, 'Yageo'],
    [/^CRCW(0201|0402|0603|0805|1206|1210|1218|2010|2512)([0-9RKM]{4})([FJDBG])([A-Z0-9]*)$/, 'Vishay'],
    [/^RK73([BHGZ])(1F|1E|1J|2A|2B|2E|2H|3A)[A-Z]{1,3}(\d{3,4}|\d{1,3}R\d{0,3}|R\d{1,3})([FJDBG])?$/, 'KOA'],
    [/^ERJ-?(12|14|1G|1T|2|3|6|8|P06|P08|P14|PA3|PB3|PB6|U03|U06|U08)([A-Z]{2,3})([FJDBG])(\d{3,4}|\d{1,3}R\d{0,3}|R\d{1,3})([A-Z0-9]*)$/, 'Panasonic'],
    [/^WR(02|04|06|08|10|12|20|25)([A-Z])(\d{3,4}|\d{1,3}R\d{0,3}|R\d{1,3})([FJDBG])([A-Z0-9]*)$/, 'Walsin'],
    [/^(0201|0402|0603|0805|1206|1210|2010|2512)[WS]([A-Z0-9])([FJDBG])(\d{3,4}|\d{1,3}R\d{0,3}|R\d{1,3})(T[A-Z0-9]*)?$/, 'Uniroyal'],
    [/^CL(03|05|10|21|31|32|43|55)([ABCFLXY])(\d{2}[0-9R]|R\d{2})([BCDFGJKMZ])([RQPOABCDEFGH])/, 'Samsung'],
    [/^G(RM|CM|RT|RJ|CJ|RF)(03|15|18|21|31|32|43|55)([A-Z0-9])(5C|7U|R6|R7|C8|C7|E7|F5)([0-9][A-Z])(\d{2}[0-9R]|R\d{2})([BCDFGJKMZ])/, 'Murata'],
    [/^C(0402|0603|1005|1608|2012|3216|3225|4532|5750)(C0G|NP0|CH|JB|X7R|X5R|X7S|X6S|X7T|X8R|Y5V)([0-9][A-Z])(\d{2}[0-9R]|R\d{2})([BCDFGJKMZ])/, 'TDK'],
    [/^CGA(2B|3E|4J|5L|6M|8L|9N)[0-9]?(C0G|NP0|X7R|X5R|X7S|X6S|X7T|X8R)([0-9][A-Z])(\d{2}[0-9R]|R\d{2})([BCDFGJKMZ])/, 'TDK'],
    [/^C(0201|0402|0603|0805|1206|1210|1808|1812|1825|2220|2225)([CKSRVWXY])(\d{2}[0-9R]|R\d{2})([BCDFGJKMZ])([0-9A])([GRPZNUH])/, 'KEMET'],
    [/^CC(0201|0402|0603|0805|1206|1210|1812)([BCDFGJKMZ])([RKPS])(X7R|X5R|NPO|NP0|C0G|Y5V|X6S|X7S)([0-9A-Z])(BB|BN|BA|BC|BK)?(\d{2}[0-9R]|R\d{2})/, 'Yageo']
  ];

  /** Manufacturer when the MPN grammar is a known passive family (Yageo, Murata …), else null. */
  function mpnFamily(mpn) {
    if (isBlank(mpn)) return null;
    var m = String(mpn).trim().toUpperCase().replace(/ /g, '');
    for (var i = 0; i < MPN_RULES.length; i++) if (MPN_RULES[i][0].test(m)) return MPN_RULES[i][1];
    return null;
  }
  function recognizes(mpn) { return mpnFamily(mpn) !== null; }

  // ------------------------------------------------------------------ Iso15434 (ISO/IEC 15434 format 06 + MH10.8.2 DIs)
  var RX_ISO_HEADER = /^(>?)\[\)>\x1E?06\x1D/;
  var RX_ISO_HEADER_LENIENT = /^(>?)[^\x1D\x1E]{0,3}06\x1D/;
  var RX_ISO_DI = /^(\d{0,2}[A-Z])([\s\S]*)$/;

  function endsNd(s) { return s != null && /-ND$/i.test(String(s).trim()); }

  function parseIso(s, d) {
    if (!s) return false;
    var h = RX_ISO_HEADER.exec(s) || RX_ISO_HEADER_LENIENT.exec(s);
    if (!h) return false;
    var mouserHeader = h[1].length > 0;
    s = s.substring(h[0].length);
    var rs = s.indexOf(RS);
    if (rs >= 0) s = s.substring(0, rs);
    s = trimEndChars(s, EOT + GS + '\r\n ');

    var fields = [], parts = s.split(GS);
    for (var i = 0; i < parts.length; i++) {
      if (!parts[i].length) continue;
      var m = RX_ISO_DI.exec(parts[i]);
      if (m) fields.push([m[1], trimEndChars(m[2], '\r\n')]);
    }
    if (!fields.length) return false;
    var count = {};
    fields.forEach(function (kv) {
      var n = (count[kv[0]] || 0) + 1;
      count[kv[0]] = n;
      d.set(n === 1 ? kv[0] : kv[0] + '#' + n, kv[1]);
    });

    var p1 = d.get('1P'), p30 = d.get('30P'), p = d.get('P'), p3 = d.get('3P');
    var v1 = d.get('1V'), v = d.get('V'), sSer = d.get('S');
    var dkMarks = d.has('11Z') || d.has('12Z') || d.has('13Z') || d.has('20Z') || endsNd(p30) || endsNd(p);
    var profile;
    if (v != null && v.trim() === '0033317') profile = 'TI';
    else if (dkMarks) profile = 'Digi-Key';
    else if (p30 != null && p1 != null && /^\d{12}$/.test(p1.trim())) profile = 'Nexperia';
    else if (mouserHeader || d.has('14K')) profile = 'Mouser';
    else if (sSer != null && p30 == null && v1 == null) profile = 'onsemi';
    else profile = 'ECIA';

    var lot = d.get('1T'), date = null, reel = null;
    var dateKeys = ['9D', '10D', '11D', 'D', '6D', '14D', '12D', '16D'];
    for (i = 0; i < dateKeys.length; i++) if (d.has(dateKeys[i])) { date = d.get(dateKeys[i]); break; }

    switch (profile) {
      case 'TI':
        d.supplier = 'TI'; d.mfr = 'Texas Instruments'; d.mpn = p1;
        lot = coalesce(d.get('31T'), lot);
        reel = d.get('1T');                       // ship track code, not the lot
        if (d.has('D')) date = d.get('D');
        break;
      case 'Nexperia':
        d.supplier = 'Nexperia'; d.mfr = 'Nexperia';
        d.mpn = p30;                              // type ("BC847B"); 1P is the 12NC order code
        d.supplierPn = p1;
        break;
      case 'Digi-Key':
        d.supplier = 'Digi-Key'; d.mpn = p1;
        d.supplierPn = p30 != null ? p30 : (endsNd(p) ? p : null);
        reel = coalesce(d.get('12Z'), d.get('1K'));
        if (v1 != null) d.mfr = v1;
        break;
      case 'Mouser':
        d.supplier = 'Mouser'; d.mpn = p1; d.supplierPn = p30;
        if (v1 != null) d.mfr = v1;
        break;
      case 'onsemi':
        d.supplier = 'onsemi'; d.mfr = 'onsemi'; d.mpn = p1; reel = sSer;
        break;
      default:
        d.mpn = p1;
        d.supplierPn = p30 != null ? p30 : p3;
        if (v1 != null) d.mfr = v1;
        break;
    }
    if (d.mpn != null) d.mpn = d.mpn.trim();
    if (d.supplierPn != null) d.supplierPn = d.supplierPn.trim();
    if (d.mfr != null) d.mfr = d.mfr.trim();
    d.setQty(d.get('Q'));
    if (lot != null && /^(MIXED|MULTI|MIX|MULTIPLE)$/.test(lot.trim().toUpperCase())) lot = null;   // mixed-lot reel
    d.setLot(lot);
    d.setDateCode(date);
    if (!isBlank(reel)) d.reelId = reel.trim();
    return true;
  }

  // ------------------------------------------------------------------ DiSingle (one identifier-prefixed 1D field)
  var RX_DS_DI = /^(1P|30P|P|Q|9D|10D|11D|6D|1T|31T|4L|1V|K|4K|11K|S|3S)([\s\S]+)$/;
  var RX_DS_SEP = /[\x1D\x1E\x04\s{}:=|;,()]/;

  function is1D(sym) {
    if (isBlank(sym)) return true;   // wedge / unknown: shape is all we have
    var s = String(sym).toUpperCase();
    return s.indexOf('CODE_') === 0 || s === 'ITF' || s === 'CODABAR' || s.indexOf('EAN') === 0 || s.indexOf('UPC') === 0;
  }

  function parseDiSingle(s, sym, d) {
    if (!s || s.length < 2 || s.length > 48 || RX_DS_SEP.test(s)) return false;
    var m = RX_DS_DI.exec(s);
    if (!m) return false;
    var di = m[1], val = m[2].trim();
    if (!val.length) return false;
    var oneD = is1D(sym);
    if (recognizes(s)) return false;   // a catalogue part number that merely starts like a DI
    switch (di) {
      case 'Q':
        if (!/^\d{1,9}$/.test(val)) return false;
        d.setQty(val);
        if (d.qty === null) return false;
        break;
      case '9D': case '10D': case '11D': case '6D':
        if (!/^\d{4}M?$|^\d{6}$/.test(val)) return false;
        d.setDateCode(val);
        break;
      case '1T': case '31T':
        if (!/^[A-Za-z0-9][A-Za-z0-9\-./]{2,24}$/.test(val) || !oneD) return false;
        d.setLot(val);
        break;
      case '4L':
        if (!/^[A-Z]{2}$/.test(val)) return false;
        break;
      case '1V':
        if (!/^[A-Za-z][A-Za-z .&-]{1,30}$/.test(val) || !oneD) return false;
        d.mfr = val;
        break;
      case 'K': case '4K': case '11K': case '3S':
        if (!oneD || !/^[A-Za-z0-9][A-Za-z0-9\-./]{2,24}$/.test(val) || (di.length === 1 && !/\d/.test(val))) return false;
        break;
      case '1P': case '30P': case 'P': case 'S': {
        if (!/^[A-Za-z0-9][A-Za-z0-9\-./+#]{3,39}$/.test(val)) return false;
        var rec = recognizes(val);
        if ((di === 'P' || di === 'S') && !rec) return false;
        if (!rec && !(oneD && val.length >= 6)) return false;
        if (di === '30P') d.supplierPn = val;
        else if (di === 'S') d.reelId = val;
        else d.mpn = val;
        break;
      }
      default:
        return false;
    }
    d.set(di, val);
    return true;
  }

  // ------------------------------------------------------------------ LCSC / JLCPCB QR
  var RX_LCSC = /^\{((?:[^:,{}]+:[^:,{}]*,)*(?:[^:,{}]+:[^:,{}]*))\}$/;

  function parseLcsc(s, d) {
    var m = RX_LCSC.exec(s.trim());
    if (!m) return false;
    var pc = null;
    m[1].split(',').forEach(function (pair) {
      var i = pair.indexOf(':');
      if (i <= 0) return;
      var k = pair.substring(0, i).trim(), v = pair.substring(i + 1).trim();
      if (!k.length) return;
      d.set(k, v);
      if (k.toLowerCase() === 'pc' && v.length) pc = v;
    });
    if (pc == null) return false;
    d.supplier = 'LCSC';
    d.supplierPn = pc;
    var pm = d.get('pm');
    if (!isBlank(pm)) d.mpn = pm.trim();
    d.setQty(d.get('qty'));
    var pdi = d.get('pdi');
    if (!isBlank(pdi)) d.reelId = pdi.trim();
    return true;
  }

  // ------------------------------------------------------------------ TME (QR "QTY:.. PN:.. MPN:.." and space separated DataMatrix)
  var RX_TME_KEY = /^(QTY|PN|PO|CPO|MFR|MPN|LOT|DC)[:：]([\s\S]*)$/i;

  function tmeFill(d, pn, mpn, qty, mfr, lot, dc) {
    d.supplier = 'TME';
    d.supplierPn = pn.trim();
    d.mpn = mpn.trim();
    if (!isBlank(mfr)) d.mfr = mfr.trim();
    d.setQty(qty);
    if (lot != null) d.setLot(lot);
    if (dc != null) d.setDateCode(dc);
  }

  function parseTme(s, d) {
    if (!s || s.length > 400 || s.indexOf('[)>') >= 0 || s[0] === '{') return false;
    if (/^QTY:\s*\d/i.test(s) && /PN:/i.test(s) && /MPN:/i.test(s)) {
      var cur = null, buf = [], vals = {};
      var flush = function () { if (cur != null) vals[cur] = buf.join(' '); };
      s.split(/[ \t\r\n]+/).forEach(function (tok) {
        if (!tok) return;
        var m = RX_TME_KEY.exec(tok);
        if (m) { flush(); cur = m[1].toUpperCase(); buf = []; if (m[2].length) buf.push(m[2]); }
        else if (/^(https?:\/\/|www\.)/i.test(tok) || /^RoHS$/i.test(tok)) { /* url / RoHS flag */ }
        else if (cur != null) buf.push(tok);
      });
      flush();
      if (vals.PN == null || vals.MPN == null) return false;
      tmeFill(d, vals.PN, vals.MPN, vals.QTY, vals.MFR, vals.LOT != null ? vals.LOT : null, vals.DC != null ? vals.DC : null);
      return true;
    }
    var toks = s.split(' ').filter(function (t) { return t.length > 0; });
    if (toks.length < 3 || toks.length > 12) return false;
    var p = null, p1 = null, q = null;
    for (var i = 0; i < toks.length; i++) {
      var m2 = /^(\d{0,2}[A-Z])(\S+)$/.exec(toks[i]);
      if (!m2) return false;
      if (m2[1] === 'P') p = m2[2];
      else if (m2[1] === '1P') p1 = m2[2];
      else if (m2[1] === 'Q') q = m2[2];
    }
    if (p == null || q == null || p1 == null) return false;
    if (!/^\d{1,9}$/.test(q) || !/[\d\-]/.test(p)) return false;
    var mpn = /^MPN/i.test(p1) && p1.length > 3 ? p1.substring(3) : p1;
    tmeFill(d, p, mpn, q, null, null, null);
    return true;
  }

  // ------------------------------------------------------------------ GS1 element strings
  // [code, fixed length (0 = variable), max length]
  var GS1_TABLE = [
    ['00', 18, 18], ['01', 14, 14], ['02', 14, 14], ['10', 0, 20], ['11', 6, 6], ['12', 6, 6], ['13', 6, 6], ['15', 6, 6], ['16', 6, 6], ['17', 6, 6],
    ['20', 2, 2], ['21', 0, 20], ['22', 0, 20], ['30', 0, 8], ['37', 0, 8],
    ['240', 0, 30], ['241', 0, 30], ['242', 0, 6], ['250', 0, 30], ['251', 0, 30], ['253', 0, 30], ['254', 0, 20],
    ['400', 0, 30], ['401', 0, 30], ['402', 17, 17], ['403', 0, 30], ['422', 3, 3], ['424', 3, 3], ['426', 3, 3],
    ['7001', 13, 13], ['7003', 10, 10], ['7007', 0, 12], ['8001', 14, 14], ['8005', 6, 6],
    ['410', 13, 13], ['411', 13, 13], ['412', 13, 13], ['413', 13, 13], ['414', 13, 13], ['415', 13, 13], ['416', 13, 13], ['417', 13, 13],
    ['90', 0, 30], ['91', 0, 90], ['92', 0, 90], ['93', 0, 90], ['94', 0, 90], ['95', 0, 90], ['96', 0, 90], ['97', 0, 90], ['98', 0, 90], ['99', 0, 90]
  ];
  var GS1_MAP = {};
  GS1_TABLE.forEach(function (t) { GS1_MAP[t[0]] = t; });
  var GS1_NUMERIC = ['00', '01', '02', '11', '12', '13', '15', '16', '17', '20', '30', '37', '242', '402', '410', '411', '412', '413', '414', '415',
    '416', '417', '422', '424', '426', '7001', '7003', '8001', '8005'];

  function gs1Lookup(s, pos) {
    for (var len = 4; len >= 2; len--) {
      if (pos + len > s.length) continue;
      var code = s.substr(pos, len);
      if (/^\d+$/.test(code) && GS1_MAP[code]) return GS1_MAP[code];
    }
    return null;
  }

  /** AI code length from its first two digits (GS1 General Specifications §3.2); 0 when it cannot be an AI. */
  function gs1AiLength(s, pos) {
    if (pos + 2 > s.length) return 0;
    var p2 = s.substr(pos, 2);
    if (!/^\d\d$/.test(p2)) return 0;
    var n = parseInt(p2, 10), len;
    if (n <= 22 || n === 30 || n === 37 || n >= 90) len = 2;
    else if (n === 24 || n === 25 || n === 40 || n === 41 || n === 42 || n === 71) len = 3;
    else if ((n >= 31 && n <= 36) || n === 39 || n === 70 || n === 72 || n === 73 || (n >= 80 && n <= 82)) len = 4;
    else len = 2;
    if (pos + len > s.length || !/^\d+$/.test(s.substr(pos, len))) return 0;
    return len;
  }

  function parseGs1(normalized, d) {
    var gs1Aim = d.aim === ']C1' || d.aim === ']d2' || d.aim === ']Q3' || d.aim === ']e0';
    var s = normalized.replace(/^\x1D+/, '');   // leading FNC1
    var leadingFnc1 = s.length !== normalized.length;
    if (s.length < 4 || s.indexOf('[)>') >= 0 || s[0] === '{') return false;

    var parsed = [], paren = false;
    if (/^(?:\(\d{2,4}\)[^()]*)+$/.test(s)) {
      // human-readable "(01)0805S8J0470T5E(30)5000(7007)MSL 1": values run to the next "("
      var re = /\((\d{2,4})\)([^()]*)/g, m, n = 0;
      while ((m = re.exec(s)) !== null) {
        n++;
        var val0 = m[2].trim();
        if (val0.length) parsed.push([m[1], val0]);
      }
      if (n >= 2) {
        if (parsed.length < 2) return false;
        paren = true;
      } else parsed = [];
    }
    var pos = paren ? s.length : 0;
    while (pos < s.length) {
      if (s[pos] === GS) { pos++; continue; }
      var ai = gs1Lookup(s, pos);
      if (!ai) {
        var len = gs1AiLength(s, pos);   // unknown AI: GS1 prefix rule for its length, read as a variable field
        if (len === 0) return false;
        ai = [s.substr(pos, len), 0, 90];
      }
      pos += ai[0].length;
      var val;
      if (ai[1] > 0) {
        if (pos + ai[1] > s.length) return false;
        val = s.substr(pos, ai[1]);
        pos += ai[1];
      } else {
        var end = s.indexOf(GS, pos);
        if (end < 0) end = s.length;
        val = s.substring(pos, end);
        if (!val.length || val.length > ai[2]) return false;
        pos = end;
      }
      if (GS1_NUMERIC.indexOf(ai[0]) >= 0 && !/^\d+$/.test(val)) return false;
      parsed.push([ai[0], val]);
    }
    if (!parsed.length) return false;
    // guard against random digit strings "parsing" as one variable-length AI
    var strongStart = parsed[0][0] === '00' || parsed[0][0] === '01' || parsed[0][0] === '02';
    if (!gs1Aim && !leadingFnc1 && parsed.length < 2 && !strongStart) return false;
    if (!gs1Aim && !leadingFnc1 && parsed.length < 2 && strongStart && !/^\d+$/.test(s)) return false;

    var count = {};
    parsed.forEach(function (kv) {
      var c = (count[kv[0]] || 0) + 1;
      count[kv[0]] = c;
      d.set(c === 1 ? kv[0] : kv[0] + '#' + c, kv[1]);
    });
    var mpn = coalesce(d.get('240'), d.get('241'));
    if (!isBlank(mpn)) d.mpn = mpn.trim();
    else {
      var ai01 = d.get('01');   // AI 01 holding something that is not a 14-digit GTIN = the part number (Far-East reel QR)
      if (!isBlank(ai01) && !/^\d{14}$/.test(ai01)) d.mpn = ai01.trim();
    }
    d.setQty(coalesce(d.get('30'), d.get('37')));
    d.setLot(d.get('10'));
    d.setDateCode(coalesce(d.get('11'), d.get('13'), d.get('15')));
    var ser = d.get('21');
    if (!isBlank(ser)) d.reelId = ser.trim();
    return true;
  }

  // ------------------------------------------------------------------ number heuristics (LabelHeuristics.ResolveNumbers)
  function looksLikeDateCode(s) {
    s = String(s).trim();
    if (!/^\d{4}$|^\d{6}$/.test(s)) return false;
    var yy, ww;
    if (s.length === 4) {
      yy = +s.substr(0, 2); ww = +s.substr(2, 2);
      return yy >= 10 && yy <= 35 && ww >= 1 && ww <= 53;
    }
    var a = +s.substr(0, 4), b = +s.substr(4, 2);
    if (a >= 2010 && a <= 2035 && b >= 1 && b <= 53) return true;
    yy = +s.substr(0, 2);
    var mm = +s.substr(2, 2), dd = +s.substr(4, 2);
    return yy >= 10 && yy <= 35 && mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31;
  }

  /** Assigns bare numbers to qty / date / lot (fills only null slots of r). */
  function resolveNumbers(numbers, r) {
    var shortNums = [];
    numbers.forEach(function (n) {
      var t = n.trim();
      if (isLongNumber(t)) { if (r.lot == null) r.lot = t; } else shortNums.push(t);
    });
    if (!shortNums.length) return;
    if (shortNums.length === 1) {
      if (r.qty == null) r.qty = shortNums[0];   // a single bare number is the quantity (cut tape 1250)
      else if (r.date == null && looksLikeDateCode(shortNums[0])) r.date = shortNums[0];
      return;
    }
    var dateLike = shortNums.filter(looksLikeDateCode);
    var plain = shortNums.filter(function (t) { return !looksLikeDateCode(t); });
    if (r.date == null && dateLike.length === 1 && plain.length >= 1) {
      r.date = dateLike[0];
      if (r.qty == null) r.qty = plain[0];
      return;
    }
    if (r.date == null && dateLike.length >= 2 && plain.length === 0) {
      var q = null, i;
      for (i = 0; i < dateLike.length; i++) if (/(000|500)$/.test(dateLike[i])) { q = dateLike[i]; break; }   // reel qty ends in 000/500
      if (q == null) q = dateLike[0];
      if (r.qty == null) r.qty = q;
      for (i = 0; i < dateLike.length; i++) if (dateLike[i] !== q) { r.date = dateLike[i]; break; }
      return;
    }
    if (r.qty == null && plain.length) r.qty = plain[0];
    if (r.date == null && dateLike.length) r.date = dateLike[0];
    if (r.qty == null && dateLike.length) r.qty = dateLike[0];
  }

  // ------------------------------------------------------------------ BraceKv ("LT:18622497}D/C:2619}0805S8J0470T5E}5000")
  var RX_BRACE_PAIR = /^\s*([A-Za-z][A-Za-z0-9_./\- ]{0,11}?)\s*[:=]\s*(.+?)\s*$/;

  function parseBraceKv(s, d) {
    if (!s || s.length > 400 || s[0] === '{' || s.indexOf('[)>') >= 0 || s.indexOf('}') < 0) return false;
    var segs = s.split('}').filter(function (x) { return x.length > 0; });
    if (segs.length < 2) return false;
    var r = { mpn: null, qty: null, lot: null, date: null, mfr: null, sku: null }, known = 0, bare = [];
    segs.forEach(function (raw) {
      var seg = raw.trim();
      if (!seg.length) return;
      var m = RX_BRACE_PAIR.exec(seg);
      if (!m) { bare.push(seg); return; }
      var key = m[1].toUpperCase().replace(/[ _\-.]/g, ''), val = m[2].trim();
      switch (key) {
        case 'LT': case 'LOT': case 'LOTNO': case 'BATCH': case 'L/N': case 'LN': if (r.lot == null) r.lot = val; known++; break;
        case 'D/C': case 'DC': case 'DATE': case 'DATECODE': case 'D': if (r.date == null) r.date = val; known++; break;
        case 'P/N': case 'PN': case 'MPN': case 'PART': case 'PARTNO': case 'MODEL': case 'TYPE': if (r.mpn == null) r.mpn = val; known++; break;
        case 'QTY': case 'Q': case 'QUANTITY': case 'PCS': case 'ADET': if (r.qty == null) r.qty = val; known++; break;
        case 'MFR': case 'MFG': case 'BRAND': case 'MAKER': case 'VENDOR': if (r.mfr == null) r.mfr = val; known++; break;
        case 'SKU': case 'CODE': case 'ORDERCODE': case 'CPN': if (r.sku == null) r.sku = val; known++; break;
        default: bare.push(val); break;
      }
    });
    var nums = [];
    bare.forEach(function (b) {
      if (/^\d{1,12}$/.test(b)) nums.push(b);
      else if (r.mpn == null && /^(?=.*[A-Za-z])(?=.*\d)[A-Za-z0-9][A-Za-z0-9\-./+]{4,39}$/.test(b)) r.mpn = b;
    });
    resolveNumbers(nums, r);
    if (r.mpn == null && r.qty == null && r.lot == null) return false;
    if (known === 0 && bare.length < 2) return false;
    if (r.mpn != null) d.mpn = r.mpn;
    if (r.mfr != null) d.mfr = r.mfr;
    if (r.sku != null) d.supplierPn = r.sku;
    d.setQty(r.qty);
    if (r.lot != null) d.setLot(r.lot);
    if (r.date != null) d.setDateCode(r.date);
    return true;
  }

  // ------------------------------------------------------------------ GenericKv ("MPN=BC547B; QTY=100; …")
  var RX_KV_PAIR = /^\s*([A-Za-z][A-Za-z0-9_./\- ]{0,19}?)\s*[=:]\s*(.*?)\s*$/;

  function parseGenericKv(s, d) {
    if (!s || s[0] === '{' || s.indexOf('[)>') >= 0) return false;
    var pairs = [];
    splitNonEmpty(s, ',;|\n\r\t' + GS + RS).forEach(function (tok) {
      var m = RX_KV_PAIR.exec(tok);
      if (!m) return;
      var k = m[1].trim(), v = m[2].trim();
      if (!k.length || !v.length) return;
      if (k.toLowerCase().indexOf('http') === 0 && v.indexOf('//') === 0) return;
      pairs.push([k, v]);
    });
    if (pairs.length < 2) return false;
    var mpn = null, mfr = null, qty = null, lot = null, date = null, sku = null, supplier = null;
    pairs.forEach(function (kv) {
      switch (normKey(kv[0])) {
        case 'MPN': case 'PN': case 'P/N': case 'PART': case 'PARTNO': case 'PARTNUMBER': case 'MFRPN': case 'MFGPN': case 'MPARTNO': if (mpn == null) mpn = kv[1]; break;
        case 'MFR': case 'MFG': case 'MANUFACTURER': case 'BRAND': case 'MAKER': if (mfr == null) mfr = kv[1]; break;
        case 'QTY': case 'Q': case 'QUANTITY': case 'ADET': case 'CNT': case 'COUNT': if (qty == null) qty = kv[1]; break;
        case 'LOT': case 'LOTNO': case 'BATCH': case 'BATCHNO': case 'LOTNUMBER': if (lot == null) lot = kv[1]; break;
        case 'DC': case 'DATE': case 'DATECODE': case 'D/C': case 'DATECODE(YYWW)': if (date == null) date = kv[1]; break;
        case 'SKU': case 'SUPPLIERPN': case 'VENDORPN': case 'ORDERCODE': if (sku == null) sku = kv[1]; break;
        case 'SUPPLIER': case 'VENDOR': case 'DIST': case 'DISTRIBUTOR': if (supplier == null) supplier = kv[1]; break;
      }
    });
    if (mpn == null && qty == null && lot == null && sku == null) return false;
    if (mpn != null) d.mpn = mpn;
    if (mfr != null) d.mfr = mfr;
    if (sku != null) d.supplierPn = sku;
    if (supplier != null) d.supplier = supplier;
    d.setQty(qty);
    d.setLot(lot);
    d.setDateCode(date);
    return true;
  }

  // ------------------------------------------------------------------ Universal (shape based, any label text)
  var K_UNKNOWN = 0, K_MPN = 1, K_QTY = 2, K_LOT = 3, K_DATE = 4, K_MFR = 5, K_SKU = 6, K_URL = 7, K_NOISE = 8;
  var HARD_SEPS = GS + RS + EOT + '\n\r\t}{|;~¦│●';
  var RX_KEY_SPLIT = uni('\\s+(?=\\p{L}[\\p{L}/]{0,9}[:=\\uFF1D\\uFF1A])', '\\s+(?=[A-Za-z\\u00C0-\\u024F\\u4E00-\\u9FFF][A-Za-z\\u00C0-\\u024F\\u4E00-\\u9FFF/]{0,9}[:=\\uFF1D\\uFF1A])');
  var RX_KEY_COUNT = uni('(?:^|\\s)\\p{L}[\\p{L}/]{0,9}[:=\\uFF1D\\uFF1A]', '(?:^|\\s)[A-Za-z\\u00C0-\\u024F\\u4E00-\\u9FFF][A-Za-z\\u00C0-\\u024F\\u4E00-\\u9FFF/]{0,9}[:=\\uFF1D\\uFF1A]', 'g');
  var RX_INLINE = /^(QTY|Q|PCS|LOT|LT|L\/N|BATCH|D\/C|DC|DATE|DATECODE|P\/N|PN|D\/N|MPN|PART|MODEL|TYPE|MFR|MFG|BRAND|VENDOR|SKU|CPN|T\/C)\s+([\s\S]+)$/i;
  var RX_U_PAIR = /^\s*([A-Za-zÀ-ɏ一-鿿][A-Za-z0-9À-ɏ一-鿿_./\- ]{0,15}?)\s*[:=＝：]\s*(.+?)\s*$/;
  var RX_U_NUMBER = /^(\d{1,3}(?:[.,]\d{3})+|\d{1,9})\s*(pcs|pc|adet|ea|k)?$/i;
  var RX_U_ALNUM = /^[A-Za-z0-9][A-Za-z0-9\-./+#]{2,39}$/;
  var MFR_PREFIXES = ['RC', 'AC', 'RT', 'RL', 'CRCW', 'CL', 'GRM', 'GCM', 'GJM', 'ERJ', 'ERA', 'RK73', 'WR', 'CC', 'C0', 'C1', 'C2', 'C3', 'CGA', 'TL', 'LM',
    'LMV', 'NE', 'OP', 'AD', 'ADS', 'MCP', 'PIC', 'ATMEGA', 'ATTINY', 'STM32', 'STM8', 'ESP32', 'ESP8266', 'NRF', 'CY', 'MAX', 'LT', 'LTC', 'TPS', 'TLV',
    'IRF', 'IRL', 'AO', 'SI', 'FDN', 'BSS', 'BC', 'BCP', '2N', '1N', 'MMBT', 'PMBT', 'SS', 'SMAJ', 'SMBJ', 'SMF', 'US1', 'ES1', 'M7', 'FR', 'UF', 'B5',
    'B3', 'LL', 'MBR', 'SB', 'SK', 'TVS', '74HC', '74LS', '74AHC', 'CD4', 'SN74', 'MC', 'ULN', 'L7', 'LD', 'AMS', 'ME', 'XC', 'HT', 'CH', 'FT', 'CP21',
    'USB', 'W25', 'AT24', 'AT25', 'MX25', 'GD25', 'DS', 'PCF', 'PCA', 'TCA', 'SX', 'RFM', 'LAN', 'W5', 'ENC', 'KSZ', 'DP83', 'FS', 'FSMD', 'MF-', 'MLCC',
    'BLM', 'MMZ', 'NFM', 'SRN', 'SRR', 'CDRH', 'NR', 'VLS', 'MLZ', 'LQG', 'LQM', 'CBC', 'PE-', '744', '7447', 'WE-', 'ABM', 'ABL', 'ECS', 'FA-', 'CSTCE',
    'CSTNE', 'X322', 'X50', 'JLC'];

  function tokenize(text) {
    var list = [];
    if (isBlank(text)) return list;
    splitNonEmpty(text, HARD_SEPS).forEach(function (hard) {
      var piece = hard.trim();
      if (!piece.length) return;
      var parts;
      var colonCount = (piece.match(/[:=]/g) || []).length;
      RX_KEY_COUNT.lastIndex = 0;
      if (piece.indexOf(',') >= 0 && colonCount !== 1) parts = piece.split(',');
      else if ((piece.match(RX_KEY_COUNT) || []).length >= 2) parts = piece.split(RX_KEY_SPLIT);
      else parts = [piece];
      parts.forEach(function (pr) {
        var seg = pr.trim();
        if (!seg.length) return;
        var m = RX_U_PAIR.exec(seg);
        if (m && m[2].trim().length) { list.push({ key: m[1].trim(), text: m[2].trim() }); return; }
        seg.split(/\s{2,}/).forEach(function (col) {
          var c = col.trim();
          if (!c.length) return;
          var inl = RX_INLINE.exec(c);
          if (inl) { list.push({ key: inl[1], text: inl[2].trim() }); return; }
          var pcs = /^(\d{1,3}(?:[.,]\d{3})+|\d{1,9})\s*(pcs|pc|adet|ea)$/i.exec(c);   // "1200 pcs" -> qty
          if (pcs) list.push({ key: 'QTY', text: pcs[1] });
          else list.push({ key: null, text: c });
        });
      });
    });
    return list;
  }

  function kindOfKey(key) {
    switch (normKey(key)) {
      case 'MPN': case 'PN': case 'P/N': case 'PART': case 'PARTNO': case 'PARTNUMBER': case 'MFRPN': case 'MFGPN': case 'MPARTNO': case 'MODEL':
      case 'TYPE': case 'TYPENO': case 'SPEC': case 'ITEM': case 'ITEMNO': case 'PARÇA': case 'PARCA': case 'PARCANO': case '1P': case '30P':
      case 'PM': case 'D/N': case '型号': case '规格': case '料号': case '品名': case '零件号':
        return K_MPN;
      case 'QTY': case 'Q': case 'QUANTITY': case 'PCS': case 'ADET': case 'MIKTAR': case 'CNT': case 'COUNT': case 'QT':
      case '数量': case '数目':
        return K_QTY;
      case 'LOT': case 'LT': case 'LOTNO': case 'LOTNUMBER': case 'BATCH': case 'BATCHNO': case 'L/N': case 'LN': case 'PARTI': case '1T': case '31T':
      case 'PBN': case '批号': case '批次': case 'LOT号':
        return K_LOT;
      case 'DC': case 'D/C': case 'DATE': case 'DATECODE': case 'D': case '9D': case '10D': case 'TARIH': case 'TARİH': case 'MFGDATE':
      case 'PRODDATE': case '日期': case '生产日期': case '周期': case 'T/C': case 'DATECODE(YYWW)':
        return K_DATE;
      case 'MFR': case 'MFG': case 'MANUFACTURER': case 'BRAND': case 'MAKER': case 'VENDOR': case 'URETICI': case 'ÜRETİCİ': case '1V':
      case '厂商': case '厂家': case '品牌':
        return K_MFR;
      case 'SKU': case 'CPN': case 'CODE': case 'ORDERCODE': case 'SUPPLIERPN': case 'VENDORPN': case 'PC': case 'P': case 'DKPN':
        return K_SKU;
      default:
        return K_UNKNOWN;
    }
  }

  /** Shape of a bare token -> {kind, score} (score only meaningful for MPN candidates). */
  function classifyBare(t) {
    if (isBlank(t)) return { kind: K_NOISE, score: 0 };
    var s = String(t).trim();
    if (/^(https?:\/\/|www\.)/i.test(s)) return { kind: K_URL, score: 0 };
    if (RX_U_NUMBER.test(s)) {
      if (isLongNumber(s.replace(/[^\d]/g, ''))) return { kind: K_UNKNOWN, score: 0 };   // 7+ digits: lot / internal number
      return { kind: looksLikeDateCode(s) ? K_DATE : K_QTY, score: 0 };
    }
    if (/^(LOT|LT|BATCH|L\/N)[-_#:]?\s?[A-Za-z0-9\-]{3,}$/i.test(s)) return { kind: K_LOT, score: 0 };
    if (/^(\d{4}|\d{2})[-/ ]?W[-/ ]?\d{2}$/i.test(s)) return { kind: K_DATE, score: 0 };
    if (!RX_U_ALNUM.test(s)) return { kind: s.length <= 3 ? K_NOISE : K_UNKNOWN, score: 0 };
    var hasLetter = /[A-Za-z]/.test(s), hasDigit = /\d/.test(s);
    if (!hasDigit) return { kind: s.length >= 3 && s.length <= 24 ? K_MFR : K_UNKNOWN, score: 0 };
    if (!hasLetter) return { kind: K_QTY, score: 0 };
    var up = s.toUpperCase();
    if (recognizes(up)) return { kind: K_MPN, score: 10 };
    var score = 1;
    for (var i = 0; i < MFR_PREFIXES.length; i++) if (up.indexOf(MFR_PREFIXES[i]) === 0) { score += 2; break; }
    if (up.indexOf('-') >= 0) score += 0.5;
    if (s.length >= 7 && s.length <= 22) score += 0.5;
    var letters = (up.match(/[A-Z]/g) || []).length;
    if (letters <= 2 && s.length >= 8) score -= 1;   // "18622497A": lot-like
    if (score <= 0.5 && s.length >= 8) return { kind: K_LOT, score: score };
    return { kind: K_MPN, score: score };
  }

  function parseUniversal(s, d) {
    var tokens = tokenize(s);
    if (!tokens.length) return false;
    var mpn = null, qty = null, lot = null, date = null, mfr = null, sku = null, mpnScore = -1;
    tokens.forEach(function (t) {
      var c;
      if (t.key != null) {
        var k = kindOfKey(t.key);
        if (k === K_UNKNOWN) c = classifyBare(t.text);
        else c = { kind: k, score: k === K_MPN ? 5 : 0 };
      } else {
        c = classifyBare(t.text);
        if (c.kind === K_UNKNOWN && isLongNumber(t.text) && tokens.length > 1) c.kind = K_LOT;   // long bare number next to other fields
      }
      var text = t.text;
      if (c.kind === K_DATE && t.key != null && text.indexOf(' ') > 0) text = text.substring(0, text.indexOf(' '));   // "T/C 2335 3302RE00"
      switch (c.kind) {
        case K_MPN: if (c.score > mpnScore) { mpn = text; mpnScore = c.score; } break;
        case K_QTY: if (qty == null) qty = text; else if (looksLikeDateCode(qty) && date == null) { date = qty; qty = text; } break;
        case K_LOT: if (lot == null) lot = text; break;
        case K_DATE: if (date == null) date = text; else if (qty == null) qty = text; break;
        case K_MFR: if (mfr == null) mfr = text; break;
        case K_SKU: if (sku == null) sku = text; break;
      }
    });
    if (mpn == null && !(qty != null && (lot != null || date != null))) return false;
    if (mpn != null) d.mpn = mpn;
    if (mfr != null) d.mfr = mfr;
    if (sku != null) d.supplierPn = sku;
    d.setQty(qty);
    if (lot != null) d.setLot(lot);
    if (date != null) d.setDateCode(date);
    return true;
  }

  function parseRaw(s, d) {
    var c = classifyBare(s);
    if (c.kind === K_MPN && /^[A-Z0-9][A-Z0-9\-./,+#]{3,39}$/i.test(s) && s.indexOf(GS) < 0) d.mpn = s;
    else if (c.kind === K_QTY) d.setQty(s);
    else if (c.kind === K_DATE) d.setDateCode(s);
    else if (c.kind === K_LOT) d.setLot(s);
  }

  // ------------------------------------------------------------------ chain
  var CHAIN = [
    ['Iso15434', function (s, sym, d) { return parseIso(s, d); }],
    ['DiSingle', parseDiSingle],
    ['Lcsc', function (s, sym, d) { return parseLcsc(s, d); }],
    ['Tme', function (s, sym, d) { return parseTme(s, d); }],
    ['Gs1', function (s, sym, d) { return parseGs1(s, d); }],
    ['BraceKv', function (s, sym, d) { return parseBraceKv(s, d); }],
    ['GenericKv', function (s, sym, d) { return parseGenericKv(s, d); }],
    ['Universal', function (s, sym, d) { return parseUniversal(s, d); }]
  ];

  function emptyOrNull(v) { return v == null || (typeof v === 'string' && v.trim() === '') ? null : v; }

  /**
   * Preview of one scanned label. fmt = ZXing format name (optional; improves 1D/2D decisions).
   * Returns {mpn, qty, mfr, supplierPn, lot, dc, supplier, parser, reelId}; qty is a number, everything may be null.
   * Never throws.
   */
  function preview(raw, fmt) {
    var n = normalize(raw), sym = isBlank(fmt) ? symbologyFromAim(n.aim) : String(fmt).trim();
    var d = new Draft(n.aim), parser = 'Raw';
    if (n.text.length) {
      for (var i = 0; i < CHAIN.length; i++) {
        var attempt = new Draft(n.aim), ok = false;
        try { ok = CHAIN[i][1](n.text, sym, attempt); } catch (e) { ok = false; }
        if (ok) { d = attempt; parser = CHAIN[i][0]; break; }
      }
      if (parser === 'Raw') { try { parseRaw(n.text, d); } catch (e) { /* preview never throws */ } }
    }
    if (isBlank(d.mfr)) d.mfr = mpnFamily(d.mpn);
    return {
      mpn: emptyOrNull(d.mpn),
      qty: typeof d.qty === 'number' && isFinite(d.qty) ? d.qty : null,
      mfr: emptyOrNull(d.mfr),
      supplierPn: emptyOrNull(d.supplierPn),
      lot: emptyOrNull(d.lot),
      dc: emptyOrNull(d.dc),
      supplier: emptyOrNull(d.supplier),
      parser: parser,
      reelId: emptyOrNull(d.reelId)
    };
  }

  return {
    preview: preview,
    normalize: normalize,
    symbologyFromAim: symbologyFromAim,
    tryParseQty: tryParseQty,
    mpnFamily: mpnFamily
  };
});
