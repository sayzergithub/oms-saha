/*
 * OMS field mode - signed + encrypted basket envelope (spec V5 §1-§3).
 * Byte-for-byte twin of Remote/RemoteEnvelope.cs:
 *   encKey = HMAC-SHA256(sec, "oms-saha-enc"), macKey = HMAC-SHA256(sec, "oms-saha-mac")
 *   frame  = 0x02 | flags | nonce(16) | AES-256-CTR(encKey, nonce, 128-bit counter)(plain) | HMAC-SHA256(macKey, dev 0x00 frame[0..18+n))
 *   armor  = "OMSB1:" + dev + ":" + base64url(frame)          (flags bit0 = plain was deflate-raw compressed)
 * UMD, no dependencies. Browser: global OmsEnvelope. Node 18+: module.exports.
 */
(function (root, factory) {
  'use strict';
  if (typeof module === 'object' && module.exports) module.exports = factory(root);
  else root.OmsEnvelope = factory(root);
})(typeof globalThis !== 'undefined' ? globalThis : (typeof self !== 'undefined' ? self : this), function (root) {
  'use strict';

  // ------------------------------------------------------------------ constants (Remote/RemoteModels.cs RemoteConst)
  var SUBJECT = 'oms_inv_sepet';
  var PAIR_PREFIX = 'OMSP1:';
  var ARMOR_PREFIX = 'OMSB1:';
  var BEGIN_MARKER = '-----BEGIN OMS SEPET-----';
  var END_MARKER = '-----END OMS SEPET-----';
  var ATTACHMENT_EXT = '.omsb';
  var DEVICE_ALPHABET = 'abcdefghjkmnpqrstuvwxyz23456789';
  var VERSION = 2, FLAG_DEFLATE = 1, HEADER_LEN = 18, TAG_LEN = 32, NONCE_LEN = 16;
  var MAX_FRAME = 4 * 1024 * 1024;
  var MAX_PLAIN = 8 * 1024 * 1024;
  var MAIL_COLS = 76;

  // ------------------------------------------------------------------ crypto source
  function getCrypto() {
    var c = root && root.crypto;
    if (c && c.subtle && c.getRandomValues) return c;
    // Node 18 has webcrypto only behind require('crypto').webcrypto.
    if (typeof module === 'object' && module.exports && typeof require === 'function') {
      try { var w = require('crypto').webcrypto; if (w && w.subtle) return w; } catch (e) { /* not node */ }
    }
    return null;
  }

  function needCrypto() {
    var c = getCrypto();
    if (!c) throw envError('NoCrypto', 'Bu tarayıcıda şifreleme (WebCrypto) yok. Sayfayı HTTPS adresinden açın.');
    return c;
  }

  /** True when WebCrypto (crypto.subtle) is usable here (it is missing on plain-HTTP pages). */
  function supported() { return getCrypto() !== null; }

  function envError(code, message) {
    var e = new Error(message || code);
    e.code = code;   // same names as C# EnvelopeStatus: BadFormat, BadVersion, BadTag, TooLarge, BadPayload (+ NoCrypto, BadArgument)
    return e;
  }

  // ------------------------------------------------------------------ bytes / text
  var te = typeof TextEncoder !== 'undefined' ? new TextEncoder() : null;

  function utf8(s) {
    if (te) return te.encode(s);
    // Fallback (very old engines): encodeURIComponent round trip.
    var bin = unescape(encodeURIComponent(s)), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  /** Strict UTF-8 decode (invalid bytes throw, like C# new UTF8Encoding(false, true)). */
  function fromUtf8(b) {
    if (typeof TextDecoder !== 'undefined') return new TextDecoder('utf-8', { fatal: true }).decode(b);
    var bin = '';
    for (var i = 0; i < b.length; i++) bin += String.fromCharCode(b[i]);
    return decodeURIComponent(escape(bin));
  }

  function ascii(s) {
    var out = new Uint8Array(s.length);
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      if (c > 127) throw envError('BadArgument', 'ASCII olmayan karakter');
      out[i] = c;
    }
    return out;
  }

  function concat(parts) {
    var n = 0, i;
    for (i = 0; i < parts.length; i++) n += parts[i].length;
    var out = new Uint8Array(n), o = 0;
    for (i = 0; i < parts.length; i++) { out.set(parts[i], o); o += parts[i].length; }
    return out;
  }

  function toBytes(x, what) {
    if (x instanceof Uint8Array) return x;
    if (x instanceof ArrayBuffer) return new Uint8Array(x);
    if (x && ArrayBuffer.isView(x)) return new Uint8Array(x.buffer, x.byteOffset, x.byteLength);
    if (Array.isArray(x)) return new Uint8Array(x);
    throw envError('BadArgument', (what || 'veri') + ' bayt dizisi olmalı');
  }

  // ------------------------------------------------------------------ base64url (RFC 4648 §5, no padding)
  var B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';
  var B64_REV = (function () {
    var r = new Int16Array(128);
    for (var i = 0; i < 128; i++) r[i] = -1;
    for (i = 0; i < 64; i++) r[B64.charCodeAt(i)] = i;
    r[43] = 62; r[47] = 63;   // also accept standard '+' '/' (C# B64UrlDecode does)
    return r;
  })();

  function b64urlEncode(bytes) {
    var b = toBytes(bytes), out = [], i, n = b.length, v;
    for (i = 0; i + 2 < n; i += 3) {
      v = (b[i] << 16) | (b[i + 1] << 8) | b[i + 2];
      out.push(B64[v >> 18], B64[(v >> 12) & 63], B64[(v >> 6) & 63], B64[v & 63]);
    }
    if (n - i === 1) { v = b[i] << 16; out.push(B64[v >> 18], B64[(v >> 12) & 63]); }
    else if (n - i === 2) { v = (b[i] << 16) | (b[i + 1] << 8); out.push(B64[v >> 18], B64[(v >> 12) & 63], B64[(v >> 6) & 63]); }
    return out.join('');
  }

  /** Decodes base64url (or standard base64); whitespace and '=' are skipped. Returns null on any other character / bad length. */
  function b64urlDecode(s) {
    if (typeof s !== 'string') return null;
    var vals = new Uint8Array(s.length), m = 0, i, c;
    for (i = 0; i < s.length; i++) {
      c = s.charCodeAt(i);
      if (c === 61 || c === 32 || c === 9 || c === 10 || c === 13 || c === 12 || c === 11) continue;
      if (c >= 128 || B64_REV[c] < 0) return null;
      vals[m++] = B64_REV[c];
    }
    if (m % 4 === 1) return null;
    var out = new Uint8Array(Math.floor(m * 3 / 4)), o = 0;
    for (i = 0; i + 3 < m; i += 4) {
      var v = (vals[i] << 18) | (vals[i + 1] << 12) | (vals[i + 2] << 6) | vals[i + 3];
      out[o++] = v >> 16; out[o++] = (v >> 8) & 255; out[o++] = v & 255;
    }
    if (m - i === 2) out[o++] = ((vals[i] << 18) | (vals[i + 1] << 12)) >> 16;
    else if (m - i === 3) {
      var w = (vals[i] << 18) | (vals[i + 1] << 12) | (vals[i + 2] << 6);
      out[o++] = w >> 16; out[o++] = (w >> 8) & 255;
    }
    return out;
  }

  // ------------------------------------------------------------------ ids
  function randomBytes(n) {
    var b = new Uint8Array(n);
    needCrypto().getRandomValues(b);
    return b;
  }

  function isDeviceId(s) {
    if (typeof s !== 'string' || s.length !== 8) return false;
    for (var i = 0; i < 8; i++) if (DEVICE_ALPHABET.indexOf(s[i]) < 0) return false;
    return true;
  }

  /** RFC 4122 v4 UUID (crypto.randomUUID when present, else getRandomValues). */
  function newUuid() {
    var c = needCrypto();
    if (typeof c.randomUUID === 'function') { try { return c.randomUUID(); } catch (e) { /* insecure context */ } }
    var b = randomBytes(16), h = [];
    b[6] = (b[6] & 0x0f) | 0x40;
    b[8] = (b[8] & 0x3f) | 0x80;
    for (var i = 0; i < 16; i++) h.push((b[i] + 0x100).toString(16).slice(1));
    return h.slice(0, 4).join('') + '-' + h.slice(4, 6).join('') + '-' + h.slice(6, 8).join('') + '-' +
      h.slice(8, 10).join('') + '-' + h.slice(10, 16).join('');
  }

  // ------------------------------------------------------------------ primitives
  function hmac(keyBytes, dataBytes) {
    var s = needCrypto().subtle;
    return s.importKey('raw', keyBytes, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign'])
      .then(function (k) { return s.sign('HMAC', k, dataBytes); })
      .then(function (r) { return new Uint8Array(r); });
  }

  function deriveKeys(secret) {
    var sec = toBytes(secret, 'sec');
    if (sec.length === 0) return Promise.reject(envError('BadArgument', 'Cihaz sırrı boş'));
    return Promise.all([hmac(sec, ascii('oms-saha-enc')), hmac(sec, ascii('oms-saha-mac'))])
      .then(function (k) { return { enc: k[0], mac: k[1] }; });
  }

  function aesCtr(encKey, nonce, data) {
    var s = needCrypto().subtle;
    if (data.length === 0) return Promise.resolve(new Uint8Array(0));
    return s.importKey('raw', encKey, { name: 'AES-CTR' }, false, ['encrypt'])
      // length:128 = whole 16-byte block is a big-endian counter (same as the C# Ctr loop).
      .then(function (k) { return s.encrypt({ name: 'AES-CTR', counter: nonce, length: 128 }, k, data); })
      .then(function (r) { return new Uint8Array(r); });
  }

  function tagInput(dev, frame, count) {
    return concat([ascii(dev), new Uint8Array([0]), frame.subarray(0, count)]);
  }

  // ------------------------------------------------------------------ deflate-raw (CompressionStream)
  /** True when this browser can compress with CompressionStream('deflate-raw') (Chrome 103+, Safari 16.4+, Firefox 113+). */
  function supportsDeflate() {
    if (typeof root.CompressionStream !== 'function') return false;
    try { new root.CompressionStream('deflate-raw'); return true; } catch (e) { return false; }
  }

  function supportsInflate() {
    if (typeof root.DecompressionStream !== 'function') return false;
    try { new root.DecompressionStream('deflate-raw'); return true; } catch (e) { return false; }
  }

  /** Pipes bytes through a (De)CompressionStream; aborts with TooLarge once output exceeds max. */
  function pump(stream, data, max) {
    var writer = stream.writable.getWriter();
    var reader = stream.readable.getReader();
    // Write errors surface on the reader as well; swallow them here so they are not unhandled.
    writer.write(data).then(function () { return writer.close(); }).catch(function () { });
    var chunks = [], total = 0;
    function step() {
      return reader.read().then(function (r) {
        if (r.done) return concat(chunks);
        var v = toBytes(r.value);
        total += v.length;
        if (total > max) {
          try { reader.cancel(); } catch (e) { /* ignore */ }
          throw envError('TooLarge', 'Açılmış sepet çok büyük');
        }
        chunks.push(v);
        return step();
      });
    }
    return step();
  }

  function deflateRaw(data) {
    return pump(new root.CompressionStream('deflate-raw'), data, MAX_FRAME);
  }

  function inflateRaw(data) {
    if (!supportsInflate()) return Promise.reject(envError('BadPayload', 'Bu tarayıcı sıkıştırılmış sepeti açamıyor'));
    return pump(new root.DecompressionStream('deflate-raw'), data, MAX_PLAIN).catch(function (e) {
      if (e && e.code) throw e;
      throw envError('BadPayload', 'Sıkıştırılmış veri bozuk');
    });
  }

  // ------------------------------------------------------------------ seal / open
  /**
   * Seals basket JSON for device dev with its 32-byte secret.
   * opts.deflate: compress first (ignored when the browser cannot; check supportsDeflate()).
   * opts.nonce: 16 bytes, test vectors only.
   * Resolves with the armor string "OMSB1:dev:...".
   */
  function seal(json, dev, secret, opts) {
    opts = opts || {};
    return Promise.resolve().then(function () {
      if (typeof json !== 'string') throw envError('BadArgument', 'json metin olmalı');
      if (!isDeviceId(dev)) throw envError('BadArgument', 'Geçersiz cihaz kimliği');
      var nonce = opts.nonce != null ? toBytes(opts.nonce, 'nonce') : randomBytes(NONCE_LEN);
      if (nonce.length !== NONCE_LEN) throw envError('BadArgument', 'nonce 16 bayt olmalı');
      var deflate = !!opts.deflate && supportsDeflate();
      var plain = utf8(json);
      if (plain.length > MAX_PLAIN) throw envError('TooLarge', 'Sepet çok büyük');
      return Promise.all([deriveKeys(secret), deflate ? deflateRaw(plain) : plain]).then(function (r) {
        var keys = r[0], body = r[1];
        if (HEADER_LEN + body.length + TAG_LEN > MAX_FRAME) throw envError('TooLarge', 'Sepet çok büyük');
        return aesCtr(keys.enc, nonce, body).then(function (ct) {
          var n = ct.length, f = new Uint8Array(HEADER_LEN + n + TAG_LEN);
          f[0] = VERSION;
          f[1] = deflate ? FLAG_DEFLATE : 0;
          f.set(nonce, 2);
          f.set(ct, HEADER_LEN);
          return hmac(keys.mac, tagInput(dev, f, HEADER_LEN + n)).then(function (tag) {
            f.set(tag, HEADER_LEN + n);
            return ARMOR_PREFIX + dev + ':' + b64urlEncode(f);
          });
        });
      });
    });
  }

  /** Splits "OMSB1:dev:base64url" without a key. Returns {dev, frame} or null (same rules as C# TryParseArmor). */
  function parseArmor(armor) {
    if (typeof armor !== 'string') return null;
    armor = armor.trim();
    if (armor.indexOf(ARMOR_PREFIX) !== 0) return null;
    var c = armor.indexOf(':', ARMOR_PREFIX.length);
    if (c < 0) return null;
    var dev = armor.substring(ARMOR_PREFIX.length, c);
    if (!isDeviceId(dev)) return null;
    var body = armor.substring(c + 1);
    if (body.length > Math.floor(MAX_FRAME / 3) * 4 + 8) return null;
    var frame = b64urlDecode(body);
    if (!frame || frame.length < HEADER_LEN + TAG_LEN) return null;
    return { dev: dev, frame: frame };
  }

  /** Verifies and decrypts an armor with the device secret. Resolves with the JSON text; rejects with err.code like C# EnvelopeStatus. */
  function openFull(armor, secret) {
    return Promise.resolve().then(function () {
      var p = parseArmor(armor);
      if (!p) throw envError('BadFormat', 'Zırh biçimi geçersiz');
      var f = p.frame;
      if (f.length > MAX_FRAME) throw envError('TooLarge', 'Zarf çok büyük');
      if (f[0] !== VERSION) throw envError('BadVersion', 'Desteklenmeyen zarf sürümü');
      if ((f[1] & ~FLAG_DEFLATE) !== 0) throw envError('BadVersion', 'Bilinmeyen zarf bayrağı');
      var n = f.length - HEADER_LEN - TAG_LEN;
      return deriveKeys(secret).then(function (keys) {
        var s = needCrypto().subtle;
        // subtle.verify compares in constant time.
        return s.importKey('raw', keys.mac, { name: 'HMAC', hash: 'SHA-256' }, false, ['verify'])
          .then(function (k) { return s.verify('HMAC', k, f.subarray(HEADER_LEN + n), tagInput(p.dev, f, HEADER_LEN + n)); })
          .then(function (ok) {
            if (!ok) throw envError('BadTag', 'İmza doğrulanamadı (yanlış cihaz sırrı ya da bozuk veri)');
            return aesCtr(keys.enc, f.slice(2, 2 + NONCE_LEN), f.slice(HEADER_LEN, HEADER_LEN + n));
          })
          .then(function (plain) { return (f[1] & FLAG_DEFLATE) ? inflateRaw(plain) : plain; })
          .then(function (plain) {
            if (plain.length > MAX_PLAIN) throw envError('TooLarge', 'Açılmış sepet çok büyük');
            var json;
            try { json = fromUtf8(plain); } catch (e) { throw envError('BadPayload', 'Sepet metni UTF-8 değil'); }
            return { dev: p.dev, flags: f[1], json: json };
          });
      });
    });
  }

  function open(armor, secret) {
    return openFull(armor, secret).then(function (r) { return r.json; });
  }

  // ------------------------------------------------------------------ mail transport
  /** Armor wrapped for a mail body: BEGIN marker, armor in 76-column lines, END marker (CRLF, RFC 6068 mailto bodies). */
  function wrapForMail(armor, eol) {
    eol = eol || '\r\n';
    var a = String(armor).replace(/\s+/g, ''), lines = [BEGIN_MARKER];
    for (var i = 0; i < a.length; i += MAIL_COLS) lines.push(a.substring(i, i + MAIL_COLS));
    lines.push(END_MARKER);
    return lines.join(eol) + eol;
  }

  /** Every armor found in a text (BEGIN/END blocks; else the whole text if it is a bare armor). Same cleaning as C# ExtractFromText. */
  function extractArmors(text) {
    var list = [];
    if (typeof text !== 'string' || !text) return list;
    function clean(s) { return s.replace(/<[^>]*>/g, '').replace(/[\s>]+|=(?=\r?\n)|&nbsp;/g, ''); }
    var re = /-----BEGIN OMS SEPET-----([\s\S]*?)-----END OMS SEPET-----/g, m;
    while ((m = re.exec(text)) !== null) {
      var a = clean(m[1]);
      if (a.indexOf(ARMOR_PREFIX) === 0 && list.indexOf(a) < 0) list.push(a);
    }
    if (list.length === 0) {
      var w = clean(text);
      if (w.indexOf(ARMOR_PREFIX) === 0) list.push(w);
    }
    return list;
  }

  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  /** Attachment name oms_sepet_yyyyMMdd_HHmmss.omsb (local time). */
  function fileName(date) {
    var d = date || new Date();
    return 'oms_sepet_' + d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate()) + '_' +
      pad2(d.getHours()) + pad2(d.getMinutes()) + pad2(d.getSeconds()) + ATTACHMENT_EXT;
  }

  // ------------------------------------------------------------------ pairing
  /**
   * Parses a pairing code: bare "OMSP1:...", a URL with "#OMSP1:..." (also percent-encoded), or text with line breaks / spaces in it.
   * Returns {dev, sec:Uint8Array(32), name, to, pc, subj, lan:string[], cid} or null.
   */
  function parsePairing(textOrUrl) {
    if (typeof textOrUrl !== 'string') return null;
    var t = textOrUrl;
    if (t.indexOf(PAIR_PREFIX) < 0 && /%3A/i.test(t)) { try { t = decodeURIComponent(t); } catch (e) { return null; } }
    var i = t.indexOf(PAIR_PREFIX);
    if (i < 0) return null;
    // The code runs until the first character that is not base64(url)/whitespace (e.g. '&', '?', quote, end of URL).
    var rest = t.substring(i + PAIR_PREFIX.length), m = /^[A-Za-z0-9\-_+\/=\s]*/.exec(rest);
    var bytes = b64urlDecode(m ? m[0] : '');
    if (!bytes || bytes.length === 0) return null;
    var o;
    try { o = JSON.parse(fromUtf8(bytes)); } catch (e) { return null; }
    if (!o || typeof o !== 'object' || o.v !== 1 || !isDeviceId(o.dev)) return null;
    var sec = typeof o.sec === 'string' ? b64urlDecode(o.sec) : null;
    if (!sec || sec.length !== 32) return null;
    function str(x) { return typeof x === 'string' ? x.trim() : ''; }
    // V6 spec §10: lan = PC HTTPS roots (catalog), cid = Entra client id (direct send); absent in V5 codes
    var lan = Array.isArray(o.lan) ? o.lan.filter(function (u) { return typeof u === 'string' && /^https?:\/\//i.test(u); }) : [];
    return { dev: o.dev, sec: sec, name: str(o.name), to: str(o.to), pc: str(o.pc), subj: str(o.subj) || SUBJECT, lan: lan, cid: str(o.cid) };
  }

  /** Builds a pairing code (mirror of C# BuildPairingCode; used by tests and tools). */
  function buildPairing(dev, secret, name, to, pc) {
    var o = { v: 1, dev: dev, sec: b64urlEncode(secret), name: name || '', to: to || '', pc: pc || '', subj: SUBJECT };
    return PAIR_PREFIX + b64urlEncode(utf8(JSON.stringify(o)));
  }

  return {
    SUBJECT: SUBJECT,
    PAIR_PREFIX: PAIR_PREFIX,
    ARMOR_PREFIX: ARMOR_PREFIX,
    BEGIN_MARKER: BEGIN_MARKER,
    END_MARKER: END_MARKER,
    ATTACHMENT_EXT: ATTACHMENT_EXT,
    MAX_FRAME: MAX_FRAME,
    MAX_PLAIN: MAX_PLAIN,
    supported: supported,
    supportsDeflate: supportsDeflate,
    supportsInflate: supportsInflate,
    seal: seal,
    open: open,
    openFull: openFull,
    parseArmor: parseArmor,
    wrapForMail: wrapForMail,
    extractArmors: extractArmors,
    fileName: fileName,
    parsePairing: parsePairing,
    buildPairing: buildPairing,
    isDeviceId: isDeviceId,
    newUuid: newUuid,
    b64urlEncode: b64urlEncode,
    b64urlDecode: b64urlDecode,
    deriveKeys: deriveKeys,
    utf8: utf8
  };
});
