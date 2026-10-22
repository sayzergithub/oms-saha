/*
 * OMS field mode - the "Buluta gonder" button: hands the sealed basket file to the user's cloud (spec V7 §16).
 *   share    Web Share Level 2 with a file. The user picks the cloud app (or iOS "Dosyalar'a Kaydet") and the "OMS Saha" folder.
 *   folder   File System Access (desktop Chrome/Edge): the "OMS Saha" sync folder is picked once, its handle is kept in IndexedDB,
 *            later sends write the file straight into it (permission is re-confirmed with one tap per session when needed).
 *   download a[download] blob; the user moves it from Downloads into "OMS Saha".
 *
 * File type facts this module is built on (checked 2026-10 against the sources):
 *   - Chromium (Android, Windows, macOS, ChromeOS; also Samsung Internet / Edge) only shares files whose extension AND MIME type are
 *     both on its allow-list (chrome/browser/webshare/share_service_impl.cc, MDN navigator.share): ".txt" + "text/plain" passes,
 *     ".omsb" never does. That check runs inside share() (NotAllowedError); canShare() still says true for ".omsb" (seen in
 *     headless Chrome 2026-10). Hence every non-WebKit browser shares ".omsb.txt" / text/plain (the PC accepts it, spec §14).
 *   - WebKit (iOS/iPadOS Safari and every iOS browser) has no file-type check in Navigator::canShare/share (WebCore/page/Navigator.cpp),
 *     so iOS keeps the real ".omsb" name; "Dosyalar'a Kaydet" offers any file. ".omsb.txt" is the fallback.
 *   - share() needs transient user activation and WebKit consumes it, so send() calls navigator.share() synchronously (no await first).
 *   - Firefox: no file sharing (desktop has no navigator.share); canShare({files}) is false or missing -> download.
 *   - showDirectoryPicker: Chromium desktop 86+, Chrome Android 132+ (SAF, cloud apps rarely expose folders) -> offered on desktop only.
 *     Handles are structured-cloneable into IndexedDB; Chrome 122+ can persist the permission ("Her ziyarette izin ver").
 *   - Android: no share title is passed, so a target app can not take it as the file name (the PC filters on the file name).
 *
 * UMD, no dependencies. Browser: global OmsCloud. Node: module.exports (tests inject fakes with _setDeps).
 */
(function (root, factory) {
  'use strict';
  if (typeof module === 'object' && module.exports) module.exports = factory(root, false);
  else root.OmsCloud = factory(root, true);
})(typeof globalThis !== 'undefined' ? globalThis : (typeof self !== 'undefined' ? self : this), function (root, isBrowser) {
  'use strict';

  // ------------------------------------------------------------------ constants
  var FOLDER = 'OMS Saha';
  var PICKER_ID = 'oms-saha';                       // remembered start dir per id (<= 32 chars, [A-Za-z0-9_-])
  var SHARE_TITLE = 'OMS sepeti';
  var DB_NAME = 'oms-cloud', DB_STORE = 'kv', DB_KEY = 'folder';
  var EXT = '.omsb', TXT = '.txt';
  var TYPE_BIN = 'application/octet-stream', TYPE_TXT = 'text/plain';
  var MAX_SAME_NAME = 20;

  var MSG = {
    BAD_ARMOR: 'Gönderilecek sepet verisi geçersiz.',
    UNSUPPORTED: 'Bu tarayıcı bu yolu desteklemiyor. Dosyayı indirip OMS Saha klasörüne taşıyın.',
    NO_SHARE: 'Bu tarayıcı dosya paylaşamıyor. Dosyayı indirip OMS Saha klasörüne taşıyın.',
    NEED_FOLDER: 'Önce OMS Saha klasörünü seçin.',
    NEED_TAP: 'Tarayıcı bir dokunuş daha istiyor. Düğmeye yeniden basın.',
    NO_PERMISSION: 'Klasöre yazma izni verilmedi. Yeniden basıp İzin ver\'i seçin.',
    FOLDER_GONE: 'OMS Saha klasörü bulunamadı (silinmiş ya da taşınmış olabilir). Klasörü yeniden seçin.',
    DISK_FULL: 'Diskte yer kalmadı. Yer açıp yeniden deneyin.',
    WRITE_FAILED: 'Dosya klasöre yazılamadı. Yeniden deneyin ya da dosyayı indirin.',
    BUSY: 'Önceki paylaşma penceresi hâlâ açık. Onu kapatın ya da sayfayı yenileyin.',
    SHARE_FAILED: 'Paylaşılamadı. Yeniden deneyin ya da dosyayı indirin.',
    DOWNLOAD_HINT: 'Dosya İndirilenler klasörüne indi. Onu OMS Saha klasörüne taşıyın.',
    UNKNOWN: 'Gönderilemedi.'
  };

  // ------------------------------------------------------------------ environment (overridable for Node tests)
  var deps = {};
  function dep(k) { return Object.prototype.hasOwnProperty.call(deps, k) ? deps[k] : root[k]; }
  function nav() { return dep('navigator') || null; }
  function win() { return dep('window') || root; }
  function doc() { return dep('document') || null; }
  function now() { var f = dep('now'); return typeof f === 'function' ? f() : Date.now(); }

  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function stamp(d) {
    return '' + d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate()) + '_' + pad2(d.getHours()) + pad2(d.getMinutes()) + pad2(d.getSeconds());
  }
  function toDate(x) {
    if (x instanceof Date && !isNaN(x.getTime())) return x;
    if (typeof x === 'number' && isFinite(x)) return new Date(x);
    return new Date(now());
  }
  function errName(e) { return (e && (e.name || (e.constructor && e.constructor.name))) || ''; }
  function errText(e) { return String((e && (e.message || e.name)) || e || ''); }
  function fail(method, code, e) {
    var r = { ok: false, method: method, cancelled: code === 'CANCELLED', code: code, error: code === 'CANCELLED' ? '' : (MSG[code] || MSG.UNKNOWN) };
    if (e) r.detail = errName(e) + (errText(e) ? ': ' + errText(e) : '');
    return r;
  }

  // ------------------------------------------------------------------ platform + capabilities
  /** 'ios' (iPhone/iPad, also iPadOS "desktop" Safari), 'android' or 'desktop'. */
  function platform() {
    var n = nav(); if (!n) return 'desktop';
    var ua = String(n.userAgent || '');
    if (/Android/i.test(ua) || (n.userAgentData && /android/i.test(String(n.userAgentData.platform || '')))) return 'android';
    if (/iPhone|iPad|iPod/i.test(ua)) return 'ios';
    if (/Macintosh/i.test(ua) && (n.maxTouchPoints | 0) > 1) return 'ios';   // iPadOS 13+ requests the desktop site
    return 'desktop';
  }

  /** Normalizes a caller-given name to the base "....omsb" (drops ".txt"); null when it is not ours. */
  function baseName(name) {
    var s = String(name || '').replace(/^.*[\\/]/, '');
    if (/\.txt$/i.test(s)) s = s.slice(0, -4);
    if (!/\.omsb$/i.test(s)) return null;
    if (!/^[A-Za-z0-9_.-]{6,80}$/.test(s)) return null;
    return s;
  }

  /**
   * The basket file. kind 'txt' = "<base>.omsb.txt" text/plain, 'omsb' = "<base>.omsb" application/octet-stream.
   * Default kind: Android 'txt' (share targets need an allow-listed type), elsewhere 'omsb'. opts.name overrides the base name.
   * Returns {file: File|null, name, type, text}.
   */
  function fileFor(armor, date, kind, opts) {
    if (kind && typeof kind === 'object') { opts = kind; kind = null; }
    opts = opts || {};
    var k = kind === 'txt' || kind === 'omsb' ? kind : (platform() === 'android' ? 'txt' : 'omsb');
    var base = baseName(opts.name) || ('oms_sepet_' + stamp(toDate(date)) + EXT);
    var name = k === 'txt' ? base + TXT : base, type = k === 'txt' ? TYPE_TXT : TYPE_BIN, text = String(armor == null ? '' : armor);
    var F = dep('File'), file = null;
    if (typeof F === 'function') { try { file = new F([text], name, { type: type, lastModified: toDate(date).getTime() }); } catch (e) { file = null; } }
    return { file: file, name: name, type: type, text: text };
  }

  function canShareData(data) {
    var n = nav();
    if (!n || typeof n.share !== 'function' || typeof n.canShare !== 'function') return false;
    try { return !!n.canShare(data); } catch (e) { return false; }
  }
  /** First file variant the browser accepts for sharing, or null. */
  function shareCandidate(armor, date, name) {
    var order = platform() === 'ios' ? ['omsb', 'txt'] : ['txt', 'omsb'];   // Chromium: canShare() does not check the type list
    for (var i = 0; i < order.length; i++) {
      var f = fileFor(armor, date, order[i], { name: name });
      if (f.file && canShareData({ files: [f.file] })) return f;
    }
    return null;
  }
  function shareSupported() { return !!shareCandidate('OMSB1:x', new Date(2026, 0, 1)); }

  function folderSupported() {
    var w = win();
    if (!w || typeof w.showDirectoryPicker !== 'function') return false;
    if (w.isSecureContext === false) return false;
    return true;
  }
  function downloadSupported() {
    var d = doc(), U = dep('URL'), B = dep('Blob');
    return !!(d && typeof d.createElement === 'function' && U && typeof U.createObjectURL === 'function' && typeof B === 'function');
  }

  /** Usable ways to send, best first. Phones: share, download. Desktop: folder, download, share. */
  function methods() {
    var p = platform(), m = [], sh = shareSupported(), dl = downloadSupported();
    if (p === 'desktop') {
      if (folderSupported()) m.push('folder');
      if (dl) m.push('download');
      if (sh) m.push('share');
    } else {
      if (sh) m.push('share');
      if (dl) m.push('download');
    }
    return m;
  }

  // ------------------------------------------------------------------ folder handle store (IndexedDB, memory fallback)
  var folder = { loaded: false, loading: null, handle: null, name: null };

  function idb() { var x = dep('indexedDB'); return x && typeof x.open === 'function' ? x : null; }
  function idbOpen() {
    return new Promise(function (resolve, reject) {
      var f = idb(); if (!f) { reject(new Error('no indexedDB')); return; }
      var rq;
      try { rq = f.open(DB_NAME, 1); } catch (e) { reject(e); return; }
      rq.onupgradeneeded = function () { try { var db = rq.result; if (!db.objectStoreNames.contains(DB_STORE)) db.createObjectStore(DB_STORE); } catch (e) { } };
      rq.onsuccess = function () { resolve(rq.result); };
      rq.onerror = function () { reject(rq.error || new Error('idb open')); };
      rq.onblocked = function () { reject(new Error('idb blocked')); };
    });
  }
  function idbDo(mode, fn) {
    return idbOpen().then(function (db) {
      return new Promise(function (resolve, reject) {
        var out, tx;
        try { tx = db.transaction(DB_STORE, mode); out = fn(tx.objectStore(DB_STORE)); } catch (e) { try { db.close(); } catch (x) { } reject(e); return; }
        tx.oncomplete = function () { try { db.close(); } catch (x) { } resolve(out && 'result' in out ? out.result : undefined); };
        tx.onerror = tx.onabort = function () { try { db.close(); } catch (x) { } reject(tx.error || new Error('idb tx')); };
      });
    });
  }

  /** Loads the remembered folder (once). Resolves to the folder name or null; never rejects. */
  function ready() {
    if (folder.loaded) return Promise.resolve(folder.name);
    if (folder.loading) return folder.loading;
    folder.loading = idbDo('readonly', function (st) { return st.get(DB_KEY); }).then(function (v) {
      if (!folder.loaded && v && v.handle && typeof v.handle.getFileHandle === 'function') { folder.handle = v.handle; folder.name = String(v.name || v.handle.name || FOLDER); }
    }, function () { }).then(function () { folder.loaded = true; folder.loading = null; return folder.name; });
    return folder.loading;
  }
  function remember(handle) {
    folder.handle = handle; folder.name = String(handle.name || FOLDER); folder.loaded = true;
    return idbDo('readwrite', function (st) { return st.put({ handle: handle, name: folder.name, at: now() }, DB_KEY); }).then(function () { return true; }, function () { return false; });
  }
  /** Name of the remembered folder, or null (call ready() once at start; pickFolder/forgetFolder keep it current). */
  function folderName() { return folder.name; }
  /** Forgets the remembered folder (memory + IndexedDB). */
  function forgetFolder() {
    folder.handle = null; folder.name = null; folder.loaded = true;
    return idbDo('readwrite', function (st) { return st.delete(DB_KEY); }).then(function () { return true; }, function () { return true; });
  }
  /** True when a picked folder name looks like the PC's "OMS Saha" folder. */
  function isSahaName(name) { return /^oms[\s_-]*saha$/i.test(String(name || '').trim()); }

  /** Opens the folder picker (needs a tap). Resolves {ok, name, looksRight} or {ok:false, cancelled, code, error}. */
  function pickFolder() {
    if (!folderSupported()) return Promise.resolve(fail('folder', 'UNSUPPORTED'));
    var p;
    try { p = win().showDirectoryPicker({ id: PICKER_ID, mode: 'readwrite', startIn: 'documents' }); } catch (e) { p = Promise.reject(e); }
    return Promise.resolve(p).then(function (h) {
      if (!h || typeof h.getFileHandle !== 'function') return fail('folder', 'UNKNOWN');
      return permission(h).then(function (st) {
        if (st !== 'granted') return fail('folder', 'NO_PERMISSION');
        return remember(h).then(function (saved) {
          return { ok: true, method: 'folder', name: folder.name, looksRight: isSahaName(folder.name), saved: saved };
        });
      });
    }, function (e) { return fail('folder', mapFolderError(e, true), e); });
  }

  function permission(h) {
    var opt = { mode: 'readwrite' };
    if (typeof h.queryPermission !== 'function') return Promise.resolve('granted');   // no permission model: writes decide
    return Promise.resolve(h.queryPermission(opt)).then(function (s) {
      if (s === 'granted' || typeof h.requestPermission !== 'function') return s;
      return h.requestPermission(opt);                                                // needs the user's tap (activation)
    });
  }
  function mapFolderError(e, picking) {
    switch (errName(e)) {
      case 'AbortError': return 'CANCELLED';
      case 'NotFoundError': return 'FOLDER_GONE';
      case 'NotAllowedError': return 'NO_PERMISSION';
      case 'SecurityError': return picking ? 'NO_PERMISSION' : 'NEED_TAP';          // no activation left, or a blocked system folder
      case 'QuotaExceededError': return 'DISK_FULL';
      case 'TypeMismatchError': case 'NoModificationAllowedError': case 'InvalidModificationError': case 'InvalidStateError': return 'WRITE_FAILED';
      default: return picking ? 'UNKNOWN' : 'WRITE_FAILED';
    }
  }
  /** "a.omsb" -> first of a.omsb, a_2.omsb, a_3.omsb ... that does not exist yet. */
  function freeName(dir, name) {
    var dot = name.lastIndexOf('.omsb'), stem = name.slice(0, dot), ext = name.slice(dot);
    function attempt(i) {
      var n = i < 2 ? name : stem + '_' + i + ext;
      if (i > MAX_SAME_NAME) return Promise.resolve(stem + '_' + now() + ext);
      return Promise.resolve().then(function () { return dir.getFileHandle(n); }).then(function () { return attempt(i < 2 ? 2 : i + 1); }, function (e) {
        var k = errName(e);
        if (k === 'NotFoundError') return n;
        if (k === 'TypeMismatchError') return attempt(i < 2 ? 2 : i + 1);
        throw e;
      });
    }
    return attempt(1);
  }
  function writeFile(dir, name, text) {
    var B = dep('Blob');
    return freeName(dir, name).then(function (n) {
      return dir.getFileHandle(n, { create: true }).then(function (fh) {
        return fh.createWritable();   // Chromium writes a temp *.crswap and swaps on close, so the PC never sees a half file
      }).then(function (w) {
        var body = typeof B === 'function' ? new B([text], { type: TYPE_BIN }) : text;
        return Promise.resolve(w.write(body)).then(function () { return w.close(); }).then(function () { return n; }, function (e) {
          return Promise.resolve().then(function () { if (typeof w.abort === 'function') return w.abort(); }).then(function () { throw e; }, function () { throw e; });
        });
      });
    });
  }

  // ------------------------------------------------------------------ the three ways
  function viaShare(armor, date, name) {
    var n = nav(), f = shareCandidate(armor, date, name);
    if (!f) return Promise.resolve(fail('share', 'NO_SHARE'));
    var data = { files: [f.file] };
    if (platform() !== 'android') { data.title = SHARE_TITLE; if (!canShareData(data)) data = { files: [f.file] }; }
    var p;
    try { p = n.share(data); } catch (e) { p = Promise.reject(e); }   // called synchronously: keeps the tap's activation
    return Promise.resolve(p).then(function () {
      return { ok: true, method: 'share', cancelled: false, code: 'OK', error: '', name: f.name };
    }, function (e) {
      var k = errName(e), code = k === 'AbortError' ? 'CANCELLED' : k === 'NotAllowedError' ? 'NEED_TAP' : k === 'InvalidStateError' ? 'BUSY' : 'SHARE_FAILED';
      var r = fail('share', code, e); r.name = f.name; return r;
    });
  }

  function viaDownload(armor, date, name) {
    if (!downloadSupported()) return Promise.resolve(fail('download', 'UNSUPPORTED'));
    var f = fileFor(armor, date, null, { name: name });
    try {
      var d = doc(), U = dep('URL'), B = dep('Blob');
      var blob = new B([f.text], { type: f.type === TYPE_TXT ? TYPE_TXT : TYPE_BIN });
      var u = U.createObjectURL(blob), a = d.createElement('a');
      a.href = u; a.download = f.name; a.rel = 'noopener';
      if (a.style) a.style.display = 'none';
      var host = d.body || d.documentElement;
      if (host) host.appendChild(a);
      a.click();
      if (host) host.removeChild(a);
      var st = dep('setTimeout');
      if (typeof st === 'function') st(function () { try { U.revokeObjectURL(u); } catch (e) { } }, 60000);   // iOS saves asynchronously
    } catch (e) { return Promise.resolve(fail('download', 'UNKNOWN', e)); }
    return Promise.resolve({ ok: true, method: 'download', cancelled: false, code: 'OK', error: '', name: f.name, hint: MSG.DOWNLOAD_HINT });
  }

  function viaFolder(armor, date, name, pickIfNeeded) {
    if (!folderSupported()) return Promise.resolve(fail('folder', 'UNSUPPORTED'));
    var picked = false;
    return ready().then(function () {
      if (folder.handle) return folder.handle;
      if (!pickIfNeeded) return null;
      picked = true;
      return pickFolder().then(function (r) { if (!r.ok) throw { result: r }; return folder.handle; });
    }).then(function (h) {
      if (!h) return fail('folder', 'NEED_FOLDER');
      return Promise.resolve().then(function () { return permission(h); }).then(function (st) {
        if (st !== 'granted') return fail('folder', 'NO_PERMISSION');
        var f = fileFor(armor, date, 'omsb', { name: name });
        return writeFile(h, f.name, f.text).then(function (written) {
          return { ok: true, method: 'folder', cancelled: false, code: 'OK', error: '', name: written, folder: folder.name, picked: picked };
        });
      }).then(null, function (e) {
        var code = mapFolderError(e, false);
        if (code === 'FOLDER_GONE') return forgetFolder().then(function () { return fail('folder', code, e); });
        return fail('folder', code, e);
      });
    }, function (x) {
      if (x && x.result) return x.result;
      return fail('folder', 'UNKNOWN', x);
    });
  }

  /**
   * Sends one basket. opts: {armor, date (Date|ms), prefer ('share'|'folder'|'download'), fileName (optional base name),
   * pickIfNeeded (folder: open the picker when none is remembered)}. Call it straight from the click handler, before any await.
   * Resolves {ok, method, cancelled, error, code, name, ...}; never rejects. code: OK, CANCELLED, NEED_FOLDER, NEED_TAP,
   * NO_PERMISSION, FOLDER_GONE, DISK_FULL, WRITE_FAILED, NO_SHARE, BUSY, SHARE_FAILED, UNSUPPORTED, BAD_ARMOR, UNKNOWN.
   */
  function send(opts) {
    opts = opts || {};
    var armor = typeof opts.armor === 'string' ? opts.armor.trim() : '';
    if (!/^OMSB\d+:/.test(armor)) return Promise.resolve(fail('', 'BAD_ARMOR'));
    var date = toDate(opts.date), name = opts.fileName || null;
    var list = methods(), m = opts.prefer;
    if (m === 'folder' ? !folderSupported() : list.indexOf(m) < 0) m = list[0];
    if (m === 'share') return viaShare(armor, date, name);
    if (m === 'folder') return viaFolder(armor, date, name, !!opts.pickIfNeeded);
    if (m === 'download') return viaDownload(armor, date, name);
    return Promise.resolve(fail('', 'UNSUPPORTED'));
  }

  // ------------------------------------------------------------------ user guide (3 plain steps)
  var CLOUD = { onedrive: 'OneDrive', gdrive: 'Google Drive', icloud: 'iCloud Drive' };
  function cloudName(c) { return CLOUD[c] || ''; }
  function folderPath(c) {
    if (c === 'onedrive') return 'OneDrive > OMS Saha';
    if (c === 'gdrive') return 'Google Drive > Drive\'ım > OMS Saha';
    if (c === 'icloud') return 'iCloud Drive > OMS Saha';
    return 'OMS Saha';
  }

  /**
   * guide(platform, cloud, opts) -> [{title, text, icon, note?}] x 3. cloud: 'onedrive'|'gdrive'|'icloud'|''.
   * opts.share / opts.folder: what this browser can do (default: detected when platform is this device's platform).
   */
  function guide(p, cloud, opts) {
    p = p === 'ios' || p === 'android' ? p : 'desktop';
    var c = CLOUD[cloud] ? cloud : '';
    opts = opts || {};
    var here = p === platform();
    var canShare = opts.share != null ? !!opts.share : (here ? shareSupported() : true);
    var canFolder = opts.folder != null ? !!opts.folder : (here ? folderSupported() : true);
    var first = { title: 'Buluta gönder\'e basın', text: 'Paylaşma menüsü açılır.', icon: 'send' };

    if (p === 'ios' && canShare) {
      var loc = c ? CLOUD[c] : 'iCloud Drive (ya da OneDrive / Google Drive)';
      var s3 = { title: 'OMS Saha klasörüne kaydedin', text: loc + ' içindeki OMS Saha klasörünü açın, sağ üstteki Kaydet\'e dokunun.', icon: 'folder' };
      if (c === 'onedrive' || c === 'gdrive') s3.note = CLOUD[c] + ' listede yoksa Gözat ekranında sağ üstteki ••• > Düzenle\'den açın.';
      return [first,
        { title: 'Dosyalar\'a Kaydet\'i seçin', text: 'Menüyü yukarı kaydırın, Dosyalar\'a Kaydet\'e dokunun.', icon: 'save' },
        s3];
    }
    if (p === 'android' && canShare) {
      var app = c === 'onedrive' ? 'OneDrive' : c === 'gdrive' ? 'Drive' : '';
      var s2 = { title: app ? app + '\'ı seçin' : 'Drive\'ı ya da OneDrive\'ı seçin', text: 'Listede yoksa kaydırın ya da Diğer\'e dokunun.', icon: 'share' };
      if (c === 'icloud') s2.note = 'Android\'de iCloud uygulaması yok; Drive ya da OneDrive kullanın.';
      var s3a = c === 'onedrive'
        ? { title: 'OMS Saha klasörüne yükleyin', text: 'OMS Saha klasörünü açın, sonra onaylayın (Yükle / Kaydet).', icon: 'folder' }
        : { title: 'Klasör: OMS Saha, sonra Kaydet', text: 'Klasör satırına dokunup OMS Saha\'yı seçin, sonra Kaydet\'e basın.', icon: 'folder' };
      return [first, s2, s3a];
    }
    if (p === 'desktop' && canFolder) {
      return [
        { title: 'İlk seferde OMS Saha klasörünü seçin', text: 'Buluta gönder\'e basın. Açılan pencerede ' + folderPath(c) + ' klasörünü seçip Klasör seç\'e basın.', icon: 'folder' },
        { title: 'İzin sorulursa İzin ver\'e basın', text: 'Tarayıcı klasöre kaydetme izni ister. Her ziyarette izin ver seçeneği varsa onu seçin.', icon: 'key' },
        { title: 'Sonra tek tıkla gider', text: 'Sepet doğrudan OMS Saha klasörüne kaydedilir; PC onu kendisi alır.', icon: 'check' }];
    }
    // No share / no folder access: download, then move by hand.
    var where = p === 'desktop' ? folderPath(c) : (c ? CLOUD[c] + ' > OMS Saha' : 'bulutunuzdaki OMS Saha');
    return [
      { title: 'Buluta gönder\'e basın', text: 'Sepet dosyası İndirilenler klasörüne iner.', icon: 'download' },
      { title: 'Dosyayı OMS Saha klasörüne taşıyın', text: (p === 'desktop' ? 'İndirilenler\'deki oms_sepet_ dosyasını ' : 'Dosyalar uygulamasında İndirilenler\'deki oms_sepet_ dosyasını ') + where + ' klasörüne taşıyın.', icon: 'move' },
      { title: 'PC kendisi alır', text: 'Dosya buluta eşitlenince PC sepeti kendisi işler.', icon: 'check' }];
  }

  // ------------------------------------------------------------------ test hooks
  function _setDeps(o) { deps = o || {}; }
  function _reset() { folder = { loaded: false, loading: null, handle: null, name: null }; }

  var api = {
    FOLDER: FOLDER, MSG: MSG,
    platform: platform, methods: methods, fileFor: fileFor, send: send,
    ready: ready, pickFolder: pickFolder, folderName: folderName, forgetFolder: forgetFolder, folderSupported: folderSupported,
    shareSupported: shareSupported, downloadSupported: downloadSupported, isSahaName: isSahaName,
    guide: guide, cloudName: cloudName,
    _setDeps: _setDeps, _reset: _reset
  };
  if (isBrowser) { try { ready(); } catch (e) { } }   // warm the folder handle so send() needs no IndexedDB wait
  return api;
});
