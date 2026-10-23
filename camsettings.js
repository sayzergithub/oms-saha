/* OmsCam — shared camera settings module (V6). Served as /camsettings.js (LAN web UI) and /saha/camsettings.js (field page).
 * Contract: reviews/_PROMPTS_V6.md B6, V5_SAHA_SPEC.md §15. UMD: browser global `OmsCam`, CommonJS export for node tests.
 * Plain ES5 + Promise (old Samsung Internet / iOS WebViews). No dependencies.
 *
 * Browser support researched 2026-10-22 (W3C mediacapture-image ED, WebKit main IDL, MDN BCD 8.1.6):
 *  - Chrome / Samsung Internet on Android (Image Capture, Chrome 59+): focusMode, focusDistance, exposureMode, exposureCompensation,
 *    exposureTime (100 us units), iso, whiteBalanceMode, colorTemperature (K), brightness, contrast, saturation, sharpness, zoom
 *    (exposed only when the gUM call asked for {zoom:true} = PTZ permission), torch, pointsOfInterest (constraint only, 0..1).
 *    Chromium rejects an applyConstraints() call that mixes image-capture and non-image constraints.
 *  - Safari / every iOS browser (WebKit): capabilities focusDistance (read-only), whiteBalanceMode, zoom, torch (boolean, not a
 *    sequence); settable: whiteBalanceMode, zoom, torch (torch works iOS 17.5+, WebKit bug 243075). No focus/exposure control.
 *  - Firefox: getCapabilities 132+, width/height/frameRate/facingMode/deviceId only.
 * Everything is driven by getCapabilities(): a control that is not reported is not shown.
 */
(function (root, factory) {
  'use strict';
  var M = factory(root);
  if (typeof module === 'object' && module && module.exports) module.exports = M;
  if (root) root.OmsCam = M;
})(typeof self !== 'undefined' ? self : (typeof window !== 'undefined' ? window : this), function (g) {
  'use strict';

  var KEY = 'omscam.v1';
  var VERSION = 1;
  // Resolution presets (landscape sensor orientation). label MP is computed, not hand-written.
  var PRESETS = [
    { w: 640, h: 480 }, { w: 1280, h: 720 }, { w: 1920, h: 1080 }, { w: 2560, h: 1440 },
    { w: 2592, h: 1944 }, { w: 3840, h: 2160 }, { w: 4000, h: 3000 }
  ];
  var AUTO_RES = { w: 1920, h: 1080 };          // "Otomatik" target (what the scan pages used before V6)
  var FPS = [15, 24, 30, 60];
  var MODE_KEYS = ['focusMode', 'exposureMode', 'whiteBalanceMode'];
  var RANGE_KEYS = ['focusDistance', 'exposureCompensation', 'exposureTime', 'iso', 'colorTemperature',
    'brightness', 'contrast', 'saturation', 'sharpness', 'zoom'];
  var IMAGE_KEYS = MODE_KEYS.concat(RANGE_KEYS, ['torch', 'pointsOfInterest', 'pan', 'tilt']);
  var SAVED_KEYS = MODE_KEYS.concat(RANGE_KEYS, ['torch']);

  // ---------------------------------------------------------------- environment (overridable for node tests)
  var env = {
    win: function () { return g; },
    nav: function () { return g && g.navigator; },
    doc: function () { return g && g.document; },
    ls: function () { try { return g && g.localStorage; } catch (e) { return null; } },
    setTimeout: function (f, ms) { return (g && g.setTimeout ? g.setTimeout(f, ms) : setTimeout(f, ms)); },
    clearTimeout: function (t) { return (g && g.clearTimeout ? g.clearTimeout(t) : clearTimeout(t)); }
  };

  function noop() { }
  function has(o, k) { return o != null && Object.prototype.hasOwnProperty.call(o, k); }
  function assign(t) { for (var i = 1; i < arguments.length; i++) { var s = arguments[i]; if (s) for (var k in s) if (has(s, k)) t[k] = s[k]; } return t; }
  function clone(o) { try { return JSON.parse(JSON.stringify(o)); } catch (e) { return {}; } }
  function isNum(v) { return typeof v === 'number' && isFinite(v); }
  function md() { var n = env.nav(); return n && n.mediaDevices; }
  function supported() { var m = md(); return !!(m && m.getUserMedia); }
  function supportedConstraints() { var m = md(); try { return (m && m.getSupportedConstraints && m.getSupportedConstraints()) || {}; } catch (e) { return {}; } }
  function mkErr(name, msg) { var e = new Error(msg || name); e.name = name; return e; }

  // ---------------------------------------------------------------- persistence
  var st = fresh();
  var loaded = false;
  var autoPick = null;      // heuristic back camera {id,label}, memory only
  var cur = { stream: null, track: null, el: null };
  var tstate = (typeof WeakMap === 'function') ? new WeakMap() : null;
  var saveTimer = null;
  var lastRep = null;   // V13b: what the last open() asked for, what it got, what it had to relax (lastOpen())

  function fresh() { return { v: VERSION, sel: null, last: null, dev: {} }; }

  function load() {
    var s = null, ls = env.ls();
    try { if (ls) s = JSON.parse(ls.getItem(KEY) || 'null'); } catch (e) { s = null; }
    st = fresh();
    if (s && typeof s === 'object' && s.v === VERSION) {
      if (s.sel && typeof s.sel.id === 'string') st.sel = { id: s.sel.id, label: String(s.sel.label || '') };
      if (s.last && typeof s.last.id === 'string') st.last = { id: s.last.id, label: String(s.last.label || '') };
      if (s.dev && typeof s.dev === 'object') for (var k in s.dev) if (has(s.dev, k) && s.dev[k] && typeof s.dev[k] === 'object') st.dev[k] = sanitizeRec(s.dev[k]);
    }
    loaded = true;
    return prefs();
  }
  function ensure() { if (!loaded) load(); }

  function sanitizeRec(r) {
    var o = {};
    if (typeof r.label === 'string') o.label = r.label.slice(0, 120);
    if (typeof r.res === 'string' && (r.res === 'auto' || /^\d{2,5}x\d{2,5}$/.test(r.res))) o.res = r.res;
    if (isNum(r.fps) && r.fps >= 0 && r.fps <= 240) o.fps = r.fps;
    if (typeof r.tapFocus === 'boolean') o.tapFocus = r.tapFocus;
    for (var i = 0; i < MODE_KEYS.length; i++) { var m = r[MODE_KEYS[i]]; if (m === 'continuous' || m === 'single-shot' || m === 'manual') o[MODE_KEYS[i]] = m; }
    for (var j = 0; j < RANGE_KEYS.length; j++) if (isNum(r[RANGE_KEYS[j]])) o[RANGE_KEYS[j]] = r[RANGE_KEYS[j]];
    if (typeof r.torch === 'boolean') o.torch = r.torch;
    return o;
  }

  function saveNow() {
    saveTimer = null;
    var ls = env.ls(); if (!ls) return false;
    try { ls.setItem(KEY, JSON.stringify(st)); return true; } catch (e) { return false; }
  }
  function save(debounced) {
    if (!debounced) { if (saveTimer) { env.clearTimeout(saveTimer); saveTimer = null; } return saveNow(); }
    if (!saveTimer) saveTimer = env.setTimeout(saveNow, 300);
    return true;
  }

  function keyOfId(id) { return id ? String(id) : '_'; }
  function targetDev() { return st.sel || autoPick || null; }
  function curKey() { var t = st.sel || autoPick || st.last; return keyOfId(t && t.id); }
  function trackDeviceId(track) { try { return (track && track.getSettings && track.getSettings().deviceId) || ''; } catch (e) { return ''; } }
  function keyOfTrack(track) { var id = trackDeviceId(track); return id ? id : (track === cur.track && st.last ? keyOfId(st.last.id) : '_'); }
  function rec(key, create) {
    var r = st.dev[key];
    if (!r && create) r = st.dev[key] = {};
    return r || {};
  }

  /** Effective preferences of the camera the next open() targets (or of `track` when given). */
  function prefs(track) {
    ensure();
    var key = track ? keyOfTrack(track) : curKey();
    var r = rec(key);
    var o = clone(r);
    o.deviceId = key === '_' ? '' : key;
    o.selected = st.sel ? st.sel.id : '';
    if (typeof o.tapFocus !== 'boolean') o.tapFocus = true;
    if (!o.res) o.res = 'auto';
    if (!o.fps) o.fps = 0;
    return o;
  }

  /** Store one or more preference values for a camera (deviceId; '' = current target). */
  function setPref(deviceId, values, debounced) {
    ensure();
    var key = deviceId ? String(deviceId) : curKey();
    var r = rec(key, true);
    for (var k in values) if (has(values, k)) { if (values[k] === undefined || values[k] === null) delete r[k]; else r[k] = values[k]; }
    st.dev[key] = sanitizeRec(r);
    save(debounced);
    return prefs();
  }

  /** Select a camera for the next open(); '' / null = automatic (best back camera). The previous camera's resolution and frame
   *  rate are carried over when the new one has none saved (the user picked "60 fps", then another camera: still 60 fps). */
  function select(deviceId, label) {
    ensure();
    var from = rec(curKey()), toKey = deviceId ? String(deviceId) : keyOfId(autoPick && autoPick.id);
    if (toKey !== curKey()) {
      var to = rec(toKey, true);
      if (from.res && !to.res) to.res = from.res;
      if (from.fps && !to.fps) to.fps = from.fps;
    }
    if (!deviceId) st.sel = null;
    else {
      var lab = label;
      if (lab == null && lastList) for (var i = 0; i < lastList.length; i++) if (lastList[i].deviceId === deviceId) lab = lastList[i].rawLabel;
      st.sel = { id: String(deviceId), label: String(lab || '') };
    }
    save();
    return prefs();
  }

  // ---------------------------------------------------------------- camera list + heuristics
  var lastList = null;

  function facingOf(label) {
    var s = String(label || '').toLowerCase();
    if (!s) return '';
    if (/(^|[^a-z])(back|rear|environment|world)([^a-z]|$)|arka|rück|ruck|arri[eè]re|trasera|posteriore|achter|traseira/.test(s)) return 'environment';
    if (/(^|[^a-z])(front|user|selfie|facetime)([^a-z]|$)|(^|\s)ön(\s|$)|ön kamera|vorder|avant|frontal|anteriore|frente/.test(s)) return 'user';
    return '';
  }
  function lensOf(label) {
    var s = String(label || '').toLowerCase();
    if (/ultra/.test(s)) return 'ultra';
    if (/tele/.test(s)) return 'tele';
    if (/dual|triple|çift|üçlü|ikili|dreifach/.test(s)) return 'multi';
    if (/wide|geniş/.test(s)) return 'wide';
    return '';
  }
  /** Best back camera for close-up label scanning: plain wide back camera first, then multi-lens virtual ones, ultra / tele last. */
  function pickBack(list) {
    var best = null, bestScore = -1e9;
    for (var i = 0; i < (list || []).length; i++) {
      var c = list[i];
      if (c.facing !== 'environment') continue;
      var lens = c.lens != null ? c.lens : lensOf(c.rawLabel || c.label);
      var sc = 100;
      if (lens === 'multi') sc -= 10; else if (lens === 'ultra') sc -= 30; else if (lens === 'tele') sc -= 40;
      var m = /camera2?\s*(\d+)/i.exec(c.rawLabel || c.label || '');   // Android: "camera2 0, facing back" = main sensor
      if (m) sc -= Math.min(20, +m[1]);
      sc -= i * 0.01;
      if (sc > bestScore) { bestScore = sc; best = c; }
    }
    return best;
  }

  function listCameras() {
    ensure();
    var m = md();
    if (!m || !m.enumerateDevices) return Promise.resolve([]);
    return Promise.resolve().then(function () { return m.enumerateDevices(); }).then(function (ds) {
      var out = [], n = 0;
      for (var i = 0; i < (ds || []).length; i++) {
        var d = ds[i];
        if (!d || d.kind !== 'videoinput') continue;
        n++;
        var raw = d.label || '';
        var facing = facingOf(raw);
        if (!facing && d.getCapabilities) { try { var fm = d.getCapabilities().facingMode; if (fm && fm.length) facing = fm[0] === 'environment' || fm[0] === 'user' ? fm[0] : ''; } catch (e) { } }
        out.push({ deviceId: d.deviceId || '', groupId: d.groupId || '', label: raw || ('Kamera ' + n), rawLabel: raw, facing: facing, lens: lensOf(raw), best: false });
      }
      var b = pickBack(out);
      if (b) { b.best = true; if (b.deviceId) autoPick = { id: b.deviceId, label: b.rawLabel }; }
      remapIds(out);
      lastList = out;
      return out;
    }, function () { return []; });
  }

  // deviceId can change (cleared site data, some WebViews): follow the stored label to the new id and move its settings.
  function remapIds(list) {
    function find(ref) {
      if (!ref || !ref.id) return null;
      for (var i = 0; i < list.length; i++) if (list[i].deviceId === ref.id) return list[i];
      if (!ref.label) return null;
      var hit = null, cnt = 0;
      for (var j = 0; j < list.length; j++) if (list[j].rawLabel && list[j].rawLabel === ref.label) { hit = list[j]; cnt++; }
      return cnt === 1 ? hit : null;
    }
    var changed = false;
    ['sel', 'last'].forEach(function (k) {
      var ref = st[k], hit = find(ref);
      if (ref && hit && hit.deviceId && hit.deviceId !== ref.id) {
        if (st.dev[ref.id] && !st.dev[hit.deviceId]) { st.dev[hit.deviceId] = st.dev[ref.id]; delete st.dev[ref.id]; }
        ref.id = hit.deviceId; changed = true;
      }
    });
    if (changed) save();
  }

  // ---------------------------------------------------------------- resolution helpers
  function mp(w, h) { return (w * h) / 1e6; }
  function fmtNum(v, digits) {
    var s = v.toFixed(digits);
    if (digits > 0) s = s.replace(/\.?0+$/, '');
    return s.replace('.', ',');
  }
  function fmtMP(w, h) { var v = mp(w, h); return fmtNum(v, v >= 10 ? 0 : 1) + ' MP'; }
  function resLabel(w, h) { return w + '×' + h + ' · ' + fmtMP(w, h); }
  function parseRes(s) { var m = /^(\d+)x(\d+)$/.exec(s || ''); return m ? { w: +m[1], h: +m[2] } : null; }

  /** Presets the camera can deliver (by capabilities width/height max); adds the exact sensor maximum when it is not a preset. */
  function resPresets(caps) {
    var cw = caps && caps.width && caps.width.max, ch = caps && caps.height && caps.height.max;
    var out = [];
    if (!isNum(cw) || !isNum(ch) || cw <= 0 || ch <= 0) {
      for (var i = 0; i < 3; i++) out.push(presetObj(PRESETS[i].w, PRESETS[i].h, false));
      return out;   // unknown max: offer only the safe ones (ideal constraints never fail anyway)
    }
    var cmax = Math.max(cw, ch), cmin = Math.min(cw, ch);
    for (var j = 0; j < PRESETS.length; j++) {
      var p = PRESETS[j];
      if (p.w <= cmax && p.h <= cmin) out.push(presetObj(p.w, p.h, false));
    }
    var top = out.length ? out[out.length - 1] : null;
    if (!top || mp(cmax, cmin) > mp(top.w, top.h) * 1.05) out.push(presetObj(cmax, cmin, true));
    return out;
  }
  function presetObj(w, h, isMax) { return { id: w + 'x' + h, w: w, h: h, mp: mp(w, h), label: (isMax ? 'En yüksek: ' : '') + resLabel(w, h), max: !!isMax }; }

  /** Highest frame rate the camera reports, or null when unknown. Chrome on Android takes it per size from the camera driver
   *  (VideoCaptureCamera2.getDeviceSupportedFormats: getOutputMinFrameDuration) and Blink keeps the maximum over all sizes
   *  (ComputeCapabilitiesForVideoSource). Without the MANUAL_SENSOR capability Android reports 0 → Blink a near-zero minimum:
   *  that means "unknown", not "slow". */
  function fpsMax(caps) {
    var mx = caps && caps.frameRate && caps.frameRate.max;
    return isNum(mx) && mx >= 5 ? mx : null;
  }
  function fpsPresets(caps) {
    var mx = fpsMax(caps);
    var out = [];
    for (var i = 0; i < FPS.length; i++) if (mx === null || FPS[i] <= mx + 0.5) out.push(FPS[i]);
    return out;
  }

  // ---------------------------------------------------------------- constraints / open
  function ptzWanted() { var s = supportedConstraints(); return !!(s.zoom && (s.pan || s.tilt)); }   // Chromium only (PTZ permission)

  /** MediaStreamConstraints for the selected camera + saved resolution/fps. `extra` = MediaStreamConstraints or a bare video set. */
  function constraints(extra, opt) {
    ensure();
    opt = opt || {};
    var v = {}, t = targetDev();
    if (t && t.id && !opt.noDevice) v.deviceId = { exact: t.id }; else v.facingMode = { ideal: 'environment' };
    var r = rec(curKey());
    var res = opt.res !== undefined ? opt.res : (parseRes(r.res) || AUTO_RES);
    if (res) { v.width = { ideal: res.w }; v.height = { ideal: res.h }; }
    if (r.fps && !opt.noFps) v.frameRate = opt.fpsSoft ? { ideal: r.fps } : { ideal: r.fps, min: Math.max(1, Math.floor(r.fps * 0.8)) };
    if (ptzWanted() && !opt.noZoom) v.zoom = true;
    var out = { audio: false, video: v };
    if (extra && typeof extra === 'object') {
      var ev = null;
      if (has(extra, 'video') || has(extra, 'audio')) {
        if (has(extra, 'audio')) out.audio = extra.audio;
        if (extra.video === false) out.video = false;
        else if (extra.video && typeof extra.video === 'object') ev = extra.video;
      } else ev = extra;
      if (ev && out.video) {
        if (has(ev, 'deviceId')) delete v.facingMode;
        if (has(ev, 'facingMode') && !has(ev, 'deviceId')) delete v.deviceId;
        assign(v, ev);
      }
    }
    return out;
  }

  function stopStream(s) { if (!s) return; try { s.getTracks().forEach(function (t) { try { t.stop(); } catch (e) { } }); } catch (e) { } }
  function stopEl(el) { if (!el) return; try { if (el.srcObject && el.srcObject !== cur.stream) stopStream(el.srcObject); } catch (e) { } }

  /** Stop the stream the module opened last (and detach it from its video element). */
  function close() {
    stopStream(cur.stream);
    if (cur.el) { try { cur.el.srcObject = null; } catch (e) { } }
    cur = { stream: null, track: null, el: null };
  }

  function lowerRes(res) {
    // next preset strictly smaller than `res` (by pixel count); null when none
    var px = res ? res.w * res.h : 0, best = null;
    for (var i = PRESETS.length - 1; i >= 0; i--) if (PRESETS[i].w * PRESETS[i].h < px) { best = PRESETS[i]; break; }
    return best;
  }

  function getUM(c) {
    var m = md();
    return Promise.resolve().then(function () { return m.getUserMedia(c); });
  }

  /** getUserMedia + attach to <video> + apply saved settings. Retries with lower resolution / without the stale camera id. */
  function open(videoEl, extra) {
    ensure();
    if (!supported()) return Promise.reject(mkErr('NotSupportedError', 'getUserMedia yok'));
    var explicitDev = !!(extra && ((extra.video && typeof extra.video === 'object' && has(extra.video, 'deviceId')) || has(extra, 'deviceId')));
    // Android cannot open a second camera while one is running: release ours first.
    close();
    stopEl(videoEl);
    var pre = (!st.sel && !explicitDev && !autoPick) ? listCameras().then(noop, noop) : Promise.resolve();
    var opt = {}, tries = 0, firstErr = null, deviceDropped = false, busyWaits = 0;
    var savedRes = parseRes(rec(curKey()).res) || AUTO_RES;
    opt.res = savedRes;
    var rep = { at: Date.now(), trace: [], relaxed: [], fallback: null };
    function wantId() { var t = explicitDev ? null : targetDev(); return explicitDev ? exactOf(extra.video ? extra.video.deviceId : extra.deviceId) : (t && t.id) || ''; }
    function exactOf(d) { return d && typeof d === 'object' ? String(d.exact || d.ideal || '') : String(d || ''); }
    function relax() {   // one step, never the camera itself: fps firm → preferred → none, zoom off, then smaller sizes
      if (!opt.fpsSoft && rec(curKey()).fps) { opt.fpsSoft = true; rep.relaxed.push('fps-tercih'); return true; }
      if (!opt.noFps && rec(curKey()).fps) { opt.noFps = true; rep.relaxed.push('fps-yok'); return true; }
      if (!opt.noZoom && ptzWanted()) { opt.noZoom = true; rep.relaxed.push('zoom-yok'); return true; }
      if (opt.res) { opt.res = lowerRes(opt.res); rep.relaxed.push(opt.res ? 'boyut-' + opt.res.w + 'x' + opt.res.h : 'boyut-serbest'); return true; }
      return false;
    }
    function wait(ms) { var w = env.win(); return new Promise(function (r) { if (w && typeof w.setTimeout === 'function') w.setTimeout(r, ms); else setTimeout(r, ms); }); }   // real time: the camera needs it

    function attempt() {
      var c = constraints(extra, opt);
      if (deviceDropped && c.video) { delete c.video.deviceId; c.video.facingMode = { ideal: 'environment' }; }
      tries++;
      return getUM(c).catch(function (e) {
        var name = (e && e.name) || '';
        var cons = (e && e.constraint) || '';
        var v = c.video || {};
        rep.trace.push(name + (cons ? '(' + cons + ')' : ''));
        if (!firstErr) firstErr = e;
        if (tries >= 10 || name === 'NotAllowedError' || name === 'SecurityError' || name === 'PermissionDeniedError' || name === 'TypeError') throw e;
        // Android: the previous camera is still being released → same request again after a short pause (not a smaller one)
        if ((name === 'NotReadableError' || name === 'TrackStartError' || name === 'AbortError') && busyWaits < 2) { busyWaits++; return wait(350 * busyWaits).then(attempt); }
        var asked = v.deviceId && !deviceDropped ? exactOf(v.deviceId) : '';
        if (asked && (name === 'NotFoundError' || name === 'DevicesNotFoundError' || name === 'OverconstrainedError' || name === 'ConstraintNotSatisfiedError')) {
          // "deviceId" / "" / NotFound only means "camera gone" when it is really not in the list (and not findable by its label)
          return listCameras().then(function (list) {
            var listed = (list || []).some(function (d) { return d.deviceId === asked; });
            if (!listed) {
              var t = targetDev();
              if (!explicitDev && t && t.id && t.id !== asked) { rep.relaxed.push('kimlik-yenilendi'); return attempt(); }   // same camera, new id (label)
              if (!(list || []).length) return relax() ? attempt() : Promise.reject(firstErr || e);   // list unavailable: never guess
              deviceDropped = true;   // really gone: open the back camera for now, KEEP the user's choice and say so
              rep.fallback = 'Seçilen kamera bu cihazda bulunamadı (' + ((st.sel && st.sel.label) || asked.slice(0, 8)) + ')';
              return attempt();
            }
            if (cons === 'frameRate' && !opt.fpsSoft) { opt.fpsSoft = true; rep.relaxed.push('fps-tercih'); return attempt(); }
            if (cons === 'zoom' && !opt.noZoom) { opt.noZoom = true; rep.relaxed.push('zoom-yok'); return attempt(); }
            if (relax()) return attempt();
            throw firstErr || e;
          });
        }
        if (name === 'OverconstrainedError' || name === 'ConstraintNotSatisfiedError' || name === 'NotReadableError' || name === 'TrackStartError' || name === 'AbortError') {
          if (cons === 'frameRate' && !opt.fpsSoft) { opt.fpsSoft = true; rep.relaxed.push('fps-tercih'); return attempt(); }
          if (cons === 'zoom' && !opt.noZoom) { opt.noZoom = true; rep.relaxed.push('zoom-yok'); return attempt(); }
          if (relax()) return attempt();
        }
        throw firstErr || e;
      });
    }

    return pre.then(attempt).then(function (stream) {
      var track = (stream.getVideoTracks && stream.getVideoTracks()[0]) || null;
      cur = { stream: stream, track: track, el: videoEl || null };
      var ws = wantId(), got = trackDeviceId(track), gs = getSet(track), pr = rec(curKey());
      rep.want = { id: ws, label: (st.sel && ws === st.sel.id && st.sel.label) || '', res: pr.res || 'auto', fps: pr.fps || 0 };
      rep.got = { id: got, label: track ? String(track.label || '') : '', w: gs.width || 0, h: gs.height || 0, fps: isNum(gs.frameRate) ? gs.frameRate : 0, facing: gs.facingMode || '' };
      rep.deviceOk = !ws || !got || ws === got;
      lastRep = rep;
      if (videoEl) {
        try { videoEl.muted = true; videoEl.setAttribute('muted', ''); videoEl.setAttribute('playsinline', ''); videoEl.playsInline = true; videoEl.setAttribute('autoplay', ''); } catch (e) { }
        try { videoEl.srcObject = stream; } catch (e) { }
      }
      var playP = Promise.resolve();
      if (videoEl && videoEl.play) { try { playP = Promise.resolve(videoEl.play()).catch(noop); } catch (e) { } }
      return playP.then(function () {
        if (track) {
          snapshotDefaults(track);
          var id = trackDeviceId(track);
          if (id) {
            st.last = { id: id, label: String(track.label || '') };
            var r = rec(id, true); if (track.label) r.label = String(track.label).slice(0, 120);
            save();
          }
          return apply(track).then(function (a) { rep.skipped = (a && a.skipped) || []; return a; });
        }
      }).then(function () {
        return track ? fpsGuard(track).then(noop, noop) : null;
      }).then(function () {
        var gs2 = getSet(track); if (rep.got) { rep.got.w = gs2.width || rep.got.w; rep.got.h = gs2.height || rep.got.h; if (isNum(gs2.frameRate)) rep.got.fps = gs2.frameRate; }
        return { stream: stream, track: track, caps: getCaps(track), settings: gs2, report: rep };
      });
    });
  }

  /** One Turkish sentence when the last open did not give exactly what was chosen; '' when it did. */
  function openIssue(r) {
    r = r || lastRep; if (!r || !r.got) return '';
    var out = [];
    if (r.fallback) out.push(r.fallback + '; şimdilik otomatik kamera açıldı');
    else if (!r.deviceOk) out.push('Tarayıcı seçilen kamera yerine başka bir kamera açtı (' + (r.got.label || 'bilinmeyen') + ')');
    var want = parseRes(r.want && r.want.res);
    if (want && r.got.w && r.got.w * r.got.h < want.w * want.h * 0.9) out.push('istenen ' + want.w + '×' + want.h + ' yerine ' + r.got.w + '×' + r.got.h);
    if (r.want && r.want.fps && r.got.fps > 0 && r.got.fps < r.want.fps * 0.8) out.push('istenen ' + r.want.fps + ' kare/s yerine ' + Math.round(r.got.fps));
    if (r.skipped && r.skipped.length) out.push('uygulanamayan ayar: ' + r.skipped.map(function (k) { return NAMES[k] || k; }).join(', '));
    return out.join(' · ');
  }

  // ---------------------------------------------------------------- applying settings
  function getCaps(track) { try { return (track && track.getCapabilities && track.getCapabilities()) || {}; } catch (e) { return {}; } }
  function getSet(track) { try { return (track && track.getSettings && track.getSettings()) || {}; } catch (e) { return {}; } }
  function live(track) { return !!(track && track.readyState !== 'ended' && track.applyConstraints); }

  function tsOf(track) {
    if (!track) return { acc: {}, mode: null, defaults: null };
    var s = tstate ? tstate.get(track) : track.__omscam;
    if (!s) { s = { acc: {}, mode: null, defaults: null }; if (tstate) tstate.set(track, s); else try { track.__omscam = s; } catch (e) { } }
    return s;
  }
  function snapshotDefaults(track) {
    var s = tsOf(track); if (s.defaults) return s.defaults;
    var set = getSet(track), d = {};
    for (var i = 0; i < SAVED_KEYS.length; i++) { var k = SAVED_KEYS[i]; if (set[k] !== undefined) d[k] = set[k]; }
    d.torch = false;
    s.defaults = d;
    return d;
  }

  var MODE_ORDER = ['continuous', 'single-shot', 'manual', 'none'];
  function modeList(caps, k) {
    var v = caps && caps[k];
    if (!(v && typeof v.length === 'number' && typeof v !== 'string')) return [];
    return Array.prototype.slice.call(v).sort(function (a, b) { return MODE_ORDER.indexOf(a) - MODE_ORDER.indexOf(b); });   // auto first
  }
  function hasMode(caps, k, m) { return modeList(caps, k).indexOf(m) >= 0; }
  function rangeOf(caps, k) { var r = caps && caps[k]; return (r && isNum(r.min) && isNum(r.max) && r.max > r.min) ? r : null; }
  function torchCap(caps) { var t = caps && caps.torch; return t === true || (t && typeof t.length === 'number' && t.indexOf && t.indexOf(true) >= 0); }
  function clampTo(r, v) {
    if (!r) return v;
    v = Math.max(r.min, Math.min(r.max, v));
    if (isNum(r.step) && r.step > 0) { v = r.min + Math.round((v - r.min) / r.step) * r.step; v = Math.max(r.min, Math.min(r.max, v)); }
    return +v.toFixed(6);
  }

  function baseConstraints(track) {
    var b = {}, c = {};
    try { c = (track.getConstraints && track.getConstraints()) || {}; } catch (e) { c = {}; }
    for (var k in c) if (has(c, k) && k !== 'advanced' && IMAGE_KEYS.indexOf(k) < 0) b[k] = c[k];
    return b;
  }

  /** Apply one image-capture constraint set. Chromium: image-only call; WebKit: merged with the current base set + earlier image values. */
  function applySet(track, set) {
    if (!live(track)) return Promise.reject(mkErr('InvalidStateError', 'track'));
    var s = tsOf(track);
    if (!s.mode) { var caps = getCaps(track); s.mode = (caps.focusMode || caps.exposureMode) ? 'img' : 'merged'; }
    function run(mode) {
      var c;
      if (mode === 'img') c = { advanced: [assign({}, set)] };
      else { c = baseConstraints(track); c.advanced = [assign({}, s.acc, set)]; }
      return Promise.resolve().then(function () { return track.applyConstraints(c); });
    }
    return run(s.mode).then(null, function (e) {
      var other = s.mode === 'img' ? 'merged' : 'img';
      return run(other).then(function () { s.mode = other; }, function () { throw e; });
    }).then(function () {
      for (var k in set) if (has(set, k) && k !== 'pointsOfInterest') s.acc[k] = set[k];
      return fpsGuard(track).then(function () { return true; }, function () { return true; });
    });
  }

  /** Re-asserts the chosen resolution + frame rate when the track runs clearly below the chosen fps. Promise<boolean> (true = re-applied).
   *  Image settings (focus, zoom, exposure) are camera state in Chromium/WebKit and survive a format-only applyConstraints. */
  function fpsGuard(track) {
    if (!live(track)) return Promise.resolve(false);
    var s = tsOf(track); if (s.guarding) return Promise.resolve(false);
    var pr = prefs(track), set = getSet(track);
    // Chromium/Android reports 0 when the camera lacks MANUAL_SENSOR (VideoCaptureCamera2): unknown, not slow → leave it alone
    if (!pr.fps || !isNum(set.frameRate) || set.frameRate <= 0 || set.frameRate >= pr.fps * 0.8) return Promise.resolve(false);
    var res = parseRes(pr.res) || AUTO_RES;
    var c = { width: { ideal: res.w }, height: { ideal: res.h }, frameRate: { ideal: pr.fps, min: Math.max(1, Math.floor(pr.fps * 0.8)) } };
    s.guarding = true;
    return Promise.resolve().then(function () { return track.applyConstraints(c); })
      .catch(function () { return track.applyConstraints({ width: c.width, height: c.height, frameRate: { ideal: pr.fps } }); })
      .then(function () { s.guarding = false; return true; }, function () { s.guarding = false; return false; });
  }

  // Apply a group; on failure fall back to its pieces (mode first) so one unsupported value does not drop the rest.
  function applyGroup(track, set, applied, skipped) {
    var keys = Object.keys(set);
    if (!keys.length) return Promise.resolve();
    return applySet(track, set).then(function () { applied.push.apply(applied, keys); }, function () {
      if (keys.length === 1) { skipped.push(keys[0]); return; }
      var p = Promise.resolve();
      keys.forEach(function (k) { p = p.then(function () { var one = {}; one[k] = set[k]; return applySet(track, one).then(function () { applied.push(k); }, function () { skipped.push(k); }); }); });
      return p;
    });
  }

  /** Build the ordered constraint groups for a stored record against the track's capabilities. */
  function groupsFor(r, caps) {
    var out = [], gset;
    var fr = rangeOf(caps, 'focusDistance');
    if (r.focusMode && hasMode(caps, 'focusMode', r.focusMode)) {
      gset = { focusMode: r.focusMode };
      if (r.focusMode === 'manual' && isNum(r.focusDistance) && fr) gset.focusDistance = clampTo(fr, r.focusDistance);
      out.push(gset);
    }
    gset = {};
    if (r.exposureMode && hasMode(caps, 'exposureMode', r.exposureMode)) gset.exposureMode = r.exposureMode;
    var manualExp = gset.exposureMode === 'manual';
    if (manualExp) {
      if (isNum(r.exposureTime) && rangeOf(caps, 'exposureTime')) gset.exposureTime = clampTo(rangeOf(caps, 'exposureTime'), r.exposureTime);
      if (isNum(r.iso) && rangeOf(caps, 'iso')) gset.iso = clampTo(rangeOf(caps, 'iso'), r.iso);
    } else if (isNum(r.exposureCompensation) && rangeOf(caps, 'exposureCompensation')) gset.exposureCompensation = clampTo(rangeOf(caps, 'exposureCompensation'), r.exposureCompensation);
    if (Object.keys(gset).length) out.push(gset);
    gset = {};
    if (r.whiteBalanceMode && hasMode(caps, 'whiteBalanceMode', r.whiteBalanceMode)) {
      gset.whiteBalanceMode = r.whiteBalanceMode;
      if (r.whiteBalanceMode === 'manual' && isNum(r.colorTemperature) && rangeOf(caps, 'colorTemperature')) gset.colorTemperature = clampTo(rangeOf(caps, 'colorTemperature'), r.colorTemperature);
      out.push(gset);
    }
    ['brightness', 'contrast', 'saturation', 'sharpness', 'zoom'].forEach(function (k) {
      if (isNum(r[k]) && rangeOf(caps, k)) { var o = {}; o[k] = clampTo(rangeOf(caps, k), r[k]); out.push(o); }
    });
    if (r.torch === true && torchCap(caps)) out.push({ torch: true });
    return out;
  }

  /** Apply the saved advanced settings of the track's camera. Never rejects; resolves {applied:[], skipped:[]}. */
  function apply(track) {
    ensure();
    var applied = [], skipped = [];
    if (!live(track)) return Promise.resolve({ applied: applied, skipped: skipped });
    var caps = getCaps(track), r = rec(keyOfTrack(track));
    var p = Promise.resolve();
    groupsFor(r, caps).forEach(function (gs) { p = p.then(function () { return applyGroup(track, gs, applied, skipped); }); });
    return p.then(function () { return { applied: applied, skipped: skipped }; }, function () { return { applied: applied, skipped: skipped }; });
  }

  // ---------------------------------------------------------------- quick actions
  function poiSupported(track) { return !!supportedConstraints().pointsOfInterest && !!getCaps(track).focusMode; }
  function tapFocusSupported(track) { var c = getCaps(track); return poiSupported(track) || hasMode(c, 'focusMode', 'single-shot'); }

  /** Focus (and meter) at a normalized point (0..1, top-left origin). Resolves true when the camera accepted it. */
  function focusAt(track, x, y) {
    if (!live(track)) return Promise.resolve(false);
    var caps = getCaps(track), poi = poiSupported(track);
    var single = hasMode(caps, 'focusMode', 'single-shot'), cont = hasMode(caps, 'focusMode', 'continuous');
    if (!poi && !single) return Promise.resolve(false);
    x = Math.max(0, Math.min(1, +x || 0)); y = Math.max(0, Math.min(1, +y || 0));
    var pt = [{ x: +x.toFixed(4), y: +y.toFixed(4) }];
    var set = {};
    if (poi) set.pointsOfInterest = pt;
    if (single) set.focusMode = 'single-shot'; else if (cont) set.focusMode = 'continuous';
    var used = null;
    return applySet(track, set).then(function () { used = set.focusMode; return true; }, function () {
      if (!poi) return false;
      return applySet(track, { pointsOfInterest: pt }).then(function () { return true; }, function () {
        return single ? applySet(track, { focusMode: 'single-shot' }).then(function () { used = 'single-shot'; return true; }, function () { return false; }) : false;
      });
    }).then(function (ok) {
      var want = rec(keyOfTrack(track)).focusMode || (cont ? 'continuous' : '');
      if (ok && used === 'single-shot' && want === 'continuous' && cont) {
        var s = tsOf(track); var tok = (s.focusTok = (s.focusTok || 0) + 1);
        env.setTimeout(function () {
          if (s.focusTok !== tok || !live(track)) return;
          var back = { focusMode: 'continuous' }; if (poi) back.pointsOfInterest = pt;
          applySet(track, back).then(noop, function () { applySet(track, { focusMode: 'continuous' }).then(noop, noop); });
        }, 2500);
      }
      return ok;
    });
  }

  /** Torch on/off. Rejects when unsupported. save=true stores it for this camera (panel does). */
  function torch(track, on, saveIt) {
    if (!live(track) || !torchCap(getCaps(track))) return Promise.reject(mkErr('NotSupportedError', 'torch'));
    on = !!on;
    return applySet(track, { torch: on }).then(function () {
      if (saveIt) setPref(keyOfTrack(track), { torch: on ? true : null });
      return on;
    });
  }

  /** Zoom (clamped to capabilities). Rejects when unsupported. Saved for this camera unless save === false. */
  function zoom(track, v, saveIt) {
    var r = rangeOf(getCaps(track), 'zoom');
    if (!live(track) || !r || !isNum(+v)) return Promise.reject(mkErr('NotSupportedError', 'zoom'));
    var z = clampTo(r, +v);
    return applySet(track, { zoom: z }).then(function () {
      if (saveIt !== false) setPref(keyOfTrack(track), { zoom: z }, true);
      return z;
    });
  }

  /** Forget this camera's saved settings and restore what the camera reported when it was opened. Resolves needRestart. */
  function resetDefaults(track) {
    ensure();
    var key = track ? keyOfTrack(track) : curKey();
    var old = rec(key);
    var needRestart = !!((old.res && old.res !== 'auto') || old.fps);
    var label = old.label;
    st.dev[key] = label ? { label: label } : {};
    save();
    if (!live(track)) return Promise.resolve({ needRestart: needRestart });
    var d = snapshotDefaults(track), caps = getCaps(track), applied = [], skipped = [];
    var r = {};
    for (var i = 0; i < MODE_KEYS.length; i++) {
      var mk = MODE_KEYS[i];
      if (old[mk] || d[mk]) r[mk] = (d[mk] && hasMode(caps, mk, d[mk])) ? d[mk] : (hasMode(caps, mk, 'continuous') ? 'continuous' : undefined);
    }
    for (var j = 0; j < RANGE_KEYS.length; j++) { var rk = RANGE_KEYS[j]; if (old[rk] !== undefined && isNum(d[rk])) r[rk] = d[rk]; }
    var groups = groupsFor(r, caps);
    if (torchCap(caps)) groups.push({ torch: false });
    var p = Promise.resolve();
    groups.forEach(function (gs) { p = p.then(function () { return applyGroup(track, gs, applied, skipped); }); });
    return p.then(function () { return { needRestart: needRestart, applied: applied, skipped: skipped }; });
  }

  /** Map a client point on a <video> to normalized frame coords (object-fit contain/cover/fill aware). */
  function tapPoint(videoEl, clientX, clientY, mirrored) {
    var r = videoEl.getBoundingClientRect(), vw = videoEl.videoWidth || r.width, vh = videoEl.videoHeight || r.height;
    var fit = 'contain';
    try { var w = env.win(); fit = (w && w.getComputedStyle ? w.getComputedStyle(videoEl).objectFit : '') || 'contain'; } catch (e) { }
    var sx, sy, ox, oy;
    if (fit === 'fill') { sx = r.width / vw; sy = r.height / vh; }
    else { var s = fit === 'cover' ? Math.max(r.width / vw, r.height / vh) : Math.min(r.width / vw, r.height / vh); if (fit === 'none') s = 1; sx = sy = s; }
    ox = (r.width - vw * sx) / 2; oy = (r.height - vh * sy) / 2;
    var x = (clientX - r.left - ox) / (vw * sx), y = (clientY - r.top - oy) / (vh * sy);
    if (mirrored) x = 1 - x;
    return { x: Math.max(0, Math.min(1, x)), y: Math.max(0, Math.min(1, y)), inside: x >= 0 && x <= 1 && y >= 0 && y <= 1 };
  }

  /** Tap-to-focus on a preview: honours prefs().tapFocus, draws a focus ring. opts.mirrored: ()=>bool. Returns an unbind function. */
  function bindTapFocus(videoEl, getTrack, opts) {
    opts = opts || {};
    injectCss();
    var host = opts.host || videoEl.parentNode, downX = 0, downY = 0, downT = 0;
    function down(e) { downX = e.clientX; downY = e.clientY; downT = Date.now(); }
    function up(e) {
      if (e.button > 0) return;
      if (Math.abs(e.clientX - downX) > 12 || Math.abs(e.clientY - downY) > 12 || Date.now() - downT > 800) return;   // a drag / long press is not a tap
      var t = getTrack && getTrack();
      if (!live(t) || !prefs(t).tapFocus || !tapFocusSupported(t)) return;
      var pt = tapPoint(videoEl, e.clientX, e.clientY, opts.mirrored ? !!opts.mirrored() : false);
      if (!pt.inside) return;
      ring(e.clientX, e.clientY, 'busy');
      focusAt(t, pt.x, pt.y).then(function (ok) { ring(e.clientX, e.clientY, ok ? 'ok' : 'fail'); if (opts.onFocus) try { opts.onFocus(ok, pt); } catch (x) { } });
    }
    var ringEl = null, ringTimer = null;
    function ring(cx, cy, state) {
      if (!host || !host.getBoundingClientRect) return;
      var d = env.doc(); if (!d) return;
      try { var w = env.win(); if (w && w.getComputedStyle && w.getComputedStyle(host).position === 'static') host.style.position = 'relative'; } catch (e) { }
      if (!ringEl) { ringEl = d.createElement('div'); ringEl.className = 'omscam-ring'; ringEl.setAttribute('aria-hidden', 'true'); }
      if (ringEl.parentNode !== host) host.appendChild(ringEl);
      var hr = host.getBoundingClientRect();
      ringEl.style.left = (cx - hr.left) + 'px'; ringEl.style.top = (cy - hr.top) + 'px';
      ringEl.className = 'omscam-ring omscam-ring-' + state;
      if (ringTimer) env.clearTimeout(ringTimer);
      ringTimer = env.setTimeout(function () { if (ringEl && ringEl.parentNode) ringEl.parentNode.removeChild(ringEl); }, state === 'busy' ? 4000 : 900);
    }
    var tgt = opts.target || videoEl;
    tgt.addEventListener('pointerdown', down);
    tgt.addEventListener('pointerup', up);
    return function () { tgt.removeEventListener('pointerdown', down); tgt.removeEventListener('pointerup', up); if (ringEl && ringEl.parentNode) ringEl.parentNode.removeChild(ringEl); };
  }

  function errorText(e) {
    var n = (e && e.name) || '';
    if (n === 'NotAllowedError' || n === 'PermissionDeniedError' || n === 'SecurityError') return 'Kamera izni verilmedi. Tarayıcının site ayarlarından bu sayfaya kamera izni verin.';
    if (n === 'NotReadableError' || n === 'TrackStartError') return 'Kamera açılamadı; başka bir uygulama veya sekme kullanıyor olabilir.';
    if (n === 'NotFoundError' || n === 'DevicesNotFoundError') return 'Kamera bulunamadı.';
    if (n === 'OverconstrainedError' || n === 'ConstraintNotSatisfiedError') return 'Bu kamera istenen ayarı desteklemiyor.';
    if (n === 'NotSupportedError') return 'Bu tarayıcı kamerayı sayfaya açmıyor.';
    return 'Kamera açılamadı' + (n ? ' (' + n + ')' : '') + '.';
  }

  // ---------------------------------------------------------------- panel
  var CSS = [
    '.omscam{--oc-text:var(--text,#1c2024);--oc-text2:var(--text-2,#60646c);--oc-surface:var(--surface,#fff);--oc-surface2:var(--surface-2,#f0f0f3);',
    '--oc-hover:var(--hover,#e8e8ec);--oc-line:var(--line,#d9d9e0);--oc-linectl:var(--line-ctl,#8b8d98);--oc-accent:var(--accent,#0b68c1);',
    '--oc-onaccent:var(--on-accent,#fff);--oc-accenttext:var(--accent-text,#0b68c1);--oc-accentsoft:var(--accent-soft,#e6f4fe);--oc-focus:var(--focus,#0b68c1);',
    '--oc-warn:var(--warn,#8f5300);--oc-warnsoft:var(--warn-soft,#fff7c2);--oc-err:var(--err,#ce2c31);--oc-errsoft:var(--err-soft,#feebec);',
    '--oc-ok:var(--ok,#1a7349);--oc-oksoft:var(--ok-soft,#e6f6eb);--oc-r:var(--r-md,8px);--oc-h:var(--h-ctl,44px);',
    'font-family:var(--font,system-ui,-apple-system,"Segoe UI",Roboto,sans-serif);font-size:var(--fs-body,16px);line-height:1.5;color:var(--oc-text);}',
    '@media (prefers-color-scheme:dark){:root:not([data-theme=light]) .omscam{--oc-text:var(--text,#edeef0);--oc-text2:var(--text-2,#b0b4ba);--oc-surface:var(--surface,#18191b);',
    '--oc-surface2:var(--surface-2,#212225);--oc-hover:var(--hover,#272a2d);--oc-line:var(--line,#363a3f);--oc-linectl:var(--line-ctl,#696e77);--oc-accent:var(--accent,#2870bd);',
    '--oc-accenttext:var(--accent-text,#70b8ff);--oc-accentsoft:var(--accent-soft,#0d2847);--oc-focus:var(--focus,#70b8ff);--oc-warn:var(--warn,#ffca16);',
    '--oc-warnsoft:var(--warn-soft,#302008);--oc-err:var(--err,#ff9592);--oc-errsoft:var(--err-soft,#3b1219);--oc-ok:var(--ok,#3dd68c);--oc-oksoft:var(--ok-soft,#132d21);}}',
    '.omscam *,.omscam *::before,.omscam *::after{box-sizing:border-box}',
    '.omscam-sec{padding:16px 0;border-top:1px solid var(--oc-line);min-width:0}.omscam-sec:first-child{border-top:0;padding-top:0}',
    '.omscam-h{font-size:var(--fs-ttl,18px);line-height:24px;font-weight:600;margin:0 0 8px;color:var(--oc-text)}',
    '.omscam-lab{display:flex;justify-content:space-between;align-items:baseline;gap:8px;font-size:var(--fs-sm,14px);line-height:20px;color:var(--oc-text2);margin:12px 0 4px}',
    '.omscam-val{color:var(--oc-text);font-variant-numeric:tabular-nums;text-align:right;white-space:nowrap}',
    '.omscam select{display:block;width:100%;min-height:var(--oc-h);font:inherit;font-size:16px;padding:0 12px;border:1px solid var(--oc-linectl);border-radius:var(--oc-r);background:var(--oc-surface);color:var(--oc-text);max-width:100%}',
    '.omscam button,.omscam input,.omscam select{font-family:inherit}',
    '.omscam-btn{display:inline-flex;align-items:center;justify-content:center;gap:6px;min-height:var(--oc-h);min-width:44px;padding:0 16px;border-radius:var(--oc-r);border:1px solid var(--oc-linectl);background:var(--oc-surface2);color:var(--oc-text);font-size:16px;line-height:20px;cursor:pointer;-webkit-tap-highlight-color:transparent;touch-action:manipulation}',
    '.omscam-btn:disabled{opacity:.55;cursor:default}',
    '.omscam-pri{background:var(--oc-accent);border-color:var(--oc-accent);color:var(--oc-onaccent);font-weight:600}',
    '.omscam-row{display:flex;flex-wrap:wrap;gap:8px;margin-top:8px}.omscam-row>.omscam-btn{flex:1 1 140px}',
    '.omscam-seg{display:flex;flex-wrap:wrap;border:1px solid var(--oc-linectl);border-radius:var(--oc-r);overflow:hidden;margin-top:4px}',
    '.omscam-seg label{position:relative;flex:1 1 30%;min-width:96px;display:flex;cursor:pointer;margin:0}',
    '.omscam-seg label+label{border-left:1px solid var(--oc-linectl)}',
    '.omscam-seg input{position:absolute;opacity:0;width:1px;height:1px;margin:0;pointer-events:none}',
    '.omscam-seg span{flex:1;display:flex;align-items:center;justify-content:center;min-height:var(--oc-h);padding:4px 10px;text-align:center;font-size:15px;line-height:18px;color:var(--oc-text);background:var(--oc-surface)}',
    '.omscam-seg input:checked+span{background:var(--oc-accent);color:var(--oc-onaccent);font-weight:600}',
    '.omscam-seg input:checked+span::before{content:"\\2713\\FE0E";margin-right:6px}',
    '.omscam-seg input:focus-visible+span{outline:3px solid var(--oc-focus);outline-offset:-3px}',
    '.omscam-sl{display:grid;grid-template-columns:44px minmax(0,1fr) 44px;align-items:center;gap:8px}',
    '.omscam-sl .omscam-btn{padding:0;font-size:22px;line-height:1;font-weight:600}',
    '.omscam input[type=range]{width:100%;height:32px;margin:0;accent-color:var(--oc-accent);touch-action:pan-y;background:transparent;cursor:pointer}',
    '.omscam-sw{display:flex;align-items:center;gap:12px;min-height:var(--oc-h);margin-top:8px;cursor:pointer;color:var(--oc-text)}',
    '.omscam-sw input{width:22px;height:22px;margin:0;flex:none;accent-color:var(--oc-accent)}',
    '.omscam-note{margin:8px 0 0;font-size:var(--fs-sm,14px);line-height:20px;color:var(--oc-text2)}',
    '.omscam-warn{margin:8px 0 0;padding:8px 12px;border-radius:var(--oc-r);background:var(--oc-warnsoft);color:var(--oc-text);font-size:var(--fs-sm,14px);line-height:20px;border-left:4px solid var(--oc-warn)}',
    '.omscam-kv{display:grid;grid-template-columns:auto minmax(0,1fr);gap:4px 12px;margin:0;font-size:var(--fs-sm,14px);line-height:20px}',
    '.omscam-kv dt{color:var(--oc-text2)}.omscam-kv dd{margin:0;color:var(--oc-text);font-variant-numeric:tabular-nums;overflow-wrap:anywhere}',
    '.omscam-status{min-height:20px;margin-top:8px;font-size:var(--fs-sm,14px);line-height:20px;color:var(--oc-text2)}',
    '.omscam-status.omscam-e{color:var(--oc-text);background:var(--oc-errsoft);border-left:4px solid var(--oc-err);padding:8px 12px;border-radius:var(--oc-r)}',
    '.omscam-status.omscam-o{color:var(--oc-text)}',
    '.omscam-status.omscam-w{color:var(--oc-text);border-left:4px solid #ab6400;padding:8px 12px;background:rgba(255,197,61,.16);border-radius:var(--oc-r)}',
    '.omscam-save{margin-top:16px;padding:12px;border:1px solid var(--oc-line);border-radius:var(--oc-r);background:var(--oc-accentsoft)}',
    '.omscam-save.omscam-savetop{margin:0 0 16px}',
    '.omscam-sum{font-size:14px;line-height:20px;color:var(--oc-text2);margin-bottom:8px;overflow-wrap:anywhere}',
    '.omscam-save .omscam-btn{width:100%;min-height:48px}',
    '.omscam-reset.omscam-arm{background:var(--oc-errsoft);border-color:var(--oc-err);color:var(--oc-text);font-weight:600}',
    '@media (hover:hover){.omscam-btn:not(:disabled):hover{background:var(--oc-hover)}.omscam-pri:not(:disabled):hover{background:var(--accent-hover,var(--oc-accent))}.omscam-seg label:hover input:not(:checked)+span{background:var(--oc-hover)}}',
    '.omscam-btn:focus-visible,.omscam select:focus-visible,.omscam input[type=range]:focus-visible,.omscam-sw input:focus-visible{outline:3px solid var(--oc-focus);outline-offset:2px}',
    '.omscam-ring{position:absolute;width:72px;height:72px;margin:-36px 0 0 -36px;border:2px solid #fff;border-radius:50%;box-shadow:0 0 0 1px rgba(0,0,0,.55),inset 0 0 0 1px rgba(0,0,0,.35);pointer-events:none;z-index:5;transition:transform .2s,opacity .2s}',
    '.omscam-ring-busy{transform:scale(1.15)}.omscam-ring-ok{border-color:#3dd68c;transform:scale(.85)}.omscam-ring-fail{border-color:#ffca16;border-style:dashed;opacity:.8}',
    '@media (prefers-reduced-motion:reduce){.omscam-ring{transition:none}}'
  ].join('\n');

  function injectCss() {
    var d = env.doc(); if (!d || !d.head || d.getElementById('omscam-css')) return;
    var s = d.createElement('style'); s.id = 'omscam-css'; s.textContent = CSS; d.head.appendChild(s);
  }

  var uid = 0;
  var MODE_LABEL = { continuous: 'Otomatik (sürekli)', 'single-shot': 'Tek sefer', manual: 'Elle' };
  var EXP_LABEL = { continuous: 'Otomatik', 'single-shot': 'Tek sefer', manual: 'Elle' };
  var NOSUP = 'Bu tarayıcı bu ayarı desteklemiyor.';

  function fmtExposure(v) {   // 100 us units → shutter text
    var sec = v / 10000;
    if (sec >= 1) return fmtNum(sec, 1) + ' s';
    return '1/' + Math.max(1, Math.round(1 / sec)) + ' s';
  }
  function fmtVal(k, v) {
    if (!isNum(v)) return '—';
    if (k === 'exposureTime') return fmtExposure(v);
    if (k === 'colorTemperature') return Math.round(v) + ' K';
    if (k === 'zoom') return fmtNum(v, 1) + '×';
    if (k === 'focusDistance') return fmtNum(v, 2) + ' m';
    if (k === 'exposureCompensation') return (v > 0 ? '+' : '') + fmtNum(v, 1) + ' EV';
    if (k === 'iso') return String(Math.round(v));
    return fmtNum(v, Math.abs(v) < 10 && v % 1 ? 1 : 0);
  }
  var NAMES = {
    focusDistance: 'Odak mesafesi', exposureCompensation: 'Pozlama telafisi', exposureTime: 'Pozlama süresi', iso: 'ISO',
    colorTemperature: 'Renk sıcaklığı', brightness: 'Parlaklık', contrast: 'Kontrast', saturation: 'Doygunluk', sharpness: 'Keskinlik', zoom: 'Yakınlaştırma (zoom)'
  };

  /** Full Turkish settings panel. opts.getTrack: ()=>MediaStreamTrack|null, opts.onRestart: ()=>Promise (reopen camera). Returns {refresh, destroy}. */
  function renderPanel(container, opts) {
    ensure();
    opts = opts || {};
    var d = env.doc();
    if (!d || !container) return { refresh: noop, destroy: noop };
    injectCss();
    var id = 'omscam' + (++uid);
    var root = d.createElement('div');
    root.className = 'omscam'; root.setAttribute('lang', 'tr');
    container.innerHTML = '';
    container.appendChild(root);
    var statusEl = null, kvEl = null, timer = null, dead = false, lastTrack = null, cams = null, busy = false, armTimer = null;
    var throttles = {};

    function el(tag, cls, text) { var e = d.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }
    function getTrack() { var t = null; try { t = opts.getTrack ? opts.getTrack() : cur.track; } catch (e) { t = null; } return live(t) ? t : null; }
    function status(text, kind) {
      if (!statusEl) return;
      statusEl.className = 'omscam-status' + (kind === 'err' ? ' omscam-e' : kind === 'warn' ? ' omscam-w' : kind === 'ok' ? ' omscam-o' : '');
      statusEl.setAttribute('role', kind === 'err' ? 'alert' : 'status');
      statusEl.textContent = (kind === 'err' ? 'Hata: ' : '') + (text || '');
    }
    function section(title) { var s = el('section', 'omscam-sec'); var h = el('h3', 'omscam-h', title); h.id = id + '-h' + root.children.length; s.setAttribute('aria-labelledby', h.id); s.appendChild(h); root.appendChild(s); return s; }
    function note(parent, text, warn) { var p = el('p', warn ? 'omscam-warn' : 'omscam-note', text); parent.appendChild(p); return p; }
    function btn(parent, text, cls, fn) { var b = el('button', 'omscam-btn' + (cls ? ' ' + cls : ''), text); b.type = 'button'; b.addEventListener('click', fn); parent.appendChild(b); return b; }

    function restart(msg) {
      if (!opts.onRestart) { status('Ayar kaydedildi; kamera yeniden açılınca geçerli olur.', 'ok'); return Promise.resolve(); }
      busy = true; status(msg || 'Kamera yeniden açılıyor…');
      return Promise.resolve().then(function () { return opts.onRestart(); }).then(function () {
        busy = false; cams = null; build();
        var chk = deliveredCheck(getTrack());
        status(chk.text, chk.ok ? 'ok' : 'warn');
      }, function (e) { busy = false; build(); status(errorText(e), 'err'); });
    }

    function seg(parent, name, labelText, modes, labels, value, onPick) {
      var lab = el('div', 'omscam-lab'); lab.appendChild(el('span', null, labelText)); lab.id = id + '-' + name + '-l'; parent.appendChild(lab);
      var box = el('div', 'omscam-seg'); box.setAttribute('role', 'radiogroup'); box.setAttribute('aria-labelledby', lab.id);
      modes.forEach(function (m) {
        var l = d.createElement('label'), i = d.createElement('input');
        i.type = 'radio'; i.name = id + '-' + name; i.value = m; i.id = id + '-' + name + '-' + m; i.checked = m === value;
        i.addEventListener('change', function () { if (i.checked) onPick(m); });
        l.appendChild(i); l.appendChild(el('span', null, labels[m] || m)); box.appendChild(l);
      });
      parent.appendChild(box);
      return box;
    }

    // Live slider: − range +, value label; throttled apply (one call in flight, last value wins).
    function slider(parent, k, r, value, onLive) {
      var lab = el('label', 'omscam-lab'); var inId = id + '-' + k;
      lab.setAttribute('for', inId);
      lab.appendChild(el('span', null, NAMES[k] || k));
      var out = el('span', 'omscam-val', fmtVal(k, value)); lab.appendChild(out); parent.appendChild(lab);
      var row = el('div', 'omscam-sl');
      var step = isNum(r.step) && r.step > 0 ? r.step : (r.max - r.min) / 100;
      var bigStep = Math.max(step, (r.max - r.min) / 20);
      var minus = btn(row, '−', null, function () { set(+inp.value - bigStep, true); });
      minus.setAttribute('aria-label', (NAMES[k] || k) + ' azalt');
      var inp = d.createElement('input'); inp.type = 'range'; inp.id = inId; inp.min = r.min; inp.max = r.max; inp.step = step;
      inp.value = isNum(value) ? value : r.min;
      row.appendChild(inp);
      var plus = btn(row, '+', null, function () { set(+inp.value + bigStep, true); });
      plus.setAttribute('aria-label', (NAMES[k] || k) + ' artır');
      parent.appendChild(row);
      function set(v, commit) { v = clampTo(r, v); inp.value = v; out.textContent = fmtVal(k, v); inp.setAttribute('aria-valuetext', fmtVal(k, v)); push(k, v, onLive, commit); }
      inp.setAttribute('aria-valuetext', fmtVal(k, +inp.value));
      inp.addEventListener('input', function () { set(+inp.value, false); });
      inp.addEventListener('change', function () { set(+inp.value, true); });
      return inp;
    }
    function push(k, v, fn, commit) {
      var t = throttles[k] || (throttles[k] = { busy: false, pend: null });
      t.pend = v;
      if (commit) t.commit = true;
      if (t.busy) return;
      (function run() {
        if (t.pend === null) return;
        var val = t.pend; t.pend = null; t.busy = true;
        Promise.resolve().then(function () { return fn(val); }).then(function () {
          t.busy = false;
          if (t.pend !== null) run(); else { t.commit = false; refreshKv(); }
        }, function () { t.busy = false; status((NAMES[k] || k) + ' uygulanamadı.', 'err'); if (t.pend !== null) run(); });
      })();
    }

    function liveSet(track, values, saveVals) {
      return applySet(track, values).then(function () { setPref(keyOfTrack(track), saveVals || values, true); status('Uygulandı ve kaydedildi.', 'ok'); });
    }

    // ------------------------------------------------ build
    function build() {
      if (dead) return;
      var focusId = d.activeElement && root.contains(d.activeElement) ? d.activeElement.id : '';
      var prevStatus = statusEl ? [statusEl.className, statusEl.textContent, statusEl.getAttribute('role')] : null;
      root.innerHTML = '';
      statusEl = null;
      throttles = {};
      var track = getTrack();
      lastTrack = track;
      var caps = track ? getCaps(track) : {}, set = track ? getSet(track) : {};
      var key = track ? keyOfTrack(track) : curKey();
      var r = rec(key);

      // 1. camera choice
      var s1 = section('Kamera');
      var lab = el('label', 'omscam-lab'); lab.setAttribute('for', id + '-cam'); lab.appendChild(el('span', null, 'Kullanılacak kamera')); s1.appendChild(lab);
      var sel = d.createElement('select'); sel.id = id + '-cam'; s1.appendChild(sel);
      sel.addEventListener('change', function () { select(sel.value || null); restart('Kamera değiştiriliyor…'); });
      var row = el('div', 'omscam-row'); s1.appendChild(row);
      btn(row, 'Listeyi yenile', null, function () { cams = null; fillCams(sel, null, camNote); });
      var camNote = note(s1, '');
      fillCams(sel, null, camNote);

      // 2. resolution + fps
      var s2 = section('Çözünürlük ve kare hızı');
      var lr = el('label', 'omscam-lab'); lr.setAttribute('for', id + '-res'); lr.appendChild(el('span', null, 'Çözünürlük (megapiksel)')); s2.appendChild(lr);
      var rs = d.createElement('select'); rs.id = id + '-res'; s2.appendChild(rs);
      var o0 = d.createElement('option'); o0.value = 'auto'; o0.textContent = 'Otomatik (' + AUTO_RES.w + '×' + AUTO_RES.h + ')'; rs.appendChild(o0);
      var pres = resPresets(caps), curRes = r.res || 'auto', found = curRes === 'auto';
      pres.forEach(function (p) { var o = d.createElement('option'); o.value = p.id; o.textContent = p.label; if (p.id === curRes) found = true; rs.appendChild(o); });
      if (!found) { var pr = parseRes(curRes); if (pr) { var ox = d.createElement('option'); ox.value = curRes; ox.textContent = resLabel(pr.w, pr.h) + ' (kayıtlı)'; rs.appendChild(ox); } }
      rs.value = curRes;
      rs.addEventListener('change', function () { setPref(key, { res: rs.value === 'auto' ? null : rs.value }); restart('Çözünürlük değiştiriliyor…'); });
      var lf = el('label', 'omscam-lab'); lf.setAttribute('for', id + '-fps'); lf.appendChild(el('span', null, 'Kare hızı')); s2.appendChild(lf);
      var fs = d.createElement('select'); fs.id = id + '-fps'; s2.appendChild(fs);
      var f0 = d.createElement('option'); f0.value = '0'; f0.textContent = 'Otomatik'; fs.appendChild(f0);
      var fps = fpsPresets(caps), fmx = track ? fpsMax(caps) : null;
      fps.forEach(function (f) { var o = d.createElement('option'); o.value = String(f); o.textContent = f + ' kare/s'; fs.appendChild(o); });
      if (r.fps && fps.indexOf(r.fps) < 0) {   // saved earlier (e.g. 60 picked while the preview was closed): show it, do not silently drop it
        var ofx = d.createElement('option'); ofx.value = String(r.fps); ofx.textContent = r.fps + ' kare/s (kayıtlı · bu kamera vermiyor)'; fs.appendChild(ofx);
      }
      fs.value = String(r.fps || 0); if (fs.value !== String(r.fps || 0)) fs.value = '0';
      fs.addEventListener('change', function () { setPref(key, { fps: +fs.value || null }); restart('Kare hızı değiştiriliyor…'); });
      // Android picks an auto-exposure fps RANGE closest to the target (VideoCaptureCamera2.getClosestFramerateRange); in dim light the
      // camera lengthens exposure and runs slower than chosen, so the HUD then reads e.g. '30 fps · kamera 60'.
      if (fmx !== null && fmx < FPS[FPS.length - 1] - 0.5) note(s2, 'Bu kamera tarayıcıya en fazla ' + Math.round(fmx) + ' kare/s veriyor; bu sınırı telefonun kamera yazılımı belirler. ' +
        'Daha yüksek kare hızı bu kamerada sayfadan kullanılamaz' + (cams && cams.length > 1 ? ' (başka bir kamera seçip deneyebilirsiniz).' : '.'), !!(r.fps && r.fps > fmx + 0.5));
      if (r.fps >= 50 && (fmx === null || fmx >= r.fps - 0.5)) note(s2, 'Işık azsa kamera pozlamayı uzatıp kare hızını kendiliğinden düşürebilir (ör. 30). Taramada sol üstte "30 fps · kamera 60" görürseniz Işık düğmesini açın ya da Pozlama bölümünde pozlamayı elle kısa süreye alın.');
      note(s2, 'Yüksek çözünürlük küçük kodları okumayı kolaylaştırır ama telefonu yavaşlatabilir ve pili daha çok kullanır. Telefon istenen değeri vermezse en yakını seçilir; gerçek değer aşağıda görünür.');
      if (track) {
        var cw = caps.width && caps.width.max, chh = caps.height && caps.height.max;
        var lim = note(s2, (isNum(cw) && isNum(chh) ? 'Canlı görüntü bu kamerada en fazla ' + resLabel(Math.max(cw, chh), Math.min(cw, chh)) + '. ' : '') +
          'Bu sınırı telefonun kamera yazılımı belirler: tarayıcıya yalnız video için ayrılmış boyutları verir. Kameranın 12 veya 50 MP olması canlı görüntüye yansımaz.');
        photoCaps(track).then(function (mx) {
          if (dead || !lim.parentNode) return;
          var vmp = isNum(cw) && isNum(chh) ? mp(cw, chh) : 0;
          if (mx && mp(mx.w, mx.h) <= vmp * 1.1) lim.textContent += ' Bu kamerada fotoğraf da aynı boyutta (' + fmtMP(mx.w, mx.h) + ').';
          else if (mx) lim.textContent += ' Fotoğraf ise ' + resLabel(Math.max(mx.w, mx.h), Math.min(mx.w, mx.h)) + ' çekilebilir: taramada "HD fotoğraf" düğmesi bu çözünürlükte çekip okur.' +
            (mp(mx.w, mx.h) < 30 ? ' 50 MP kameralar çoğu telefonda pikselleri birleştirerek bu boyutta fotoğraf verir; tam 50 MP modu tarayıcıya açık değildir.' : '');
          else lim.textContent += ' Bu tarayıcı tam çözünürlüklü fotoğrafı sayfa içinden çekemiyor; "Fotoğraf çek" telefonun kamera uygulamasını açar.';
        });
      }

      if (!track) {
        var sx = section('Odak, pozlama ve ışık');
        note(sx, 'Kamera açık değil. Önizlemeyi açınca bu kameranın desteklediği ayarlar burada görünür.');
        tail(null, r, key, prevStatus);
        return restoreFocus(focusId);
      }


      // 3. focus
      var s3 = section('Odak');
      var fm = modeList(caps, 'focusMode').filter(function (m) { return MODE_LABEL[m]; });
      var curFm = (r.focusMode && fm.indexOf(r.focusMode) >= 0) ? r.focusMode : (set.focusMode || '');
      if (fm.length) {
        seg(s3, 'focus', 'Odak modu', fm, MODE_LABEL, curFm, function (m) {
          var vals = { focusMode: m };
          liveSet(track, vals).then(build, function () { status('Odak modu uygulanamadı.', 'err'); build(); });
        });
        var fr = rangeOf(caps, 'focusDistance');
        if (curFm === 'manual' && fr) slider(s3, 'focusDistance', fr, isNum(set.focusDistance) ? set.focusDistance : r.focusDistance, function (v) { return liveSet(track, { focusMode: 'manual', focusDistance: v }); });
        else if (curFm === 'manual') note(s3, 'Bu kamera elle odak mesafesi ayarı sunmuyor.');
        if (fm.indexOf('single-shot') >= 0) { var rr = el('div', 'omscam-row'); s3.appendChild(rr); btn(rr, 'Şimdi odakla', null, function () { status('Odaklanıyor…'); focusAt(track, 0.5, 0.5).then(function (ok) { status(ok ? 'Ortadaki noktaya odaklandı.' : 'Odaklanamadı.', ok ? 'ok' : 'err'); refreshKv(); }); }); }
      } else note(s3, NOSUP + (rangeOf(caps, 'focusDistance') ? ' (Odak otomatik; mesafe yalnız okunabiliyor.)' : ''));
      if (tapFocusSupported(track)) {
        var sw = el('label', 'omscam-sw'); var cb = d.createElement('input'); cb.type = 'checkbox'; cb.id = id + '-tap'; cb.setAttribute('role', 'switch');
        cb.checked = r.tapFocus !== false;
        cb.addEventListener('change', function () { setPref(key, { tapFocus: cb.checked }); status(cb.checked ? 'Dokunarak odaklama açık: önizlemede bir noktaya dokunun.' : 'Dokunarak odaklama kapalı.', 'ok'); });
        sw.appendChild(cb); sw.appendChild(el('span', null, 'Dokunarak odakla (önizlemede dokunulan noktaya)')); s3.appendChild(sw);
      } else if (fm.length) note(s3, 'Dokunarak odaklama bu tarayıcıda desteklenmiyor.');

      // 4. exposure
      var s4 = section('Pozlama');
      var em = modeList(caps, 'exposureMode').filter(function (m) { return m === 'continuous' || m === 'manual'; });
      var curEm = (r.exposureMode && em.indexOf(r.exposureMode) >= 0) ? r.exposureMode : (set.exposureMode || '');
      var shown4 = false;
      if (em.length > 1 || (em.length === 1 && em[0] === 'manual')) {
        shown4 = true;
        seg(s4, 'exp', 'Pozlama modu', em, EXP_LABEL, curEm, function (m) { liveSet(track, { exposureMode: m }).then(build, function () { status('Pozlama modu uygulanamadı.', 'err'); build(); }); });
      }
      if (curEm !== 'manual' && rangeOf(caps, 'exposureCompensation')) { shown4 = true; slider(s4, 'exposureCompensation', rangeOf(caps, 'exposureCompensation'), isNum(set.exposureCompensation) ? set.exposureCompensation : (isNum(r.exposureCompensation) ? r.exposureCompensation : 0), function (v) { return liveSet(track, { exposureCompensation: v }); }); }
      if (curEm === 'manual') {
        if (rangeOf(caps, 'exposureTime')) { shown4 = true; slider(s4, 'exposureTime', rangeOf(caps, 'exposureTime'), isNum(set.exposureTime) ? set.exposureTime : r.exposureTime, function (v) { return liveSet(track, { exposureMode: 'manual', exposureTime: v }); }); }
        if (rangeOf(caps, 'iso')) { shown4 = true; slider(s4, 'iso', rangeOf(caps, 'iso'), isNum(set.iso) ? set.iso : r.iso, function (v) { return liveSet(track, { exposureMode: 'manual', iso: v }); }); }
      } else if (em.indexOf('manual') >= 0 && (rangeOf(caps, 'exposureTime') || rangeOf(caps, 'iso'))) note(s4, 'Pozlama süresi ve ISO "Elle" modunda ayarlanır.');
      if (!shown4) note(s4, NOSUP);

      // 5. white balance
      var s5 = section('Beyaz denge');
      var wm = modeList(caps, 'whiteBalanceMode').filter(function (m) { return m === 'continuous' || m === 'manual' || m === 'single-shot'; });
      var curWm = (r.whiteBalanceMode && wm.indexOf(r.whiteBalanceMode) >= 0) ? r.whiteBalanceMode : (set.whiteBalanceMode || '');
      if (wm.length > 1 || (wm.length === 1 && rangeOf(caps, 'colorTemperature'))) {
        seg(s5, 'wb', 'Beyaz denge modu', wm, EXP_LABEL, curWm, function (m) { liveSet(track, { whiteBalanceMode: m }).then(build, function () { status('Beyaz denge uygulanamadı.', 'err'); build(); }); });
        if (curWm === 'manual' && rangeOf(caps, 'colorTemperature')) slider(s5, 'colorTemperature', rangeOf(caps, 'colorTemperature'), isNum(set.colorTemperature) ? set.colorTemperature : r.colorTemperature, function (v) { return liveSet(track, { whiteBalanceMode: 'manual', colorTemperature: v }); });
        else if (curWm === 'manual') note(s5, 'Elle modu o anki beyaz dengeyi kilitler; bu kamera renk sıcaklığını sayı olarak ayarlatmıyor.');
      } else note(s5, NOSUP);

      // 6. image
      var s6 = section('Görüntü');
      var shown6 = false;
      ['brightness', 'contrast', 'saturation', 'sharpness'].forEach(function (k) {
        var rg = rangeOf(caps, k); if (!rg) return; shown6 = true;
        slider(s6, k, rg, isNum(set[k]) ? set[k] : r[k], function (v) { var o = {}; o[k] = v; return liveSet(track, o); });
      });
      if (!shown6) note(s6, 'Parlaklık, kontrast, doygunluk ve keskinlik: ' + NOSUP.charAt(0).toLowerCase() + NOSUP.slice(1));

      // 7. zoom + torch
      var s7 = section('Yakınlaştırma ve fener');
      var zr = rangeOf(caps, 'zoom'), shown7 = false;
      if (zr) { shown7 = true; slider(s7, 'zoom', zr, isNum(set.zoom) ? set.zoom : (isNum(r.zoom) ? r.zoom : zr.min), function (v) { return zoom(track, v, true).then(function () { status('Uygulandı ve kaydedildi.', 'ok'); }); }); }
      if (torchCap(caps)) {
        shown7 = true;
        var ts = tsOf(track);
        var on = has(ts.acc, 'torch') ? !!ts.acc.torch : !!set.torch;
        var tw = el('label', 'omscam-sw'); var tc = d.createElement('input'); tc.type = 'checkbox'; tc.id = id + '-torch'; tc.setAttribute('role', 'switch'); tc.checked = on;
        tc.addEventListener('change', function () {
          var want = tc.checked;
          torch(track, want, true).then(function () { status(want ? 'Fener açık (kaydedildi; kamera her açılışta yakar).' : 'Fener kapalı.', 'ok'); refreshKv(); },
            function () { tc.checked = !want; status('Fener bu kamerada kullanılamıyor.', 'err'); });
        });
        tw.appendChild(tc); tw.appendChild(el('span', null, 'Fener (ışık)')); s7.appendChild(tw);
      }
      if (!shown7) note(s7, NOSUP);

      if (!caps.focusMode && !caps.exposureMode) {
        note(s3, 'Bu tarayıcı odak ve pozlamayı sayfalara açmıyor. iPhone ve iPad\'de (Safari dahil tüm tarayıcılar) yalnız yakınlaştırma, fener ve beyaz denge ayarlanabilir; en çok ayar Android\'de Chrome ile sunulur.', true);
      }
      tail(track, r, key, prevStatus);
      restoreFocus(focusId);
    }

    function tail(track, r, key, prevStatus) {
      var s8 = section('Şu anki gerçek değerler');
      kvEl = el('dl', 'omscam-kv'); s8.appendChild(kvEl);
      refreshKv();
      var s9 = section('Varsayılanlar');
      note(s9, 'Bu kamera için kaydedilen tüm ayarları siler ve kameranın kendi otomatik ayarlarına döner.');
      var rowR = el('div', 'omscam-row'); s9.appendChild(rowR);
      var bR = btn(rowR, 'Varsayılanlara dön', 'omscam-reset', function () {
        if (!bR.classList.contains('omscam-arm')) {
          bR.classList.add('omscam-arm'); bR.textContent = 'Emin misiniz? Onaylamak için tekrar dokunun';
          if (armTimer) env.clearTimeout(armTimer);
          armTimer = env.setTimeout(function () { bR.classList.remove('omscam-arm'); bR.textContent = 'Varsayılanlara dön'; }, 5000);
          return;
        }
        if (armTimer) env.clearTimeout(armTimer);
        bR.disabled = true;
        resetDefaults(track).then(function (res) {
          if (res.needRestart && track) return restart('Varsayılan çözünürlükle yeniden açılıyor…');
          build(); status('Varsayılanlara dönüldü.', 'ok');
        }, function () { build(); status('Varsayılanlara dönülemedi.', 'err'); });
      });
      statusEl = el('div', 'omscam-status'); statusEl.setAttribute('role', 'status'); statusEl.setAttribute('aria-live', 'polite');
      root.appendChild(statusEl);
      if (prevStatus) { statusEl.className = prevStatus[0]; statusEl.textContent = prevStatus[1]; statusEl.setAttribute('role', prevStatus[2] || 'status'); }
      saveBar(false); saveBar(true);
    }

    /** Saved preference vs what the open track really delivers (the phone may refuse 60 fps or a size). */
    function deliveredCheck(t) {
      var pr = prefs(t || undefined), s = t ? getSet(t) : {}, parts = [], ok = true;
      var want = parseRes(pr.res);
      if (t) {
        if (isNum(s.width) && isNum(s.height)) parts.push(s.width + '×' + s.height + ' (' + fmtMP(s.width, s.height) + ')');
        if (isNum(s.frameRate) && s.frameRate > 0) parts.push(fmtNum(s.frameRate, 0) + ' kare/s');
        if (want && isNum(s.width) && Math.max(s.width, s.height) * Math.min(s.width, s.height) < want.w * want.h * 0.9) ok = false;
        if (pr.fps && isNum(s.frameRate) && s.frameRate > 0 && s.frameRate < pr.fps * 0.8) ok = false;
      }
      var iss = t && lastRep && cur.track === t ? openIssue(lastRep) : '';
      if (iss) return { ok: false, text: 'Kaydedildi, ama: ' + iss + '.' + (lastRep && lastRep.fallback ? ' Kamerayı listeden yeniden seçin.' : '') };
      var text = 'Kaydedildi. Tarama bu kamerayla açılır' + (parts.length ? ': ' + parts.join(' · ') : '') + '.';
      if (!ok) text = 'Kaydedildi, ama kamera istenen değeri vermiyor' + (parts.length ? ' (şu an ' + parts.join(' · ') + ')' : '') +
        '. Telefon bu kamerada bu çözünürlük/kare hızı birleşimini desteklemiyor; daha düşük bir çözünürlük seçin.';
      return { ok: ok, text: text };
    }
    /** top = first thing in the panel (seen without scrolling), else after the status line at the end */
    function saveBar(top) {
      var bar = el('div', 'omscam-save' + (top ? ' omscam-savetop' : ''));   // not '.top': the field page header owns that class
      var sum = el('div', 'omscam-sum'); bar.appendChild(sum);
      var pr = prefs(getTrack() || undefined), t = getTrack();
      var camTxt = st.sel ? (st.sel.label || 'seçili kamera') : 'Otomatik (arka kamera)';
      var resTxt = pr.res && pr.res !== 'auto' ? (function (x) { return x ? resLabel(x.w, x.h) : pr.res; })(parseRes(pr.res)) : 'Otomatik';
      sum.textContent = 'Taramada: ' + camTxt + ' · ' + resTxt + ' · ' + (pr.fps ? pr.fps + ' kare/s' : 'kare hızı otomatik');
      var b = btn(bar, opts.onDone ? 'Kaydet ve taramaya dön' : 'Kaydet', 'omscam-pri', function () {
        var okSave = save();
        var chk = deliveredCheck(t);
        status(okSave === false ? 'Kaydedilemedi: tarayıcı bu sitede kayıt tutmaya izin vermiyor (gizli sekme?).' : chk.text, okSave === false ? 'err' : chk.ok ? 'ok' : 'warn');
        if (okSave !== false && opts.onDone) { try { opts.onDone(); } catch (e) { } }
      });
      if (!top) b.id = id + '-save';
      if (top) root.insertBefore(bar, root.firstChild); else root.appendChild(bar);
    }

    function restoreFocus(fid) { if (!fid) return; var e = d.getElementById(fid); if (e && e.focus) try { e.focus({ preventScroll: true }); } catch (x) { try { e.focus(); } catch (y) { } } }

    function fillCams(sel, bChange, camNote) {
      sel.innerHTML = '';
      var oa = d.createElement('option'); oa.value = ''; oa.textContent = 'Otomatik (arka kamera)'; sel.appendChild(oa);
      sel.value = '';
      var p = cams ? Promise.resolve(cams) : listCameras();
      p.then(function (list) {
        if (dead) return;
        cams = list;
        var curId = trackDeviceId(getTrack());
        var noLabels = true;
        list.forEach(function (c) {
          if (!c.deviceId) return;
          var o = d.createElement('option'); o.value = c.deviceId;
          var t = c.label + (c.facing === 'environment' && !facingOf(c.rawLabel) ? ' · arka' : c.facing === 'user' && !facingOf(c.rawLabel) ? ' · ön' : '');
          if (c.best) t += ' · önerilen';
          if (c.deviceId === curId) t += ' · şu an açık';
          if (c.rawLabel) noLabels = false;
          o.textContent = t; sel.appendChild(o);
        });
        sel.value = st.sel ? st.sel.id : '';
        if (sel.value !== (st.sel ? st.sel.id : '')) sel.value = '';
        var txt = list.length ? (list.length + ' kamera bulundu.') : 'Kamera listesi alınamadı.';
        if (noLabels && list.length) txt += ' Kamera adları, kameraya izin verildikten sonra görünür.';
        if (list.length > 1) txt += ' Birden çok arka kamera varsa "ultra geniş" ve "tele" kameralar yakından net çekmeyebilir; "önerilen" ana kameradır.';
        camNote.textContent = txt;
      });
    }

    function refreshKv() {
      if (!kvEl) return;
      var t = getTrack(), s = t ? getSet(t) : {};
      var rows = [];
      if (!t) rows.push(['Durum', 'Kamera kapalı']);
      else {
        rows.push(['Kamera', t.label || (s.deviceId ? 'Kimlik ' + String(s.deviceId).slice(0, 8) : '—')]);
        if (lastRep && cur.track === t) { var iss = openIssue(lastRep); rows.push(['Seçilenler', iss ? '⚠ ' + iss : '✓ seçilen kamera ve ayarlar uygulandı']); }
        if (isNum(s.width) && isNum(s.height)) { rows.push(['Çözünürlük', s.width + '×' + s.height]); rows.push(['Megapiksel', fmtMP(s.width, s.height)]); }
        if (isNum(s.frameRate)) rows.push(['Kare hızı', fmtNum(s.frameRate, 1) + ' kare/s']);
        if (s.facingMode) rows.push(['Yön', s.facingMode === 'environment' ? 'Arka' : s.facingMode === 'user' ? 'Ön' : s.facingMode]);
        if (s.focusMode) rows.push(['Odak', (MODE_LABEL[s.focusMode] || s.focusMode) + (isNum(s.focusDistance) ? ' · ' + fmtVal('focusDistance', s.focusDistance) : '')]);
        else if (isNum(s.focusDistance)) rows.push(['Odak mesafesi', fmtVal('focusDistance', s.focusDistance)]);
        if (s.exposureMode) rows.push(['Pozlama', (EXP_LABEL[s.exposureMode] || s.exposureMode) + (isNum(s.exposureTime) ? ' · ' + fmtExposure(s.exposureTime) : '') + (isNum(s.iso) ? ' · ISO ' + Math.round(s.iso) : '')]);
        if (s.whiteBalanceMode) rows.push(['Beyaz denge', (EXP_LABEL[s.whiteBalanceMode] || s.whiteBalanceMode) + (isNum(s.colorTemperature) && s.colorTemperature > 0 ? ' · ' + Math.round(s.colorTemperature) + ' K' : '')]);
        if (isNum(s.zoom)) rows.push(['Yakınlaştırma', fmtVal('zoom', s.zoom)]);
        var tsx = tsOf(t);
        if (torchCap(getCaps(t))) rows.push(['Fener', (has(tsx.acc, 'torch') ? tsx.acc.torch : s.torch) ? 'Açık' : 'Kapalı']);
      }
      var key2 = rows.map(function (x) { return x.join('='); }).join('|');
      if (kvEl.getAttribute('data-k') === key2) return;
      kvEl.setAttribute('data-k', key2);
      kvEl.innerHTML = '';
      rows.forEach(function (x) { kvEl.appendChild(el('dt', null, x[0])); kvEl.appendChild(el('dd', null, x[1])); });
    }

    function tick() {
      if (dead) return;
      if (!root.isConnected && root.isConnected !== undefined) { destroy(); return; }
      var dd = env.doc();
      if (dd && dd.hidden) return;
      if (busy) return;
      var t = getTrack();
      if (t !== lastTrack) { cams = null; build(); return; }
      refreshKv();
    }
    function destroy() { dead = true; if (timer) { clearInterval(timer); timer = null; } if (armTimer) env.clearTimeout(armTimer); }

    build();
    timer = setInterval(tick, 1500);
    return { refresh: function () { cams = null; build(); }, destroy: destroy, root: root };
  }

  // ---------------------------------------------------------------- V12b: full-resolution still (ImageCapture)
  // Why the live picture stops at e.g. 1920x1080 on a 12 / 50 MP phone (researched 2026-10-23, primary sources):
  //  - Chrome on Android lists only the camera driver's YUV video output sizes (VideoCaptureCamera2.getDeviceSupportedFormats,
  //    video_capture_device_factory_android.cc) and Blink caps every candidate at that native size
  //    (media_stream_constraints_util_video_device.cc): getCapabilities().width.max IS the ceiling, a bigger "ideal" cannot exceed it.
  //    Android leaves slow full-sensor sizes out of getOutputSizes() (StreamConfigurationMap: getHighResolutionOutputSizes /
  //    maximum-resolution map), so a 50 MP sensor never shows up as a video size.
  //  - Stills take a separate JPEG path: ImageCapture.getPhotoCapabilities() / takePhoto({imageWidth, imageHeight}) gives the
  //    largest JPEG size (typically the 12 MP class). Chrome closes the preview session for the shot (one short freeze).
  //    Support: Chrome / Samsung Internet 59+, Safari 18.4+ (MDN browser-compat-data api/ImageCapture.json); Firefox: flag only.
  //    One report of takePhoto hanging on Safari → every call has a timeout.
  var photoMemo = (typeof WeakMap === 'function') ? new WeakMap() : null;
  function imageCaptureFor(track) { var w = env.win(); return (w && typeof w.ImageCapture === 'function' && live(track)) ? new w.ImageCapture(track) : null; }
  function canPhoto(track) { var w = env.win(); return !!(w && typeof w.ImageCapture === 'function' && live(track) && typeof w.ImageCapture.prototype.takePhoto === 'function'); }
  function withTimeout(p, ms, what) {
    return new Promise(function (res, rej) {
      var t = env.setTimeout(function () { rej(mkErr('TimeoutError', what + ' zaman aşımı')); }, ms);
      Promise.resolve(p).then(function (v) { env.clearTimeout(t); res(v); }, function (e) { env.clearTimeout(t); rej(e); });
    });
  }
  /** Largest still size of the camera behind `track`: Promise<{w,h}|null> (null = not supported / unknown). */
  function photoCaps(track) {
    if (!canPhoto(track)) return Promise.resolve(null);
    if (photoMemo && photoMemo.has(track)) return photoMemo.get(track);
    var p = Promise.resolve().then(function () {
      var ic = imageCaptureFor(track);
      if (!ic || typeof ic.getPhotoCapabilities !== 'function') return null;
      return withTimeout(ic.getPhotoCapabilities(), 4000, 'getPhotoCapabilities').then(function (pc) {
        var w = pc && pc.imageWidth && pc.imageWidth.max, h = pc && pc.imageHeight && pc.imageHeight.max;
        return isNum(w) && isNum(h) && w > 0 && h > 0 ? { w: w, h: h } : null;
      });
    }).catch(function () { return null; });
    if (photoMemo) photoMemo.set(track, p);
    return p;
  }
  /** Full-resolution JPEG from the running camera: Promise<Blob>. Largest size when known; plain takePhoto() as fallback. */
  function takePhoto(track) {
    if (!canPhoto(track)) return Promise.reject(mkErr('NotSupportedError', 'Bu tarayıcı tam çözünürlüklü fotoğraf çekemiyor'));
    return photoCaps(track).then(function (mx) {
      var ic = imageCaptureFor(track);
      var shot = function (o) { return withTimeout(o ? ic.takePhoto(o) : ic.takePhoto(), 10000, 'takePhoto'); };
      return (mx ? shot({ imageWidth: mx.w, imageHeight: mx.h }).catch(function (e) { if (e && e.name === 'TimeoutError') throw e; return shot(null); }) : shot(null));
    });
  }

  // ---------------------------------------------------------------- public API
  // ---------------------------------------------------------------- V12 live HUD: what the camera is REALLY delivering
  // User report: after switching camera / resolution it was impossible to tell whether anything changed. The HUD reads the live
  // stream, not the saved preference: camera = the running track's label (+ its place among the cameras), resolution = the frame
  // size the <video> actually decodes (videoWidth x videoHeight), fps = frames presented during the last second
  // (requestVideoFrameCallback; "~" + settings.frameRate where that API is missing). It flashes when the camera or size changes.
  var HUD_CSS = '.camhud{font:10px/12px ui-monospace,SFMono-Regular,Menlo,Consolas,"Liberation Mono",monospace;color:#8dffa0;' +
    'background:rgba(0,0,0,.72);border-radius:6px;padding:2px 6px;pointer-events:none;white-space:nowrap;overflow:hidden;' +
    'font-variant-numeric:tabular-nums;transition:color .3s,background-color .3s}' +
    '.camhud[hidden]{display:none}' +
    '.camhud.inrow{flex:0 1 auto;min-width:0;margin-left:8px}' +
    '.camhud.abs{position:absolute;left:8px;top:8px;max-width:calc(100% - 16px);z-index:2}' +
    '.camhud div{overflow:hidden;text-overflow:ellipsis}' +
    '.camhud.chg{color:#000;background:#ffca16}';

  function hudShortName(label) {
    var s = String(label || '');
    var lens = lensOf(s);
    if (lens) return { ultra: 'ultra geniş', tele: 'tele', multi: 'çoklu', wide: 'geniş' }[lens];
    var m = /camera2?\s*(\d+)/i.exec(s);   // Android "camera2 0, facing back"
    if (m) return 'cam' + m[1];
    return s.replace(/\s*\([0-9a-f]{4}:[0-9a-f]{4}\)\s*$/i, '');   // USB webcams: drop the "(046d:0825)" vendor id
  }
  /** Pure formatter (node-tested): {label, facing, idx, n, w, h, fps, approx} -> three short lines. */
  function hudLines(i) {
    i = i || {};
    var face = (i.wrong ? '⚠ ' : '') + (i.facing === 'environment' ? 'ARKA' : i.facing === 'user' ? 'ÖN' : 'KAM');
    var l1 = face + (i.n > 1 && i.idx > 0 ? ' ' + i.idx + '/' + i.n : '');
    var nm = hudShortName(i.label);
    if (nm) l1 += ' · ' + nm;
    var l2 = i.w > 0 && i.h > 0 ? i.w + '×' + i.h + ' ' + fmtMP(i.w, i.h) : '—×— MP';
    var l3 = isNum(i.fps) && i.fps > 0 ? (i.approx ? '~' : '') + (i.fps < 10 ? fmtNum(i.fps, 1) : String(Math.round(i.fps))) + ' fps' : '-- fps';
    if (!i.approx && isNum(i.camFps) && i.camFps > 0 && isNum(i.fps) && Math.abs(i.camFps - i.fps) > Math.max(3, i.camFps * 0.15)) l3 += ' · kamera ' + Math.round(i.camFps);
    return [l1, l2, l3];
  }

  /** Attaches a HUD to a <video>. When the video's stage has a top button row (.stop) the HUD sits in that row, left, after the
   *  buttons (never over a control); otherwise it floats top-left in the video's parent. Returns {el, update, destroy}; idempotent. */
  function hud(videoEl) {
    var doc = env.doc();
    if (!videoEl || !doc || !doc.createElement) return null;
    if (videoEl.__omsHud) return videoEl.__omsHud;
    if (!doc.getElementById('omscamhudcss')) {
      var css = doc.createElement('style'); css.id = 'omscamhudcss'; css.textContent = HUD_CSS;
      (doc.head || doc.documentElement).appendChild(css);
    }
    var host = videoEl.parentElement, row = host && host.querySelector ? host.querySelector('.stop') : null;
    var el = doc.createElement('div');
    el.className = 'camhud ' + (row ? 'inrow' : 'abs'); el.setAttribute('aria-hidden', 'true'); el.hidden = true;
    var lines = [doc.createElement('div'), doc.createElement('div'), doc.createElement('div')];
    lines.forEach(function (d) { el.appendChild(d); });
    if (row) row.insertBefore(el, row.querySelector('.sp') || null); else if (host) host.appendChild(el);
    var w = env.win();
    var S = { timer: null, rvfc: false, f0: 0, t0: 0, n: 0, fps: null, key: '', devKey: null, list: null, flashT: null, dead: false };

    function track() { var so = videoEl.srcObject; try { return so && so.getVideoTracks ? so.getVideoTracks()[0] || null : null; } catch (e) { return null; } }
    function onFrame(now, meta) {
      if (S.dead || !videoEl.srcObject) { S.rvfc = false; return; }
      var n = meta && isNum(meta.presentedFrames) ? meta.presentedFrames : S.n + 1;
      if (!S.t0) { S.t0 = now; S.f0 = n; }
      else if (now - S.t0 >= 1000) { S.fps = (n - S.f0) * 1000 / (now - S.t0); S.t0 = now; S.f0 = n; }
      S.n = n;
      videoEl.requestVideoFrameCallback(onFrame);
    }
    function cams(devId) {
      if (S.devKey === devId) return;
      S.devKey = devId; S.list = lastList;
      var m = md();
      if (m && m.enumerateDevices) Promise.resolve().then(function () { return m.enumerateDevices(); }).then(function (ds) {
        S.list = (ds || []).filter(function (d) { return d && d.kind === 'videoinput'; }); paint();
      }, noop);
    }
    function paint() {
      var t = track();
      if (S.dead || !t || t.readyState === 'ended') { el.hidden = true; return; }
      var s = {}; try { s = t.getSettings ? t.getSettings() : {}; } catch (e) { }
      var devId = s.deviceId || '';
      cams(devId);
      var idx = 0, n = S.list ? S.list.length : 0;
      for (var i = 0; i < n; i++) if (S.list[i].deviceId === devId) idx = i + 1;
      var vw = videoEl.videoWidth || s.width || 0, vh = videoEl.videoHeight || s.height || 0;
      var fps = S.rvfc ? S.fps : (isNum(s.frameRate) ? s.frameRate : null);
      var txt = hudLines({ label: t.label, facing: s.facingMode || facingOf(t.label), idx: idx, n: n, w: vw, h: vh, fps: fps, approx: !S.rvfc, camFps: s.frameRate, wrong: !!(lastRep && lastRep.got && lastRep.got.id === devId && (lastRep.fallback || !lastRep.deviceOk)) });
      for (var j = 0; j < 3; j++) if (lines[j].textContent !== txt[j]) lines[j].textContent = txt[j];
      el.title = (t.label || '') + ' · ' + txt[1] + ' · ' + txt[2];
      el.hidden = false;
      var key = devId + '|' + vw + 'x' + vh;
      if (S.key && vw && key !== S.key) {   // camera or resolution really changed -> visible flash
        el.classList.add('chg'); env.clearTimeout(S.flashT);
        S.flashT = env.setTimeout(function () { el.classList.remove('chg'); }, 1800);
      }
      if (vw) S.key = key;
    }
    function start() {
      if (S.dead) return;
      if (!S.rvfc && videoEl.requestVideoFrameCallback) { S.rvfc = true; S.t0 = 0; S.fps = null; videoEl.requestVideoFrameCallback(onFrame); }
      if (!S.timer && w && w.setInterval) S.timer = w.setInterval(function () { if (!videoEl.srcObject) stop(); else paint(); }, 1000);
      paint();
    }
    function stop() {
      if (S.timer && w && w.clearInterval) w.clearInterval(S.timer);
      S.timer = null; S.fps = null; S.t0 = 0; el.hidden = true;   // S.key kept: reopening on another camera still flashes
    }
    var evs = { loadedmetadata: start, playing: start, resize: function () { if (videoEl.srcObject) start(); }, emptied: stop };
    for (var k in evs) if (has(evs, k)) videoEl.addEventListener(k, evs[k]);
    if (videoEl.srcObject) start();
    var ctl = {
      el: el, update: paint,
      destroy: function () {
        S.dead = true; stop();
        for (var k2 in evs) if (has(evs, k2)) videoEl.removeEventListener(k2, evs[k2]);
        if (el.parentNode) el.parentNode.removeChild(el);
        videoEl.__omsHud = null;
      }
    };
    videoEl.__omsHud = ctl;
    return ctl;
  }
  /** Both pages: the scan stage (#cvid) and the LAN page's camera-settings preview (#kvid). Safe to call more than once. */
  function hudAuto() {
    var doc = env.doc(); if (!doc || !doc.getElementById) return;
    ['cvid', 'kvid'].forEach(function (id) { var v = doc.getElementById(id); if (v) hud(v); });
  }
  (function () {
    var doc = env.doc();
    if (!doc || !doc.addEventListener || !doc.getElementById) return;   // node tests: no DOM
    if (doc.readyState === 'loading') doc.addEventListener('DOMContentLoaded', hudAuto); else hudAuto();
  })();

  var API = {
    version: '6.0',
    KEY: KEY,
    load: load,
    prefs: prefs,
    setPref: setPref,
    select: select,
    selected: function () { ensure(); return st.sel ? st.sel.id : ''; },
    supported: supported,
    listCameras: listCameras,
    pickBack: pickBack,
    facingOf: facingOf,
    lensOf: lensOf,
    constraints: constraints,
    open: open,
    close: close,
    current: function () { return { stream: cur.stream, track: cur.track }; },
    apply: apply,
    renderPanel: renderPanel,
    focusAt: focusAt,
    tapFocusSupported: tapFocusSupported,
    bindTapFocus: bindTapFocus,
    tapPoint: tapPoint,
    torch: torch,
    torchSupported: function (track) { return live(track) && torchCap(getCaps(track)); },
    zoom: zoom,
    resetDefaults: resetDefaults,
    errorText: errorText,
    resPresets: resPresets,
    fpsPresets: fpsPresets,
    fpsMax: fpsMax,
    lastOpen: function () { return lastRep; },
    openIssue: openIssue,
    resLabel: resLabel,
    fmtMP: fmtMP,
    hud: hud,
    canPhoto: canPhoto,
    fpsGuard: fpsGuard,
    photoCaps: photoCaps,
    takePhoto: takePhoto,
    hudLines: hudLines,
    // test hooks
    _env: env,
    _reset: function () { st = fresh(); loaded = false; autoPick = null; lastList = null; cur = { stream: null, track: null, el: null }; if (tstate) tstate = new WeakMap(); },
    _state: function () { return st; },
    _groupsFor: groupsFor
  };
  return API;
});
