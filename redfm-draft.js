/* RED FM — local draft autosave for report forms.
   Saves a report's typed data to this device's browser (localStorage) so a
   dropped connection, a browser reset, or navigating away does NOT lose the
   day's work. Drafts are per-device only (never synced) and are cleared once
   the report is submitted. File attachments (photos, lab PDFs) are NOT saved.

   Usage (call once, from the form's own inline script after state/draw exist):
     REDFMDraft.init({
       key:   'weekly',          // unique per form (required)
       state: state,             // the form's answer object (optional)
       draw:  draw,              // re-render fn for grid forms (optional)
       finishId: 'finish',       // submit button id (default 'finish')
       extra: {                  // for values not held in id'd inputs (optional)
         save:    () => ({ severity }),
         restore: (o) => { severity = o.severity || ''; ...repaint... }
       }
     });
   On a successful submit the form must call REDFMDraft.clear().
*/
window.REDFMDraft = (function () {
  function init(cfg) {
    cfg = cfg || {};
    var KEY = 'redfm_draft_v1_' + (cfg.key || location.pathname);
    var finishId = cfg.finishId || 'finish';
    var ok = storageOK();
    var userTouched = false, pendingRestore = false, done = false, saveTimer = null;

    // ---- snapshot / restore -------------------------------------------------
    function inputEls() {
      return Array.prototype.slice.call(
        document.querySelectorAll('input[id],textarea[id],select[id]')
      ).filter(function (el) {
        var t = (el.type || '').toLowerCase();
        return t !== 'file' && t !== 'button' && t !== 'submit' && t !== 'reset';
      });
    }
    function snapshot() {
      var snap = { ts: Date.now(), inputs: {} };
      inputEls().forEach(function (el) {
        var t = (el.type || '').toLowerCase();
        snap.inputs[el.id] = (t === 'checkbox' || t === 'radio') ? { c: el.checked } : el.value;
      });
      if (cfg.state) { try { snap.state = JSON.parse(JSON.stringify(cfg.state)); } catch (e) {} }
      if (cfg.extra && cfg.extra.save) { try { snap.extra = cfg.extra.save(); } catch (e) {} }
      return snap;
    }
    function hasContent(s) {
      if (!s) return false;
      if (s.state && Object.keys(s.state).length) return true;
      if (s.extra) { for (var k in s.extra) { if (s.extra[k]) return true; } }
      var inp = s.inputs || {};
      for (var id in inp) {
        var v = inp[id];
        if (v && typeof v === 'object') { if (v.c) return true; }
        else if (v != null && String(v).trim() !== '') {
          // ignore auto-set date fields that equal today (not real user input)
          if (!/^\d{4}-\d{2}-\d{2}$/.test(v) || v !== new Date().toISOString().slice(0, 10)) return true;
        }
      }
      return false;
    }
    function applyRestore(s) {
      if (!s) return;
      if (cfg.state) { Object.keys(cfg.state).forEach(function (k) { delete cfg.state[k]; });
        if (s.state) Object.assign(cfg.state, s.state); }
      var inp = s.inputs || {};
      Object.keys(inp).forEach(function (id) {
        var el = document.getElementById(id); if (!el) return;
        var v = inp[id];
        if (v && typeof v === 'object' && 'c' in v) el.checked = v.c; else el.value = v;
      });
      if (cfg.extra && cfg.extra.restore) { try { cfg.extra.restore(s.extra || {}); } catch (e) {} }
      if (cfg.draw) { try { cfg.draw(); } catch (e) {} }
    }

    // ---- persistence --------------------------------------------------------
    function save(manual) {
      if (!ok || done) return;
      try {
        var s = snapshot();
        localStorage.setItem(KEY, JSON.stringify(s));
        setStatus('Saved ' + hhmm(s.ts), manual);
      } catch (e) { setStatus('Could not save (storage full?)', true); }
    }
    function clear() {
      done = true;
      try { localStorage.removeItem(KEY); } catch (e) {}
      var pill = document.getElementById('rfDraftPill'); if (pill) pill.style.display = 'none';
      var ban = document.getElementById('rfDraftBanner'); if (ban) ban.remove();
    }
    function read() {
      if (!ok) return null;
      try { return JSON.parse(localStorage.getItem(KEY) || 'null'); } catch (e) { return null; }
    }

    // ---- UI: floating Save pill + restore banner ---------------------------
    function pill() {
      var d = document.createElement('div');
      d.id = 'rfDraftPill';
      d.style.cssText = 'position:fixed;right:14px;bottom:78px;z-index:9998;display:flex;'
        + 'flex-direction:column;align-items:flex-end;gap:4px;font-family:inherit;';
      var b = document.createElement('button');
      b.type = 'button'; b.textContent = '💾 Save draft';
      b.style.cssText = 'border:0;border-radius:22px;padding:11px 16px;background:#222;color:#fff;'
        + 'font-weight:700;font-size:14px;box-shadow:0 2px 8px rgba(0,0,0,.25);cursor:pointer;';
      b.onclick = function () { save(true); };
      var s = document.createElement('div');
      s.id = 'rfDraftStatus';
      s.style.cssText = 'font-size:11px;color:#444;background:rgba(255,255,255,.9);'
        + 'padding:2px 8px;border-radius:10px;display:none;';
      d.appendChild(b); d.appendChild(s); document.body.appendChild(d);
    }
    function setStatus(msg, flash) {
      var s = document.getElementById('rfDraftStatus'); if (!s) return;
      s.textContent = msg; s.style.display = 'block';
      if (flash) { s.style.background = '#e7f5ec'; setTimeout(function () { s.style.background = 'rgba(255,255,255,.9)'; }, 1200); }
    }
    function banner(s) {
      pendingRestore = true;
      var d = document.createElement('div');
      d.id = 'rfDraftBanner';
      d.style.cssText = 'position:sticky;top:0;z-index:9999;background:#fff7e0;border-bottom:1px solid #e8d28a;'
        + 'padding:10px 14px;display:flex;gap:10px;align-items:center;flex-wrap:wrap;'
        + 'font-family:inherit;font-size:14px;color:#5a4a1a;';
      var txt = document.createElement('span');
      txt.style.flex = '1';
      txt.innerHTML = 'You have an unsaved draft from <b>' + when(s.ts) + '</b> on this device.';
      var r = mkbtn('Restore draft', '#1e8e3e');
      var x = mkbtn('Discard', '#b00');
      r.onclick = function () { applyRestore(s); pendingRestore = false; d.remove(); setStatus('Draft restored', true); };
      x.onclick = function () { try { localStorage.removeItem(KEY); } catch (e) {} pendingRestore = false; d.remove(); };
      d.appendChild(txt); d.appendChild(r); d.appendChild(x);
      document.body.insertBefore(d, document.body.firstChild);
    }
    function mkbtn(label, color) {
      var b = document.createElement('button');
      b.type = 'button'; b.textContent = label;
      b.style.cssText = 'border:0;border-radius:8px;padding:8px 14px;color:#fff;font-weight:700;'
        + 'font-size:13px;cursor:pointer;background:' + color + ';';
      return b;
    }

    // ---- helpers ------------------------------------------------------------
    function hhmm(t) { var d = new Date(t); return pad(d.getHours()) + ':' + pad(d.getMinutes()); }
    function when(t) { var d = new Date(t), n = new Date();
      var tm = pad(d.getHours()) + ':' + pad(d.getMinutes());
      return d.toDateString() === n.toDateString() ? ('today ' + tm)
        : (d.getDate() + '/' + (d.getMonth() + 1) + ' ' + tm); }
    function pad(n) { return (n < 10 ? '0' : '') + n; }
    function storageOK() { try { var k = '__rf'; localStorage.setItem(k, '1'); localStorage.removeItem(k); return true; } catch (e) { return false; } }

    // ---- wire up ------------------------------------------------------------
    if (!ok) return { save: function () {}, clear: function () {}, restore: function () {} };

    pill();
    var existing = read();
    if (hasContent(existing)) banner(existing);

    function markTouched() { userTouched = true; }
    ['input', 'change'].forEach(function (ev) {
      document.addEventListener(ev, function (e) {
        var el = e.target;
        if (el && (el.matches('input,textarea,select') || el.closest('.chk,.sevbtn,[data-id]'))) markTouched();
        schedule();
      }, true);
    });
    document.addEventListener('click', function (e) {
      if (e.target && e.target.closest('.chk,.sevbtn,[data-id],.tab,button')) { markTouched(); schedule(); }
    }, true);

    function schedule() {
      if (done || pendingRestore || !userTouched) return;
      clearTimeout(saveTimer);
      saveTimer = setTimeout(function () { save(false); }, 600);
    }

    // safety net: save when leaving/hiding the page
    window.addEventListener('pagehide', function () { if (userTouched && !pendingRestore && !done) save(false); });
    document.addEventListener('visibilitychange', function () {
      if (document.visibilityState === 'hidden' && userTouched && !pendingRestore && !done) save(false);
    });

    return { save: function () { save(true); }, clear: clear, restore: function () { var s = read(); if (s) applyRestore(s); } };
  }
  return { init: init };
})();
