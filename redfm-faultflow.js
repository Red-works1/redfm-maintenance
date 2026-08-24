/* RED FM — fault capture flow (phase 1, items 1 + 2)
   Shared by every capture form. After a report is submitted, every line item the engineer
   flagged as a fault must be documented here before the visit counts as complete.

   Usage from a form, after the visit header + readings have saved:

     REDFMFaultFlow.start({
       visitId | visitIds:[...], completeStatus, visitRef, visitDate, engineer, plant,
       items: [{assetId, section, item, col}],
       buildReport: function(faults){ ... return {blob, filename}; }   // optional
     }).then(function(res){ ...res.faults, res.withdrawn... });

   Requires redfm-data.js (REDFM) to be loaded first. Degrades to a local-only run if
   SharePoint is unreachable — the queue is kept in localStorage and can be resumed.
*/
window.REDFMFaultFlow = (function () {
  "use strict";

  var QKEY = "redfm_faultqueue_v1";
  var STYLE_ID = "redfm-faultflow-style";

  /* ---------- canonical asset identity -------------------------------------
     Faults must carry the same asset code the reading tables already use, so
     "click E2, see everything" and repeat counters can join the two together. */
  function assetId(variant, section, col) {
    var v = String(variant || "").trim();
    var c = String(col || "").trim();
    if (!c) return v;
    if (/^(kwhr|°c)$/i.test(c)) return (v + " — " + (section || "")).trim();  // units, not assets
    if (/^(no\d\s+)?[a-e]\d/i.test(c)) return c;      // E2, C1, A4, No1 B2, E1 Compr1
    if (/^[a-e]$/i.test(c)) return c.toUpperCase();    // the A / B packs on the A&B sheet
    return (v + " " + c).trim();
  }

  /* ---------- Plant is a fixed CHOICE column on FaultRegister ---------------
     Confirmed live 24 Aug 2026: A / B / C1-C4 / D1-D4 / E1 / E2 / A&B Plant Room /
     General. AssetId carries the precise identity, so Plant just has to resolve to a
     legal choice — never write a raw group name like "C & D" into it. */
  var PLANT_CHOICES = ["A", "B", "C1", "C2", "C3", "C4", "D1", "D2", "D3", "D4",
                       "E1", "E2", "A&B Plant Room", "General"];
  function plantChoice(asset, group) {
    var a = String(asset || "").trim();
    var A = a.toUpperCase();
    if (/a\s*&\s*b/i.test(a)) return "A&B Plant Room";   // plant-room level, not a chamber
    // longest choice the asset code starts with: "E1 Compr1" -> "E1", "C1" -> "C1"
    var hit = PLANT_CHOICES.filter(function (c) { return A.indexOf(c.toUpperCase()) === 0; })
                           .sort(function (x, y) { return y.length - x.length; })[0];
    if (hit) return hit;
    var m = /^(?:no\d\s+)?([a-e])/i.exec(a);   // "A1" -> A, "No1 B2" -> B
    if (m && PLANT_CHOICES.indexOf(m[1].toUpperCase()) >= 0) return m[1].toUpperCase();
    if (/a\s*&\s*b/i.test(group || "")) return "A&B Plant Room";
    return "General";
  }

  /* ---------- storage ------------------------------------------------------ */
  function load() {
    try { return JSON.parse(localStorage.getItem(QKEY) || "null"); } catch (e) { return null; }
  }
  function save(q) {
    try { localStorage.setItem(QKEY, JSON.stringify(q)); } catch (e) {}
  }
  function clear() {
    try { localStorage.removeItem(QKEY); } catch (e) {}
  }
  function pending() {
    var q = load();
    if (!q || !q.items) return null;
    var outstanding = q.items.filter(function (i) { return !i.done; }).length +
      (q.closures || []).filter(function (i) { return !i.done; }).length;
    return outstanding ? { ref: q.visitRef, date: q.visitDate, outstanding: outstanding,
      total: q.items.length + (q.closures || []).length } : null;
  }

  /* ---------- styling ------------------------------------------------------ */
  function injectStyle() {
    if (document.getElementById(STYLE_ID)) return;
    var s = document.createElement("style");
    s.id = STYLE_ID;
    s.textContent = [
      ".ff-wrap{position:fixed;inset:0;z-index:9999;background:rgba(21,21,26,.72);display:flex;align-items:flex-start;justify-content:center;overflow-y:auto;padding:16px;-webkit-overflow-scrolling:touch;}",
      ".ff-card{background:#fff;border-radius:14px;max-width:640px;width:100%;margin:auto;box-shadow:0 18px 50px rgba(0,0,0,.35);overflow:hidden;font-family:inherit;}",
      ".ff-head{background:#15151a;color:#fff;padding:16px 18px;}",
      ".ff-head h3{margin:0;font-size:17px;font-weight:800;letter-spacing:.2px;}",
      ".ff-head p{margin:5px 0 0;font-size:12.5px;color:#b9b9c2;line-height:1.45;}",
      ".ff-bar{height:4px;background:#2a2a33;}",
      ".ff-bar i{display:block;height:100%;background:#E01322;transition:width .25s ease;}",
      ".ff-body{padding:18px;}",
      ".ff-ctx{background:#f5f5f7;border:1px solid #e3e3e8;border-left:4px solid #E01322;border-radius:10px;padding:12px 14px;margin-bottom:16px;}",
      ".ff-asset{font-size:21px;font-weight:800;color:#15151a;line-height:1.15;}",
      ".ff-meta{font-size:12.5px;color:#6b6b72;margin-top:4px;line-height:1.5;}",
      ".ff-body label{display:block;font-size:13px;font-weight:700;color:#15151a;margin:14px 0 6px;}",
      ".ff-body label span{color:#E01322;}",
      ".ff-body textarea,.ff-body input[type=text]{width:100%;box-sizing:border-box;font:inherit;font-size:15px;padding:11px 12px;border:1.5px solid #e3e3e8;border-radius:10px;background:#fff;color:#15151a;}",
      ".ff-body textarea{min-height:76px;resize:vertical;}",
      ".ff-body textarea:focus,.ff-body input:focus{outline:none;border-color:#E01322;}",
      ".ff-sev{display:flex;gap:8px;}",
      ".ff-sev button{flex:1;font:inherit;font-size:15px;font-weight:700;padding:12px 0;border-radius:10px;border:1.5px solid #e3e3e8;background:#fff;color:#6b6b72;cursor:pointer;}",
      '.ff-sev button.on[data-v="Low"]{background:#eef7f0;color:#1a7f37;border-color:#1a7f37;}',
      '.ff-sev button.on[data-v="Medium"]{background:#fdf3e3;color:#b9770b;border-color:#b9770b;}',
      '.ff-sev button.on[data-v="High"]{background:#fdecee;color:#E01322;border-color:#E01322;}',
      ".ff-ph{display:flex;gap:8px;flex-wrap:wrap;margin-top:6px;}",
      ".ff-ph .thumb{width:74px;height:74px;border-radius:8px;object-fit:cover;border:1px solid #e3e3e8;}",
      ".ff-add{width:74px;height:74px;border-radius:8px;border:1.5px dashed #e3e3e8;background:#fff;color:#E01322;font-size:24px;cursor:pointer;display:flex;align-items:center;justify-content:center;}",
      ".ff-actions{display:flex;gap:10px;align-items:center;margin-top:20px;padding-top:16px;border-top:1px solid #e3e3e8;}",
      ".ff-btn{font:inherit;font-size:15px;font-weight:700;padding:13px 20px;border-radius:10px;border:0;cursor:pointer;}",
      ".ff-btn.primary{background:#E01322;color:#fff;flex:1;}",
      ".ff-btn.primary:disabled{opacity:.5;cursor:default;}",
      ".ff-link{background:none;border:0;font:inherit;font-size:12.5px;color:#6b6b72;text-decoration:underline;cursor:pointer;padding:6px 0;}",
      ".ff-hint{font-size:12.5px;color:#b9770b;margin-top:10px;line-height:1.5;}",
      ".ff-hint.ok{color:#1a7f37;}",
      ".ff-warn{background:#fdecee;border:1px solid #f6c3c8;color:#8c1018;border-radius:8px;padding:10px 12px;font-size:12.5px;margin-top:12px;line-height:1.5;}",
      ".ff-done{text-align:center;padding:34px 24px;}",
      ".ff-done .tick{font-size:38px;color:#1a7f37;}",
      ".ff-done h3{margin:10px 0 6px;font-size:19px;color:#15151a;}",
      ".ff-done p{margin:0;font-size:13.5px;color:#6b6b72;line-height:1.55;}",
      ".ff-banner{background:#fdf3e3;border:1px solid #f0d9a8;border-left:4px solid #b9770b;border-radius:10px;padding:12px 14px;margin:12px 0;font-size:13.5px;color:#7a4e05;display:flex;gap:12px;align-items:center;flex-wrap:wrap;}",
      ".ff-banner button{font:inherit;font-size:13.5px;font-weight:700;padding:9px 15px;border-radius:8px;border:0;background:#b9770b;color:#fff;cursor:pointer;}"
    ].join("\n");
    document.head.appendChild(s);
  }

  /* ---------- helpers ------------------------------------------------------ */
  function el(tag, cls, html) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (html != null) e.innerHTML = html;
    return e;
  }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }
  function slug(s) {
    return String(s || "").toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "").slice(0, 80);
  }
  function downscale(dataURL, maxDim) {
    return new Promise(function (res) {
      var img = new Image();
      img.onload = function () {
        var w = img.width, h = img.height, s = Math.min(1, maxDim / Math.max(w, h));
        var c = document.createElement("canvas");
        c.width = Math.round(w * s); c.height = Math.round(h * s);
        c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
        res(c.toDataURL("image/jpeg", 0.85));
      };
      img.onerror = function () { res(dataURL); };
      img.src = dataURL;
    });
  }

  /* ---------- SharePoint writes -------------------------------------------
     The extra columns (AssetId, SourceVisitRef, …) are added to SharePoint as part
     of this phase. Until they exist a POST carrying them returns 400, so the first
     failure falls back to the legacy field set for the rest of the session and the
     asset/visit link is carried in the description instead. Never lose a fault
     because a column is missing. */
  var extraFieldsOk = true;

  function faultFields(rec, withExtras) {
    var base = {
      Title: rec.ref, FaultDate: rec.date, Plant: rec.plant, Severity: rec.severity,
      Description: rec.description, ActionTaken: rec.action, Stage: "Reported", RaisedBy: rec.raisedBy
    };
    if (!withExtras) {
      base.Description = "[" + rec.assetId + " · " + rec.line + " · from " + rec.visitRef + "] " + rec.description;
      return base;
    }
    base.AssetId = rec.assetId;
    base.SourceVisitRef = rec.visitRef;
    base.SourceType = "Engineer report";
    base.TriageStatus = "Not required";
    base.RepeatCount = 1;
    base.LastFlagged = rec.date;
    base.SourceLine = rec.line || "";
    base.LastReviewed = rec.date;
    return base;
  }

  /* ---------- Item 6 — repeat flags increment, they do not duplicate --------
     Ray's "E2 E2 E2 E2 down the page". A re-flag of the SAME asset on the SAME
     check, while an earlier fault for it is still open, bumps that fault's
     RepeatCount instead of adding a row. Both AssetId and SourceLine must match
     and both must be non-empty — matching on asset alone would wrongly merge two
     genuinely different faults on the same chamber. */
  var RESOLVED_STAGES = ["Closed", "Reported and fixed"];
  function isOpenFault(f) {
    var st = (f && f.Stage) || "";
    return st !== "Call out" && RESOLVED_STAGES.indexOf(st) < 0;
  }
  async function findOpenMatch(rec) {
    if (!rec.assetId || !rec.line || !REDFM.getFaultsWithId) return null;
    var all;
    try { all = await REDFM.getFaultsWithId(); } catch (e) { return null; }
    var a = String(rec.assetId).trim().toLowerCase();
    var l = String(rec.line).trim().toLowerCase();
    var hits = (all || []).filter(function (f) {
      return isOpenFault(f) &&
        String(f.AssetId || "").trim().toLowerCase() === a &&
        String(f.SourceLine || "").trim().toLowerCase() === l;
    });
    // most recently flagged wins if somehow there is more than one
    hits.sort(function (x, y) {
      return String(y.LastFlagged || y.FaultDate || "").localeCompare(String(x.LastFlagged || x.FaultDate || ""));
    });
    return hits[0] || null;
  }

  // Returns {ref, repeated, count} — ref is the EXISTING fault's ref on a repeat,
  // so the engineer and the report both quote the original ID, not a new one.
  async function writeFault(rec) {
    if (extraFieldsOk) {
      var match = null;
      try { match = await findOpenMatch(rec); } catch (e) { match = null; }
      if (match && (match._id || match.id)) {
        var count = (parseInt(match.RepeatCount, 10) || 1) + 1;
        var note = (match.ActionTaken || "") +
          (match.ActionTaken ? "\n" : "") +
          "[" + rec.date + " — flagged again (x" + count + ") by " + rec.raisedBy + "] " + rec.action;
        try {
          await REDFM.updateFault(match._id || match.id, {
            RepeatCount: count, LastFlagged: rec.date, LastReviewed: rec.date,
            ActionTaken: note,
            // three strikes escalates the severity as well as the colour
            Severity: count >= 3 ? "High" : (match.Severity || rec.severity)
          });
          return { ref: match.Title || match.FaultRef || rec.ref, repeated: true, count: count };
        } catch (e) { /* fall through and create a new record rather than lose it */ }
      }
      try {
        await REDFM.addItem("FaultRegister", faultFields(rec, true));
        return { ref: rec.ref, repeated: false, count: 1 };
      } catch (e) {
        extraFieldsOk = false; // columns not there yet — degrade for the rest of the session
      }
    }
    await REDFM.addItem("FaultRegister", faultFields(rec, false));
    return { ref: rec.ref, repeated: false, count: 1 };
  }

  var REVIEW_AFTER_DAYS = 7;
  function daysBetween(fromISO, toISO) {
    if (!fromISO || !toISO) return null;
    var a = new Date(String(fromISO).slice(0, 10)), b = new Date(String(toISO).slice(0, 10));
    if (isNaN(a) || isNaN(b)) return null;
    return Math.max(0, Math.round((b - a) / 86400000));
  }

  /* Open faults that have not been looked at for a week. Anything already dealt
     with on THIS visit (raised or re-flagged just now) is excluded — the engineer
     has only just answered for it. */
  async function loadClosures(q) {
    if (!REDFM.getFaultsWithId) return [];
    var all;
    try { all = await REDFM.getFaultsWithId(); } catch (e) { return []; }
    var today = q.visitDate;
    return (all || []).filter(function (f) {
      if (!isOpenFault(f)) return false;
      if (String(f.SourceVisitRef || "") === String(q.visitRef || "")) return false;
      var reviewed = f.LastReviewed ? String(f.LastReviewed).slice(0, 10) : null;
      if (reviewed === String(today).slice(0, 10)) return false;
      var ago = daysBetween(reviewed, today);
      return reviewed == null || ago >= REVIEW_AFTER_DAYS;
    }).map(function (f) {
      var raised = f.FaultDate ? String(f.FaultDate).slice(0, 10) : null;
      var reviewed = f.LastReviewed ? String(f.LastReviewed).slice(0, 10) : null;
      return {
        id: f._id || f.id, ref: f.Title || f.FaultRef || "", assetId: f.AssetId || "",
        plant: f.Plant || "", line: f.SourceLine || "", description: f.Description || "",
        severity: f.Severity || "", stage: f.Stage || "", faultDate: raised,
        closureNote: f.ClosureNote || "",
        repeatCount: parseInt(f.RepeatCount, 10) || 1,
        ageDays: daysBetween(raised, today) || 0,
        reviewedAgo: reviewed == null ? null : daysBetween(reviewed, today),
        done: false, outcome: null
      };
    }).sort(function (a, b) { return b.ageDays - a.ageDays; });   // oldest first
  }

  async function patchVisit(visitIds, fields) {
    if (!REDFM.updateVisit) return;
    var ids = (Array.isArray(visitIds) ? visitIds : [visitIds]).filter(Boolean);
    for (var i = 0; i < ids.length; i++) {
      try { await REDFM.updateVisit(ids[i], fields); }
      catch (e) {
        // Retry without the new counter columns
        var lean = {};
        if (fields.Status) lean.Status = fields.Status;
        if (Object.keys(lean).length) { try { await REDFM.updateVisit(ids[i], lean); } catch (e2) {} }
      }
    }
  }

  async function nextRefBase() {
    try {
      var existing = await REDFM.getFaults();
      var max = 1000;
      existing.forEach(function (it) {
        var m = /-(\d{3,})$/.exec(it.Title || it.FaultRef || "");
        if (m) { var n = parseInt(m[1], 10); if (!isNaN(n) && n > max) max = n; }
      });
      return max;
    } catch (e) {
      return null; // fall back to timestamp refs
    }
  }

  /* ---------- the flow ----------------------------------------------------- */
  function start(opts) {
    injectStyle();
    var q = load();
    // Resume the stored queue only if it is the same visit; otherwise start fresh.
    if (!q || q.visitRef !== opts.visitRef) {
      q = {
        visitId: opts.visitId || null,
        visitIds: opts.visitIds || (opts.visitId ? [opts.visitId] : []),
        // A form that submits in parts (quarterly, certs) must NOT be forced to
        // "Fully completed" just because its faults are now documented.
        completeStatus: opts.completeStatus || "Fully completed",
        visitRef: opts.visitRef, visitDate: opts.visitDate,
        engineer: opts.engineer, plant: opts.plant || "", formUrl: location.pathname,
        items: (opts.items || []).map(function (i) {
          return {
            assetId: i.assetId, group: i.group || "", section: i.section || "", item: i.item || "", col: i.col || "",
            done: false, ref: null, withdrawn: false
          };
        }),
        faults: [], withdrawn: []
      };
      save(q);
    } else {
      if (opts.visitId && !q.visitId) q.visitId = opts.visitId;
      if (opts.visitIds && opts.visitIds.length) q.visitIds = opts.visitIds;
      if (opts.completeStatus) q.completeStatus = opts.completeStatus;
      save(q);
    }
    return run(q, opts);
  }

  function resume(opts) {
    injectStyle();
    var q = load();
    if (!q) return Promise.resolve(null);
    return run(q, opts || {});
  }

  function run(q, opts) {
    return new Promise(function (resolve) {
      var refBase = null;
      var wrap = el("div", "ff-wrap");
      var card = el("div", "ff-card");
      wrap.appendChild(card);
      document.body.appendChild(wrap);
      document.body.style.overflow = "hidden";

      nextRefBase().then(function (b) { refBase = b; });

      function finish() {
        document.body.style.overflow = "";
        wrap.remove();
      }

      function nextIndex() {
        for (var i = 0; i < q.items.length; i++) if (!q.items[i].done) return i;
        return -1;
      }

      async function complete() {
        var faults = q.faults || [];
        card.innerHTML = "";
        var d = el("div", "ff-done");
        d.innerHTML = '<div class="tick">&#10003;</div><h3>All faults documented</h3>' +
          '<p>Finishing the report&hellip;</p>';
        card.appendChild(d);

        var msg = "";
        // Build + file the report now that the faults block has something to show
        if (typeof opts.buildReport === "function") {
          try {
            var out = await opts.buildReport(faults, q.withdrawn || [], q.reviewed || []);
            if (out && out.blob) {
              await REDFM.uploadReportPdf(out.filename, out.blob);
              msg = "Report PDF filed.";
            }
          } catch (e) { msg = "Report PDF could not be filed (" + (e.message || e) + ") — the data is saved."; }
        }
        await patchVisit(q.visitIds && q.visitIds.length ? q.visitIds : q.visitId, {
          Status: q.completeStatus || "Fully completed",
          FaultsFlagged: q.items.length,
          FaultsDocumented: faults.length
        });
        clear();
        var rev = q.reviewed || [];
        var fixedN = rev.filter(function (r) { return r.outcome === "fixed"; }).length;
        d.innerHTML = '<div class="tick">&#10003;</div><h3>Visit complete</h3>' +
          "<p>" + faults.length + " fault" + (faults.length === 1 ? "" : "s") + " logged to the Fault Register" +
          (q.withdrawn.length ? ", " + q.withdrawn.length + " flag withdrawn" : "") +
          (rev.length ? ", " + rev.length + " open fault" + (rev.length === 1 ? "" : "s") + " reviewed" +
            (fixedN ? " (" + fixedN + " closed as fixed)" : "") : "") + ". " + esc(msg) + "</p>" +
          '<div style="margin-top:18px"><button class="ff-btn primary" style="flex:0 0 auto">Done</button></div>';
        d.querySelector("button").onclick = function () {
          finish();
          resolve({ faults: faults, withdrawn: q.withdrawn, reviewed: q.reviewed || [] });
        };
      }

      function nextClosure() {
        var c = q.closures || [];
        for (var i = 0; i < c.length; i++) if (!c[i].done) return i;
        return -1;
      }

      var loadingClosures = false;
      function render() {
        var idx = nextIndex();
        if (idx >= 0) return renderFault(q.items[idx]);
        // Fault cards are done — now bring back anything still open and unreviewed.
        if (!q.closures) {
          if (loadingClosures) return;
          loadingClosures = true;
          card.innerHTML = '<div class="ff-done"><h3>Checking open faults&hellip;</h3>' +
            "<p>Looking for anything still open that has not been reviewed for a week.</p></div>";
          loadClosures(q).then(function (list) {
            q.closures = list; save(q); loadingClosures = false; render();
          }).catch(function () { q.closures = []; save(q); loadingClosures = false; render(); });
          return;
        }
        var ci = nextClosure();
        if (ci >= 0) return renderClosure(q.closures[ci]);
        complete();
      }

      /* ---------- Items 4 + 5 — reverse sign-off ------------------------------
         Every open fault that has not been looked at for a week comes back at the
         point of completion. The engineer either closes it with evidence or says
         why it is still open. This is the control that would have made the
         20 August report impossible. */
      function renderClosure(f) {
        var doneCount = (q.closures || []).filter(function (x) { return x.done; }).length;
        var total = (q.closures || []).length;
        var photos = [];
        var outcome = "";

        card.innerHTML = "";
        var head = el("div", "ff-head");
        head.innerHTML = "<h3>Open fault review &mdash; " + (doneCount + 1) + " of " + total + "</h3>" +
          "<p>This fault has been open " + f.ageDays + " day" + (f.ageDays === 1 ? "" : "s") +
          " and was last looked at " + (f.reviewedAgo == null ? "never" : f.reviewedAgo + " days ago") +
          ". Close it with evidence, or say why it is still open.</p>";
        var bar = el("div", "ff-bar");
        bar.innerHTML = '<i style="width:' + Math.round((doneCount / Math.max(1, total)) * 100) + '%"></i>';
        card.appendChild(head); card.appendChild(bar);

        var body = el("div", "ff-body");
        var ctx = el("div", "ff-ctx");
        ctx.innerHTML = '<div class="ff-asset">' + esc(f.assetId || f.plant || "Plant") +
          (f.repeatCount > 1 ? ' <span style="font-size:13px;font-weight:700;color:#E01322">&times;' + f.repeatCount + ' flagged</span>' : "") +
          "</div>" +
          '<div class="ff-meta">' + esc(f.ref) + (f.line ? " &middot; " + esc(f.line) : "") +
          "<br>" + esc(f.description || "") +
          "<br>Raised " + esc(f.faultDate || "") + " &middot; " + esc(f.severity || "") + " &middot; " + esc(f.stage || "") + "</div>";
        body.appendChild(ctx);

        body.appendChild(el("label", null, "What is the position now? <span>*</span>"));
        var pick = el("div", "ff-sev");
        // data-k carries the logical value; data-v only drives the existing colour rules
        // (Low = green for fixed, High = red for still open). Keep them separate — reusing
        // one attribute for both is how the highlight broke on the second click.
        [["fixed", "Fixed", "Low"], ["open", "Still open", "High"]].forEach(function (o) {
          var b = el("button", null, o[1]);
          b.type = "button";
          b.dataset.k = o[0];
          b.setAttribute("data-v", o[2]);
          b.onclick = function () {
            outcome = o[0];
            Array.prototype.forEach.call(pick.children, function (x) {
              x.className = x.dataset.k === outcome ? "on" : "";
            });
            fixedBox.style.display = outcome === "fixed" ? "" : "none";
            openBox.style.display = outcome === "open" ? "" : "none";
            validate();
          };
          pick.appendChild(b);
        });
        body.appendChild(pick);

        // --- fixed: date + note + photo ---
        var fixedBox = el("div");
        fixedBox.style.display = "none";
        fixedBox.appendChild(el("label", null, "Date fixed <span>*</span>"));
        var fdate = el("input"); fdate.type = "date"; fdate.value = q.visitDate;
        fixedBox.appendChild(fdate);
        fixedBox.appendChild(el("label", null, "What was done <span>*</span>"));
        var fnote = el("textarea");
        fnote.placeholder = "Parts fitted, adjustment made, retested and within range…";
        fixedBox.appendChild(fnote);
        fixedBox.appendChild(el("label", null, "Evidence photo <span>*</span>"));
        var phWrap = el("div", "ff-ph");
        var add = el("button", "ff-add", "+"); add.type = "button";
        var file = el("input");
        file.type = "file"; file.accept = "image/*"; file.multiple = true;
        file.setAttribute("capture", "environment"); file.style.display = "none";
        add.onclick = function () { file.click(); };
        file.onchange = function () {
          Array.prototype.forEach.call(file.files, function (fl) {
            var r = new FileReader();
            r.onload = function () {
              photos.push({ name: fl.name, data: r.result, file: fl });
              var img = el("img", "thumb"); img.src = r.result;
              phWrap.insertBefore(img, add); validate();
            };
            r.readAsDataURL(fl);
          });
          file.value = "";
        };
        phWrap.appendChild(add);
        fixedBox.appendChild(phWrap); fixedBox.appendChild(file);
        body.appendChild(fixedBox);

        // --- still open: why ---
        var openBox = el("div");
        openBox.style.display = "none";
        openBox.appendChild(el("label", null, "Why is it still open? <span>*</span>"));
        var onote = el("textarea");
        onote.placeholder = "Awaiting parts, awaiting Border instruction, quoted and not yet approved, access needed…";
        openBox.appendChild(onote);
        body.appendChild(openBox);

        var warn = el("div", "ff-warn"); warn.style.display = "none";
        body.appendChild(warn);

        var actions = el("div", "ff-actions");
        var go = el("button", "ff-btn primary", "Save"); go.type = "button"; go.disabled = true;
        actions.appendChild(go);
        body.appendChild(actions);
        var hint = el("div", "ff-hint");
        body.appendChild(hint);
        card.appendChild(body);

        function validate() {
          var missing = [];
          if (!outcome) missing.push("fixed or still open");
          else if (outcome === "fixed") {
            if (!fdate.value) missing.push("the date it was fixed");
            if (fnote.value.trim().length < 4) missing.push("what was done");
            if (!photos.length) missing.push("an evidence photo");
          } else {
            if (onote.value.trim().length < 10) missing.push("a reason it is still open");
          }
          go.disabled = missing.length > 0;
          if (!missing.length) { hint.className = "ff-hint ok"; hint.textContent = "Ready to save."; }
          else {
            hint.className = "ff-hint";
            hint.textContent = "Still needed: " + (missing.length === 1 ? missing[0]
              : missing.slice(0, -1).join(", ") + " and " + missing[missing.length - 1]) + ".";
          }
        }
        fnote.oninput = validate; onote.oninput = validate; fdate.onchange = validate;
        validate();

        go.onclick = async function () {
          go.disabled = true; go.textContent = "Saving…";
          var stamp = "[" + q.visitDate + " — " + q.engineer + "] ";
          var fields;
          if (outcome === "fixed") {
            fields = {
              Stage: "Reported and fixed",
              ClosedDate: fdate.value, ClosedBy: q.engineer,
              ClosureNote: (f.closureNote ? f.closureNote + "\n" : "") + stamp + "FIXED: " + fnote.value.trim(),
              LastReviewed: q.visitDate
            };
          } else {
            fields = {
              ClosureNote: (f.closureNote ? f.closureNote + "\n" : "") + stamp + "STILL OPEN: " + onote.value.trim(),
              LastReviewed: q.visitDate
            };
          }
          try {
            await REDFM.updateFault(f.id, fields);
          } catch (e) {
            warn.style.display = "";
            warn.textContent = "Could not save (" + (e.message || e) + "). Press again when you have signal.";
            go.disabled = false; go.textContent = "Save";
            return;
          }
          // evidence photo → Faults library, named against the fault ref
          if (outcome === "fixed" && photos.length) {
            try {
              var base = "Fixed_" + fdate.value + "_" + slug(f.assetId || f.plant || "plant") + "_" + f.ref;
              for (var i = 0; i < photos.length; i++) {
                var p = photos[i];
                var ext = (p.name.split(".").pop() || "jpg").toLowerCase();
                await REDFM.uploadFile("Faults", base + "_evidence" + (i + 1) + "." + ext, p.file, p.file.type || "image/jpeg");
              }
            } catch (e) { /* best effort */ }
          }
          f.done = true;
          f.outcome = outcome;
          f.note = (outcome === "fixed" ? fnote.value.trim() : onote.value.trim());
          f.fixedDate = outcome === "fixed" ? fdate.value : "";
          (q.reviewed = q.reviewed || []).push({
            ref: f.ref, assetId: f.assetId, line: f.line, outcome: outcome,
            note: f.note, fixedDate: f.fixedDate, ageDays: f.ageDays
          });
          save(q);
          render();
        };
      }

      function renderFault(it) {
        var doneCount = q.items.filter(function (x) { return x.done; }).length;
        var photos = [];
        var severity = "";

        card.innerHTML = "";
        var head = el("div", "ff-head");
        head.innerHTML = "<h3>Fault detail required &mdash; " + (doneCount + 1) + " of " + q.items.length + "</h3>" +
          "<p>The report is saved but sits at <b>Awaiting fault detail</b> until every flagged item is documented. " +
          "No PDF is filed and nothing goes to Border until then.</p>";
        var bar = el("div", "ff-bar");
        bar.innerHTML = '<i style="width:' + Math.round((doneCount / q.items.length) * 100) + '%"></i>';
        card.appendChild(head); card.appendChild(bar);

        var body = el("div", "ff-body");
        var ctx = el("div", "ff-ctx");
        ctx.innerHTML = '<div class="ff-asset">' + esc(it.assetId || q.plant || "Plant") + "</div>" +
          '<div class="ff-meta">' + esc(it.section) + (it.item ? " &middot; " + esc(it.item) : "") +
          "<br>Visit " + esc(q.visitRef) + " &middot; " + esc(q.visitDate) + " &middot; " + esc(q.engineer) + "</div>";
        body.appendChild(ctx);

        body.appendChild(el("label", null, 'What is wrong? <span>*</span>'));
        var desc = el("textarea");
        desc.placeholder = "Describe the fault as you found it";
        body.appendChild(desc);

        body.appendChild(el("label", null, "Severity <span>*</span>"));
        var sev = el("div", "ff-sev");
        ["Low", "Medium", "High"].forEach(function (v) {
          var b = el("button", null, v);
          b.type = "button"; b.dataset.v = v;
          b.onclick = function () {
            severity = v;
            Array.prototype.forEach.call(sev.children, function (x) {
              x.className = x.dataset.v === severity ? "on" : "";
            });
            validate();
          };
          sev.appendChild(b);
        });
        body.appendChild(sev);

        body.appendChild(el("label", null, "Immediate action taken <span>*</span>"));
        var act = el("textarea");
        act.placeholder = "What you did on the day — isolated, adjusted, made safe, parts required, left running…";
        body.appendChild(act);

        body.appendChild(el("label", null, "Photo <span>*</span>"));
        var phWrap = el("div", "ff-ph");
        var add = el("button", "ff-add", "+");
        add.type = "button";
        var file = el("input");
        file.type = "file"; file.accept = "image/*"; file.multiple = true;
        file.setAttribute("capture", "environment");
        file.style.display = "none";
        add.onclick = function () { file.click(); };
        file.onchange = function () {
          Array.prototype.forEach.call(file.files, function (f) {
            var r = new FileReader();
            r.onload = function () {
              photos.push({ name: f.name, data: r.result, file: f });
              var img = el("img", "thumb");
              img.src = r.result;
              phWrap.insertBefore(img, add);
              validate();
            };
            r.readAsDataURL(f);
          });
          file.value = "";
        };
        phWrap.appendChild(add);
        body.appendChild(phWrap);
        body.appendChild(file);

        var warn = el("div", "ff-warn");
        warn.style.display = "none";
        body.appendChild(warn);

        var actions = el("div", "ff-actions");
        var go = el("button", "ff-btn primary", "Log this fault");
        go.type = "button"; go.disabled = true;
        var wd = el("button", "ff-link", "Ticked in error");
        wd.type = "button";
        actions.appendChild(go); actions.appendChild(wd);
        body.appendChild(actions);
        var hint = el("div", "ff-hint");
        body.appendChild(hint);
        card.appendChild(body);

        // A disabled button with no explanation is useless on a phone in a cold store.
        // Say exactly what is still outstanding, and keep the bar low enough that a real
        // one-word answer ("Iced", "Leak") passes while "ok" / "na" does not.
        function validate() {
          var missing = [];
          if (desc.value.trim().length < 4) missing.push("what is wrong");
          if (!severity) missing.push("a severity");
          if (act.value.trim().length < 3) missing.push("the action you took");
          if (!photos.length) missing.push("a photo");
          go.disabled = missing.length > 0;
          if (!missing.length) {
            hint.className = "ff-hint ok";
            hint.textContent = "Ready to log.";
          } else {
            hint.className = "ff-hint";
            hint.textContent = "Still needed: " + (missing.length === 1 ? missing[0]
              : missing.slice(0, -1).join(", ") + " and " + missing[missing.length - 1]) + ".";
          }
        }
        desc.oninput = validate; act.oninput = validate;
        validate();

        /* --- withdraw a mis-tapped flag: allowed, but it has to be explained --- */
        wd.onclick = function () {
          warn.style.display = "";
          warn.innerHTML = "Withdrawing a flag is recorded against the visit and shown on the dashboard. " +
            "Type why this was ticked in error (at least 10 characters):" +
            '<input type="text" id="ffwhy" style="margin-top:8px" placeholder="e.g. mis-tap, reading was within range">' +
            '<div style="margin-top:10px;display:flex;gap:8px"><button class="ff-btn primary" id="ffwok" style="flex:0 0 auto;padding:9px 16px;font-size:13.5px" disabled>Withdraw flag</button>' +
            '<button class="ff-btn" id="ffwno" style="flex:0 0 auto;padding:9px 16px;font-size:13.5px;background:#fff;color:#6b6b72;border:1.5px solid #e3e3e8">Cancel</button></div>';
          var why = warn.querySelector("#ffwhy");
          var ok = warn.querySelector("#ffwok");
          why.oninput = function () { ok.disabled = why.value.trim().length < 10; };
          warn.querySelector("#ffwno").onclick = function () { warn.style.display = "none"; };
          ok.onclick = function () {
            it.done = true; it.withdrawn = true;
            q.withdrawn.push({ assetId: it.assetId, line: it.section + " · " + it.item, reason: why.value.trim() });
            save(q);
            render();
          };
        };

        /* --- log the fault --- */
        go.onclick = async function () {
          go.disabled = true; go.textContent = "Saving…";
          var line = (it.section ? it.section + " · " : "") + it.item;
          var ref = refBase != null
            ? "FLT-" + (refBase + 1 + q.faults.length)
            : "FLT-" + Date.now().toString(36).toUpperCase();
          var rec = {
            ref: ref, date: q.visitDate, plant: plantChoice(it.assetId, it.group || q.plant),
            assetId: it.assetId, line: line, visitRef: q.visitRef,
            severity: severity, description: desc.value.trim(), action: act.value.trim(),
            raisedBy: q.engineer
          };

          var saved = false, res = null;
          try {
            res = await writeFault(rec);
            if (res && res.repeated) { rec.ref = res.ref; rec.repeatCount = res.count; }
            saved = true;
          } catch (e) {
            warn.style.display = "";
            warn.textContent = "Could not save to the Fault Register (" + (e.message || e) +
              "). It is kept on this device — reconnect and press again.";
            go.disabled = false; go.textContent = "Log this fault";
            return;
          }

          // Fault PDF + original photos → Faults library (never block the save on this)
          try {
            var pdfPhotos = await Promise.all(photos.map(function (p) { return downscale(p.data, 1400); }));
            var blob = REDFM.buildFaultPdf({
              ref: rec.ref, date: rec.date, plant: rec.assetId + " — " + rec.plant, severity: rec.severity,
              raisedBy: rec.raisedBy, stage: rec.repeatCount ? "Reported (repeat x" + rec.repeatCount + ")" : "Reported",
              description: rec.description, action: rec.action,
              subtitle: "Border Holdings, Avonmouth · " + (rec.repeatCount
                ? "RE-FLAGGED on visit " + rec.visitRef + " — occurrence " + rec.repeatCount
                : "raised on visit " + rec.visitRef)
            }, pdfPhotos);
            var base = "Fault_" + rec.date + "_" + slug(rec.assetId) + "_" + rec.ref;
            await REDFM.uploadFile("Faults", base + ".pdf", blob, "application/pdf");
            var n = 0;
            for (var i = 0; i < photos.length; i++) {
              n++;
              var p = photos[i];
              var ext = (p.name.split(".").pop() || "jpg").toLowerCase();
              await REDFM.uploadFile("Faults", base + "_photo" + n + "." + ext, p.file, p.file.type || "image/jpeg");
            }
          } catch (e) { /* PDF/photo filing is best-effort */ }

          if (saved) {
            it.done = true; it.ref = rec.ref;
            q.faults.push(rec);
            save(q);
            await patchVisit(q.visitIds && q.visitIds.length ? q.visitIds : q.visitId, {
              FaultsFlagged: q.items.length,
              FaultsDocumented: q.faults.length
            });
            render();
          }
        };
      }

      render();
    });
  }

  /* ---------- resume banner for an abandoned queue ------------------------- */
  function checkPending(container) {
    var p = pending();
    if (!p) return null;
    injectStyle();
    var b = el("div", "ff-banner");
    b.innerHTML = "<div><b>" + p.outstanding + " fault" + (p.outstanding === 1 ? "" : "s") +
      "</b> still to document from visit " + esc(p.ref) + " (" + esc(p.date) + "). " +
      "That report is not complete and has not been filed.</div>";
    var btn = el("button", null, "Finish it now");
    btn.type = "button";
    btn.onclick = function () { b.remove(); resume(); };
    b.appendChild(btn);
    var host = container || document.querySelector(".wrap") || document.querySelector("main") || document.body;
    host.insertBefore(b, host.firstChild);
    return b;
  }

  return {
    assetId: assetId,
    start: start,
    resume: resume,
    pending: pending,
    checkPending: checkPending,
    clear: clear
  };
})();
