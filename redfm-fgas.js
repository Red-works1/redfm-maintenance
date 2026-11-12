/* RED FM — F-Gas refrigerant register: shared logic.
   Loaded by RED_FM_FGas_Register.html, RED_FM_Compliance_Dashboard.html and view.html so the
   three never drift apart (view.html is otherwise a hand-kept copy of the dashboard).

   Three SharePoint lists, all readable by the shared Coolstream login — so NO money in any of
   them, ever:
     RefrigerantSystems    one row per refrigeration system (Title = plant code), installed charge
     RefrigerantCylinders  one row per cylinder on site, with a running ContentsKg
     RefrigerantMovements  every weighing: opening stock, deliveries, additions, recoveries, checks

   The running ContentsKg is a convenience. The ledger is the movements: the balance is
   re-derived from them and any cylinder whose stored figure disagrees by more than 0.5 kg is
   flagged rather than silently trusted. */
(function () {
  "use strict";

  // GWP (100-year), AR4 — the values in Annex I of the GB F-Gas Regulation (retained 517/2014),
  // which is what the leak-check thresholds are written against.
  var GWP = { R407C: 1774, R407F: 1825, R449A: 1397, R448A: 1387, R404A: 3922, R134a: 1430 };

  var PLANTS = ["A", "B", "C1", "C2", "C3", "C4", "D1", "D2", "D3", "D4", "E1", "E2"];
  var MOVE_TYPES = ["Opening stock", "Delivered to site", "Added to system", "Recovered from system",
    "Removed from site", "Stock check"];
  var REASONS = ["Leak repair", "Top-up - no leak found", "Initial charge / commissioning",
    "Component replacement", "Stock record", "Other"];
  var TOLERANCE_KG = 0.5;

  function n(v) { if (v === null || v === undefined || v === "") return null; var x = Number(v); return isNaN(x) ? null : x; }
  function d10(v) { return v ? String(v).slice(0, 10) : ""; }
  function gasOf(v) { var s = String(v || "").trim(); return GWP[s] ? s : null; }

  /* Leak-check band from tonnes CO2e. GB F-Gas Regulation Art. 4: 5–<50 t every 12 months,
     50–<500 t every 6 months, 500 t+ every 3 months; a working leak-detection system doubles
     the interval (and is itself mandatory at 500 t+). Below 5 t no periodic check is required. */
  function band(t, hasLD) {
    if (t === null) return { key: "unknown", label: "Charge not yet recorded", months: null };
    var v = String(hasLD || "");
    var ld = v === "Yes", noLd = v === "No";
    if (t < 5) return { key: "none", label: "Below 5 t CO2e - no periodic check", months: null };
    var base = t >= 500 ? 3 : t >= 50 ? 6 : 12;
    var words = { 3: "Quarterly", 6: "Six-monthly", 12: "Annual", 24: "Every 24 months" };
    /* Leak detection not yet confirmed: show both intervals and do NOT raise the red
       "required" flag — an unconfirmed fact must not read as a breach on Border's view. */
    if (!ld && !noLd) return { key: "m" + base, label: words[base] + " (" + words[base * 2].toLowerCase() + " if leak detection is fitted - to confirm)", months: base, ldRequired: false };
    var m = ld ? base * 2 : base;
    var label = words[m] + (ld ? " (leak detection fitted)" : "");
    if (t >= 500 && noLd) label += " - leak detection REQUIRED at 500 t+";
    return { key: "m" + m, label: label, months: m, ldRequired: t >= 500 && noLd };
  }

  function systemRow(s) {
    var gas = gasOf(s.GasType), kg = n(s.ChargeKg);
    var t = (gas && kg !== null) ? Math.round(kg * GWP[gas] / 100) / 10 : null;   // t CO2e, 1 dp
    return { plant: s.Title || "", gas: s.GasType || "Not yet confirmed", gasKnown: !!gas, kg: kg,
      t: t, ld: s.HasLeakDetection || "Not yet confirmed", band: band(t, s.HasLeakDetection),
      source: s.ChargeSource || "", notes: s.Notes || "", _id: s._id };
  }

  function sortMoves(moves) {
    return moves.slice().sort(function (a, b) {
      var c = d10(a.MoveDate).localeCompare(d10(b.MoveDate));
      if (c) return c;
      return String(a.Title || "").localeCompare(String(b.Title || ""), undefined, { numeric: true });
    });
  }

  /* Signed effect of one movement on its cylinder's NET contents.
     StartKg / EndKg are the scale readings (gross, cylinder + gas) the engineer writes down —
     they are evidence, not the balance. QtyKg is always net refrigerant: the amount moved, or
     for Opening stock / Stock check the net contents counted, which resets the balance. */
  function effect(m) {
    var t = String(m.MoveType || ""), q = n(m.QtyKg);
    if (t === "Opening stock" || t === "Stock check") return { set: q };
    if (q === null) return { delta: 0 };
    if (t === "Delivered to site" || t === "Recovered from system") return { delta: +q };
    if (t === "Added to system" || t === "Removed from site") return { delta: -q };
    return { delta: 0 };
  }

  /* Re-derive each cylinder's balance from its movements, in date order. */
  function cylinderBalances(cylinders, moves) {
    var byRef = {};
    sortMoves(moves).forEach(function (m) {
      var ref = String(m.CylinderRef || "").trim(); if (!ref) return;
      var b = byRef[ref] || (byRef[ref] = { bal: null, count: 0, last: "" });
      var e = effect(m);
      if (e.set !== undefined && e.set !== null) b.bal = e.set;
      else b.bal = (b.bal || 0) + e.delta;
      b.bal = Math.round(b.bal * 100) / 100;
      b.count++; b.last = d10(m.MoveDate);
    });
    return cylinders.map(function (c) {
      var ref = String(c.Title || "").trim(), b = byRef[ref] || { bal: null, count: 0, last: "" };
      var stored = n(c.ContentsKg);
      var mismatch = (b.bal !== null && stored !== null && Math.abs(b.bal - stored) > TOLERANCE_KG) ||
                     (b.bal === null && stored !== null && stored > 0);
      return { ref: ref, gas: c.GasType || "", verified: String(c.GasVerified || "") === "Yes",
        stored: stored, derived: b.bal, moves: b.count, last: b.last || d10(c.LastMovement),
        nominal: n(c.NominalKg), serial: c.Serial || "", tare: n(c.TareKg), owner: c.Owner || "", type: c.CylinderType || "",
        location: c.LocationNote || "", status: c.Status || "On site", source: c.SourceNote || "",
        mismatch: mismatch, _id: c._id };
    });
  }

  function onSite(c) { return c.status === "On site" || c.status === "Empty on site" || !c.status; }

  /* Additions to systems in the last N days, split by whether they were leak-caused. */
  function additions(moves, days, today) {
    var now = today ? new Date(today + "T12:00:00") : new Date();
    var from = new Date(now.getTime() - days * 864e5).toISOString().slice(0, 10);
    var rows = moves.filter(function (m) { return m.MoveType === "Added to system" && d10(m.MoveDate) >= from; });
    var out = { total: 0, leak: 0, nonCommission: 0, tCO2e: 0, rows: rows };
    rows.forEach(function (m) {
      var q = n(m.QtyKg) || 0, gas = gasOf(m.GasType);
      out.total += q;
      if (m.Reason === "Leak repair") out.leak += q;
      if (m.Reason !== "Initial charge / commissioning") out.nonCommission += q;
      if (gas) out.tCO2e += q * GWP[gas] / 1000;
    });
    ["total", "leak", "nonCommission"].forEach(function (k) { out[k] = Math.round(out[k] * 10) / 10; });
    out.tCO2e = Math.round(out.tCO2e * 10) / 10;
    return out;
  }

  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function kg(v) { return v === null || v === undefined ? "—" : (Math.round(v * 10) / 10) + " kg"; }
  function dmy(v) { var s = d10(v); return s ? s.slice(8, 10) + "/" + s.slice(5, 7) + "/" + s.slice(0, 4) : "—"; }

  /* The dashboard / client-view card. data = {systems, cylinders, movements} — any of them null
     means the list could not be read, and the card says so instead of showing a false all-clear. */
  function renderBox(el, data, opts) {
    opts = opts || {};
    if (!el) return;
    if (!data || data.systems == null || data.cylinders == null || data.movements == null) {
      el.innerHTML = '<div class="empty">The refrigerant register is not available in this view yet.</div>';
      return;
    }
    var sys = data.systems.map(systemRow).sort(function (a, b) { return PLANTS.indexOf(a.plant) - PLANTS.indexOf(b.plant); });
    var cyl = cylinderBalances(data.cylinders, data.movements).filter(onSite);
    var add12 = additions(data.movements, 365, opts.today);
    var stockByGas = {};
    cyl.forEach(function (c) { var v = c.derived !== null ? c.derived : c.stored; if (v === null) return;
      var g = c.gas || "Unknown"; stockByGas[g] = (stockByGas[g] || 0) + v; });
    var unverified = cyl.filter(function (c) { return !c.verified; }).length;
    var mism = cyl.filter(function (c) { return c.mismatch; }).length;
    var missingCharge = sys.filter(function (s) { return s.kg === null || !s.gasKnown; }).length;

    var stockTiles = Object.keys(stockByGas).sort().map(function (g) {
      return '<div class="stage"><div class="c">' + (Math.round(stockByGas[g] * 10) / 10) + '</div><div class="s">kg ' + esc(g) + ' in stock</div></div>';
    }).join("") || '<div class="stage"><div class="c">0</div><div class="s">kg in stock</div></div>';
    var h = '<div class="cohead">Stock on site <span>· ' + cyl.length + ' cylinder' + (cyl.length === 1 ? "" : "s") + '</span></div>' +
      '<div class="pipe">' + stockTiles +
      '<div class="stage ' + (unverified ? "us" : "fixed") + '"><div class="c">' + unverified + '</div><div class="s">cylinders not label-checked</div></div>' +
      (mism ? '<div class="stage them"><div class="c">' + mism + '</div><div class="s">balances to reconcile</div></div>' : "") +
      '</div>';
    h += '<div class="cohead">Added to systems <span>· last 12 months</span></div><div class="pipe">' +
      '<div class="stage"><div class="c">' + add12.total + '</div><div class="s">kg added</div></div>' +
      '<div class="stage ' + (add12.leak ? "them" : "fixed") + '"><div class="c">' + add12.leak + '</div><div class="s">kg after a leak</div></div>' +
      '<div class="stage co"><div class="c">' + add12.tCO2e + '</div><div class="s">t CO2e added</div></div></div>';

    h += '<div class="cohead">Systems and leak-check frequency' + (missingCharge ? ' <span>· ' + missingCharge + ' awaiting charge figures</span>' : "") + '</div>' +
      '<div class="tblscroll"><table><thead><tr><th>System</th><th>Gas</th><th>Charge</th><th>t CO2e</th><th>Leak detection</th><th>Leak checks</th></tr></thead><tbody>' +
      sys.map(function (s) {
        return '<tr><td data-label="System"><b>' + esc(s.plant) + '</b></td><td data-label="Gas">' + esc(s.gas) + '</td>' +
          '<td data-label="Charge">' + kg(s.kg) + '</td><td data-label="t CO2e">' + (s.t === null ? "—" : s.t) + '</td>' +
          '<td data-label="Leak detection">' + esc(s.ld) + '</td>' +
          '<td data-label="Leak checks"' + (s.band.ldRequired ? ' style="color:var(--red);font-weight:700"' : "") + '>' + esc(s.band.label) + '</td></tr>';
      }).join("") + '</tbody></table></div>';

    var leaks = add12.rows.filter(function (m) { return m.Reason === "Leak repair" || m.Reason === "Top-up - no leak found"; })
      .sort(function (a, b) { return d10(b.MoveDate).localeCompare(d10(a.MoveDate)); });
    h += '<div class="cohead">Refrigerant added after a leak or loss <span>· last 12 months</span></div>' +
      (leaks.length ? '<div class="tblscroll"><table><thead><tr><th>Date</th><th>System</th><th>Added</th><th>Gas</th><th>Reason</th><th>Leak / repair</th><th>Ref</th></tr></thead><tbody>' +
        leaks.map(function (m) {
          return '<tr><td data-label="Date">' + dmy(m.MoveDate) + '</td><td data-label="System"><b>' + esc(m.Plant || "—") + '</b></td>' +
            '<td data-label="Added">' + kg(n(m.QtyKg)) + '</td><td data-label="Gas">' + esc(m.GasType || "") + '</td>' +
            '<td data-label="Reason">' + esc(m.Reason || "") + '</td><td class="wide" data-label="Leak / repair">' + esc(m.LeakDetail || "—") + '</td>' +
            '<td data-label="Ref">' + esc([m.Title, m.LeakRef, m.JobRef].filter(Boolean).join(" · ")) + '</td></tr>';
        }).join("") + '</tbody></table></div>'
        : '<div class="empty ok">No refrigerant added after a leak in the last 12 months.</div>');
    if (opts.note) h += '<div class="srcnote">' + esc(opts.note) + '</div>';
    el.innerHTML = h;
  }

  window.REDFMFGas = { GWP: GWP, PLANTS: PLANTS, MOVE_TYPES: MOVE_TYPES, REASONS: REASONS, TOLERANCE_KG: TOLERANCE_KG,
    band: band, systemRow: systemRow, cylinderBalances: cylinderBalances, additions: additions, effect: effect,
    sortMoves: sortMoves, onSite: onSite, renderBox: renderBox, n: n, d10: d10 };
})();
