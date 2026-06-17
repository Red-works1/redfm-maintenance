/* RED FM Maintenance — shared data layer
   Signs in the shared maintenance account (MSAL) and reads/writes the
   SharePoint lists via Microsoft Graph. Used by every capture form + the dashboard.

   SETUP: after IT does the app registration, paste the Application (client) ID below.
   Load MSAL before this file:
   <script src="https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js"></script>
   <script src="redfm-data.js"></script>
*/
window.REDFM = (function () {
  const CONFIG = {
    clientId: "9de89aff-4367-4355-808d-f9fccbb1f974",     // RED FM Maintenance Web App registration
    tenantId: "63dfe382-90e1-467d-926e-af043415e67f",     // Red Refrigeration Bristol
    siteHostname: "redbristol.sharepoint.com",
    sitePath: "/sites/REDFMMaintenance",
    scopes: ["Sites.ReadWrite.All"]
  };

  const msalConfig = {
    auth: {
      clientId: CONFIG.clientId,
      authority: "https://login.microsoftonline.com/" + CONFIG.tenantId,
      redirectUri: window.location.origin   // must match a registered SPA redirect URI (site root)
    },
    cache: { cacheLocation: "localStorage" }   // keeps the engineer signed in between visits
  };

  let msalApp, account, siteId = null;
  const listIds = {};

  async function init() {
    msalApp = new window.msal.PublicClientApplication(msalConfig);
    if (msalApp.initialize) await msalApp.initialize();
    const redirect = await msalApp.handleRedirectPromise();
    account = redirect ? redirect.account : (msalApp.getAllAccounts()[0] || null);
    return !!account;
  }

  async function signIn() {
    const r = await msalApp.loginPopup({ scopes: CONFIG.scopes });
    account = r.account;
    return account;
  }

  function signOut() { return msalApp.logoutPopup({ account }); }
  function getAccount() { return account; }

  async function getToken() {
    try {
      const r = await msalApp.acquireTokenSilent({ scopes: CONFIG.scopes, account });
      return r.accessToken;
    } catch (e) {
      const r = await msalApp.acquireTokenPopup({ scopes: CONFIG.scopes });
      account = r.account; return r.accessToken;
    }
  }

  async function graph(path, method, body) {
    const token = await getToken();
    const res = await fetch("https://graph.microsoft.com/v1.0" + path, {
      method: method || "GET",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: body ? JSON.stringify(body) : undefined
    });
    if (!res.ok) throw new Error("Graph " + res.status + ": " + (await res.text()));
    return res.status === 204 ? null : res.json();
  }

  async function getSiteId() {
    if (siteId) return siteId;
    const s = await graph("/sites/" + CONFIG.siteHostname + ":" + CONFIG.sitePath);
    return (siteId = s.id);
  }

  async function getListId(title) {
    if (listIds[title]) return listIds[title];
    const sid = await getSiteId();
    const r = await graph("/sites/" + sid + "/lists?$select=id,displayName&$top=100");
    r.value.forEach(l => (listIds[l.displayName] = l.id));
    return listIds[title];
  }

  async function listItems(title, filterFn) {
    const sid = await getSiteId();
    const lid = await getListId(title);
    let url = "/sites/" + sid + "/lists/" + lid + "/items?$expand=fields&$top=500";
    const out = [];
    while (url) {
      const page = await graph(url);
      page.value.forEach(i => { if (!filterFn || filterFn(i.fields)) out.push(i.fields); });
      url = page["@odata.nextLink"] ? page["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "") : null;
    }
    return out;
  }

  // The reading definitions for a service type, in order (e.g. "Weekly", "Quarterly A&B")
  async function getCatalogue(serviceType) {
    const rows = await listItems("ReadingCatalogue", f => f.ServiceType === serviceType);
    return rows.sort((a, b) => (a.SortOrder || 0) - (b.SortOrder || 0));
  }

  async function addItem(title, fields) {
    const sid = await getSiteId();
    const lid = await getListId(title);
    return graph("/sites/" + sid + "/lists/" + lid + "/items", "POST", { fields });
  }

  // Create many items fast using Graph $batch (20 per request) — a weekly sheet is 100+ rows.
  async function addItemsBatch(title, fieldsArray) {
    const sid = await getSiteId();
    const lid = await getListId(title);
    const url = "/sites/" + sid + "/lists/" + lid + "/items";
    const results = [];
    for (let i = 0; i < fieldsArray.length; i += 20) {
      const requests = fieldsArray.slice(i, i + 20).map((f, n) => ({
        id: String(n), method: "POST", url: url,
        headers: { "Content-Type": "application/json" }, body: { fields: f }
      }));
      const res = await graph("/$batch", "POST", { requests });
      results.push(...res.responses);
    }
    const failed = results.filter(r => r.status >= 300);
    if (failed.length) throw new Error(failed.length + " of " + fieldsArray.length + " readings failed to save");
    return results;
  }

  // Save a visit header + all its reading line-items. Returns the created visit.
  async function saveVisit(visitFields, readingRows) {
    const visit = await addItem("ServiceVisits", visitFields);
    const ref = visitFields.Title || ("SV-" + visit.id);
    const rows = readingRows.map(rd => Object.assign({ VisitRef: ref }, rd));
    if (rows.length) await addItemsBatch("Readings", rows);
    return visit;
  }

  // Read everything the dashboard needs
  async function getServiceVisits() { return listItems("ServiceVisits"); }
  async function getFaults() { return listItems("FaultRegister"); }

  // ---- PDF report (client-side) ----
  let reportsDriveId = null;
  async function getReportsDriveId() {
    if (reportsDriveId) return reportsDriveId;
    const sid = await getSiteId();
    const d = await graph("/sites/" + sid + "/drives?$select=id,name");
    const drive = d.value.find(x => x.name === "Reports") || d.value[0];
    return (reportsDriveId = drive.id);
  }

  // Upload a Blob to the Reports document library. Returns the created file (incl. webUrl).
  async function uploadReportPdf(filename, blob) {
    const token = await getToken();
    const sid = await getSiteId();
    const did = await getReportsDriveId();
    const safe = filename.replace(/[\\/:*?"<>|#%]+/g, "-");
    const res = await fetch("https://graph.microsoft.com/v1.0/sites/" + sid + "/drives/" + did +
      "/root:/" + encodeURIComponent(safe) + ":/content", {
      method: "PUT",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/pdf" },
      body: blob
    });
    if (!res.ok) throw new Error("PDF upload " + res.status + ": " + (await res.text()));
    return res.json();
  }

  const RED = [0xE0, 0x13, 0x22];
  // Coolstream (UK) Ltd logo — extracted from their letterhead; shown on F-Gas certificates only.
  const COOLSTREAM_LOGO = "data:image/jpeg;base64,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";

  // Draw the RED FM logo lockup as vector (red rules + RED / FM). Returns its bottom Y.
  function drawLogo(doc, x, y, w) {
    doc.setFillColor.apply(doc, RED);
    doc.rect(x, y, w, 3.5, "F");                                  // top red rule
    doc.setTextColor(20, 20, 26); doc.setFont("helvetica", "bold");
    doc.setFontSize(30); doc.text("RED", x + w / 2, y + 32, { align: "center" });
    doc.setFontSize(13);  doc.text("F  M", x + w / 2, y + 50, { align: "center" });
    doc.setFillColor.apply(doc, RED);
    doc.rect(x, y + 58, w, 3.5, "F");                             // bottom red rule
    return y + 61.5;
  }

  // Stamp the footer (tagline + Coolstream credit + page number) on every page.
  // fgasNo (optional): Coolstream's F-Gas company certification number — shown on F-Gas certificates.
  function stampFooter(doc, fgasNo) {
    const W = doc.internal.pageSize.getWidth();
    const H = doc.internal.pageSize.getHeight();
    const n = doc.internal.getNumberOfPages();
    const credit = "Maintenance works carried out by Coolstream (UK) Ltd"
      + (fgasNo ? " · F-Gas (Refcom) No. " + fgasNo : "")
      + " on behalf of RED FM (Red Electrical Bristol Ltd).";
    for (let p = 1; p <= n; p++) {
      doc.setPage(p);
      doc.setDrawColor(225, 225, 228); doc.setLineWidth(0.5); doc.line(40, H - 40, W - 40, H - 40);
      // tagline (two-tone, italic)
      doc.setFont("times", "italic"); doc.setFontSize(10);
      doc.setTextColor(20, 20, 26); doc.text("The work that ", 40, H - 26);
      const tw = doc.getTextWidth("The work that ");
      doc.setTextColor.apply(doc, RED); doc.text("keeps you working.", 40 + tw, H - 26);
      // Coolstream credit (+ F-Gas cert number on certificates)
      doc.setFont("helvetica", "normal"); doc.setFontSize(7.5); doc.setTextColor(120, 120, 126);
      doc.text(credit, 40, H - 15);
      doc.text("Page " + p + " of " + n, W - 40, H - 15, { align: "right" });
    }
  }

  // Build a branded RED FM service-sheet PDF (Blob). Needs jsPDF + autotable loaded on the page.
  // meta: {title, subtitle, ref, date, engineer, plant, extra:[[label,value],...]}
  // sections: [{name, cols:[...], rows:[[item, v1, v2, ...]]}]  — one grid table per section, like the paper sheet.
  function buildReportPdf(meta, sections) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: "pt", format: "a4" });
    const W = doc.internal.pageSize.getWidth();
    const H = doc.internal.pageSize.getHeight();

    // ---- Header: logo left, title block right ----
    drawLogo(doc, 40, 28, 150);
    doc.setTextColor(20, 20, 26); doc.setFont("helvetica", "bold"); doc.setFontSize(16);
    doc.text(meta.title || "Service report", W - 40, 46, { align: "right" });
    doc.setFont("helvetica", "normal"); doc.setFontSize(9.5); doc.setTextColor(95, 95, 102);
    doc.text(meta.subtitle || "Border Holdings, Avonmouth", W - 40, 62, { align: "right" });
    if (meta.date) doc.text("Date: " + meta.date, W - 40, 76, { align: "right" });
    doc.setDrawColor.apply(doc, RED); doc.setLineWidth(1.5); doc.line(40, 100, W - 40, 100);

    // ---- Details block ----
    const head = [["Ref", meta.ref], ["Date", meta.date], ["Engineer", meta.engineer]]
      .concat(meta.plant ? [["Plant", meta.plant]] : []).concat(meta.extra || []);
    doc.autoTable({
      startY: 110, theme: "plain", styles: { fontSize: 9, cellPadding: 1.5 }, margin: { left: 40, right: 40 },
      body: head.filter(([, v]) => v != null && v !== "")
        .map(([k, v]) => [{ content: k + ":", styles: { fontStyle: "bold", cellWidth: 80, textColor: [80, 80, 86] } }, String(v)])
    });

    // ---- One grid table per section ----
    let y = doc.lastAutoTable.finalY + 16;
    (sections || []).forEach(s => {
      if (y > H - 90) { doc.addPage(); y = 56; }
      doc.setFont("helvetica", "bold"); doc.setFontSize(10.5); doc.setTextColor.apply(doc, RED);
      doc.text(s.name, 40, y);
      doc.autoTable({
        startY: y + 6, margin: { left: 40, right: 40 },
        head: [["Item"].concat(s.cols)],
        body: s.rows.map(r => r.map(c => String(c == null ? "" : c))),
        styles: { fontSize: 8, cellPadding: 3, overflow: "linebreak", lineColor: [225, 225, 228], lineWidth: 0.5 },
        headStyles: { fillColor: RED, textColor: 255, fontSize: 8 },
        columnStyles: { 0: { fontStyle: "bold", cellWidth: 130, textColor: [40, 40, 46] } },
        alternateRowStyles: { fillColor: [248, 248, 250] }
      });
      y = doc.lastAutoTable.finalY + 16;
    });

    // Contractor (Coolstream) logo at the bottom of F-Gas certificates
    if (meta.fgasNo) {
      const lw = 120, lh = lw * 105 / 400, ly = H - 84;
      doc.setFont("helvetica", "normal"); doc.setFontSize(6.8); doc.setTextColor(120, 120, 126);
      doc.text("Refrigeration maintenance carried out by:", 40, ly - 5);
      try { doc.addImage(COOLSTREAM_LOGO, "JPEG", 40, ly, lw, lh); } catch (e) {}
    }
    stampFooter(doc, meta.fgasNo);
    return doc.output("blob");
  }

  return {
    init, signIn, signOut, getAccount,
    getCatalogue, saveVisit, addItem, addItemsBatch,
    getServiceVisits, getFaults,
    buildReportPdf, uploadReportPdf
  };
})();
