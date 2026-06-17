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
