/* Publishes a read-only dashboard snapshot for the public view.html, AND syncs the
   report PDFs from SharePoint into reports/ + reports-manifest.json so the public
   reports library (reports.html) can list & download them without a login.

   Runs in GitHub Actions (Node 20, global fetch) with app-only Graph creds.
   Secrets required: TENANT_ID, CLIENT_ID, CLIENT_SECRET (set in repo Actions secrets).
   App registration needs APPLICATION permission Sites.Read.All (admin-consented). */
const fs = require('fs');
const TENANT = process.env.TENANT_ID, CLIENT = process.env.CLIENT_ID, SECRET = process.env.CLIENT_SECRET;
const HOST = "redbristol.sharepoint.com", SITEPATH = "/sites/REDFMMaintenance";
const RAW_BASE = "https://raw.githubusercontent.com/Red-works1/redfm-maintenance/main/reports/";
const PHOTO_BASE = "https://raw.githubusercontent.com/Red-works1/redfm-maintenance/main/photos/";
let TOK;

async function getToken() {
  const r = await fetch(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`, {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: CLIENT, client_secret: SECRET,
      scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" })
  });
  const j = await r.json();
  if (!j.access_token) throw new Error("token error: " + JSON.stringify(j));
  return j.access_token;
}
async function g(path) {
  const r = await fetch("https://graph.microsoft.com/v1.0" + path, { headers: { Authorization: "Bearer " + TOK } });
  if (!r.ok) throw new Error("graph " + r.status + ": " + (await r.text()));
  return r.json();
}
async function listAll(sid, title) {
  const lists = (await g(`/sites/${sid}/lists?$select=id,displayName&$top=100`)).value;
  const lid = (lists.find(l => l.displayName === title) || {}).id;
  if (!lid) throw new Error("list not found: " + title);
  let url = `/sites/${sid}/lists/${lid}/items?$expand=fields&$top=500`, out = [];
  while (url) {
    const p = await g(url);
    p.value.forEach(i => out.push(i.fields));
    url = p["@odata.nextLink"] ? p["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "") : null;
  }
  return out;
}

// Same reader, but for the lists added by the Contract Performance Notice: one that does not
// exist yet (or is empty) must publish as [] rather than failing the whole run.
async function listAllSafe(sid, title) {
  try { return await listAll(sid, title); }
  catch (e) { console.error("list " + title + " unavailable: " + e.message); return []; }
}

// Group a report file into a library heading by its filename prefix.
function classify(name) {
  const n = name.toLowerCase();
  if (n.startsWith("weekly_")) return "Weekly checks";
  if (n.startsWith("quarterly_")) return "Quarterly service";
  if (n.startsWith("sixmonthly_") || n.includes("condenser")) return "Six-monthly condenser wash";
  if (n.startsWith("fgas_") || n.includes("f-gas") || n.includes("fgas")) return "F-Gas leak testing";
  if (n.includes("cutout") || n.includes("safety")) return "Annual safety cut-outs";
  if (n.includes("probe")) return "Temperature probes";
  if (n.includes("leakdetection") || n.includes("leak_detection") || n.includes("ableak")) return "A&B leak detection";
  if (n.includes("oil")) return "Oil analysis";
  if (n.includes("eicr") || n.includes("fixedwiring") || n.includes("fixed_wiring")) return "Fixed wiring (EICR)";
  if (n.includes("emergency") || n.includes("emerglight")) return "Emergency lighting";
  if (n.includes("firealarm") || n.includes("fire_alarm") || n.includes("fire alarm")) return "Fire alarm services";
  if (n.startsWith("fault_")) return "Fault reports";
  if (n.startsWith("border_compliance") || n.includes("compliance")) return "Monthly compliance reports";
  return "Other reports";
}

async function driveIdFor(sid, title) {
  const lists = (await g(`/sites/${sid}/lists?$select=id,displayName&$top=100`)).value;
  const lid = (lists.find(l => l.displayName === title) || {}).id;
  if (!lid) return null;
  const drive = await g(`/sites/${sid}/lists/${lid}/drive`);
  return drive.id;
}

// Download every .pdf from a document library into reports/ and return manifest rows.
async function syncLib(sid, title) {
  const did = await driveIdFor(sid, title);
  if (!did) { console.log("library not found: " + title); return []; }
  let url = `/drives/${did}/root/children?$top=200`;
  const out = [];
  while (url) {
    const p = await g(url);
    for (const it of p.value) {
      if (!it.file || !/\.pdf$/i.test(it.name)) continue;   // PDFs only (skip .heic photos, folders)
      const dl = it["@microsoft.graph.downloadUrl"];
      if (!dl) continue;
      const buf = Buffer.from(await (await fetch(dl)).arrayBuffer());
      fs.writeFileSync("reports/" + it.name, buf);
      out.push({
        name: it.name,
        type: classify(it.name),
        date: it.lastModifiedDateTime || it.createdDateTime || null,
        size: it.size || buf.length,
        url: RAW_BASE + encodeURIComponent(it.name)
      });
    }
    url = p["@odata.nextLink"] ? p["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "") : null;
  }
  console.log(`synced ${out.length} PDFs from "${title}"`);
  return out;
}

/* Photos reported from site live in the Faults document library, which Border staff have no
   account for — a SharePoint link is useless to them. Same answer as the report PDFs: copy the
   images out into the repo so the public dashboard can show them with no login.
   Filenames are <ref>_<date>_<slug>[_n].jpg, so the fault ref comes straight off the name and
   no lookup is needed. .heic is skipped deliberately: browsers cannot display it, and anything
   arriving through the reporting form is already re-encoded to JPEG on the phone. */
async function syncPhotos(sid, title) {
  const did = await driveIdFor(sid, title);
  if (!did) { console.log("library not found: " + title); return []; }
  let url = `/drives/${did}/root/children?$top=200`;
  const out = [];
  while (url) {
    const p = await g(url);
    for (const it of p.value) {
      if (!it.file || !/\.(jpe?g|png|webp)$/i.test(it.name)) continue;
      const m = it.name.match(/^([A-Za-z]+-[A-Za-z0-9]+)_/);
      const dl = it["@microsoft.graph.downloadUrl"];
      if (!dl) continue;
      try {
        // Already copied and unchanged? Skip the download so the hourly run stays cheap and
        // git sees no diff.
        const path = "photos/" + it.name;
        if (!fs.existsSync(path) || fs.statSync(path).size !== it.size) {
          fs.writeFileSync(path, Buffer.from(await (await fetch(dl)).arrayBuffer()));
        }
        out.push({
          name: it.name,
          ref: m ? m[1].toUpperCase() : "",
          date: it.lastModifiedDateTime || it.createdDateTime || null,
          size: it.size || 0,
          url: PHOTO_BASE + encodeURIComponent(it.name)
        });
      } catch (e) { console.error("photo " + it.name + " failed: " + e.message); }
    }
    url = p["@odata.nextLink"] ? p["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "") : null;
  }
  console.log(`synced ${out.length} photos from "${title}"`);
  return out;
}

(async () => {
  TOK = await getToken();
  const sid = (await g(`/sites/${HOST}:${SITEPATH}`)).id;

  // 1) the dashboard data snapshot (unchanged)
  const [visits, readings, faults, cat] = await Promise.all([
    listAll(sid, "ServiceVisits"), listAll(sid, "Readings"),
    listAll(sid, "FaultRegister"), listAll(sid, "ReadingCatalogue")]);
  // RDM alarms (notice 5.3), parameter changes (clause 8.8) and the forward works register
  const [rdmalarms, parameterchanges, forwardworks] = await Promise.all([
    listAllSafe(sid, "RDMAlarms"), listAllSafe(sid, "ParameterChanges"),
    listAllSafe(sid, "ForwardWorks")]);
  // Photos go into the same snapshot the dashboards already fetch, so neither page needs a
  // second request. Done before the write so an empty result still publishes as [].
  fs.mkdirSync("photos", { recursive: true });
  let photos = [];
  try { photos = await syncPhotos(sid, "Faults"); }
  catch (e) { console.error("photo sync failed: " + e.message); }

  const snap = { generatedAt: new Date().toISOString(), visits, readings, faults, cat, rdmalarms, parameterchanges, forwardworks, photos };
  fs.writeFileSync("data-snapshot.json", JSON.stringify(snap));
  fs.writeFileSync("photos-manifest.json", JSON.stringify({ generatedAt: new Date().toISOString(), photos }));
  console.log(`snapshot: ${visits.length} visits, ${readings.length} readings, ${faults.length} faults, ${cat.length} catalogue rows, ${rdmalarms.length} RDM alarms, ${parameterchanges.length} parameter changes, ${forwardworks.length} forward works, ${photos.length} photos`);

  // 2) the public reports library
  fs.mkdirSync("reports", { recursive: true });
  let reports = [];
  for (const lib of ["Reports", "Client Reports", "Faults"]) {
    try { reports = reports.concat(await syncLib(sid, lib)); }
    catch (e) { console.error("sync " + lib + " failed: " + e.message); }
  }
  // Also list any PDFs committed straight into reports/ (e.g. building-compliance certs
  // seeded into the repo) that aren't already sourced from a SharePoint library.
  try {
    const seen = new Set(reports.map(r => r.name));
    for (const fn of fs.readdirSync("reports")) {
      if (!/\.pdf$/i.test(fn) || seen.has(fn)) continue;
      const st = fs.statSync("reports/" + fn);
      reports.push({ name: fn, type: classify(fn), date: st.mtime.toISOString(), size: st.size, url: RAW_BASE + encodeURIComponent(fn) });
      console.log("included committed report: " + fn);
    }
  } catch (e) { console.error("local reports scan failed: " + e.message); }
  reports.sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  fs.writeFileSync("reports-manifest.json", JSON.stringify({ generatedAt: new Date().toISOString(), reports }));
  console.log(`reports library: ${reports.length} PDFs total`);
})().catch(e => { console.error(e); process.exit(1); });
