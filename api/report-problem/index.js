/*
  RED FM — anonymous site problem intake.

  Border staff have no accounts and never will. This function is the only thing that
  holds a credential; the page that calls it carries none, which is why that page can
  be left open to anyone on site or behind a QR label.

  Everything written here lands at TriageStatus 'Awaiting triage'. Nothing raised by a
  member of the public escalates, emails Coolstream or ages against RED FM until a human
  has looked at it and put it In scope. That is deliberate — see notice 5.2.

  Required application settings (Azure portal → Static Web App → Configuration):
    REDFM_TENANT_ID       directory (tenant) id
    REDFM_CLIENT_ID       app registration (application) id
    REDFM_CLIENT_SECRET   client secret value
    REDFM_SITE_HOST       redbristol.sharepoint.com
    REDFM_SITE_PATH       /sites/REDFMMaintenance
*/

const GRAPH = "https://graph.microsoft.com/v1.0";
const MAX_PHOTO_BYTES = 4 * 1024 * 1024;

const PLANTS = ["A","B","C1","C2","C3","C4","D1","D2","D3","D4","E1","E2","A&B Plant Room","General"];
const SEVERITIES = ["Low","Medium","High"];
/* Trade decides who turns up, and it decides whether this report belongs in the
   refrigeration compliance figures at all. Anything that is not Refrigeration must not
   land in the numbers Coolstream are measured against. */
const TRADES = ["Refrigeration","Electrical","Doors & loading bays","Building fabric","Other / not sure"];

const clip = (v, n) => String(v == null ? "" : v).replace(/\s+/g, " ").trim().slice(0, n);

async function token() {
  const body = new URLSearchParams({
    client_id: process.env.REDFM_CLIENT_ID,
    client_secret: process.env.REDFM_CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials"
  });
  const r = await fetch(`https://login.microsoftonline.com/${process.env.REDFM_TENANT_ID}/oauth2/v2.0/token`,
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body });
  if (!r.ok) throw new Error("token " + r.status + " " + (await r.text()).slice(0, 300));
  return (await r.json()).access_token;
}

const g = async (tk, path, init = {}) => {
  const r = await fetch(GRAPH + path, { ...init, headers: { Authorization: "Bearer " + tk, ...(init.headers || {}) } });
  if (!r.ok) throw new Error(path + " → " + r.status + " " + (await r.text()).slice(0, 300));
  return r.status === 204 ? null : r.json();
};

/* Sequential SR- reference. The volume here is a handful a week, so a read-then-write is
   safe; if the read fails we fall back to a timestamp rather than losing the report. */
async function nextRef(tk, siteId, listId) {
  try {
    const q = `/sites/${siteId}/lists/${listId}/items?$expand=fields($select=Title)&$top=50&$orderby=id desc`;
    const j = await g(tk, q);
    let max = 1000;
    for (const it of (j.value || [])) {
      const m = /^SR-(\d+)$/.exec(((it.fields || {}).Title) || "");
      if (m) max = Math.max(max, parseInt(m[1], 10));
    }
    return "SR-" + (max + 1);
  } catch (e) {
    return "SR-" + Date.now().toString(36).toUpperCase();
  }
}

module.exports = async function (context, req) {
  const reply = (status, body) => {
    context.res = { status, headers: { "Content-Type": "application/json", "Cache-Control": "no-store" }, body };
  };

  try {
    const b = req.body || {};

    const description = clip(b.description, 2000);
    const reportedBy  = clip(b.reportedBy, 120);
    const placeCode   = clip(b.placeCode, 60);
    const place       = clip(b.place, 120);
    const contact     = clip(b.contact, 160);
    const via         = clip(b.via, 40) === "QR label" ? "QR label" : "Dashboard";
    const severity    = SEVERITIES.includes(clip(b.severity, 10)) ? clip(b.severity, 10) : "Medium";
    const trade       = TRADES.includes(clip(b.trade, 40)) ? clip(b.trade, 40) : "Other / not sure";

    if (description.length < 10) return reply(400, { error: "Tell us a bit more about what is wrong." });
    if (!reportedBy)             return reply(400, { error: "We need a name." });
    if (place.length < 2)        return reply(400, { error: "We need to know where it is." });

    /* Whatever they typed is kept verbatim. If it happens to name a chamber or plant room
       we tag it too, so repeat detection and per-asset history still work; if it does not,
       the location survives in full in the description rather than being forced into a box. */
    const plant  = PLANTS.includes(placeCode) ? placeCode : "General";
    const assetId = PLANTS.includes(placeCode) && placeCode !== "General" ? placeCode : "";

    let photo = String(b.photo || "");
    if (photo && !/^data:image\/(jpeg|jpg|png|webp);base64,/i.test(photo)) photo = "";
    let photoBuf = null;
    if (photo) {
      photoBuf = Buffer.from(photo.split(",")[1] || "", "base64");
      if (!photoBuf.length || photoBuf.length > MAX_PHOTO_BYTES) photoBuf = null;
    }

    const tk = await token();
    const site = await g(tk, `/sites/${process.env.REDFM_SITE_HOST}:${process.env.REDFM_SITE_PATH}`);
    const lists = await g(tk, `/sites/${site.id}/lists?$select=id,name,displayName`);
    const list = (lists.value || []).find(l => l.name === "FaultRegister" || l.displayName === "FaultRegister");
    if (!list) throw new Error("FaultRegister list not found");

    const ref = await nextRef(tk, site.id, list.id);
        // Azure runs in UTC; between 00:00 and 01:00 BST that is still yesterday, and FaultDate starts the 48h clock in clause 4.1.
    const today = (() => { try { const d = new Date().toLocaleDateString("en-CA", { timeZone: "Europe/London" }); if (/^\d{4}-\d{2}-\d{2}$/.test(d)) return d; } catch (e) {} return new Date().toISOString().slice(0, 10); })();

    const detail = description
      + `\n\nLocation as given: ${place}`
      + `\nReported as: ${trade}`
      + `\n\nReported by ${reportedBy}${contact ? ` (${contact})` : ""} via the ${via.toLowerCase()}.`
      + `\nRaised by a member of site staff, not by an attending engineer. Not yet triaged.`;

    await g(tk, `/sites/${site.id}/lists/${list.id}/items`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        fields: {
          Title: ref,
          FaultDate: today,
          Plant: plant,
          AssetId: assetId,
          Severity: severity,
          Description: detail,
          Stage: "Reported",
          RaisedBy: reportedBy,
          Trade: trade,
          SourceType: "Site QR",
          TriageStatus: "Awaiting triage",
          RepeatCount: 1
        }
      })
    });

    /* The photo is the whole reason this exists rather than a phone call, but a failed
       upload must never lose the report — the record is already in by this point. */
    let photoNote = "";
    if (photoBuf) {
      try {
        const drives = await g(tk, `/sites/${site.id}/drives?$select=id,name`);
        const lib = (drives.value || []).find(d => d.name === "Faults") || (drives.value || [])[0];
        const name = `${ref}_${today}_${(assetId || plant).replace(/[^A-Za-z0-9]+/g, "")}.jpg`;
        const up = await fetch(`${GRAPH}/drives/${lib.id}/root:/${encodeURIComponent(name)}:/content`, {
          method: "PUT",
          headers: { Authorization: "Bearer " + tk, "Content-Type": "image/jpeg" },
          body: photoBuf
        });
        photoNote = up.ok ? "" : " (photo did not attach)";
      } catch (e) { photoNote = " (photo did not attach)"; }
    }

    context.log(`site report ${ref} · ${trade} · ${place} · ${severity} · via ${via}${photoNote}`);
    return reply(200, { ref, ok: true });

  } catch (err) {
    context.log.error("report-problem failed: " + (err && err.message));
    return reply(500, { error: "Could not log that. Please tell RED FM directly." });
  }
};
