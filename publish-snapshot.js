/* Publishes a read-only dashboard snapshot for the public view.html.
   Runs in GitHub Actions (Node 20, global fetch) with app-only Graph creds.
   Secrets required: TENANT_ID, CLIENT_ID, CLIENT_SECRET (set in repo Actions secrets).
   App registration needs APPLICATION permission Sites.Read.All (admin-consented). */
const fs = require('fs');
const TENANT = process.env.TENANT_ID, CLIENT = process.env.CLIENT_ID, SECRET = process.env.CLIENT_SECRET;
const HOST = "redbristol.sharepoint.com", SITEPATH = "/sites/REDFMMaintenance";
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
(async () => {
  TOK = await getToken();
  const sid = (await g(`/sites/${HOST}:${SITEPATH}`)).id;
  const [visits, readings, faults, cat] = await Promise.all([
    listAll(sid, "ServiceVisits"), listAll(sid, "Readings"),
    listAll(sid, "FaultRegister"), listAll(sid, "ReadingCatalogue")]);
  const snap = { generatedAt: new Date().toISOString(), visits, readings, faults, cat };
  fs.writeFileSync("data-snapshot.json", JSON.stringify(snap));
  console.log(`snapshot: ${visits.length} visits, ${readings.length} readings, ${faults.length} faults, ${cat.length} catalogue rows`);
})().catch(e => { console.error(e); process.exit(1); });
