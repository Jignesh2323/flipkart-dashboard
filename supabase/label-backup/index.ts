// Cloud backup of the daily label run (label-run.ps1) for when the PC is off (user, 19/29 Oct).
// pg_cron calls this once per account a few minutes after each slot: POST ?slot=10-15|11-45|01-45&acc=AVEDIA
// with header x-key = fk_cfg.cron_key. Does nothing on Sunday / holidays or when the PC already started that
// slot (label_mark src 'pc'). Otherwise same steps as the PC: RTD due PACKED orders, label all APPROVED
// orders, crop to the 3x5 label, save to Drive (Apps Script) as <ACC>-cloud-HHMM.pdf, then RTD.
import { PDFDocument, StandardFonts, rgb } from "npm:pdf-lib@1.17.1";
// SKUs the listings API cannot look up ('/' in the SKU gives 404): GST rate from the HSN fix sheets
const TAX: Record<string, number> = { "PVD-APG-B/W/Purple": 18, "PVD-APG-G/W/Pink": 18 };

const API = "https://api.flipkart.net";
const SB = Deno.env.get("SUPABASE_URL")!;
// legacy service_role JWT, or the new sb_secret_ key (apikey header only)
const SK = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") || Object.values(JSON.parse(Deno.env.get("SUPABASE_SECRET_KEYS") || "{}"))[0] as string;
const SH: Record<string, string> = SK.startsWith("ey") ? { apikey: SK, Authorization: `Bearer ${SK}` } : { apikey: SK };
const HOLIDAYS = ["2026-11-09", "2026-11-10", "2026-11-11"];
const SLOT_DIR: Record<string, string> = { "10-15": "1) 10-15", "11-45": "2) 11-45", "01-45": "3) 01-45" };
const BOX = [185, 462, 410, 817];
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
const ist = () => new Date(Date.now() + 5.5 * 3600e3); // read with getUTC* = IST wall clock
const istDay = (d: Date) => d.toISOString().slice(0, 10);

async function db(path: string, init: RequestInit = {}) {
  const r = await fetch(`${SB}/rest/v1/${path}`, { ...init, headers: { ...SH, "Content-Type": "application/json", ...(init.headers || {}) } });
  return r.status === 204 || init.method === "POST" ? null : r.json();
}
const cfg = async (k: string) => (await db(`fk_cfg?key=eq.${k}&select=value`))[0]?.value;

Deno.serve(async (req) => {
  const u = new URL(req.url), slot = u.searchParams.get("slot") || "", acc = (u.searchParams.get("acc") || "").toUpperCase();
  if (req.headers.get("x-key") !== await cfg("cron_key")) return new Response("no", { status: 401 });
  if (SLOT_DIR[slot] && u.searchParams.get("summary")) return new Response(await summary(slot));
  if (!SLOT_DIR[slot] || !acc) return new Response("bad args", { status: 400 });
  const lines: string[] = [];
  const log = (m: string) => { lines.push(m); };
  try { await run(slot, acc, log); } catch (e) { log(`${acc} ERROR ${e}`); }
  if (lines.length) await db("label_log", { method: "POST", body: JSON.stringify(lines.map((msg) => ({ slot, acc, msg }))) });
  return new Response(JSON.stringify(lines), { headers: { "Content-Type": "application/json" } });
});

async function run(slot: string, acc: string, log: (m: string) => void) {
  const now = ist(), day = istDay(now);
  if (now.getUTCDay() === 0 || HOLIDAYS.includes(day)) return;
  if ((await db(`label_mark?day=eq.${day}&slot=eq.${slot}&src=eq.pc&select=day`)).length) return; // PC did this slot
  await db("label_mark", { method: "POST", headers: { Prefer: "resolution=ignore-duplicates" }, body: JSON.stringify({ day, slot, src: "cloud" }) });
  const c = (await db(`fk_cred?acc=eq.${acc}&select=app_id,app_secret`))[0];
  if (!c) throw "no credentials";
  const tk = await (await fetch(`${API}/oauth-service/oauth/token?grant_type=client_credentials&scope=Seller_Api`, { headers: { Authorization: "Basic " + btoa(`${c.app_id}:${c.app_secret}`) } })).json();
  const H = { Authorization: `Bearer ${tk.access_token}`, "Content-Type": "application/json" };
  const get = async (p: string) => (await fetch(`${API}${p}`, { headers: H })).json();
  const post = async (p: string, b: unknown) => (await fetch(`${API}${p}`, { method: "POST", headers: H, body: JSON.stringify(b) })).json();
  const filter = async (states: string[], extra = {}) => {
    let r = await post("/sellers/v3/shipments/filter", { filter: { type: "preDispatch", states, ...extra }, pagination: { pageSize: 20 } });
    const all = [...(r.shipments || [])];
    while (r.hasMore && r.nextPageUrl) { r = await get(`/sellers${r.nextPageUrl}`); all.push(...(r.shipments || [])); }
    return all;
  };
  const labels = async (ids: string[]) => new Uint8Array(await (await fetch(`${API}/sellers/v3/shipments/${ids.join(",")}/labels`, { headers: { Authorization: H.Authorization } })).arrayBuffer());
  const hhmm = `${String(now.getUTCHours()).padStart(2, "0")}${String(now.getUTCMinutes()).padStart(2, "0")}`;
  const dueToday = (s: any) => !s.hold && istDay(new Date(new Date(s.dispatchAfterDate).getTime() + 5.5 * 3600e3)) <= day;
  const info = (s: any) => `${s.orderItems?.[0]?.orderId} ${s.orderItems?.[0]?.sku}`;

  // RTD with Flipkart's rate limit retried; returns ids that are RTD now
  const dispatch = async (group: any[]) => {
    const done: string[] = [];
    const byLoc: Record<string, any[]> = {};
    for (const s of group) (byLoc[s.locationId] ||= []).push(s);
    for (const [loc, ss] of Object.entries(byLoc)) for (let i = 0; i < ss.length; i += 20) {
      let left = ss.slice(i, i + 20).map((s) => s.shipmentId);
      for (let k = 0; k < 4 && left.length; k++) {
        if (k) await sleep(8000);
        const rr = await post("/sellers/v3/shipments/dispatch", { shipmentIds: left, locationId: loc });
        const busy: string[] = [];
        for (const x of rr.shipments || []) {
          if (x.status === "SUCCESS") done.push(x.shipmentId);
          else if (/Too many requests/.test(x.errorMessage || "") && k < 3) busy.push(x.shipmentId);
          else {
            const st = (await get(`/sellers/v3/shipments?shipmentIds=${x.shipmentId}`)).shipments?.[0];
            if (st?.orderItems?.[0]?.status === "READY_TO_DISPATCH") done.push(x.shipmentId);
            else log(`${acc} RTD FAIL ${x.shipmentId} [${st ? info(st) : ""}]: ${x.status} ${x.errorMessage || ""}`);
          }
        }
        left = busy;
      }
    }
    return done;
  };

  // 1) PACKED but due and not RTD (labels made earlier / by staff): print first, then RTD
  const packed = (await filter(["PACKED"])).filter(dueToday);
  if (packed.length) {
    const pdf = await labels(packed.map((s) => s.shipmentId));
    const done = await dispatch(packed);
    if (done.length) { await save(acc, slot, day, `${acc}-RTD-cloud-${hhmm}.pdf`, [done.length < packed.length ? await labels(done) : pdf]); log(`${acc} : ${done.length} purane PACKED order RTD kiye`); }
  }

  // 2) all APPROVED orders
  const from = new Date(Date.now() - 30 * 86400e3).toISOString().slice(0, 10) + "T00:00:00.000+05:30";
  const to = new Date(Date.now() + 86400e3).toISOString().slice(0, 10) + "T23:59:59.000+05:30";
  const ids = [...new Set((await filter(["APPROVED"], { orderDate: { from, to } })).map((s) => s.shipmentId))];
  if (!ids.length) { log(`${acc} : koi naya order nahi`); return; }
  const full: any[] = [];
  for (let i = 0; i < ids.length; i += 25) full.push(...((await get(`/sellers/v3/shipments?shipmentIds=${ids.slice(i, i + 25).join(",")}`)).shipments || []));
  const taxC: Record<string, number> = {}, hsnRate: Record<string, number> = {};
  const hsnOf: Record<string, string> = {};
  for (const s of full) for (const it of s.orderItems) hsnOf[it.sku] = it.hsn;
  for (const sku of Object.keys(hsnOf).sort((a, b) => Number(a.includes("/")) - Number(b.includes("/")))) {
    let rate: number | undefined;
    try { const m = JSON.stringify(await get(`/sellers/listings/v3/${encodeURIComponent(sku)}`)).match(/"tax_code":\s*"([^"]+)"/); if (m) rate = Number(m[1].replace(/[^\d.]/g, "")); } catch { /* fallbacks below */ }
    if (rate === undefined) rate = TAX[sku] ?? hsnRate[hsnOf[sku]];
    if (rate === undefined) throw `GST rate nahi mila: ${sku} (HSN ${hsnOf[sku]})`;
    taxC[sku] = rate; hsnRate[hsnOf[sku]] ??= rate;
  }
  const today = day;
  const reqs = full.map((s) => {
    const qty = s.orderItems.reduce((a: number, x: any) => a + x.quantity, 0);
    const r: any = {
      shipmentId: s.shipmentId, locationId: s.locationId,
      subShipments: s.subShipments.map((ss: any) => { const d = ss.packages[0].dimensions; return { subShipmentId: ss.subShipmentId, dimensions: { length: +d.length, breadth: +d.breadth, height: qty > 1 ? Math.round(d.height * qty * 10) / 10 : +d.height, weight: qty > 1 ? Math.round(d.weight * qty * 1000) / 1000 : +d.weight } }; }),
      invoices: [{ orderId: s.orderItems[0].orderId, invoiceDate: today }],
      taxItems: s.orderItems.map((x: any) => ({ orderItemId: x.orderItemId, taxRate: taxC[x.sku], quantity: x.quantity })),
    };
    const sn = s.orderItems.filter((x: any) => /GPU Stand/.test(x.sku)).map((x: any) => ({ orderItemId: x.orderItemId, serialNumbers: Array.from({ length: x.quantity }, (_, k) => [`GPUSTAND${x.orderItemId}${k + 1}`]) }));
    if (sn.length) r.serialNumbers = sn;
    return r;
  });
  const failed: string[] = [], later: string[] = [];
  const nowMs = Date.now(), tomorrow = istDay(new Date(nowMs + 5.5 * 3600e3 + 86400e3));
  for (let i = 0; i < reqs.length; i += 25) {
    const resp = await post("/sellers/v3/shipments/labels", { shipments: reqs.slice(i, i + 25) });
    for (const x of resp.shipments || []) if (x.status && x.status !== "SUCCESS") {
      const why = `${x.errorMessage || ""} ${x.skippedReason || ""}`.trim();
      const sh = full.find((s) => s.shipmentId === x.shipmentId);
      const notDue = sh && (new Date(sh.dispatchAfterDate).getTime() > nowMs || istDay(new Date(new Date(sh.dispatchByDate).getTime() + 5.5 * 3600e3)) > tomorrow);
      if (notDue || /incorrect quantity/.test(why)) later.push(x.shipmentId);
      else failed.push(`${x.shipmentId} [${sh ? info(sh) : ""}]: ${x.status} ${why}`);
    }
  }
  if (failed.length) log(`${acc} label FAIL: ${failed.join(" | ")}`);
  if (later.length) log(`${acc} : ${later.length} order abhi due nahi (agle slot me)`);
  const ok = ids.filter((id) => !later.includes(id) && !failed.some((f) => f.startsWith(id)));
  if (!ok.length) return;
  for (let t = 0; t < 9; t++) { // wait until PACKED, ~90 s (function time limit)
    await sleep(10000);
    const st: any[] = [];
    for (let i = 0; i < ok.length; i += 25) st.push(...((await get(`/sellers/v3/shipments?shipmentIds=${ok.slice(i, i + 25).join(",")}`)).shipments || []));
    if (!st.some((s) => s.orderItems.some((x: any) => x.status !== "PACKED"))) break;
  }
  const pdfs: Uint8Array[] = [];
  for (let i = 0; i < ok.length; i += 25) pdfs.push(await labels(ok.slice(i, i + 25)));
  await save(acc, slot, day, `${acc}-cloud-${hhmm}.pdf`, pdfs);
  const rtd = full.filter((s) => ok.includes(s.shipmentId) && dueToday(s));
  const done = await dispatch(rtd);
  log(`DONE: ${acc} ${ok.length} labels, ${done.length} RTD${rtd.length < ok.length ? `, ${ok.length - rtd.length} RTD baad me` : ""}`);
}

// crop every A4 page to the 3x5 label and upload one PDF to Drive (Flipkart Labels/<dd-MM-yyyy>/<slot dir>/)
async function save(acc: string, slot: string, day: string, name: string, pdfs: Uint8Array[]) {
  const out = await PDFDocument.create();
  for (const b of pdfs) {
    const src = await PDFDocument.load(b);
    for (const p of await out.copyPages(src, src.getPageIndices())) {
      p.setMediaBox(BOX[0], BOX[1], BOX[2] - BOX[0], BOX[3] - BOX[1]);
      p.setCropBox(BOX[0], BOX[1], BOX[2] - BOX[0], BOX[3] - BOX[1]);
      out.addPage(p);
    }
  }
  await upload(slot, day, name, await out.save());
}

async function upload(slot: string, day: string, name: string, bytes: Uint8Array) {
  let bin = ""; for (let i = 0; i < bytes.length; i += 32768) bin += String.fromCharCode(...bytes.subarray(i, i + 32768));
  const [y, m, d] = day.split("-");
  const r = await fetch(await cfg("drive_url"), { method: "POST", body: JSON.stringify({ token: await cfg("drive_token"), path: `${d}-${m}-${y}/${SLOT_DIR[slot]}`, name, b64: btoa(bin) }) });
  const j = await r.json().catch(() => ({ ok: false, error: `HTTP ${r.status}` }));
  if (!j.ok) throw `Drive upload fail: ${j.error}`;
}

// Summary.pdf for a slot the cloud ran (PC off): same table as the PC's Summary.png. Labels per slot from
// label_log, failures from each account's latest cloud call of this slot, Pending Handover live from Flipkart.
const ACCS = ["ARKZO", "VUSTICA", "BRAINLE", "KIDOCA", "TAUDI", "ZOTIMO", "PIVDO", "WISKA", "AVEDIA"];
const ORDER = ["10-15", "11-45", "01-45"];
async function summary(slot: string) {
  const now = ist(), day = istDay(now);
  const marks = (await db(`label_mark?day=eq.${day}&slot=eq.${slot}&select=src`)).map((m: any) => m.src);
  if (marks.includes("pc") || !marks.includes("cloud")) return "skip: PC made it / cloud did not run";
  const from = new Date(`${day}T00:00:00+05:30`).toISOString();
  const rows: any[] = await db(`label_log?at=gte.${from}&select=at,slot,acc,msg&order=id`);
  const ci = ORDER.indexOf(slot);
  const R: Record<string, Record<string, number>> = {}, fails: Record<string, string[]> = {}, last: Record<string, string> = {};
  for (const a of ACCS) { R[a] = {}; fails[a] = []; }
  for (const r of rows) {
    if (!R[r.acc] || ORDER.indexOf(r.slot) > ci) continue;
    let m = r.msg.match(/DONE: \w+ \d+ labels, (\d+) RTD/) || r.msg.match(/: (\d+) purane PACKED order RTD kiye/);
    if (m) R[r.acc][r.slot] = (R[r.acc][r.slot] || 0) + Number(m[1]);
    if (r.slot !== slot) continue;
    if (last[r.acc] !== r.at) { last[r.acc] = r.at; fails[r.acc] = []; } // one call = one insert = same time
    if ((m = r.msg.match(/RTD FAIL \S+ \[(.*?)\]:? ?(.*)/))) fails[r.acc].push(`${m[1]}: LABEL NIKAL CHUKA, RTD BAAKI - ${m[2]}`);
    else if ((m = r.msg.match(/label FAIL: (.*)/))) for (const x of m[1].split(" | ")) { const k = x.match(/\[(.+?)\]: (?:\w+ )?(.*)/); fails[r.acc].push(k ? `${k[1]}: LABEL + RTD BAAKI - ${k[2]}` : `LABEL + RTD BAAKI - ${x}`); }
    else if ((m = r.msg.match(/ERROR (.*)/))) fails[r.acc].push(`Pura account: ${m[1]}`);
  }
  // live: Pending Handover and due orders still PACKED (not RTD)
  const H: Record<string, string | number> = {};
  for (const a of ACCS) {
    try {
      const c = (await db(`fk_cred?acc=eq.${a}&select=app_id,app_secret`))[0];
      const tk = await (await fetch(`${API}/oauth-service/oauth/token?grant_type=client_credentials&scope=Seller_Api`, { headers: { Authorization: "Basic " + btoa(`${c.app_id}:${c.app_secret}`) } })).json();
      const hd = { Authorization: `Bearer ${tk.access_token}`, "Content-Type": "application/json" };
      const count = async (st: string) => {
        let r = await (await fetch(`${API}/sellers/v3/shipments/filter`, { method: "POST", headers: hd, body: JSON.stringify({ filter: { type: "preDispatch", states: [st] }, pagination: { pageSize: 20 } }) })).json();
        const all = [...(r.shipments || [])];
        while (r.hasMore && r.nextPageUrl) { r = await (await fetch(`${API}/sellers${r.nextPageUrl}`, { headers: hd })).json(); all.push(...(r.shipments || [])); }
        return all;
      };
      H[a] = (await count("READY_TO_DISPATCH")).length;
      const stuck = (await count("PACKED")).filter((s) => !s.hold && istDay(new Date(new Date(s.dispatchAfterDate).getTime() + 5.5 * 3600e3)) <= day);
      if (stuck.length) fails[a].push(`${stuck.length} order LABEL NIKAL CHUKA, RTD BAAKI (PACKED) - Seller Hub me RTD karein`);
    } catch { H[a] = "?"; }
  }
  // draw (A4 landscape-ish page, same columns as the PC image)
  const pdf = await PDFDocument.create(), page = pdf.addPage([640, 760]);
  const f = await pdf.embedFont(StandardFonts.Helvetica), fb = await pdf.embedFont(StandardFonts.HelveticaBold);
  const txt = (s: string, x: number, y: number, o: any = {}) => page.drawText(s.replace(/[^\x20-\x7E]/g, ""), { x, y, size: o.size || 11, font: o.bold ? fb : f, color: o.color || rgb(0, 0, 0), maxWidth: o.w });
  const grey = rgb(0.45, 0.45, 0.45), red = rgb(0.8, 0, 0), orange = rgb(0.9, 0.47, 0);
  const [y4, mo, d] = day.split("-");
  txt(`Flipkart Labels  -  ${d}-${mo}-${y4}  -  ${slot.replace(/^0/, "")}  (CLOUD - PC band tha)`, 24, 725, { size: 15, bold: true });
  txt(`Supabase se bana: ${String(now.getUTCHours()).padStart(2, "0")}:${String(now.getUTCMinutes()).padStart(2, "0")}`, 24, 705, { color: grey, size: 10 });
  const cols: [string, number][] = [["Account", 24], ["10-15", 150], ["11-45", 220], ["1-45", 290], ["Failed", 360], ["Pending Handover", 440]];
  let y = 670; page.drawRectangle({ x: 16, y: y - 8, width: 608, height: 26, color: rgb(0.95, 0.95, 0.95) });
  for (const [c, x] of cols) txt(c, x, y, { bold: true });
  const tot: Record<string, number> = {}; let hTot = 0;
  for (const a of ACCS) {
    y -= 26; txt(a, 24, y);
    ORDER.forEach((s, i) => { const v = R[a][s] || 0; tot[s] = (tot[s] || 0) + v; txt(i > ci ? "-" : String(v), cols[i + 1][1], y, i === ci ? { bold: true } : { color: grey }); });
    txt(fails[a].length ? String(fails[a].length) : "-", 360, y, { color: fails[a].length ? red : grey });
    txt(String(H[a]), 440, y); hTot += Number(H[a]) || 0;
  }
  y -= 28; txt("Total", 24, y, { bold: true }); ORDER.forEach((s, i) => { if (i <= ci) txt(String(tot[s]), cols[i + 1][1], y, { bold: true }); }); txt(String(hTot), 440, y, { bold: true });
  const fl = ACCS.flatMap((a) => fails[a].map((x) => [a, x]));
  if (fl.length) {
    y -= 40; txt("Failed orders - Seller Hub me khud process karein:", 24, y, { bold: true, color: red, size: 12 });
    for (const [a, x] of fl) {
      y -= 30; if (y < 30) break;
      const t = x.includes("LABEL + RTD BAAKI") ? ["LABEL + RTD DONO BAAKI", red] : x.includes("LABEL NIKAL CHUKA") ? ["LABEL NIKAL CHUKA - RTD BAAKI", orange] : null;
      let tx = 24;
      if (t) { const w = fb.widthOfTextAtSize(t[0] as string, 9) + 10; page.drawRectangle({ x: 24, y: y - 5, width: w, height: 17, color: t[1] as any }); txt(t[0] as string, 29, y, { bold: true, size: 9, color: rgb(1, 1, 1) }); tx = 24 + w + 8; }
      txt(`${a}  -  ${x.replace(/LABEL \+ RTD BAAKI - |LABEL NIKAL CHUKA, RTD BAAKI( \(PACKED\))? - /, "")}`, tx, y, { size: 9, w: 600 - tx });
    }
  }
  await upload(slot, day, "Summary-cloud.pdf", await pdf.save());
  return "ok";
}
