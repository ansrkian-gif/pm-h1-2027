import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import XLSX from "xlsx";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");
const extractPath = process.env.EXTRACT_PATH || path.join(root, "cloud", "extract.xlsx");
const fallbackExtract = path.join(root, "PM H1 2027.xlsx");
const src = fs.existsSync(extractPath) ? extractPath : fallbackExtract;
const outDir = path.join(root, "dashboard");
const templatePath = path.join(outDir, "template.html");

const ACTIVE = "OGK Active General";
const PASSIVE = "OGK Passive General";
const SMALL = "OGK Active Small Cell /Book RRU/Easy Macro";
const CATEGORY = "OGK Active and Passive Routine Maintenance";

const FME_ALIAS = {
  "Mohd Yasin Mohd Matin Ansari": "Nabijohn Piyarjan Piyarjan",
};
function canonFme(name) {
  const t = String(name || "").trim();
  return FME_ALIAS[t] || t;
}

function ymd(d) {
  return d.toISOString().slice(0, 10);
}
function parseDay(ct) {
  if (!ct) return null;
  const s = String(ct).slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return null;
  return new Date(s + "T00:00:00Z");
}
function weekStart(d) {
  const x = new Date(d.getTime());
  x.setUTCDate(x.getUTCDate() - x.getUTCDay());
  return x;
}

const wb = XLSX.readFile(src);
const sheet = wb.Sheets[wb.SheetNames[0]];
const raw = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
const pm = raw.filter((r) => r["Task Type"] === "PM" && r["Task Category"] === CATEGORY);

const rows = [];
for (const r of pm) {
  const day = parseDay(r["Complete Time"]);
  if (!day) continue;
  const op = `${r["Accept Operator"]}|${r["Arrive Operator"]}|${r["Complete Operator"]}`;
  rows.push({
    taskId: String(r["Task Id"] || ""),
    siteId: String(r["Site ID"] || "").trim(),
    subcat: String(r["Task Subcategory"] || "").trim(),
    fme: canonFme(r["Assign To FME Full Name"]),
    day,
    week: ymd(weekStart(day)),
    complete: String(r["Complete Time"] || ""),
    isNte: /NTE/.test(op),
  });
}

const exceptions = [];
const pairUnits = [];
const groups = new Map();
for (const row of rows.filter((x) => x.subcat === ACTIVE || x.subcat === PASSIVE)) {
  const key = row.siteId + "|" + row.week;
  if (!groups.has(key)) groups.set(key, []);
  groups.get(key).push(row);
}
for (const [, items] of groups) {
  const act = items.filter((x) => x.subcat === ACTIVE).sort((a, b) => a.complete.localeCompare(b.complete));
  const pas = items.filter((x) => x.subcat === PASSIVE).sort((a, b) => a.complete.localeCompare(b.complete));
  const site = items[0].siteId;
  const week = items[0].week;
  if (act.length && pas.length) {
    pairUnits.push({
      kind: "Active+Passive",
      siteId: site,
      weekStart: week,
      completeDay: pas[0].day,
      complete: pas[0].complete,
      nteFme: pas[0].fme,
      activeTask: act[0].taskId,
      passiveTask: pas[0].taskId,
    });
    act.slice(1).forEach((d) => exceptions.push({ reason: "Duplicate Active row same site + same week", siteId: site, day: ymd(d.day), week, subcat: d.subcat, taskId: d.taskId, fme: d.fme }));
    pas.slice(1).forEach((d) => exceptions.push({ reason: "Duplicate Passive row same site + same week", siteId: site, day: ymd(d.day), week, subcat: d.subcat, taskId: d.taskId, fme: d.fme }));
  } else {
    items.forEach((x) => exceptions.push({ reason: "Unpaired Active/Passive in the same week (ignored)", siteId: site, day: ymd(x.day), week, subcat: x.subcat, taskId: x.taskId, fme: x.fme }));
  }
}

pairUnits.sort((a, b) => a.complete.localeCompare(b.complete));
const counted = [];
const seenSite = new Set();
for (const u of pairUnits) {
  if (seenSite.has(u.siteId)) {
    exceptions.push({ reason: "Duplicate Active+Passive visit later week/site (ignored)", siteId: u.siteId, day: ymd(u.completeDay), week: u.weekStart, subcat: "Active+Passive pair", taskId: `${u.activeTask} / ${u.passiveTask}`, fme: u.nteFme });
    continue;
  }
  seenSite.add(u.siteId);
  counted.push(u);
}

const scBySite = new Map();
for (const row of rows.filter((x) => x.subcat === SMALL).sort((a, b) => a.complete.localeCompare(b.complete))) {
  if (!scBySite.has(row.siteId)) scBySite.set(row.siteId, []);
  scBySite.get(row.siteId).push(row);
}
for (const [site, items] of scBySite) {
  const keep = items[0];
  counted.push({
    kind: "Small Cell",
    siteId: site,
    weekStart: keep.week,
    completeDay: keep.day,
    complete: keep.complete,
    nteFme: keep.fme,
    activeTask: "",
    passiveTask: keep.taskId,
  });
  items.slice(1).forEach((d) => exceptions.push({ reason: "Duplicate Small Cell for same site (ignored)", siteId: site, day: ymd(d.day), week: d.week, subcat: d.subcat, taskId: d.taskId, fme: d.fme }));
}

const perfMap = new Map();
for (const c of counted) {
  if (c.completeDay.getUTCDay() === 5) continue;
  const fme = (c.nteFme || "").trim();
  const day = ymd(c.completeDay);
  const key = fme + "|" + day;
  if (!perfMap.has(key)) perfMap.set(key, { fme, day, ap: 0, sc: 0 });
  const p = perfMap.get(key);
  if (c.kind === "Active+Passive") p.ap += 1;
  else p.sc += 1;
}
const perf = [...perfMap.values()].map((p) => {
  const actual = p.ap + p.sc;
  const mode = p.ap > 0 ? "Active+Passive" : "Small Cell only";
  const target = p.ap > 0 ? 3 : 5;
  return { ...p, actual, mode, target, hit: actual >= target };
});

const days = [];
if (counted.length) {
  const min = counted.reduce((a, c) => (c.completeDay < a ? c.completeDay : a), counted[0].completeDay);
  const max = counted.reduce((a, c) => (c.completeDay > a ? c.completeDay : a), counted[0].completeDay);
  for (let t = min.getTime(); t <= max.getTime(); t += 86400000) {
    const d = new Date(t);
    if (d.getUTCDay() !== 5) days.push(ymd(d));
  }
}

const fmeNames = [...new Set(perf.map((p) => p.fme))].sort();
const fme = fmeNames.map((name) => {
  const daysW = perf.filter((p) => p.fme === name);
  const actual = daysW.reduce((s, p) => s + p.actual, 0);
  const expected = daysW.reduce((s, p) => s + p.target, 0);
  const ap = daysW.reduce((s, p) => s + p.ap, 0);
  const sc = daysW.reduce((s, p) => s + p.sc, 0);
  const on = daysW.filter((p) => p.hit).length;
  const below = daysW.filter((p) => !p.hit).length;
  const hitPct = daysW.length ? Math.round((100 * on) / daysW.length) : 0;
  const shape = hitPct >= 80 ? "GOOD" : hitPct >= 50 ? "WATCH" : "BAD";
  const daily = days.map((day) => {
    const row = daysW.find((p) => p.day === day);
    return row ? { day, actual: row.actual, target: row.target, hit: row.hit } : { day, actual: null, target: null, hit: null };
  });
  return {
    name, daysWorked: daysW.length, ap, sc, actual, expected,
    avg: Number((actual / Math.max(daysW.length, 1)).toFixed(1)),
    onTarget: on, below, hitPct, shape, daily,
  };
}).sort((a, b) => b.hitPct - a.hitPct || a.name.localeCompare(b.name));

const teamOn = perf.filter((p) => p.hit).length;
const teamHit = perf.length ? Math.round((100 * teamOn) / perf.length) : 0;
const teamShape = teamHit >= 80 ? "GOOD" : teamHit >= 50 ? "WATCH" : "BAD";
const payload = {
  generatedAt: new Date().toISOString().slice(0, 16).replace("T", " "),
  sourceFile: path.basename(src),
  owsUrl: process.env.OWS_URL || "https://106d-sg.teleows.com/",
  kpis: {
    tasks: counted.length,
    sites: new Set(counted.map((c) => c.siteId)).size,
    fmes: fmeNames.length,
    onTarget: teamOn,
    fmeDays: perf.length,
    hitPct: teamHit,
    shape: teamShape,
    below: perf.filter((p) => !p.hit).length,
    expected: perf.reduce((s, p) => s + p.target, 0),
    actual: perf.reduce((s, p) => s + p.actual, 0),
    pairs: counted.filter((c) => c.kind === "Active+Passive").length,
    smallCell: counted.filter((c) => c.kind === "Small Cell").length,
  },
  days,
  fme,
  below: perf.filter((p) => !p.hit).sort((a, b) => a.fme.localeCompare(b.fme) || a.day.localeCompare(b.day)).map((p) => ({
    fme: p.fme, day: p.day, mode: p.mode, actual: p.actual, target: p.target, ap: p.ap, sc: p.sc,
  })),
  exceptions,
  rules: [
    "PM only - OGK Active and Passive Routine Maintenance",
    "Active + Passive, same site, same week (Sun-Sat) = 1 task",
    "Small Cell = 1 task by itself",
    "NTE FME accounts only",
    "Mohd Yasin counted under Nabijohn (replacement)",
    "Daily target: 3 paired tasks, or 5 if Small Cell only",
    "Friday is not a working day",
  ],
};

fs.mkdirSync(outDir, { recursive: true });
const json = JSON.stringify(payload);
fs.writeFileSync(path.join(outDir, "data.json"), json);
let html = fs.readFileSync(templatePath, "utf8");
html = html.replace("__DASHBOARD_DATA__", json);
html = html.replace("__DASHBOARD_PIN__", process.env.DASHBOARD_PIN || "");
fs.writeFileSync(path.join(outDir, "index.html"), html);
const summary = `${teamShape} | ${payload.kpis.tasks} tasks | ${teamOn}/${perf.length} days on target (${teamHit}%)`;
fs.writeFileSync(path.join(outDir, "telegram.txt"), summary + "\n" + payload.below.slice(0, 12).map((b) => `${b.fme}: ${b.actual}/${b.target} on ${b.day}`).join("\n"));
console.log("Wrote dashboard", summary);
