import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { chromium } from "playwright";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outFile = path.join(__dirname, "extract.xlsx");
const base = (process.env.OWS_URL || "https://106d-sg.teleows.com/").replace(/\/$/, "");
const queryHash = "#/WFMBase_custom/mission_control_service_separate/Query Task Ooredoo";
const homeUrl = `${base}/portal-web/portal/homepage.html${queryHash}`;
const user = process.env.OWS_USER;
const pass = process.env.OWS_PASS;
const fromDate = process.env.OWS_FROM || "2026-08-01 00:00:00";
const searchText = process.env.OWS_SEARCH || "_OGK Active and Passive";

if (!user || !pass) {
  console.error("OWS_USER and OWS_PASS must be set as cloud secrets.");
  process.exit(1);
}

async function shot(page, name) {
  await page.screenshot({ path: path.join(__dirname, name), fullPage: true }).catch(() => {});
}

async function debugDump(page, name) {
  const html = await page.content().catch(() => "");
  const safe = html
    .replace(/type=["']password["'][^>]*>/gi, 'type="password">')
    .replace(new RegExp(pass.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g"), "***");
  const text = [
    `url=${page.url()}`,
    `title=${await page.title().catch(() => "")}`,
    `htmlLength=${html.length}`,
    safe.slice(0, 80000),
  ].join("\n");
  fs.writeFileSync(path.join(__dirname, name), text);
}

async function ctxs(page) {
  const frames = [...page.frames()];
  const rank = (f) => {
    const u = f.url();
    if (/mission_task_work_grid_ooredoo/i.test(u)) return 0;
    if (/c_mission_control_service|adc-ui\/spl/i.test(u)) return 1;
    if (/adc-web\/ui/i.test(u)) return 2;
    return 3;
  };
  frames.sort((a, b) => rank(a) - rank(b));
  return frames;
}

async function firstVisible(page, builder) {
  for (const ctx of await ctxs(page)) {
    const loc = builder(ctx);
    const n = await loc.count();
    for (let i = 0; i < n; i++) {
      const el = loc.nth(i);
      if (await el.isVisible().catch(() => false)) return el;
    }
  }
  return null;
}

async function logControls(page) {
  for (const f of page.frames()) {
    const info = await f
      .evaluate(() => {
        const out = [];
        document.querySelectorAll("span, button, input, label, div").forEach((el) => {
          const t = (el.innerText || el.value || "").trim().replace(/\s+/g, " ").slice(0, 48);
          if (!t || t.length > 40) return;
          const rec = el.getBoundingClientRect();
          if (rec.width < 8 || rec.height < 8) return;
          if (!/(Select|PM|Task Type|Status|Search|Export|Completion|Creation|closed|completed|Form|From)/i.test(t)) return;
          out.push({ t, cls: String(el.className || "").slice(0, 60), w: Math.round(rec.width), h: Math.round(rec.height) });
        });
        return out.slice(0, 50);
      })
      .catch(() => []);
    if (info && info.length) console.log("controls", f.url(), JSON.stringify(info));
  }
}

async function clickMatching(page, pattern) {
  for (const f of page.frames()) {
    const clicked = await f
      .evaluate((pat) => {
        const r = new RegExp(pat);
        const els = [...document.querySelectorAll("span, div, button, a, label, li")];
        for (const el of els) {
          const t = (el.innerText || "").trim().replace(/\s+/g, " ");
          if (t.length > 24) continue;
          if (!r.test(t)) continue;
          const rec = el.getBoundingClientRect();
          if (rec.width < 8 || rec.height < 8) continue;
          el.click();
          return t;
        }
        return "";
      })
      .catch(() => "");
    if (clicked) {
      console.log("clicked", clicked);
      return true;
    }
  }
  return false;
}

async function fillNearLabel(page, label, value) {
  const lab = await firstVisible(page, (c) => c.getByText(new RegExp("^\\s*" + label + "\\s*$", "i")));
  if (!lab) {
    console.log("label missing", label);
    return false;
  }
  const input = lab.locator("xpath=following::input[1]");
  await input.click({ timeout: 8000 });
  await input.fill("");
  await input.fill(value);
  const ok = await firstVisible(page, (c) => c.getByRole("button", { name: /^OK$/i }));
  if (ok) await ok.click();
  await page.waitForTimeout(300);
  return true;
}

async function loginIfNeeded(page) {
  const passBox = await firstVisible(page, (c) => c.locator('input[type="password"]'));
  if (!passBox) {
    console.log("No password box yet at", page.url());
    return false;
  }
  const userBox = await firstVisible(page, (c) =>
    c.locator('input[type="text"], input[type="email"], input[name*="user" i], input[id*="user" i], #username')
  );
  if (!userBox) throw new Error("Login user box not found");
  await userBox.fill(user);
  await passBox.fill(pass);
  const btn = await firstVisible(page, (c) =>
    c.locator('button[type="submit"], input[type="submit"], button:has-text("Login"), button:has-text("Sign in"), button:has-text("Log in")')
  );
  if (btn) await btn.click();
  else await page.keyboard.press("Enter");
  await page.waitForTimeout(8000);
  await page.waitForLoadState("networkidle").catch(() => {});
  console.log("After login", page.url());
  return true;
}

const browser = await chromium.launch({ headless: true });
const context = await browser.newContext({
  acceptDownloads: true,
  viewport: { width: 1440, height: 900 },
  userAgent:
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
});
const page = await context.newPage();

try {
  const resp = await page.goto(homeUrl, { waitUntil: "load", timeout: 120000 });
  console.log("goto", resp?.status(), page.url());
  await page.waitForTimeout(4000);
  await loginIfNeeded(page);
  await shot(page, "01-after-login.png");
  await debugDump(page, "01-after-login.txt");

  if (!/homepage|portal|Query/i.test(page.url())) {
    await page.goto(homeUrl, { waitUntil: "load", timeout: 90000 });
    await page.waitForTimeout(5000);
    await loginIfNeeded(page);
  } else if (!page.url().includes("Query")) {
    await page.goto(homeUrl, { waitUntil: "load", timeout: 90000 });
    await page.waitForTimeout(5000);
  }

  await shot(page, "02-query-page.png");
  await debugDump(page, "02-query-page.txt");
  await page.waitForTimeout(2000);
  await logControls(page);

  let ows = page.frames().find((f) => /mission_task_work_grid_ooredoo/i.test(f.url()));
  if (!ows) {
    const queryOpen = await firstVisible(page, (c) => c.getByText(/Query Task\s*\(?Ooredoo\)?/i));
    if (queryOpen) {
      await queryOpen.click();
      await page.waitForTimeout(2500);
    }
    ows = page.frames().find((f) => /mission_task_work_grid_ooredoo/i.test(f.url()));
  }
  if (!ows) throw new Error("Ooredoo query frame not found");
  console.log("ows frame", ows.url());
  await page.keyboard.press("Escape").catch(() => {});

  async function owsEval(fn, arg) {
    return ows.evaluate(fn, arg);
  }

  const typeClick = await owsEval(() => {
    const row = [...document.querySelectorAll(".toolbar_each")].find((el) =>
      (el.innerText || "").includes("Task Type")
    );
    const input = row && row.querySelector("input");
    if (!input) return "no-input";
    input.click();
    return "clicked:" + (input.className || "");
  });
  console.log("task type", typeClick);
  await page.waitForTimeout(600);
  const pmPick = await owsEval(() => {
    const items = [...document.querySelectorAll(".x-combo-list-item, .x-boundlist-item, .x-combo-list div, .x-layer div")];
    const pm = items.find((el) => (el.innerText || "").trim() === "PM");
    if (!pm) return "no-pm:" + items.slice(0, 12).map((e) => (e.innerText || "").trim()).filter(Boolean).join("|");
    pm.click();
    return "ok";
  });
  console.log("pm pick", pmPick);
  if (!String(pmPick).startsWith("ok")) {
    await page.keyboard.type("PM", { delay: 40 });
    await page.keyboard.press("Enter");
  }
  await page.keyboard.press("Escape").catch(() => {});
  await page.waitForTimeout(300);

  const searchFill = await owsEval((text) => {
    const inputs = [...document.querySelectorAll(".toolbar_each input")];
    const box =
      inputs.find((i) => /Task Id|Title|Site|FME/i.test(i.placeholder || i.title || "")) ||
      inputs.find((i) => {
        const row = i.closest(".toolbar_each");
        const t = (row && row.innerText) || "";
        return !/Task Type|Creation|Completion|Search|Export|Refresh/i.test(t);
      });
    if (!box) return "no-search";
    box.focus();
    box.value = text;
    box.dispatchEvent(new Event("input", { bubbles: true }));
    box.dispatchEvent(new Event("change", { bubbles: true }));
    return "ok:" + (box.placeholder || box.className);
  }, searchText);
  console.log("search", searchFill);
  if (!String(searchFill).startsWith("ok")) throw new Error("Task Id, Title, Site, FME box not found");

  async function owsFillDate(label, value) {
    const res = await owsEval(({ label, value }) => {
      const row = [...document.querySelectorAll(".toolbar_each")].find((el) =>
        (el.innerText || "").includes(label)
      );
      if (!row) return "no-row";
      const input = row.querySelector("input");
      if (!input) return "no-input:" + (row.className || "");
      input.focus();
      input.click();
      input.value = value;
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(new Event("change", { bubbles: true }));
      input.blur();
      return "ok:" + input.className + ":" + input.value;
    }, { label, value });
    console.log("date", label, res);
    await page.keyboard.press("Enter").catch(() => {});
    await page.waitForTimeout(200);
  }
  await owsFillDate("Creation Time From", "2026-08-01 00:00:00");
  await owsFillDate("Creation Time To", "2027-01-31 23:59:59");
  await owsFillDate("Completion Time Form", fromDate);
  await page.keyboard.press("Escape").catch(() => {});

  await shot(page, "03-filters.png");

  const searched = await owsEval(() => {
    const row = [...document.querySelectorAll(".toolbar_each")].find((el) => (el.innerText || "").trim() === "Search");
    const btn = row && row.querySelector(".sdm_splitbutton_text, button, .sdm_button");
    if (!btn) return "no-search-btn";
    btn.click();
    return "ok";
  });
  console.log("search click", searched);
  if (searched !== "ok") throw new Error("Search button not found");
  await page.waitForTimeout(8000);
  await shot(page, "04-after-search.png");

  const exportWait = page.waitForEvent("download", { timeout: 180000 });
  const exported = await owsEval(() => {
    const row = [...document.querySelectorAll(".toolbar_each")].find((el) => (el.innerText || "").trim() === "Export");
    const btn = row && row.querySelector(".sdm_splitbutton_text, button, .sdm_button");
    if (!btn) return "no-export";
    btn.click();
    return "ok";
  });
  console.log("export click", exported);
  if (exported !== "ok") throw new Error("Export button not found");
  const download = await exportWait;
  await download.saveAs(outFile);
  if (!fs.existsSync(outFile) || fs.statSync(outFile).size < 1000) {
    throw new Error("Export did not produce a valid xlsx");
  }
  console.log("Saved", outFile, fs.statSync(outFile).size, download.suggestedFilename());
} catch (err) {
  await shot(page, "last-ows.png");
  await debugDump(page, "last-ows.txt");
  console.error(err.message);
  process.exit(1);
} finally {
  await browser.close();
}
