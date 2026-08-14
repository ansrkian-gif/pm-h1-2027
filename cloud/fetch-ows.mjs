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

  const typeInput = ows.locator("input.x-superboxselect-input-field").first();
  await typeInput.click({ timeout: 15000 });
  await typeInput.fill("PM");
  await page.waitForTimeout(500);
  const pmItem = ows.locator(".x-combo-list-item, .x-boundlist-item").filter({ hasText: /^PM$/ }).first();
  if (await pmItem.count()) await pmItem.click({ force: true });
  else await page.keyboard.press("Enter");
  await page.waitForTimeout(400);

  const superBoxes = ows.locator("input.x-superboxselect-input-field");
  if ((await superBoxes.count()) > 1) {
    const statusInput = superBoxes.nth(1);
    await statusInput.click();
    await statusInput.fill("closed");
    await page.keyboard.press("Enter");
    await page.waitForTimeout(200);
    await statusInput.fill("completed");
    await page.keyboard.press("Enter");
    await page.keyboard.press("Escape");
  }

  const searchBox = ows.getByPlaceholder(/Task Id, Title, Site, FME/i);
  if (await searchBox.count()) {
    await searchBox.fill(searchText);
  } else {
    const row = ows
      .locator(".toolbar_each")
      .filter({ hasNotText: /Task Type|Creation|Completion|Search|Export|Refresh/i })
      .locator("input.x-form-field")
      .first();
    if (!(await row.count())) throw new Error("Task Id, Title, Site, FME box not found");
    await row.fill(searchText);
  }

  async function fillToolbarDate(label, value) {
    const row = ows.locator(".toolbar_each").filter({ hasText: label }).first();
    if (!(await row.count())) {
      console.log("date missing", label);
      return;
    }
    const inp = row.locator("input.x-form-field").first();
    await inp.click();
    await inp.fill(value);
    await page.keyboard.press("Enter");
    const ok = ows.getByRole("button", { name: /^OK$/i });
    if (await ok.count()) await ok.first().click().catch(() => {});
  }
  await fillToolbarDate("Creation Time From", "2026-08-01 00:00:00");
  await fillToolbarDate("Creation Time To", "2027-01-31 23:59:59");
  await fillToolbarDate("Completion Time Form", fromDate);

  await shot(page, "03-filters.png");

  await ows.locator(".toolbar_each").filter({ hasText: /^Search$/ }).locator(".sdm_splitbutton_text").first().click();
  await page.waitForTimeout(6000);
  await shot(page, "04-after-search.png");

  const [download] = await Promise.all([
    page.waitForEvent("download", { timeout: 180000 }),
    ows.locator(".toolbar_each").filter({ hasText: /^Export$/ }).locator(".sdm_splitbutton_text").first().click(),
  ]);
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
