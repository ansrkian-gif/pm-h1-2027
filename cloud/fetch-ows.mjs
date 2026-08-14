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
  return [page, ...page.frames()];
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

  let queryOpen = await firstVisible(page, (c) => c.getByText(/Query Task\s*\(?Ooredoo\)?/i));
  if (!queryOpen) {
    const taskMenu = await firstVisible(page, (c) => c.getByText(/^\s*Task\s*$/));
    if (taskMenu) {
      await taskMenu.click();
      await page.waitForTimeout(1500);
    }
    queryOpen = await firstVisible(page, (c) => c.getByText(/Query Task\s*\(?Ooredoo\)?/i));
    if (!queryOpen) throw new Error("Query Task(Ooredoo) not found at " + page.url());
    await queryOpen.click();
    await page.waitForTimeout(2500);
  }

  const typeTrigger = await firstVisible(page, (c) =>
    c.locator(".sdm_splitbutton_text").filter({ hasText: /^-Select-$|^PM$/ })
  );
  if (!typeTrigger) throw new Error("Task Type dropdown not visible");
  await typeTrigger.click();
  await page.waitForTimeout(500);
  const pm = await firstVisible(page, (c) => c.getByText(/^\s*PM\s*$/));
  if (!pm) throw new Error("PM option not visible");
  await pm.click();
  await page.waitForTimeout(400);

  const statusTrigger = await firstVisible(page, (c) =>
    c.getByText(/^-Select-$/).or(c.getByPlaceholder(/status/i)).or(c.getByLabel(/^Status$/i))
  );
  if (statusTrigger) await statusTrigger.click();
  await page.waitForTimeout(400);
  const closed = await firstVisible(page, (c) => c.getByText(/^\s*closed\s*$/i));
  if (closed) await closed.click();
  const completed = await firstVisible(page, (c) => c.getByText(/^\s*completed\s*$/i));
  if (completed) await completed.click();
  await page.keyboard.press("Escape").catch(() => {});

  const searchBox = await firstVisible(page, (c) =>
    c.getByPlaceholder(/Task Id, Title, Site, FME/i).or(c.getByLabel(/Task Id, Title, Site, FME/i))
  );
  if (!searchBox) throw new Error("Task Id, Title, Site, FME box not found");
  await searchBox.fill(searchText);

  await fillNearLabel(page, "Creation Time From", "2026-08-01 00:00:00");
  await fillNearLabel(page, "Creation Time To", "2027-01-31 23:59:59");
  await fillNearLabel(page, "Completion Time Form", fromDate);
  await fillNearLabel(page, "Completion Time From", fromDate);

  await shot(page, "03-filters.png");

  const searchBtn = await firstVisible(page, (c) => c.getByRole("button", { name: /^Search$/i }));
  if (!searchBtn) throw new Error("Search button not found");
  await searchBtn.click();
  await page.waitForTimeout(5000);
  await shot(page, "04-after-search.png");

  const exportBtn = await firstVisible(page, (c) => c.getByRole("button", { name: /^Export$/i }));
  if (!exportBtn) throw new Error("Export button not found");
  const [download] = await Promise.all([
    page.waitForEvent("download", { timeout: 180000 }),
    exportBtn.click(),
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
