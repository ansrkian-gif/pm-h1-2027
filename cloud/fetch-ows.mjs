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

async function ctxs(page) {
  return [page, ...page.frames()];
}

async function firstVisible(page, builder) {
  for (const ctx of await ctxs(page)) {
    const loc = builder(ctx);
    if (await loc.count()) return loc.first();
  }
  return null;
}

const browser = await chromium.launch({ headless: true });
const context = await browser.newContext({ acceptDownloads: true });
const page = await context.newPage();

try {
  await page.goto(homeUrl, { waitUntil: "domcontentloaded", timeout: 90000 });
  const passBox = page.locator('input[type="password"]').first();
  if (await passBox.count()) {
    await passBox.waitFor({ timeout: 20000 });
    await page.locator('input[type="text"], input[type="email"], input[name*="user" i], input[id*="user" i]').first().fill(user);
    await passBox.fill(pass);
    await page.locator('button[type="submit"], input[type="submit"], button:has-text("Login"), button:has-text("Sign in"), button:has-text("Log in")').first().click();
    await page.waitForTimeout(5000);
  }
  await shot(page, "01-after-login.png");

  if (!page.url().includes("Query")) {
    await page.goto(homeUrl, { waitUntil: "domcontentloaded", timeout: 60000 });
    await page.waitForTimeout(3000);
  }

  const queryOpen = await firstVisible(page, (c) => c.getByText(/Query Task\(Ooredoo\)/i));
  if (!queryOpen) {
    const taskMenu = await firstVisible(page, (c) => c.getByText(/^\s*Task\s*$/));
    if (taskMenu) {
      await taskMenu.click();
      await page.waitForTimeout(1000);
    }
    const q = await firstVisible(page, (c) => c.getByText(/Query Task\(Ooredoo\)/i));
    if (!q) throw new Error("Query Task(Ooredoo) not found");
    await q.click();
  }
  await page.waitForTimeout(2500);
  await shot(page, "02-query-page.png");

  const typeBox = await firstVisible(page, (c) => c.getByText(/^\s*PM\s*$/).or(c.getByLabel(/task type/i)));
  if (typeBox) {
    await typeBox.click();
    const pm = await firstVisible(page, (c) => c.getByText(/^\s*PM\s*$/));
    if (pm) await pm.click();
  }

  const status = await firstVisible(page, (c) => c.getByText(/-Select-/).or(c.getByLabel(/status/i)));
  if (status) await status.click();
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

  const fromBox = await firstVisible(page, (c) =>
    c.getByLabel(/Completion Time Form/i).or(c.getByLabel(/Completion Time From/i))
  );
  if (fromBox) {
    await fromBox.click();
    await fromBox.fill("");
    await fromBox.fill(fromDate);
    const ok = await firstVisible(page, (c) => c.getByRole("button", { name: /^OK$/i }));
    if (ok) await ok.click();
  }

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
  console.error(err.message);
  process.exit(1);
} finally {
  await browser.close();
}
