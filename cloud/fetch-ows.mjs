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
const fromRaw = (process.env.OWS_FROM || "2026-08-01 00:00:00").trim();
const fromDate = /\d{2}:\d{2}:\d{2}/.test(fromRaw) ? fromRaw : `${fromRaw} 00:00:00`;
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

async function clickListItem(page, want) {
  for (const f of page.frames()) {
    const hit = await f
      .evaluate((label) => {
        const nodes = [
          ...document.querySelectorAll(
            ".x-menu-item-text, .x-menu-item, .x-combo-list-item, .x-boundlist-item, .x-layer .x-combo-list div, .x-superboxselect-item"
          ),
        ];
        const el = nodes.find((n) => {
          const t = (n.innerText || "").trim();
          if (t !== label) return false;
          const r = n.getBoundingClientRect();
          return r.width >= 8 && r.height >= 8 && r.width < 320 && r.height < 48;
        });
        if (!el) return "";
        el.click();
        return String(el.className || "ok").slice(0, 80);
      }, want)
      .catch(() => "");
    if (hit) {
      console.log("list item", want, hit);
      return true;
    }
  }
  return false;
}

async function hideExtPopups(target) {
  await target
    .evaluate(() => {
      try {
        if (window.Ext && Ext.menu && Ext.menu.MenuMgr) Ext.menu.MenuMgr.hideAll();
      } catch {}
      try {
        if (window.Ext && Ext.WindowMgr) Ext.WindowMgr.hideAll();
      } catch {}
      document.querySelectorAll(".x-date-picker, .x-datetime-picker, .x-combo-list").forEach((el) => {
        el.style.display = "none";
        el.style.visibility = "hidden";
      });
      document.querySelectorAll(".x-layer").forEach((el) => {
        const t = el.innerText || "";
        if (/Current Page/.test(t)) return;
        if (/Clear|Update|NPM|\bPM\b|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec/.test(t)) {
          el.style.display = "none";
          el.style.visibility = "hidden";
        }
      });
    })
    .catch(() => {});
}

async function closeOverlays(page, ows) {
  if (ows) await hideExtPopups(ows);
  for (const f of page.frames()) await hideExtPopups(f);
  for (let i = 0; i < 2; i++) {
    await page.keyboard.press("Escape").catch(() => {});
    await page.waitForTimeout(120);
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
      }, pattern)
      .catch(() => "");
    if (clicked) {
      console.log("clicked", clicked);
      return true;
    }
  }
  return false;
}

async function dumpLayers(page, label) {
  for (const f of page.frames()) {
    const layers = await f
      .evaluate(() =>
        [...document.querySelectorAll(".x-menu, .x-layer, .x-combo-list, .x-date-picker, .sdm_splitbutton")]
          .filter((el) => el.offsetWidth > 0 && el.offsetHeight > 0)
          .slice(0, 12)
          .map((el) => ({
            cls: String(el.className || "").slice(0, 70),
            t: (el.innerText || "").replace(/\s+/g, " ").slice(0, 70),
            w: el.offsetWidth,
            h: el.offsetHeight,
          }))
      )
      .catch(() => []);
    if (layers && layers.length) console.log(label, f.url().slice(-60), JSON.stringify(layers));
  }
}

async function clickExportAll(page) {
  for (const f of page.frames()) {
    const hit = await f
      .evaluate(() => {
        const nodes = [
          ...document.querySelectorAll(".x-menu-item-text, .x-menu-item, .x-menu span, .x-menu a, .x-menu li, .x-layer span"),
        ];
        const el = nodes.find((n) => {
          const t = (n.innerText || "").trim();
          if (t !== "All") return false;
          const r = n.getBoundingClientRect();
          return r.width >= 8 && r.height >= 8 && r.width < 320 && r.height < 48;
        });
        if (!el) return "";
        el.click();
        return String(el.className || "ok").slice(0, 80);
      })
      .catch(() => "");
    if (hit) {
      console.log("export all", hit);
      return true;
    }
  }
  for (const f of page.frames()) {
    const loc = f.getByText(/^All$/, { exact: true });
    const n = await loc.count().catch(() => 0);
    for (let i = 0; i < n; i++) {
      const el = loc.nth(i);
      if (await el.isVisible().catch(() => false)) {
        await el.click();
        console.log("export all locator");
        return true;
      }
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
page.on("dialog", (d) => d.accept().catch(() => {}));

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
  await closeOverlays(page, ows);

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
    return "clicked:" + (input.className || "") + ":" + (input.value || "");
  });
  console.log("task type", typeClick);
  await page.waitForTimeout(400);
  let pmPick = false;
  for (let i = 0; i < 8 && !pmPick; i++) {
    pmPick = await clickListItem(page, "PM");
    if (!pmPick) await page.waitForTimeout(400);
  }
  if (!pmPick) {
    pmPick = await owsEval(() => {
      const items = [...document.querySelectorAll(".x-combo-list-item, .x-boundlist-item, .x-combo-list div, .x-layer div")];
      const pm = items.find((el) => (el.innerText || "").trim() === "PM");
      if (!pm) return "no-pm:" + items.slice(0, 12).map((e) => (e.innerText || "").trim()).filter(Boolean).join("|");
      pm.click();
      return "ok";
    });
    console.log("pm pick", pmPick);
  }
  if (!pmPick || String(pmPick).startsWith("no-")) {
    await page.keyboard.type("PM", { delay: 40 });
    await page.keyboard.press("Enter");
  }
  await closeOverlays(page, ows);

  const statusOpen = await owsEval(() => {
    const boxes = [...document.querySelectorAll("input.x-superboxselect-input-field, input.x-form-empty-field")];
    const status = boxes.find((i) => {
      const row = i.closest(".toolbar_each");
      const t = ((row && row.innerText) || "") + " " + (i.value || "") + " " + (i.placeholder || "");
      return /Status|Select/i.test(t) && !/Task Type/i.test((row && row.innerText) || "");
    }) || boxes[1];
    if (!status) return "no-status";
    status.click();
    return "clicked:" + ((status.closest(".toolbar_each") || {}).innerText || "").slice(0, 40);
  });
  console.log("status", statusOpen);
  await page.waitForTimeout(500);
  await clickListItem(page, "closed");
  await page.waitForTimeout(250);
  if (statusOpen && statusOpen !== "no-status") {
    await owsEval(() => {
      const boxes = [...document.querySelectorAll("input.x-superboxselect-input-field")];
      const status = boxes.find((i) => {
        const row = i.closest(".toolbar_each");
        return row && !/Task Type/i.test(row.innerText || "");
      }) || boxes[1];
      if (status) status.click();
    });
    await page.waitForTimeout(300);
  }
  await clickListItem(page, "completed");
  await closeOverlays(page, ows);

  const searchFill = await owsEval((text) => {
    const inputs = [...document.querySelectorAll(".toolbar_each input")];
    const box =
      inputs.find((i) => /Task Id|Title|Site|FME/i.test(i.placeholder || i.title || "")) ||
      inputs.find((i) => {
        const row = i.closest(".toolbar_each");
        const t = (row && row.innerText) || "";
        return !/Task Type|Creation|Completion|Search|Export|Refresh|Select/i.test(t);
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

  async function owsFillDates(label, fromVal, toVal) {
    const res = await owsEval(({ label, fromVal, toVal }) => {
      const row = [...document.querySelectorAll(".toolbar_each")].find((el) =>
        (el.innerText || "").includes(label)
      );
      if (!row) return "no-row";
      const inputs = [...row.querySelectorAll("input")];
      if (!inputs.length) return "no-input";
      const setVal = (input, value) => {
        input.value = value;
        input.dispatchEvent(new Event("input", { bubbles: true }));
        input.dispatchEvent(new Event("change", { bubbles: true }));
        input.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true }));
        input.dispatchEvent(new Event("blur", { bubbles: true }));
        try {
          if (window.Ext && input.id) {
            const cmp = Ext.getCmp(input.id);
            if (cmp && cmp.setRawValue) cmp.setRawValue(value);
            else if (cmp && cmp.setValue) cmp.setValue(value);
          }
        } catch {}
      };
      setVal(inputs[0], fromVal);
      if (toVal && inputs[1]) setVal(inputs[1], toVal);
      return "ok:" + inputs.length + ":" + inputs.map((i) => i.value).join("|");
    }, { label, fromVal, toVal });
    console.log("date", label, res);
    await closeOverlays(page, ows);
  }
  await owsFillDates("Creation Time From", "2026-08-01 00:00:00", "2027-01-31 23:59:59");
  await owsFillDates("Completion Time Form", fromDate, "");
  await closeOverlays(page, ows);

  await shot(page, "03-filters.png");

  const searched = await owsEval(() => {
    const row = [...document.querySelectorAll(".toolbar_each")].find((el) => (el.innerText || "").trim() === "Search");
    const btn = row && row.querySelector(".sdm_splitbutton_text, button, .sdm_button, .toolbar_each_input");
    if (!btn) return "no-search-btn";
    btn.click();
    return "ok";
  });
  console.log("search click", searched);
  if (searched !== "ok") throw new Error("Search button not found");
  await page.waitForTimeout(10000);
  await closeOverlays(page, ows);
  await shot(page, "04-after-search.png");

  const blobs = [];
  const onRes = async (res) => {
    try {
      const headers = res.headers();
      const cd = headers["content-disposition"] || "";
      const ct = headers["content-type"] || "";
      const url = res.url();
      if (/\.(js|css|png|gif|jpg|woff|svg)(\?|$)/i.test(url) || /javascript|css|image|font/.test(ct)) return;
      if (cd || /xlsx|excel|octet-stream|spreadsheet|export|download/i.test(ct + url + cd)) {
        console.log("net-file", res.status(), ct.slice(0, 60), cd.slice(0, 80), url.slice(0, 180));
      } else if (!/\/(extjs|static|assets|adc-web\/ui)\//i.test(url)) {
        console.log("net", res.status(), ct.slice(0, 40), url.slice(0, 160));
      }
      const buf = await res.body().catch(() => null);
      if (!buf || buf.length < 1000) return;
      const head = Buffer.from(buf.slice(0, 4)).toString();
      if (
        head.startsWith("PK") ||
        /\.xlsx/i.test(cd + url) ||
        /spreadsheetml|officedocument|ms-excel/i.test(ct) ||
        (/octet-stream/i.test(ct) && /export|download|xlsx/i.test(url + cd))
      ) {
        blobs.push(buf);
        console.log("captured blob", buf.length, ct, cd.slice(0, 80));
      }
    } catch {}
  };
  context.on("response", onRes);

  const downloadPromise = page.waitForEvent("download", { timeout: 180000 }).catch(() => null);
  const popupPromise = page.waitForEvent("popup", { timeout: 25000 }).catch(() => null);

  await closeOverlays(page, ows);
  await dumpLayers(page, "layers pre-export");
  const exported = await owsEval(() => {
    const row = [...document.querySelectorAll(".toolbar_each")].find((el) => (el.innerText || "").trim() === "Export");
    if (!row) return "no-export";
    const input = row.querySelector(".toolbar_each_input");
    const btn = row.querySelector(".sdm_splitbutton_text, button, .sdm_button");
    if (input) input.click();
    else if (btn) btn.click();
    else return "no-btn";
    return "ok";
  });
  console.log("export click", exported);
  if (exported !== "ok") throw new Error("Export button not found");
  await page.waitForTimeout(800);
  await dumpLayers(page, "layers post-export");
  await shot(page, "05-export-menu.png");
  let allClicked = await clickExportAll(page);
  if (!allClicked) {
    await owsEval(() => {
      const row = [...document.querySelectorAll(".toolbar_each")].find((el) => (el.innerText || "").trim() === "Export");
      const arrow = row && row.querySelector(".sdm_splitbutton_arrow, .x-btn-split-right, em, .sdm_splitbutton i");
      if (arrow) arrow.click();
    });
    await page.waitForTimeout(600);
    await dumpLayers(page, "layers post-arrow");
    allClicked = await clickExportAll(page);
  }
  console.log("export all clicked", allClicked);
  await shot(page, "06-after-export-all.png");

  let download = await downloadPromise;
  if (!download) {
    const popup = await popupPromise;
    if (popup) download = await popup.waitForEvent("download", { timeout: 120000 }).catch(() => null);
  }
  if (download) {
    await download.saveAs(outFile);
    console.log("Saved download", outFile, fs.statSync(outFile).size, download.suggestedFilename());
  } else if (blobs.length) {
    fs.writeFileSync(outFile, blobs[blobs.length - 1]);
    console.log("Saved response body", outFile, fs.statSync(outFile).size);
  }
  context.off("response", onRes);
  if (!fs.existsSync(outFile) || fs.statSync(outFile).size < 1000) {
    throw new Error("Export did not produce a valid xlsx");
  }
} catch (err) {
  await shot(page, "last-ows.png");
  await debugDump(page, "last-ows.txt");
  console.error(err.message);
  process.exit(1);
} finally {
  await browser.close();
}
