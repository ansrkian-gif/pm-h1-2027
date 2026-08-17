"""
Daily 17:30 alert: today's WO target vs closed, and which FMEs are behind.
Sends a mobile push via ntfy.sh (install free ntfy app on Android).
"""
from __future__ import annotations

import json
import os
import secrets
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DASH = ROOT / "dashboard"
CONFIG = ROOT / "alert_config.json"
EXTRACT = ROOT / "scripts" / "extract_data.py"
DATA_JS = DASH / "dashboard-data.js"
KUWAIT = timezone(timedelta(hours=3))


def kuwait_today() -> date:
    return datetime.now(KUWAIT).date()


FRIDAY = 4  # Python weekday(): Mon=0 ... Fri=4


def load_config() -> dict:
    if CONFIG.exists():
        return json.loads(CONFIG.read_text(encoding="utf-8"))
    return {}


def save_config(cfg: dict) -> None:
    CONFIG.write_text(json.dumps(cfg, indent=2), encoding="utf-8")


def load_dashboard_data() -> dict:
    if not DATA_JS.exists():
        raise SystemExit("Missing dashboard-data.js — run extract_data.py first")
    text = DATA_JS.read_text(encoding="utf-8").strip()
    if text.startswith("window.DASHBOARD_DATA"):
        text = text.split("=", 1)[1].strip()
    if text.endswith(";"):
        text = text[:-1]
    return json.loads(text)


def refresh_data() -> None:
    if EXTRACT.exists():
        subprocess.run([sys.executable, str(EXTRACT)], cwd=str(ROOT), check=False)


def build_summary(today: date | None = None) -> tuple[str, str, dict]:
    """Returns title, body, details dict."""
    data = load_dashboard_data()
    targets = data.get("targets", {})
    team_target = int(targets.get("dailyTeam", 21))
    per_fme = int(targets.get("perFmeDaily", 3))
    today = today or kuwait_today()
    today_key = today.isoformat()
    day_name = today.strftime("%a %d %b")

    if today.weekday() == FRIDAY:
        title = f"PM NTE · {day_name}"
        body = "Friday — rest day (no WO target)."
        details = {
            "id": f"fri-{today_key}-{int(datetime.now().timestamp())}",
            "friday": True,
            "closed": 0,
            "target": 0,
            "achieved": True,
            "behind": [],
            "on_track": [],
            "date": today_key,
            "dayName": day_name,
            "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "incompletePairCount": int(((data.get("siteChecks") or {}).get("generalPairs") or {}).get("incompleteSiteCount") or 0),
            "smallCellSiteCount": int(((data.get("siteChecks") or {}).get("smallCell") or {}).get("siteCount") or 0),
            "smallCellWoCount": int(((data.get("siteChecks") or {}).get("smallCell") or {}).get("woCount") or 0),
        }
        return title, body, details

    wos = [w for w in data.get("workOrders", []) if w.get("completeDate") == today_key]
    closed = len(wos)
    by_fme = Counter(w.get("fmeShort") or "?" for w in wos)

    # All known FME seats from full dataset (so missing ones show as 0 today)
    all_fmes = sorted({w.get("fmeShort") for w in data.get("workOrders", []) if w.get("fmeShort")})
    if not all_fmes:
        all_fmes = sorted(by_fme.keys())

    behind = []
    on_track = []
    for name in all_fmes:
        count = by_fme.get(name, 0)
        gap = count - per_fme
        if gap < 0:
            behind.append(f"{name} {count}/{per_fme} ({gap})")
        else:
            on_track.append(f"{name} {count}/{per_fme}")

    achieved = closed >= team_target
    status = "TARGET MET" if achieved else "BEHIND TARGET"
    title = f"PM NTE · {status}"

    lines = [
        f"{day_name}",
        f"Today: {closed}/{team_target} WOs",
        ("Result: Achieved" if achieved else f"Result: Short by {team_target - closed}"),
        "",
    ]
    if behind:
        lines.append("FMEs behind:")
        lines.extend(f"- {x}" for x in behind)
    else:
        lines.append("All FMEs met daily target.")

    if on_track:
        lines.append("")
        lines.append("On track: " + ", ".join(on_track))

    checks = data.get("siteChecks") or {}
    gp = checks.get("generalPairs") or {}
    sc = checks.get("smallCell") or {}
    incomplete_n = int(gp.get("incompleteSiteCount") or 0)
    small_n = int(sc.get("siteCount") or 0)
    small_wo = int(sc.get("woCount") or 0)
    dup_n = int(sc.get("duplicateCount") or 0)
    incomplete_sites = gp.get("incompleteSites") or []

    lines.append("")
    lines.append("Site integrity (full report):")
    lines.append(f"- Incomplete Active/Passive pairs: {incomplete_n}")
    if incomplete_n:
        sample = ", ".join(
            f"{row.get('siteId')} miss {'/'.join(row.get('missing') or [])}"
            for row in incomplete_sites[:6]
        )
        lines.append(f"  {sample}")
        if incomplete_n > 6:
            lines.append(f"  … +{incomplete_n - 6} more")
    lines.append(f"- Small Cell sites: {small_n} (WOs {small_wo})")
    if dup_n:
        lines.append(f"- Small Cell sites with >1 entry: {dup_n}")

    ex = checks.get("exemptions") or {}
    exempt_wo = int(ex.get("matchedWoCount") or 0)
    exempt_sites = int(ex.get("commonSiteIdCount") or 0)
    lines.append(f"- Exempted WOs: {exempt_wo}")
    lines.append(f"- Exempted common Site IDs: {exempt_sites}")

    body = "\n".join(lines)
    details = {
        "id": f"{today_key}-{closed}-{len(behind)}-{incomplete_n}-{exempt_sites}-{int(datetime.now().timestamp())}",
        "friday": False,
        "closed": closed,
        "target": team_target,
        "achieved": achieved,
        "behind": behind,
        "on_track": on_track,
        "date": today_key,
        "dayName": day_name,
        "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "incompletePairCount": incomplete_n,
        "missingPassiveCount": int(gp.get("missingPassiveCount") or 0),
        "missingActiveCount": int(gp.get("missingActiveCount") or 0),
        "incompletePairSites": [
            {
                "siteId": row.get("siteId"),
                "missing": row.get("missing") or [],
            }
            for row in incomplete_sites[:20]
        ],
        "smallCellSiteCount": small_n,
        "smallCellWoCount": small_wo,
        "smallCellDuplicateCount": dup_n,
        "pairCompleteCount": int(gp.get("completeSiteCount") or 0),
        "exemptedWoCount": exempt_wo,
        "exemptedSiteIdCount": exempt_sites,
    }
    return title, body, details


def site_base_url() -> str | None:
    env = (os.environ.get("PAGES_URL") or "").rstrip("/")
    if env:
        return env
    return "https://ansrkian-gif.github.io/pm-h1-2027"


def write_alert_data(details: dict) -> Path:
    """Write alert-data.js used by the full-screen mobile alert page."""
    DASH.mkdir(parents=True, exist_ok=True)
    out = DASH / "alert-data.js"
    out.write_text(
        "window.ALERT_DATA = " + json.dumps(details, ensure_ascii=False, indent=2) + ";\n",
        encoding="utf-8",
    )
    return out


def alert_page_url() -> str | None:
    base = site_base_url()
    return f"{base}/alert.html" if base else None


def send_ntfy(topic: str, title: str, body: str, achieved: bool | None, click_url: str | None = None) -> None:
    # Max priority so Android heads-up appears; tap opens full-screen page
    if achieved is True:
        priority = "5"
        tags = "white_check_mark,iphone"
    elif achieved is False:
        priority = "5"
        tags = "warning,iphone"
    else:
        priority = "4"
        tags = "calendar,iphone"

    click = click_url or alert_page_url()
    short_body = "TAP THIS NOTIFICATION\n→ opens BIG full-screen alert\n→ stays until you press CLEAR"
    if body:
        short_body = short_body + "\n\n" + body[:500]

    headers = {
        "Title": title[:250],
        "Priority": priority,
        "Tags": tags,
        "Content-Type": "text/plain; charset=utf-8",
    }
    if click:
        headers["Click"] = click
        headers["Actions"] = f"view, Open full alert, {click}, clear=true"

    base = (os.environ.get("NTFY_URL") or "https://ntfy.sh").rstrip("/")
    url = f"{base}/{urllib.parse.quote(topic)}"
    token = os.environ.get("NTFY_TOKEN") or ""
    if token:
        headers["Authorization"] = f"Bearer {token}"
    req = urllib.request.Request(
        url,
        data=short_body.encode("utf-8"),
        method="POST",
        headers=headers,
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            resp.read()
    except urllib.error.HTTPError as exc:
        raise SystemExit(f"ntfy send failed: {exc.code} {exc.read().decode('utf-8', errors='ignore')}") from exc
    except urllib.error.URLError as exc:
        raise SystemExit(f"ntfy send failed (network): {exc}") from exc


def ensure_config(interactive: bool = False) -> dict:
    cfg = load_config()
    if cfg.get("topic"):
        return cfg
    if not interactive:
        raise SystemExit("Alert not set up. Run: Setup Daily Alert.bat")

    topic = f"pm-nte-{secrets.token_hex(4)}"
    cfg = {
        "topic": topic,
        "time": "17:30",
        "createdAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    save_config(cfg)
    return cfg


def create_windows_task(time_hhmm: str = "17:30") -> None:
    """Prefer no-admin methods. 17:30 alert is also handled by auto_sync.py."""
    runner = ROOT / "run_daily_alert.bat"
    # Try current-user scheduled task (often allowed without admin)
    cmd = [
        "schtasks",
        "/Create",
        "/TN",
        "PM NTE Daily Alert",
        "/TR",
        f'cmd /c ""{runner}""',
        "/SC",
        "DAILY",
        "/ST",
        time_hhmm,
        "/F",
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode == 0:
        print("Windows task created: PM NTE Daily Alert @ " + time_hhmm + " daily")
        return

    print("Scheduled task not available on this laptop (no admin) — OK.")
    print("Use Install Auto Sync.bat instead: it sends the 5:30 PM alert")
    print("automatically while Auto Sync is running (no admin needed).")



def write_phone_instructions(topic: str) -> None:
    url = f"https://ntfy.sh/{topic}"
    alert_page = alert_page_url() or "(run Setup Phone App.bat / Publish Now.bat first)"
    text = f"""PM NTE daily mobile alert (05:30 PM)

1) Install free app on Android: ntfy
   Play Store: https://play.google.com/store/apps/details?id=io.heckel.ntfy

2) Open ntfy → + → Subscribe to topic (exact name):
   {topic}

   Or open this on phone:
   {url}

3) Keep notifications allowed for ntfy.
   In ntfy topic settings, turn ON Instant delivery if available.

4) IMPORTANT for big full-screen alert:
   TAP the notification (do not only glance at the small shade).
   It opens a large-font page that stays until you press CLEAR ALERT.

Full-screen alert page (also open anytime):
   {alert_page}

Every day at 5:30 PM your PC will send:
- Today closed vs 21 target
- Whether target was achieved
- Which FMEs are behind (need 3 WOs each)

PC must be ON at 5:30 PM (or Auto Sync running) for the alert to send.
Update/save the Excel before 5:30 so numbers are current.

Test anytime: double-click "Send Test Alert.bat"
"""
    (ROOT / "DAILY-ALERT-SETUP.txt").write_text(text, encoding="utf-8")


def run_alert(send: bool = True) -> dict:
    refresh_data()
    title, body, details = build_summary()
    write_alert_data(details)

    print(title)
    print(body)
    click = alert_page_url()
    if click:
        print(f"Full-screen alert: {click}")
    print()

    if send:
        topic = (os.environ.get("NTFY_TOPIC") or "").strip() or load_config().get("topic")
        if not topic:
            print("NTFY_TOPIC not set; wrote alert page but skipped popup")
            return details
        achieved = None if details.get("friday") else details.get("achieved")
        send_ntfy(topic, title, body, achieved, click_url=click)
        print("Sent ntfy popup")
        print("On phone: TAP the notification to open the big full-screen alert.")
    return details


def setup() -> None:
    print()
    print("=" * 60)
    print("  Setup Daily Mobile Alert (05:30 PM)")
    print("=" * 60)
    print()
    cfg = ensure_config(interactive=True)
    write_phone_instructions(cfg["topic"])
    create_windows_task(cfg.get("time", "17:30"))
    print()
    print("Your private alert topic:")
    print(f"  {cfg['topic']}")
    print()
    print("NEXT — on your Android phone:")
    print("  1. Install app: ntfy (from Play Store)")
    print(f"  2. Subscribe to topic: {cfg['topic']}")
    print(f"  3. Or open: https://ntfy.sh/{cfg['topic']}")
    print()
    print("Sending a test notification now…")
    run_alert(send=True)
    print()
    print("Instructions saved in DAILY-ALERT-SETUP.txt")
    print("Done.")


def main(argv: list[str]) -> None:
    if "--setup" in argv:
        setup()
        return
    if "--print" in argv:
        refresh_data()
        title, body, _ = build_summary()
        print(title)
        print(body)
        return
    run_alert(send=True)


if __name__ == "__main__":
    main(sys.argv[1:])
