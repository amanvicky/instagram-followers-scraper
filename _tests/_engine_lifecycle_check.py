"""Lifecycle check: REAL scraper.collect_followers + REAL Chromium + fake Instagram.

Playwright route interception serves a fake instagram.com (profile page,
followers dialog, paginated GraphQL endpoint). The real engine code runs
unmodified against it. Four runs, each asserting observable behavior:

  1. full lifecycle: counter parse, dedup, stall -> captured-request
     replay (nested node.user.username shape, cursor pagination,
     completion), finish_reason == "exhausted", progress callbacks
  2. max_followers stop -> finish_reason == "cap"
  3. login-wall dialog -> ScrapeError mentioning login
  4. missing profile -> ScrapeError "not available"
  plus a real save_to_excel -> re-read via openpyxl asserts

Run: .venv/Scripts/python.exe _tests/_engine_lifecycle_check.py
"""
import datetime as dt
import faulthandler
import json
import re
import sys
import tempfile
from pathlib import Path

# If the engine wedges, dump every thread's stack and exit after 90s.
faulthandler.dump_traceback_later(90, exit=True)

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import scraper  # noqa: E402

# Real pacing constants are rate-limit safety for real Instagram; against a
# local fake they only slow the check.
scraper.SCROLL_PAUSE_SECONDS = 0.05

PAGE_SIZE = 50
API_TOTAL = 100  # fb0..fb99 served by the fake GraphQL endpoint
state = {"mode": "big"}  # big | small | wall


def dialog_html(rows, wall=False):
    links = "".join(
        f'<div class="row"><a href="/{n}">{n}</a></div>' for n in rows
    )
    login_link = '<a href="/accounts/login/">Log in to see followers</a>' if wall else ""
    return f"""<!DOCTYPE html><html><body>
<div role="dialog">{login_link}
  <div class="scroll" style="height:300px;overflow-y:auto;">
    <div style="height:900px;">{links}</div>
  </div>
</div>
<script>
  document.querySelector('.scroll').addEventListener('scroll', () => {{
    fetch('/api/graphql/query', {{method:'POST',
      body:'variables=' + encodeURIComponent(JSON.stringify({{first:50, after:''}}))}});
  }});
</script>
</body></html>"""


def profile_html():
    return """<!DOCTYPE html><html><body>
<h1>Test account</h1><span>1,500 followers</span><span>312 following</span>
<a href="/nasa/followers/">followers</a>
</body></html>"""


def api_response(after: str):
    start = 0 if after == "" else int(after[1:]) * PAGE_SIZE
    usernames = [f"fb{i}" for i in range(start, min(start + PAGE_SIZE, API_TOTAL))]
    has_more = start + PAGE_SIZE < API_TOTAL
    cursor = f"C{start // PAGE_SIZE + 1}" if has_more else None
    return {"data": {"user": {"edge_followed_by": {
        "edges": [{"node": {"user": {"username": u}}} for u in usernames],
        "page_info": {"end_cursor": cursor, "has_next_page": has_more},
    }}}}


def route_handler(route, request):
    url = request.url
    if "/api/graphql/query" in url:
        after = ""
        m = re.search(r"variables=([^&]+)", request.post_data or "")
        if m:
            from urllib.parse import unquote_plus
            after = (json.loads(unquote_plus(m.group(1))) or {}).get("after", "") or ""
        return route.fulfill(status=200, content_type="application/json",
                             body=json.dumps(api_response(after)))
    if url.startswith("https://www.instagram.com/ghost"):
        return route.fulfill(status=200, content_type="text/html; charset=utf-8",
                             body="<html><body>Sorry, this page isn't available.</body></html>")
    if "/followers" in url or "/following" in url:
        mode = state["mode"]
        if mode == "big":
            body = dialog_html([f"fb{i}" for i in range(40)])
        elif mode == "small":
            body = dialog_html([f"fb{i}" for i in range(8)])
        else:  # wall
            body = dialog_html([], wall=True)
        return route.fulfill(status=200, content_type="text/html; charset=utf-8", body=body)
    return route.fulfill(status=200, content_type="text/html; charset=utf-8",
                         body=profile_html())


failures = []
def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + (f"  [{detail}]" if detail and not cond else ""),
          flush=True)
    if not cond:
        failures.append(label)


with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=True)
    context = browser.new_context()
    context.route("**/*", route_handler)

    followers = []

    # --- Run 1: full lifecycle with API fallback ----------------------
    page = context.new_page()
    progress_calls = []
    print("[harness] run1 starting", flush=True)
    followers, wall, early, expected, reason, suggested = scraper.collect_followers(
        page, "nasa", timeout=15,
        progress=lambda c, **kw: progress_calls.append((c, dict(kw))),
    )
    print(f"[harness] run1 done: {len(followers)} names, reason={reason}", flush=True)
    check("collected exactly 100 unique followers", len(followers) == 100, f"got {len(followers)}")
    check("no duplicates (case-insensitive)",
          len({f.lower() for f in followers}) == len(followers))
    check("names are exactly fb0..fb99",
          sorted(followers, key=lambda s: int(s[2:])) == [f"fb{i}" for i in range(100)])
    check("no login wall hit", wall is False)
    check("not stopped early", early is False)
    check("expected counter parsed as 1500", expected == 1500, f"got {expected}")
    check("finish_reason is 'exhausted' (API completed the list)", reason == "exhausted", f"got {reason}")
    counts = [c for c, _ in progress_calls]
    check("progress counts are non-decreasing and end at 100",
          bool(counts) and all(a <= b for a, b in zip(counts, counts[1:])) and counts[-1] == 100,
          f"counts={counts}")
    last_kw = progress_calls[-1][1]
    check("progress reports expected to the UI", last_kw.get("expected") == 1500)
    check("progress recent-window carries the tail names",
          (last_kw.get("recent") or [])[-1] == "fb99")
    page.close()

    # --- Run 2: cap stop ----------------------------------------------
    state["mode"] = "small"
    page2 = context.new_page()
    print("[harness] run2 starting", flush=True)
    f2, w2, e2, x2, r2, _sug2 = scraper.collect_followers(page2, "nasa", timeout=15,
                                                   max_followers=25)
    print(f"[harness] run2 done: {len(f2)} names, reason={r2}", flush=True)
    check("cap stop: collection stops near the cap (dialog batch)", 25 <= len(f2) < 100,
          f"got {len(f2)}")
    check("cap stop flagged", e2 is True)
    check("cap finish_reason is 'cap'", r2 == "cap", f"got {r2}")
    page2.close()

    # --- Run 3: login wall returns an anonymous empty result -------------
    state["mode"] = "wall"
    page3 = context.new_page()
    (wall_names, hit_wall, wall_stopped, _, wall_reason, _) = scraper.collect_followers(
        page3, "nasa", timeout=15)
    check("login wall returns no public names", wall_names == [])
    check("login wall is recorded", hit_wall is True)
    check("login wall is not a manual stop", wall_stopped is False)
    check("login wall has its own finish reason", wall_reason == "login_wall")
    page3.close()

    # --- Run 4: missing profile ----------------------------------------
    page4 = context.new_page()
    try:
        scraper.collect_followers(page4, "ghost", timeout=15)
        check("missing profile raises ScrapeError", False, "no exception raised")
    except scraper.ScrapeError as exc:
        check("missing profile raises ScrapeError", True)
        check("missing-profile message says not available", "not available" in str(exc), str(exc))
    page4.close()

    # --- Run 5: following list (different dialog URL + counter) --------
    state["mode"] = "small"
    page5 = context.new_page()
    f5, w5, e5, x5, r5, _sug5 = scraper.collect_followers(
        page5, "nasa", timeout=15, max_followers=5, list_type="following")
    check("following run collects the dialog rows", len(f5) == 8, f"got {len(f5)}")
    check("following run reads the following counter (312)", x5 == 312, f"got {x5}")
    check("following run honors the cap flag", e5 is True and r5 == "cap", f"{e5}/{r5}")
    page5.close()

    # --- Excel: real file, real reader --------------------------------------
    if followers:
        with tempfile.TemporaryDirectory() as tmp:
            path = scraper.save_to_excel("nasa", followers,
                                         dt.datetime(2026, 9, 28, 12, 0, 0),
                                         Path(tmp), partial=False, expected=1500,
                                         finish_reason="exhausted")
            check("xlsx file created in target dir", path.exists() and path.parent == Path(tmp))
            wb = load_workbook(path, read_only=True)
            check("workbook has Followers + Info sheets",
                  set(wb.sheetnames) == {"Followers", "Info"}, str(wb.sheetnames))
            ws = wb["Followers"]
            rows = list(ws.iter_rows(values_only=True))
            check("header row is #/Username/Profile URL",
                  list(rows[0]) == ["#", "Username", "Profile URL"], str(rows[0]))
            data = rows[1:]
            check("one row per follower (100)", len(data) == 100, f"got {len(data)}")
            check("first row fb0 with profile URL",
                  data[0] == (1, "fb0", "https://www.instagram.com/fb0/"), str(data[0]))
            check("last row fb99 numbered 100",
                  data[-1] == (100, "fb99", "https://www.instagram.com/fb99/"), str(data[-1]))
            check("row numbers are 1..100 in order",
                  [r[0] for r in data] == list(range(1, 101)))
            info = {r[0]: r[1] for r in wb["Info"].iter_rows(values_only=True) if r[0]}
            wb.close()
            check("Info: collected count is 100", info.get("Usernames collected") == 100)
            check("Info: Instagram counter is 1500",
                  info.get("Instagram's follower counter") == 1500)
            check("Info: honest incompleteness (100 < 1500)",
                  str(info.get("List complete?", "")).startswith("Possibly incomplete"),
                  str(info.get("List complete?")))
            check("Info: mode says anonymous",
                info.get("Mode") == "Anonymous (no login)", str(info.get("Mode")))

    context.close()
    browser.close()

print("ALL PASS" if not failures else f"{len(failures)} CHECK(S) FAILED", flush=True)
sys.exit(1 if failures else 0)
