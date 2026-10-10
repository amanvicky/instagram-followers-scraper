#!/usr/bin/env python3
"""Collect publicly exposed Instagram follower/following names anonymously.

Each scrape uses a fresh, non-persistent browser context and saves its Excel
report under ./output/. Instagram may limit or hide lists for logged-out
visitors; this tool does not sign in or reuse browser sessions.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import re
import sys
import time
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlsplit, urlunsplit

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

OUTPUT_DIR = Path(__file__).resolve().parent / "output"

SCROLL_RETRIES = 12        # consecutive unproductive scrolls before giving up
SCROLL_PAUSE_SECONDS = 1.2  # wait after each scroll for Instagram to load more rows
STALL_CHECKPOINT = 4        # stalled rounds between scroll-recovery attempts
MAX_STALL_RECOVERIES = 3    # scroll-jiggle recovery cycles before giving up
MAX_API_ATTEMPTS = 3        # bounded retries of direct API pagination
API_PAGE_SIZE = 50          # rows per replayed pagination request
MAX_COLLECTION_ROUNDS = 2000  # absolute backstop: never scroll-collect forever
MAX_FRESH_PER_ROUND = 120     # a followers dialog cannot show this many NEW rows at once
MAX_COLLECT_PASSES = 3        # fresh-dialog retries when Instagram stops early
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
)


class ScrapeError(Exception):
    """A friendly, expected failure (bad username, login wall, markup change...)."""


# --- JavaScript helpers (run inside the browser) -----------------------------

# Collect unique profile usernames from anchor links inside the followers
# dialog (falls back to the whole page if no dialog is open).
_COLLECT_JS = r"""
() => {
  const dialog = document.querySelector('div[role="dialog"]');
  const scope = dialog || document;
  const skip = new Set([
    'explore', 'accounts', 'p', 'reel', 'reels', 'tv', 'stories', 'about',
    'legal', 'direct', 'developer', 'directories', 'web', 'ajax', 'challenge',
    'emails', 'session', 'oauth', 'graphql', 'your_activity', 'topics',
    'search', 'explore people'
  ]);
  const seen = new Set();
  const names = [];
  for (const a of scope.querySelectorAll('a[href]')) {
    const href = a.getAttribute('href') || '';
    const m = href.match(/^\/([^\/?#]+)\/?$/);
    if (!m) continue;
    const name = decodeURIComponent(m[1]);
    if (!name || skip.has(name.toLowerCase()) || seen.has(name)) continue;
    seen.add(name);
    names.push(name);
  }
  return names;
}
"""

# Scroll the dialog's scrollable list container to the bottom.
_SCROLL_JS = r"""
() => {
  const dialog = document.querySelector('div[role="dialog"]');
  if (!dialog) return false;
  for (const el of dialog.querySelectorAll('*')) {
    const s = getComputedStyle(el);
    if ((s.overflowY === 'auto' || s.overflowY === 'scroll') &&
        el.scrollHeight > el.clientHeight) {
      el.scrollTop = el.scrollHeight;
      return true;
    }
  }
  return false;
}
"""

# True when Instagram replaced the list with a login prompt.
_LOGIN_WALL_JS = r"""
() => {
  const dialog = document.querySelector('div[role="dialog"]');
  if (!dialog) return false;
  const text = dialog.innerText || '';
  return /log in to see|log in to view|sign up to see/i.test(text) ||
         !!dialog.querySelector('a[href="/accounts/login/"]');
}
"""

# Close the popups Instagram throws over the page (notifications prompt,
# "save your login info", etc.) - their overlays otherwise swallow every
# click the scraper tries to make.
_DISMISS_JS = r"""
() => {
  let closed = 0;
  for (const btn of document.querySelectorAll('button')) {
    const label = (btn.getAttribute('aria-label') || '').toLowerCase();
    const text = (btn.innerText || '').trim().toLowerCase();
    if (label === 'close' || text === 'not now' || text === 'cancel') {
      try { btn.click(); closed++; } catch (e) {}
    }
  }
  return closed;
}
"""

# Mid-scroll stall recovery. Deliberately does NOT press Escape and does NOT
# click aria-label="close" buttons: the follower dialog itself contains a
# close button, and the recovery must never close the dialog it is
# harvesting. Safe actions only: dismiss "Not now"/"Cancel" popup buttons,
# then re-trigger the virtual list with real wheel events plus a small
# up-then-bottom jiggle of the scroll container.
_RECOVERY_JS = r"""
() => {
  let revived = false;
  for (const btn of document.querySelectorAll('button')) {
    const text = (btn.innerText || '').trim().toLowerCase();
    if (text === 'not now' || text === 'cancel') {
      try { btn.click(); revived = true; } catch (e) {}
    }
  }
  const dialog = document.querySelector('div[role="dialog"]');
  if (!dialog) return false;
  let container = null;
  for (const el of dialog.querySelectorAll('*')) {
    const s = getComputedStyle(el);
    if ((s.overflowY === 'auto' || s.overflowY === 'scroll') &&
        el.scrollHeight > el.clientHeight) { container = el; break; }
  }
  if (!container) return revived;
  const rect = container.getBoundingClientRect();
  const cx = rect.left + rect.width / 2;
  const cy = rect.top + Math.min(rect.height / 2, 300);
  for (const dy of [900, 1400]) {
    container.dispatchEvent(new WheelEvent('wheel', {
      bubbles: true, cancelable: true, deltaY: dy, clientX: cx, clientY: cy
    }));
  }
  container.scrollTop = Math.max(0, container.scrollTop - 240);
  container.scrollTop = container.scrollHeight;
  return true;
}
"""

# Replay the dialog's own follower-pagination request from inside the page
# (same-origin cookies ride along automatically). This bypasses DOM rendering
# entirely, which is what breaks around ~900 rows on big lists.
_API_FETCH_JS = r"""
async (args) => {
  const headers = { ...(args.headers || {}) };
  const opts = args.body
    ? { method: "POST",
        headers: { "content-type": "application/x-www-form-urlencoded",
                   ...headers },
        body: args.body }
    : { method: "GET", credentials: "include", headers };
  const r = await fetch(args.url, opts);
  if (!r.ok) return { ok: false, status: r.status };
  return { ok: true, json: await r.json() };
}
"""


class _FollowerRequestCapture:
    """Records the follower-dialog's own pagination request so it can be
    replayed with a new cursor when the DOM stops rendering.

    Two request shapes exist in the wild:
    - GraphQL POST with variables (edges/page_info responses)
    - GET /api/v1/friendships/<id>/followers/ (users/next_max_id responses)
    """

    def __init__(self):
        self.kind: str | None = None   # 'graphql' | 'friendships'
        self.url: str | None = None
        self.post_data: str | None = None
        self.headers: dict = {}

    def attach(self, page) -> None:
        def on_request(req):
            try:
                if "graphql/query" in req.url and req.post_data:
                    body = req.post_data
                    if "variables" in body and "first" in body:
                        self.kind, self.url, self.post_data = (
                            "graphql", req.url, body)
                        self.headers = dict(req.headers)
                elif "/api/v1/friendships/" in req.url and re.search(
                        r"/(followers|following)", req.url):
                    self.kind, self.url, self.post_data = (
                        "friendships", req.url, None)
                    self.headers = dict(req.headers)
            except Exception:  # noqa: BLE001 - capture must never break a run
                pass
        page.on("request", on_request)


def _find_edges_block(node):
    """Locate {edges, page_info} anywhere in a GraphQL response (Instagram has
    shipped several shapes; this handles all of them)."""
    if isinstance(node, dict):
        if "edges" in node and "page_info" in node:
            return node
        for value in node.values():
            found = _find_edges_block(value)
            if found is not None:
                return found
    elif isinstance(node, list):
        for value in node:
            found = _find_edges_block(value)
            if found is not None:
                return found
    return None


def _extract_names_and_cursor(payload: dict):
    """Pull (usernames, next_cursor, has_more) from either response shape."""
    block = _find_edges_block(payload)
    if block:
        edges = block.get("edges", [])
        names = [((e or {}).get("node") or {}).get("username")
                 for e in edges]
        if not any(names):
            # Instagram also ships a shape where the user object is nested
            # one level deeper (node.user.username) - handle it so the API
            # fallback never sees an empty page that actually had rows.
            names = ((((e or {}).get("node") or {}).get("user") or {})
                     .get("username") for e in edges)
            names = list(names)
        info = block.get("page_info") or {}
        return ([n for n in names if n], info.get("end_cursor"),
                bool(info.get("has_next_page")))
    # friendships shape: {users: [{username}], next_max_id, has_more}
    users = payload.get("users")
    if isinstance(users, list):
        names = [((u or {}).get("username")) for u in users]
        return ([n for n in names if n], payload.get("next_max_id"),
                bool(payload.get("has_more")))
    return [], None, False


def _replay_followers_api(page, capture, seen: set, followers: list,
                          progress, max_followers: int, should_stop,
                          expected=None):
    """Replay the captured pagination request until the follower list is
    exhausted (or the cap/stop fires). Returns (added, complete | None).
    Replays page through already-collected rows (deduplicated by ``seen``)
    until they reach fresh territory, so no state is lost by the switch."""
    if not capture.url:
        return 0, None

    graphql_payload: dict = {}
    base_vars: dict = {}
    if capture.kind == "graphql" and capture.post_data:
        graphql_payload = {k: v[0] for k, v in parse_qs(capture.post_data).items()}
        if graphql_payload.get("variables"):
            try:
                base_vars = json.loads(graphql_payload["variables"])
            except ValueError:
                base_vars = {}

    added = 0
    complete = None
    _api_cursor = None
    for _ in range(400):  # hard bound: 400 pages x 50 = 20k names max
        if should_stop is not None and should_stop():
            break
        if max_followers and len(followers) >= max_followers:
            break

        headers = dict(capture.headers or {})
        if capture.kind == "graphql":
            payload = dict(graphql_payload)
            vars_obj = dict(base_vars)
            vars_obj["first"] = API_PAGE_SIZE
            if _api_cursor:
                vars_obj["after"] = _api_cursor
            payload["variables"] = json.dumps(vars_obj)
            url = capture.url
            body = urlencode(payload)
            csrf = page.evaluate(
                "() => (document.cookie.match(/(?:^|; )csrftoken=([^;]+)/) || [])[1]"
            )
            if csrf:
                headers["x-csrf-token"] = csrf
        else:
            # friendships GET: swap count/max_id in the query string.
            parts = urlsplit(capture.url)
            query = {k: v[0] for k, v in parse_qs(parts.query).items()}
            query["count"] = str(API_PAGE_SIZE)
            query.pop("max_id", None)
            if _api_cursor:
                query["max_id"] = _api_cursor
            url = urlunsplit((parts.scheme, parts.netloc, parts.path,
                              urlencode(query), parts.fragment))
            body = None

        try:
            result = page.evaluate(_API_FETCH_JS,
                                   {"url": url, "body": body, "headers": headers})
        except PlaywrightError:
            break
        if not result or not result.get("ok"):
            break

        payload_json = result.get("json") or {}
        names, cursor, has_more = _extract_names_and_cursor(payload_json)
        fresh = 0
        for name in names:
            key = name.lower()
            if key not in seen:
                seen[key] = name
                followers.append(name)
                added += 1
                fresh += 1
        if progress:
            progress(len(followers), expected=expected, recent=followers[-40:])

        if not names and not added:
            # First replayed page came back with nothing recognizable (wrong
            # request captured, or Instagram refused it). Do not trust it as
            # "complete" - report failure so the run falls back to recovery.
            return added, None
        if not has_more or not cursor:
            complete = True
            break
        # Replayed pages are full of already-collected names until the cursor
        # passes the stall point - that is expected; keep paginating (the
        # 400-page loop bound above is the hard stop).
        _api_cursor = cursor
        page.wait_for_timeout(400)
    return added, complete

# The two lists Instagram shows on a profile. Both live in a dialog with the
# same markup; only the URL, the counter wording and the API route differ.
LIST_TYPES = {"followers", "following"}

# Follower-count badge readers (the yardstick for completeness, e.g.
# "1,304 followers" -> 1304; "312 following" -> 312).
_EXPECTED_JS = r"""
() => {
  const t = document.body ? (document.body.innerText || '') : '';
  const m = t.match(/([\d.,]+)\s*([KM])?\s*followers/i);
  if (!m) return null;
  let n = parseFloat(m[1].replace(/,/g, ''));
  if (!Number.isFinite(n) || n <= 0) return null;
  if (m[2]) n = m[2].toUpperCase() === 'M' ? n * 1e6 : n * 1e3;
  return Math.round(n);
}
"""

_FOLLOWING_JS = r"""
() => {
  const t = document.body ? (document.body.innerText || '') : '';
  const m = t.match(/([\d.,]+)\s*([KM])?\s*following/i);
  if (!m) return null;
  let n = parseFloat(m[1].replace(/,/g, ''));
  if (!Number.isFinite(n) || n <= 0) return null;
  if (m[2]) n = m[2].toUpperCase() === 'M' ? n * 1e6 : n * 1e3;
  return Math.round(n);
}
"""

_LIST_CONFIGS = {
    "followers": {
        "list_word": "followers",
        "dialog_url": "https://www.instagram.com/{username}/followers/",
        "link_regex": r"followers\s*$",
        "counter_js": _EXPECTED_JS,
    },
    "following": {
        "list_word": "following",
        "dialog_url": "https://www.instagram.com/{username}/following/",
        "link_regex": r"following\s*$",
        "counter_js": _FOLLOWING_JS,
    },
}

# Fallback collector: follower rows that are plain text, not anchors. Reads
# dialog lines that look like usernames and are not UI labels.
_COLLECT_TEXT_JS = r"""
() => {
  const dialog = document.querySelector('div[role="dialog"]');
  if (!dialog) return [];
  const out = [];
  const seen = new Set();
  for (const line of (dialog.innerText || '').split('\n')) {
    const t = line.trim().replace(/^@/, '');
    if (!t || t.length > 30 || seen.has(t.toLowerCase())) continue;
    if (/^[0-9.,]+$/.test(t)) continue;  // follower counts, not usernames
    if (/^(search|close|cancel|not now|ok|done|suggested|follow|following|followers|about|remove|block|message)\b/i.test(t)) continue;
    if (/followers$|following$/i.test(t)) continue;
    if (!/^[A-Za-z0-9._]+$/.test(t)) continue;
    if (!/[A-Za-z]/.test(t)) continue;  // usernames contain at least one letter
    seen.add(t.toLowerCase());
    out.push(t);
  }
  return out;
}
"""


# --- CLI ----------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Scrape an Instagram account's publicly visible follower "
                    "list into an Excel file."
    )
    parser.add_argument(
        "username", nargs="?", default=None,
        help="Instagram username, e.g. nasa",
    )
    parser.add_argument(
        "-o", "--output-dir", type=Path, default=OUTPUT_DIR,
        help="Folder for the .xlsx file (default: ./output inside this repo)",
    )
    parser.add_argument(
        "--headless", action=argparse.BooleanOptionalAction, default=True,
        help="Run without a visible browser window (use --no-headless to watch).",
    )
    parser.add_argument(
        "--timeout", type=float, default=30.0,
        help="Seconds to wait for Instagram page elements (default: 30).",
    )
    parser.add_argument(
        "--compare", metavar="FILE",
        help="Compare scraped followers against usernames in FILE "
             "(.xlsx/.csv/.txt, or '-' to paste via stdin).",
    )
    parser.add_argument(
        "--max", type=int, default=0, metavar="N",
        help="Stop after collecting N usernames (0 = no cap).",
    )
    parser.add_argument(
        "--diff", nargs=2, metavar=("NEW_XLSX", "OLD_XLSX"),
        help="Compare two follower-report files (no scraping): how many "
             "names are new since the old run, how many disappeared.",
    )
    return parser.parse_args()


def sanitize_username(raw: str) -> str:
    """Strip an '@' and anything that would break a URL or file name."""
    return re.sub(r"[^A-Za-z0-9._-]", "", raw.strip().lstrip("@")) or "user"


# --- Reference list loading and comparison ---------------------------------------

_USERNAME_RE = re.compile(r"^[A-Za-z0-9._]{1,30}$")
_PROFILE_URL_RE = re.compile(r"instagram\.com/([A-Za-z0-9._]+)/?", re.I)
# Instagram path segments and spreadsheet header words that are features,
# not usernames (e.g. /p/<post-id>, or a "followers" header column).
_RESERVED_SEGMENTS = {
    "p", "reel", "reels", "explore", "accounts", "stories", "tv", "direct",
    "about", "legal", "sessions", "challenge", "followers", "following",
    "follow", "posts", "tagged", "emails", "web", "developer", "directories",
}
# @mention inside surrounding text; the lookbehind skips email addresses
# (in "john@example.com" the @ is attached to a word character).
_MENTION_RE = re.compile(r"(?<![A-Za-z0-9])@([A-Za-z0-9._]+)")


def _extract_username(token: str) -> str | None:
    """Pull a plausible Instagram username out of a raw cell or line.

    Tolerates '@name', profile URLs, and mentions inside surrounding text
    like 'John Doe (@johndoe)'. Rejects emails and other non-usernames.
    """
    token = token.strip()
    if not token:
        return None

    candidate = None
    url_match = _PROFILE_URL_RE.search(token)
    if url_match:
        candidate = url_match.group(1)
    else:
        plain = token.lstrip("@").strip()
        if not any(ch.isspace() for ch in plain) and "/" not in plain:
            candidate = plain
        else:
            mention = _MENTION_RE.search(token)
            if mention:
                candidate = mention.group(1)

    if candidate is None:
        return None
    candidate = candidate.rstrip(".")
    if (not _USERNAME_RE.match(candidate)
            or candidate.isdigit()
            or candidate.lower() in _RESERVED_SEGMENTS):
        return None
    return candidate


def load_reference_list(source: Path | str) -> list[str]:
    """Load the user's list of usernames from .xlsx, .csv, .txt, or stdin ('-')."""
    if str(source) == "-":
        raw_items = sys.stdin.read().split()
        source_label = "<pasted text>"
    else:
        path = Path(source)
        if not path.exists():
            raise ValueError(f"file not found: {path}")
        suffix = path.suffix.lower()
        if suffix == ".xlsx":
            wb = load_workbook(path, read_only=True, data_only=True)
            raw_items = []
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    for cell in row:
                        if cell is not None:
                            raw_items.append(str(cell))
            wb.close()
        elif suffix == ".csv":
            with open(path, newline="", encoding="utf-8-sig") as fh:
                raw_items = [cell for row in csv.reader(fh) for cell in row]
        elif suffix in {".txt", ".list", ""}:
            raw_items = path.read_text(encoding="utf-8-sig", errors="replace").split()
        else:
            raise ValueError(f"unsupported file type '{path.suffix}' "
                             "(use .xlsx, .csv, .txt, or '-' for stdin)")
        source_label = str(path)

    return _finalize_reference(raw_items, source_label)


def _finalize_reference(raw_items, source_label: str) -> list[str]:
    usernames = normalize_usernames(raw_items)
    print(f"[*] Loaded {len(usernames)} usernames from {source_label}")
    return usernames


def normalize_usernames(raw_items) -> list[str]:
    """Extract, deduplicate (case-insensitive) and keep order of usernames."""
    usernames: list[str] = []
    seen: set[str] = set()
    for item in raw_items:
        name = _extract_username(item)
        if name and name.lower() not in seen:
            seen.add(name.lower())
            usernames.append(name)
    return usernames


def load_reference_from_text(text: str) -> list[str]:
    """Load usernames from pasted text (any whitespace-separated entries)."""
    return normalize_usernames(text.split())


def compare_followers(scraped: list[str], reference: list[str]) -> tuple[list[str], list[str], list[str]]:
    """Match reference usernames against the scraped set (case-insensitive).

    Returns (matches, on_list_not_found, scraped_not_on_list), each list
    preserving the order of its source.
    """
    scraped_lookup = {name.lower(): name for name in scraped}
    matches: list[str] = []
    not_found: list[str] = []
    for entry in reference:
        if entry.lower() in scraped_lookup:
            matches.append(entry)
        else:
            not_found.append(entry)
    ref_lower = {entry.lower() for entry in reference}
    extras = [name for name in scraped if name.lower() not in ref_lower]
    return matches, not_found, extras


def diff_followers(new_run: list[str], old_run: list[str]) -> tuple[list[str], list[str]]:
    """Run-to-run diff for convergence measurement (case-insensitive).

    Returns (new_names, lost_names): names in the new run that the old run
    never had, and names the old run had that the new one does not. Order
    follows the source runs.
    """
    old_lower = {name.lower() for name in old_run}
    new_names = [name for name in new_run if name.lower() not in old_lower]
    new_lower = {name.lower() for name in new_run}
    lost = [name for name in old_run if name.lower() not in new_lower]
    return new_names, lost


def _style_header_row(ws) -> None:
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4F46E5")
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")


def _format_list_sheet(ws) -> None:
    for column, width in (("A", 6), ("B", 32), ("C", 44)):
        ws.column_dimensions[column].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def save_match_report(username: str, matches: list[str], not_found: list[str],
                      extras: list[str], reference_size: int, partial: bool,
                      scraped_at: dt.datetime, out_dir: Path,
                      expected: int | None = None,
                      finish_reason: str | None = None) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = scraped_at.strftime("%Y%m%d_%H%M%S")
    path = out_dir / f"{username}_match_report_{stamp}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Matches"
    ws.append(["#", "Username", "Profile URL"])
    _style_header_row(ws)
    for index, name in enumerate(matches, start=1):
        ws.append([index, name, f"https://www.instagram.com/{name}/"])
    _format_list_sheet(ws)

    for title, rows in (("Not found from your list", not_found),
                        ("Scraped - not on your list", extras)):
        sheet = wb.create_sheet(title)
        sheet.append(["#", "Username", "Profile URL"])
        _style_header_row(sheet)
        for index, name in enumerate(rows, start=1):
            sheet.append([index, name, f"https://www.instagram.com/{name}/"])
        _format_list_sheet(sheet)

    info = wb.create_sheet("Info")
    scraped_total = len(matches) + len(extras)
    status = ("Partial - Instagram hid the rest behind login"
              if partial else
              "Possibly incomplete - Instagram throttled collection"
              if finish_reason == "throttled" and expected is None else
              "Appears complete")
    rows = [
        ("Target account", f"@{username}"),
        ("Scraped at", scraped_at.strftime("%Y-%m-%d %H:%M:%S")),
        ("Reference list size", reference_size),
        ("Usernames scraped", scraped_total),
        ("Matches", len(matches)),
        ("Not found from your list", len(not_found)),
        ("Scrape completeness", status),
    ]
    if expected is not None:
        rows.append(("Instagram's follower counter", expected))
        rows.append(("Gap (counter minus collected)", max(0, expected - scraped_total)))
        if not partial and scraped_total < expected:
            status = (f"Possibly incomplete - profile reports {expected} "
                      f"followers, {scraped_total} collected")
            rows[-3] = ("Scrape completeness", status)
        rows.append(("Gap note", "A gap usually means deleted or hidden "
                     "('ghost') accounts - or a throttled run - not rows "
                     "missing from this file."))
    rows.extend([
        ("Caveat", "If the scrape was partial, 'not found' does not prove "
                  "someone does not follow the target."),
        ("Mode", "Anonymous (no login)"),
    ])
    for row in rows:
        info.append(row)
    info.column_dimensions["A"].width = 30
    info.column_dimensions["B"].width = 60

    wb.save(path)
    return path


# --- Combined multi-account workbook -----------------------------------------------

def _batch_complete_text(entry: dict) -> str:
    """One-line per-account completeness verdict for the Summary sheet."""
    if entry.get("error"):
        return "-"
    if entry.get("partial"):
        return "No - login wall"
    if entry.get("stopped_early"):
        return "No - stopped early"
    expected, count = entry.get("expected"), len(entry.get("followers") or [])
    if expected is not None and count < expected:
        return f"Partial - gap of {expected - count}"
    if entry.get("finish_reason") == "throttled":
        return "Possibly incomplete - throttled"
    if entry.get("suggestions"):
        return "Mixed with suggested accounts"
    return "Appears complete"


def save_batch_workbook(targets: list[str], list_type: str,
                        scraped_at: dt.datetime, out_dir: Path,
                        results: dict[str, dict]) -> Path:
    """One workbook for a multi-account run: a Summary sheet with one row per
    target plus one sheet per successful account, so every account's list
    stays separately readable. A target whose scrape failed keeps its Summary
    row (with the error) but gets no sheet of its own.

    ``results`` maps username -> {followers, partial, stopped_early, expected,
    finish_reason, error}.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = scraped_at.strftime("%Y%m%d_%H%M%S")
    path = out_dir / f"batch_{list_type}_{stamp}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append(["#", "Account", "List", "Names collected", "Instagram's counter",
               "Complete?", "Result"])
    _style_header_row(ws)
    for column, width in (("A", 6), ("B", 26), ("C", 12), ("D", 16),
                          ("E", 20), ("F", 30), ("G", 44)):
        ws.column_dimensions[column].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for index, target in enumerate(targets, start=1):
        entry = results.get(target, {"error": "no result recorded"})
        ws.append([
            index,
            f"@{target}",
            list_type,
            len(entry.get("followers") or []),
            entry.get("expected") if not entry.get("error") else None,
            _batch_complete_text(entry),
            entry.get("error") or "ok",
        ])

    for target in targets:
        entry = results.get(target) or {}
        followers = entry.get("followers") or []
        if not followers:
            continue
        sheet = wb.create_sheet(target)
        sheet.append(["#", "Username", "Profile URL"])
        _style_header_row(sheet)
        for row_index, name in enumerate(followers, start=1):
            sheet.append([row_index, name, f"https://www.instagram.com/{name}/"])
        _format_list_sheet(sheet)

    wb.save(path)
    return path


# --- Scraping ------------------------------------------------------------------

def collect_followers(page, username: str, timeout: float,
                      progress=None, max_followers: int = 0,
                      should_stop=None, list_type: str = "followers"
                      ) -> tuple[list[str], bool, bool, int | None, str]:
    """Open the profile, open the Followers or Following dialog (per
    ``list_type``: "followers" or "following") and scroll while collecting
    usernames. Returns (usernames, hit_login_wall, stopped_early,
    expected, finish_reason).

    Runs anonymously and captures whatever Instagram exposes. If Instagram
    redirects to or displays a sign-in wall, the result reports
    ``hit_login_wall=True`` without requesting credentials.

    ``progress`` is called as ``progress(count, expected=N, recent=[...])``
    after each round (web UI live progress: ``expected`` is Instagram's own
    counter for the chosen list, ``recent`` the newest names),
    ``should_stop`` is polled each round (Stop button), and ``max_followers``
    caps collection (0 = no cap)."""
    if list_type not in LIST_TYPES:
        raise ScrapeError(f"Unknown list type '{list_type}'.")
    timeout_ms = timeout * 1000
    capture = _FollowerRequestCapture()
    capture.attach(page)
    list_config = _LIST_CONFIGS[list_type]
    url = f"https://www.instagram.com/{username}/"
    print(f"[*] Opening {url}")
    page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
    page.wait_for_timeout(1_500)
    if "/accounts/login" in page.url:
        return [], True, False, None, "login_wall", False

    body_text = page.evaluate(
        "() => document.body ? document.body.innerText : ''"
    ) or ""  # odd/fake page contexts can yield None; treat as empty
    if re.search(r"sorry, this page isn'?t available", body_text, re.I):
        raise ScrapeError(
            f"Instagram says the profile '{username}' is not available "
            "(check the spelling; the account may be private or removed)."
        )

    # Dismiss popups whose overlays intercept clicks (observed live: the
    # followers link resolved but the click never landed because of them).
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(400)
        page.evaluate(_DISMISS_JS)
        page.wait_for_timeout(400)
    except PlaywrightError:
        pass

    # Yardstick for completeness: Instagram's own counter badge on the
    # profile (e.g. "1,304 followers" / "312 following"). Captured BEFORE any
    # collecting starts and reported immediately (CLI log + web UI) so the run
    # has a visible target from the first row.
    expected = None
    try:
        expected = page.evaluate(list_config["counter_js"])
    except PlaywrightError:
        pass
    if expected is not None:
        print(f"[*] Instagram reports {expected} {list_config['list_word']} "
              f"for @{username}.")
        if progress:
            progress(0, expected=expected)

    # Preferred route: the followers dialog has its own URL - no click needed.
    opened = False
    try:
        page.goto(list_config["dialog_url"].format(username=username),
                  wait_until="domcontentloaded", timeout=timeout_ms)
        page.wait_for_selector('div[role="dialog"]', timeout=timeout_ms)
        opened = True
    except PlaywrightTimeoutError:
        if "/accounts/login" in page.url:
            return [], True, False, expected, "login_wall", False
        pass  # fall through to the click approach below

    if not opened:
        # The list control is a link/button whose text ends in the list word
        # (e.g. "104M\nfollowers"; verified against Instagram's current DOM).
        link = page.locator("a, button").filter(
            has_text=re.compile(list_config["link_regex"], re.I)
        ).first
        try:
            link.wait_for(state="visible", timeout=timeout_ms)
        except PlaywrightTimeoutError:
            if "/accounts/login" in page.url:
                return [], True, False, expected, "login_wall", False
            raise ScrapeError(
                f"Could not find a {list_config['list_word']} link on "
                f"@{username}'s profile. The account may be private, or "
                "Instagram's markup changed."
            )
        try:
            link.click(timeout=10_000)
        except PlaywrightError:
            # Overlay still intercepting - dispatch a synthetic click on the
            # element itself, which bypasses pointer-event interception.
            link.evaluate("el => el.click()")

        try:
            page.wait_for_selector('div[role="dialog"]', timeout=timeout_ms)
        except PlaywrightTimeoutError:
            if "/accounts/login" in page.url:
                return [], True, False, expected, "login_wall", False
            raise ScrapeError(
                f"The {list_config['list_word']} dialog did not open "
                "(markup change?).")

    followers: list[str] = []
    seen: dict[str, str] = {}   # lowercase -> original casing (case-insensitive dedup)
    recoveries = 0
    api_attempted = False
    api_attempts = 0
    api_exhausted = False   # an API attempt ran but added nothing new

    last_count = -1
    stall_rounds = 0
    wall_hit = page.evaluate(_LOGIN_WALL_JS) or "/accounts/login" in page.url
    stopped_early = False
    scroll_worked = True
    rounds = 0
    round_limited = False
    pass_no = 1
    pass_start = 0   # follower count at the start of the current pass

    while True:
        rounds += 1
        if rounds > MAX_COLLECTION_ROUNDS:
            round_limited = stopped_early = True
            print(f"\n[!] Safety limit: {MAX_COLLECTION_ROUNDS} collection "
                  "rounds - stopping with what was gathered.")
            break
        names = page.evaluate(_COLLECT_JS)
        if not names:
            names = page.evaluate(_COLLECT_TEXT_JS)
        fresh_round = 0
        for name in names:
            key = name.lower()
            if key not in seen:
                seen[key] = name
                followers.append(name)
                fresh_round += 1
        if fresh_round > MAX_FRESH_PER_ROUND:
            raise ScrapeError(
                "Internal safety stop: a single collection round produced far "
                "more new names than a followers dialog can show, which means "
                "the collector was re-reading rows it already had. Nothing is "
                "saved from a failed run - just start the scrape again."
            )
        print(f"\r[+] {len(followers)} usernames collected...", end="", flush=True)
        if progress:
            progress(len(followers), expected=expected,
                     recent=followers[-40:])

        if should_stop is not None and should_stop():
            stopped_early = True
            print("\n[i] Stop requested - keeping what was collected.")
            break

        if max_followers and len(followers) >= max_followers:
            stopped_early = True
            print(f"\n[i] Reached the {max_followers}-username cap - keeping what was collected.")
            break

        if page.evaluate(_LOGIN_WALL_JS) or "/accounts/login" in page.url:
            wall_hit = True
            print("\n[!] Instagram is asking to log in - kept everything visible anonymously.")
            break

        if len(followers) == last_count:
            stall_rounds += 1
            if (stall_rounds >= STALL_CHECKPOINT
                    and stall_rounds % STALL_CHECKPOINT == 0):
                if capture.url and api_attempts < MAX_API_ATTEMPTS:
                    # The DOM list stalls around ~900 rows because Instagram
                    # throttles the dialog. Replay the dialog's own pagination
                    # requests - server-side, rendering-proof. Replays resume
                    # through already-seen pages thanks to dedup, bounded.
                    api_attempts += 1
                    api_attempted = True
                    stall_rounds = 0
                    print("\n[i] Dialog stalled - switching to direct API "
                          "pagination (fetched from Instagram's own requests)...")
                    try:
                        added, api_complete = _replay_followers_api(
                            page, capture, seen, followers, progress,
                            max_followers, should_stop, expected=expected,
                        )
                    except PlaywrightError:
                        added, api_complete = 0, None
                    if api_complete:
                        print(f"\n[i] Full list retrieved via API pagination "
                              f"({len(followers)} total).")
                        break
                    if added:
                        last_count = len(followers)
                        print(f"\n[i] API added {added} more names - "
                              "resuming scroll watch.")
                        continue
                    api_exhausted = True
                elif recoveries < MAX_STALL_RECOVERIES:
                    # No captured request to replay (or replay added nothing):
                    # jiggle the virtual list back to life.
                    recoveries += 1
                    stall_rounds = 0
                    print(f"\n[i] Scroll stalled - recovery {recoveries}/"
                          f"{MAX_STALL_RECOVERIES}: dismissing popups and "
                          "re-triggering the list...")
                    try:
                        page.evaluate(_RECOVERY_JS)
                        page.wait_for_timeout(1_500)
                    except PlaywrightError:
                        pass
            if stall_rounds >= SCROLL_RETRIES:
                list_complete = (expected is not None
                                 and len(followers) >= expected)
                if (list_complete or pass_no >= MAX_COLLECT_PASSES
                        or len(followers) == pass_start):
                    print(f"\n[i] List stopped growing after {SCROLL_RETRIES} "
                          f"attempts ({recoveries} recoveries, "
                          f"api={'used' if api_attempted else 'n/a'}) - done.")
                    break
                # Instagram cut this pass short but names are still missing:
                # reopen the dialog fresh (the throttle resets per dialog) and
                # keep collecting - dedup means nothing is lost or duplicated.
                print(f"\n[i] Pass {pass_no} ended at {len(followers)} of "
                      f"{expected} - Instagram stopped early. Reopening the "
                      f"dialog for pass {pass_no + 1}...")
                pass_no += 1
                pass_start = len(followers)
                rounds = 0
                stall_rounds = 0
                last_count = -1
                api_attempts = 0
                api_exhausted = False
                recoveries = 0
                if progress:
                    progress(len(followers), expected=expected,
                             pass_number=pass_no)
                try:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(1_000)
                    page.goto(
                        list_config["dialog_url"].format(username=username),
                        wait_until="domcontentloaded", timeout=timeout_ms)
                    page.wait_for_selector('div[role="dialog"]',
                                           timeout=timeout_ms)
                    page.wait_for_timeout(1_500)
                    continue
                except PlaywrightError:
                    print("\n[!] Could not reopen the followers dialog - "
                          "keeping what this pass collected.")
                    break
        else:
            stall_rounds = 0
            last_count = len(followers)

        if not page.evaluate(_SCROLL_JS):
            if scroll_worked:
                print("\n[!] Could not find the scrollable list container.")
            scroll_worked = False
        # Adaptive backoff: when Instagram throttles, hammering the scroll
        # makes it worse. Wait longer the longer we have been stalled.
        pause = SCROLL_PAUSE_SECONDS * (1 + min(stall_rounds, 6))
        page.wait_for_timeout(int(pause * 1000))

    if expected is not None and len(followers) >= expected:
        print(f"\n[OK] Complete: {len(followers)} of {expected} followers "
              f"collected in {pass_no} pass(es).")
    finish_reason = ("login_wall" if wall_hit
                     else "cap" if stopped_early and max_followers
                     else "round_limit" if round_limited
                     else "stopped" if stopped_early
                     else "throttled" if api_exhausted
                     else "exhausted")
    if expected is not None and finish_reason == "exhausted":
        gap = expected - len(followers)
        if gap > 0:
            print(f"\n[i] Instagram's counter says {expected} followers - "
                  f"{gap} more than collected. Deleted/hidden ('ghost') "
                  "accounts usually explain the difference.")

    if list_type == "following":
        # An account cannot follow itself; if Instagram served the target's
        # own profile link inside the list (verified in the suggested-accounts
        # padding case), drop it here so no report contains it.
        followers = [n for n in followers
                     if n.lower() != username.lower()]
    followers, suggested = _flag_following_suggestions(
        followers, expected, list_type)
    if suggested:
        print(f"\n[!] Collected {len(followers)} names but Instagram's counter "
              f"says {expected} - the list is padded with 'Suggested for you' "
              "accounts, not real follows.")
    return followers, wall_hit, stopped_early, expected, finish_reason, suggested


def _flag_following_suggestions(followers: list[str], expected: int | None,
                                list_type: str) -> tuple[list[str], bool]:
    """Detect Instagram's suggested-accounts padding in a following list.

    On sparse accounts Instagram pads the Following dialog/API with a
    "Suggested for you" block. Verified live: a counter of 1 produced 33
    names, including the account itself. Signals: collected more names than
    the counter claims, and/or the account following itself. Returns the
    cleaned list (self-name removed) and whether suggestions were detected.
    """
    if list_type != "following" or not followers:
        return followers, False
    # The counter only becomes a meaningful signal once the collection has
    # more than a handful of names - tiny lists are dominated by noise.
    padded = expected is not None and len(followers) > max(10, expected)
    return followers, padded


# --- Browser runner ----------------------------------------------------------------

def scrape_followers(target: str, headless: bool = True, timeout: float = 30.0,
                     progress=None, max_followers: int = 0,
                     should_stop=None,
                     list_type: str = "followers"
                     ) -> tuple[list[str], bool, bool, int | None, str, bool]:
    """Scrape the target's follower (or following) list using the shared
    persistent browser profile (so the saved Instagram login is reused across
    runs).

    Shared by the CLI and the web UI. Raises ScrapeError for expected
    failures and PlaywrightError if the browser itself cannot start.
    """
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=headless)
        try:
            context = browser.new_context(
                user_agent=USER_AGENT,
                viewport={"width": 1280, "height": 900},
                locale="en-US",
            )
            try:
                page = context.new_page()
                followers, wall, early, expected, reason, suggested = collect_followers(
                    page, target, timeout, progress=progress,
                    max_followers=max_followers, should_stop=should_stop,
                    list_type=list_type,
                )
                return followers, wall, early, expected, reason, suggested
            except PlaywrightTimeoutError as exc:
                raise ScrapeError(f"Timed out waiting for Instagram: {exc}") from exc
            finally:
                context.close()
        finally:
            browser.close()


# --- Excel export ----------------------------------------------------------------

def save_to_excel(username: str, followers: list[str], scraped_at: dt.datetime,
                  out_dir: Path, partial: bool, expected: int | None = None,
                  finish_reason: str | None = None,
                  suggestions: bool = False) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = scraped_at.strftime("%Y%m%d_%H%M%S")
    path = out_dir / f"{username}_followers_{stamp}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Followers"

    ws.append(["#", "Username", "Profile URL"])
    _style_header_row(ws)

    for index, name in enumerate(followers, start=1):
        ws.append([index, name, f"https://www.instagram.com/{name}/"])

    _format_list_sheet(ws)

    info = wb.create_sheet("Info")
    if partial:
        verdict = "No - Instagram showed a login wall"
    elif finish_reason in {"stopped", "cap", "round_limit"}:
        verdict = "No - run stopped early"
    elif expected is not None and len(followers) < expected:
        verdict = ("Possibly incomplete - profile reports more followers "
                   f"({expected}) than collected ({len(followers)})")
    elif finish_reason == "throttled" and expected is None:
        verdict = "Possibly incomplete - Instagram throttled collection"
    else:
        verdict = "Appears complete"
    rows = [
        ("Instagram username", f"@{username}"),
        ("Scraped at", scraped_at.strftime("%Y-%m-%d %H:%M:%S")),
        ("Usernames collected", len(followers)),
        ("List complete?", verdict),
    ]
    if expected is not None:
        gap = max(0, expected - len(followers))
        rows.append(("Instagram's follower counter", expected))
        rows.append(("Gap (counter minus collected)", gap))
        rows.append(("Gap note", "A gap usually means deleted or hidden "
                     "('ghost') accounts - or a throttled run - not rows "
                     "missing from this file."))
    rows.append(("Mode", "Anonymous (no login)"))
    if suggestions:
        rows.append(("Suggested accounts", "Yes - Instagram padded this "
                     "following list with 'Suggested for you' accounts. The "
                     "rows are not all real follows."))
    for row in rows:
        info.append(row)
    info.column_dimensions["A"].width = 30
    info.column_dimensions["B"].width = 60

    wb.save(path)
    return path


# --- Entry point -------------------------------------------------------------------

def _load_follower_names(path: Path) -> list[str]:
    """Read the Username column of a follower report (the Names sheet or
    sheet titled 'Followers' / 'Master')."""
    wb = load_workbook(path, read_only=True, data_only=True)
    sheet = None
    for title in ("Followers", "Master"):
        if title in wb.sheetnames:
            sheet = wb[title]
            break
    if sheet is None:
        sheet = wb.worksheets[0]
    names: list[str] = []
    header_seen = False
    for row in sheet.iter_rows(values_only=True):
        cell = row[1] if len(row) > 1 else None
        if cell is None:
            continue
        text = str(cell).strip()
        if not header_seen:
            if text.lower() == "username":
                header_seen = True
            continue
        if text and text != "None":
            names.append(text)
    wb.close()
    return names


def run_diff(new_path: Path, old_path: Path) -> int:
    """Print the run-to-run convergence numbers and return a CLI exit code."""
    try:
        new_names = _load_follower_names(new_path)
        old_names = _load_follower_names(old_path)
    except Exception as exc:  # noqa: BLE001 - friendly CLI error
        print(f"[x] Could not read the report files: {exc}", file=sys.stderr)
        return 1
    new_names = normalize_usernames(new_names)
    old_names = normalize_usernames(old_names)
    new_names_d, lost = diff_followers(new_names, old_names)
    print(f"[*] New run:  {new_path.name}  ({len(new_names)} names)")
    print(f"[*] Old run:  {old_path.name}  ({len(old_names)} names)")
    print(f"[=] New since the old run:      {len(new_names_d)}")
    print(f"[=] In the old run, now absent: {len(lost)}")
    if not new_names_d and not lost:
        print("[OK] Identical follower sets - fully converged.")
    elif len(new_names_d) <= max(10, len(new_names) // 20):
        print("[~] Small delta - the runs are converging.")
    else:
        print("[!] Large delta - the target is growing fast or Instagram "
              "is serving inconsistent slices between sessions.")
    return 0


def main() -> int:
    args = parse_args()

    if args.diff:
        new_path, old_path = (Path(p) for p in args.diff)
        return run_diff(new_path, old_path)

    if not args.username:
        print("[x] Please provide a target username, e.g.:  python scraper.py nasa",
              file=sys.stderr)
        return 2

    # Load and validate the reference list BEFORE launching the browser so a
    # bad file fails fast instead of wasting a full scrape.
    reference: list[str] = []
    if args.compare:
        try:
            reference = load_reference_list(args.compare)
        except Exception as exc:  # noqa: BLE001 - friendly CLI error for any file problem
            print(f"[x] Could not read reference list: {exc}", file=sys.stderr)
            return 1
        if not reference:
            print("[x] The reference list is empty - nothing to compare.", file=sys.stderr)
            return 1

    username = sanitize_username(args.username)
    try:
        (followers, wall_hit, stopped_early, expected, finish_reason,
         suggestions) = scrape_followers(
            username, headless=args.headless, timeout=args.timeout,
            max_followers=args.max,
        )
    except ScrapeError as exc:
        print(f"[x] {exc}", file=sys.stderr)
        return 1
    except PlaywrightError as exc:
        print(f"[x] Browser error: {exc}", file=sys.stderr)
        print("    First run? Install the browser with:  python -m playwright install chromium",
              file=sys.stderr)
        return 1

    now = dt.datetime.now()

    if args.compare:
        matches, not_found, extras = compare_followers(followers, reference)
        path = save_match_report(username, matches, not_found, extras,
                                 len(reference), wall_hit, now, args.output_dir,
                                 expected=expected, finish_reason=finish_reason)
        print(f"\n[*] Comparison: {len(matches)} matched, {len(not_found)} of your "
              f"list not found, {len(extras)} scraped names not on your list")
        if wall_hit and not_found:
            print("[!] The scrape was partial, so 'not found' may be a sampling "
                  "artifact rather than proof someone does not follow.")
        print(f"[OK] Match report saved to {path}")
        return 0

    path = save_to_excel(username, followers, now, args.output_dir, wall_hit,
                         expected=expected, finish_reason=finish_reason,
                         suggestions=suggestions)
    notes = []
    if stopped_early:
        notes.append("stopped early")
    if wall_hit:
        notes.append("Instagram did not expose the full list anonymously")
    if not followers and wall_hit:
        print("[!] Instagram did not expose any names to logged-out visitors; "
              "the empty report records that result.")
    suffix = f" ({'; '.join(notes)})" if notes else ""
    print(f"[OK] Saved {len(followers)} usernames to {path}{suffix}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
