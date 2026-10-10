#!/usr/bin/env python3
"""
Local web UI for the Instagram followers scraper & comparator.

One command (``python app.py``) starts the app and opens your browser
automatically. Type a target account, watch the anonymous collection live,
and download the Excel report.

This is a local, single-user tool: the server binds to 127.0.0.1 only and
holds state in memory. Nothing is sent anywhere except to Instagram itself.
"""

from __future__ import annotations

import datetime as dt
import socket
import sys
import tempfile
import threading
import time
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

try:
    import scraper
    from scraper import (
        LIST_TYPES,
        PlaywrightError,
        ScrapeError,
        load_reference_from_text,
        load_reference_list,
        save_batch_workbook,
        save_match_report,
        save_to_excel,
        scrape_followers,
    )
except ImportError as exc:
    print("[x] Dependencies are missing. Run:")
    print("       pip install -r requirements.txt && python -m playwright install chromium")
    print(f"    (detail: {exc})")
    sys.exit(1)

APP_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = APP_DIR / "output"
MAX_TARGETS = 10   # accounts per batch run; one browser is used sequentially
RETRY_DELAY_SECONDS = 3.0   # cool-down before the one end-of-batch retry per failed account

app = Flask(__name__)

# Very small in-memory job store: this is a single-user local tool.
_JOBS: dict[str, dict] = {}
_JOBS_LOCK = threading.Lock()
_job_counter = 0
_stop_flags: dict[str, threading.Event] = {}


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def _set_job(job_id: str, **fields) -> None:
    with _JOBS_LOCK:
        _JOBS[job_id].update(fields)


def _reference_from_upload(storage) -> list[str]:
    """Extract usernames from an uploaded .xlsx/.csv/.txt file."""
    suffix = Path(storage.filename or "").suffix.lower()
    if suffix not in {".xlsx", ".csv", ".txt", ".list"}:
        raise ValueError(
            f"Unsupported file type '{suffix or '(none)'}'. "
            "Upload .xlsx, .csv, .txt, or paste the names instead."
        )
    tmp = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
    tmp_path = Path(tmp.name)
    tmp.close()  # close the handle before Flask writes (Windows file locks)
    try:
        storage.save(tmp_path)
        return load_reference_list(tmp_path)
    finally:
        tmp_path.unlink(missing_ok=True)


def parse_targets(raw: str) -> list[str]:
    """Split the multi-account field into clean, unique targets.

    Separators are commas and/or whitespace; each entry is sanitized the
    same way as the old single-target field. Entries that sanitize to
    nothing usable (empty, pure punctuation) are dropped.
    """
    targets: list[str] = []
    seen: set[str] = set()
    for piece in (raw or "").replace(",", " ").split():
        target = scraper.sanitize_username(piece)
        # Case-insensitive dedup: Instagram names are case-insensitive, and
        # 'NASA' + 'nasa' must not become two sheets for one account.
        if target != "user" and target.lower() not in seen:
            seen.add(target.lower())
            targets.append(target)
    return targets


def _friendly_browser_error(exc: Exception) -> str:
    message = str(exc)
    if "Executable" in message or "playwright install" in message.lower():
        return ("Browser not installed. Run:  pip install -r requirements.txt "
                "&& python -m playwright install chromium")
    if "ProcessSingleton" in message or "SingletonLock" in message:
        return ("Another browser process could not be started. Close extra "
            "Chromium windows and try again.")
    if "has been closed" in message or "Target page, context or browser" in message:
        return ("The browser window was closed before the run finished - "
                "start the scrape again and leave the window open.")
    return message or "Unexpected error - see the console window."


def _attempt_target(job_id: str, target: str, index: int, total: int,
                    headless: bool, max_followers: int, list_type: str,
                    stop_flag: threading.Event,
                    label: str = "") -> tuple[dict | None, str | None, bool]:
    """One scrape attempt for one account (used by the batch loop and its
    end-of-batch retry). Returns (entry, error, abort): exactly one of
    entry/error is set; ``abort`` means the browser itself died, so further
    accounts are pointless."""
    suffix = f" ({label})" if label else ""
    _set_job(job_id, phase=f"Opening Instagram (@{target} "
                           f"{index + 1}/{total}){suffix}", count=0,
             expected=None, recent=[])

    pass_no_seen = 0   # engine reopens the list when Instagram pauses it

    def on_progress(count: int, **info) -> None:
        nonlocal pass_no_seen
        if info.get("pass_number"):
            pass_no_seen = max(pass_no_seen, int(info["pass_number"]))
        note = (f" - Instagram paused the list; re-opening it "
                f"(attempt {pass_no_seen})" if pass_no_seen else "")
        update = {"phase": f"Collecting from @{target} "
                           f"({index + 1}/{total}){suffix}{note}", "count": count}
        if info.get("expected"):
            update["expected"] = info["expected"]
        if info.get("recent") is not None:
            # Prefix each name with its account so the live list
            # stays unambiguous in multi-account runs.
            update["recent"] = [f"{target}:{n}" for n in info["recent"]]
        _set_job(job_id, **update)

    try:
        (followers, partial, stopped_early,
         expected, finish_reason, suggested) = scrape_followers(
            target, headless=headless, progress=on_progress,
            max_followers=max_followers,
            should_stop=stop_flag.is_set if stop_flag else None,
            list_type=list_type,
        )
    except ScrapeError as exc:
        return None, str(exc), False
    except PlaywrightError as exc:
        return None, _friendly_browser_error(exc), True
    return {
        "followers": followers,
        "partial": partial,
        "stopped_early": stopped_early,
        "expected": expected,
        "finish_reason": finish_reason,
        "suggestions": suggested,
    }, None, False


def _run_job(job_id: str, targets: list[str], reference: list[str],
             headless: bool, max_followers: int, list_type: str,
             stop_flag: threading.Event) -> None:
    """Worker: scrape each target in turn (one shared browser, sequentially -
    the profile cannot be shared between parallel runs), then write the
    workbook and update the job record. A failed account never aborts the
    batch: it is recorded and the next account starts, and after the main
    pass every failed account gets exactly one retry (transient Instagram
    hiccups - rate-limit pages, momentary errors - otherwise lose accounts
    that would succeed a moment later)."""
    results: dict[str, dict] = {}
    errors: dict[str, str] = {}
    try:
        total = len(targets)
        for index, target in enumerate(targets):
            if stop_flag is not None and stop_flag.is_set():
                break   # user pressed Stop: no further accounts are started
            entry, error, abort = _attempt_target(
                job_id, target, index, total, headless, max_followers,
                list_type, stop_flag)
            if entry:
                results[target] = entry
                continue
            errors[target] = error
            if abort:
                break

        # One end-of-batch retry per failed account (batch runs only; a
        # single-account failure just reports, and the user can re-run).
        if len(targets) > 1 and errors and not (
                stop_flag is not None and stop_flag.is_set()):
            if RETRY_DELAY_SECONDS:
                time.sleep(RETRY_DELAY_SECONDS)
            for index, target in enumerate(targets):
                if target not in errors:
                    continue
                if stop_flag is not None and stop_flag.is_set():
                    break   # user pressed Stop mid-retry: keep what we have
                entry, error, abort = _attempt_target(
                    job_id, target, index, total, headless, max_followers,
                    list_type, stop_flag, label="retry")
                if entry:
                    del errors[target]
                    results[target] = entry
                    continue
                errors[target] = error   # newest message wins
                if abort:
                    break

        if not results:
            detail = "; ".join(f"@{t}: {errors[t]}" for t in targets if t in errors)
            raise ScrapeError("No accounts could be collected. " + detail)

        _set_job(job_id, phase="Writing Excel report",
                 count=sum(len(e["followers"]) for e in results.values()))
        now = dt.datetime.now()

        if reference:
            # Compare mode is single-target only (enforced in /start).
            target = targets[0]
            entry = results[target]
            matches, not_found, extras = scraper.compare_followers(
                entry["followers"], reference)
            path = save_match_report(
                target, matches, not_found, extras,
                len(reference), entry["partial"], now, OUTPUT_DIR,
                expected=entry["expected"],
                finish_reason=entry["finish_reason"],
            )
            result = {
                "mode": "compare",
                "reference": len(reference),
                "scraped": len(entry["followers"]),
                "matches": len(matches),
                "not_found": len(not_found),
                "extras": len(extras),
                "partial": entry["partial"],
                "stopped_early": entry["stopped_early"],
                "expected": entry["expected"],
                "finish_reason": entry["finish_reason"],
                "rows": {
                    "matches": matches[:500],
                    "not_found": not_found[:500],
                    "extras": extras[:500],
                },
                "file": path.name,
            }
        elif len(targets) == 1:
            target = targets[0]
            entry = results[target]
            path = save_to_excel(
                target, entry["followers"], now, OUTPUT_DIR, entry["partial"],
                expected=entry["expected"],
                finish_reason=entry["finish_reason"],
                suggestions=entry["suggestions"],
            )
            result = {
                "mode": "scrape",
                "list_type": list_type,
                "scraped": len(entry["followers"]),
                "partial": entry["partial"],
                "stopped_early": entry["stopped_early"],
                "expected": entry["expected"],
                "finish_reason": entry["finish_reason"],
                "suggestions": entry["suggestions"],
                "rows": {"followers": entry["followers"][:500]},
                "file": path.name,
            }
        else:
            path = save_batch_workbook(targets, list_type, now, OUTPUT_DIR,
                                       results)
            accounts = []
            for target in targets:
                entry = results.get(target)
                if entry:
                    accounts.append({
                        "account": target,
                        "collected": len(entry["followers"]),
                        "expected": entry["expected"],
                        "status": "ok",
                        "error": None,
                        "suggestions": entry["suggestions"],
                        "rows": entry["followers"][:500],
                    })
                elif target in errors:
                    accounts.append({
                        "account": target, "collected": 0, "expected": None,
                        "status": "failed", "error": errors[target], "rows": [],
                    })
                else:
                    accounts.append({
                        "account": target, "collected": 0, "expected": None,
                        "status": "not run (stopped)", "error": None, "rows": [],
                    })
            result = {
                "mode": "batch_scrape",
                "list_type": list_type,
                "accounts": accounts,
                "done": sum(1 for a in accounts if a["status"] == "ok"),
                "failed": sum(1 for a in accounts if a["status"] != "ok"),
                "stopped_early": any(e["stopped_early"]
                                     for e in results.values()),
                "file": path.name,
            }

        _set_job(job_id, status="done", phase="Done", result=result,
                 message=None)
    except ScrapeError as exc:
        _set_job(job_id, status="error", message=str(exc))
    except Exception as exc:  # noqa: BLE001 - surface anything to the UI
        _set_job(job_id, status="error", message=_friendly_browser_error(exc))
    finally:
        _stop_flags.pop(job_id, None)


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/start")
def start():
    global _job_counter
    raw_targets = request.form.get("target", "")
    if not raw_targets.strip():
        return jsonify({"error": "Please type at least one target Instagram account."}), 400
    targets = parse_targets(raw_targets)
    if not targets:
        return jsonify({"error": "No usable Instagram usernames in the target field."}), 400
    if len(targets) > MAX_TARGETS:
        return jsonify({
            "error": f"Please enter no more than {MAX_TARGETS} unique accounts per run."
        }), 400
    if len(targets) > 1 and (request.form.get("names", "").strip()
                             or (request.files.get("file")
                                 and request.files.get("file").filename)):
        return jsonify({"error": "Compare mode works with a single account - "
                                 "put one username in the target field."}), 400

    pasted = request.form.get("names", "")
    upload = request.files.get("file")
    reference: list[str] = []
    try:
        if upload is not None and upload.filename:
            reference = _reference_from_upload(upload)
        elif pasted.strip():
            reference = load_reference_from_text(pasted)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    # One scrape at a time: each job launches its own anonymous browser.
    with _JOBS_LOCK:
        if any(job.get("status") == "running" for job in _JOBS.values()):
            return jsonify({
                "error": "A scrape is already running - wait for it or press Stop.",
                "busy": True,
            }), 409

    # Same digits-only rule the UI enforces: anything else means the user
    # probably mistyped ("5hundred") and an uncapped run would be a surprise.
    raw_max = (request.form.get("max") or "").strip()
    if raw_max and (not raw_max.isdigit() or int(raw_max) <= 0):
        return jsonify({
            "error": "The 'Stop after' field needs a plain number like 500 - "
                     "or leave it empty for the whole list."
        }), 400
    max_followers = int(raw_max) if raw_max else 0

    list_type = (request.form.get("list") or "followers").strip()
    if list_type not in LIST_TYPES:
        return jsonify({"error": "List must be 'followers' or 'following'."}), 400
    if reference:
        list_type = "followers"

    stop_flag = threading.Event()

    with _JOBS_LOCK:
        _job_counter += 1
        job_id = str(_job_counter)
        _stop_flags[job_id] = stop_flag
        _JOBS[job_id] = {
            "status": "running",
            "phase": "Starting browser",
            "count": 0,
            "expected": None,
            "recent": [],
            "result": None,
            "message": None,
        }

    threading.Thread(
        target=_run_job,
        args=(job_id, targets, reference, request.form.get("headless") != "no",
              max_followers, list_type, stop_flag),
        daemon=True,
    ).start()
    return jsonify({
        "job_id": job_id,
        "reference_size": len(reference),
        "accounts": len(targets),
        "list": list_type,
        "mode": "anonymous",
    })


@app.post("/stop/<job_id>")
def stop(job_id: str):
    flag = _stop_flags.get(job_id)
    if flag is None:
        return jsonify({"error": "Job not running"}), 404
    flag.set()
    return jsonify({"ok": True})


@app.get("/status/<job_id>")
def status(job_id: str):
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
        if job is None:
            return jsonify({"error": "Unknown job"}), 404
        return jsonify(job)


@app.get("/download/<job_id>")
def download(job_id: str):
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
    if not job or job.get("status") != "done":
        return jsonify({"error": "Report not ready"}), 404
    path = OUTPUT_DIR / job["result"]["file"]
    if not path.exists():
        return jsonify({"error": "Report file is missing"}), 404
    return send_file(path, as_attachment=True)


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    port = _free_port()
    url = f"http://127.0.0.1:{port}"
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    print(f"[*] Instagram Follower Collector running at {url}  (Ctrl+C to stop)")
    print("[!] Keep this window open while you use the tool.")
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
