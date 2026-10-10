# Performance ledger

Measurements taken 2026-09-28. Environment: Chrome via Freebuff preview
(browser) and Pyodide 0.26.4 (Python sandbox, no threads; openpyxl runs
slower there than under CPython, so those numbers are upper bounds).
Terminal was unavailable this session; all measurements were taken on the
real shipped code paths via these two runtimes.

## Baseline measurements (real shipped code)

| Path | Measurement | Budget check |
|---|---|---|
| `_COLLECT_JS` per round (40-row dialog) | 0.036 ms → ~36 ms per 1,000-round run | invisible (< 10 ms/round) |
| `_COLLECT_JS` worst case (500-row rendered list) | 0.030 ms | invisible |
| `_COLLECT_TEXT_JS` per round (~120-line dialog text) | 0.017 ms → ~17 ms per run | invisible |
| Poll-loop chip render (40 chips/poll, ~1.4 polls/sec) | 0.005 ms/poll → ~2.2 ms per run | INP-safe by 3 orders |
| Result table HTML build (500-row UI cap) | 2.3 ms once | invisible |
| `normalize_usernames` (pathological 10k-line paste) | 42.6 ms → 8,000 names | invisible |
| `_extract_names_and_cursor` (400 API pages × 50) | 4.1 ms flat shape / 8.1 ms nested | invisible |
| `compare_followers` (1,300 scraped × 500 reference) | 0.80 ms | invisible |
| `save_to_excel` (1,300 rows, in-memory FS) | 576 ms in Pyodide (~1–2 s native worst case) | already surfaced as its own UI phase |

## Verdict: no optimization applied

Every measured cost is 2–5 orders of magnitude below the nearest perceptible
threshold. A real run's duration (~4–6 min for ~1,300 followers) is dominated
by **deliberate** throttle-avoidance waits (`SCROLL_PAUSE_SECONDS` 1.2 s+ per
round, adaptive backoff, per-page API pacing of 400 ms). Those waits protect
the account from Instagram rate limits; removing them would be a correctness
regression wearing a perf win's clothing.

## Attempts log (kept and rejected alike)

| Idea | Baseline → Result | Verdict | Why |
|---|---|---|---|
| Batch DOM reads / optimize `_COLLECT_JS` | 0.036 ms/round measured | rejected — never written | 3 orders below noise; nothing to gain |
| Incremental chip rendering (delta instead of 40-node innerHTML) | 0.005 ms/poll measured | rejected — never written | invisible; innerHTML rebuild is already free at this size |
| Stream/write Excel incrementally during collection | 576 ms/1,300 rows measured | rejected | write happens once per run, already has its own UI phase; streaming would couple engine to IO mid-collection |
| Cache `csrftoken` lookup across API pages | not measured separately | rejected — never written | one `page.evaluate` per 50-name page ≈ microseconds inside a 400 ms pacing wait |
| Reduce `SCROLL_PAUSE_SECONDS` / waits | n/a | rejected | rate-limit safety is a product requirement, not overhead |

## Rule for future sessions

Do not re-run these micro-benchmarks. If a *new* symptom appears (UI jank on
huge accounts, slow report writing at 20k rows — the 400-page loop's hard
bound), measure that specific path first; the numbers above bound everything
below ~1,300 rows as already free.
