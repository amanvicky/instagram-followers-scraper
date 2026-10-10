# Living Master List

## Problem Statement
How might we keep one complete, always-current follower list per target — where every run adds value instead of creating another disposable file?

## Recommended Direction
**Master Archive (A)** is the core: every scrape of a target merges into a per-target master workbook (`master_<target>.xlsx`) instead of only producing a timestamped snapshot. The master holds every username ever seen with `first_seen`/`last_seen` dates, a **"New this run"** sheet, and a **Runs** history sheet. Repeated runs converge on a stable core (evidence so far: 890 → 1,173 → 1,302 across sessions for the same target) and never lose a name Instagram served once.

**Audience Compare (C)** falls out almost for free: the same case-insensitive match logic already built for user lists retargets two master files — common followers of two accounts, unique to each. It ships in the same MVP as a simple "Compare two targets" box in the UI.

**The Watcher (B) is deliberately gated.** The user wants scheduled autonomy; the honest constraint is that unattended background automation against Instagram from a personal session is how accounts get action-blocked. So Watcher ships as a conservative daily-catch-up mode (max 1–2 targets, one run each per day, randomized execution time inside a window, stop-on-any-error, no automatic retries), triggered by Windows Task Scheduler rather than an in-app daemon. If two weeks of such runs produce no warnings, the caps can be revisited.

## Key Assumptions to Validate
- [ ] Reruns converge: new names per run shrink toward zero for an unchanged target (test: three daily reruns on `unmaskwithaj`, record new-per-run counts).
- [ ] openpyxl merge stays fast and reliable at master-file sizes of 10–20k rows (test: synthetic 20k-row merge timing; measured baseline: 1,300-row write ≈ 0.6–2 s).
- [ ] Instagram tolerates ~1 modest logged-in run/day/target without warnings (test: two weeks of Task-Scheduler daily runs on one target).
- [ ] The ghost gap (counter − served ≈ 166 for this target) stays stable; a large jump is itself a signal worth a "counter anomaly" note (test: compare counter across the reruns above).

## MVP Scope
- Merge-on-scrape: after a successful run, merge into `master_<target>.xlsx` (sheets: Master with first_seen/last_seen, New this run, Runs history, Info). The existing per-run timestamped file is still produced.
- "Compare two targets" in the web UI: pick two master files → workbook with Common / Only in A / Only in B.
- Task Scheduler recipe (a documented `.bat` + one scheduled task the user creates once) for the daily catch-up run.
- Counter + gap recorded per run in the Runs sheet, so the archive itself accumulates the evidence.

## Not Doing (and Why)
- **In-app background daemon/scheduler** — account risk and a always-on-process burden; Task Scheduler does the job with zero new moving parts.
- **Engagement/DM data, "who interacted with whom"** — different question, different tool, higher risk.
- **Fuzzy matching of display names** — masks data-quality problems; exact usernames only, as in v1.
- **Multi-account session switching** — complexity and ban-risk concentration for no current need.
- **Auto-retry on failures** — failures during watching must stop and surface, not hammer.

## Open Questions
- Should "New this run" also flag *departures* (names in master but absent from the newest full list)? Requires enough runs to know whether absence is meaningful.
- Does the Runs history need export (e.g., a growth sparkline in the UI), or is the sheet enough?
- When two targets are compared, is the per-target master the right denominator, or should comparison run on the freshest snapshot?
