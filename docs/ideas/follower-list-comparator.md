# Follower List Comparator

## Problem Statement
How might we find out which people from an existing list follow a given Instagram account, using only anonymous access?

## Recommended Direction
Extend `scraper.py` with a `--compare <file>` mode. The tool scrapes the target's followers exactly as it does today, loads the user's list, normalizes usernames (strip `@`, lowercase, accept pasted profile URLs), and writes a match workbook: **Matches**, **On your list — not found in scrape**, **Scraped — not on your list**, plus the existing **Info** sheet.

The reason this is the right direction: the user told us the actual job is comparison, not collection. Collection is the commodity part (v1 exists); the comparison layer is where the value sits, and it costs almost nothing to build on top of the existing pipeline.

## Key Assumptions to Validate
- [ ] Your list contains **usernames**, not display names or emails — if it's names, exact matching fails silently (test: open the file and look)
- [ ] Target accounts are small enough that the anonymous scrape is effectively complete — otherwise matches are unreliable (test: run on a real target and note where the login wall bites)
- [ ] Zero matches will be read as "sampling artifact" rather than "ground truth" (test: seed the list with one manually confirmed follower of a small account and confirm the tool finds them)

## MVP Scope
- `python scraper.py <target> --compare mylist.xlsx|csv|txt` (stdin with `-` for pasted text)
- Username normalization on both sides; case-insensitive matching
- Four-sheet output workbook (Matches / Not found from your list / Extra from scrape / Info with completeness warning)
- Explicit warning in console and Info sheet when the scrape was partial

## Not Doing (and Why)
- **Growth/counts tracking** — parked; not the stated job
- **Engagement scraping** — different question, different tool
- **Team portal GUI** — CLI is fine until comparison proves valuable
- **Logged-in or paid data for complete lists** — explicitly declined for now
- **Fuzzy matching of real names** — masks data-quality problems; exact usernames only in v1

## Open Questions
- ~~What format is the existing list actually in?~~ **Resolved: xlsx, csv, or pasted text — all supported in `--compare`.**
- How big are the target accounts, typically? (Determines how trustworthy "not found" really is.)
- What happens downstream on a match — does the report shape need to serve outreach?
