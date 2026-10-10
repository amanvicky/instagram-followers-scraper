# Instagram Follower Collector

A small local web app that scrapes public Instagram follower/following lists into an Excel report. It is designed to work without a login and without a saved browser profile, so it works for anyone running it locally.

## What it does

- Collects followers or following names for one or more public Instagram accounts
- Exports a report to Excel in the `output/` folder
- Supports compare mode against a pasted or uploaded list
- Accepts up to 10 unique targets per batch run
- Keeps the process local-only: it runs from your machine and talks to Instagram directly from a fresh browser context

## No login required

This app intentionally works in anonymous public mode. It opens a fresh Playwright browser context, loads the public dialog, and records whatever Instagram exposes without a stored session or login prompt.

That means:

- no account setup is required
- no saved browser session is reused
- some large or protected accounts may return a small sample or zero names
- the workbook still records when the public list was blocked or incomplete

## Quick start

Windows:

```bat
start.bat
```

macOS/Linux:

```bash
./start.sh
```

If you are starting manually:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python -m playwright install chromium
python app.py
```

## Output format

Each run creates an Excel workbook in `output/`.

- single account: a follower/following sheet plus an Info sheet
- compare mode: Matches / Not found / Extras / Info sheets
- batch mode: a Summary sheet plus one sheet per successful target

## Validation

Run the regression suite from the project root:

```bash
python -m unittest discover -s tests -t . -v
```

The tests cover route validation, scraper behavior, public-mode login wall handling, compare mode, and workbook metadata. They do not hit Instagram directly.

## Project status

This project is intentionally a local-first, public-source utility rather than a hosted service. The app is designed to be easy to run on any machine, inspect in source form, and improve through pull requests.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE).

## Notes

- The app is local-only and does not require persistent login state.
- Some Instagram profiles intentionally hide or limit their public lists.
- Do not share a browser profile or commit local session data.
- `browser_data/` and `output/` are ignored by git to keep the workspace clean.
