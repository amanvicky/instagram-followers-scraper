# Instagram Public List Collector

Collect the follower or following usernames Instagram exposes on public profiles and export them to an Excel workbook. The app runs locally and does not ask for Instagram credentials.

Instagram controls what logged-out visitors can see. A run may be partial or empty, and the app reports that rather than promising a complete list.

## What it does

- Collects followers or following from public profiles in a fresh browser context.
- Processes up to 10 target accounts in one batch.
- Compares a collected list with usernames you paste or upload.
- Exports single-account, batch, or comparison reports as `.xlsx` workbooks.
- Shows collection progress and records when Instagram limits the visible results.

## Quick start

You need Python 3.9 or newer. On first launch, the startup script creates a virtual environment, installs the dependencies, and installs Chromium for Playwright.

Windows:

```bat
start.bat
```

macOS or Linux:

```bash
./start.sh
```

The app opens in your browser and listens on your own computer. Keep the terminal window open while using it.

### Manual setup

```bash
python -m venv .venv
```

Activate the environment, then install the project dependencies and browser:

```bash
python -m pip install -r requirements.txt
python -m playwright install chromium
python app.py
```

On Windows, activate with `.venv\Scripts\activate`. On macOS or Linux, use `source .venv/bin/activate`.

## Use a run

1. Enter one or more public Instagram usernames, separated by commas or spaces.
2. Choose Followers or Following and, optionally, set a per-account limit.
3. Optionally paste a reference list or upload an `.xlsx`, `.csv`, `.txt`, or `.list` file to compare.
4. Review the result and its status, then download the Excel workbook.

Reports are written to `output/`. Compare workbooks contain Matches, Not found, Extras, and Info sheets. Batch workbooks include a summary and a sheet for each successful target.

## Privacy and limitations

- The Flask app binds to `127.0.0.1`; it is not a hosted scraping service.
- The browser uses a fresh context. The app does not request or reuse an Instagram login session.
- Browser automation visits Instagram. Instagram may show a login prompt, throttle requests, change its interface, or hide some or all names.
- Results describe only what was visible during that run. They are not guaranteed to be complete, current, or suitable as an authoritative record.
- Use responsibly, respect Instagram's terms and rate limits, and handle exported usernames with care.

## Tests

Run the regression suite from the repository root:

```bash
python -m unittest discover -s tests -t . -v
```

The suite covers request validation, job handling, scraper behavior, anonymous login-wall handling, comparisons, and workbook output.

Two longer manual integration checks are kept under `tests/manual/`. They exercise a real local HTTP server and Playwright against a locally intercepted mock Instagram site:

```bash
python tests/manual/app_http_check.py
python tests/manual/engine_lifecycle_check.py
```

## License

MIT. See [LICENSE](LICENSE).
