# Contributing

Thanks for helping improve this project.

## Local setup

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python -m playwright install chromium
```

## Run the app

```bash
python app.py
```

## Run tests

```bash
python -m unittest discover -s tests -t . -v
```

## Pull requests

- Keep changes focused and easy to review
- Prefer small, testable improvements
- Update docs when behavior changes
- Verify the test suite before opening a PR

## Public-mode expectation

This project is intentionally designed for local, anonymous public scraping. Keep changes aligned with the public-source, local-first direction rather than adding login-required or hosted-service behavior.
