# Get new ATS jobs in a Google Sheet every morning

This repository has a Google Apps Script (`Code.gs`) that adds new job postings to a Google Sheet once a day. It reads Greenhouse, Lever and Ashby boards (and, through the same Actor, Workday, Workable, Recruitee and Teamtailor boards) through the **ats-jobs-scraper** Actor on Apify. You don't need a server, and the script runs inside your Google account.

Actor used: [fecklesspleb/ats-jobs-scraper](https://apify.com/fecklesspleb/ats-jobs-scraper) (`fecklesspleb~ats-jobs-scraper`).

## What it does

1. A time-driven trigger runs `runDaily()` once a day (08:00 in the project time zone, default UTC).
2. The script calls the Actor's `run-sync-get-dataset-items` endpoint with the boards you list and waits for the results. Each run is capped at `maxTotalChargeUsd` 0.10 (at most 50 jobs at the Actor's $0.002-per-job price).
3. It skips any job whose `uid` (`ats:board:id`) is already in the sheet and appends the rest, one row per job. The job URL is not used as the key because Greenhouse URLs often point at the company's own careers site and can change.
4. Failures are written as a row on the **Run log** sheet (time, status, added, message). The token is never written to the sheet.

Columns, in order:

`uid, ats, board, company, title, location, workplace, posted_at, salary, url, first_seen`

## Setup (about 5 minutes)

1. Create a Google Sheet, then open **Extensions > Apps Script**.
2. Paste the contents of `Code.gs` and save. You can also copy `appsscript.json` into Project Settings if you are using the Apps Script CLI (`clasp`); the important parts are the three OAuth scopes and the time zone.
3. Under **Project Settings > Script Properties**, add `APIFY_TOKEN` (from Apify Console > Settings > API & Integrations) and `BOARDS`, a comma-separated list of board URLs or slugs (for example `https://boards.greenhouse.io/airbnb, lever:spotify, ashby:ramp`).
4. Run `runDaily()` once by hand and approve the permissions Google asks for. The first rows should appear.
5. Run `installTrigger()` once to schedule the daily run. It replaces any earlier `runDaily` trigger.

## Notes

- Your token stays in Script Properties. Don't paste it into the code, and don't share the sheet with edit access to people you don't trust.
- Each run uses Apify credits at the Actor's pay-per-result price ($0.002 per job returned; unresolved-board and status records are free). The script also turns off job descriptions (`includeDescription: false`) and turns on the Actor's own only-new-jobs state. Keep the board list short to start.
- Apps Script stops any execution that runs longer than 6 minutes. If you list many boards, split them across two sheets or triggers. A run that hits Apify's 300-second sync limit is logged and not retried until the next scheduled run.
- Prefer no Apify at all? The free Python script in [fecklesspleb/ats-new-jobs-feed](https://github.com/fecklesspleb/ats-new-jobs-feed) does the same thing with GitHub Actions, calling each board's public JSON directly.

## License

MIT. Copyright (c) 2026 fecklesspleb.
