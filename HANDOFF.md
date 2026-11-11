# Handoff

State of the timetable tool as of November 2026. Read `CLAUDE.md` first for
the rules of the codebase.

## Where things stand

Everything in `PLAN.md` (phases 0–9) is built, merged to `main` and tested by
the maintainer in Google Sheets. `npm test` passes (94 tests).

| PR | What |
|---|---|
| jaceranieri/timetable-tool#1 | The full overhaul: correct compliance results, template cleanup, name rules, the generator rewrite (far fewer spreadsheet calls, Sheets API batch copies), deployment polish, the guided sheet workflow (Start here tab, checks, copy helpers, Google Docs), the wizard redesign, editing an existing timetable, and the shared-code build. Also the OVERVIEW fix so each class is judged against its own minimums. |
| jaceranieri/timetable-tool#2 | Copy and Google Doc tools work when the user is signed into several Google accounts (they used `google.script.run` windows, which fail with PERMISSION_DENIED). |

## What the maintainer has to have done in Google

These can't be done from the repo; check they're in place if something
behaves oddly:

- Wizard project: deployed as a web app, *Execute as: User accessing*, *Who
  has access: Anyone within Sydney Catholic Schools*; **Sheets API** advanced
  service added. `LoadTest.gs` is not in the live project.
- Template project: **Drive API** v3 advanced service added; latest
  `TemplateSync.gs` and `TimetableTools.gs` pasted in.
- Master template spreadsheet: all the steps in `template/SHEET_CHANGES.md`
  applied (Amadeus check, OVERVIEW colour rules on `G6:CH13` reading each
  class tab's F and G, English minimum F5 `=50 - MAX(20, E6)`, ...).
- Master template shared with the domain as Viewer.

## Known issues and follow-ups

- **Timetables generated before PR #2 still have the old copy and Doc
  tools.** Anyone signed into several accounts gets PERMISSION_DENIED in
  them until the new `template/TimetableTools.gs` is pasted into that
  timetable's own Apps Script project (or the timetable is regenerated).
- **Timetables generated before the OVERVIEW fix** keep the old OVERVIEW
  colouring (fixed English 30%). The same manual edits fix them; see
  `template/SHEET_CHANGES.md`, "OVERVIEW: compare each class with its own
  minimums".
- **No browser tests in the repo for the wizard page.** During development
  the page was checked with Playwright scripts kept outside the repo
  (mocking `google.script.run`, with Tailwind built locally); they were not
  committed. If the wizard changes a lot, consider adding a small Playwright
  suite under `tests/browser/`, kept separate from `npm test` so the
  dependency-free tests stay that way.
- **Deploying is by copy and paste.** `clasp` (pushing from the repo to Apps
  Script) is an optional, not-started item in `PLAN.md` Phase 0. It would
  remove the paste step for the wizard and the template, but not for
  timetables already generated.
- Parked ideas (`PLAN.md`, "Future ideas"): a sidebar fill tool (pick a
  subject, fill the selected cells), and specialist working days. A sidebar
  would need to avoid `google.script.run` problems with several accounts, so
  think about that before starting it.

## Things learned the hard way

- `google.script.run` from a dialog or sidebar in a spreadsheet runs as the
  browser's *first* signed-in account. Incognito with one account works,
  which is the quickest way to confirm this when someone reports
  "PERMISSION_DENIED" or "error occurred while reading from storage".
- Console messages about `csp.withgoogle.com` (ERR_SSL_PROTOCOL_ERROR) or
  `chrome-extension://invalid` are noise from the school network and Docs
  Offline, not script errors. The real error is under Extensions > Apps
  Script > Executions.
- Tabs copied with the Sheets API `duplicateSheet` can't be found through
  SpreadsheetApp later in the same run, so conditional formatting goes on
  the first tab of each cycle before the copies are made.
- English and Mathematics minimums depend on each other (together at least
  50%, Mathematics at least 20%), and Religious Education's on the grade (10% K-2, 12%
  Years 3-6), so any compliance colouring has to read the class tab's own
  F and G columns, not fixed numbers.

## Where to look

| Need | File |
|---|---|
| Why something is the way it is | `PLAN.md` (each item explains the change and the reason) |
| What was changed in the spreadsheet by hand, and how to test it | `template/SHEET_CHANGES.md` |
| Deploying, first-run permissions, testing against a template copy | `README.md` |
| Generating, editing an existing timetable, drafts | `wizard/Code.gs` (`generateTimetable`, `updateTimetable`, `saveDraft`) |
| The wizard UI | `wizard/Index.html` (`App`; steps Core Schedule, Classes, Activities, Specialists, Summary) |
| Sync, lock, clash formatting | `template/TemplateSync.gs` |
| Checks, copy helpers, Google Docs | `template/TimetableTools.gs` |
