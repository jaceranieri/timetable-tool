# CLAUDE.md

Guidance for working in this repository. `HANDOFF.md` has the current state
and open items; `README.md` has setup and deployment; `PLAN.md` has the
decisions and the history of every change.

## What this is

A Google Apps Script tool for Sydney Catholic Schools primary schools. Two
Apps Script projects live here, and nothing in the repo runs in Google by
itself: the person maintaining it pastes the files into the Apps Script
editor and tests in Google Sheets.

| Folder | Apps Script project | Notes |
|---|---|---|
| `wizard/` | Standalone web app (`Code.gs` server, `Index.html` page) | Copies the master template and builds the class and specialist tabs. Also edits an existing timetable. Runs as the user accessing it, domain only. |
| `template/` | Bound to the master template spreadsheet (`TemplateSync.gs`, `TimetableTools.gs`) | The **Timetable Admin** menu: sync, lock, clash formatting, checks, copy helpers, Google Docs. |
| `shared/` | Copied into both | Edit here, then `npm run build`. |
| `tests/` | Node only | Runs the scripts against an in-memory fake of Google Sheets. |

## Commands

```
npm test         # checks shared copies are up to date, then runs node --test
npm run build    # copies shared/ files into the scripts between marker lines
```

Node 18 or later, no dependencies. Run `npm test` before every commit.

## Rules that are easy to break

- **No `google.script.run` in template dialogs.** It fails with
  PERMISSION_DENIED for anyone signed into more than one Google account (most
  teachers). Ask questions with `ui.alert` / `ui.prompt` from the menu
  function itself; HtmlService windows may only *display* results (links,
  tables). The wizard is a web app and isn't affected.
- **Don't edit between `// ---- shared/...` marker lines.** The build
  overwrites them and `npm test` fails while a copy is stale. A shared name
  defined twice also fails the tests.
- **Tabs created through the Sheets API are invisible to SpreadsheetApp for
  the rest of that run.** The generator tracks tabs and positions itself
  instead of reading them back.
- **Every generated timetable carries its own copy of the template script.**
  A change to `template/` only reaches new timetables; existing ones need the
  file pasted into their own Apps Script project. Say so in rollout notes.
- **Teacher names can't be blank or TBC** (compliance): `TEACHER_PLACEHOLDER`
  in `shared/setupRules.js` is checked on the page and on the server.
- **The decisions table in `PLAN.md` is settled** (full year of 4 terms,
  5-minute rows, cycles of 1–10 weeks, Amadeus for Years 3–6, no composite
  classes...). Don't reopen them without being asked.
- Changes made in the template spreadsheet itself (formulas, conditional
  formatting) can't be done from code: write numbered steps and a check list
  in `template/SHEET_CHANGES.md` for the maintainer to apply by hand.

## Sheet layout (class and specialist tabs)

- Times in column I; the grid starts at J5 (`FIRST_DATA_ROW` 5,
  `FIRST_GRID_COLUMN` 10). One row per 5 minutes.
- Row 1 holds "Term X - Week Y"; row 4 the day names. Each week is 6 columns:
  5 days and a spacer.
- Class tab compliance table B5:G21: E is the %, F the minimum, G the
  maximum. C16 is overall compliance, C18 Amadeus, C27 the teacher.
  TEMPLATE!F5 (English minimum) is `=50 - MAX(20, E6)` and F6 (Mathematics)
  `=MAX(20, 50 - E5)`: together at least 50%, Mathematics at least 20%.
- Specialist tabs are named `SPEC_<name>`. Synced lessons carry a note
  starting `[SPEC_SYNC]`, conflicts `[SPEC_CONFLICT]`.
- OVERVIEW colours each class against its own tab's F and G columns.

## Code style

- Match the surrounding code: plain functions, `const`/arrow functions (V8
  runtime), a short JSDoc comment above each function saying what it does and
  why, in plain English. No frameworks on the server.
- The wizard page is React 18, Babel and Tailwind loaded from CDNs, with no
  build step, in the single `Index.html`. Font is Figtree.
- Keep spreadsheet calls few: read and write whole ranges, not cell by cell
  (the tests count calls).
- User-facing text is Australian English ("colour", "organise"), short and
  friendly, written for teachers rather than developers. Explain what will
  happen and what to do next; never show a raw error code alone.

## Tests

- `tests/helpers/fakeSpreadsheet.js` is a hand-written stand-in for
  SpreadsheetApp, the Sheets advanced service, DriveApp, Drive, DocumentApp,
  CacheService and PropertiesService. When code uses a method it doesn't
  have, add the method to the fake (small, faithful to Google's behaviour)
  rather than working around it.
- `world.answers` queues alert answers (default YES), `world.promptAnswers`
  queues prompt text (`null` = Cancel), `world.dialogs` collects HtmlService
  windows, `world.alerts` the alert messages.
- `tests/fixtures/Code.before-phase4.gs` is the old generator, kept to check
  the new one builds the same timetable.
- There are no browser tests for `wizard/Index.html` in the repo (see
  `HANDOFF.md`).

## Working with the maintainer

- Work happens phase by phase: implement, run the tests, push, then the
  maintainer pastes the files into Apps Script, tests in Sheets and reports
  back ("all passed" means tested in Google).
- End each change with exactly what to paste where, and anything to do by hand
  in the template.
- Open a PR only when asked. After a PR is merged, start follow-up work on a
  fresh branch from `main`.
