# Implementation plan

Prioritised plan for fixing and improving the timetable tool. Work top to
bottom; tick items off as they are completed. Sizes: **S** = under an hour,
**M** = a few hours, **L** = a day or more.

## Decisions

| Topic | Decision |
|---|---|
| Web app deployment | Execute as **user accessing the web app**, access limited to the Sydney Catholic Schools domain |
| Terms | Every timetable covers the **full year (4 terms)**. This is a compliance requirement. |
| Row size | Fixed **5-minute rows** so timetables are consistent across schools |
| Bell times | Same for all grades and all days |
| Amadeus | Applies to **Years 3–6** in all schools. Set automatically, no wizard toggle. |
| Amadeus compliance (`C18`) | Creative Arts within range **and** Amadeus Tutorials present and within range |
| Overall compliance (`C16`) | Requires `C18` to be compliant when Amadeus applies |
| Cycle length | 1–**10** weeks |
| Composite classes | Not supported |
| Fixed per-class events (assembly, library…) | Entered by class teachers in the sheet, not in the wizard |
| Specialist class lists | Specialists see every class, so no per-specialist class list |
| Teacher emails / per-tab protection | Not wanted |
| Spaces as resources (hall, library…) | Not wanted |
| OVERVIEW tab | Keep and fix |
| HELPER, Example, Advice tabs | Purge |

## Phase 0: Setup

- [x] **1. Put the code in the repo (S).** `wizard/` holds the web app (`Code.gs`, `Index.html`); `template/` holds the script bound to the master template (`TemplateSync.gs`). Both have an `appsscript.json` manifest.
- [x] **2. Separate test and live templates (S).** `TEMPLATE_ID` is read from Script Properties, falling back to the live template.
- [ ] Optional: set up `clasp` to push from the repo to Apps Script.

## Phase 1: Correct results

- [x] **3. Amadeus check (S).**
  - The script writes `B31` "Amadeus" and `C31` TRUE (Years 3–6) or FALSE (K–2) on each class tab.
  - New `C18`: `"N/A"` when `C31` is FALSE; otherwise TRUE only if Creative Arts (`E11`) is within `F11:G11`, Amadeus Tutorials minutes (`C21`) are above 0, and `E21` is within `F21:G21`.
  - `C16` also requires `C18` to be TRUE or `"N/A"`.
  - Add a grey conditional-format rule on `C18` for `"N/A"`.
  - The wizard's "Amadeus - Tutorials" activity can't be renamed or removed, and the server always adds it, because `C21` counts that exact name.
- [x] **4. Fix `C15` and `SPEC_Template!D5` (S).** `C15` checks the full `J:AAK` grid, not just `J:Z`. `D5` sums the whole class list, not just `D6:D13`.
- [x] **5. Break rows shown as clashes (S).** Create the `BreakList` named range in `populateDataSheet`. Keep one version of `applySpecialistClashFormatting` (TemplateSync's, which handles breaks) and copy it into `Code.gs`.
- [x] **6. Specialist sync correctness (M).**
  - Clear old `[SPEC_SYNC]` and `CONFLICT:` cells before re-syncing.
  - Flag a conflict when a class teacher's own entry is already in the slot.
  - Warn when a class's cycle isn't a multiple of the specialist's cycle.
  - A clashing class entry is kept in the conflict cell's note and put back on the next sync once the clash is gone.
  - "Lock Synced Cells" rebuilds its own locks each run, so moved or deleted lessons are unlocked.
- [x] **7. OVERVIEW (S).** `G5` reads the class list from `DATA!F2:F`; `G4`, `E6:E13` and `G17` cover the same number of classes; unhide the tab.
- [x] **8. Remove Non-Curricular from the specialist KLA dropdown (S).**

## Phase 2: Template cleanup

- [x] **9. Purge leftovers (S).**
  - Tabs: HELPER, Example, "Amadeus + Creative Arts Advice".
  - Named ranges pointing to `#REF!`: `SchoolData`, `ExtraCurriculars`, `ReferenceList`.
  - All conditional formatting on `SPEC_Template` (the script deletes it anyway).
  - Stale TEMPLATE colour rules on `B19:B21` and `A12:B12` (the "Text is exactly" rules; keep the label-colour formula rules); the empty `B32` "Resources" heading.
  - DATA column N (ActivityList).
  - TemplateSync's old `#E8F0FE` colour-marker code and the "unexpected samples" diagnostics.
  - Dead code in `Index.html` (`errors.general`, `metadata.generatedAt`, `isSubmitting`, unused `CardTitle`/`CardDescription`).

## Phase 3: Stop silent data loss

- [x] **10. Name rules in the wizard and the server (M).** Unique class and specialist names (ignoring case and the characters replaced in tab names); reserved names blocked (DATA, TEMPLATE, OVERVIEW…); no `SPEC_` prefix; no characters that tab names can't hold (`/ \ ? * [ ] :`), because OVERVIEW looks tabs up by class name. The server returns a clear error instead of skipping. Also: names can't start with `=`, class names can't contain `'` or match a break name, and specialist names can't be "Template". The setup is checked before the template is copied.
- [x] **11. Store DATA as plain text (S)** so `3/4` isn't turned into a date and `=…` isn't turned into a formula.
- [x] **12. Detect the template tab reliably (S).** Compare `SPEC_Template` without case, or tag generated tabs with developer metadata.
- [x] **13. Check inputs on the server (S).** Times on the 5-minute grid; breaks inside the day and not overlapping; cycle length 1–10; at least one class.

## Phase 4: Generator rewrite

- [x] **14. Local tests first (M).** `npm test` runs both scripts against an in-memory fake spreadsheet (`tests/`): setup rules, grid layout, full generation, sync, lock, and a comparison with the previous generator.
- [x] **15. Rebuild tab generation (L).** Build one finished class tab per cycle length and copy it for each class; batched writes; conditional-formatting rules built once. Each step is written as "make sure this tab exists and is set up" so step 32 can reuse it.
- [x] **16. Clean up after a failed run (S).** The copy is named `INCOMPLETE - …` until generation finishes (this also covers timeouts, which can't run clean-up code); plain-English error messages; one in-app error panel with Try again.
- [x] **17. Load test (S).** 28 classes (3-week cycle) and 6 specialists: 504s before, **93s** after copying tabs through the Google Sheets API (target 120s). Remaining time: building the first tab of each cycle (45s), clash formatting (18s), setup (10s), copies (20s). Use about 1.5 minutes for a school this size in step 28.
- [x] **18a. Live clash check across cycle lengths (M).** The red clash rule on specialist tabs compared the same cell on every tab, which only lines up for equal cycle lengths: a weekly lesson against a week-B lesson wasn't flagged, and the same column in different terms was. The rule now works out, for each other specialist, the cells in the same real week, term, day and time (the same definition the sync uses). Equal cycles keep the simple same-cell formula. The rules are put on the first tab of each cycle before the copies are made, because tabs copied through the Sheets API can't be reached through SpreadsheetApp in the same run (that also put class tabs in the wrong order; positions are now counted, not read back).
- [x] **18. Faster sync (S).** Header rows are read in bulk in `buildScheduleColumnMap` and the clash formatting.

## Phase 5: Deployment polish

- [x] **19. "User accessing" settings (S).** Progress uses the per-user cache; timetables are saved in a "Timetables" folder in the user's My Drive (created if needed; `OUTPUT_FOLDER_NAME`); the file name includes the school name. No in-app note about Google's permission prompt: with "execute as user", Google asks when the wizard is first opened, before the wizard loads, so the note is in the README for the rollout message instead.
- [x] **20. Pin CDN versions (S).** React 18.3.1, ReactDOM 18.3.1, Babel standalone 7.26.10, Tailwind Play CDN 3.4.17.

## Phase 6: Guided workflow in the sheet

- [ ] **21. "Start here" tab (M).**
  - Stages in order:
    1. Specialists fill their tabs.
    2. Sync and resolve clashes.
    3. Lock specialist slots.
    4. Class teachers fill the rest.
    5. Compliance review.
  - Live status for each stage: specialist tabs filled, last sync time, open conflicts, locked or not, classes compliant.
  - Timetable Admin menu reordered to match. The order is guided and warned about, not enforced.
- [ ] **22. "Check my timetable" (M).** For one class: each KLA's minutes over or under, empty slots by term and week, Amadeus status, open conflicts. Whole-school version writes to OVERVIEW.
- [ ] **23. Copy helpers (M).** "Copy Week A → other weeks" and "Copy Term 1 → Terms 2–4". Confirmation first; never overwrite break rows, locked specialist slots or non-empty cells unless chosen.
- [ ] **24. Printable outputs (M).** A clean weekly timetable per class (consecutive same-subject slots merged into blocks) as PDF, and a one-page compliance summary per school. Needs the external-request permission.

## Phase 7: Wizard usability

- [ ] **25. Draft autosave (S)** with Resume and Start new; "Back to wizard" and "Start new" on the success screen.
- [ ] **26. Step 3 errors (S).** ~~Jump to the first grade tab with an error~~ (done in Phase 3); show a summary, ~~clear errors as fields are edited~~ (done in Phase 5). Allow "TBC" teachers; ignore fully blank rows.
- [x] **26a. Grade tabs in Step 3 (S).** A grade's classes sit in a panel with its own small scroll area (`max-h-[350px]`) and a thin scrollbar, so with four or more classes the last one is hidden below the fold; when it has an error the wizard reports it but you can't see where. Let the class list grow with the page (or make the scroll obvious), show "4 classes, 1 needs attention" on each tab, and scroll to and highlight the first class with an error. *Done: no inner scroll area; each tab shows its class count and "N to fix"; the panel says "4 classes · 1 needs attention"; Continue opens the tab, scrolls to and highlights the first class with a problem and puts the cursor in the field; a field's error clears as soon as it's edited.*
- [ ] **27. ~~School name field~~ (done in Phase 5) and specialist/class cycle warning (S).** (5-minute time steps, overlapping-break errors and the 10-week cycle cap were done in Phase 3.)
- [ ] **28. Preview before generating (S):** "14 class tabs, 3 specialist tabs, about 1 minute".
- [ ] **29. Paste from a spreadsheet (S):** paste class names and teachers in bulk.
- [ ] **30. Save and load the setup as a file (S).** Same format step 32 reads from DATA.
- [ ] **31. Accessibility pass (M)** and Edit links on the summary step.

## Phase 8: Edit setup after generating, and start next year

- [ ] **32. Open an existing file in the wizard (L).**
  - Paste the file's link; the wizard prefills from `RawData_JSON`.
  - Apply only the changes: add tabs; rename a class or specialist (updating DATA and specialist entries that use the old name); archive removed classes by hiding and renaming, never deleting.
  - Changes that rebuild a grid (bell times, a class's cycle length) are blocked or need explicit confirmation.
- [ ] **33. "Start next year from this year's setup" (S).** Fresh copy of the template, prefilled from the old file's setup.

## Phase 9: Maintainability

- [ ] **34. One source for shared helpers (M).** An Apps Script library, or a build step that copies them into both projects.

## Future ideas (parked)

- Sidebar fill tool: pick a subject and fill the selected cells.
- Specialist working days, once schools know them at timetabling time.
