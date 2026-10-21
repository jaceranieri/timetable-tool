# Sheet changes to apply in the template

Changes that live in the template spreadsheet itself rather than in the
scripts. Apply them to the **test** template first, generate a timetable to
check, then repeat on the live template. Step numbers match `PLAN.md`.

## Phase 1 (applied and tested)

### Step 3: Amadeus check (TEMPLATE tab)

The wizard now writes `B31` = "Amadeus" and `C31` = TRUE (Years 3–6) or
FALSE (K–2) on every class tab. In the template, `C31` is blank, so `C18`
shows "N/A" there.

- [x] **`B31`:** optionally type `Amadeus` and copy the formatting from `B30`, so the template shows the row. Leave `C31` blank.
- [x] **`C18`:**
  ```
  =IF(C31<>TRUE, "N/A",
    AND(
      ROUND(E11,1) >= ROUND(F11,1), ROUND(E11,1) <= ROUND(G11,1),
      C21 > 0,
      ROUND(E21,1) >= ROUND(F21,1), ROUND(E21,1) <= ROUND(G21,1)
    )
  )
  ```
  Creative Arts (row 11) within range, Amadeus Tutorials (row 21) present and within range.
- [x] **`C16`:** add the Amadeus condition:
  ```
  =AND(
    C15=0,
    ARRAYFORMULA(
      (ROUND(E5:E12,1) >= ROUND(F5:F12,1)) * (ROUND(E5:E12,1) <= ROUND(G5:G12,1))
    ),
    OR(C18="N/A", C18=TRUE)
  )
  ```
- [x] **Conditional formatting on `C18`:** add "Text is exactly" `N/A` with grey text and no fill, alongside the existing TRUE/FALSE rules.
- [x] Optional: on `B20:G21`, add a custom-formula rule `=$C$18="N/A"` with grey text, placed above the existing `=$C$18=FALSE` rule, so the Amadeus block looks inactive for K–2.
- [x] Check `B21` is exactly `Amadeus - Tutorials`. The wizard's Amadeus activity uses that name, and it can no longer be renamed or removed in the wizard.

### Step 4: Unallocated blocks and specialist hours

- [x] **TEMPLATE `C15`:** check the whole grid, not just columns J–Z:
  ```
  =SUM(
    ARRAYFORMULA(
      IF(
        (I5:I <> "") * (J4:AAK4 <> ""),
        IF(ISBLANK(J5:AAK), 1, 0),
        0
      )
    )
  )
  ```
- [x] **SPEC_Template `D5`:** `=SUM(D6:D)`. It currently sums `D6:D13`, which misses classes after the eighth.

### Step 7: OVERVIEW tab

- [x] **`G5`:** read the class list from DATA (there is no CLASSES tab):
  ```
  =IFERROR(TRANSPOSE(FILTER(DATA!F2:F, DATA!F2:F <> "")), "")
  ```
- [x] **`G4`:** widen from `G5:Z5` (20 classes) to `G5:CH5` (80 classes), matching `G17`:
  ```
  =IFERROR(BYCOL(G5:CH5, LAMBDA(sheetName, IF(sheetName<>"", INDIRECT("'" & sheetName & "'!C29"), ""))), "")
  ```
- [x] **`E6:E13`:** widen from `G:AK` to `G:CH`, for example `E6` = `=IFERROR(AVERAGE(G6:CH6),"n/a")`, then fill down to `E13`.
- [x] **Unhide** the OVERVIEW tab.

## After applying Phase 1 (done)

- [x] Paste the updated `template/TemplateSync.gs` into the template's script project (Extensions → Apps Script).
- [x] Paste the updated `wizard/Code.gs` and `wizard/Index.html` into the wizard project.
- [x] Generate a test timetable with at least one Year 3–6 class, one K–2 class and two specialists, and check:
  - K–2 class tabs show `C18` = N/A; Years 3–6 show TRUE/FALSE.
  - Recess and lunch rows on specialist tabs are grey, not red.
  - OVERVIEW lists every class.
  - The specialist KLA dropdown in the wizard no longer offers Non-Curricular.

## Phase 2 (applied and tested)

### Step 9: Purge leftovers

Nothing in the scripts or the remaining tabs reads any of these.

- [x] **Delete tabs:** HELPER, Example, "Amadeus + Creative Arts Advice" (right-click the tab → Delete; show hidden tabs first via View → Hidden sheets).
- [x] **Delete named ranges** that point to `#REF!` (Data → Named ranges): `SchoolData`, `ExtraCurriculars`, `ReferenceList`.
- [x] **SPEC_Template:** remove all conditional formatting (select all → Format → Conditional formatting → delete each rule). The script already deletes these rules on every generated specialist tab.
- [x] **TEMPLATE `B19:B21`:** delete the "Text is exactly" rules for Break, English, Mathematics, Religion, Science, HSIE, PDHPE and Creative Arts. Those cells only ever hold the Amadeus labels. Keep the custom-formula rule `=AND(B19<>"", ISNA(MATCH(B19, $B$5:$B$11, 0)))` if it gives the Amadeus labels their colour.
- [x] **TEMPLATE `A12:B12`:** delete the "Text is exactly Break" rule. Keep the custom-formula `ISNA(MATCH(…))` rule if it gives the Non-Curricular label its colour.
- [x] **TEMPLATE `B32`:** delete the empty "Resources" heading (and any formatting under it).
- [x] **DATA:** nothing to do. Column N (ActivityList) is no longer written by the wizard.

### After applying Phase 2 (done)

- [x] Paste the updated `template/TemplateSync.gs` into the template's script project.
- [x] Paste the updated `wizard/Code.gs` and `wizard/Index.html` into the wizard project.
- [x] Generate a test timetable and check:
  - Class tabs still show subject colours, and the compliance table (rows 5–21) looks as before.
  - Specialist tabs show class colours and grey break rows.
  - Sync, then Lock Synced Cells, then delete one specialist lesson and run both again: the deleted slot is cleared and no longer locked.

## Phase 3 (tested)

No sheet changes. Paste the updated `wizard/Code.gs`, `wizard/Index.html`
and `template/TemplateSync.gs`, then check:

- [x] The wizard flags, on the field itself:
  - a duplicate class name in another grade (the wizard opens that grade's tab),
  - a class called `Data`, `SPEC_Test` or `Year 3/4`, or containing an apostrophe,
  - a specialist called `Template` or two specialists with the same name,
  - a start, end or break time not on a 5-minute mark (e.g. 10:32),
  - two overlapping breaks.
- [x] A normal setup generates as before, and DATA shows times such as `08:30` as typed (plain text).
- [x] The hidden SPEC_Template tab gets no clash formatting, and "Sync Specialist Schedules" doesn't list it.

## Phase 4

No sheet changes. Paste the updated `wizard/Code.gs`, `wizard/Index.html`
and `template/TemplateSync.gs`.

- [x] **Load test:** first run took 504s (too slow; see PLAN step 17). In the wizard project, add a script file `LoadTest` and paste `wizard/LoadTest.gs` into it. Choose `runLoadTest` in the toolbar and click Run. The log shows how long 28 classes and 6 specialists took (the limit is 360s). Note the time for step 28, delete the generated file, then delete the `LoadTest` file again.
- [x] Generate a normal timetable from the wizard and check:
  - Class tabs look as before: headers, times, grey break rows, dropdowns, subject colours, compliance table.
  - Two classes with the same cycle length have identical grids; a class with a different cycle has its own layout.
  - Specialist tabs: class colours, grey break rows, red clashes between two specialists.
  - The file name has no `INCOMPLETE - ` prefix.
- [x] Force an error (for example, set the `TEMPLATE_ID` script property to `x`) and check the wizard shows the error panel with **Try again** instead of a pop-up. Then set it back.
- [x] Sync and lock still work as in Phase 2.
- [x] **Load test with timings:** 429s. Copying class tabs (244s) and writing their details (117s) were 85% of the time.
- [x] **Load test with the Sheets API:** 93s. Afterwards delete the test timetables and the `LoadTest` file.
  1. In the wizard project, click **Services** (the + next to it in the left sidebar), choose **Google Sheets API**, keep the identifier `Sheets`, and click **Add**.
  2. Paste the updated `wizard/Code.gs` (and keep `LoadTest.gs` for now).
  3. Run `runLoadTest` and copy the whole log. Copies should now appear as lines like "Copy Kindergarten Red, … (one request)".
  4. Generate a normal timetable from the wizard and check the tabs are in the same order as the classes in the wizard, and each copied class shows its own teacher, name, grade and cycle in `C27:C31`.

### Live clash check across cycle lengths (PLAN step 18a)

- [ ] Paste the updated `template/TemplateSync.gs` into the **template's** script project, and the updated `wizard/Code.gs` into the wizard project.
- [ ] Generate a new timetable (timetables generated earlier keep their own copy of the old script and rules) with a 1-week and a 2-week specialist, then check:
  - 1-week specialist: a class in Term 1 Monday 9:00 (column J). 2-week specialist: the same class in Term 1 **Week B** Monday 9:00 (column P). Both cells turn red.
  - 1-week specialist: a class in Term 2 Monday 9:00 (column P). 2-week specialist: the same class in Term 1 Week B Monday 9:00 (also column P). Neither turns red.
  - Two specialists with the same cycle length and the same class in the same cell still turn red.

**Result:** different cycle lengths turned red, but two specialists with the same cycle length didn't. Cause: the second specialist with a cycle length is a copy made through the Sheets API, and tabs made that way can't be reached through SpreadsheetApp for the rest of the run. So the clash formatting skipped the copies, and the class tabs were counted into the wrong positions. Fixed by putting the clash rules on the first tab of each cycle before copying (the copies inherit them) and counting tab positions instead of reading them back.

- [x] Paste the updated `wizard/Code.gs` and `template/TemplateSync.gs` again, generate a **new** timetable, and repeat the three checks above, plus:
  - Tabs appear in the wizard's order: DATA, OVERVIEW, TEMPLATE, SPEC_Template, the specialists, then the classes grade by grade.
  - Each copied class shows its own details in `C27:C31`.
- [ ] Optional, on the earlier timetable: **Timetable Admin → Refresh Specialist Clash Formatting** should make the same-cycle clash turn red there too (the menu runs separately, so it sees every tab).

## Phase 5 and step 26a

No sheet changes. Paste the updated `wizard/Code.gs` and `wizard/Index.html`.

- [x] The wizard loads and looks the same as before (React, Babel and Tailwind now load pinned versions; check the page is styled, i.e. Tailwind loaded).
- [x] Step 1 asks for the **School Name** and won't continue without it.
- [x] Step 3: give a grade 4 or more classes. All of them are visible without a small scroll box, and the tab shows the class count.
- [x] Leave one class's teacher blank, switch to another grade's tab and click Continue: the wizard opens the right tab, scrolls to the class, highlights it and puts the cursor in the empty field. The tab shows "1 to fix" until it's filled in.
- [x] The summary step mentions the school name.
- [x] Generate: the timetable is saved in **My Drive › Timetables** (created the first time) and its name includes the school name. A second timetable goes into the same folder.

## Phase 6

No sheet changes, but the template's script project gets a **new file**.

- [ ] In the template's script project (Extensions → Apps Script): paste the updated `template/TemplateSync.gs`, then add a script file named `TimetableTools` and paste `template/TimetableTools.gs` into it.
- [ ] Paste the updated `wizard/Code.gs` and `wizard/Index.html` into the wizard project.
- [ ] Generate a new timetable and check the **Start here** tab is first:
  - The five stages with a status for each, and "Not run yet" under Last run for the sync and the lock.
  - Every specialist and class tab listed, each name a working link, with lessons entered / empty slots / conflicts / compliant filling in.
- [ ] Reload the timetable; the **Timetable Admin** menu is in stage order. Lock before syncing warns first. After a sync and a lock, Start here shows when they ran, by whom and the result.
- [ ] **Check this timetable** on a class tab (partly filled), on a specialist tab, and on DATA (asks for a timetable tab). **Check all class timetables** lists every class.
- [ ] **Copy Week A to the other weeks** on a 2-week class: Yes fills only empty slots; No also replaces entries; breaks and synced specialist lessons stay as they are. **Copy Term 1 to Terms 2-4** on a 1-week class.
- [ ] **Print this timetable**: enter `1A`; the Print view tab shows that week with merged blocks and times, and **Download PDF** gives a landscape A4 page. **Print the compliance overview** downloads OVERVIEW.

### Phase 6 changes: clearer copy wording, Google Docs instead of PDFs

- [ ] In the **template's** script project: click **Services** (+), choose **Drive API**, keep the identifier `Drive` (version v3) and click **Add**.
- [ ] Paste the updated `template/TemplateSync.gs` and `template/TimetableTools.gs` into the template's script project, and the updated `wizard/Code.gs` and `wizard/Index.html` into the wizard project.
- [ ] Generate a new timetable, then check:
  - **Copy Week A into the other weeks**: a window explains what will be copied, with the buttons **Fill empty slots only**, **Fill empty slots and replace what's there** and **Cancel**; the result shows in the same window.
  - **Make a Google Doc of this timetable** on a class tab (`1A`): the Doc opens from the link, is landscape, each lesson is one merged block with its times, and subjects have the same colours as the class tab. Try it on a specialist tab too (class colours).
  - **Make a Google Doc of the compliance summary**: one row per class, out-of-range percentages in red.
  - Both Docs are saved in the same folder as the timetable.

### Phase 6 changes: Google Docs details, All terms, Start here list

No sheet changes. Paste the updated `template/TemplateSync.gs` and `template/TimetableTools.gs` into the template's script project, and the updated `wizard/Code.gs` into the wizard project. Remove `LoadTest.gs` from the wizard project if it's there (it's for development only).

- [ ] Generate a new timetable. It's named `<school name> - Timetable - <date>`.
- [ ] **Start here** has a **Google Docs** section below the class list, reading "None yet…".
- [ ] **Make a Google Doc of this timetable** on a class tab: the window has a drop-down of weeks (terms, for a 1-week class) plus **All terms**.
  - One week: the Doc starts with a details table (class, teacher, grade, cycle, school, school day, breaks, generated date and by whom), then the week.
  - **All terms**: details once at the top, then one page per week (per term for a 1-week class), each with its heading.
- [ ] On a specialist tab the details show the specialist, KLA and cycle.
- [ ] **Make a Google Doc of the compliance summary**: details show how many classes are compliant and when it was made.
- [ ] Each Doc is added to the **Google Docs** list on Start here (a link, what it was made from, when and by whom), newest at the bottom.

### Phase 6 changes: Doc layout, specialist names, compliance colours, Start here style

No sheet changes. Paste the updated `template/TimetableTools.gs` into the template's script project, and the updated `wizard/Code.gs` into the wizard project.

- [ ] Generate a new timetable. **Start here** has no gridlines, a navy title band that stays at the top when scrolling, section headings in navy, light grey table headings and thin dividers. In the class list, Compliant shows green Yes / red No, Conflicts above 0 are red, and 0 empty slots are green. The sync stage's status turns red when there are conflicts and green when there are none.
- [ ] Sync a specialist, then **Make a Google Doc of this timetable** on a class tab:
  - The tables stretch across the page.
  - Every lesson shows its start and end times, including 5-minute ones (e.g. the last slot of the day).
  - Synced specialist lessons show the specialist's name under the KLA.
- [ ] **Make a Google Doc of the compliance summary**: percentages within their range are green, outside it red; the table stretches across the page.

## Phase 7

No sheet changes. Paste the updated `wizard/Code.gs` and `wizard/Index.html` into the wizard project and redeploy (Deploy → Manage deployments → edit → New version).

- [ ] **Drafts:** fill in Step 1 and part of Step 2; "Draft saved" pops up at the top of the screen. Reload the page: "Continue where you left off?" appears. **Resume** goes back to the same step with everything filled in; reload again and **Start new** gives an empty wizard.
- [ ] **Problem list:** click Continue on Step 1 with no school name. A red box lists the problem; clicking it puts the cursor in the field, and the box goes once it's fixed. On Step 3, make two classes in different grades share a name: both are listed, and clicking one opens its grade tab.
- [ ] **Teachers:** leave a class's teacher blank, and type TBC for another: both are listed as problems and Continue is blocked. Clear both fields of a third class: it says "Blank: skipped" and isn't counted on the summary.
- [ ] **Colours:** the default classes have matching colours (Kindergarten Blue is blue, Year 1 Gold is yellow...). Rename a class "Year 3 Purple": it turns purple. Pick a colour for it, then rename it again: your colour stays.
- [ ] **Paste:** in a spreadsheet, select class names and teachers (two columns) and copy. On Step 3, open the **...** at the end of the grade tabs, **Paste from a spreadsheet**, and paste. It shows how many it found; **Use these N classes** replaces that grade's classes.
- [ ] **Cycle warning:** give a specialist a 2-week cycle while a grade has 1 or 3 weeks: a warning shows under the specialist and on the summary.
- [ ] **Header:** shows "Step N of 4", the step name, a line about it and a figure (teaching time, classes...), with a four-part progress bar. Click a finished step's name to go back to it.
- [ ] **Summary:** info cards at the top (class tabs, specialist tabs, teaching time, time to generate), then a card for each part with Edit links back to its step.
- [ ] **Setup file:** the **Save / Load** button at the top right of the wizard has **Save setup to a file** (downloads `<school> - timetable setup.json`) and **Load a saved setup**. Loading it in a new wizard fills every step the same way; another file gets "isn't a timetable setup file" at the top of the screen.
- [ ] **Generating:** a progress bar; the line under it alternates between the time reminder and teacher in-jokes every few seconds; **View the log** shows the steps. Compare the real time with the estimate.
- [ ] **After generating:** **Back to the wizard** returns to the summary with the setup kept; **Start a new timetable** gives an empty wizard.
- [ ] **Keyboard:** on Step 3, Tab to the grade tabs and use the left/right arrow keys; open a class colour and close it with Escape; open a **...** menu and use the up/down arrows and Enter.
