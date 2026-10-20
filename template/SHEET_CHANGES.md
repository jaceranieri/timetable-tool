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

- [ ] **Load test:** first run took 504s (too slow; see PLAN step 17). in the wizard project, add a script file `LoadTest` and paste `wizard/LoadTest.gs` into it. Choose `runLoadTest` in the toolbar and click Run. The log shows how long 28 classes and 6 specialists took (the limit is 360s). Note the time for step 28, delete the generated file, then delete the `LoadTest` file again.
- [x] Generate a normal timetable from the wizard and check:
  - Class tabs look as before: headers, times, grey break rows, dropdowns, subject colours, compliance table.
  - Two classes with the same cycle length have identical grids; a class with a different cycle has its own layout.
  - Specialist tabs: class colours, grey break rows, red clashes between two specialists.
  - The file name has no `INCOMPLETE - ` prefix.
- [x] Force an error (for example, set the `TEMPLATE_ID` script property to `x`) and check the wizard shows the error panel with **Try again** instead of a pop-up. Then set it back.
- [x] Sync and lock still work as in Phase 2.
- [x] **Load test with timings:** 429s. Copying class tabs (244s) and writing their details (117s) were 85% of the time.
- [ ] **Load test with the Sheets API:**
  1. In the wizard project, click **Services** (the + next to it in the left sidebar), choose **Google Sheets API**, keep the identifier `Sheets`, and click **Add**.
  2. Paste the updated `wizard/Code.gs` (and keep `LoadTest.gs` for now).
  3. Run `runLoadTest` and copy the whole log. Copies should now appear as lines like "Copy Kindergarten Red, … (one request)".
  4. Generate a normal timetable from the wizard and check the tabs are in the same order as the classes in the wizard, and each copied class shows its own teacher, name, grade and cycle in `C27:C31`.
