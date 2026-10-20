/**
 * Specialist -> Class Timetable Sync
 * -----------------------------------
 * Paste this file into the Apps Script project BOUND to the master
 * TEMPLATE spreadsheet (the one the wizard duplicates via
 * DriveApp.getFileById(TEMPLATE_ID).makeCopy(...)). Because it's bound
 * directly to the template file, every copy the wizard generates for a
 * school will automatically inherit this script too, including the
 * "Timetable Admin" menu below - no extra setup needed per copy.
 *
 * What it does:
 *  - Reads every specialist's filled-in schedule (SPEC_ sheets).
 *  - Maps each entry onto the correct cell(s) on the matching class sheet,
 *    correctly handling specialists and classes that have different
 *    cycle lengths (see buildScheduleColumnMap / computeLcm below).
 *  - Writes the specialist's KLA into that class cell. "Lock Synced Cells"
 *    then locks those cells to the spreadsheet owner.
 *  - Each sync starts by removing the slots written by the previous sync,
 *    so lessons a specialist has since moved or deleted don't linger.
 *  - If two specialists, or a specialist and a class teacher's own entry,
 *    want the same class cell, flags it visibly in red instead of silently
 *    overwriting. The class teacher's entry is restored on the next sync
 *    once the clash is gone.
 *  - Warns (without blocking) when a class's cycle isn't a multiple of the
 *    specialist's cycle, because the class timetable then can't show every
 *    specialist week separately.
 */

const SYNC_NOTE_PREFIX = '[SPEC_SYNC]';
const CONFLICT_NOTE_PREFIX = '[SPEC_CONFLICT]';
const CONFLICT_VALUE_PREFIX = 'CONFLICT:';
const LOCK_DESCRIPTION = 'Locked Specialist Slot(s)';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Timetable Admin')
    .addItem('Sync Specialist Schedules', 'syncSpecialistSchedules')
    .addItem('Lock Synced Cells', 'lockSyncedSpecialistCells')
    .addItem('Refresh Specialist Clash Formatting', 'refreshSpecialistClashFormatting')
    .addToUi();
}

function refreshSpecialistClashFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  applySpecialistClashFormatting(ss);
  SpreadsheetApp.getUi().alert('Specialist clash formatting refreshed.');
}

function syncSpecialistSchedules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const confirmResponse = ui.alert(
    'Sync Specialist Schedules',
    "This will read every specialist's timetable and copy their subject into the matching class " +
      "timetable cells. It does NOT lock anything - run \"Lock Synced Cells\" afterwards once " +
      "you're happy with the result.\n\n" +
      'Before continuing, please check each specialist sheet for scheduling collisions (cells ' +
      'highlighted red) and resolve them - this sync will not stop for them.\n\n' +
      'Continue with the sync?',
    ui.ButtonSet.YES_NO
  );
  if (confirmResponse !== ui.Button.YES) return;

  const mappingRange = ss.getRangeByName('SpecialistMapping');
  if (!mappingRange) {
    ui.alert("Error: 'SpecialistMapping' range not found on DATA sheet.");
    return;
  }

  const specialistKlaMap = {};
  mappingRange.getValues().forEach(row => {
    if (row[0] && row[1]) {
      specialistKlaMap[String(row[0]).trim()] = String(row[1]).trim();
    }
  });

  const specialistSheets = ss.getSheets()
    .filter(sheet => sheet.getName().startsWith('SPEC_') && sheet.getName() !== 'SPEC_TEMPLATE');

  ss.toast('Reading ' + specialistSheets.length + ' specialist sheet(s)...', 'Sync', 5);

  // Each class sheet's schedule grid is read into memory ONCE (a single bulk
  // read), mutated there as specialists are processed, and written back with
  // a single bulk write per sheet at the end - instead of one read+write per
  // individual cell, which is what made earlier runs slow.
  const classCache = {};
  const conflicts = [];
  const cycleWarnings = [];
  const seenCyclePairs = {};
  let syncedCount = 0;

  function getClassInfo(className) {
    if (className in classCache) return classCache[className];
    const classSheet = ss.getSheetByName(className) || ss.getSheetByName(sanitizeSheetName(className));
    if (!classSheet || classSheet.getLastRow() < 5) {
      classCache[className] = null;
      return null;
    }
    const numRows = classSheet.getLastRow() - 4;
    const numCols = classSheet.getLastColumn() - 9;
    const region = classSheet.getRange(5, 10, numRows, numCols);
    const info = {
      sheet: classSheet,
      cycle: parseCycleFromClassSheet(classSheet),
      columnMap: buildScheduleColumnMap(classSheet, classSheet.getLastColumn()),
      times: classSheet.getRange(5, 9, numRows, 1).getValues(),
      values: region.getValues(),
      backgrounds: region.getBackgrounds(),
      fontColors: region.getFontColors(),
      fontWeights: region.getFontWeights(),
      notes: region.getNotes(),
      sources: Array.from({ length: numRows }, () => new Array(numCols).fill(null)),
      dirty: Array.from({ length: numRows }, () => new Array(numCols).fill(false)),
      released: Array.from({ length: numRows }, () => new Array(numCols).fill(false)),
      numRows: numRows,
      numCols: numCols
    };
    releasePreviousSyncCells(info, specialistKlaMap);
    // What each cell held before this sync wrote anything (a class
    // teacher's entry or blank), so a conflict can be undone later.
    info.original = info.values.map(row => row.slice());
    classCache[className] = info;
    return info;
  }

  // Load every class sheet up front, not only the ones a specialist still
  // mentions, so slots from the previous sync are removed everywhere.
  const classListRange = ss.getRangeByName('ClassList');
  if (classListRange) {
    classListRange.getValues().flat()
      .map(value => String(value).trim())
      .filter(Boolean)
      .forEach(getClassInfo);
  }

  specialistSheets.forEach((specialistSheet, specialistIndex) => {
    const metadataName = String(specialistSheet.getRange('B1').getDisplayValue()).trim();
    const tabName = specialistSheet.getName().replace('SPEC_', '').trim();
    const specialistName = metadataName || tabName;
    const assignedKla = specialistKlaMap[specialistName] || 'Specialist';
    const specialistCycle = parseCycleFromSpecialistSheet(specialistSheet);

    const lastRow = specialistSheet.getLastRow();
    const lastColumn = specialistSheet.getLastColumn();

    if (lastRow >= 5 && lastColumn >= 10) {
      const columnMap = buildScheduleColumnMap(specialistSheet, lastColumn);
      const timeSlots = specialistSheet.getRange(5, 9, lastRow - 4, 1).getValues();
      const schedule = specialistSheet.getRange(5, 10, lastRow - 4, lastColumn - 9).getValues();

      for (let columnIndex = 0; columnIndex < schedule[0].length; columnIndex++) {
        const column = columnIndex + 10;
        const colInfo = columnMap[column];
        if (!colInfo) continue; // spacer column or unrecognised header

        for (let rowIndex = 0; rowIndex < schedule.length; rowIndex++) {
          const assignedClass = schedule[rowIndex][columnIndex];
          if (!assignedClass) continue;

          const className = String(assignedClass).trim();
          const classInfo = getClassInfo(className);
          if (!classInfo) continue;

          const pairLcm = computeLcm(specialistCycle, classInfo.cycle);
          const cyclePairKey = specialistName + '|' + className;
          if (!seenCyclePairs[cyclePairKey]) {
            seenCyclePairs[cyclePairKey] = true;
            if (classInfo.cycle % specialistCycle !== 0) {
              cycleWarnings.push(
                specialistName + ' (' + specialistCycle + '-week cycle) and ' + className +
                ' (' + classInfo.cycle + '-week cycle)'
              );
            }
          }

          let matchingRowIndex = -1;
          for (let timeIndex = 0; timeIndex < classInfo.times.length; timeIndex++) {
            if (classInfo.times[timeIndex][0] === timeSlots[rowIndex][0]) {
              matchingRowIndex = timeIndex;
              break;
            }
          }
          if (matchingRowIndex === -1) continue;

          // Real-week-aware mapping: this specialist's week can land on
          // several different class-weeks across the shared repeat cycle
          // (pairLcm). Work out all of them and write to whichever actually
          // exist as columns on the class sheet for this term + day.
          const repeats = pairLcm / specialistCycle;
          const classWeeksToWrite = new Set();
          for (let i = 0; i < repeats; i++) {
            classWeeksToWrite.add((colInfo.week + i * specialistCycle) % classInfo.cycle);
          }

          classWeeksToWrite.forEach(classWeek => {
            let targetColumn = -1;
            for (const colKey in classInfo.columnMap) {
              const info = classInfo.columnMap[colKey];
              if (info.term === colInfo.term && info.week === classWeek && info.day === colInfo.day) {
                targetColumn = parseInt(colKey, 10);
                break;
              }
            }
            if (targetColumn === -1) return; // that week block doesn't exist on the class sheet

            const gridRow = matchingRowIndex;
            const gridCol = targetColumn - 10;
            const existing = classInfo.sources[gridRow][gridCol];

            if (existing && existing.value === assignedKla && existing.source === specialistName) {
              return; // identical repeat write (e.g. broadcast across weeks) - nothing new to do
            }
            if (existing && existing.value === 'CONFLICT') {
              return; // already flagged; leave the conflict visible for a human to resolve
            }

            // A class teacher's own entry (or a break label) already in the
            // slot. Matching the specialist's KLA is fine: the sync takes it over.
            const currentValue = String(classInfo.values[gridRow][gridCol] || '').trim();
            const classEntry = !existing && currentValue && currentValue !== assignedKla
              ? currentValue
              : '';

            if (existing || classEntry) {
              // Genuine conflict: two different sources/values want this cell.
              // Conflicts still get a loud literal red flag (rare/exceptional,
              // so it's fine for this to override any subject-colour formatting).
              // The note keeps the cell's original value so the next sync can
              // put it back once the clash is resolved; it is never locked.
              const otherSide = existing
                ? existing.source + ' (' + existing.value + ')'
                : 'class entry (' + classEntry + ')';
              classInfo.values[gridRow][gridCol] = CONFLICT_VALUE_PREFIX + ' ' +
                (existing ? existing.source : classEntry) + ' vs ' + specialistName;
              classInfo.backgrounds[gridRow][gridCol] = '#FF0000';
              classInfo.fontColors[gridRow][gridCol] = '#FFFFFF';
              classInfo.fontWeights[gridRow][gridCol] = 'bold';
              classInfo.notes[gridRow][gridCol] = CONFLICT_NOTE_PREFIX + ' ' +
                String(classInfo.original[gridRow][gridCol] || '');
              conflicts.push(
                classInfo.sheet.getName() + ' ' + classInfo.sheet.getRange(gridRow + 5, targetColumn).getA1Notation() +
                ': ' + otherSide + ' vs ' + specialistName + ' (' + assignedKla + ')'
              );
              classInfo.sources[gridRow][gridCol] = { value: 'CONFLICT', source: 'CONFLICT' };
              classInfo.dirty[gridRow][gridCol] = true;
              return;
            }

            // Successful, non-conflict write: only the value changes. Background/
            // font are explicitly reset to default (in case this cell was
            // previously flagged red as a conflict, or marked with the old
            // flat-blue marker from an earlier version of this script) so
            // your own subject-colour conditional formatting is free to
            // control how the cell looks. A note (invisible, no effect on
            // appearance) marks the cell as sync-managed so "Lock Synced
            // Cells" can find it later.
            classInfo.values[gridRow][gridCol] = assignedKla;
            classInfo.backgrounds[gridRow][gridCol] = null;
            classInfo.fontColors[gridRow][gridCol] = null;
            classInfo.fontWeights[gridRow][gridCol] = 'normal';
            classInfo.notes[gridRow][gridCol] = SYNC_NOTE_PREFIX + ' ' + specialistName;
            classInfo.sources[gridRow][gridCol] = { value: assignedKla, source: specialistName };
            classInfo.dirty[gridRow][gridCol] = true;
            syncedCount++;
          });
        }
      }
    }

    ss.toast(
      'Processed ' + (specialistIndex + 1) + ' of ' + specialistSheets.length + ' specialist sheet(s)...',
      'Sync',
      4
    );
  });

  // Slots from the previous sync that no specialist wrote again this time,
  // i.e. lessons that have since been moved or deleted.
  let removedCount = 0;
  Object.keys(classCache).forEach(className => {
    const info = classCache[className];
    if (!info) return;
    for (let r = 0; r < info.numRows; r++) {
      for (let c = 0; c < info.numCols; c++) {
        if (info.released[r][c] && !info.sources[r][c]) removedCount++;
      }
    }
  });

  ss.toast('Writing changes to ' + Object.keys(classCache).length + ' class timetable(s)...', 'Sync', 5);

  // Flush every touched class sheet in one bulk write per format type. If a
  // sheet has cells already locked from a previous "Lock Synced Cells" run
  // and the current user isn't the owner, the bulk write for that sheet will
  // fail - fall back to writing just the touched cells individually so the
  // rest of the sync isn't lost, and report which specific cells couldn't be
  // updated.
  const writeFailures = [];
  Object.keys(classCache).forEach(className => {
    const info = classCache[className];
    if (!info) return;
    const region = info.sheet.getRange(5, 10, info.numRows, info.numCols);
    try {
      region.setValues(info.values);
      region.setBackgrounds(info.backgrounds);
      region.setFontColors(info.fontColors);
      region.setFontWeights(info.fontWeights);
      region.setNotes(info.notes);
    } catch (bulkErr) {
      for (let r = 0; r < info.numRows; r++) {
        for (let c = 0; c < info.numCols; c++) {
          if (!info.dirty[r][c]) continue;
          const cell = info.sheet.getRange(r + 5, c + 10);
          try {
            cell.setValue(info.values[r][c])
              .setBackground(info.backgrounds[r][c])
              .setFontColor(info.fontColors[r][c])
              .setFontWeight(info.fontWeights[r][c])
              .setNote(info.notes[r][c]);
          } catch (cellErr) {
            writeFailures.push(info.sheet.getName() + ' ' + cell.getA1Notation());
          }
        }
      }
    }
  });

  let message = 'Sync Complete! Populated ' + syncedCount + ' specialist slot(s) across class timetables.\n\n' +
    'Nothing has been locked yet - run "Lock Synced Cells" from the menu when you\'re ready.';

  if (removedCount > 0) {
    message += '\n\nRemoved ' + removedCount + ' slot(s) left over from the previous sync ' +
      '(specialist lessons that have since been moved or deleted, or conflicts that are now resolved).';
  }

  if (conflicts.length > 0) {
    message += '\n\n' + conflicts.length + ' conflict(s) were found and flagged in red on the ' +
      'class sheets - please resolve these manually:\n' + conflicts.slice(0, 10).join('\n');
    if (conflicts.length > 10) message += '\n...and ' + (conflicts.length - 10) + ' more.';
    message += '\n\nTo resolve a clash with a class entry, move the specialist lesson or the ' +
      'class entry, then sync again. The class entry is put back automatically.';
  }

  if (cycleWarnings.length > 0) {
    message += '\n\n' + cycleWarnings.length + ' specialist/class pair(s) have cycles that ' +
      "don't line up (the class cycle isn't a multiple of the specialist's), so lessons from " +
      'different specialist weeks are shown together on the class timetable. Check these ' +
      'minutes carefully:\n' + cycleWarnings.slice(0, 10).join('\n');
    if (cycleWarnings.length > 10) message += '\n...and ' + (cycleWarnings.length - 10) + ' more.';
  }

  if (writeFailures.length > 0) {
    message += '\n\n' + writeFailures.length + ' cell(s) could not be updated because they were ' +
      "locked by a previous run and you aren't the spreadsheet owner:\n" +
      writeFailures.slice(0, 10).join('\n');
  }

  ui.alert(message);
}

/**
 * Undoes the previous sync on one class sheet's in-memory grid, before the
 * specialist timetables are applied again:
 *  - Synced slots still holding the specialist's KLA are emptied. If someone
 *    has since typed something else there, their value is kept.
 *  - Conflict cells get back the value they held before the clash, unless
 *    someone has already replaced the conflict text by hand.
 * Changed cells are marked in info.released and info.dirty.
 */
function releasePreviousSyncCells(info, specialistKlaMap) {
  for (let r = 0; r < info.numRows; r++) {
    for (let c = 0; c < info.numCols; c++) {
      const note = String(info.notes[r][c] || '');
      const value = String(info.values[r][c] || '').trim();
      const isSyncNote = note.indexOf(SYNC_NOTE_PREFIX) === 0;
      const isConflictNote = note.indexOf(CONFLICT_NOTE_PREFIX) === 0;
      const isConflictValue = value.indexOf(CONFLICT_VALUE_PREFIX) === 0;
      if (!isSyncNote && !isConflictNote && !isConflictValue) continue;

      if (isSyncNote) {
        const source = note.substring(SYNC_NOTE_PREFIX.length).trim();
        if (value === (specialistKlaMap[source] || 'Specialist')) {
          info.values[r][c] = '';
        }
      } else if (isConflictValue) {
        // Conflicts flagged by older versions of this script have no note,
        // so there is nothing to restore.
        info.values[r][c] = isConflictNote ? note.substring(CONFLICT_NOTE_PREFIX.length + 1) : '';
      }

      info.notes[r][c] = '';
      info.backgrounds[r][c] = null;
      info.fontColors[r][c] = null;
      info.fontWeights[r][c] = 'normal';
      info.dirty[r][c] = true;
      info.released[r][c] = true;
    }
  }
}

/**
 * Locks every specialist-synced cell (marked with an invisible note by
 * syncSpecialistSchedules, prefixed "[SPEC_SYNC]") on class timetables so
 * only the spreadsheet owner can edit them. Cells flagged as conflicts (red) are left
 * alone until a human resolves them and re-runs the sync.
 *
 * This is a separate step from syncSpecialistSchedules on purpose: locking
 * via Range.protect() is the slow part (Apps Script only supports creating
 * one Protection per range, so it's fundamentally one API call per range),
 * so it's kept out of the fast, frequently-rerun sync step. To cut down the
 * number of protect() calls, adjacent cells that need locking are merged
 * into the fewest rectangular ranges that cover them (mergeCellsIntoRectangles)
 * before protecting, instead of protecting one cell at a time.
 */
function lockSyncedSpecialistCells() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const owner = ss.getOwner();
  if (!owner) {
    ui.alert(
      'Cells were not locked: this spreadsheet has no individual owner (this usually happens ' +
      'with files stored in a Shared Drive). Move the file to My Drive first, then try again.'
    );
    return;
  }

  const confirmResponse = ui.alert(
    'Lock Synced Cells',
    'This will lock every specialist-synced cell on class timetables so only the ' +
      'spreadsheet owner (' + owner.getEmail() + ') can edit them. Cells flagged as conflicts ' +
      '(red) will be left editable until resolved.\n\nExisting specialist locks are rebuilt, so ' +
      'slots that are no longer synced are unlocked.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (confirmResponse !== ui.Button.YES) return;

  const classListRange = ss.getRangeByName('ClassList');
  if (!classListRange) {
    ui.alert("Error: 'ClassList' range not found on DATA sheet.");
    return;
  }
  const classNames = classListRange.getValues().flat().map(v => String(v).trim()).filter(Boolean);
  const runner = Session.getEffectiveUser();

  ss.toast('Scanning ' + classNames.length + ' class timetable(s)...', 'Lock Synced Cells', 5);

  let rangeCount = 0;
  let cellCount = 0;
  let scannedCellCount = 0;
  let matchedCellCount = 0;
  let alreadyProtectedCellCount = 0;
  let removedLockCount = 0;
  let sheetsScanned = 0;
  const failures = [];
  const unexpectedSamples = []; // {sheet, cell, text, note} for populated cells NOT matching the sync marker

  classNames.forEach((className, index) => {
    const classSheet = ss.getSheetByName(className) || ss.getSheetByName(sanitizeSheetName(className));
    if (!classSheet) return;

    const numRows = classSheet.getLastRow() - 4;
    const numCols = classSheet.getLastColumn() - 9;
    if (numRows <= 0 || numCols <= 0) return;

    sheetsScanned++;
    scannedCellCount += numRows * numCols;

    const region = classSheet.getRange(5, 10, numRows, numCols);
    const notes = region.getNotes();
    const backgrounds = region.getBackgrounds(); // fallback check for cells synced by an older version of this script
    const values = region.getValues();

    // Remove the locks made by the previous run, so slots that are no longer
    // synced (moved or deleted lessons) don't stay locked, then lock the
    // current synced cells again below.
    classSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(protection => {
      if (protection.getDescription() === LOCK_DESCRIPTION && protection.canEdit()) {
        protection.remove();
        removedLockCount++;
      }
    });

    // Don't re-protect cells already covered by another protection
    // (e.g. one a school admin added by hand).
    const alreadyProtected = Array.from({ length: numRows }, () => new Array(numCols).fill(false));
    classSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(protection => {
      const range = protection.getRange();
      if (!range) return;
      const rStart = range.getRow() - 5;
      const rEnd = rStart + range.getNumRows() - 1;
      const cStart = range.getColumn() - 10;
      const cEnd = cStart + range.getNumColumns() - 1;
      for (let r = Math.max(0, rStart); r <= Math.min(numRows - 1, rEnd); r++) {
        for (let c = Math.max(0, cStart); c <= Math.min(numCols - 1, cEnd); c++) {
          alreadyProtected[r][c] = true;
        }
      }
    });

    const needsLock = Array.from({ length: numRows }, (_, r) =>
      Array.from({ length: numCols }, (_, c) => {
        const hasNoteMarker = String(notes[r][c] || '').indexOf('[SPEC_SYNC]') === 0;
        const hasLegacyColourMarker = String(backgrounds[r][c]).toUpperCase() === '#E8F0FE';
        const isMarked = hasNoteMarker || hasLegacyColourMarker;
        if (isMarked) {
          matchedCellCount++;
          if (alreadyProtected[r][c]) alreadyProtectedCellCount++;
        } else if (values[r][c] && unexpectedSamples.length < 5) {
          unexpectedSamples.push({
            sheet: className,
            cell: classSheet.getRange(r + 5, c + 10).getA1Notation(),
            text: String(values[r][c]),
            note: notes[r][c] || '(none)'
          });
        }
        return isMarked && !alreadyProtected[r][c];
      })
    );

    const rectangles = mergeCellsIntoRectangles(needsLock, numRows, numCols);
    rangeCount += rectangles.length;

    rectangles.forEach(rect => {
      const range = classSheet.getRange(
        rect.rowStart + 5,
        rect.colStart + 10,
        rect.rowEnd - rect.rowStart + 1,
        rect.colEnd - rect.colStart + 1
      );
      try {
        range.setBackground(null); // clean up any leftover marker colour from an older script version
        const protection = range.protect().setDescription(LOCK_DESCRIPTION);
        protection.addEditor(owner);
        if (runner.getEmail() !== owner.getEmail()) {
          protection.removeEditors([runner]);
        }
        try { protection.setDomainEdit(false); } catch (permErr) {}
        cellCount += (rect.rowEnd - rect.rowStart + 1) * (rect.colEnd - rect.colStart + 1);
      } catch (protectErr) {
        failures.push(className + ' ' + range.getA1Notation());
      }
    });

    ss.toast('Locked ' + (index + 1) + ' of ' + classNames.length + ' class timetable(s)...', 'Lock Synced Cells', 4);
  });

  let message = 'Locking complete! Protected ' + cellCount + ' cell(s) using ' + rangeCount +
    ' range(s) (adjacent cells merged together to keep this fast).';

  message += '\n\nDiagnostics: scanned ' + sheetsScanned + ' class sheet(s), ' + scannedCellCount +
    ' schedule cell(s) total. Found ' + matchedCellCount + ' cell(s) marked as synced, ' +
    'of which ' + alreadyProtectedCellCount + ' are inside other protected ranges and were left ' +
    'as they are. Replaced ' + removedLockCount + ' lock(s) from the previous run.';

  if (matchedCellCount === 0) {
    message += "\n\nNo synced cells were found at all. This usually means either: (1) " +
      '"Sync Specialist Schedules" hasn\'t been run yet on this spreadsheet, or (2) it ran but ' +
      "didn't actually match any specialist entries to a class (check the sync's own completion " +
      'message for a "Populated 0 specialist slot(s)" result).';

    if (unexpectedSamples.length > 0) {
      message += '\n\nHere is what was actually found instead, on populated (non-blank) cells:\n' +
        unexpectedSamples.map(s =>
          s.sheet + ' ' + s.cell + ': text="' + s.text + '", note=' + s.note
        ).join('\n');
    }
  }

  if (failures.length > 0) {
    message += '\n\n' + failures.length + ' range(s) could not be locked:\n' + failures.slice(0, 10).join('\n');
  }

  ui.alert(message);
}

/**
 * Given a grid of booleans (needsLock[row][col]), returns the fewest
 * rectangular ranges that together cover every "true" cell and no "false"
 * cell. Not a globally-optimal rectangle cover (that's a much harder
 * problem) - it's a practical two-pass heuristic: merge each row's
 * consecutive true cells into horizontal strips, then merge vertically
 * adjacent strips that share the exact same column span into taller
 * rectangles. This handles the common real-world cases well (a specialist
 * teaching several consecutive periods with one class = a vertical merge;
 * a specialist teaching the same class across a whole day = a horizontal
 * merge), while staying simple and fast to run in Apps Script.
 */
function mergeCellsIntoRectangles(needsLock, numRows, numCols) {
  const stripsByRow = {};
  const allStrips = [];

  for (let r = 0; r < numRows; r++) {
    let c = 0;
    while (c < numCols) {
      if (!needsLock[r][c]) { c++; continue; }
      const colStart = c;
      while (c < numCols && needsLock[r][c]) c++;
      const strip = { row: r, colStart: colStart, colEnd: c - 1 };
      allStrips.push(strip);
      stripsByRow[r] = stripsByRow[r] || [];
      stripsByRow[r].push(strip);
    }
  }

  const consumed = new Set();
  const rectangles = [];

  allStrips.forEach(strip => {
    const key = strip.row + ':' + strip.colStart + ':' + strip.colEnd;
    if (consumed.has(key)) return;
    consumed.add(key);

    let rowEnd = strip.row;
    while (true) {
      const candidates = stripsByRow[rowEnd + 1] || [];
      const match = candidates.find(s => s.colStart === strip.colStart && s.colEnd === strip.colEnd);
      if (!match) break;
      const matchKey = match.row + ':' + match.colStart + ':' + match.colEnd;
      if (consumed.has(matchKey)) break;
      consumed.add(matchKey);
      rowEnd = match.row;
    }

    rectangles.push({ rowStart: strip.row, rowEnd: rowEnd, colStart: strip.colStart, colEnd: strip.colEnd });
  });

  return rectangles;
}

/**
 * Keep this function identical in wizard/Code.gs and
 * template/TemplateSync.gs: the wizard applies it when generating a
 * timetable, and the Timetable Admin menu re-applies it later.
 *
 * NOTE ON CYCLE MISMATCHES: this live conditional-formatting clash check
 * compares specialist sheets column-for-column, which is only guaranteed
 * correct when the specialist sheets being compared share the same cycle
 * length. It still gives useful, immediate feedback while specialists are
 * filling in their sheets, but it is a heuristic, not the source of truth.
 * The authoritative, cycle-aware clash + conflict check is the Timetable
 * Admin menu's "Sync Specialist Schedules" (TemplateSync.gs), which
 * resolves real repeating weeks via each sheet's own cycle length before
 * comparing anything. Re-run it periodically (or after specialists finish
 * editing), then run "Lock Synced Cells" once you're happy with the result,
 * to catch clashes this live check might miss when cycles differ.
 */
function applySpecialistClashFormatting(ss) {
  const specialistSheets = ss.getSheets()
    .filter(sheet => sheet.getName().startsWith('SPEC_') && sheet.getName() !== 'SPEC_TEMPLATE');

  if (!specialistSheets.length) return;

  const breakListRange = ss.getRangeByName('BreakList');

  specialistSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (lastRow < 5 || lastColumn < 10) return;

    const scheduleRanges = [];
    for (let column = 10; column <= lastColumn; column++) {
      if (String(sheet.getRange(4, column).getDisplayValue()).trim()) {
        scheduleRanges.push(sheet.getRange(5, column, lastRow - 4, 1));
      }
    }

    if (!scheduleRanges.length) return;

    const referenceCell = scheduleRanges[0].getCell(1, 1).getA1Notation();

    // Keep any rules that aren't the break/clash rules this function manages
    // (e.g. the class-colour rules added elsewhere).
    const retainedRules = sheet.getConditionalFormatRules().filter(rule => {
      const condition = rule.getBooleanCondition();
      if (!condition) return true;
      const isManagedRule = condition.getCriteriaValues().some(value =>
        typeof value === 'string' &&
        ((value.indexOf('INDIRECT') !== -1 && value.indexOf('COUNTIF') !== -1) ||
          value.indexOf('BreakList') !== -1)
      );
      return !isManagedRule;
    });

    const newRules = [];

    // Break periods (Recess, Lunch, etc, as configured on the DATA tab) are
    // intentionally identical across every specialist at the same time slot,
    // so the clash rule below would otherwise flag them as false-positive
    // collisions. This rule sits above the clash rule so it always wins.
    // INDIRECT because conditional formatting can't refer to other tabs
    // (or named ranges on them) directly.
    if (breakListRange) {
      newRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=COUNTIF(INDIRECT("BreakList"),' + referenceCell + ')>0')
          .setBackground('#D9D9D9')
          .setFontColor('#434343')
          .setRanges(scheduleRanges)
          .build()
      );
    }

    if (specialistSheets.length >= 2) {
      const counts = specialistSheets.map(specialistSheet => {
        const escapedName = specialistSheet.getName().replace(/'/g, "''");
        return 'COUNTIF(INDIRECT("\'' + escapedName +
          '\'!"&ADDRESS(ROW(),COLUMN(),4)),' + referenceCell + ')';
      }).join('+');

      newRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=AND(LEN(' + referenceCell + ')>0,(' + counts + ')>1)')
          .setBackground('#FF0000')
          .setFontColor('#FFFFFF')
          .setBold(true)
          .setRanges(scheduleRanges)
          .build()
      );
    }

    sheet.setConditionalFormatRules(newRules.concat(retainedRules));
  });
}

function computeGcd(a, b) {
  a = Math.abs(a); b = Math.abs(b);
  while (b) { const t = b; b = a % b; a = t; }
  return a || 1;
}

function computeLcm(a, b) {
  if (!a || !b) return Math.max(a, b, 1);
  return Math.abs(a * b) / computeGcd(a, b);
}

function weekLetterToIndex(letter) {
  return String(letter).toUpperCase().charCodeAt(0) - 65;
}

function parseTermWeekLabel(text) {
  const match = String(text || '').match(/Term\s+(\d+)\s*-\s*Week\s+([A-Za-z])/i);
  if (!match) return null;
  return { term: parseInt(match[1], 10), week: weekLetterToIndex(match[2]) };
}

function parseCycleFromClassSheet(classSheet) {
  const text = String(classSheet.getRange('C30').getDisplayValue() || '');
  const match = text.match(/(\d+)/);
  return match ? Math.max(1, parseInt(match[1], 10)) : 1;
}

function parseCycleFromSpecialistSheet(specialistSheet) {
  const parsed = parseInt(specialistSheet.getRange('D3').getValue(), 10);
  return isNaN(parsed) || parsed < 1 ? 1 : parsed;
}

/**
 * Maps every schedule column (from column 10 onward) to the {term, week, day}
 * it represents, by reading the header labels already written to the sheet
 * (row 1 = "Term X - Week Y", row 4 = day name) rather than assuming a fixed
 * column arithmetic. This makes the mapping correct regardless of how many
 * weeks are in this particular sheet's cycle, and naturally skips spacer
 * columns (which have no day name).
 */
function buildScheduleColumnMap(sheet, lastColumn) {
  const map = {};
  let currentLabel = null;
  for (let column = 10; column <= lastColumn; column++) {
    const headerText = String(sheet.getRange(1, column).getDisplayValue() || '').trim();
    if (headerText) {
      currentLabel = parseTermWeekLabel(headerText);
    }
    const dayName = String(sheet.getRange(4, column).getDisplayValue() || '').trim();
    if (dayName && currentLabel) {
      map[column] = { term: currentLabel.term, week: currentLabel.week, day: dayName };
    }
  }
  return map;
}

function sanitizeSheetName(name) {
  const sanitized = String(name || '')
    .replace(/[\/\\?*\[\]:]/g, '_')
    .trim();

  return sanitized || 'Untitled';
}
