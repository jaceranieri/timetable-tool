/**
 * Timetable Configuration Wizard backend.
 * Complete clean-paste version with dynamic sheet sizing, replicated template
 * formatting, class-only specialist dropdowns, activity-aware KLA dropdowns,
 * specialist clash detection, and TEMPLATE!B5 conditional formatting.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Timetable Configuration Wizard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * The master template the wizard copies. Set a TEMPLATE_ID script property
 * (Project Settings > Script Properties) to point the wizard at a test copy
 * while developing; delete the property to use the live template again.
 */
const LIVE_TEMPLATE_ID = '1x3pzraH_eG2nsQyiZPE9WGAzZ1wZIFGM2JUVF4W6FCk';

function getTemplateId() {
  const configuredId = PropertiesService.getScriptProperties().getProperty('TEMPLATE_ID');
  return (configuredId || '').trim() || LIVE_TEMPLATE_ID;
}

// ---- shared/setupRules.js (copied here by "npm run build"; edit that file, not this block) ----
/*
 * Setup rules shared by the wizard page (wizard/Index.html), which checks
 * them as you type, and the server (wizard/Code.gs, validateConfig), which
 * checks them again before generating: "npm run build" copies this file into
 * both. Edit it here, run the build, then paste both files into Apps Script.
 */
const SLOT_MINUTES = 5;
const MAX_CYCLE_WEEKS = 10;
const MAX_NAME_LENGTH = 90;
const MAX_SCHOOL_NAME_LENGTH = 100;
// Tabs that already exist in the template (or are planned), compared
// case-insensitively because Google Sheets tab names are.
const RESERVED_TAB_NAMES = ['data', 'template', 'overview', 'spec_template', 'start here'];
// Amadeus runs in Years 3-6 in every school. The activity name must match
// TEMPLATE!B21 exactly, because C21 counts cells with this value.
const AMADEUS_ACTIVITY_NAME = 'Amadeus - Tutorials';
// Every class needs its teacher's name for compliance, so stand-ins are refused.
const TEACHER_PLACEHOLDER = /^(tbc|tba|tbd|n\/?a|none|unknown|teacher|[-?.\s]*)$/i;
// ---- end of shared/setupRules.js ----

function isAmadeusGrade(grade) {
  return /^Year\s*[3-6]$/i.test(String(grade || '').trim());
}


function checkProgress(jobId) {
  if (!jobId) return null;
  try {
    return CacheService.getUserCache().get(jobId);
  } catch (error) {
    return null;
  }
}


/**
 * The wizard's unfinished setup, kept per user (User Properties) so it can be
 * resumed on any computer. Stored in parts because one property holds at
 * most 9 KB.
 */
const DRAFT_PROPERTY = 'WIZARD_DRAFT';
const DRAFT_PART_LENGTH = 8000;
const DRAFT_MAX_PARTS = 40;

function saveDraft(draftJson) {
  const text = String(draftJson || '');
  const partCount = Math.ceil(text.length / DRAFT_PART_LENGTH);
  if (!partCount || partCount > DRAFT_MAX_PARTS) return false;
  const properties = PropertiesService.getUserProperties();
  const values = { [DRAFT_PROPERTY]: String(partCount) };
  for (let i = 0; i < partCount; i++) {
    values[DRAFT_PROPERTY + '_' + i] = text.substring(i * DRAFT_PART_LENGTH, (i + 1) * DRAFT_PART_LENGTH);
  }
  properties.setProperties(values);
  // Parts left over from a longer draft.
  for (let i = partCount; i < DRAFT_MAX_PARTS; i++) {
    if (properties.getProperty(DRAFT_PROPERTY + '_' + i) === null) break;
    properties.deleteProperty(DRAFT_PROPERTY + '_' + i);
  }
  return true;
}


function loadDraft() {
  const properties = PropertiesService.getUserProperties();
  const partCount = Number(properties.getProperty(DRAFT_PROPERTY)) || 0;
  let text = '';
  for (let i = 0; i < partCount; i++) {
    const part = properties.getProperty(DRAFT_PROPERTY + '_' + i);
    if (part === null) return null; // incomplete: ignore it
    text += part;
  }
  return text || null;
}


function clearDraft() {
  const properties = PropertiesService.getUserProperties();
  const partCount = Number(properties.getProperty(DRAFT_PROPERTY)) || 0;
  properties.deleteProperty(DRAFT_PROPERTY);
  for (let i = 0; i < Math.max(partCount, 1); i++) properties.deleteProperty(DRAFT_PROPERTY + '_' + i);
  return true;
}


function updateJobProgress(jobId, percent, statusText) {
  if (!jobId) return;
  try {
    const normalizedPercent = percent === -1
      ? -1
      : Math.min(100, Math.max(0, Math.round(percent)));


    // Per-user cache: the wizard runs as each user, and only they poll
    // for their own job's progress.
    CacheService.getUserCache().put(jobId, JSON.stringify({
      percent: normalizedPercent,
      status: statusText
    }), 600);
  } catch (error) {
    console.warn('Failed to update cache progress:', error);
  }
}


// Generated timetables are saved in this folder in the user's My Drive,
// which is created if it doesn't exist. Set to '' to save them in My Drive
// itself.
const OUTPUT_FOLDER_NAME = 'Timetables';
// The copy keeps this prefix until generation finishes, so a copy left
// behind by a failed or timed-out run is easy to spot and delete.
const INCOMPLETE_FILE_PREFIX = 'INCOMPLETE - ';

/**
 * options.timing (used by the load test) logs how long each step took. It
 * flushes pending changes after every step so each step's cost, including
 * Google recalculating the spreadsheet, is counted against that step.
 */
function generateTimetable(payload, options) {
  const jobId = payload ? payload.jobId : null;
  const startedAt = Date.now();
  const timer = createStepTimer(Boolean(options && options.timing));


  try {
    // Check the setup before copying anything, so a bad setup doesn't leave
    // a half-built file in the user's Drive.
    const config = parsePayloadConfig(payload);
    validateConfig(config);

    updateJobProgress(jobId, 5, 'Initializing template duplication...');


    const templateFile = DriveApp.getFileById(getTemplateId());
    const fileName = config.schoolName + ' - Timetable - ' + new Date().toLocaleDateString('en-AU');
    const outputFolder = getOutputFolder();
    const newFile = outputFolder
      ? templateFile.makeCopy(INCOMPLETE_FILE_PREFIX + fileName, outputFolder)
      : templateFile.makeCopy(INCOMPLETE_FILE_PREFIX + fileName);
    timer.mark('Copy the template file', 'Setup');


    updateJobProgress(jobId, 12, 'Opening duplicated spreadsheet...');


    const ss = SpreadsheetApp.openById(newFile.getId());
    const dataSheet = ss.getSheetByName('DATA');


    if (!dataSheet) {
      throw new Error("Could not find a sheet named 'DATA' in the template.");
    }


    updateJobProgress(jobId, 18, 'Building DATA reference tables...');


    timer.mark('Open the copy', 'Setup');
    populateDataSheet(dataSheet, payload, config, ss);
    timer.mark('Fill the DATA tab', 'Setup');
    const tabs = createTabTracker(ss);
    buildSpecialistSheets(ss, config, jobId, timer, tabs);
    buildAllClassSheets(ss, config, jobId, timer, tabs);
    buildStartHereSheet(ss, config, tabs);
    timer.mark('Start here tab', 'Start here tab');
    newFile.setName(fileName);


    console.log(
      'Generated ' + config.classes.length + ' class and ' + config.specialists.length +
      ' specialist tab(s) in ' + Math.round((Date.now() - startedAt) / 1000) + 's'
    );
    timer.summary();
    updateJobProgress(jobId, 100, 'Timetable generation complete!');
    return ss.getUrl();
  } catch (error) {
    console.error('Error generating timetable:', error);
    const message = describeGenerationError(error);
    updateJobProgress(jobId, -1, message);
    throw new Error(message);
  }
}


/*
 * Changing a timetable that has already been generated (PLAN step 32), and
 * starting a new one from an existing timetable's setup (step 33).
 *
 * The wizard reads the setup kept in the timetable's DATA tab
 * (RawData_JSON), the user changes it, and updateTimetable applies only the
 * differences: new tabs are added, renamed classes and specialists keep their
 * tabs and lessons, and removed ones are archived (renamed and hidden, never
 * deleted). Bell times, breaks and the cycle of an existing class or
 * specialist can't be changed, because every tab's grid is built from them.
 */
const ARCHIVED_TAB_PREFIX = 'Archived - ';

/** Reads a timetable's setup for the wizard, from its link (or file ID). */
function readTimetableSetup(link) {
  const ss = openTimetable(link);
  const payload = readStoredPayload(ss);
  return {
    fileId: ss.getId(),
    fileName: ss.getName(),
    url: ss.getUrl(),
    exportVariables: payload.exportVariables
  };
}


/** The file ID in a Google Sheets link, or the ID itself. */
function getFileIdFromLink(link) {
  const text = String(link || '').trim();
  const fromLink = text.match(/\/d\/([A-Za-z0-9_-]+)/);
  if (fromLink) return fromLink[1];
  return /^[A-Za-z0-9_-]{20,}$/.test(text) ? text : '';
}


function openTimetable(link) {
  const fileId = getFileIdFromLink(link);
  if (!fileId) {
    throw new Error("That doesn't look like a link to a Google Sheet. Open the timetable and copy the address from your browser's address bar.");
  }
  try {
    return SpreadsheetApp.openById(fileId);
  } catch (error) {
    throw new Error("The timetable couldn't be opened. Check the link, and that the timetable is shared with you.");
  }
}


/** The wizard's setup as it was when the timetable was generated or last updated. */
function readStoredPayload(ss) {
  const dataSheet = ss.getSheetByName('DATA');
  const rows = dataSheet ? dataSheet.getRange(1, 1, Math.max(1, dataSheet.getLastRow()), 2).getValues() : [];
  const row = rows.find(item => String(item[0]).trim() === 'RawData_JSON');
  let payload = null;
  try {
    payload = row ? JSON.parse(String(row[1])) : null;
  } catch (error) {
    payload = null;
  }
  if (!payload || !Array.isArray(payload.exportVariables)) {
    throw new Error("This spreadsheet wasn't made by the timetable wizard (its DATA tab has no saved setup), so it can't be opened here.");
  }
  return payload;
}


/**
 * Applies the wizard's changed setup to an existing timetable.
 * request: { jobId, fileId, exportVariables, renames: { classes: [{from, to}], specialists: [{from, to}] } }
 * Returns { url, changes: [text], notes: [text] }.
 */
function updateTimetable(request) {
  const jobId = request ? request.jobId : null;
  try {
    const payload = { jobId: jobId, exportVariables: request.exportVariables };
    const config = parsePayloadConfig(payload);
    validateConfig(config);

    updateJobProgress(jobId, 5, 'Opening the timetable...');
    const ss = openTimetable(request.fileId);
    const oldConfig = parsePayloadConfig(readStoredPayload(ss));
    const plan = planTimetableUpdate(oldConfig, config, request.renames || {});
    if (plan.blocked.length) {
      throw new Error("These changes can't be made to an existing timetable, because every tab's grid is built from them:\n- " +
        plan.blocked.join('\n- ') + '\n\nUndo them, or start a new timetable from this one to make them.');
    }

    applyTimetableUpdate(ss, payload, oldConfig, config, plan, jobId);
    updateJobProgress(jobId, 100, 'Timetable updated!');
    return { url: ss.getUrl(), changes: plan.changes, notes: plan.notes };
  } catch (error) {
    console.error('Error updating timetable:', error);
    const text = String((error && error.message) || error || '').replace(/^(Error|Exception):\s*/i, '');
    const message = /can't be made to an existing timetable|Please fix the following|couldn't be opened|doesn't look like a link|wasn't made by the timetable wizard/.test(text)
      ? text
      : /do(n't| not) have permission|access denied|read-only/i.test(text)
        ? "You can open this timetable but not change it. Ask its owner to share it with you as an editor, then try again."
        : 'Something went wrong while updating the timetable: ' + text + '\n\nSome changes may already have been made; open the timetable to check, then try again.';
    updateJobProgress(jobId, -1, message);
    throw new Error(message);
  }
}


/**
 * Works out what an update changes, matching classes and specialists by
 * name (or by the renames the wizard sends). blocked lists changes that
 * would need a tab's grid rebuilt.
 */
function planTimetableUpdate(oldConfig, config, renames) {
  const key = name => String(name || '').trim().toLowerCase();
  const blocked = [];
  const changes = [];
  const notes = [];

  if (oldConfig.startTime !== config.startTime || oldConfig.endTime !== config.endTime) {
    blocked.push('School day: ' + oldConfig.startTime + ' - ' + oldConfig.endTime + ' to ' + config.startTime + ' - ' + config.endTime);
  }
  const describeBreaks = breaks => breaks.map(item => item.label + ' ' + item.start + '-' + item.end).join(', ');
  if (describeBreaks(oldConfig.breaks) !== describeBreaks(config.breaks)) {
    blocked.push('Breaks: ' + (describeBreaks(oldConfig.breaks) || 'none') + ' to ' + (describeBreaks(config.breaks) || 'none'));
  }

  // Pairs each new class or specialist with the old one it continues, if any.
  const match = (oldItems, newItems, renameList) => {
    const oldByKey = {};
    oldItems.forEach(item => { oldByKey[key(item.name)] = item; });
    const fromOf = {};
    (renameList || []).forEach(rename => {
      if (oldByKey[key(rename.from)]) fromOf[key(rename.to)] = key(rename.from);
    });
    const used = {};
    const pairs = newItems.map(item => {
      const fromKey = fromOf[key(item.name)] || key(item.name);
      const old = !used[fromKey] && oldByKey[fromKey];
      if (old) used[fromKey] = true;
      return { old: old || null, item: item };
    });
    return {
      kept: pairs.filter(pair => pair.old),
      added: pairs.filter(pair => !pair.old).map(pair => pair.item),
      removed: oldItems.filter(item => !used[key(item.name)])
    };
  };

  const classes = match(oldConfig.classes, config.classes, renames.classes);
  const specialists = match(oldConfig.specialists, config.specialists, renames.specialists);

  classes.kept.forEach(pair => {
    if (pair.old.cycle !== pair.item.cycle) {
      blocked.push('Class "' + pair.item.name + '": cycle ' + pair.old.cycle + ' to ' + pair.item.cycle + ' weeks');
    }
    if (pair.old.name !== pair.item.name) changes.push('Rename class "' + pair.old.name + '" to "' + pair.item.name + '"');
    if (pair.old.teacher !== pair.item.teacher) changes.push(pair.item.name + ': teacher ' + pair.old.teacher + ' to ' + pair.item.teacher);
    if ((pair.old.color || '') !== (pair.item.color || '')) changes.push(pair.item.name + ': new colour');
  });
  classes.added.forEach(item => changes.push('Add class "' + item.name + '" (' + item.grade + ', ' + item.teacher + ')'));
  classes.removed.forEach(item => changes.push('Archive class "' + item.name + '"'));

  specialists.kept.forEach(pair => {
    if (pair.old.cycle !== pair.item.cycle) {
      blocked.push('Specialist "' + pair.item.name + '": cycle ' + pair.old.cycle + ' to ' + pair.item.cycle + ' weeks');
    }
    if (pair.old.name !== pair.item.name) changes.push('Rename specialist "' + pair.old.name + '" to "' + pair.item.name + '"');
    if (pair.old.kla !== pair.item.kla) changes.push(pair.item.name + ': KLA ' + pair.old.kla + ' to ' + pair.item.kla);
  });
  specialists.added.forEach(item => changes.push('Add specialist "' + item.name + '" (' + item.kla + ')'));
  specialists.removed.forEach(item => changes.push('Archive specialist "' + item.name + '"'));

  const activityNames = list => list.map(item => item.name);
  const oldActivities = activityNames(oldConfig.activities);
  const newActivities = activityNames(config.activities);
  newActivities.filter(name => oldActivities.indexOf(name) === -1).forEach(name => changes.push('Add activity "' + name + '"'));
  oldActivities.filter(name => newActivities.indexOf(name) === -1).forEach(name => changes.push('Remove activity "' + name + '"'));
  if (oldConfig.schoolName !== config.schoolName) {
    changes.push('School name "' + oldConfig.schoolName + '" to "' + config.schoolName + '"');
  }

  if (classes.removed.length) {
    notes.push('Archived class tabs are hidden, not deleted: right-click a tab name > Show sheet to see one. ' +
      'Lessons with an archived class on the specialist tabs are kept; change or clear them, then sync again.');
  }
  if (specialists.removed.length) {
    notes.push("Archived specialists' synced lessons were taken off the class tabs.");
  }

  const klaList = list => uniqueNonBlank(list.klas.map(item => item.name).concat(activityNames(list.activities)));
  return {
    blocked: blocked,
    changes: changes,
    notes: notes,
    classes: classes,
    specialists: specialists,
    klaListGrew: klaList(config).length > klaList(oldConfig).length,
    classListGrew: config.classes.length > oldConfig.classes.length
  };
}


/** Makes the planned changes, then rebuilds DATA and the Start here tab. */
function applyTimetableUpdate(ss, payload, oldConfig, config, plan, jobId) {
  const timer = createStepTimer(false);
  const slots = createTimeSlots(config.startTime, config.endTime);
  const breakLabels = getBreakLabels(slots, config.breaks);

  // 1. Archive removed tabs, then rename (through temporary names, so two
  // classes can swap names).
  updateJobProgress(jobId, 15, 'Archiving and renaming tabs...');
  plan.classes.removed.forEach(item => archiveTab(ss, getClassTabName(item)));
  plan.specialists.removed.forEach(item => archiveTab(ss, getSpecialistTabName(item)));
  const renames = [];
  plan.classes.kept.forEach(pair => renames.push([getClassTabName(pair.old), getClassTabName(pair.item)]));
  plan.specialists.kept.forEach(pair => renames.push([getSpecialistTabName(pair.old), getSpecialistTabName(pair.item)]));
  const moving = renames
    .filter(rename => rename[0] !== rename[1])
    .map((rename, index) => ({ sheet: ss.getSheetByName(rename[0]), to: rename[1], temporary: 'Renaming ' + index + ' ' + Date.now() }))
    .filter(item => item.sheet);
  moving.forEach(item => item.sheet.setName(item.temporary));
  moving.forEach(item => item.sheet.setName(item.to));

  // 2. DATA: lists, named ranges and the saved setup.
  updateJobProgress(jobId, 30, 'Updating the DATA tab...');
  populateDataSheet(ss.getSheetByName('DATA'), payload, config, ss);

  // 3. Lessons that name a renamed or changed class or specialist.
  updateJobProgress(jobId, 40, 'Updating lessons on the class and specialist tabs...');
  updateSyncedLessons(ss, plan);
  updateSpecialistTabClasses(ss, config, plan);

  // 4. Dropdowns on existing tabs only list what DATA held when they were
  // made, so they're pointed at the longer lists.
  const validations = [];
  if (plan.klaListGrew) {
    const klaCount = uniqueNonBlank(config.klas.map(item => item.name).concat(config.activities.map(item => item.name))).length;
    plan.classes.kept.forEach(pair => validations.push({
      sheet: ss.getSheetByName(getClassTabName(pair.item)), cycle: pair.item.cycle, list: 'DATA!$D$2:$D$' + (klaCount + 1)
    }));
  }
  if (plan.classListGrew) {
    plan.specialists.kept.forEach(pair => validations.push({
      sheet: ss.getSheetByName(getSpecialistTabName(pair.item)), cycle: pair.item.cycle, list: 'DATA!$F$2:$F$' + (config.classes.length + 1)
    }));
  }
  refreshGridValidations(ss, validations.filter(item => item.sheet), slots.length, breakLabels);

  // 5. New tabs, and details on the existing ones (class teacher, specialist
  // KLA...). The specialists' clash formatting is redone for every tab.
  updateJobProgress(jobId, 55, 'Adding new tabs...');
  const startHere = captureStartHere(ss);
  const tabs = createTabTracker(ss);
  buildSpecialistSheets(ss, config, jobId, timer, tabs);
  buildAllClassSheets(ss, config, jobId, timer, tabs);

  // 6. Start here lists every tab, so it's rebuilt, keeping its "Last run"
  // lines and its list of Google Docs.
  updateJobProgress(jobId, 90, 'Updating the Start here tab...');
  if (startHere) ss.deleteSheet(startHere.sheet);
  buildStartHereSheet(ss, config, tabs);
  if (startHere) restoreStartHere(ss, startHere);

  if (oldConfig.schoolName !== config.schoolName && ss.getName().indexOf(oldConfig.schoolName) === 0) {
    ss.rename(config.schoolName + ss.getName().substring(oldConfig.schoolName.length));
  }
}


/** Renames a tab "Archived - <name>" (made unique) and hides it. */
function archiveTab(ss, tabName) {
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return;
  let name = (ARCHIVED_TAB_PREFIX + tabName).substring(0, 95);
  for (let n = 2; ss.getSheetByName(name); n++) {
    name = (ARCHIVED_TAB_PREFIX + tabName).substring(0, 90) + ' (' + n + ')';
  }
  sheet.setName(name).hideSheet();
}


/**
 * Class tabs: synced specialist lessons are matched to their specialist by
 * name (the "[SPEC_SYNC] <name>" note), so a renamed specialist's notes are
 * renamed, a changed KLA is written into their lessons, and an archived
 * specialist's lessons are taken off. Otherwise the next sync would treat
 * them as the class teacher's own entries.
 */
function updateSyncedLessons(ss, plan) {
  const key = name => String(name || '').trim().toLowerCase();
  const bySpecialist = {};
  plan.specialists.kept.forEach(pair => {
    if (pair.old.name !== pair.item.name || pair.old.kla !== pair.item.kla) {
      bySpecialist[key(pair.old.name)] = { name: pair.item.name, oldKla: pair.old.kla, kla: pair.item.kla };
    }
  });
  plan.specialists.removed.forEach(item => { bySpecialist[key(item.name)] = { removed: true, oldKla: item.kla }; });
  if (!Object.keys(bySpecialist).length) return;

  const prefix = SYNC_NOTE_PREFIX;
  plan.classes.kept.forEach(pair => {
    const sheet = ss.getSheetByName(getClassTabName(pair.item));
    if (!sheet || sheet.getLastRow() < FIRST_DATA_ROW || sheet.getLastColumn() < FIRST_GRID_COLUMN) return;
    const range = sheet.getRange(FIRST_DATA_ROW, FIRST_GRID_COLUMN,
      sheet.getLastRow() - FIRST_DATA_ROW + 1, sheet.getLastColumn() - FIRST_GRID_COLUMN + 1);
    const values = range.getValues();
    const notes = range.getNotes();
    let changed = false;
    notes.forEach((row, r) => row.forEach((note, c) => {
      if (String(note).indexOf(prefix) !== 0) return;
      const change = bySpecialist[key(String(note).substring(prefix.length))];
      if (!change) return;
      const holdsLesson = String(values[r][c]).trim() === change.oldKla;
      if (change.removed) {
        if (holdsLesson) values[r][c] = '';
        notes[r][c] = '';
      } else {
        notes[r][c] = prefix + ' ' + change.name;
        if (holdsLesson) values[r][c] = change.kla;
      }
      changed = true;
    }));
    if (changed) {
      range.setValues(values);
      range.setNotes(notes);
    }
  });
}


/**
 * Specialist tabs: lessons with a renamed class get its new name, and the
 * class colour rules are made again from the current classes.
 */
function updateSpecialistTabClasses(ss, config, plan) {
  const renamed = {};
  plan.classes.kept.forEach(pair => {
    if (pair.old.name !== pair.item.name) renamed[pair.old.name] = pair.item.name;
  });
  const coloursChanged = plan.classes.added.length || plan.classes.removed.length || Object.keys(renamed).length ||
    plan.classes.kept.some(pair => (pair.old.color || '') !== (pair.item.color || ''));

  plan.specialists.kept.forEach(pair => {
    const sheet = ss.getSheetByName(getSpecialistTabName(pair.item));
    if (!sheet || sheet.getLastRow() < FIRST_DATA_ROW || sheet.getLastColumn() < FIRST_GRID_COLUMN) return;
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (Object.keys(renamed).length) {
      const range = sheet.getRange(FIRST_DATA_ROW, FIRST_GRID_COLUMN, lastRow - FIRST_DATA_ROW + 1, lastColumn - FIRST_GRID_COLUMN + 1);
      const values = range.getValues();
      let changed = false;
      values.forEach(row => row.forEach((value, c) => {
        const name = String(value).trim();
        if (renamed.hasOwnProperty(name)) {
          row[c] = renamed[name];
          changed = true;
        }
      }));
      if (changed) range.setValues(values);
    }

    if (coloursChanged) {
      // Keep the break and clash rules (redone later); replace the class colours.
      const managed = sheet.getConditionalFormatRules().filter(rule => {
        const condition = rule.getBooleanCondition();
        return condition && condition.getCriteriaValues().some(value =>
          typeof value === 'string' && (value.indexOf('BreakList') !== -1 ||
            (value.indexOf('INDIRECT') !== -1 && value.indexOf('COUNTIF') !== -1)));
      });
      sheet.setConditionalFormatRules(managed);
      applyClassColorFormatting(sheet, config.classes, FIRST_DATA_ROW, lastRow, lastColumn);
    }
  });
}


/**
 * Points the dropdowns of existing tabs at a DATA list (A1 notation), on
 * every day column of the grid except break rows: in one request with the
 * Google Sheets API service, otherwise block by block.
 * items: [{ sheet, cycle, list }]
 */
function refreshGridValidations(ss, items, slotCount, breakLabels) {
  if (!items.length) return;
  const breakRuns = getBreakRuns(breakLabels);

  if (typeof Sheets !== 'undefined' && Sheets.Spreadsheets) {
    const requests = [];
    items.forEach(item => {
      const sheetId = item.sheet.getSheetId();
      getGridLayout(item.cycle).blocks.forEach(block => {
        const range = (startRow, rowCount) => ({
          sheetId: sheetId,
          startRowIndex: startRow - 1, endRowIndex: startRow - 1 + rowCount,
          startColumnIndex: block.startColumn - 1, endColumnIndex: block.startColumn - 1 + DAYS.length
        });
        requests.push({ setDataValidation: {
          range: range(FIRST_DATA_ROW, slotCount),
          rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: '=' + item.list }] }, strict: true, showCustomUi: true }
        } });
        breakRuns.forEach(run => requests.push({ setDataValidation: { range: range(FIRST_DATA_ROW + run.startIndex, run.length) } }));
      });
    });
    SpreadsheetApp.flush();
    Sheets.Spreadsheets.batchUpdate({ requests: requests }, ss.getId());
    return;
  }

  items.forEach(item => {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(ss.getSheetByName('DATA').getRange(item.list.split('!')[1].replace(/\$/g, '')), true)
      .setAllowInvalid(false)
      .build();
    const layout = getGridLayout(item.cycle);
    layout.blocks.forEach(block => {
      item.sheet.getRange(FIRST_DATA_ROW, block.startColumn, slotCount, DAYS.length).setDataValidation(rule);
    });
    const breakRanges = [];
    breakRuns.forEach(run => layout.blocks.forEach(block => {
      breakRanges.push(toA1(FIRST_DATA_ROW + run.startIndex, block.startColumn, run.length, DAYS.length));
    }));
    if (breakRanges.length) item.sheet.getRangeList(breakRanges).clearDataValidations();
  });
}


/** What the Start here tab holds that isn't rebuilt from the setup. */
function captureStartHere(ss) {
  const sheet = ss.getSheetByName(START_HERE_SHEET_NAME);
  if (!sheet) return null;
  const valueOf = name => {
    const range = ss.getRangeByName(name);
    return range ? range.getValue() : '';
  };
  const docs = [];
  const headings = ss.getRangeByName('StartHereDocs');
  if (headings) {
    const firstRow = headings.getRow() + 1;
    const rowCount = sheet.getLastRow() - firstRow + 1;
    if (rowCount > 0) {
      const range = sheet.getRange(firstRow, 1, rowCount, 4);
      const formulas = range.getFormulas();
      range.getValues().forEach((row, i) => {
        const first = formulas[i][0] || row[0];
        if (!String(first).trim() || String(row[0]).indexOf('None yet') === 0) return;
        docs.push([first, row[1], row[2], row[3]]);
      });
    }
  }
  return { sheet: sheet, sync: valueOf('StartHereSyncStatus'), lock: valueOf('StartHereLockStatus'), docs: docs };
}


function restoreStartHere(ss, saved) {
  if (saved.sync) ss.getRangeByName('StartHereSyncStatus').setValue(saved.sync);
  if (saved.lock) ss.getRangeByName('StartHereLockStatus').setValue(saved.lock);
  if (saved.docs.length) {
    const headings = ss.getRangeByName('StartHereDocs');
    headings.getSheet().getRange(headings.getRow() + 1, 1, saved.docs.length, 4)
      .setValues(saved.docs)
      .setFontColor(START_HERE_STYLE.text);
  }
}


/**
 * Times the steps of a generation for the load test. When disabled, mark()
 * and summary() do nothing, so normal runs aren't slowed by extra flushes.
 */
function createStepTimer(enabled) {
  const groups = {};
  const order = [];
  const startedAt = Date.now();
  let lastAt = startedAt;

  return {
    mark: function (step, group) {
      if (!enabled) return;
      SpreadsheetApp.flush();
      const now = Date.now();
      const seconds = (now - lastAt) / 1000;
      lastAt = now;
      if (!groups[group]) {
        groups[group] = { count: 0, seconds: 0 };
        order.push(group);
      }
      groups[group].count++;
      groups[group].seconds += seconds;
      console.log(seconds.toFixed(1) + 's  ' + step);
    },
    summary: function () {
      if (!enabled) return;
      const lines = order.map(group => {
        const item = groups[group];
        return group + ': ' + item.seconds.toFixed(1) + 's' +
          (item.count > 1 ? ' (' + item.count + ' x ' + (item.seconds / item.count).toFixed(1) + 's)' : '');
      });
      lines.push('Total: ' + ((Date.now() - startedAt) / 1000).toFixed(1) + 's');
      console.log('Timing summary\n' + lines.join('\n'));
    }
  };
}


/** The OUTPUT_FOLDER_NAME folder in the user's My Drive, created if needed. */
function getOutputFolder() {
  if (!OUTPUT_FOLDER_NAME) return null;
  const root = DriveApp.getRootFolder();
  const existing = root.getFoldersByName(OUTPUT_FOLDER_NAME);
  return existing.hasNext() ? existing.next() : root.createFolder(OUTPUT_FOLDER_NAME);
}


/** Turns an error from generateTimetable into a message for the wizard. */
function describeGenerationError(error) {
  const text = String((error && error.message) || error || '')
    .replace(/^(Error|Exception):\s*/i, '');

  // Setup problems from validateConfig are already written for the user.
  if (text.indexOf('Please fix the following') === 0) return text;

  if (/no item with the given id|do not have permission|access denied/i.test(text)) {
    return "You don't have access to the master timetable template. Ask the person who " +
      'manages the timetable tool to share it with you (view access is enough).';
  }

  return 'Something went wrong while generating the timetable: ' + text + '\n\n' +
    'If a file named "' + INCOMPLETE_FILE_PREFIX + '..." was created in your Drive, it can ' +
    'be deleted. Please try again.';
}


function parsePayloadConfig(payload) {
  const varMap = {};


  if (payload && Array.isArray(payload.exportVariables)) {
    payload.exportVariables.forEach(item => {
      varMap[item.name] = item.value;
    });
  }


  const breaks = [];
  let index = 1;
  while (varMap['Break_' + index + '_Name'] !== undefined) {
    const name = String(varMap['Break_' + index + '_Name'] || '').trim();
    const start = String(varMap['Break_' + index + '_Start'] || '').trim();
    const end = String(varMap['Break_' + index + '_End'] || '').trim();


    if (name || start || end) {
      breaks.push({ label: name, start: start, end: end });
    }
    index++;
  }


  const klas = [];
  index = 1;
  while (varMap['KLA_' + index + '_Name'] !== undefined) {
    const name = String(varMap['KLA_' + index + '_Name'] || '').trim();
    if (name && !/^non[- ]?curricul(?:ar|um)$/i.test(name)) {
      klas.push({ name: name });
    }
    index++;
  }


  const activities = [];
  index = 1;
  while (varMap['Activity_' + index + '_Name'] !== undefined) {
    const name = String(varMap['Activity_' + index + '_Name'] || '').trim();
    if (name) activities.push({ name: name });
    index++;
  }

  // Amadeus is compulsory, so make sure it is always selectable on class tabs.
  const hasAmadeus = activities.some(
    item => item.name.toLowerCase() === AMADEUS_ACTIVITY_NAME.toLowerCase()
  );
  if (!hasAmadeus) activities.push({ name: AMADEUS_ACTIVITY_NAME });


  const specialists = [];
  index = 1;
  while (varMap['Specialist_' + index + '_Name'] !== undefined) {
    const name = String(varMap['Specialist_' + index + '_Name'] || '').trim();
    const kla = String(varMap['Specialist_' + index + '_KLA'] || '').trim();
    const cycle = Number(varMap['Specialist_' + index + '_Cycle']);
    if (name) specialists.push({ name: name, kla: kla, cycle: cycle });
    index++;
  }


  const classes = [];
  index = 1;
  while (varMap['Class_' + index + '_Name'] !== undefined) {
    const name = String(varMap['Class_' + index + '_Name'] || '').trim();
    const teacher = String(varMap['Class_' + index + '_Teacher'] || '').trim();
    const grade = String(varMap['Class_' + index + '_Grade'] || '').trim();
    const cycle = Number(varMap['Class_' + index + '_Cycle']);
    const color = String(varMap['Class_' + index + '_Color'] || '').trim();


    if (name) {
      classes.push({
        name: name,
        teacher: teacher,
        grade: grade,
        cycle: cycle,
        color: color
      });
    }
    index++;
  }


  return {
    schoolName: String(varMap.SchoolName || '').trim(),
    startTime: String(varMap.SchoolStart || '').trim(),
    endTime: String(varMap.SchoolEnd || '').trim(),
    breaks: breaks,
    klas: klas,
    activities: activities,
    specialists: specialists,
    classes: classes
  };
}


/**
 * Checks the parsed setup and throws one error listing every problem, so
 * nothing is generated from a setup that would produce missing or broken
 * tabs.
 */
function validateConfig(config) {
  const problems = [];

  if (!config.schoolName) {
    problems.push('Enter the school name.');
  } else if (startsWithFormula(config.schoolName)) {
    problems.push("The school name can't start with =.");
  } else if (config.schoolName.length > MAX_SCHOOL_NAME_LENGTH) {
    problems.push('The school name is longer than ' + MAX_SCHOOL_NAME_LENGTH + ' characters.');
  }

  const dayStart = timeToMinutes(config.startTime);
  const dayEnd = timeToMinutes(config.endTime);
  if (dayStart === null || dayEnd === null) {
    problems.push('School start and end times must be valid times.');
  } else if (dayStart % SLOT_MINUTES || dayEnd % SLOT_MINUTES) {
    problems.push('School start and end times must be on a 5-minute mark (e.g. 8:30, 8:35).');
  } else if (dayStart >= dayEnd) {
    problems.push('School end time must be after the start time.');
  }

  const validBreaks = [];
  config.breaks.forEach((item, index) => {
    const label = item.label ? '"' + item.label + '"' : 'Break ' + (index + 1);
    const start = timeToMinutes(item.start);
    const end = timeToMinutes(item.end);
    if (!item.label) problems.push(label + ' needs a name.');
    else if (startsWithFormula(item.label)) problems.push(label + " can't start with =.");
    if (start === null || end === null) {
      problems.push(label + ' needs a valid start and end time.');
    } else if (start % SLOT_MINUTES || end % SLOT_MINUTES) {
      problems.push(label + ' must start and end on a 5-minute mark.');
    } else if (start >= end) {
      problems.push(label + ' must end after it starts.');
    } else if (dayStart !== null && dayEnd !== null && (start < dayStart || end > dayEnd)) {
      problems.push(label + ' must be within the school day.');
    } else {
      validBreaks.push({ label: label, start: start, end: end });
    }
  });
  validBreaks.sort((a, b) => a.start - b.start);
  for (let i = 1; i < validBreaks.length; i++) {
    if (validBreaks[i].start < validBreaks[i - 1].end) {
      problems.push(validBreaks[i - 1].label + ' and ' + validBreaks[i].label + ' overlap.');
    }
  }

  if (!config.classes.length) {
    problems.push('Add at least one class.');
  }

  const breakNames = config.breaks.map(item => item.label.toLowerCase());
  const seenClasses = {};
  config.classes.forEach(item => {
    const label = 'Class "' + item.name + '"';
    const nameProblem = getNameProblem(item.name, true);
    if (nameProblem) problems.push(label + ' ' + nameProblem);
    const key = item.name.toLowerCase();
    if (seenClasses[key]) problems.push(label + ' is used more than once.');
    seenClasses[key] = true;
    if (breakNames.indexOf(key) !== -1) problems.push(label + ' has the same name as a break.');
    if (TEACHER_PLACEHOLDER.test(item.teacher)) {
      problems.push(label + " needs its teacher's name" + (item.teacher ? ' ("' + item.teacher + '" isn\'t a name).' : '.'));
    } else if (startsWithFormula(item.teacher)) {
      problems.push(label + ": the teacher name can't start with =.");
    }
    if (!isValidCycle(item.cycle)) {
      problems.push(label + ' needs a cycle length of 1 to ' + MAX_CYCLE_WEEKS + ' weeks.');
    }
  });

  const klaNames = config.klas.map(item => item.name.toLowerCase());
  const seenSpecialists = {};
  config.specialists.forEach(item => {
    const label = 'Specialist "' + item.name + '"';
    const nameProblem = getNameProblem(item.name, false);
    if (nameProblem) problems.push(label + ' ' + nameProblem);
    const key = item.name.toLowerCase();
    if (seenSpecialists[key]) problems.push(label + ' is used more than once.');
    seenSpecialists[key] = true;
    if (klaNames.indexOf(item.kla.toLowerCase()) === -1) {
      problems.push(label + ' needs a KLA from the list.');
    }
    if (!isValidCycle(item.cycle)) {
      problems.push(label + ' needs a cycle length of 1 to ' + MAX_CYCLE_WEEKS + ' weeks.');
    }
  });

  config.activities.forEach(item => {
    if (startsWithFormula(item.name)) {
      problems.push('Activity "' + item.name + '" can\'t start with =.');
    }
  });

  if (problems.length) {
    const shown = problems.slice(0, 15);
    if (problems.length > shown.length) {
      shown.push('...and ' + (problems.length - shown.length) + ' more.');
    }
    throw new Error('Please fix the following and try again:\n- ' + shown.join('\n- '));
  }
}

/**
 * Returns why a class or specialist name can't be used as a tab name, or ''
 * if it's fine. Class names also can't contain an apostrophe, because the
 * OVERVIEW tab looks class tabs up by name with INDIRECT.
 */
function getNameProblem(name, isClass) {
  const lower = name.toLowerCase();
  if (startsWithFormula(name)) return "can't start with =.";
  if (name.length > MAX_NAME_LENGTH) return 'is longer than ' + MAX_NAME_LENGTH + ' characters.';
  if (isClass ? /[\/\\?*\[\]:']/.test(name) : /[\/\\?*\[\]:]/.test(name)) {
    return "can't contain " + (isClass ? "/ \\ ? * [ ] : or '" : '/ \\ ? * [ ] or :') + '.';
  }
  if (isClass && (RESERVED_TAB_NAMES.indexOf(lower) !== -1 || /^spec_/.test(lower))) {
    return 'is a name the timetable uses for its own tabs.';
  }
  if (!isClass && lower === 'template') {
    return 'is a name the timetable uses for its own tabs.';
  }
  return '';
}

function startsWithFormula(value) {
  return String(value || '').trim().charAt(0) === '=';
}

function isValidCycle(value) {
  return Number.isInteger(value) && value >= 1 && value <= MAX_CYCLE_WEEKS;
}

/** 'HH:MM' to minutes after midnight, or null if it isn't a valid time. */
function timeToMinutes(value) {
  const match = String(value || '').trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!match) return null;
  const hours = parseInt(match[1], 10);
  const minutes = parseInt(match[2], 10);
  if (hours > 23 || minutes > 59) return null;
  return hours * 60 + minutes;
}


function populateDataSheet(dataSheet, payload, config, ss) {
  dataSheet.clearContents();


  const variableRows = [];
  if (payload && Array.isArray(payload.exportVariables)) {
    payload.exportVariables.forEach(item => {
      variableRows.push([item.name, item.value]);
    });
  }
  variableRows.push(['RawData_JSON', JSON.stringify(payload)]);


  const klaNames = uniqueNonBlank(
    config.klas.map(item => item.name)
      .concat(config.activities.map(item => item.name))
  );
  const classNames = uniqueNonBlank(config.classes.map(item => item.name));
  const specialistRows = config.specialists.map(item => [item.name, item.kla]);
  const classColorRows = config.classes.map(item => [item.name, item.color || '']);
  const breakNames = uniqueNonBlank(config.breaks.map(item => item.label));


  const requiredRows = Math.max(
    2,
    variableRows.length + 1,
    klaNames.length + 1,
    classNames.length + 1,
    specialistRows.length + 1,
    classColorRows.length + 1,
    breakNames.length + 1
  );


  ensureSheetSize(dataSheet, requiredRows, 16);
  // Plain text, so values like "3-4" or "08:30" are stored exactly as typed
  // instead of being turned into dates or times.
  dataSheet.getRange(1, 1, dataSheet.getMaxRows(), 16).setNumberFormat('@');


  dataSheet.getRange('A1:B1').setValues([['VariableName', 'Value']]);
  if (variableRows.length) {
    dataSheet.getRange(2, 1, variableRows.length, 2).setValues(variableRows);
  }


  dataSheet.getRange('D1').setValue('KLAList');
  if (klaNames.length) {
    const klaValues = klaNames.map(name => [name]);
    dataSheet.getRange(2, 4, klaValues.length, 1).setValues(klaValues);
    ss.setNamedRange('KLAList', dataSheet.getRange(2, 4, klaValues.length, 1));
  }


  dataSheet.getRange('F1').setValue('ClassList');
  if (classNames.length) {
    const classValues = classNames.map(name => [name]);
    dataSheet.getRange(2, 6, classValues.length, 1).setValues(classValues);
    ss.setNamedRange(
      'ClassList',
      dataSheet.getRange(2, 6, classValues.length, 1)
    );
  }


  dataSheet.getRange('H1:I1').setValues([['SpecialistName', 'AssignedKLA']]);
  if (specialistRows.length) {
    dataSheet.getRange(2, 8, specialistRows.length, 2).setValues(specialistRows);
    ss.setNamedRange(
      'SpecialistMapping',
      dataSheet.getRange(2, 8, specialistRows.length, 2)
    );
  }


  dataSheet.getRange('K1:L1').setValues([['ClassName', 'ClassColor']]);
  if (classColorRows.length) {
    dataSheet.getRange(2, 11, classColorRows.length, 2).setValues(classColorRows);
    ss.setNamedRange(
      'ClassColorMapping',
      dataSheet.getRange(2, 11, classColorRows.length, 2)
    );
  }


  dataSheet.getRange('P1').setValue('Breaks');
  if (breakNames.length) {
    const breakValues = breakNames.map(name => [name]);
    dataSheet.getRange(2, 16, breakValues.length, 1).setValues(breakValues);
    // Used by the specialist clash formatting so break rows (identical on
    // every specialist tab) are shown grey instead of as clashes.
    ss.setNamedRange('BreakList', dataSheet.getRange(2, 16, breakValues.length, 1));
  }
}


/**
 * Grid layout shared by class and specialist tabs. TemplateSync.gs reads
 * the same layout back from the headers ("Term X - Week Y" in row 1, day
 * names in row 4), so changing it here needs no change there.
 */
const DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
const TERM_COUNT = 4;
const FIRST_DATA_ROW = 5;
const TIME_COLUMN = 9;
const FIRST_GRID_COLUMN = 10;
const DAY_COLUMN_WIDTH = 120;
const WEEK_SPACER_WIDTH = 20;
const TERM_SPACER_WIDTH = 40;
const BREAK_BACKGROUND = '#F2F2F2';

/**
 * Builds a specialist tab for every specialist. The first specialist with
 * each cycle length gets a tab built from SPEC_Template; the others with the
 * same cycle length are copies of it (see copyTabs).
 */
function buildSpecialistSheets(ss, config, jobId, timer, tabs) {
  timer = timer || createStepTimer(false);
  tabs = tabs || createTabTracker(ss);
  const specTemplateSheet = ss.getSheetByName('SPEC_Template');
  const classListRange = ss.getRangeByName('ClassList');


  if (!config.specialists.length) return;
  if (!specTemplateSheet) {
    throw new Error("Could not find a sheet named 'SPEC_Template' in the template.");
  }
  if (!classListRange) {
    throw new Error("Named range 'ClassList' was not created.");
  }


  const slots = createTimeSlots(config.startTime, config.endTime);
  const breakLabels = getBreakLabels(slots, config.breaks);
  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(classListRange, true)
    .setAllowInvalid(false)
    .build();
  const prototypes = {}; // cycle length -> first tab built with it
  const copies = [];
  const firstTabIndex = tabs.nextIndex;
  const tabNames = config.specialists.map(getSpecialistTabName);


  config.specialists.forEach((specialist, specialistIndex) => {
    updateJobProgress(
      jobId,
      25 + (specialistIndex / config.specialists.length) * 15,
      'Building Specialist Tab ' + (specialistIndex + 1) + ' of ' +
        config.specialists.length + ': ' + specialist.name
    );

    const sheetName = tabNames[specialistIndex];
    const existing = ss.getSheetByName(sheetName);

    if (existing) {
      writeDetails(existing, getSpecialistDetails(specialist));
      tabs.sheetIds[sheetName] = existing.getSheetId();
    } else if (prototypes[specialist.cycle]) {
      copies.push({
        prototype: prototypes[specialist.cycle],
        name: sheetName,
        index: firstTabIndex + specialistIndex,
        details: getSpecialistDetails(specialist)
      });
    } else {
      const sheet = specTemplateSheet.copyTo(ss).setName(sheetName).showSheet();
      // Start from a clean slate: don't inherit any conditional formatting
      // rules that happen to exist on SPEC_Template. Only the rules this
      // script adds below should apply.
      sheet.clearConditionalFormatRules();
      const grid = buildGrid(
        sheet, specTemplateSheet, specialist.cycle, slots, breakLabels, validationRule
      );
      applyClassColorFormatting(
        sheet,
        config.classes,
        FIRST_DATA_ROW,
        grid.lastDataRow,
        grid.layout.lastColumn
      );
      sheet.getRangeList(['B1', 'B3']).setNumberFormat('@');
      writeDetails(sheet, getSpecialistDetails(specialist));
      prototypes[specialist.cycle] = sheet;
      tabs.nextIndex++;
      tabs.sheetIds[sheetName] = sheet.getSheetId();
      timer.mark(sheetName + ': build (' + specialist.cycle + '-week cycle)', 'Specialist tabs: build');
    }
  });


  // Clash formatting goes on the tabs built so far, listing every specialist
  // including the ones still to be copied, so the copies inherit it. (Tabs
  // copied through the Sheets API can't be reached through SpreadsheetApp in
  // this run, so they couldn't be formatted afterwards.)
  applySpecialistClashFormatting(ss, config.specialists.map((specialist, index) => ({
    name: tabNames[index],
    cycle: specialist.cycle
  })));
  timer.mark('Specialist clash formatting', 'Specialist clash formatting');

  updateJobProgress(jobId, 40, 'Copying ' + copies.length + ' specialist tab(s)...');
  copyTabs(ss, copies, timer, 'Specialist tabs: copy', tabs);
}


/**
 * Builds a class tab for every class, the same way as the specialist tabs:
 * built from TEMPLATE once per cycle length, copied for the rest.
 *
 * A tab that already exists only has its details (C27:C31) refreshed, so
 * its timetable entries are never touched.
 */
function buildAllClassSheets(ss, config, jobId, timer, tabs) {
  timer = timer || createStepTimer(false);
  tabs = tabs || createTabTracker(ss);
  const templateSheet = ss.getSheetByName('TEMPLATE');
  const klaListRange = ss.getRangeByName('KLAList');


  if (!templateSheet || !config.classes.length) return;
  if (!klaListRange) {
    throw new Error("Named range 'KLAList' was not created.");
  }


  const slots = createTimeSlots(config.startTime, config.endTime);
  const breakLabels = getBreakLabels(slots, config.breaks);
  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(klaListRange, true)
    .setAllowInvalid(false)
    .build();
  const prototypes = {}; // cycle length -> first tab built with it
  const copies = [];
  const firstTabIndex = tabs.nextIndex;


  config.classes.forEach((classConfig, classIndex) => {
    const sheetName = getClassTabName(classConfig);
    const existing = ss.getSheetByName(sheetName);

    if (existing) {
      writeClassDetails(existing, classConfig);
      tabs.sheetIds[sheetName] = existing.getSheetId();
    } else if (prototypes[classConfig.cycle]) {
      copies.push({
        prototype: prototypes[classConfig.cycle],
        name: sheetName,
        index: firstTabIndex + classIndex,
        details: getClassDetails(classConfig)
      });
    } else {
      updateJobProgress(
        jobId,
        45 + (classIndex / config.classes.length) * 25,
        'Building Classroom Tab for ' + classConfig.cycle + '-week cycle: ' + classConfig.name
      );
      const sheet = templateSheet.copyTo(ss).setName(sheetName).showSheet();
      const grid = buildGrid(
        sheet, templateSheet, classConfig.cycle, slots, breakLabels, validationRule
      );
      pasteTemplateConditionalFormatting(templateSheet, sheet, grid.layout, slots.length);
      writeClassDetails(sheet, classConfig);
      prototypes[classConfig.cycle] = sheet;
      tabs.nextIndex++;
      tabs.sheetIds[sheetName] = sheet.getSheetId();
      timer.mark(sheetName + ': build (' + classConfig.cycle + '-week cycle)', 'Class tabs: build');
    }
  });


  updateJobProgress(jobId, 75, 'Copying ' + copies.length + ' classroom tab(s)...');
  copyTabs(ss, copies, timer, 'Class tabs: copy', tabs);
}


/**
 * Where new tabs go and which sheet ids are taken, kept here rather than
 * read back from the spreadsheet: tabs added through the Sheets API aren't
 * visible to SpreadsheetApp for the rest of the run.
 */
function createTabTracker(ss) {
  const usedIds = {};
  const sheets = ss.getSheets();
  sheets.forEach(sheet => { usedIds[sheet.getSheetId()] = true; });
  // sheetIds: tab name -> sheet id of every class and specialist tab, for
  // the links on the Start here tab.
  return { nextIndex: sheets.length, usedIds: usedIds, sheetIds: {} };
}


/**
 * Creates tabs as copies of tabs that are already built, each at its index
 * in the tab order, and writes their details.
 *
 * Every change makes Google recalculate the spreadsheet (all the compliance
 * formulas and the OVERVIEW lookups), and that grows with every tab. So with
 * the Google Sheets API service enabled in the project, the copies and their
 * details are sent in a few large requests instead of one change at a time.
 * Without the service, tabs are copied one by one, which is much slower for
 * large schools.
 *
 * copies: [{ prototype: Sheet, name, index, details: [{ row, column, values }] }]
 */
const COPIES_PER_REQUEST = 10;

function copyTabs(ss, copies, timer, timerGroup, tabs) {
  if (!copies.length) return;
  tabs.nextIndex += copies.length;

  if (typeof Sheets === 'undefined' || !Sheets.Spreadsheets) {
    copies.forEach(copy => {
      const sheet = copy.prototype.copyTo(ss).setName(copy.name);
      tabs.sheetIds[copy.name] = sheet.getSheetId();
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(copy.index + 1);
      writeDetails(sheet, copy.details);
      timer.mark(copy.name + ': copy, one at a time', timerGroup);
    });
    return;
  }

  // The tabs being copied must be complete before Google copies them.
  SpreadsheetApp.flush();

  const usedIds = tabs.usedIds;
  ss.getSheets().forEach(sheet => { usedIds[sheet.getSheetId()] = true; });
  const newSheetId = () => {
    let id;
    do {
      id = Math.floor(Math.random() * 2000000000) + 1;
    } while (usedIds[id]);
    usedIds[id] = true;
    return id;
  };

  for (let start = 0; start < copies.length; start += COPIES_PER_REQUEST) {
    const batch = copies.slice(start, start + COPIES_PER_REQUEST);
    const requests = [];
    batch.forEach(copy => {
      const sheetId = newSheetId();
      tabs.sheetIds[copy.name] = sheetId;
      requests.push({
        duplicateSheet: {
          sourceSheetId: copy.prototype.getSheetId(),
          insertSheetIndex: copy.index,
          newSheetId: sheetId,
          newSheetName: copy.name
        }
      });
      copy.details.forEach(detail => {
        requests.push({
          updateCells: {
            start: { sheetId: sheetId, rowIndex: detail.row - 1, columnIndex: detail.column - 1 },
            rows: detail.values.map(row => ({ values: row.map(toCellData) })),
            fields: 'userEnteredValue'
          }
        });
      });
    });
    Sheets.Spreadsheets.batchUpdate({ requests: requests }, ss.getId());
    timer.mark(
      'Copy ' + batch.map(copy => copy.name).join(', ') + ' (one request)',
      timerGroup
    );
  }
}


function toCellData(value) {
  if (typeof value === 'boolean') return { userEnteredValue: { boolValue: value } };
  if (typeof value === 'number') return { userEnteredValue: { numberValue: value } };
  return { userEnteredValue: { stringValue: String(value) } };
}


/** Cells written on each class tab, as {row, column, values} blocks. */
function getClassDetails(classConfig) {
  return [
    {
      row: 27,
      column: 3,
      values: [
        [classConfig.teacher],
        [classConfig.name],
        [classConfig.grade],
        [classConfig.cycle + ' week cycle']
      ]
    },
    // C18 (Amadeus compliance) reads C31 to decide whether Amadeus applies.
    { row: 31, column: 2, values: [['Amadeus', isAmadeusGrade(classConfig.grade)]] }
  ];
}


function getSpecialistDetails(specialist) {
  return [
    { row: 1, column: 2, values: [[specialist.name]] },
    { row: 3, column: 2, values: [[specialist.kla]] },
    { row: 3, column: 4, values: [[specialist.cycle]] }
  ];
}


function writeClassDetails(sheet, classConfig) {
  sheet.getRange('C27:C29').setNumberFormat('@');
  writeDetails(sheet, getClassDetails(classConfig));
}


function writeDetails(sheet, details) {
  details.forEach(detail => {
    sheet.getRange(detail.row, detail.column, detail.values.length, detail.values[0].length)
      .setValues(detail.values);
  });
}


/**
 * The Start here tab: the first tab of every generated timetable. It lists
 * the order of work (specialists first, then sync and lock, then class
 * teachers, then the compliance check), who does each stage and a live
 * status for each, followed by a link to every specialist and class tab
 * with its progress.
 *
 * The status formulas refer to the tabs by name. "Sync Specialist
 * Schedules" and "Lock Synced Cells" (TemplateSync.gs) write their "Last
 * run" text through the named ranges StartHereSyncStatus and
 * StartHereLockStatus.
 */
const START_HERE_SHEET_NAME = 'Start here';
// Shown under "Google Docs" until the first one is made; TimetableTools.gs
// replaces it with the first document's row.
const START_HERE_NO_DOCS_TEXT = 'None yet. Timetable Admin > Make a Google Doc adds them here.';

function buildStartHereSheet(ss, config, tabs) {
  if (ss.getSheetByName(START_HERE_SHEET_NAME)) return;

  const sheet = ss.insertSheet(START_HERE_SHEET_NAME, 0);
  const quote = name => "'" + name.replace(/'/g, "''") + "'";
  const link = name => {
    const id = tabs.sheetIds[name];
    const label = name.replace(/"/g, '""');
    return id ? '=HYPERLINK("#gid=' + id + '","' + label + '")' : name;
  };
  const grid = name => quote(name) + '!J5:AAK';

  // Where each part goes, so the stage formulas can refer to the lists.
  const specialistCount = config.specialists.length;
  const classCount = config.classes.length;
  const specialistFirstRow = 13;
  const specialistLastRow = specialistFirstRow + Math.max(specialistCount, 1) - 1;
  const classHeadingRow = specialistLastRow + 2;
  const classFirstRow = classHeadingRow + 2;
  const classLastRow = classFirstRow + classCount - 1;
  const docsHeadingRow = classLastRow + 2;
  const lessons = 'D' + specialistFirstRow + ':D' + specialistLastRow;
  const emptySlots = 'D' + classFirstRow + ':D' + classLastRow;
  const conflicts = 'E' + classFirstRow + ':E' + classLastRow;
  const compliant = 'F' + classFirstRow + ':F' + classLastRow;

  const rows = [];
  const put = (rowNumber, values) => {
    while (rows.length < rowNumber) rows.push(['', '', '', '', '', '']);
    rows[rowNumber - 1] = values.concat(new Array(6 - values.length).fill(''));
  };

  put(1, [config.schoolName + ' - Timetable']);
  put(2, ['Work through the stages in order. The status column updates as the timetable is filled in.']);
  put(4, ['Stage', 'Who', 'What to do', 'Status', 'Last run']);
  put(5, [
    '1. Fill in the specialist timetables',
    'Specialist teachers',
    'Open your SPEC_ tab (listed below) and choose a class for each lesson. A cell turns red when another specialist has the same class at the same time.',
    specialistCount
      ? '=SUM(' + lessons + ')&" lessons entered on "&COUNTIF(' + lessons + ',">0")&" of ' +
        specialistCount + ' specialist tab(s)"'
      : 'No specialists in this timetable'
  ]);
  put(6, [
    '2. Sync to the class timetables',
    'Timetable coordinator',
    'Timetable Admin > Sync Specialist Schedules. Copies every specialist lesson onto the class tabs and flags clashes in red. Run it again after any change to a specialist tab.',
    '=IF(SUM(' + conflicts + ')>0,SUM(' + conflicts + ')&" conflict(s) to resolve on the class tabs","No conflicts")',
    'Not run yet'
  ]);
  put(7, [
    '3. Lock the specialist lessons',
    'Timetable coordinator',
    'Timetable Admin > Lock Synced Cells, once the sync shows no conflicts. Locked lessons can only be changed by the spreadsheet owner.',
    '=IF(SUM(' + conflicts + ')>0,"Resolve the conflicts first","Ready to lock after syncing")',
    'Not run yet'
  ]);
  put(8, [
    '4. Fill in the class timetables',
    'Class teachers',
    'Open your class tab (listed below) and fill every empty slot. Specialist lessons are already in place. To save typing, Timetable Admin can copy Week A into the other weeks, and Term 1 into Terms 2-4.',
    '=COUNTIF(' + emptySlots + ',0)&" of ' + classCount + ' class(es) have no empty slots"'
  ]);
  put(9, [
    '5. Check compliance',
    'Class teachers and coordinator',
    'Timetable Admin > Check this timetable (on a class tab) lists what to fix. OVERVIEW shows the whole school. Timetable Admin can also make a Google Doc of a timetable or of the compliance summary to print or share.',
    '=COUNTIF(' + compliant + ',"Yes")&" of ' + classCount + ' class(es) compliant"'
  ]);

  put(11, ['Specialist timetables']);
  put(12, ['Tab', 'KLA', 'Cycle', 'Lessons entered']);
  if (!specialistCount) put(specialistFirstRow, ['No specialists in this timetable']);
  const hasBreaks = config.breaks.length > 0;
  config.specialists.forEach((specialist, index) => {
    const name = getSpecialistTabName(specialist);
    const lessonFormula = hasBreaks
      ? '=SUMPRODUCT((LEN(' + grid(name) + ')>0)*ISNA(MATCH(' + grid(name) + ',BreakList,0)))'
      : '=COUNTA(' + grid(name) + ')';
    put(specialistFirstRow + index, [
      link(name), specialist.kla, specialist.cycle + ' week cycle', lessonFormula
    ]);
  });

  put(classHeadingRow, ['Class timetables']);
  put(classHeadingRow + 1, ['Tab', 'Teacher', 'Grade', 'Empty slots', 'Conflicts', 'Compliant']);
  config.classes.forEach((classConfig, index) => {
    const name = getClassTabName(classConfig);
    put(classFirstRow + index, [
      link(name),
      classConfig.teacher,
      classConfig.grade,
      '=' + quote(name) + '!C15',
      '=COUNTIF(' + grid(name) + ',"' + CONFLICT_VALUE_PREFIX + '*")',
      '=IF(' + quote(name) + '!C16=TRUE,"Yes","No")'
    ]);
  });

  // Google Docs made from the timetable are listed here by TimetableTools.gs
  // (named range StartHereDocs: the column headings row).
  put(docsHeadingRow, ['Google Docs']);
  put(docsHeadingRow + 1, ['Document', 'Made from', 'Made on', 'Made by']);
  put(docsHeadingRow + 2, [START_HERE_NO_DOCS_TEXT]);

  // Names typed in the wizard stay exactly as typed (no dates from "3-4");
  // the formula cells keep the normal format so they calculate.
  const textCells = ['A1', 'B' + specialistFirstRow + ':B' + specialistLastRow];
  if (classCount) textCells.push('B' + classFirstRow + ':C' + classLastRow);
  sheet.getRangeList(textCells).setNumberFormat('@');
  sheet.getRange(1, 1, rows.length, 6).setValues(rows);

  styleStartHereSheet(sheet, {
    lastRow: rows.length,
    headingRows: [4, 12, classHeadingRow + 1, docsHeadingRow + 1],
    sectionRows: [11, classHeadingRow, docsHeadingRow],
    tableRanges: ['A5:E9', 'A' + specialistFirstRow + ':D' + specialistLastRow,
      'A' + classFirstRow + ':F' + Math.max(classFirstRow, classLastRow)],
    docsFirstRow: docsHeadingRow + 2,
    classRows: classCount ? { first: classFirstRow, last: classLastRow } : null
  });

  ss.setNamedRange('StartHereSyncStatus', sheet.getRange('E6'));
  ss.setNamedRange('StartHereLockStatus', sheet.getRange('E7'));
  ss.setNamedRange('StartHereDocs', sheet.getRange('A' + (docsHeadingRow + 1) + ':D' + (docsHeadingRow + 1)));
}


/** Start here's colours: a navy title band, light grey headings, thin dividers. */
const START_HERE_STYLE = {
  font: 'Roboto',
  text: '#1F2933',
  muted: '#5F6B7A',
  accent: '#1E3A5F',
  accentText: '#FFFFFF',
  accentMuted: '#C9D6E8',
  headingBackground: '#EEF2F6',
  headingBorder: '#CBD5E0',
  divider: '#E5E9F0',
  good: { background: '#E3F4E8', font: '#1E7B3A' },
  bad: { background: '#FDECEC', font: '#B42318' }
};


/**
 * Gives the Start here tab a clean look: no gridlines, a title band, section
 * headings in the accent colour, light table headings, thin row dividers, and
 * green or red status chips (Yes/No, conflicts, empty slots) that update live.
 */
function styleStartHereSheet(sheet, layout) {
  const style = START_HERE_STYLE;
  const solid = SpreadsheetApp.BorderStyle.SOLID;
  const lastRow = Math.max(layout.lastRow, layout.docsFirstRow + 20);

  sheet.setHiddenGridlines(true);
  sheet.setTabColor(style.accent);
  sheet.getRange(1, 1, lastRow, 6)
    .setFontFamily(style.font).setFontSize(10).setFontColor(style.text)
    .setVerticalAlignment('middle');

  // Title band.
  sheet.getRange('A1:F1').merge();
  sheet.getRange('A2:F2').merge();
  sheet.getRange('A1:F2').setBackground(style.accent);
  sheet.getRange('A1').setFontSize(20).setFontWeight('bold').setFontColor(style.accentText);
  sheet.getRange('A2').setFontColor(style.accentMuted);
  sheet.setRowHeights(1, 1, 52);
  sheet.setRowHeights(2, 1, 28);
  sheet.setRowHeights(3, 1, 16);
  sheet.setFrozenRows(2);

  // Section headings and table headings.
  sheet.getRangeList(layout.sectionRows.map(row => 'A' + row))
    .setFontSize(14).setFontWeight('bold').setFontColor(style.accent).setVerticalAlignment('bottom');
  layout.sectionRows.forEach(row => sheet.setRowHeights(row, 1, 40));
  sheet.getRangeList(layout.headingRows.map(row => 'A' + row + ':F' + row))
    .setBackground(style.headingBackground).setFontColor(style.muted).setFontWeight('bold').setFontSize(9)
    .setBorder(null, null, true, null, null, null, style.headingBorder, solid);
  layout.headingRows.forEach(row => sheet.setRowHeights(row, 1, 28));

  // Rows: thin dividers; the stages wrap, the lists stay one line each.
  sheet.getRangeList(layout.tableRanges.concat(['A' + layout.docsFirstRow + ':D' + lastRow]))
    .setBorder(null, null, true, null, null, true, style.divider, solid);
  sheet.getRange('A5:E9').setVerticalAlignment('top').setWrap(true);
  sheet.getRange('A5:A9').setFontWeight('bold').setFontColor(style.accent);
  sheet.getRange('B5:B9').setFontColor(style.muted);
  sheet.getRange('D5:D9').setFontWeight('bold');
  sheet.getRange('E5:E9').setFontColor(style.muted);
  sheet.getRange('A' + layout.docsFirstRow).setFontColor(style.muted);

  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 170);
  sheet.setColumnWidth(3, 440);
  sheet.setColumnWidth(4, 260);
  sheet.setColumnWidth(5, 220);
  sheet.setColumnWidth(6, 110);

  // Live status chips.
  const chip = (range, formula, colours) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground(colours.background).setFontColor(colours.font)
    .setRanges([sheet.getRange(range)])
    .build();
  const rules = [
    chip('D6', '=ISNUMBER(SEARCH("to resolve",D6))', style.bad),
    chip('D6', '=D6="No conflicts"', style.good),
    chip('D7', '=ISNUMBER(SEARCH("Resolve",D7))', style.bad)
  ];
  if (layout.classRows) {
    const first = layout.classRows.first;
    const last = layout.classRows.last;
    rules.push(
      chip('D' + first + ':D' + last, '=AND(ISNUMBER(D' + first + '),D' + first + '=0)', style.good),
      chip('E' + first + ':E' + last, '=AND(ISNUMBER(E' + first + '),E' + first + '>0)', style.bad),
      chip('F' + first + ':F' + last, '=F' + first + '="Yes"', style.good),
      chip('F' + first + ':F' + last, '=F' + first + '="No"', style.bad)
    );
    sheet.getRange('D' + first + ':F' + last).setHorizontalAlignment('center');
  }
  sheet.setConditionalFormatRules(rules);
}


function getSpecialistTabName(specialist) {
  return ('SPEC_' + sanitizeSheetName(specialist.name)).substring(0, 99);
}


function getClassTabName(classConfig) {
  return sanitizeSheetName(classConfig.name).substring(0, 99);
}


/**
 * Lays out the timetable grid on a freshly copied TEMPLATE or SPEC_Template
 * tab: times, term/week headers, day names, column widths, the template's
 * cell formatting, dropdowns and break rows. Writes are batched (one call
 * per row or block rather than per cell) to stay well inside Apps Script's
 * time limit.
 */
function buildGrid(sheet, formatSource, cycle, slots, breakLabels, validationRule) {
  const layout = getGridLayout(cycle);
  const lastDataRow = FIRST_DATA_ROW + slots.length - 1;
  const gridWidth = layout.lastColumn - FIRST_GRID_COLUMN + 1;

  ensureSheetSize(sheet, Math.max(lastDataRow, 30), layout.lastColumn);
  sheet.getRange(
    4,
    TIME_COLUMN,
    sheet.getMaxRows() - 3,
    layout.lastColumn - TIME_COLUMN + 1
  ).clearContent().clearDataValidations();

  // The template's formatting for one week, repeated for every week block.
  const sourceLastRow = Math.max(5, Math.min(formatSource.getMaxRows(), lastDataRow));
  const sourceHeader = formatSource.getRange(4, FIRST_GRID_COLUMN, 1, DAYS.length);
  const sourceData = formatSource.getRange(5, FIRST_GRID_COLUMN, sourceLastRow - 4, DAYS.length);
  layout.blocks.forEach(block => {
    const blockLastColumn = block.startColumn + DAYS.length - 1;
    sourceHeader.copyFormatToRange(sheet, block.startColumn, blockLastColumn, 4, 4);
    sourceData.copyFormatToRange(sheet, block.startColumn, blockLastColumn, FIRST_DATA_ROW, lastDataRow);
  });

  sheet.getRange(FIRST_DATA_ROW, TIME_COLUMN, slots.length, 1)
    .setValues(slots.map(slot => [slot.label]));
  sheet.setColumnWidth(TIME_COLUMN, 100);

  // Row 1 (term/week headers), row 4 (day names) and the grid itself
  // (blank, apart from break labels) are each written in one call.
  const headerRow = new Array(gridWidth).fill('');
  const dayRow = new Array(gridWidth).fill('');
  const gridValues = slots.map(() => new Array(gridWidth).fill(''));
  layout.blocks.forEach(block => {
    const offset = block.startColumn - FIRST_GRID_COLUMN;
    headerRow[offset] = block.label;
    DAYS.forEach((day, dayIndex) => {
      dayRow[offset + dayIndex] = day;
      breakLabels.forEach((label, slotIndex) => {
        if (label) gridValues[slotIndex][offset + dayIndex] = label;
      });
    });
  });

  sheet.getRange(1, FIRST_GRID_COLUMN, 1, gridWidth).breakApart().setValues([headerRow]);
  layout.blocks.forEach(block => {
    sheet.getRange(1, block.startColumn, 1, DAYS.length).merge();
  });
  sheet.getRangeList(layout.blocks.map(block => toA1(1, block.startColumn, 1, DAYS.length)))
    .setFontSize(12)
    .setFontFamily('Poppins')
    .setFontWeight('bold');

  sheet.getRange(4, FIRST_GRID_COLUMN, 1, gridWidth).setValues([dayRow]);
  sheet.getRangeList(layout.blocks.map(block => toA1(4, block.startColumn, 1, DAYS.length)))
    .setFontWeight('bold');
  sheet.getRange(FIRST_DATA_ROW, FIRST_GRID_COLUMN, slots.length, gridWidth)
    .setValues(gridValues);

  layout.blocks.forEach(block => {
    sheet.setColumnWidths(block.startColumn, DAYS.length, DAY_COLUMN_WIDTH);
    sheet.getRange(FIRST_DATA_ROW, block.startColumn, slots.length, DAYS.length)
      .setDataValidation(validationRule);
  });
  layout.spacers.forEach(spacer => {
    sheet.setColumnWidth(spacer.column, spacer.width);
  });

  // Break rows: no dropdown, grey, centred. One range per block and run of
  // consecutive break slots.
  const breakRanges = [];
  getBreakRuns(breakLabels).forEach(run => {
    layout.blocks.forEach(block => {
      breakRanges.push(
        toA1(FIRST_DATA_ROW + run.startIndex, block.startColumn, run.length, DAYS.length)
      );
    });
  });
  if (breakRanges.length) {
    sheet.getRangeList(breakRanges)
      .clearDataValidations()
      .setBackground(BREAK_BACKGROUND)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }

  return { layout: layout, lastDataRow: lastDataRow };
}


/**
 * Copies only the conditional formatting from TEMPLATE!B5 to every week
 * block of a class tab (one paste per block, so the spacer columns between
 * weeks and terms do not receive timetable conditional formatting).
 */
function pasteTemplateConditionalFormatting(templateSheet, sheet, layout, slotCount) {
  const sourceRange = templateSheet.getRange('B5');
  layout.blocks.forEach(block => {
    sourceRange.copyTo(
      sheet.getRange(FIRST_DATA_ROW, block.startColumn, slotCount, DAYS.length),
      SpreadsheetApp.CopyPasteType.PASTE_CONDITIONAL_FORMATTING,
      false
    );
  });
}


/**
 * Column layout of the grid for a cycle length: a block of five day columns
 * per term and week, a narrow spacer between weeks and a wider one between
 * terms.
 */
function getGridLayout(cycle) {
  const blocks = [];
  const spacers = [];
  let column = FIRST_GRID_COLUMN;

  for (let term = 1; term <= TERM_COUNT; term++) {
    for (let week = 0; week < cycle; week++) {
      blocks.push({
        term: term,
        week: week,
        label: 'Term ' + term + ' - Week ' + String.fromCharCode(65 + week),
        startColumn: column
      });
      column += DAYS.length;
      if (week < cycle - 1) {
        spacers.push({ column: column, width: WEEK_SPACER_WIDTH });
        column++;
      }
    }
    if (term < TERM_COUNT) {
      spacers.push({ column: column, width: TERM_SPACER_WIDTH });
      column++;
    }
  }

  return { blocks: blocks, spacers: spacers, lastColumn: column - 1 };
}


/** Every 5-minute slot of the school day, labelled "HH:MM - HH:MM". */
function createTimeSlots(startTime, endTime) {
  const start = timeToMinutes(startTime);
  const end = timeToMinutes(endTime);

  if (start === null || end === null || start >= end) {
    throw new Error('Invalid school start or end time.');
  }

  const slots = [];
  for (let minutes = start; minutes < end; minutes += SLOT_MINUTES) {
    slots.push({
      start: minutes,
      end: minutes + SLOT_MINUTES,
      label: minutesToTime(minutes) + ' - ' + minutesToTime(minutes + SLOT_MINUTES)
    });
  }
  return slots;
}


/** The break name for each slot that overlaps a break, '' for the rest. */
function getBreakLabels(slots, breaks) {
  return slots.map(slot => {
    const match = breaks.find(item =>
      slot.start < timeToMinutes(item.end) && slot.end > timeToMinutes(item.start)
    );
    return match ? match.label : '';
  });
}


/** Runs of consecutive break slots, as {startIndex, length}. */
function getBreakRuns(breakLabels) {
  const runs = [];
  breakLabels.forEach((label, index) => {
    if (!label) return;
    const lastRun = runs[runs.length - 1];
    if (lastRun && lastRun.startIndex + lastRun.length === index) {
      lastRun.length++;
    } else {
      runs.push({ startIndex: index, length: 1 });
    }
  });
  return runs;
}


function minutesToTime(totalMinutes) {
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  return (hours < 10 ? '0' : '') + hours + ':' + (minutes < 10 ? '0' : '') + minutes;
}


/** A1 notation for a range, e.g. toA1(1, 10, 1, 5) is "J1:N1". */
function toA1(row, column, numRows, numColumns) {
  const start = columnToLetters(column) + row;
  if (numRows === 1 && numColumns === 1) return start;
  return start + ':' + columnToLetters(column + numColumns - 1) + (row + numRows - 1);
}


function columnToLetters(column) {
  let letters = '';
  while (column > 0) {
    const remainder = (column - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    column = Math.floor((column - 1) / 26);
  }
  return letters;
}


/**
 * Applies conditional formatting mapping each class's chosen colour onto a
 * specialist sheet, both in the roster helper column (B6:B, where classes
 * assigned to that specialist are listed via a preset formula) and across
 * the actual schedule grid (from the first time-slot/day cell through the
 * last used column and row), so that any cell showing e.g. "Kindergarten-1"
 * is shaded with that class's chosen colour.
 */
function applyClassColorFormatting(sheet, classes, firstDataRow, lastDataRow, lastColumn) {
  const validClasses = (classes || []).filter(
    cls => cls.name && isValidHexColor(cls.color)
  );
  if (!validClasses.length) return;

  const rosterLastRow = Math.max(6, sheet.getMaxRows());
  const rosterRange = sheet.getRange(6, 2, rosterLastRow - 5, 1); // B6:B(lastRow)
  const rosterAnchor = rosterRange.getCell(1, 1).getA1Notation();

  const ranges = [{ range: rosterRange, anchor: rosterAnchor }];

  if (lastColumn >= 10 && lastDataRow >= firstDataRow) {
    const gridRange = sheet.getRange(
      firstDataRow,
      10,
      lastDataRow - firstDataRow + 1,
      lastColumn - 9
    ); // J5 (or equivalent) through the last used column/row
    const gridAnchor = gridRange.getCell(1, 1).getA1Notation();
    ranges.push({ range: gridRange, anchor: gridAnchor });
  }

  const colorRules = [];
  validClasses.forEach(cls => {
    const escapedName = cls.name.replace(/"/g, '""');
    ranges.forEach(({ range, anchor }) => {
      colorRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=' + anchor + '="' + escapedName + '"')
          .setBackground(cls.color)
          .setRanges([range])
          .build()
      );
    });
  });

  const existingRules = sheet.getConditionalFormatRules();
  sheet.setConditionalFormatRules(existingRules.concat(colorRules));
}

function isValidHexColor(value) {
  return typeof value === 'string' && /^#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$/.test(value.trim());
}


function ensureSheetSize(sheet, requiredRows, requiredColumns) {
  const missingRows = requiredRows - sheet.getMaxRows();
  if (missingRows > 0) {
    sheet.insertRowsAfter(sheet.getMaxRows(), missingRows);
  }

  const missingColumns = requiredColumns - sheet.getMaxColumns();
  if (missingColumns > 0) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), missingColumns);
  }
}

function uniqueNonBlank(values) {
  const seen = {};
  return values.filter(value => {
    const cleaned = String(value || '').trim();
    const key = cleaned.toLowerCase();
    if (!cleaned || seen[key]) return false;
    seen[key] = true;
    return true;
  }).map(value => String(value).trim());
}


// ---- shared/timetableTabs.gs (copied here by "npm run build"; edit that file, not this block) ----
/*
 * Shared by the wizard (wizard/Code.gs) and the template's own script
 * (template/TemplateSync.gs): "npm run build" copies this file into both.
 * Edit it here, run the build, then paste both files into Apps Script.
 */

// Notes and text the sync puts on class tabs. The wizard reads them too
// (the Start here tab counts conflicts; an update moves synced lessons when
// a specialist is renamed or archived).
const SYNC_NOTE_PREFIX = '[SPEC_SYNC]';
const CONFLICT_NOTE_PREFIX = '[SPEC_CONFLICT]';
const CONFLICT_VALUE_PREFIX = 'CONFLICT:';

/**
 * Generated specialist tabs are named
 * "SPEC_<name>"; SPEC_Template is the master copy and is never one of them.
 * Compared case-insensitively, like Google Sheets tab names.
 */
function isSpecialistSheet(sheet) {
  const name = sheet.getName().toUpperCase();
  return name.indexOf('SPEC_') === 0 && name !== 'SPEC_TEMPLATE';
}

/**
 * The wizard applies this when generating a timetable, and the Timetable
 * Admin menu re-applies it later.
 *
 * Adds two conditional-formatting rules to every specialist tab: break rows
 * in grey, and a live red clash warning when another specialist has the
 * same class at the same time in the same real week. Specialists can have
 * different cycle lengths, so their columns don't line up; the clash rule
 * works out which cells to compare (see buildClashFormula), the same way
 * "Sync Specialist Schedules" does.
 */
function applySpecialistClashFormatting(ss, allSpecialists) {
  const specialistSheets = ss.getSheets()
    .filter(isSpecialistSheet);

  if (!specialistSheets.length) return;

  const breakListRange = ss.getRangeByName('BreakList');
  // Every specialist tab's name and cycle length (from D3, written by the
  // wizard). The wizard passes the full list instead, because it formats the
  // first tab of each cycle length before the rest are copied from it.
  const cycleOf = sheet => Math.max(1, parseInt(sheet.getRange('D3').getValue(), 10) || 1);
  const specialists = allSpecialists || specialistSheets.map(sheet => ({
    name: sheet.getName(),
    cycle: cycleOf(sheet)
  }));

  specialistSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (lastRow < 5 || lastColumn < 10) return;

    // Day-name columns only (spacer columns have no heading in row 4).
    const dayHeadings = sheet.getRange(4, 10, 1, lastColumn - 9).getDisplayValues()[0];
    const scheduleRanges = [];
    dayHeadings.forEach((heading, index) => {
      if (String(heading).trim()) {
        scheduleRanges.push(sheet.getRange(5, index + 10, lastRow - 4, 1));
      }
    });

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

    if (specialists.length >= 2) {
      newRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(
            buildClashFormula(referenceCell, cycleOf(sheet), specialists)
          )
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

/**
 * The clash rule's formula for a specialist tab with the given cycle length.
 * It counts how many specialist tabs (this one included) have the same
 * class at the same time, on the same day, in the same term and in a week
 * that is the same real week, and is true when that count is above 1.
 *
 * Grid layout (from the wizard): columns start at J (10); each week block is
 * 5 day columns plus a spacer, so 6 columns, and a term is 6 x cycle
 * columns wide. For the cell being checked, COLUMN() gives its term, week
 * and day. Cycles restart each term, so week w of a C-week cycle falls on
 * real weeks w, w + C, w + 2C, ...; it shares a real week with week x of a
 * D-week cycle when w and x are equal modulo gcd(C, D). Those are the cells
 * compared on each other tab. With equal cycle lengths that is just the
 * cell in the same position.
 */
function buildClashFormula(referenceCell, cycle, specialists) {
  const gcd = (a, b) => (b ? gcd(b, a % b) : a);
  const offset = '(COLUMN()-10)';
  const term = 'INT(' + offset + '/' + (6 * cycle) + ')';
  const week = 'INT(MOD(' + offset + ',' + (6 * cycle) + ')/6)';
  const day = 'MOD(' + offset + ',6)';

  const counts = [];
  specialists.forEach(other => {
    const sheetName = "'" + other.name.replace(/'/g, "''").replace(/"/g, '""') + "'!";
    if (other.cycle === cycle) {
      counts.push('COUNTIF(INDIRECT("' + sheetName + '"&ADDRESS(ROW(),COLUMN(),4)),' + referenceCell + ')');
      return;
    }
    const shared = gcd(cycle, other.cycle);
    const firstWeek = shared === 1 ? '0' : 'MOD(' + week + ',' + shared + ')';
    for (let k = 0; k < other.cycle / shared; k++) {
      const otherWeek = firstWeek + '+' + (k * shared);
      const column = '10+' + term + '*' + (6 * other.cycle) + '+6*(' + otherWeek + ')+' + day;
      counts.push('COUNTIF(INDIRECT("' + sheetName + '"&ADDRESS(ROW(),' + column + ',4)),' + referenceCell + ')');
    }
  });

  return '=AND(LEN(' + referenceCell + ')>0,(' + counts.join('+') + ')>1)';
}

/** A tab name with the characters Google Sheets doesn't allow replaced. */
function sanitizeSheetName(name) {
  const sanitized = String(name || '')
    .replace(/[\/\\?*\[\]:]/g, '_')
    .trim();

  return sanitized || 'Untitled';
}
// ---- end of shared/timetableTabs.gs ----
