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

/**
 * Amadeus runs in Years 3-6 in every school. The activity name must match
 * TEMPLATE!B21 exactly, because C21 counts cells with this value.
 */
const AMADEUS_ACTIVITY_NAME = 'Amadeus - Tutorials';

function isAmadeusGrade(grade) {
  return /^Year\s*[3-6]$/i.test(String(grade || '').trim());
}


function checkProgress(jobId) {
  if (!jobId) return null;
  try {
    return CacheService.getScriptCache().get(jobId);
  } catch (error) {
    return null;
  }
}


function updateJobProgress(jobId, percent, statusText) {
  if (!jobId) return;
  try {
    const normalizedPercent = percent === -1
      ? -1
      : Math.min(100, Math.max(0, Math.round(percent)));


    CacheService.getScriptCache().put(jobId, JSON.stringify({
      percent: normalizedPercent,
      status: statusText
    }), 600);
  } catch (error) {
    console.warn('Failed to update cache progress:', error);
  }
}


const GENERATED_FILE_PREFIX = 'Generated Timetable - Resource First - ';
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
    const fileName = GENERATED_FILE_PREFIX + new Date().toLocaleDateString('en-AU');
    const newFile = templateFile.makeCopy(INCOMPLETE_FILE_PREFIX + fileName);
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
    const teacher = String(
      varMap['Class_' + index + '_Teacher'] || 'N/A'
    ).trim();
    const grade = String(varMap['Class_' + index + '_Grade'] || '').trim();
    const cycle = Number(varMap['Class_' + index + '_Cycle']);
    const color = String(varMap['Class_' + index + '_Color'] || '').trim();


    if (name) {
      classes.push({
        name: name,
        teacher: teacher || 'N/A',
        grade: grade,
        cycle: cycle,
        color: color
      });
    }
    index++;
  }


  return {
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
 * Rules shared with the wizard (Index.html), which checks the same things
 * before submitting. Keep the two in step.
 */
const SLOT_MINUTES = 5;
const MAX_CYCLE_WEEKS = 10;
const MAX_NAME_LENGTH = 90;
// Tabs that already exist in the template (or are planned), compared
// case-insensitively because Google Sheets tab names are.
const RESERVED_TAB_NAMES = ['data', 'template', 'overview', 'spec_template', 'start here'];

/**
 * Checks the parsed setup and throws one error listing every problem, so
 * nothing is generated from a setup that would produce missing or broken
 * tabs.
 */
function validateConfig(config) {
  const problems = [];

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
    if (startsWithFormula(item.teacher)) problems.push(label + ": the teacher name can't start with =.");
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
  const tabNames = config.specialists.map(specialist =>
    ('SPEC_' + sanitizeSheetName(specialist.name)).substring(0, 99)
  );


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
    const sheetName = sanitizeSheetName(classConfig.name).substring(0, 99);
    const existing = ss.getSheetByName(sheetName);

    if (existing) {
      writeClassDetails(existing, classConfig);
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
  return { nextIndex: sheets.length, usedIds: usedIds };
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

/**
 * Keep this function identical in wizard/Code.gs and
 * template/TemplateSync.gs. Generated specialist tabs are named
 * "SPEC_<name>"; SPEC_Template is the master copy and is never one of them.
 * Compared case-insensitively, like Google Sheets tab names.
 */
function isSpecialistSheet(sheet) {
  const name = sheet.getName().toUpperCase();
  return name.indexOf('SPEC_') === 0 && name !== 'SPEC_TEMPLATE';
}

/**
 * Keep this function identical in wizard/Code.gs and
 * template/TemplateSync.gs: the wizard applies it when generating a
 * timetable, and the Timetable Admin menu re-applies it later.
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
 * Keep this function identical in wizard/Code.gs and
 * template/TemplateSync.gs.
 *
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

function sanitizeSheetName(name) {
  const sanitized = String(name || '')
    .replace(/[\/\\?*\[\]:]/g, '_')
    .trim();

  return sanitized || 'Untitled';
}
