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


function generateTimetable(payload) {
  const jobId = payload ? payload.jobId : null;


  try {
    // Check the setup before copying anything, so a bad setup doesn't leave
    // a half-built file in the user's Drive.
    const config = parsePayloadConfig(payload);
    validateConfig(config);

    updateJobProgress(jobId, 5, 'Initializing template duplication...');


    const templateFile = DriveApp.getFileById(getTemplateId());
    const timeStamp = new Date().toLocaleDateString('en-AU');
    const newFile = templateFile.makeCopy(
      'Generated Timetable - Resource First - ' + timeStamp
    );


    updateJobProgress(jobId, 12, 'Opening duplicated spreadsheet...');


    const ss = SpreadsheetApp.openById(newFile.getId());
    const dataSheet = ss.getSheetByName('DATA');


    if (!dataSheet) {
      throw new Error("Could not find a sheet named 'DATA' in the template.");
    }


    updateJobProgress(jobId, 18, 'Building DATA reference tables...');


    populateDataSheet(dataSheet, payload, config, ss);
    buildSpecialistSheets(ss, config, jobId);
    buildAllClassSheets(ss, config, jobId);


    updateJobProgress(jobId, 100, 'Timetable generation complete!');
    return ss.getUrl();
  } catch (error) {
    console.error('Error generating timetable:', error);
    updateJobProgress(jobId, -1, 'Error: ' + error.toString());
    throw new Error(error.toString());
  }
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


function buildSpecialistSheets(ss, config, jobId) {
  const specTemplateSheet = ss.getSheetByName('SPEC_Template');
  const classListRange = ss.getRangeByName('ClassList');


  if (!config.specialists.length) return;
  if (!specTemplateSheet) {
    throw new Error("Could not find a sheet named 'SPEC_Template' in the template.");
  }
  if (!classListRange) {
    throw new Error("Named range 'ClassList' was not created.");
  }


  const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
  const timeSlots = createTimeSlots(
    config.startTime,
    config.endTime,
    ss.getSpreadsheetTimeZone()
  );
  const firstDataRow = 5;
  const lastDataRow = firstDataRow + timeSlots.length - 1;


  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(classListRange, true)
    .setAllowInvalid(false)
    .build();


  config.specialists.forEach((specialist, specialistIndex) => {
    const sheetName = (
      'SPEC_' + sanitizeSheetName(specialist.name)
    ).substring(0, 99);


    updateJobProgress(
      jobId,
      25 + (specialistIndex / config.specialists.length) * 20,
      'Generating Specialist Tab ' + (specialistIndex + 1) + ' of ' +
        config.specialists.length + ': ' + specialist.name
    );


    if (ss.getSheetByName(sheetName)) return;


    const columnsPerTerm =
      specialist.cycle * days.length + Math.max(0, specialist.cycle - 1);
    const totalScheduleColumns = 4 * columnsPerTerm + 3;
    const lastRequiredColumn = 9 + totalScheduleColumns;


    const sheet = specTemplateSheet.copyTo(ss);
    sheet.setName(sheetName).showSheet();
    ensureSheetSize(sheet, Math.max(lastDataRow, 30), lastRequiredColumn);


    // Start from a clean slate: don't inherit any conditional formatting
    // rules that happen to exist on SPEC_Template (e.g. a leftover "cell is
    // not empty" rule). Only the rules this script adds below should apply.
    sheet.clearConditionalFormatRules();


    sheet.getRangeList(['B1', 'B3']).setNumberFormat('@');
    sheet.getRange('B1').setValue(specialist.name);
    sheet.getRange('B3').setValue(specialist.kla);
    sheet.getRange('D3').setValue(specialist.cycle);


    sheet.getRange(
      4,
      9,
      sheet.getMaxRows() - 3,
      lastRequiredColumn - 8
    ).clearContent().clearDataValidations();


    sheet.getRange(firstDataRow, 9, timeSlots.length, 1).setValues(timeSlots);
    sheet.setColumnWidth(9, 100);


    let currentColumn = 10;
    for (let term = 1; term <= 4; term++) {
      for (let week = 0; week < specialist.cycle; week++) {
        const blockStartColumn = currentColumn;
        const weekLabel = String.fromCharCode(65 + week);


        applyTemplateGridFormatting(
          specTemplateSheet,
          sheet,
          blockStartColumn,
          firstDataRow,
          lastDataRow,
          days.length
        );


        sheet.getRange(1, blockStartColumn, 1, days.length)
          .breakApart()
          .merge()
          .setValue('Term ' + term + ' - Week ' + weekLabel)
          .setFontSize(12)
          .setFontFamily('Poppins')
          .setFontWeight('bold');


        for (let dayIndex = 0; dayIndex < days.length; dayIndex++) {
          sheet.getRange(4, currentColumn)
            .setValue(days[dayIndex])
            .setFontWeight('bold');
          sheet.setColumnWidth(currentColumn, 120);
          currentColumn++;
        }


        sheet.getRange(
          firstDataRow,
          blockStartColumn,
          timeSlots.length,
          days.length
        ).setDataValidation(validationRule);


        if (week < specialist.cycle - 1) {
          sheet.setColumnWidth(currentColumn, 20);
          currentColumn++;
        }
      }


      if (term < 4) {
        sheet.setColumnWidth(currentColumn, 40);
        currentColumn++;
      }
    }


    applyBreakRows(
      sheet,
      specialist,
      config.breaks,
      days,
      firstDataRow,
      lastDataRow
    );


    applyClassColorFormatting(
      sheet,
      config.classes,
      firstDataRow,
      lastDataRow,
      lastRequiredColumn
    );
  });


  applySpecialistClashFormatting(ss);
}


function buildAllClassSheets(ss, config, jobId) {
  const templateSheet = ss.getSheetByName('TEMPLATE');
  const klaListRange = ss.getRangeByName('KLAList');


  if (!templateSheet || !config.classes.length) return;
  if (!klaListRange) {
    throw new Error("Named range 'KLAList' was not created.");
  }


  const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
  const timeSlots = createTimeSlots(
    config.startTime,
    config.endTime,
    ss.getSpreadsheetTimeZone()
  );
  const firstDataRow = 5;
  const lastDataRow = firstDataRow + timeSlots.length - 1;


  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(klaListRange, true)
    .setAllowInvalid(false)
    .build();


  config.classes.forEach((classConfig, classIndex) => {
    const sheetName = sanitizeSheetName(classConfig.name).substring(0, 99);
    if (ss.getSheetByName(sheetName)) return;


    updateJobProgress(
      jobId,
      45 + (classIndex / config.classes.length) * 50,
      'Generating Classroom Tab ' + (classIndex + 1) + ' of ' +
        config.classes.length + ': ' + classConfig.name
    );


    const columnsPerTerm =
      classConfig.cycle * days.length + Math.max(0, classConfig.cycle - 1);
    const totalScheduleColumns = 4 * columnsPerTerm + 3;
    const lastRequiredColumn = 9 + totalScheduleColumns;


    const sheet = templateSheet.copyTo(ss);
    sheet.setName(sheetName).showSheet();
    ensureSheetSize(sheet, Math.max(lastDataRow, 30), lastRequiredColumn);


    sheet.getRange('C27:C29').setNumberFormat('@');
    sheet.getRange('C27:C30').setValues([
      [classConfig.teacher],
      [classConfig.name],
      [classConfig.grade],
      [classConfig.cycle + ' week cycle']
    ]);

    // C18 (Amadeus compliance) reads C31 to decide whether Amadeus applies.
    sheet.getRange('B31:C31').setValues([
      ['Amadeus', isAmadeusGrade(classConfig.grade)]
    ]);


    sheet.getRange(
      4,
      9,
      sheet.getMaxRows() - 3,
      lastRequiredColumn - 8
    ).clearContent().clearDataValidations();


    sheet.getRange(firstDataRow, 9, timeSlots.length, 1).setValues(timeSlots);
    sheet.setColumnWidth(9, 100);


    let currentColumn = 10;
    for (let term = 1; term <= 4; term++) {
      for (let week = 0; week < classConfig.cycle; week++) {
        const blockStartColumn = currentColumn;
        const weekLabel = String.fromCharCode(65 + week);


        applyTemplateGridFormatting(
          templateSheet,
          sheet,
          blockStartColumn,
          firstDataRow,
          lastDataRow,
          days.length
        );


        sheet.getRange(1, blockStartColumn, 1, days.length)
          .breakApart()
          .merge()
          .setValue('Term ' + term + ' - Week ' + weekLabel)
          .setFontSize(12)
          .setFontFamily('Poppins')
          .setFontWeight('bold');


        for (let dayIndex = 0; dayIndex < days.length; dayIndex++) {
          sheet.getRange(4, currentColumn)
            .setValue(days[dayIndex])
            .setFontWeight('bold');
          sheet.setColumnWidth(currentColumn, 120);
          currentColumn++;
        }


        sheet.getRange(
          firstDataRow,
          blockStartColumn,
          timeSlots.length,
          days.length
        ).setDataValidation(validationRule);


        if (week < classConfig.cycle - 1) {
          sheet.setColumnWidth(currentColumn, 20);
          currentColumn++;
        }
      }


      if (term < 4) {
        sheet.setColumnWidth(currentColumn, 40);
        currentColumn++;
      }
    }


    applyBreakRows(
      sheet,
      classConfig,
      config.breaks,
      days,
      firstDataRow,
      lastDataRow
    );


    copyConditionalFormattingToClassSheet(
      templateSheet,
      sheet,
      firstDataRow,
      lastDataRow
    );
  });
}


function applyTemplateGridFormatting(
  templateSheet,
  targetSheet,
  targetStartColumn,
  firstDataRow,
  lastDataRow,
  dayCount
) {
  const sourceLastRow = Math.max(
    5,
    Math.min(templateSheet.getMaxRows(), lastDataRow)
  );
  const sourceHeader = templateSheet.getRange(4, 10, 1, dayCount);
  const sourceData = templateSheet.getRange(
    5,
    10,
    sourceLastRow - 4,
    dayCount
  );


  sourceHeader.copyFormatToRange(
    targetSheet,
    targetStartColumn,
    targetStartColumn + dayCount - 1,
    4,
    4
  );


  sourceData.copyFormatToRange(
    targetSheet,
    targetStartColumn,
    targetStartColumn + dayCount - 1,
    firstDataRow,
    lastDataRow
  );
}


/**
 * Copies only the conditional formatting from TEMPLATE!B5 to all completed
 * timetable cells on a newly generated class sheet.
 *
 * The target sheet is passed directly because a web-app execution does not
 * reliably have the newly generated class tab as its active sheet.
 */
function copyConditionalFormattingToClassSheet(
  templateSheet,
  targetSheet,
  firstDataRow,
  lastDataRow
) {
  if (!templateSheet) {
    throw new Error("Could not find a sheet named 'TEMPLATE'.");
  }
  const sourceRange = templateSheet.getRange('B5');
  const startColumn = 10;
  const lastColumn = targetSheet.getLastColumn();
  if (lastDataRow < firstDataRow || lastColumn < startColumn) return;
  // Apply to each weekday column separately so spacer columns between weeks
  // and terms do not receive timetable conditional formatting.
  for (let column = startColumn; column <= lastColumn; column++) {
    const dayHeading = String(
      targetSheet.getRange(4, column).getDisplayValue()
    ).trim();

    if (!dayHeading) continue;
    const targetRange = targetSheet.getRange(
      firstDataRow,
      column,
      lastDataRow - firstDataRow + 1,
      1
    );

    sourceRange.copyTo(      targetRange,
      SpreadsheetApp.CopyPasteType.PASTE_CONDITIONAL_FORMATTING,      false    );
  }
}


function applyBreakRows(sheet, cycleConfig, breaks, days, firstDataRow, lastDataRow) {
  for (let row = firstDataRow; row <= lastDataRow; row++) {
    const timeRange = parseTimeRange(sheet.getRange(row, 9).getValue());
    if (!timeRange) continue;

    let breakLabel = '';
    for (let index = 0; index < breaks.length; index++) {
      if (
        timeRange.start < parseTime(breaks[index].end) &&
        timeRange.end > parseTime(breaks[index].start)
      ) {
        breakLabel = breaks[index].label;
        break;
      }
    }

    if (!breakLabel) continue;

    let currentColumn = 10;
    for (let term = 1; term <= 4; term++) {
      for (let week = 0; week < cycleConfig.cycle; week++) {
        sheet.getRange(row, currentColumn, 1, days.length)
          .setValue(breakLabel)
          .clearDataValidations()
          .setBackground('#F2F2F2')
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');

        currentColumn += days.length;
        if (week < cycleConfig.cycle - 1) currentColumn++;
      }
      if (term < 4) currentColumn++;
    }
  }
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
    .filter(isSpecialistSheet);

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

function createTimeSlots(startValue, endValue, timezone) {
  const start = parseTime(startValue);
  const end = parseTime(endValue);

  if (!start || !end || start >= end) {
    throw new Error('Invalid school start or end time.');
  }

  const slots = [];
  let current = start;

  while (current < end) {
    const next = new Date(current.getTime() + 5 * 60 * 1000);
    slots.push([formatTimeRange(current, next, timezone)]);
    current = next;
  }

  return slots;
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

function isValidTime(value) {
  return (value instanceof Date && !isNaN(value.getTime())) ||
    (typeof value === 'string' && /^(?:2[0-3]|[01]?[0-9]):[0-5][0-9]$/.test(value));
}

function parseTime(value) {
  if (value instanceof Date) return value;
  if (isValidTime(value)) return new Date('2000-01-01T' + value + ':00');
  return null;
}

function formatTimeRange(start, end, timezone) {
  return Utilities.formatDate(start, timezone, 'HH:mm') + ' - ' +
    Utilities.formatDate(end, timezone, 'HH:mm');
}

function parseTimeRange(value) {
  if (!value || typeof value !== 'string') return null;

  const parts = value.split(' - ');
  if (parts.length !== 2) return null;

  const start = parseTime(parts[0]);
  const end = parseTime(parts[1]);
  return start && end ? { start: start, end: end } : null;
}

function sanitizeSheetName(name) {
  const sanitized = String(name || '')
    .replace(/[\/\\?*\[\]:]/g, '_')
    .trim();

  return sanitized || 'Untitled';
}
