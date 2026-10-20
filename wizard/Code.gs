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


    const config = parsePayloadConfig(payload);
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


    if (name && isValidTime(start) && isValidTime(end)) {
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


  const specialists = [];
  index = 1;
  while (varMap['Specialist_' + index + '_Name'] !== undefined) {
    const name = String(varMap['Specialist_' + index + '_Name'] || '').trim();
    const kla = String(varMap['Specialist_' + index + '_KLA'] || '').trim();
    const cycle = Math.max(
      1,
      parseInt(varMap['Specialist_' + index + '_Cycle'], 10) || 1
    );
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
    const cycle = Math.max(
      1,
      parseInt(varMap['Class_' + index + '_Cycle'], 10) || 1
    );
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
    startTime: varMap.SchoolStart || '08:30',
    endTime: varMap.SchoolEnd || '15:15',
    breaks: breaks,
    klas: klas,
    activities: activities,
    specialists: specialists,
    classes: classes
  };
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
  const activityNames = uniqueNonBlank(config.activities.map(item => item.name));
  const breakNames = uniqueNonBlank(config.breaks.map(item => item.label));


  const requiredRows = Math.max(
    2,
    variableRows.length + 1,
    klaNames.length + 1,
    classNames.length + 1,
    specialistRows.length + 1,
    classColorRows.length + 1,
    activityNames.length + 1,
    breakNames.length + 1
  );


  ensureSheetSize(dataSheet, requiredRows, 16);


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


  dataSheet.getRange('N1').setValue('ActivityList');
  if (activityNames.length) {
    const activityValues = activityNames.map(name => [name]);
    dataSheet.getRange(2, 14, activityValues.length, 1).setValues(activityValues);
  }


  dataSheet.getRange('P1').setValue('Breaks');
  if (breakNames.length) {
    const breakValues = breakNames.map(name => [name]);
    dataSheet.getRange(2, 16, breakValues.length, 1).setValues(breakValues);
  }
}


function buildSpecialistSheets(ss, config, jobId) {
  const specTemplateSheet = ss.getSheetByName('SPEC_TEMPLATE');
  const classListRange = ss.getRangeByName('ClassList');


  if (!config.specialists.length) return;
  if (!specTemplateSheet) {
    throw new Error("Could not find a sheet named 'SPEC_TEMPLATE' in the template.");
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
    // rules that happen to exist on SPEC_TEMPLATE (e.g. a leftover "cell is
    // not empty" rule). Only the rules this script adds below should apply.
    sheet.clearConditionalFormatRules();


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


    sheet.getRange('C27:C30').setValues([
      [classConfig.teacher],
      [classConfig.name],
      [classConfig.grade],
      [classConfig.cycle + ' week cycle']
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
 * NOTE ON CYCLE MISMATCHES: this live conditional-formatting clash check
 * compares specialist sheets column-for-column, which is only guaranteed
 * correct when the specialist sheets being compared share the same cycle
 * length. It still gives useful, immediate feedback while specialists are
 * filling in their sheets, but it is a heuristic, not the source of truth.
 * The authoritative, cycle-aware clash + conflict check runs via
 * "Sync & Lock Specialist Schedules" (see TemplateSync.gs, which is bound
 * directly to the master template spreadsheet and travels with every copy
 * the wizard generates), which resolves real repeating weeks via each
 * sheet's own cycle length before comparing anything. Re-run Sync & Lock
 * periodically (or after specialists finish editing) to catch clashes this
 * live check might miss when cycles differ.
 */
function applySpecialistClashFormatting(ss) {
  const specialistSheets = ss.getSheets()
    .filter(sheet => sheet.getName().startsWith('SPEC_') && sheet.getName() !== 'SPEC_TEMPLATE');

  if (specialistSheets.length < 2) return;

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
    const counts = specialistSheets.map(specialistSheet => {
      const escapedName = specialistSheet.getName().replace(/'/g, "''");
      return 'COUNTIF(INDIRECT("\'' + escapedName +
        '\'!"&ADDRESS(ROW(),COLUMN(),4)),' + referenceCell + ')';
    }).join('+');

    const formula = '=AND(LEN(' + referenceCell + ')>0,(' + counts + ')>1)';

    const clashRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground('#FF0000')
      .setFontColor('#FFFFFF')
      .setBold(true)
      .setRanges(scheduleRanges)
      .build();

    const retainedRules = sheet.getConditionalFormatRules().filter(rule => {
      const condition = rule.getBooleanCondition();
      if (!condition) return true;
      return !condition.getCriteriaValues().some(value =>
        typeof value === 'string' &&
        value.indexOf('INDIRECT') !== -1 &&
        value.indexOf('COUNTIF') !== -1
      );
    });

    retainedRules.unshift(clashRule);
    sheet.setConditionalFormatRules(retainedRules);
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
