/**
 * Timetable tools for the Timetable Admin menu (see onOpen in
 * TemplateSync.gs): checking a timetable, copying weeks and terms, and
 * printing. Add this file to the Apps Script project bound to the master
 * template, next to TemplateSync.gs; it uses that file's helpers
 * (buildScheduleColumnMap, isSpecialistSheet, the note prefixes).
 *
 * Grid layout (from the wizard): row 1 "Term X - Week Y" headers, row 4 day
 * names, column I times, timetable from J5. Class tabs hold the compliance
 * table in B5:G21 (KLA rows 5-12, total in row 13, unallocated blocks C15,
 * compliant C16, Amadeus C18, Amadeus Tutorials row 21) and the class
 * details in C27:C31.
 */

const PRINT_SHEET_NAME = 'Print view';


// ---------------------------------------------------------------------------
// Check this timetable / Check all class timetables
// ---------------------------------------------------------------------------

function checkThisTimetable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getActiveSheet();

  if (isSpecialistSheet(sheet)) {
    showDialog(ui, 'Check: ' + sheet.getName(), describeSpecialistTimetable(sheet));
    return;
  }
  if (!isClassSheet(ss, sheet)) {
    ui.alert('Check this timetable', 'Open a class tab (or a specialist tab) first, then run this again.', ui.ButtonSet.OK);
    return;
  }

  const result = checkClassTimetable(ss, sheet);
  const sections = [];

  sections.push(result.compliant
    ? '<p class="ok"><b>Compliant.</b> Every KLA is within its range and there are no empty slots.</p>'
    : '<p class="bad"><b>Not compliant yet.</b> Fix the items below.</p>');

  if (!result.totalMinutes) {
    sections.push('<p>Nothing has been entered on this timetable yet.</p>');
  } else if (result.klaIssues.length) {
    sections.push('<h3>KLAs outside their range</h3><ul>' + result.klaIssues.map(issue =>
      '<li><b>' + escapeHtml(issue.kla) + '</b>: ' + issue.percent.toFixed(1) + '% of teaching time, needs ' +
      issue.min + '-' + issue.max + '%. ' + (issue.direction === 'under'
        ? 'About <b>' + formatDuration(issue.minutes) + ' more</b> per week.'
        : 'About <b>' + formatDuration(issue.minutes) + ' less</b> per week.') + '</li>'
    ).join('') + '</ul>');
  } else {
    sections.push('<p class="ok">Every KLA is within its range.</p>');
  }

  if (result.amadeus === 'N/A') {
    sections.push('<h3>Amadeus</h3><p>Doesn\'t apply to this grade.</p>');
  } else if (result.amadeusIssues.length) {
    sections.push('<h3>Amadeus</h3><ul>' + result.amadeusIssues.map(text => '<li>' + text + '</li>').join('') + '</ul>');
  } else {
    sections.push('<h3>Amadeus</h3><p class="ok">Creative Arts and Amadeus Tutorials are within range.</p>');
  }

  if (result.emptyBlocks.length) {
    sections.push('<h3>Empty slots (' + result.emptySlots + ')</h3><ul>' + result.emptyBlocks.map(block =>
      '<li>' + escapeHtml(block.label) + ': ' + block.count + ' slot(s), ' + formatDuration(block.count * 5) + '</li>'
    ).join('') + '</ul>');
  } else {
    sections.push('<h3>Empty slots</h3><p class="ok">None.</p>');
  }

  if (result.conflicts.length) {
    sections.push('<h3>Conflicts (' + result.conflicts.length + ')</h3><p>Move the specialist lesson or the class entry, ' +
      'then run "1. Sync Specialist Schedules" again.</p><ul>' + result.conflicts.slice(0, 10).map(conflict =>
      '<li>' + conflict.cell + ': ' + escapeHtml(conflict.text) + '</li>').join('') +
      (result.conflicts.length > 10 ? '<li>...and ' + (result.conflicts.length - 10) + ' more</li>' : '') + '</ul>');
  }

  showDialog(ui, 'Check: ' + sheet.getName(), sections.join(''));
}


function checkAllClassTimetables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const classSheets = getClassNames(ss).map(name => findClassSheet(ss, name)).filter(Boolean);
  if (!classSheets.length) {
    ui.alert('Check all class timetables', 'No class tabs were found.', ui.ButtonSet.OK);
    return;
  }

  ss.toast('Checking ' + classSheets.length + ' class timetable(s)...', 'Check', 5);
  const results = classSheets.map(sheet => ({ name: sheet.getName(), result: checkClassTimetable(ss, sheet) }));
  results.sort((a, b) => Number(a.result.compliant) - Number(b.result.compliant));
  const compliantCount = results.filter(item => item.result.compliant).length;

  const rows = results.map(item => {
    const r = item.result;
    const problems = r.klaIssues.map(issue => issue.kla + ' ' + issue.direction)
      .concat(r.amadeusIssues.length ? ['Amadeus'] : []);
    return '<tr><td>' + escapeHtml(item.name) + '</td>' +
      '<td class="' + (r.compliant ? 'ok' : 'bad') + '">' + (r.compliant ? 'Yes' : 'No') + '</td>' +
      '<td>' + r.emptySlots + '</td><td>' + r.conflicts.length + '</td>' +
      '<td>' + escapeHtml(problems.join(', ') || '-') + '</td></tr>';
  }).join('');

  showDialog(ui, 'All class timetables',
    '<p><b>' + compliantCount + ' of ' + results.length + '</b> class timetable(s) are compliant. ' +
    'Open a class tab and choose "Check this timetable" for the details.</p>' +
    '<table><tr><th>Class</th><th>Compliant</th><th>Empty slots</th><th>Conflicts</th><th>Outside range</th></tr>' +
    rows + '</table>');
}


/**
 * Reads a class tab's compliance table and grid. Returns what is out of
 * range (with roughly how many minutes per week to add or remove), the
 * empty slots per term and week, and any sync conflicts.
 */
function checkClassTimetable(ss, sheet) {
  const table = sheet.getRange('B5:G21').getValues(); // rows 5-21, columns B-G
  const number = value => Number(value) || 0;
  const totalMinutes = number(table[8][1]); // C13
  const roundUp = minutes => Math.max(5, Math.ceil(minutes / 5) * 5);

  const outOfRange = (label, minutes, percent, min, max) => {
    if (Math.round(percent * 10) / 10 < min) {
      return { kla: label, direction: 'under', percent, min, max, minutes: roundUp(min / 100 * totalMinutes - minutes) };
    }
    if (Math.round(percent * 10) / 10 > max) {
      return { kla: label, direction: 'over', percent, min, max, minutes: roundUp(minutes - max / 100 * totalMinutes) };
    }
    return null;
  };

  const klaIssues = [];
  for (let i = 0; i <= 7; i++) { // rows 5-12
    const row = table[i];
    const issue = outOfRange(String(row[0]), number(row[1]), number(row[3]), number(row[4]), number(row[5]));
    if (issue) klaIssues.push(issue);
  }

  const amadeus = table[13][1]; // C18
  const amadeusIssues = [];
  if (amadeus !== 'N/A' && amadeus !== true) {
    const creativeArts = klaIssues.find(issue => issue.kla === String(table[6][0]));
    if (creativeArts) amadeusIssues.push('Creative Arts is ' + creativeArts.direction + ' its range (see above).');
    const tutorials = table[16]; // row 21
    if (!number(tutorials[1])) {
      amadeusIssues.push('No <b>' + escapeHtml(String(tutorials[0])) + '</b> lessons yet.');
    } else {
      const issue = outOfRange(String(tutorials[0]), number(tutorials[1]), number(tutorials[3]),
        number(tutorials[4]), number(tutorials[5]));
      if (issue) {
        amadeusIssues.push(escapeHtml(issue.kla) + ' is ' + issue.percent.toFixed(1) + '%, needs ' + issue.min +
          '-' + issue.max + '%: about ' + formatDuration(issue.minutes) + (issue.direction === 'under' ? ' more' : ' less') +
          ' per week.');
      }
    }
  }

  const grid = readTimetableGrid(sheet);
  const emptyByBlock = {};
  const conflicts = [];
  grid.values.forEach((row, r) => {
    if (!grid.times[r]) return;
    grid.columns.forEach(column => {
      const value = String(row[column.index] || '').trim();
      if (!value) {
        emptyByBlock[column.blockLabel] = (emptyByBlock[column.blockLabel] || 0) + 1;
      } else if (value.indexOf(CONFLICT_VALUE_PREFIX) === 0) {
        conflicts.push({ cell: sheet.getRange(r + 5, column.column).getA1Notation(), text: value });
      }
    });
  });
  const emptyBlocks = grid.blockLabels
    .filter(label => emptyByBlock[label])
    .map(label => ({ label: label, count: emptyByBlock[label] }));

  return {
    compliant: table[11][1] === true, // C16
    totalMinutes: totalMinutes,
    klaIssues: klaIssues,
    amadeus: amadeus,
    amadeusIssues: amadeusIssues,
    emptySlots: emptyBlocks.reduce((sum, block) => sum + block.count, 0),
    emptyBlocks: emptyBlocks,
    conflicts: conflicts
  };
}


/** For specialist tabs: minutes per week with each class (the tab's own B/C columns). */
function describeSpecialistTimetable(sheet) {
  const rows = sheet.getRange(6, 2, Math.max(1, sheet.getLastRow() - 5), 2).getValues()
    .filter(row => row[0] && row[0] !== '#N/A');
  const teaching = rows.filter(row => Number(row[1]) > 0);
  if (!teaching.length) {
    return '<p>No lessons entered on this tab yet.</p>';
  }
  return '<p>Average minutes per week with each class, over the year:</p><ul>' + teaching.map(row =>
    '<li><b>' + escapeHtml(String(row[0])) + '</b>: ' + formatDuration(Math.round(Number(row[1]))) + '</li>'
  ).join('') + '</ul><p>Clashes with other specialists show in red on the tab, and in full when you run ' +
    '"1. Sync Specialist Schedules".</p>';
}


// ---------------------------------------------------------------------------
// Copy Week A to the other weeks / Copy Term 1 to Terms 2-4
// ---------------------------------------------------------------------------

function copyWeekAToOtherWeeks() {
  copyTimetableBlocks('weeks');
}

function copyTermOneToOtherTerms() {
  copyTimetableBlocks('terms');
}

/**
 * Copies entries on the active class or specialist tab, either from each
 * term's Week A to that term's other weeks, or from each Term 1 week to the
 * same week of Terms 2-4. Never touches break rows or specialist lessons
 * placed by the sync (including conflicts), never copies them, and never
 * clears anything; entries already in a target slot are kept unless the
 * user chooses to replace them.
 */
function copyTimetableBlocks(mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getActiveSheet();
  const title = mode === 'weeks' ? 'Copy Week A to the other weeks' : 'Copy Term 1 to Terms 2-4';

  if (!isSpecialistSheet(sheet) && !isClassSheet(ss, sheet)) {
    ui.alert(title, 'Open a class or specialist tab first, then run this again.', ui.ButtonSet.OK);
    return;
  }

  const grid = readTimetableGrid(sheet);
  const byKey = {};
  grid.columns.forEach(column => { byKey[column.term + '|' + column.week + '|' + column.day] = column; });
  const pairs = grid.columns
    .map(target => {
      const sourceKey = mode === 'weeks'
        ? target.term + '|0|' + target.day
        : '1|' + target.week + '|' + target.day;
      const source = byKey[sourceKey];
      return source && source !== target ? { source: source, target: target } : null;
    })
    .filter(Boolean);

  if (!pairs.length) {
    ui.alert(title, mode === 'weeks'
      ? 'This timetable has a 1-week cycle, so there are no other weeks to copy to.'
      : 'There are no other terms to copy to.', ui.ButtonSet.OK);
    return;
  }

  const answer = ui.alert(
    title,
    (mode === 'weeks'
      ? 'Copy each term\'s Week A entries into that term\'s other weeks?'
      : 'Copy each Term 1 week into the same week of Terms 2, 3 and 4?') +
      '\n\nBreaks and specialist lessons are never changed, and nothing is cleared.\n\n' +
      'Yes: fill only the empty slots.\nNo: also replace entries that are already there.\nCancel: stop.',
    ui.ButtonSet.YES_NO_CANCEL
  );
  if (answer !== ui.Button.YES && answer !== ui.Button.NO) return;
  const replaceExisting = answer === ui.Button.NO;

  const breakNames = getBreakNames(ss);
  const isSpecialistOwned = (value, note) =>
    String(note || '').indexOf(SYNC_NOTE_PREFIX) === 0 ||
    String(note || '').indexOf(CONFLICT_NOTE_PREFIX) === 0 ||
    String(value || '').indexOf(CONFLICT_VALUE_PREFIX) === 0;

  const changed = [];
  let keptExisting = 0;
  let keptSpecialist = 0;
  grid.values.forEach((row, r) => {
    if (!grid.times[r]) return;
    pairs.forEach(pair => {
      const source = String(row[pair.source.index] || '').trim();
      if (!source || breakNames[source] || isSpecialistOwned(source, grid.notes[r][pair.source.index])) return;
      const target = String(row[pair.target.index] || '').trim();
      if (target === source || breakNames[target]) return;
      if (isSpecialistOwned(target, grid.notes[r][pair.target.index])) { keptSpecialist++; return; }
      if (target && !replaceExisting) { keptExisting++; return; }
      row[pair.target.index] = source;
      changed.push({ r: r, index: pair.target.index });
    });
  });

  const failed = writeGridChanges(sheet, grid, changed);
  let message = 'Copied ' + (changed.length - failed) + ' slot(s).';
  if (keptExisting) message += '\n\nKept ' + keptExisting + ' slot(s) that already had an entry.';
  if (keptSpecialist) message += '\n\nLeft ' + keptSpecialist + ' specialist lesson(s) as they were.';
  if (failed) message += '\n\n' + failed + ' slot(s) are locked and could not be changed.';
  ui.alert(title, message, ui.ButtonSet.OK);
}


/**
 * Writes the changed cells of a grid read by readTimetableGrid. One write
 * for the whole grid; if that fails because some cells are locked, each
 * changed cell is written on its own. Returns the number that failed.
 */
function writeGridChanges(sheet, grid, changed) {
  if (!changed.length) return 0;
  try {
    sheet.getRange(5, 10, grid.values.length, grid.width).setValues(grid.values);
    return 0;
  } catch (bulkError) {
    let failed = 0;
    changed.forEach(cell => {
      try {
        sheet.getRange(cell.r + 5, cell.index + 10).setValue(grid.values[cell.r][cell.index]);
      } catch (cellError) {
        failed++;
      }
    });
    return failed;
  }
}


// ---------------------------------------------------------------------------
// Printing
// ---------------------------------------------------------------------------

/**
 * Lays out one week of the active class or specialist tab on the Print view
 * tab (consecutive slots with the same entry merged into one block, with its
 * times), then offers a link that downloads it as a landscape A4 PDF.
 */
function printThisTimetable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getActiveSheet();
  const isSpecialist = isSpecialistSheet(sheet);

  if (!isSpecialist && !isClassSheet(ss, sheet)) {
    ui.alert('Print this timetable', 'Open a class or specialist tab first, then run this again.', ui.ButtonSet.OK);
    return;
  }

  const grid = readTimetableGrid(sheet);
  const response = ui.prompt(
    'Print this timetable',
    'Which week? Enter the term and week, for example 1A for Term 1 Week A. Leave blank for 1A.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const match = (response.getResponseText().trim() || '1A').match(/^([1-4])\s*([A-Za-z])?$/);
  const term = match ? parseInt(match[1], 10) : NaN;
  const week = match ? weekLetterToIndex(match[2] || 'A') : NaN;
  const days = grid.columns.filter(column => column.term === term && column.week === week);
  if (days.length !== 5) {
    ui.alert('Print this timetable', 'That week isn\'t on this timetable. Enter something like 1A or 2B.', ui.ButtonSet.OK);
    return;
  }

  const blockLabel = days[0].blockLabel;
  const heading = isSpecialist
    ? String(sheet.getRange('B1').getDisplayValue() || sheet.getName())
    : sheet.getName() + ' - ' + String(sheet.getRange('C27').getDisplayValue());
  const printSheet = buildPrintSheet(ss, heading + ' - ' + blockLabel, grid, days, getBreakNames(ss));

  SpreadsheetApp.flush();
  showPdfLink(ui, ss, printSheet,
    'The Print view tab now shows ' + escapeHtml(heading) + ', ' + escapeHtml(blockLabel) + '.');
}


function printComplianceOverview() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const overview = ss.getSheetByName('OVERVIEW');
  if (!overview) {
    ui.alert('Print the compliance overview', 'There is no OVERVIEW tab in this timetable.', ui.ButtonSet.OK);
    return;
  }
  showPdfLink(ui, ss, overview, 'The OVERVIEW tab shows every class\'s compliance.');
}


/** Replaces the Print view tab with one week of a timetable. */
function buildPrintSheet(ss, title, grid, days, breakNames) {
  const existing = ss.getSheetByName(PRINT_SHEET_NAME);
  if (existing) ss.deleteSheet(existing);
  const sheet = ss.insertSheet(PRINT_SHEET_NAME);

  const slotRows = grid.times.map((time, r) => ({ time: String(time || ''), r: r })).filter(slot => slot.time);
  const firstRow = 4;
  const lastRow = firstRow + slotRows.length - 1;
  const startOf = slot => slot.time.split(' - ')[0];
  const endOf = slot => slot.time.split(' - ')[1];

  // One row per 5-minute slot; the time shows on the hour and half hour.
  const values = slotRows.map(slot => {
    const minutes = parseInt(startOf(slot).split(':')[1], 10);
    return [minutes % 30 === 0 ? startOf(slot) : ''].concat(
      days.map(day => String(grid.values[slot.r][day.index] || '').trim())
    );
  });

  sheet.getRange('A1').setValue(title).setFontSize(14).setFontWeight('bold');
  sheet.getRange(3, 1, 1, 6).setValues([['Time'].concat(days.map(day => day.day))])
    .setFontWeight('bold').setHorizontalAlignment('center').setBackground('#E8EEF2');
  sheet.getRange(firstRow, 1, values.length, 6).setNumberFormat('@').setValues(values)
    .setFontSize(8).setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true);
  sheet.getRange(firstRow, 1, values.length, 1).setFontColor('#666666').setVerticalAlignment('top');
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidths(2, 5, 150);
  sheet.setRowHeights(firstRow, values.length, 9);

  // Merge each run of slots with the same entry into one block showing its times.
  const breakRanges = [];
  for (let d = 0; d < days.length; d++) {
    let start = 0;
    while (start < values.length) {
      const value = values[start][d + 1];
      let end = start;
      while (end + 1 < values.length && value && values[end + 1][d + 1] === value) end++;
      if (value) {
        const range = sheet.getRange(firstRow + start, d + 2, end - start + 1, 1);
        if (end > start) range.merge();
        range.setValue(value + (end - start >= 2
          ? '\n' + startOf(slotRows[start]) + ' - ' + endOf(slotRows[end])
          : ''));
        if (breakNames[value]) breakRanges.push(range.getA1Notation());
      }
      start = end + 1;
    }
  }
  if (breakRanges.length) sheet.getRangeList(breakRanges).setBackground('#EEEEEE').setFontColor('#555555');
  sheet.getRange(3, 1, lastRow - 2, 6).setBorder(true, true, true, true, true, false);

  return sheet;
}


/** A dialog with a link that downloads one tab as a landscape A4 PDF. */
function showPdfLink(ui, ss, sheet, intro) {
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf' +
    '&gid=' + sheet.getSheetId() + '&size=A4&portrait=false&fitw=true&gridlines=false' +
    '&printtitle=false&sheetnames=false&pagenum=UNDEFINED&fzr=false' +
    '&top_margin=0.3&bottom_margin=0.3&left_margin=0.3&right_margin=0.3';
  showDialog(ui, 'Print',
    '<p>' + intro + '</p>' +
    '<p><a class="button" href="' + url + '" target="_blank">Download PDF</a></p>' +
    '<p class="note">The PDF opens in a new browser tab. You can also use File &gt; Download &gt; ' +
    'PDF with this tab selected.</p>');
}


// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * The timetable grid of a class or specialist tab, read in a few calls:
 * values and notes from J5, the times in column I, and every day column
 * with its term, week and block label ("Term 1 - Week A").
 */
function readTimetableGrid(sheet) {
  const lastRow = Math.max(5, sheet.getLastRow());
  const lastColumn = Math.max(10, sheet.getLastColumn());
  const numRows = lastRow - 4;
  const width = lastColumn - 9;
  const region = sheet.getRange(5, 10, numRows, width);
  const columnMap = buildScheduleColumnMap(sheet, lastColumn);

  const blockLabels = [];
  const columns = Object.keys(columnMap).map(key => parseInt(key, 10)).sort((a, b) => a - b).map(column => {
    const info = columnMap[column];
    const blockLabel = 'Term ' + info.term + ' - Week ' + String.fromCharCode(65 + info.week);
    if (blockLabels.indexOf(blockLabel) === -1) blockLabels.push(blockLabel);
    return { column: column, index: column - 10, term: info.term, week: info.week, day: info.day, blockLabel: blockLabel };
  });

  return {
    values: region.getValues(),
    notes: region.getNotes(),
    times: sheet.getRange(5, 9, numRows, 1).getValues().map(row => row[0]),
    width: width,
    columns: columns,
    blockLabels: blockLabels
  };
}


function getClassNames(ss) {
  const range = ss.getRangeByName('ClassList');
  return range ? range.getValues().flat().map(value => String(value).trim()).filter(Boolean) : [];
}

function findClassSheet(ss, className) {
  return ss.getSheetByName(className) || ss.getSheetByName(sanitizeSheetName(className));
}

function isClassSheet(ss, sheet) {
  const name = sheet.getName().toLowerCase();
  return getClassNames(ss).some(className =>
    className.toLowerCase() === name || sanitizeSheetName(className).toLowerCase() === name);
}

/** Break names (DATA's BreakList) as a lookup object. */
function getBreakNames(ss) {
  const names = {};
  const range = ss.getRangeByName('BreakList');
  if (range) range.getValues().flat().forEach(value => { if (value) names[String(value).trim()] = true; });
  return names;
}

function formatDuration(totalMinutes) {
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  if (hours && minutes) return hours + ' hr ' + minutes + ' min';
  if (hours) return hours + ' hr';
  return minutes + ' min';
}

function escapeHtml(text) {
  return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function showDialog(ui, title, bodyHtml) {
  const html = '<style>' +
    'body{font-family:Arial,sans-serif;font-size:13px;color:#202124;line-height:1.45}' +
    'h3{font-size:14px;margin:14px 0 4px}ul{margin:4px 0;padding-left:20px}' +
    '.ok{color:#137333}.bad{color:#b3261e}.note{color:#5f6368;font-size:12px}' +
    'table{border-collapse:collapse;width:100%}th,td{border:1px solid #dadce0;padding:4px 6px;text-align:left}' +
    'th{background:#f1f3f4}a.button{display:inline-block;background:#007A48;color:#fff;padding:8px 16px;' +
    'border-radius:4px;text-decoration:none;font-weight:bold}' +
    '</style>' + bodyHtml;
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(560).setHeight(520), title);
}
