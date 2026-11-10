/**
 * Timetable tools for the Timetable Admin menu (see onOpen in
 * TemplateSync.gs): checking a timetable, copying weeks and terms, and
 * making Google Docs of a timetable or the compliance summary. Add this file to the Apps Script project bound to the master
 * template, next to TemplateSync.gs; it uses that file's helpers
 * (buildScheduleColumnMap, isSpecialistSheet, the note prefixes).
 *
 * Grid layout (from the wizard): row 1 "Term X - Week Y" headers, row 4 day
 * names, column I times, timetable from J5. Class tabs hold the compliance
 * table in B5:G21 (KLA rows 5-12, total in row 13, unallocated blocks C15,
 * compliant C16, Amadeus C18, Amadeus Tutorials row 21) and the class
 * details in C27:C31.
 */

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
// Copy Week A into the other weeks / Copy Term 1 into Terms 2-4
// ---------------------------------------------------------------------------

const COPY_MODES = {
  weeks: {
    title: 'Copy Week A into the other weeks',
    explanation: 'Copies what is in Week A of each term into the other weeks of the same term ' +
      '(Term 1 Week A into Term 1 Week B, and so on).'
  },
  terms: {
    title: 'Copy Term 1 into Terms 2-4',
    explanation: 'Copies each week of Term 1 into the same week of Terms 2, 3 and 4 ' +
      '(Term 1 Week A into Term 2 Week A, Term 3 Week A and Term 4 Week A, and so on).'
  }
};

function copyWeekAToOtherWeeks() {
  runCopyTool('weeks');
}

function copyTermOneToOtherTerms() {
  runCopyTool('terms');
}

/**
 * Explains what will be copied on the active tab, asks whether to keep
 * what's already there, copies and shows the result. Only Sheets' own
 * alerts are used, not an HTML window calling back with google.script.run:
 * that fails with PERMISSION_DENIED for anyone signed into more than one
 * Google account in the browser.
 */
function runCopyTool(mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getActiveSheet();
  const copyMode = COPY_MODES[mode];

  if (!isSpecialistSheet(sheet) && !isClassSheet(ss, sheet)) {
    ui.alert(copyMode.title, 'Open a class or specialist tab first, then choose this again.', ui.ButtonSet.OK);
    return;
  }
  if (!getCopyPairs(readTimetableGrid(sheet), mode).length) {
    ui.alert(copyMode.title, mode === 'weeks'
      ? 'This tab has a 1-week timetable, so there are no other weeks to copy into. ' +
        'Use "Copy Term 1 into Terms 2-4" instead.'
      : 'There are no other terms to copy into.', ui.ButtonSet.OK);
    return;
  }

  const answer = ui.alert(copyMode.title,
    copyMode.explanation + '\n\n' +
    'This only changes the "' + sheet.getName() + '" tab. Breaks and specialist lessons are never ' +
    'changed, and nothing is deleted.\n\n' +
    'Keep what is already in the slots you are copying into?\n\n' +
    'Yes: fill empty slots only, and keep what is already there.\n' +
    'No: fill empty slots, and replace what is already there.\n' +
    'Cancel: don\'t copy anything.',
    ui.ButtonSet.YES_NO_CANCEL);
  if (answer !== ui.Button.YES && answer !== ui.Button.NO) return;

  ss.toast('Copying...', copyMode.title, 30);
  ui.alert(copyMode.title, describeCopyResult(copyTimetableBlocks(ss, sheet, mode, answer === ui.Button.NO)),
    ui.ButtonSet.OK);
}

/** The result of copyTimetableBlocks as a message for an alert. */
function describeCopyResult(result) {
  const lines = [];
  lines.push(result.copied
    ? 'Done. Filled ' + result.copied + ' slot(s) (' + formatDuration(result.copied * 5) + ').'
    : 'Nothing to copy. Every slot already matched or was kept.');
  if (result.keptExisting) {
    lines.push(result.keptExisting + ' slot(s) already had something in them and were left as they were.');
  }
  if (result.keptSpecialist) {
    lines.push(result.keptSpecialist + ' specialist lesson slot(s) were left as they were.');
  }
  if (result.locked) {
    lines.push(result.locked + ' slot(s) are locked and could not be changed.');
  }
  return lines.join('\n\n');
}

/** Each target day column with the column it copies from. */
function getCopyPairs(grid, mode) {
  const byKey = {};
  grid.columns.forEach(column => { byKey[column.term + '|' + column.week + '|' + column.day] = column; });
  return grid.columns
    .map(target => {
      const source = byKey[mode === 'weeks'
        ? target.term + '|0|' + target.day
        : '1|' + target.week + '|' + target.day];
      return source && source !== target ? { source: source, target: target } : null;
    })
    .filter(Boolean);
}

/**
 * Copies entries on a class or specialist tab, either from each term's
 * Week A into that term's other weeks, or from each Term 1 week into the
 * same week of Terms 2-4. Never touches break rows or specialist lessons
 * placed by the sync (including conflicts), never copies them, and never
 * clears anything; entries already in a target slot are kept unless
 * replaceExisting is true.
 */
function copyTimetableBlocks(ss, sheet, mode, replaceExisting) {
  const grid = readTimetableGrid(sheet);
  const pairs = getCopyPairs(grid, mode);
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

  const locked = writeGridChanges(sheet, grid, changed);
  return { copied: changed.length - locked, keptExisting: keptExisting, keptSpecialist: keptSpecialist, locked: locked };
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
// Google Docs: a timetable (one week, or all terms), and the compliance summary
// ---------------------------------------------------------------------------

/**
 * Asks which week (or term, for a 1-week timetable), or all terms, to make
 * a Google Doc of for the active class or specialist tab, makes it and
 * shows the link. The question is a Sheets prompt, not an HTML window
 * calling back with google.script.run: that fails with PERMISSION_DENIED
 * for anyone signed into more than one Google account in the browser. The
 * window with the link only shows it, which works for everyone.
 */
function makeTimetableDoc() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getActiveSheet();
  const title = 'Make a Google Doc of this timetable';

  if (!isSpecialistSheet(sheet) && !isClassSheet(ss, sheet)) {
    ui.alert(title, 'Open a class or specialist tab first, then choose this again.', ui.ButtonSet.OK);
    return;
  }
  if (!hasDriveService(ui, title)) return;

  const grid = readTimetableGrid(sheet);
  const choice = askDocWeekChoice(ui, title, sheet, grid);
  if (!choice) return;

  ss.toast('Making the Google Doc...', title, 60);
  const doc = makeTimetableDocOfTab(ss, sheet, choice);
  showDialog(ui, title,
    '<p class="ok"><b>Your Google Doc is ready:</b> ' + escapeHtml(doc.title) + '.</p>' +
    (doc.pages > 1 ? '<p>' + doc.pages + ' pages, one per ' + doc.pageUnit + '.</p>' : '') +
    '<p>It is saved in the same folder as this timetable and listed on the Start here tab.</p>' +
    '<p><a class="button" href="' + doc.url + '" target="_blank">Open the document</a></p>');
}


/**
 * Asks for the part of the year by number. Leaving it blank picks the week
 * (or term) the selected cell is in, when it is in one. Asks again after
 * an answer that isn't one of the numbers. Returns the choice's value
 * ("term|week" or "all"), or null if cancelled.
 */
function askDocWeekChoice(ui, title, sheet, grid) {
  const choices = getDocWeekChoices(grid);
  const selected = getSelectedDocWeekChoice(sheet, grid, choices);
  const question = 'Makes a Google Doc of the "' + sheet.getName() + '" timetable, coloured like the tab.\n\n' +
    'Which part of the year? Type its number:\n\n' +
    choices.map((choice, i) => (i + 1) + '   ' + choice.label).join('\n') +
    (selected ? '\n\nOr leave it blank for ' + selected.label + ', where your selected cell is.' : '');

  let problem = '';
  for (;;) {
    const response = ui.prompt(title, problem + question, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) return null;
    const text = String(response.getResponseText()).trim();
    if (!text && selected) return selected.value;
    const number = /^\d+$/.test(text) ? Number(text) : 0;
    if (number >= 1 && number <= choices.length) return choices[number - 1].value;
    problem = (text ? '"' + text + '" isn\'t one of the numbers. ' : '') +
      'Type a number from 1 to ' + choices.length + '.\n\n';
  }
}


/** The week (or term) choice the selected cell's column is in, or null. */
function getSelectedDocWeekChoice(sheet, grid, choices) {
  const cell = sheet.getActiveCell();
  if (!cell) return null;
  const column = cell.getColumn();
  // Each week takes 6 columns: 5 days and a gap, which counts as the week before it.
  const day = grid.columns.filter(item => item.column <= column && column - item.column < 6).pop();
  if (!day) return null;
  return choices.filter(choice => choice.value === day.term + '|' + day.week)[0] || null;
}


/** The weeks of a timetable (or terms, for a 1-week cycle), then "All terms". */
function getDocWeekChoices(grid) {
  const oneWeekCycle = !grid.columns.some(column => column.week > 0);
  const choices = [];
  grid.columns.forEach(column => {
    const value = column.term + '|' + column.week;
    if (choices.some(choice => choice.value === value)) return;
    choices.push({ value: value, label: oneWeekCycle ? 'Term ' + column.term : column.blockLabel });
  });
  choices.push({ value: 'all', label: 'All terms (one page per ' + (oneWeekCycle ? 'term' : 'week') + ')' });
  return choices;
}


/**
 * Makes the Doc of a class or specialist tab for the chosen week
 * ("term|week") or "all", and lists it on Start here. Returns its title,
 * link and how many pages (one per week, or per term) it has.
 */
function makeTimetableDocOfTab(ss, sheet, choice) {
  const grid = readTimetableGrid(sheet);
  const oneWeekCycle = !grid.columns.some(column => column.week > 0);
  const blocks = getDocWeekChoices(grid)
    .filter(item => item.value !== 'all' && (choice === 'all' || item.value === choice))
    .map(item => {
      const parts = item.value.split('|').map(Number);
      return {
        label: item.label,
        days: grid.columns.filter(column => column.term === parts[0] && column.week === parts[1])
      };
    });
  if (!blocks.length) throw new Error('That part of the year isn\'t on this timetable.');

  const details = getTimetableDetails(ss, sheet, grid);
  const colours = getEntryColours(sheet);
  const breakNames = getBreakNames(ss);
  const madeFrom = choice === 'all' ? 'All terms' : blocks[0].label;
  const docTitle = details.name + ' - ' + madeFrom;

  const sections = blocks.map((block, i) =>
    (blocks.length > 1 || i > 0 ? '<h2>' + escapeHtml(block.label) + '</h2>' : '') +
    buildWeekTableHtml(grid, block.days, colours, breakNames));
  const html = docHtmlPage(
    '<h1>' + escapeHtml(details.name) + (blocks.length === 1 ? ' - ' + escapeHtml(blocks[0].label) : '') + '</h1>' +
    buildDetailsHtml(details) + sections.join('')
  );

  // One page per week (or term): a page break before every heading but the first.
  const doc = createDocFromHtml(ss, docTitle, html, {
    pageBreaks: blocks.length > 1,
    tableWidths: [DETAILS_TABLE_WIDTHS].concat(blocks.map(() => WEEK_TABLE_WIDTHS))
  });
  recordDocOnStartHere(ss, doc.url, docTitle, sheet.getName() + ', ' + madeFrom);

  return { title: docTitle, url: doc.url, pages: blocks.length, pageUnit: oneWeekCycle ? 'term' : 'week' };
}


/**
 * The details printed at the top of a timetable Doc: who and what it is
 * for, the cycle, the school day, and when and by whom it was made.
 */
function getTimetableDetails(ss, sheet, grid) {
  const schoolName = getSchoolName(ss);
  const times = grid.times.map(String).filter(Boolean);
  const schoolDay = times.length
    ? times[0].split(' - ')[0] + ' - ' + times[times.length - 1].split(' - ')[1]
    : '';
  const breaks = describeBreaks(grid, getBreakNames(ss));
  const generated = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'd MMMM yyyy, h:mm a');
  const who = Session.getActiveUser().getEmail();
  const rows = [];

  if (isSpecialistSheet(sheet)) {
    const name = String(sheet.getRange('B1').getDisplayValue() || sheet.getName().replace(/^SPEC_/, ''));
    const cycle = Number(sheet.getRange('D3').getValue()) || 1;
    rows.push(['Specialist', name], ['KLA', String(sheet.getRange('B3').getDisplayValue())]);
    rows.push(['Timetable', cycle + '-week cycle']);
    if (schoolName) rows.push(['School', schoolName]);
    if (schoolDay) rows.push(['School day', schoolDay]);
    if (breaks) rows.push(['Breaks', breaks]);
    rows.push(['Generated', generated + (who ? ' by ' + who : '')]);
    return { name: name, rows: rows };
  }

  const values = sheet.getRange('C27:C30').getDisplayValues().map(row => String(row[0]));
  rows.push(['Class', values[1] || sheet.getName()], ['Teacher', values[0]], ['Grade', values[2]]);
  rows.push(['Timetable', parseCycleFromClassSheet(sheet) + '-week cycle']);
  if (schoolName) rows.push(['School', schoolName]);
  if (schoolDay) rows.push(['School day', schoolDay]);
  if (breaks) rows.push(['Breaks', breaks]);
  rows.push(['Generated', generated + (who ? ' by ' + who : '')]);
  return { name: sheet.getName(), rows: rows.filter(row => row[1]) };
}


/** "Recess 10:30 - 11:00, Lunch 13:00 - 13:45", from the first day's break rows. */
function describeBreaks(grid, breakNames) {
  const firstDay = grid.columns[0];
  if (!firstDay) return '';
  const breaks = [];
  grid.values.forEach((row, r) => {
    const value = String(row[firstDay.index] || '').trim();
    const time = String(grid.times[r] || '');
    if (!breakNames[value] || !time) return;
    const last = breaks[breaks.length - 1];
    if (last && last.name === value && last.lastRow === r - 1) {
      last.end = time.split(' - ')[1];
      last.lastRow = r;
    } else {
      breaks.push({ name: value, start: time.split(' - ')[0], end: time.split(' - ')[1], lastRow: r });
    }
  });
  return breaks.map(item => item.name + ' ' + item.start + ' - ' + item.end).join(', ');
}


/** The details as a two-column table, two details per row. */
function buildDetailsHtml(details) {
  let html = '<table class="details">';
  for (let i = 0; i < details.rows.length; i += 2) {
    html += '<tr>' + [details.rows[i], details.rows[i + 1]].map(row => row
      ? '<td class="label">' + escapeHtml(row[0]) + '</td><td>' + escapeHtml(row[1]) + '</td>'
      : '<td class="label"></td><td></td>').join('') + '</tr>';
  }
  return html + '</table>';
}


/**
 * Makes a Google Doc listing every class with its KLA percentages
 * (out-of-range values in red), empty slots and whether it is compliant.
 */
function makeComplianceDoc() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const title = 'Make a Google Doc of the compliance summary';
  if (!hasDriveService(ui, title)) return;

  const classSheets = getClassNames(ss).map(name => findClassSheet(ss, name)).filter(Boolean);
  if (!classSheets.length) {
    ui.alert(title, 'No class tabs were found.', ui.ButtonSet.OK);
    return;
  }
  ss.toast('Checking ' + classSheets.length + ' class timetable(s)...', title, 10);

  const klaNames = classSheets[0].getRange('B5:B12').getValues().map(row => String(row[0]));
  const header = '<tr><th>Class</th><th>Teacher</th>' +
    klaNames.map(kla => '<th>' + escapeHtml(kla) + '</th>').join('') +
    '<th>Empty slots</th><th>Compliant</th></tr>';
  let compliantCount = 0;
  const rows = classSheets.map(sheet => {
    const result = checkClassTimetable(ss, sheet);
    if (result.compliant) compliantCount++;
    const table = sheet.getRange('B5:G12').getValues();
    const outOfRange = {};
    result.klaIssues.forEach(issue => { outOfRange[issue.kla] = true; });
    return '<tr><td><b>' + escapeHtml(sheet.getName()) + '</b></td><td>' +
      escapeHtml(String(sheet.getRange('C27').getDisplayValue())) + '</td>' +
      table.map(row => {
        const percent = (Number(row[3]) || 0).toFixed(1) + '%';
        return outOfRange[String(row[0])]
          ? '<td style="background-color:#F4CCCC;color:#990000"><b>' + percent + '</b></td>'
          : '<td style="background-color:#D9EAD3;color:#274E13">' + percent + '</td>';
      }).join('') +
      '<td>' + result.emptySlots + '</td>' +
      (result.compliant
        ? '<td style="background-color:#D9EAD3;color:#274E13"><b>Yes</b></td>'
        : '<td style="background-color:#F4CCCC;color:#990000"><b>No</b></td>') + '</tr>';
  }).join('');

  const schoolName = getSchoolName(ss);
  const who = Session.getActiveUser().getEmail();
  const generated = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'd MMMM yyyy, h:mm a');
  const docTitle = (schoolName ? schoolName + ' - ' : '') + 'Compliance summary';
  const html = docHtmlPage(
    '<h1>' + escapeHtml(docTitle) + '</h1>' +
    buildDetailsHtml({ rows: [
      ['Classes compliant', compliantCount + ' of ' + classSheets.length],
      ['Generated', generated + (who ? ' by ' + who : '')]
    ] }) +
    '<p>Percentage of teaching time per KLA, averaged over the year. Green is within its range, red is outside it.</p>' +
    '<table>' + header + rows + '</table>'
  );
  const kla = 0.64 / klaNames.length;
  const doc = createDocFromHtml(ss, docTitle, html, {
    tableWidths: [DETAILS_TABLE_WIDTHS, [0.1, 0.12].concat(klaNames.map(() => kla), [0.07, 0.07])]
  });
  recordDocOnStartHere(ss, doc.url, docTitle, 'Compliance summary, all classes');
  showDialog(ui, title,
    '<p class="ok"><b>Your Google Doc is ready.</b> It is saved in the same folder as this timetable and ' +
    'listed on the Start here tab.</p>' +
    '<p><a class="button" href="' + doc.url + '" target="_blank">Open the document</a></p>');
}


/**
 * The HTML table for one week: one row per stretch of time between changes
 * on any day, and each lesson as one cell spanning its rows (rowspan), so
 * Google Docs shows it as a merged block.
 */
function buildWeekTableHtml(grid, days, colours, breakNames) {
  const slots = grid.times
    .map((time, r) => ({ r: r, start: String(time || '').split(' - ')[0], end: String(time || '').split(' - ')[1] }))
    .filter(slot => slot.start && slot.end);
  const valueAt = (slotIndex, day) => String(grid.values[slots[slotIndex].r][day.index] || '').trim();
  // The specialist teaching a synced lesson, from the note the sync leaves on the cell.
  const specialistAt = (slotIndex, day) => {
    const note = String((grid.notes[slots[slotIndex].r] || [])[day.index] || '');
    return note.indexOf(SYNC_NOTE_PREFIX) === 0 ? note.substring(SYNC_NOTE_PREFIX.length).trim() : '';
  };

  // Each day's runs of the same entry (empty runs too); back-to-back lessons
  // of the same KLA with different specialists stay separate.
  const runsByDay = days.map(day => {
    const runs = [];
    slots.forEach((slot, i) => {
      const value = valueAt(i, day);
      const specialist = value ? specialistAt(i, day) : '';
      const last = runs[runs.length - 1];
      if (last && last.value === value && last.specialist === specialist) last.end = i;
      else runs.push({ value: value, specialist: specialist, start: i, end: i });
    });
    return runs;
  });

  // Table rows: the stretches of time between any two changes.
  const boundaries = {};
  boundaries[slots.length] = true;
  runsByDay.forEach(runs => runs.forEach(run => { boundaries[run.start] = true; }));
  const rowStarts = Object.keys(boundaries).map(Number).sort((a, b) => a - b);
  const rowIndexOf = {};
  rowStarts.forEach((start, i) => { rowIndexOf[start] = i; });

  const rows = [];
  for (let i = 0; i < rowStarts.length - 1; i++) {
    const startSlot = rowStarts[i];
    const endSlot = rowStarts[i + 1] - 1;
    const minutes = (endSlot - startSlot + 1) * 5;
    let html = '<tr style="height:' + Math.max(14, Math.round(minutes * 0.95)) + 'pt">' +
      '<td class="time">' + slots[startSlot].start + '<br>' + slots[endSlot].end + '</td>';
    runsByDay.forEach(runs => {
      const run = runs.find(item => item.start === startSlot);
      if (!run) return; // covered by a cell above (rowspan)
      const span = rowIndexOf[run.end + 1] - rowIndexOf[run.start];
      const colour = breakNames[run.value]
        ? { background: '#EEEEEE', font: '#555555' }
        : (run.value ? colours(run.value) : null);
      const style = colour
        ? ' style="background-color:' + colour.background + (colour.font ? ';color:' + colour.font : '') + '"'
        : '';
      const text = run.value
        ? '<b>' + escapeHtml(run.value) + '</b>' +
          (run.specialist ? '<br><span class="small">' + escapeHtml(run.specialist) + '</span>' : '') +
          '<br><span class="small">' + slots[run.start].start + ' - ' + slots[run.end].end + '</span>'
        : '';
      html += '<td' + (span > 1 ? ' rowspan="' + span + '"' : '') + style + '>' + text + '</td>';
    });
    rows.push(html + '</tr>');
  }

  return '<table><tr><th class="time">Time</th>' + days.map(day => '<th>' + escapeHtml(day.day) + '</th>').join('') +
    '</tr>' + rows.join('') + '</table>';
}


function docHtmlPage(body) {
  return '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,sans-serif;font-size:9pt}' +
    'h1{font-size:16pt;margin:0 0 4pt 0}h2{font-size:13pt;margin:0 0 4pt 0}' +
    'table{border-collapse:collapse;width:100%}' +
    'th,td{border:1px solid #999999;padding:2pt 4pt;text-align:center;vertical-align:middle;font-size:9pt}' +
    'th{background-color:#E8EEF2}td.time,th.time{width:12%;color:#555555;font-size:8pt}' +
    'table.details{margin:0 0 8pt 0}table.details td{border:none;text-align:left;padding:1pt 8pt 1pt 0}' +
    'table.details td.label{color:#555555;font-weight:bold;width:12%}' +
    '.small{font-size:7pt}' +
    '</style></head><body>' + body + '</body></html>';
}


// Column widths as fractions of the page width (Google's conversion ignores
// width:100%, so the widths are set once the Doc exists).
const DOC_PAGE_WIDTH = 842; // A4 landscape, in points
const DOC_MARGIN = 28;
const DETAILS_TABLE_WIDTHS = [0.12, 0.38, 0.12, 0.38];
const WEEK_TABLE_WIDTHS = [0.1, 0.18, 0.18, 0.18, 0.18, 0.18];

/**
 * Converts HTML into a Google Doc in the timetable's folder, sets it to
 * landscape A4 and stretches each table to the page width
 * (options.tableWidths: one list of column fractions per table, in order;
 * equal columns otherwise). With options.pageBreaks, starts a new page
 * before every second-level heading after the first. Returns { id, url }.
 */
function createDocFromHtml(ss, name, html, options) {
  options = options || {};
  const parents = Drive.Files.get(ss.getId(), { fields: 'parents' }).parents || [];
  const file = Drive.Files.create(
    { name: name, mimeType: 'application/vnd.google-apps.document', parents: parents },
    Utilities.newBlob(html, 'text/html', name + '.html')
  );
  const doc = DocumentApp.openById(file.id);
  const body = doc.getBody();
  body.setPageWidth(DOC_PAGE_WIDTH).setPageHeight(595);
  body.setMarginTop(DOC_MARGIN).setMarginBottom(DOC_MARGIN).setMarginLeft(DOC_MARGIN).setMarginRight(DOC_MARGIN);

  const usableWidth = DOC_PAGE_WIDTH - 2 * DOC_MARGIN;
  const tableWidths = options.tableWidths || [];
  let headingsSeen = 0;
  let tablesSeen = 0;
  const headingIndexes = [];
  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    const type = child.getType();
    if (type === DocumentApp.ElementType.TABLE) {
      const table = child.asTable();
      const columnCount = table.getNumRows() ? table.getRow(0).getNumCells() : 0;
      const fractions = tableWidths[tablesSeen++];
      const widths = fractions && fractions.length === columnCount
        ? fractions
        : new Array(columnCount).fill(1 / columnCount);
      // Whole points; the last column takes what rounding leaves, so the table is exactly the page width.
      let used = 0;
      widths.forEach((fraction, column) => {
        const width = column === widths.length - 1 ? usableWidth - used : Math.round(usableWidth * fraction);
        used += width;
        table.setColumnWidth(column, width);
      });
    } else if (type === DocumentApp.ElementType.PARAGRAPH &&
        child.asParagraph().getHeading() === DocumentApp.ParagraphHeading.HEADING2) {
      headingsSeen++;
      if (headingsSeen > 1) headingIndexes.push(i);
    }
  }
  // From the last heading back, so each insert doesn't shift the ones still to do.
  if (options.pageBreaks) headingIndexes.reverse().forEach(index => body.insertPageBreak(index));
  doc.saveAndClose();
  return { id: file.id, url: 'https://docs.google.com/document/d/' + file.id + '/edit' };
}


/**
 * Adds a made Doc to the "Google Docs" list on the Start here tab (below the
 * named range StartHereDocs, the list's column headings). Timetables made
 * before the list existed are skipped.
 */
function recordDocOnStartHere(ss, url, name, madeFrom) {
  const headings = ss.getRangeByName('StartHereDocs');
  if (!headings) return;
  const sheet = headings.getSheet();
  const firstRow = headings.getRow() + 1;
  const existing = sheet.getRange(firstRow, 1, 200, 1).getDisplayValues();
  let row = firstRow;
  while (row - firstRow < existing.length) {
    const text = String(existing[row - firstRow][0]);
    if (!text || text.indexOf('None yet') === 0) break;
    row++;
  }
  const who = Session.getActiveUser().getEmail();
  sheet.getRange(row, 1, 1, 4).setValues([[
    '=HYPERLINK("' + url + '","' + name.replace(/"/g, '""') + '")',
    madeFrom,
    Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'd MMM yyyy, h:mm a'),
    who || ''
  ]]).setFontColor('#1F2933'); // Start here's text colour (the "None yet" line is grey)
}


function hasDriveService(ui, title) {
  if (typeof Drive !== 'undefined' && Drive.Files) return true;
  ui.alert(title,
    'Making Google Docs needs the Drive API service, which isn\'t turned on for this timetable. ' +
    'Ask the person who manages the timetable tool to turn it on in the template\'s script ' +
    '(Extensions > Apps Script > Services > Drive API).', ui.ButtonSet.OK);
  return false;
}


/**
 * How the tab colours each entry, worked out from its conditional-formatting
 * rules in order (the first matching rule wins, as in Sheets): "text is
 * exactly" rules (the KLA colours on class tabs), formulas like ="K Blue"
 * (the class colours on specialist tabs) and "not empty" formulas like
 * LEN(TRIM(J5))>0. Returns a function from an entry to {background, font},
 * or null when no rule colours it.
 */
function getEntryColours(sheet) {
  const matchers = [];
  sheet.getConditionalFormatRules().forEach(rule => {
    const condition = rule.getBooleanCondition();
    if (!condition) return;
    const background = colourToHex(condition, 'Background');
    if (!background) return;
    const colour = { background: background, font: colourToHex(condition, 'FontColor') };
    const type = String(condition.getCriteriaType());
    const values = condition.getCriteriaValues();

    if (type === 'TEXT_EQUAL_TO' || type === 'NUMBER_EQUAL_TO') {
      const text = String(values[0]).replace(/^=/, '').replace(/^"(.*)"$/, '$1').replace(/""/g, '"');
      matchers.push({ test: value => value === text, colour: colour });
    } else if (type === 'CUSTOM_FORMULA') {
      // Spaces are allowed around the parts of the formula, but not changed
      // inside the quoted text ("K Blue").
      const formula = String(values[0]).trim();
      const equals = formula.match(/^=\s*\$?[A-Z]+\$?\d+\s*=\s*"((?:[^"]|"")*)"\s*$/);
      if (equals) {
        const text = equals[1].replace(/""/g, '"');
        matchers.push({ test: value => value === text, colour: colour });
      } else if (/^=\s*LEN\(\s*TRIM\(\s*\$?[A-Z]+\$?\d+\s*\)\s*\)\s*>\s*0\s*$/i.test(formula)) {
        matchers.push({ test: value => value !== '', colour: colour });
      }
    }
  });
  return value => {
    const match = matchers.find(matcher => matcher.test(value));
    return match ? match.colour : null;
  };
}


/** A conditional-format colour as #RRGGBB, or '' (theme colours are skipped). */
function colourToHex(condition, kind) {
  try {
    const colour = condition['get' + kind + 'Object'] && condition['get' + kind + 'Object']();
    if (colour) return colour.asRgbColor().asHexString().toUpperCase();
  } catch (error) {
    // Theme colours have no fixed RGB value; fall back below.
  }
  const hex = condition['get' + kind] && condition['get' + kind]();
  return hex ? String(hex).toUpperCase() : '';
}


function getSchoolName(ss) {
  const range = ss.getRangeByName('StartHereSyncStatus');
  if (!range) return '';
  const title = String(range.getSheet().getRange('A1').getValue());
  return title.replace(/ - Timetable$/, '');
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
    'button{font-size:13px;padding:7px 14px;border:1px solid #dadce0;border-radius:4px;background:#fff;cursor:pointer}' +
    'button.primary{background:#007A48;color:#fff;border-color:#007A48;font-weight:bold}' +
    '</style>' + bodyHtml;
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(560).setHeight(520), title);
}
