'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createWorld, makeTemplate } = require('./helpers/fakeSpreadsheet');
const { loadScript } = require('./helpers/loadScript');
const { makePayload } = require('./helpers/payload');

const TEMPLATE_ID = '1x3pzraH_eG2nsQyiZPE9WGAzZ1wZIFGM2JUVF4W6FCk';

/** A generated timetable, opened later with the template's script (both files). */
function setUp() {
  const world = createWorld();
  makeTemplate(world, TEMPLATE_ID);
  loadScript('wizard/Code.gs', world.globals).generateTimetable(makePayload());
  world.newExecution();
  const ss = world.files[world.lastCopyId].spreadsheet;
  world.active = ss;
  const tools = loadScript(['template/TemplateSync.gs', 'template/TimetableTools.gs'], world.globals);
  const open = name => { ss.setActiveSheet(ss.getSheetByName(name)); return ss.getSheetByName(name); };
  const lastDialog = () => world.dialogs[world.dialogs.length - 1];
  return { world, ss, tools, open, lastDialog };
}

/** Stands in for Google calculating a class tab's compliance table. */
function setCompliance(sheet, rows, extra) {
  // [KLA, mins, %, min, max] for rows 5-12
  rows.forEach((row, i) => {
    sheet.getRange(5 + i, 2, 1, 6).setValues([[row[0], row[1], row[1] / 60, row[2], row[3], row[4]]]);
  });
  sheet.getRange('C13').setValue(rows.reduce((sum, row) => sum + row[1], 0));
  sheet.getRange('C15').setValue(extra.empty);
  sheet.getRange('C16').setValue(extra.compliant);
  sheet.getRange('C18').setValue(extra.amadeus);
  sheet.getRange('B21:G21').setValues([['Amadeus - Tutorials', extra.tutorials, 0, extra.tutorialsPercent, 0.1, 2]]);
}

const BALANCED = [
  ['English', 600, 33, 30, 35], ['Mathematics', 450, 25, 20, 35], ['Religion', 216, 12, 12, 15],
  ['Science & Tech', 162, 9, 8, 10], ['HSIE', 162, 9, 8, 10], ['PDHPE', 162, 9, 8, 10],
  ['Creative Arts', 54, 3, 8, 10], ['Non-Curricular', 0, 0, 0, 20]
];

test('the menu lists every action, and each one exists', () => {
  const { world, tools } = setUp();
  tools.onOpen();
  const labels = world.menu.filter(item => item[0] !== '-');
  assert.deepEqual(labels.map(item => item[0]).slice(0, 3), ['Go to Start here', '1. Sync Specialist Schedules', '2. Lock Synced Cells']);
  labels.forEach(([label, fn]) => assert.equal(typeof tools[fn], 'function', label));
});

test('Go to Start here opens the Start here tab', () => {
  const { ss, tools, open } = setUp();
  open('K Blue');
  tools.goToStartHere();
  assert.equal(ss.getActiveSheet().getName(), 'Start here');
});

test('sync and lock record their result on Start here; locking before syncing warns first', () => {
  const { world, ss, tools } = setUp();
  world.answers = ['NO'];
  tools.lockSyncedSpecialistCells();
  assert.match(world.alerts.pop(), /haven't been synced yet/);
  assert.equal(ss.getRangeByName('StartHereLockStatus').getValue(), 'Not run yet');

  ss.getSheetByName('SPEC_Mr Smith').getRange('K11').setValue('K Blue');
  tools.syncSpecialistSchedules();
  assert.match(ss.getRangeByName('StartHereSyncStatus').getValue(), /by coordinator@school\.example: 1 specialist slot\(s\) synced, 0 conflict\(s\)$/);
  tools.lockSyncedSpecialistCells();
  assert.match(ss.getRangeByName('StartHereLockStatus').getValue(), /: 1 cell\(s\) locked$/);
});

test('Check this timetable: KLAs out of range with minutes to add, Amadeus, empty slots, conflicts', () => {
  const { tools, open, lastDialog } = setUp();
  const sheet = open('Year 3 Gold');
  setCompliance(sheet, BALANCED, { empty: 3, compliant: false, amadeus: false, tutorials: 0, tutorialsPercent: 0 });
  // Fill the whole grid except three Term 4 Week B slots, plus one conflict.
  const grid = sheet.getRange('J5:BD85');
  grid.setValues(grid.getValues().map(row => row.map(value => value || 'English')));
  sheet.getRange('O5:O85').setValue(''); // keep spacer columns empty
  sheet.getRange('U5:U85').setValue('');
  ['AA', 'AG', 'AM', 'AS', 'AY'].forEach(col => sheet.getRange(col + '5:' + col + '85').setValue(''));
  sheet.getRange('BA20:BC20').setValue('');
  sheet.getRange('K11').setValue('CONFLICT: English vs Mr Smith');

  tools.checkThisTimetable();
  const html = lastDialog().html;
  assert.match(html, /Not compliant yet/);
  assert.match(html, /<b>Creative Arts<\/b>: 3\.0% of teaching time, needs 8-10%\. About <b>1 hr 35 min more<\/b> per week\./);
  assert.match(html, /Creative Arts is under its range/);
  assert.match(html, /No <b>Amadeus - Tutorials<\/b> lessons yet/);
  assert.match(html, /Empty slots \(3\)<\/h3><ul><li>Term 4 - Week B: 3 slot\(s\), 15 min<\/li>/);
  assert.match(html, /Conflicts \(1\)[\s\S]*K11: CONFLICT: English vs Mr Smith/);
});

test('Check this timetable: compliant class, and a K-2 class where Amadeus does not apply', () => {
  const { tools, open, lastDialog } = setUp();
  const sheet = open('K Blue');
  const rows = BALANCED.map(row => row[0] === 'Creative Arts' ? ['Creative Arts', 162, 9, 8, 10] : row);
  setCompliance(sheet, rows, { empty: 0, compliant: true, amadeus: 'N/A', tutorials: 0, tutorialsPercent: 0 });
  const grid = sheet.getRange('J5:AF85');
  grid.setValues(grid.getValues().map((row, r) => row.map((value, c) => ([5, 11, 17].includes(c) ? '' : value || 'English'))));
  tools.checkThisTimetable();
  const html = lastDialog().html;
  assert.match(html, /<b>Compliant\.<\/b>/);
  assert.match(html, /Every KLA is within its range/);
  assert.match(html, /Doesn't apply to this grade/);
  assert.match(html, /Empty slots<\/h3><p class="ok">None/);
});

test('Check this timetable on a specialist tab lists minutes per class; elsewhere it asks for a timetable tab', () => {
  const { world, tools, open, lastDialog } = setUp();
  const spec = open('SPEC_Mr Smith');
  spec.getRange('B6:C8').setValues([['K Blue', 45], ['K Red', 0], ['Year 3 Gold', 90]]);
  tools.checkThisTimetable();
  assert.match(lastDialog().html, /<b>K Blue<\/b>: 45 min<\/li><li><b>Year 3 Gold<\/b>: 1 hr 30 min/);
  open('DATA');
  tools.checkThisTimetable();
  assert.match(world.alerts.pop(), /Open a class tab/);
});

test('Check all class timetables lists every class, non-compliant first', () => {
  const { tools, ss, lastDialog } = setUp();
  setCompliance(ss.getSheetByName('K Red'), BALANCED, { empty: 0, compliant: true, amadeus: 'N/A', tutorials: 0, tutorialsPercent: 0 });
  tools.checkAllClassTimetables();
  const html = lastDialog().html;
  assert.match(html, /<b>1 of 3<\/b> class timetable\(s\) are compliant/);
  assert.ok(html.indexOf('K Blue') < html.indexOf('K Red'), 'non-compliant first');
});

test('Copy Week A: fills empty slots in each term, keeps breaks, existing entries and specialist lessons', () => {
  const { world, tools, open } = setUp();
  const sheet = open('Year 3 Gold'); // 2-week cycle: Term 1 Week A = J-N, Week B = P-T; Term 2 Week A = V-Z, Week B = AB-AF
  sheet.getRange('J11:J13').setValues([['English'], ['English'], ['Mathematics']]);
  sheet.getRange('V11').setValue('HSIE');
  sheet.getRange('P12').setValue('Religion');                               // existing entry
  sheet.getRange('P13').setValue('PDHPE').setNote('[SPEC_SYNC] Mr Smith');  // specialist lesson
  sheet.getRange('J29').setValue('Recess');                                 // break row (already Recess)

  world.answers = ['YES']; // fill empty slots only
  tools.copyWeekAToOtherWeeks();
  assert.equal(sheet.getRange('P11').getValue(), 'English');
  assert.equal(sheet.getRange('P12').getValue(), 'Religion', 'existing entry kept');
  assert.equal(sheet.getRange('P13').getValue(), 'PDHPE', 'specialist lesson kept');
  assert.equal(sheet.getRange('AB11').getValue(), 'HSIE', 'Term 2 copies from Term 2 Week A');
  assert.equal(sheet.getRange('P29').getValue(), 'Recess');
  assert.match(world.alerts.pop(), /Copied 2 slot\(s\)\.\n\nKept 1 slot\(s\) that already had an entry\.\n\nLeft 1 specialist lesson/);

  world.answers = ['NO']; // also replace existing entries
  tools.copyWeekAToOtherWeeks();
  assert.equal(sheet.getRange('P12').getValue(), 'English');
  assert.equal(sheet.getRange('P13').getValue(), 'PDHPE');

  world.answers = ['CANCEL'];
  sheet.getRange('J14').setValue('Religion');
  tools.copyWeekAToOtherWeeks();
  assert.equal(sheet.getRange('P14').getValue(), '');
});

test('Copy Term 1 to Terms 2-4, and the 1-week cycle message for Copy Week A', () => {
  const { world, tools, open } = setUp();
  const sheet = open('K Blue'); // 1-week cycle: Term 1 = J-N, Term 2 = P-T, Term 3 = V-Z, Term 4 = AB-AF
  sheet.getRange('J11').setValue('English');
  world.answers = ['YES'];
  tools.copyTermOneToOtherTerms();
  assert.deepEqual(['P11', 'V11', 'AB11'].map(a1 => sheet.getRange(a1).getValue()), ['English', 'English', 'English']);
  tools.copyWeekAToOtherWeeks();
  assert.match(world.alerts.pop(), /1-week cycle/);
});

test('Print this timetable: one week on the Print view tab, merged blocks with times, PDF link', () => {
  const { world, ss, tools, open, lastDialog } = setUp();
  const sheet = open('Year 3 Gold');
  sheet.getRange('P11:P22').setValue('English'); // Term 1 Week B Monday 9:00-10:00
  world.promptAnswers = ['1b'];
  tools.printThisTimetable();

  const print = ss.getSheetByName('Print view');
  assert.equal(print.getRange('A1').getValue(), 'Year 3 Gold - Ms Chen - Term 1 - Week B');
  assert.deepEqual(print.getRange('A3:F3').getValues()[0], ['Time', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']);
  assert.equal(print.getRange('A4').getValue(), '08:30');
  assert.equal(print.getRange('B10').getValue(), 'English\n09:00 - 10:00');
  assert.ok(print._merges.some(m => m.row === 10 && m.column === 2 && m.numRows === 12));
  assert.equal(print.getRange('B28').getValue(), 'Recess\n10:30 - 11:00');
  assert.equal(print.getRange('B28').getBackgrounds()[0][0], '#EEEEEE');
  assert.match(lastDialog().html, new RegExp('export\\?format=pdf&gid=' + print.getSheetId() + '&size=A4&portrait=false'));

  // Printing again replaces the tab; an unknown week is refused.
  world.promptAnswers = ['4C'];
  tools.printThisTimetable();
  assert.match(world.alerts.pop(), /That week isn't on this timetable/);
  world.promptAnswers = [''];
  tools.printThisTimetable();
  assert.equal(ss.getSheets().filter(s => s.getName() === 'Print view').length, 1);
  assert.match(ss.getSheetByName('Print view').getRange('A1').getValue(), /Term 1 - Week A$/);
});

test('Print the compliance overview links the OVERVIEW tab as a PDF', () => {
  const { ss, tools, lastDialog } = setUp();
  tools.printComplianceOverview();
  assert.match(lastDialog().html, new RegExp('gid=' + ss.getSheetByName('OVERVIEW').getSheetId() + '&'));
});
