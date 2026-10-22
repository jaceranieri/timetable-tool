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

/** Runs a copy tool on the active tab, answering its Yes/No/Cancel question; returns the alerts it showed. */
function runCopy(world, tools, mode, answer) {
  world.alerts = [];
  world.answers = [answer];
  if (mode === 'weeks') tools.copyWeekAToOtherWeeks(); else tools.copyTermOneToOtherTerms();
  return world.alerts;
}

test('Copy Week A: asks in a Sheets alert, not an HTML window (which fails with more than one Google account)', () => {
  const { world, tools, open } = setUp();
  open('Year 3 Gold').getRange('J11').setValue('English');
  const alerts = runCopy(world, tools, 'weeks', 'CANCEL');
  assert.equal(alerts.length, 1);
  assert.match(alerts[0], /^Copies what is in Week A of each term into the other weeks of the same term/);
  assert.match(alerts[0], /only changes the "Year 3 Gold" tab/);
  assert.match(alerts[0], /Yes: fill empty slots only, and keep what is already there\.\nNo: fill empty slots, and replace what is already there\.\nCancel: don't copy anything\./);
  assert.equal(open('Year 3 Gold').getRange('P11').getValue(), '', 'cancel copies nothing');
  assert.equal(world.dialogs, undefined, 'no HTML window');
});

test('Copy Week A: fills empty slots in each term, keeps breaks, existing entries and specialist lessons', () => {
  const { world, tools, open } = setUp();
  const sheet = open('Year 3 Gold'); // 2-week cycle: Term 1 Week A = J-N, Week B = P-T; Term 2 Week A = V-Z, Week B = AB-AF
  sheet.getRange('J11:J13').setValues([['English'], ['English'], ['Mathematics']]);
  sheet.getRange('V11').setValue('HSIE');
  sheet.getRange('P12').setValue('Religion');                               // existing entry
  sheet.getRange('P13').setValue('PDHPE').setNote('[SPEC_SYNC] Mr Smith');  // specialist lesson
  sheet.getRange('J29').setValue('Recess');                                 // break row (already Recess)

  const alerts = runCopy(world, tools, 'weeks', 'YES'); // fill empty slots only
  assert.equal(sheet.getRange('P11').getValue(), 'English');
  assert.equal(sheet.getRange('P12').getValue(), 'Religion', 'existing entry kept');
  assert.equal(sheet.getRange('P13').getValue(), 'PDHPE', 'specialist lesson kept');
  assert.equal(sheet.getRange('AB11').getValue(), 'HSIE', 'Term 2 copies from Term 2 Week A');
  assert.equal(sheet.getRange('P29').getValue(), 'Recess');
  assert.equal(alerts[1], 'Done. Filled 2 slot(s) (10 min).\n\n' +
    '1 slot(s) already had something in them and were left as they were.\n\n' +
    '1 specialist lesson slot(s) were left as they were.');

  runCopy(world, tools, 'weeks', 'NO'); // also replace what's there
  assert.equal(sheet.getRange('P12').getValue(), 'English');
  assert.equal(sheet.getRange('P13').getValue(), 'PDHPE');
  assert.match(runCopy(world, tools, 'weeks', 'NO')[1], /Nothing to copy/);
});

test('Copy Term 1 into Terms 2-4, and the 1-week timetable message for Copy Week A', () => {
  const { world, tools, open } = setUp();
  const sheet = open('K Blue'); // 1-week cycle: Term 1 = J-N, Term 2 = P-T, Term 3 = V-Z, Term 4 = AB-AF
  sheet.getRange('J11').setValue('English');
  assert.match(runCopy(world, tools, 'terms', 'YES')[1], /Filled 3 slot\(s\) \(15 min\)/);
  assert.deepEqual(['P11', 'V11', 'AB11'].map(a1 => sheet.getRange(a1).getValue()), ['English', 'English', 'English']);
  tools.copyWeekAToOtherWeeks();
  assert.match(world.alerts.pop(), /1-week timetable, so there are no other weeks to copy into\. Use "Copy Term 1 into Terms 2-4" instead\./);
});

/** Lays out a doc's HTML table rows and checks every row has 6 cells once rowspans are counted. */
function tableShape(html) {
  const rows = html.split('<tr').slice(1).map(row => row.split('</tr>')[0]);
  const covered = [];
  rows.forEach((row, r) => {
    const cells = row.match(/<t[dh][^>]*>/g) || [];
    let column = 0;
    cells.forEach(cell => {
      while (covered[r] && covered[r][column]) column++;
      const span = Number((cell.match(/rowspan="(\d+)"/) || [0, 1])[1]);
      for (let k = 0; k < span; k++) {
        covered[r + k] = covered[r + k] || [];
        covered[r + k][column] = true;
      }
      column++;
    });
  });
  return covered.map(row => row.filter(Boolean).length);
}

/** Makes a Google Doc of a tab from the menu, typing the answers into its prompt. */
function makeDoc(world, tools, open, name, ...answers) {
  open(name);
  world.promptAnswers = answers;
  tools.makeTimetableDoc();
}

test('Google Doc: a Sheets prompt numbers every week, then All terms (no window calling the script back)', () => {
  const { world, tools, open, lastDialog } = setUp();
  makeDoc(world, tools, open, 'Year 3 Gold', null); // cancelled
  const question = world.prompts.pop();
  assert.match(question, /^Makes a Google Doc of the "Year 3 Gold" timetable/);
  assert.match(question, /\n1   Term 1 - Week A\n2   Term 1 - Week B\n3   Term 2 - Week A\n/);
  assert.match(question, /\n9   All terms \(one page per week\)$/, 'A1 is selected: no blank option');
  assert.equal(world.docs, undefined);
  assert.equal(world.dialogs, undefined);

  open('K Blue').setActiveSelection('Q20'); // Term 2
  makeDoc(world, tools, open, 'K Blue', null);
  assert.match(world.prompts.pop(), /1   Term 1\n2   Term 2\n3   Term 3\n4   Term 4\n5   All terms \(one page per term\)\n\nOr leave it blank for Term 2, where your selected cell is\.$/);

  // Blank picks the selected week; a wrong answer asks again.
  open('Year 3 Gold').setActiveSelection('U12'); // the gap after Term 1 Week B counts as Week B
  makeDoc(world, tools, open, 'Year 3 Gold', '12', '');
  assert.match(world.prompts[world.prompts.length - 1], /^"12" isn't one of the numbers\. Type a number from 1 to 9\.\n\nMakes a Google Doc/);
  assert.equal(world.docs[0].name, 'Year 3 Gold - Term 1 - Week B');
  assert.match(lastDialog().html, /Your Google Doc is ready:<\/b> Year 3 Gold - Term 1 - Week B\./);
  assert.ok(!/google\.script\.run/.test(lastDialog().html));
});

test('Google Doc of a week: details, merged lessons with their times, coloured like the tab, listed on Start here', () => {
  const { world, ss, tools, open, lastDialog } = setUp();
  const sheet = open('Year 3 Gold');
  sheet.getRange('J11:J22').setValue('English');     // Monday 9:00-10:00
  sheet.getRange('K11:K16').setValue('Mathematics'); // Tuesday 9:00-9:30
  sheet.getRange('K17:K22').setValue('Assembly');    // Tuesday 9:30-10:00 (no colour rule but "not empty")
  makeDoc(world, tools, open, 'Year 3 Gold', '1');

  const doc = world.docs[0];
  assert.equal(doc.name, 'Year 3 Gold - Term 1 - Week A');
  assert.equal(doc.mimeType, 'application/vnd.google-apps.document');
  assert.deepEqual(doc.parents, ['Timetables']);
  assert.deepEqual(doc.page, { PageWidth: 842, PageHeight: 595, MarginTop: 28, MarginBottom: 28, MarginLeft: 28, MarginRight: 28 });
  assert.match(doc.html, /<h1>Year 3 Gold - Term 1 - Week A<\/h1>/);
  const details = doc.html.match(/<table class="details">([\s\S]*?)<\/table>/)[1];
  const pairs = [...details.matchAll(/<td class="label">([^<]*)<\/td><td>([^<]*)<\/td>/g)].map(m => [m[1], m[2]]);
  assert.deepEqual(pairs, [
    ['Class', 'Year 3 Gold'], ['Teacher', 'Ms Chen'], ['Grade', 'Year 3'], ['Timetable', '2-week cycle'],
    ['School', 'St Example Primary'], ['School day', '08:30 - 15:15'],
    ['Breaks', 'Recess 10:30 - 11:00, Lunch 13:00 - 13:45'],
    ['Generated', '21 October 2026, 2:05 PM by coordinator@school.example']
  ]);
  assert.match(doc.html, /<td rowspan="2" style="background-color:#D9EAD3;color:#6AA84F"><b>English<\/b><br><span class="small">09:00 - 10:00<\/span><\/td>/);
  assert.match(doc.html, /<td style="background-color:#FCE5CD;color:#E69138"><b>Mathematics<\/b><br><span class="small">09:00 - 09:30<\/span><\/td>/);
  assert.match(doc.html, /<td style="background-color:#F3F3F3;color:#B7B7B7"><b>Assembly<\/b>/);
  assert.match(doc.html, /style="background-color:#EEEEEE;color:#555555"><b>Recess<\/b><br><span class="small">10:30 - 11:00<\/span>/);
  const weekTable = doc.html.slice(doc.html.indexOf('<table><tr><th class="time">'));
  tableShape(weekTable).forEach((cells, r) => assert.equal(cells, 6, 'row ' + r));
  assert.ok(!doc.children.some(child => child.type === 'PAGE_BREAK'), 'one week, no page breaks');
  // The tables are stretched to the page: the time column narrow, the days equal.
  const tables = doc.children.filter(child => child.type === 'TABLE');
  assert.deepEqual(tables.map(table => table.widths), [[94, 299, 94, 299], [79, 141, 141, 141, 141, 143]]);
  assert.match(lastDialog().html, /href="https:\/\/docs\.google\.com\/document\/d\/doc1\/edit" target="_blank">Open the document</);

  // Listed on Start here, replacing the "None yet" line; the next one goes below it.
  const docsHeadings = ss.getRangeByName('StartHereDocs');
  const startHere = docsHeadings.getSheet();
  const row = docsHeadings.getRow() + 1;
  assert.deepEqual(startHere.getRange(row, 1, 1, 4).getValues()[0], [
    '=HYPERLINK("https://docs.google.com/document/d/doc1/edit","Year 3 Gold - Term 1 - Week A")',
    'Year 3 Gold, Term 1 - Week A', '21 October 2026, 2:05 PM', 'coordinator@school.example'
  ]);
  makeDoc(world, tools, open, 'K Blue', '2');
  assert.deepEqual(startHere.getRange(row + 1, 1, 1, 2).getValues()[0], [
    '=HYPERLINK("https://docs.google.com/document/d/doc2/edit","K Blue - Term 2")', 'K Blue, Term 2'
  ]);
});

test('Google Doc of all terms: one page per week, details once at the top', () => {
  const { world, tools, open, lastDialog } = setUp();
  makeDoc(world, tools, open, 'Year 3 Gold', '9');
  assert.match(lastDialog().html, /8 pages, one per week\./);
  const doc = world.docs[0];
  assert.equal(doc.name, 'Year 3 Gold - All terms');
  assert.match(doc.html, /<h1>Year 3 Gold<\/h1><table class="details">/);
  const weekHeadings = [...doc.html.matchAll(/<h2>([^<]*)<\/h2>/g)].map(m => m[1]);
  assert.deepEqual(weekHeadings, ['Term 1 - Week A', 'Term 1 - Week B', 'Term 2 - Week A', 'Term 2 - Week B',
    'Term 3 - Week A', 'Term 3 - Week B', 'Term 4 - Week A', 'Term 4 - Week B']);
  // A page break before every week heading but the first.
  const order = doc.children.map(child => child.type === 'PARAGRAPH' ? child.heading : child.type);
  assert.deepEqual(order.slice(0, 7), ['HEADING1', 'TABLE', 'HEADING2', 'TABLE', 'PAGE_BREAK', 'HEADING2', 'TABLE']);
  assert.equal(doc.children.filter(child => child.type === 'PAGE_BREAK').length, 7);
});

test('Google Doc of a class tab: synced specialist lessons show the specialist, short lessons keep their times', () => {
  const { world, tools, open } = setUp();
  open('SPEC_Mr Smith').getRange('J11:J16').setValue('K Blue'); // Monday 9:00-9:30
  tools.syncSpecialistSchedules();
  const sheet = open('K Blue');
  sheet.getRange('J17').setValue('Religion');                    // Monday 9:30-9:35, one row
  makeDoc(world, tools, open, 'K Blue', '1');
  const html = world.docs[0].html;
  assert.match(html, /<b>PDHPE<\/b><br><span class="small">Mr Smith<\/span><br><span class="small">09:00 - 09:30<\/span>/);
  assert.match(html, /<b>Religion<\/b><br><span class="small">09:30 - 09:35<\/span>/);
});

test('Google Doc of a specialist tab: specialist details, classes keep their colours', () => {
  const { world, tools, open } = setUp();
  const spec = open('SPEC_Mr Smith');
  spec.getRange('J11:J16').setValue('K Blue');
  makeDoc(world, tools, open, 'SPEC_Mr Smith', '1');
  const doc = world.docs[0];
  assert.equal(doc.name, 'Mr Smith - Term 1');
  assert.match(doc.html, /<td class="label">Specialist<\/td><td>Mr Smith<\/td><td class="label">KLA<\/td><td>PDHPE<\/td>/);
  assert.match(doc.html, /<td class="label">Timetable<\/td><td>1-week cycle<\/td>/);
  assert.match(doc.html, /style="background-color:#FFADAD"><b>K Blue<\/b>/);
});

test('Google Doc of the compliance summary: a row per class, in-range values green, out-of-range red, page width', () => {
  const { world, ss, tools, lastDialog } = setUp();
  setCompliance(ss.getSheetByName('K Red'), BALANCED, { empty: 0, compliant: true, amadeus: 'N/A', tutorials: 0, tutorialsPercent: 0 });
  tools.makeComplianceDoc();
  const doc = world.docs[0];
  assert.equal(doc.name, 'St Example Primary - Compliance summary');
  assert.match(doc.html, /<td class="label">Classes compliant<\/td><td>1 of 3<\/td><td class="label">Generated<\/td><td>21 October 2026, 2:05 PM by coordinator@school\.example<\/td>/);
  assert.match(doc.html, /<th>Class<\/th><th>Teacher<\/th><th>English<\/th>/);
  assert.match(doc.html, /<td><b>K Red<\/b><\/td><td>Mr Brown<\/td><td style="background-color:#D9EAD3;color:#274E13">33\.0%<\/td>/);
  assert.match(doc.html, /<td style="background-color:#F4CCCC;color:#990000"><b>3\.0%<\/b><\/td>/); // Creative Arts under
  assert.match(doc.html, /<td style="background-color:#D9EAD3;color:#274E13"><b>Yes<\/b><\/td>/);
  assert.match(lastDialog().html, /Open the document/);
  // Both tables fill the width inside the margins (786pt).
  doc.children.filter(child => child.type === 'TABLE').forEach(table => {
    assert.equal(table.widths.reduce((sum, width) => sum + width, 0), 786);
  });
});

test('without the Drive API service, making a Google Doc explains what to turn on', () => {
  const world = createWorld();
  makeTemplate(world, TEMPLATE_ID);
  loadScript('wizard/Code.gs', world.globals).generateTimetable(makePayload());
  world.newExecution();
  const ss = world.files[world.lastCopyId].spreadsheet;
  world.active = ss;
  delete world.globals.Drive;
  const tools = loadScript(['template/TemplateSync.gs', 'template/TimetableTools.gs'], world.globals);
  ss.setActiveSheet(ss.getSheetByName('K Blue'));
  tools.makeTimetableDoc();
  assert.match(world.alerts.pop(), /needs the Drive API service/);
  assert.equal(world.docs, undefined);
});
