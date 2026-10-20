'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createWorld, makeTemplate } = require('./helpers/fakeSpreadsheet');
const { loadScript } = require('./helpers/loadScript');
const { makePayload, makeLargeSchoolPayload } = require('./helpers/payload');

const TEMPLATE_ID = '1x3pzraH_eG2nsQyiZPE9WGAzZ1wZIFGM2JUVF4W6FCk';
const DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];

/** Values made inside the script sandbox, as ordinary objects for deepEqual. */
const plain = value => JSON.parse(JSON.stringify(value));

function setUp() {
  const world = createWorld();
  makeTemplate(world, TEMPLATE_ID);
  const wizard = loadScript('wizard/Code.gs', world.globals);
  return { world, wizard };
}

function generate(payload) {
  const { world, wizard } = setUp();
  const url = wizard.generateTimetable(payload || makePayload());
  const file = world.files[world.lastCopyId];
  return { world, wizard, url, file, ss: file.spreadsheet };
}

/** The old formula for the last grid column, before getGridLayout existed. */
function oldLastColumn(cycle) {
  return 9 + 4 * (cycle * 5 + Math.max(0, cycle - 1)) + 3;
}

test('getGridLayout matches the original column arithmetic for 1-10 week cycles', () => {
  const { wizard } = setUp();
  for (let cycle = 1; cycle <= 10; cycle++) {
    const layout = wizard.getGridLayout(cycle);
    assert.equal(layout.lastColumn, oldLastColumn(cycle), 'cycle ' + cycle);
    assert.equal(layout.blocks.length, 4 * cycle);
    assert.equal(layout.spacers.length, 4 * (cycle - 1) + 3);
  }
  const layout = wizard.getGridLayout(2);
  assert.deepEqual(plain(layout.blocks.map(b => [b.label, b.startColumn])), [
    ['Term 1 - Week A', 10], ['Term 1 - Week B', 16], ['Term 2 - Week A', 22], ['Term 2 - Week B', 28],
    ['Term 3 - Week A', 34], ['Term 3 - Week B', 40], ['Term 4 - Week A', 46], ['Term 4 - Week B', 52]
  ]);
  assert.deepEqual(plain(layout.spacers.slice(0, 2)), [{ column: 15, width: 20 }, { column: 21, width: 40 }]);
});

test('createTimeSlots labels every 5-minute slot like the old Utilities.formatDate version', () => {
  const { wizard } = setUp();
  const slots = wizard.createTimeSlots('08:30', '15:15');
  assert.equal(slots.length, 81);
  assert.equal(slots[0].label, '08:30 - 08:35');
  assert.equal(slots[7].label, '09:05 - 09:10');
  assert.equal(slots[80].label, '15:10 - 15:15');
  assert.throws(() => wizard.createTimeSlots('10:00', '09:00'), /Invalid school start or end time/);
});

test('getBreakLabels and getBreakRuns find the slots each break covers', () => {
  const { wizard } = setUp();
  const slots = wizard.createTimeSlots('08:30', '15:15');
  const labels = wizard.getBreakLabels(slots, [
    { label: 'Recess', start: '10:30', end: '11:00' },
    { label: 'Lunch', start: '13:00', end: '13:45' }
  ]);
  assert.equal(labels.filter(l => l === 'Recess').length, 6);
  assert.equal(labels.filter(l => l === 'Lunch').length, 9);
  assert.equal(labels[24], 'Recess'); // 10:30 - 10:35
  assert.equal(labels[23], '');
  assert.deepEqual(plain(wizard.getBreakRuns(labels)), [
    { startIndex: 24, length: 6 }, { startIndex: 54, length: 9 }
  ]);
});

test('toA1 and columnToLetters', () => {
  const { wizard } = setUp();
  assert.equal(wizard.toA1(1, 10, 1, 5), 'J1:N1');
  assert.equal(wizard.toA1(5, 28, 1, 1), 'AB5');
  assert.equal(wizard.columnToLetters(713), 'AAK');
});

test('generateTimetable builds the file, names it, and lists tabs in order', () => {
  const { url, file, ss } = generate();
  assert.match(url, /^https:\/\/docs\.google\.com\/spreadsheets\/d\//);
  assert.match(file.handle.getName(), /^Generated Timetable - Resource First - /);
  assert.deepEqual(ss.getSheets().map(s => s.getName()), [
    'DATA', 'OVERVIEW', 'TEMPLATE', 'SPEC_Template',
    'SPEC_Mr Smith', 'SPEC_Ms Jones', 'K Blue', 'K Red', 'Year 3 Gold'
  ]);
  ss.getSheets().slice(4).forEach(sheet => assert.equal(sheet.isSheetHidden(), false, sheet.getName()));
});

test('DATA tab: plain text, lists and named ranges', () => {
  const { ss } = generate();
  const data = ss.getSheetByName('DATA');
  assert.equal(data.getRange('A2').getNumberFormats()[0][0], '@');
  assert.deepEqual(ss.getRangeByName('ClassList').getValues().flat(), ['K Blue', 'K Red', 'Year 3 Gold']);
  assert.deepEqual(ss.getRangeByName('BreakList').getValues().flat(), ['Recess', 'Lunch']);
  assert.deepEqual(ss.getRangeByName('SpecialistMapping').getValues(), [['Mr Smith', 'PDHPE'], ['Ms Jones', 'Creative Arts']]);
  const klas = ss.getRangeByName('KLAList').getValues().flat();
  assert.ok(klas.includes('Amadeus - Tutorials'));
  assert.ok(!klas.includes('Non-Curricular'));
  // DATA column N (ActivityList) is no longer written.
  assert.equal(data.getRange('N1').getValue(), '');
});

test('class tab: details, headers, times, breaks and dropdowns', () => {
  const { ss } = generate();
  const sheet = ss.getSheetByName('Year 3 Gold');

  assert.deepEqual(sheet.getRange('C27:C30').getValues().flat(), ['Ms Chen', 'Year 3 Gold', 'Year 3', '2 week cycle']);
  assert.deepEqual(sheet.getRange('B31:C31').getValues()[0], ['Amadeus', true]);
  assert.equal(ss.getSheetByName('K Blue').getRange('C31').getValue(), false);

  // Headers: block labels on row 1 (merged), day names on row 4, spacers blank.
  assert.equal(sheet.getRange('J1').getValue(), 'Term 1 - Week A');
  assert.equal(sheet.getRange('P1').getValue(), 'Term 1 - Week B');
  assert.equal(sheet.getRange('K1').getValue(), '');
  assert.deepEqual(sheet.getRange('J4:O4').getValues()[0], DAYS.concat(['']));
  assert.equal(sheet._merges.filter(m => m.row === 1).length, 8);
  assert.equal(sheet.getColumnWidth(15), 20);
  assert.equal(sheet.getColumnWidth(21), 40);
  assert.equal(sheet.getColumnWidth(10), 120);

  // Times in column I, rows 5 to 85.
  assert.equal(sheet.getRange('I5').getValue(), '08:30 - 08:35');
  assert.equal(sheet.getRange('I85').getValue(), '15:10 - 15:15');

  // Recess (10:30-11:00) is rows 29-34: labelled in day columns, grey, no dropdown.
  assert.deepEqual(sheet.getRange('J29:O29').getValues()[0], ['Recess', 'Recess', 'Recess', 'Recess', 'Recess', '']);
  assert.equal(sheet.getRange('J34').getValue(), 'Recess');
  assert.equal(sheet.getRange('J35').getValue(), '');
  assert.equal(sheet.getRange('J29').getBackgrounds()[0][0], '#F2F2F2');
  assert.equal(sheet.getRange('J29').getDataValidations()[0][0], null);

  // Other grid cells offer the KLA list; spacers don't.
  const rule = sheet.getRange('J5').getDataValidations()[0][0];
  assert.ok(rule && rule.range === ss.getRangeByName('KLAList'));
  assert.equal(sheet.getRange('O5').getDataValidations()[0][0], null);

  // TEMPLATE!B5's conditional formatting pasted once per week block.
  const pasted = sheet.getConditionalFormatRules().filter(r => r.pastedFrom === 'B5');
  assert.equal(pasted.length, 8 * 2); // 8 blocks x the 2 rules on B5
});

test('copied class tabs get the same grid as the first tab with that cycle', () => {
  const { ss } = generate();
  const first = ss.getSheetByName('K Blue');
  const copy = ss.getSheetByName('K Red');
  const grid = sheet => sheet.getRange(1, 9, 85, 24).getValues();
  assert.deepEqual(grid(copy), grid(first));
  assert.deepEqual(copy.getRange('C27:C30').getValues().flat(), ['Mr Brown', 'K Red', 'Kindergarten', '1 week cycle']);
  assert.equal(copy.getConditionalFormatRules().length, first.getConditionalFormatRules().length);
});

test('specialist tab: details, class dropdowns, class colours, break and clash rules', () => {
  const { ss } = generate();
  const sheet = ss.getSheetByName('SPEC_Ms Jones');
  assert.deepEqual([sheet.getRange('B1').getValue(), sheet.getRange('B3').getValue(), sheet.getRange('D3').getValue()],
    ['Ms Jones', 'Creative Arts', 2]);
  assert.equal(sheet.getRange('J5').getDataValidations()[0][0].range, ss.getRangeByName('ClassList'));
  assert.equal(sheet.getRange('J29').getValue(), 'Recess');

  const formulas = sheet.getConditionalFormatRules().map(r => r._formula);
  assert.ok(formulas[0].includes('INDIRECT("BreakList")'), 'break rule comes first');
  assert.ok(formulas[1].includes("COUNTIF(INDIRECT(\"'SPEC_Mr Smith'!\""), 'clash rule second');
  assert.ok(formulas.includes('=J5="K Blue"'), 'class colour rule kept');
  // SPEC_Template's own rules are not carried over, and it is not treated as a specialist.
  assert.ok(!formulas.some(f => f && f.includes('SPEC_Template')));
  assert.equal(ss.getSheetByName('SPEC_Template').getConditionalFormatRules().length, 1);
});

test('setup problems are reported before anything is copied', () => {
  const { world, wizard } = setUp();
  assert.throws(
    () => wizard.generateTimetable(makePayload({ classes: [['Data', 'A', 'Year 1', 1]] })),
    /^Error: Please fix the following and try again:\n- Class "Data" is a name the timetable uses/
  );
  assert.equal(world.lastCopyId, undefined, 'no copy made');
  assert.match(JSON.parse(world.cache.job_test).status, /^Please fix/);
});

test('no access to the template gives a plain-English message', () => {
  const { world, wizard } = setUp();
  world.properties.TEMPLATE_ID = 'someone-elses-template';
  assert.throws(() => wizard.generateTimetable(makePayload()), /You don't have access to the master timetable template/);
});

test('a run that fails part-way leaves the copy marked INCOMPLETE', () => {
  const world = createWorld();
  const template = makeTemplate(world, TEMPLATE_ID);
  template._sheets.splice(template._sheets.findIndex(s => s._name === 'SPEC_Template'), 1);
  const wizard = loadScript('wizard/Code.gs', world.globals);
  assert.throws(() => wizard.generateTimetable(makePayload()), /Something went wrong[\s\S]*SPEC_Template[\s\S]*INCOMPLETE - /);
  assert.match(world.files[world.lastCopyId].handle.getName(), /^INCOMPLETE - Generated Timetable/);
});

test('large school: far fewer spreadsheet calls than the previous generator', () => {
  const payload = makeLargeSchoolPayload();

  const before = setUp();
  const oldWizard = loadScript('tests/fixtures/Code.before-phase4.gs', before.world.globals);
  before.world.resetCalls();
  oldWizard.generateTimetable(payload);
  const oldCalls = before.world.calls;

  const after = setUp();
  after.world.resetCalls();
  after.wizard.generateTimetable(payload);
  const newCalls = after.world.calls;

  console.log(`  28 classes (3-week cycle), 6 specialists: ${oldCalls} calls before, ${newCalls} after`);
  assert.ok(newCalls * 5 < oldCalls, `expected at least 5x fewer calls (${oldCalls} -> ${newCalls})`);

  // Same visible result: compare one class tab and one specialist tab.
  const pick = (world, name) => world.files[world.lastCopyId].spreadsheet.getSheetByName(name);
  ['Year 6 Green', 'SPEC_RE Teacher'].forEach(name => {
    const oldSheet = pick(before.world, name);
    const newSheet = pick(after.world, name);
    assert.deepEqual(newSheet.getRange(1, 1, 85, newSheet.getMaxColumns()).getValues(),
      oldSheet.getRange(1, 1, 85, oldSheet.getMaxColumns()).getValues(), name + ' values');
    assert.deepEqual(newSheet.getRange(5, 10, 81, 20).getBackgrounds(),
      oldSheet.getRange(5, 10, 81, 20).getBackgrounds(), name + ' backgrounds');
    assert.deepEqual(newSheet._merges.map(m => JSON.stringify(m)).sort(),
      oldSheet._merges.map(m => JSON.stringify(m)).sort(), name + ' merges');

    const dropdowns = sheet => sheet.getRange(5, 9, 81, sheet.getMaxColumns() - 8).getDataValidations()
      .map(row => row.map(rule => (rule ? rule.range.getA1Notation() : null)));
    assert.deepEqual(dropdowns(newSheet), dropdowns(oldSheet), name + ' dropdowns');

    const widths = sheet => Array.from({ length: sheet.getMaxColumns() }, (_, i) => sheet.getColumnWidth(i + 1));
    assert.deepEqual(widths(newSheet), widths(oldSheet), name + ' column widths');

    // Which cells TEMPLATE!B5's pasted conditional formatting covers (the old
    // version pasted per column, the new one per week block).
    const pastedCoverage = sheet => {
      const covered = new Set();
      sheet.getConditionalFormatRules().filter(r => r.pastedFrom).forEach(rule => rule._ranges.forEach(range => {
        for (let r = 0; r < range.getNumRows(); r++) {
          for (let c = 0; c < range.getNumColumns(); c++) covered.add((range.getRow() + r) + ':' + (range.getColumn() + c));
        }
      }));
      return [...covered].sort();
    };
    assert.deepEqual(pastedCoverage(newSheet), pastedCoverage(oldSheet), name + ' conditional formatting coverage');
  });
});

test('the Apps Script load test (LoadTest.gs) runs against a valid setup', () => {
  const world = createWorld();
  makeTemplate(world, TEMPLATE_ID);
  const wizard = loadScript(['wizard/Code.gs', 'wizard/LoadTest.gs'], world.globals);
  assert.equal(typeof wizard.runLoadTest(), 'number');
  const ss = world.files[world.lastCopyId].spreadsheet;
  assert.equal(ss.getSheets().length, 4 + 6 + 28);
});
