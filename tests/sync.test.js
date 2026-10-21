'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createWorld, makeTemplate } = require('./helpers/fakeSpreadsheet');
const { loadScript } = require('./helpers/loadScript');
const { makePayload } = require('./helpers/payload');

const TEMPLATE_ID = '1x3pzraH_eG2nsQyiZPE9WGAzZ1wZIFGM2JUVF4W6FCk';
const plain = value => JSON.parse(JSON.stringify(value));

/**
 * Generates a timetable with the wizard, then loads TemplateSync.gs as the
 * generated file's own script (a separate project, as in Google).
 */
function setUp(payloadOptions) {
  const world = createWorld();
  makeTemplate(world, TEMPLATE_ID);
  loadScript('wizard/Code.gs', world.globals).generateTimetable(makePayload(payloadOptions));
  world.newExecution(); // the menu runs later, in its own execution
  const ss = world.files[world.lastCopyId].spreadsheet;
  world.active = ss;
  const sync = loadScript('template/TemplateSync.gs', world.globals);
  return { world, ss, sync, sheet: name => ss.getSheetByName(name) };
}

// Row 5 is 08:30, so 09:00 is row 11. Week A Monday is column J on every tab.
const cell = (sheet, a1) => ({
  value: sheet.getRange(a1).getValue(),
  note: sheet.getRange(a1).getNotes()[0][0],
  background: sheet.getRange(a1).getBackgrounds()[0][0]
});

test('sync copies a specialist lesson onto the class tab built by the wizard', () => {
  const { sync, sheet, world } = setUp();
  sheet('SPEC_Mr Smith').getRange('K11').setValue('K Blue'); // Tuesday 09:00
  sync.syncSpecialistSchedules();
  assert.deepEqual(cell(sheet('K Blue'), 'K11'), { value: 'PDHPE', note: '[SPEC_SYNC] Mr Smith', background: '#ffffff' });
  // The same Tuesday in Term 4 is a different column and stays empty.
  assert.equal(sheet('K Blue').getRange('AC11').getValue(), '');
  assert.match(world.alerts.pop(), /Populated 1 specialist slot/);
});

test('a 1-week specialist lands in every week of a 2-week class', () => {
  const { sync, sheet } = setUp();
  sheet('SPEC_Mr Smith').getRange('M13').setValue('Year 3 Gold'); // Thursday 09:10
  sync.syncSpecialistSchedules();
  const yearThree = sheet('Year 3 Gold');
  assert.equal(yearThree.getRange('M13').getValue(), 'PDHPE'); // Term 1 week A Thursday
  assert.equal(yearThree.getRange('S13').getValue(), 'PDHPE'); // Term 1 week B Thursday
});

test('a 2-week specialist on a 1-week class is synced with a cycle warning', () => {
  const { sync, sheet, world } = setUp();
  sheet('SPEC_Ms Jones').getRange('T13').setValue('K Blue'); // week B Friday 09:10
  sync.syncSpecialistSchedules();
  assert.equal(sheet('K Blue').getRange('N13').getValue(), 'Creative Arts');
  assert.match(world.alerts.pop(), /Ms Jones \(2-week cycle\) and K Blue \(1-week cycle\)/);
});

test("a clash with a teacher's entry is flagged, then the entry comes back once resolved", () => {
  const { sync, sheet, world } = setUp();
  const kBlue = sheet('K Blue');
  kBlue.getRange('J11').setValue('English');
  sheet('SPEC_Mr Smith').getRange('J11').setValue('K Blue');

  sync.syncSpecialistSchedules();
  assert.deepEqual(cell(kBlue, 'J11'), {
    value: 'CONFLICT: English vs Mr Smith', note: '[SPEC_CONFLICT] English', background: '#FF0000'
  });
  assert.match(world.alerts.pop(), /K Blue J11: class entry \(English\) vs Mr Smith \(PDHPE\)/);

  sync.syncSpecialistSchedules(); // still clashing: flagged again, original kept
  assert.equal(kBlue.getRange('J11').getNotes()[0][0], '[SPEC_CONFLICT] English');

  sheet('SPEC_Mr Smith').getRange('J11').setValue('');
  sync.syncSpecialistSchedules();
  assert.deepEqual(cell(kBlue, 'J11'), { value: 'English', note: '', background: '#ffffff' });
});

test("a teacher's entry that already matches the specialist's KLA is taken over, not flagged", () => {
  const { sync, sheet } = setUp();
  sheet('K Blue').getRange('K11').setValue('PDHPE');
  sheet('SPEC_Mr Smith').getRange('K11').setValue('K Blue');
  sync.syncSpecialistSchedules();
  assert.equal(sheet('K Blue').getRange('K11').getNotes()[0][0], '[SPEC_SYNC] Mr Smith');
});

test('two specialists in the same class slot are flagged as a conflict', () => {
  const { sync, sheet, world } = setUp();
  sheet('SPEC_Mr Smith').getRange('L11').setValue('K Blue');
  sheet('SPEC_Ms Jones').getRange('L11').setValue('K Blue');
  sync.syncSpecialistSchedules();
  assert.equal(sheet('K Blue').getRange('L11').getValue(), 'CONFLICT: Mr Smith vs Ms Jones');
  assert.match(world.alerts.pop(), /1 conflict\(s\)/);
});

test('moved or deleted lessons are removed on the next sync; an unchanged re-sync changes nothing', () => {
  const { sync, sheet, world } = setUp();
  const spec = sheet('SPEC_Mr Smith');
  spec.getRange('J11').setValue('K Blue');
  spec.getRange('K12').setValue('K Blue');
  sync.syncSpecialistSchedules();

  spec.getRange('K12').setValue('');
  sync.syncSpecialistSchedules();
  assert.equal(sheet('K Blue').getRange('K12').getValue(), '');
  assert.equal(sheet('K Blue').getRange('J11').getValue(), 'PDHPE');
  assert.match(world.alerts.pop(), /Removed 1 slot\(s\) left over/);

  const snapshot = () => JSON.stringify([sheet('K Blue').getRange('J5:AF85').getValues(), sheet('K Blue').getRange('J5:AF85').getNotes()]);
  const before = snapshot();
  sync.syncSpecialistSchedules();
  assert.equal(snapshot(), before);
});

test("a synced slot someone has since edited by hand is kept as their entry", () => {
  const { sync, sheet } = setUp();
  sheet('SPEC_Mr Smith').getRange('J11').setValue('K Blue');
  sync.syncSpecialistSchedules();
  sheet('K Blue').getRange('J11').setValue('Mathematics');
  sheet('SPEC_Mr Smith').getRange('J11').setValue('');
  sync.syncSpecialistSchedules();
  assert.deepEqual(cell(sheet('K Blue'), 'J11'), { value: 'Mathematics', note: '', background: '#ffffff' });
});

test('break labels on specialist tabs are not treated as classes', () => {
  const { sync, sheet, world } = setUp();
  sync.syncSpecialistSchedules(); // every specialist tab has Recess/Lunch rows
  assert.match(world.alerts.pop(), /Populated 0 specialist slot/);
  assert.equal(sheet('K Blue').getRange('J29').getValue(), 'Recess');
});

test('Lock Synced Cells locks synced slots and rebuilds its locks after changes', () => {
  const { sync, sheet } = setUp();
  const spec = sheet('SPEC_Mr Smith');
  spec.getRange('K11:K12').setValues([['K Blue'], ['K Blue']]);
  spec.getRange('L12').setValue('K Blue');
  sync.syncSpecialistSchedules();
  sync.lockSyncedSpecialistCells();
  const locks = () => sheet('K Blue').getProtections().map(p => p.getRange().getA1Notation()).sort();
  assert.deepEqual(locks(), ['K11', 'K12:L12']);

  spec.getRange('L12').setValue('');
  sync.syncSpecialistSchedules();
  sync.lockSyncedSpecialistCells();
  assert.deepEqual(locks(), ['K11:K12']);
});

test('mergeCellsIntoRectangles covers every marked cell with few rectangles', () => {
  const { sync } = setUp();
  const grid = [
    [true, true, false],
    [true, true, false],
    [false, false, true]
  ];
  assert.deepEqual(plain(sync.mergeCellsIntoRectangles(grid, 3, 3)), [
    { rowStart: 0, rowEnd: 1, colStart: 0, colEnd: 1 },
    { rowStart: 2, rowEnd: 2, colStart: 2, colEnd: 2 }
  ]);
});

test('header parsing and cycle helpers', () => {
  const { sync } = setUp();
  assert.deepEqual(plain(sync.parseTermWeekLabel('Term 3 - Week C')), { term: 3, week: 2 });
  assert.equal(sync.parseTermWeekLabel('Week A - Term 1'), null);
  assert.equal(sync.computeLcm(2, 3), 6);
  assert.equal(sync.computeLcm(1, 4), 4);
});
