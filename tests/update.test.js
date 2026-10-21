'use strict';

// Changing a timetable that has already been generated (PLAN step 32) and
// reading its setup back into the wizard (steps 32 and 33).

const test = require('node:test');
const assert = require('node:assert/strict');
const { createWorld, makeTemplate } = require('./helpers/fakeSpreadsheet');
const { loadScript } = require('./helpers/loadScript');
const { makePayload } = require('./helpers/payload');

const TEMPLATE_ID = '1x3pzraH_eG2nsQyiZPE9WGAzZ1wZIFGM2JUVF4W6FCk';
const plain = value => JSON.parse(JSON.stringify(value));

// Default school: Mr Smith (PDHPE, 1 week), Ms Jones (Creative Arts, 2 weeks);
// K Blue, K Red (Kindergarten, 1 week), Year 3 Gold (Year 3, 2 weeks).
const CLASSES = [['K Blue', 'Ms Adams', 'Kindergarten', 1], ['K Red', 'Mr Brown', 'Kindergarten', 1], ['Year 3 Gold', 'Ms Chen', 'Year 3', 2]];
const SPECIALISTS = [['Mr Smith', 'PDHPE', 1], ['Ms Jones', 'Creative Arts', 2]];

function setUp() {
  const world = createWorld();
  makeTemplate(world, TEMPLATE_ID);
  const wizard = loadScript('wizard/Code.gs', world.globals);
  const url = wizard.generateTimetable(makePayload());
  world.newExecution();
  const ss = world.files[world.lastCopyId].spreadsheet;
  world.active = ss;
  const tools = loadScript(['template/TemplateSync.gs', 'template/TimetableTools.gs'], world.globals);
  const sheet = name => ss.getSheetByName(name);
  const update = (options, renames) => {
    const result = wizard.updateTimetable({ jobId: 'job_update', fileId: url, exportVariables: makePayload(options).exportVariables, renames: renames || {} });
    world.newExecution(); // look at the result as someone opening the file would
    return result;
  };
  return { world, wizard, ss, url, tools, sheet, update };
}

test('readTimetableSetup gives the wizard the setup saved in a timetable', () => {
  const { wizard, url, ss } = setUp();
  const setup = wizard.readTimetableSetup(url);
  assert.deepEqual(plain(setup.exportVariables), plain(makePayload().exportVariables));
  assert.equal(setup.fileId, ss.getId());
  assert.equal(setup.fileName, 'St Example Primary - Timetable - ' + new Date().toLocaleDateString('en-AU'));
  assert.equal(setup.url, url);
  // A link with more after the ID, or the ID on its own, work too.
  const id = '1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789_-abcde';
  assert.equal(wizard.getFileIdFromLink('https://docs.google.com/spreadsheets/d/' + id + '/edit#gid=0'), id);
  assert.equal(wizard.getFileIdFromLink(' ' + id + ' '), id);
  assert.equal(wizard.getFileIdFromLink('hello'), '');
});

test('readTimetableSetup explains a bad link, or a sheet the wizard did not make', () => {
  const { wizard } = setUp();
  assert.throws(() => wizard.readTimetableSetup('my timetable'), /doesn't look like a link to a Google Sheet/);
  assert.throws(() => wizard.readTimetableSetup('https://docs.google.com/spreadsheets/d/nothing-here/edit'), /couldn't be opened/);
  assert.throws(() => wizard.readTimetableSetup(TEMPLATE_ID), /wasn't made by the timetable wizard/);
});

test('classes: added, renamed (keeping their tab and lessons), archived, and teachers changed', () => {
  const { sheet, update, wizard, url, ss } = setUp();
  sheet('SPEC_Mr Smith').getRange('J11').setValue('K Red');
  sheet('K Red').getRange('K11').setValue('English');
  sheet('Year 3 Gold').getRange('J11').setValue('Mathematics');

  const result = update({ classes: [
    ['K Blue', 'Ms Adams-Lee', 'Kindergarten', 1],
    ['K Green', 'Mr Brown', 'Kindergarten', 1],
    ['Year 3 Silver', 'Mr Diaz', 'Year 3', 2]
  ] }, { classes: [{ from: 'K Red', to: 'K Green' }] });

  // Renamed: same tab (its lessons kept), new name everywhere.
  assert.equal(sheet('K Red'), null);
  assert.equal(sheet('K Green').getRange('K11').getValue(), 'English');
  assert.equal(sheet('K Green').getRange('C28').getValue(), 'K Green');
  assert.equal(sheet('SPEC_Mr Smith').getRange('J11').getValue(), 'K Green');
  // Archived: renamed and hidden, never deleted.
  assert.equal(sheet('Year 3 Gold'), null);
  const archived = sheet('Archived - Year 3 Gold');
  assert.equal(archived.isSheetHidden(), true);
  assert.equal(archived.getRange('J11').getValue(), 'Mathematics');
  // Added: a new tab with its details.
  assert.deepEqual(sheet('Year 3 Silver').getRange('C27:C30').getValues().flat(), ['Mr Diaz', 'Year 3 Silver', 'Year 3', '2 week cycle']);
  assert.equal(sheet('K Blue').getRange('C27').getValue(), 'Ms Adams-Lee');

  // DATA and the saved setup follow.
  assert.deepEqual(ss.getRangeByName('ClassList').getValues().flat(), ['K Blue', 'K Green', 'Year 3 Silver']);
  assert.deepEqual(plain(wizard.readTimetableSetup(url).exportVariables.filter(v => v.name === 'Class_2_Name').map(v => v.value)), ['K Green']);
  // Start here lists the classes as they are now, each linked to its tab.
  const startHere = sheet('Start here');
  const listed = startHere.getRange('A18:A20').getValues().flat();
  assert.deepEqual(listed.map(f => f.match(/"([^"]+)"\)$/)[1]), ['K Blue', 'K Green', 'Year 3 Silver']);
  assert.match(listed[2], new RegExp('#gid=' + sheet('Year 3 Silver').getSheetId()));

  assert.deepEqual(plain(result.changes), [
    'K Blue: teacher Ms Adams to Ms Adams-Lee',
    'Rename class "K Red" to "K Green"',
    'Add class "Year 3 Silver" (Year 3, Mr Diaz)',
    'Archive class "Year 3 Gold"'
  ]);
  assert.match(result.notes[0], /hidden, not deleted/);
});

test('specialists: a rename and a KLA change carry their synced lessons; an archived specialist\'s are taken off', () => {
  const { sheet, update, tools, world } = setUp();
  sheet('SPEC_Mr Smith').getRange('J11').setValue('K Blue');
  sheet('SPEC_Ms Jones').getRange('J13').setValue('Year 3 Gold');
  tools.syncSpecialistSchedules();
  assert.equal(sheet('K Blue').getRange('J11').getValue(), 'PDHPE');
  assert.equal(sheet('Year 3 Gold').getRange('J13').getValue(), 'Creative Arts');

  update({ specialists: [['Mr Smithers', 'HSIE', 1]] }, { specialists: [{ from: 'Mr Smith', to: 'Mr Smithers' }] });

  assert.equal(sheet('SPEC_Mr Smith'), null);
  assert.deepEqual([sheet('SPEC_Mr Smithers').getRange('B1').getValue(), sheet('SPEC_Mr Smithers').getRange('B3').getValue()], ['Mr Smithers', 'HSIE']);
  assert.equal(sheet('SPEC_Mr Smithers').getRange('J11').getValue(), 'K Blue');
  assert.deepEqual([sheet('K Blue').getRange('J11').getValue(), sheet('K Blue').getRange('J11').getNotes()[0][0]], ['HSIE', '[SPEC_SYNC] Mr Smithers']);
  assert.deepEqual([sheet('Year 3 Gold').getRange('J13').getValue(), sheet('Year 3 Gold').getRange('J13').getNotes()[0][0]], ['', '']);
  assert.equal(sheet('Archived - SPEC_Ms Jones').isSheetHidden(), true);

  // Syncing again finds the renamed specialist's lesson already in place: no clash.
  world.active = world.files[world.lastCopyId].spreadsheet;
  tools.syncSpecialistSchedules();
  const message = world.alerts.pop();
  assert.match(message, /Populated 1 specialist slot/);
  assert.doesNotMatch(message, /conflict/i);
  assert.equal(sheet('K Blue').getRange('J11').getValue(), 'HSIE');
});

test('bell times, breaks and the cycle of an existing class or specialist are refused, and nothing changes', () => {
  const { update, wizard, url, sheet } = setUp();
  const before = plain(wizard.readTimetableSetup(url).exportVariables);
  assert.throws(() => update({ start: '08:45', breaks: [['Recess', '10:30', '11:00']] }),
    /can't be made to an existing timetable[\s\S]*School day: 08:30 - 15:15 to 08:45 - 15:15[\s\S]*Breaks: /);
  assert.throws(() => update({ classes: [CLASSES[0], CLASSES[1], ['Year 3 Gold', 'Ms Chen', 'Year 3', 1]], specialists: [['Mr Smith', 'PDHPE', 2], SPECIALISTS[1]] }),
    /Class "Year 3 Gold": cycle 2 to 1 weeks[\s\S]*Specialist "Mr Smith": cycle 1 to 2 weeks/);
  assert.deepEqual(plain(wizard.readTimetableSetup(url).exportVariables), before);
  assert.ok(sheet('Year 3 Gold'));
});

test('dropdowns on existing tabs are pointed at the longer lists when an activity or class is added', () => {
  const { sheet, update, ss } = setUp();
  update({
    classes: CLASSES.concat([['K Gold', 'Ms Evans', 'Kindergarten', 1]]),
    activities: ['Assembly', 'Library', 'Amadeus - Tutorials', 'Choir']
  });
  // 7 KLAs (Non-Curricular is left out) and 4 activities; 4 classes.
  assert.equal(sheet('K Blue').getRange('J11').getDataValidations()[0][0].list, '=DATA!$D$2:$D$12');
  assert.equal(sheet('Year 3 Gold').getRange('AN40').getDataValidations()[0][0].list, '=DATA!$D$2:$D$12'); // term 4 week B
  assert.equal(sheet('K Blue').getRange('J29').getDataValidations()[0][0], null); // recess: no dropdown
  assert.equal(sheet('SPEC_Ms Jones').getRange('J11').getDataValidations()[0][0].list, '=DATA!$F$2:$F$5');
  assert.ok(ss.getRangeByName('KLAList').getValues().flat().includes('Choir'));
});

test('the Start here tab keeps its "Last run" lines and its list of Google Docs', () => {
  const { ss, update, sheet } = setUp();
  ss.getRangeByName('StartHereSyncStatus').setValue('21 Oct 2026 by someone: 3 synced');
  const docsRow = ss.getRangeByName('StartHereDocs').getRow() + 1;
  sheet('Start here').getRange(docsRow, 1, 1, 4).setValues([['=HYPERLINK("https://docs.google.com/document/d/d1/edit","K Blue - Term 1")', 'K Blue, Term 1', '21 Oct 2026', 'me@school']]);

  update({ classes: CLASSES.concat([['K Gold', 'Ms Evans', 'Kindergarten', 1]]) });

  assert.equal(ss.getRangeByName('StartHereSyncStatus').getValue(), '21 Oct 2026 by someone: 3 synced');
  assert.equal(ss.getRangeByName('StartHereLockStatus').getValue(), 'Not run yet');
  const newDocsRow = ss.getRangeByName('StartHereDocs').getRow() + 1;
  assert.equal(newDocsRow, docsRow + 1); // one more class row above it
  assert.deepEqual(sheet('Start here').getRange(newDocsRow, 1, 1, 2).getValues()[0],
    ['=HYPERLINK("https://docs.google.com/document/d/d1/edit","K Blue - Term 1")', 'K Blue, Term 1']);
  assert.equal(ss.getSheets()[0].getName(), 'Start here');
});

test('two classes can swap names', () => {
  const { sheet, update } = setUp();
  sheet('K Blue').getRange('J11').setValue('English');
  update({ classes: [['K Red', 'Ms Adams', 'Kindergarten', 1], ['K Blue', 'Mr Brown', 'Kindergarten', 1], CLASSES[2]] },
    { classes: [{ from: 'K Blue', to: 'K Red' }, { from: 'K Red', to: 'K Blue' }] });
  assert.equal(sheet('K Red').getRange('J11').getValue(), 'English');
  assert.deepEqual(sheet('K Red').getRange('C27:C28').getValues().flat(), ['Ms Adams', 'K Red']);
  assert.deepEqual(sheet('K Blue').getRange('C27:C28').getValues().flat(), ['Mr Brown', 'K Blue']);
});

test('without the Google Sheets API service, dropdowns are refreshed block by block', () => {
  const world = createWorld();
  makeTemplate(world, TEMPLATE_ID);
  delete world.globals.Sheets;
  const wizard = loadScript('wizard/Code.gs', world.globals);
  const url = wizard.generateTimetable(makePayload());
  world.newExecution();
  wizard.updateTimetable({ jobId: 'j', fileId: url, exportVariables: makePayload({ activities: ['Assembly', 'Library', 'Amadeus - Tutorials', 'Choir'] }).exportVariables });
  world.newExecution();
  const ss = world.files[world.lastCopyId].spreadsheet;
  const rule = ss.getSheetByName('K Blue').getRange('J11').getDataValidations()[0][0];
  assert.equal(rule.range.getA1Notation(), 'D2:D12');
  assert.equal(rule.range.getSheet().getName(), 'DATA');
  assert.equal(ss.getSheetByName('K Blue').getRange('J29').getDataValidations()[0][0], null);
});

test('someone who can only view the timetable is told to ask for edit access', () => {
  const { wizard, url, ss } = setUp();
  ss.getSheetByName('DATA').clearContents = () => { throw new Error("You do not have permission to access the requested document."); };
  assert.throws(() => wizard.updateTimetable({ jobId: 'j', fileId: url, exportVariables: makePayload({ activities: ['Assembly'] }).exportVariables }),
    /You can open this timetable but not change it. Ask its owner to share it with you as an editor/);
});
