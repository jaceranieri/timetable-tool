'use strict';

// The wizard's plain helper functions (Index.html, before the React app),
// run outside the browser.

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makePayload, makeLargeSchoolPayload } = require('./helpers/payload');

const html = fs.readFileSync(path.join(__dirname, '..', 'wizard', 'Index.html'), 'utf8');
const helperSource = html.match(/(const CLASS_COLOR_PALETTE[\s\S]*?)\n\s*const App = /)[1];
const wizard = vm.runInNewContext(helperSource + `;({ makeInitialData, buildExportVariables, setupFromVariables,
  getClassesToGenerate, findCycleMismatches, describeCycleMismatch, estimateGeneration, parsePastedClasses,
  getTeacherError, colourForName, describeUpdate, findRenames, withoutOriginals, asOriginals })`, {});

const plain = value => JSON.parse(JSON.stringify(value));
const valuesOf = variables => Object.fromEntries(variables.map(item => [item.name, item.value]));

test('a setup loaded from a timetable payload (RawData_JSON) gives back the same payload', () => {
  const payload = makePayload();
  const data = wizard.setupFromVariables(payload.exportVariables);
  assert.equal(data.schoolName, 'St Example Primary');
  assert.deepEqual(plain(data.grades.filter(g => g.classes).map(g => [g.name, g.classes, g.cycle])),
    [['Kindergarten', 2, 1], ['Year 3', 1, 2]]);
  assert.ok(data.activities.find(a => a.name === 'Amadeus - Tutorials').locked);
  assert.deepEqual(plain(wizard.buildExportVariables(data)), plain(payload.exportVariables));
});

test('grades the wizard does not list are added; teachers and colours come back as saved', () => {
  const payload = makePayload({ classes: [['Stage 1 Blue', 'TBC', 'Stage 1', 2], ['K Blue', 'Ms Adams', 'Kindergarten', 1]] });
  const data = wizard.setupFromVariables(payload.exportVariables);
  const stage = data.grades.find(g => g.name === 'Stage 1');
  assert.deepEqual([stage.classes, stage.cycle, data.grades.indexOf(stage)], [1, 2, 7]);
  // Kept as saved; Step 3 asks for a real name before generating.
  assert.equal(data.staffing.find(s => s.className === 'Stage 1 Blue').teacher, 'TBC');
  assert.equal(data.staffing.find(s => s.className === 'K Blue').color, '#FFADAD');
});

test('blank class and specialist rows are left out of what is generated', () => {
  const data = wizard.makeInitialData();
  data.schoolName = 'St X';
  data.staffing[1].className = '';             // K Red: fully blank (teacher already blank)
  data.staffing.push({ id: 'g1', gradeId: 'y1', className: 'Year 1 Gold', teacher: 'Ms Gold', color: '#FDFFB6' });
  data.grades[1].classes = 1;
  const values = valuesOf(wizard.buildExportVariables(data));
  assert.equal(values.Class_1_Name, 'Kindergarten Blue');
  assert.equal(values.Class_1_Teacher, '');
  assert.equal(values.Class_2_Name, 'Year 1 Gold');
  assert.equal(values.Class_2_Teacher, 'Ms Gold');
  assert.equal(values.Specialist_1_Name, undefined); // the starting specialist row is blank
});

test('cycle warnings: a class cycle has to be a multiple of the specialist cycle', () => {
  const data = wizard.setupFromVariables(makePayload({
    classes: [['K Blue', 'A', 'Kindergarten', 1], ['Y3 Gold', 'B', 'Year 3', 2], ['Y5 Red', 'C', 'Year 5', 4]],
    specialists: [['Mr Smith', 'PDHPE', 2], ['Ms Jones', 'HSIE', 1]]
  }).exportVariables);
  const mismatches = wizard.findCycleMismatches(data);
  const smith = data.specialists.find(s => s.name === 'Mr Smith');
  assert.deepEqual(Object.keys(mismatches), [String(smith.id)]);
  assert.deepEqual(plain(mismatches[smith.id].map(g => g.name)), ['Kindergarten']);
  assert.match(wizard.describeCycleMismatch(smith, mismatches[smith.id]), /^Mr Smith's 2-week cycle doesn't fit Kindergarten \(1-week cycle\)\./);
});

test('the time estimate matches the load test and small schools', () => {
  const large = wizard.estimateGeneration(wizard.setupFromVariables(makeLargeSchoolPayload().exportVariables));
  assert.deepEqual([large.classTabs, large.specialistTabs], [28, 6]);
  assert.ok(large.seconds > 75 && large.seconds < 120, String(large.seconds));
  assert.equal(large.text, 'about 1½ minutes');
  // Each cycle length needs its own first tab, so it counts more than the number of classes.
  assert.equal(wizard.estimateGeneration(wizard.setupFromVariables(makePayload().exportVariables)).text, 'about 1 minute');
  const oneCycle = makePayload({
    classes: [['K Blue', 'A', 'Kindergarten', 1], ['K Red', 'B', 'Kindergarten', 1]],
    specialists: [['Mr Smith', 'PDHPE', 1]]
  });
  assert.equal(wizard.estimateGeneration(wizard.setupFromVariables(oneCycle.exportVariables)).text, 'under a minute');
});

test('pasted rows: tab-separated name and teacher, heading and blank lines dropped', () => {
  assert.deepEqual(plain(wizard.parsePastedClasses('Class Name\tTeacher\r\n3 Blue\tMs Lee\n\n3 Gold\n 3 Red \t Mr Tran \n')), [
    { className: '3 Blue', teacher: 'Ms Lee' },
    { className: '3 Gold', teacher: '' },
    { className: '3 Red', teacher: 'Mr Tran' }
  ]);
});

test('a teacher is required, and stand-ins like TBC are refused', () => {
  assert.deepEqual(['', '  ', 'TBC', 'tba', 'N/A', 'na', '-', '?', 'Teacher', '=A1', 'Ms Adams', 'Mr T. B. Cole']
    .map(wizard.getTeacherError), ['Required', 'Required', 'Enter their name', 'Enter their name', 'Enter their name',
    'Enter their name', 'Enter their name', 'Enter their name', 'Enter their name', "Can't start with =", '', '']);
});

test('default and pasted classes named after a colour get that colour', () => {
  const defaults = wizard.makeInitialData().staffing.map(s => [s.className, s.color]);
  assert.deepEqual(plain(defaults), [['Kindergarten Blue', '#A0C4FF'], ['Kindergarten Red', '#FFADAD']]);
  assert.equal(wizard.colourForName('Year 1 Gold'), '#FDFFB6');
  assert.equal(wizard.colourForName('Year 2 Jade'), '#C1FFD7');
  assert.equal(wizard.colourForName('3/4 Purple'), '#BDB2FF');
  assert.equal(wizard.colourForName('Year 3 - 1'), null);
});

test('the default setup has two Kindergarten classes and empty grades ready for classes', () => {
  const data = wizard.makeInitialData();
  assert.deepEqual(plain(data.grades.map(g => [g.name, data.staffing.filter(s => s.gradeId === g.id).length, g.classes])), [
    ['Kindergarten', 2, 2], ['Year 1', 0, 0], ['Year 2', 0, 0], ['Year 3', 0, 0], ['Year 4', 0, 0], ['Year 5', 0, 0], ['Year 6', 0, 0]
  ]);
});

test('editing: the changes are listed, renames are found, and a copy continues nothing', () => {
  const original = wizard.setupFromVariables(makePayload().exportVariables);
  assert.ok(original.staffing.every(s => s.originalName === s.className));
  const data = JSON.parse(JSON.stringify(original));
  data.staffing[0].teacher = 'Ms Adams-Lee';                 // K Blue
  data.staffing[1].className = 'K Green';                    // K Red renamed
  data.staffing = data.staffing.filter(s => s.className !== 'Year 3 Gold');
  data.staffing.push({ id: 'n1', gradeId: data.grades[3].id, className: 'Year 3 Silver', teacher: 'Mr Diaz', color: '' });
  data.specialists[0].kla = 'HSIE';                          // Mr Smith
  data.specialists[1].name = 'Ms Jonas';                     // Ms Jones renamed
  data.activities.push({ id: 'a9', name: 'Choir' });
  assert.deepEqual(plain(wizard.describeUpdate(original, data)), [
    'K Blue: teacher Ms Adams to Ms Adams-Lee',
    'Rename class "K Red" to "K Green"',
    'Add class "Year 3 Silver" (Year 3, Mr Diaz)',
    'Archive class "Year 3 Gold"',
    'Mr Smith: KLA PDHPE to HSIE',
    'Rename specialist "Ms Jones" to "Ms Jonas"',
    'Add activity "Choir"'
  ]);
  assert.deepEqual(plain(wizard.findRenames(data)), {
    classes: [{ from: 'K Red', to: 'K Green' }],
    specialists: [{ from: 'Ms Jones', to: 'Ms Jonas' }]
  });
  assert.deepEqual(plain(wizard.describeUpdate(original, original)), []);
  // After an update, the setup is the new starting point.
  assert.deepEqual(plain(wizard.describeUpdate(wizard.asOriginals(data), wizard.asOriginals(data))), []);
  assert.deepEqual(plain(wizard.findRenames(wizard.asOriginals(data))), { classes: [], specialists: [] });
  // Starting a new timetable from one: nothing is a rename.
  const copy = wizard.withoutOriginals(data);
  assert.ok(copy.staffing.every(s => !('originalName' in s)) && copy.specialists.every(s => !('originalName' in s)));
});
