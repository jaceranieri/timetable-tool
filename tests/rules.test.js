'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createWorld } = require('./helpers/fakeSpreadsheet');
const { loadScript, loadWizardRules } = require('./helpers/loadScript');
const { makePayload } = require('./helpers/payload');

const server = loadScript('wizard/Code.gs', createWorld().globals);
const wizard = loadWizardRules();

function problems(options) {
  try {
    server.validateConfig(server.parsePayloadConfig(makePayload(options)));
    return [];
  } catch (error) {
    return error.message.split('\n- ').slice(1);
  }
}

test('the default setup is valid', () => {
  assert.deepEqual(problems({}), []);
});

test('the wizard and the server agree on which names are allowed', () => {
  const names = [
    'K Blue', 'Year 3 Gold', 'Year 3/4', "Mrs O'Neill", 'Data', 'data', 'OVERVIEW', 'Template',
    'SPEC_Template', 'spec_x', 'Start here', '=SUM(1)', 'a:b', 'x*y', '[1]', 'a\\b', 'Q?', 'x'.repeat(91)
  ];
  [true, false].forEach(isClass => {
    names.forEach(name => {
      const serverOk = server.getNameProblem(name, isClass) === '';
      const wizardOk = wizard.getNameError(name, isClass) === '';
      assert.equal(wizardOk, serverOk, `${isClass ? 'class' : 'specialist'} "${name}"`);
    });
  });
  assert.equal(server.getNameProblem("Mr O'Brien", false), '', 'apostrophes are fine for specialists');
});

test('the school name is required, at most 100 characters, and not a formula', () => {
  assert.deepEqual(problems({ schoolName: '' }), ['Enter the school name.']);
  assert.deepEqual(problems({ schoolName: '=x' }), ["The school name can't start with =."]);
  assert.deepEqual(problems({ schoolName: 'x'.repeat(101) }), ['The school name is longer than 100 characters.']);
  assert.deepEqual(problems({ schoolName: "St Mary's Primary" }), []);
});

test('times must be on the 5-minute grid, inside the day and not overlapping', () => {
  assert.deepEqual(problems({ start: '08:32' }), ['School start and end times must be on a 5-minute mark (e.g. 8:30, 8:35).']);
  assert.deepEqual(problems({ start: '15:00', end: '09:00', breaks: [] }), ['School end time must be after the start time.']);
  assert.deepEqual(problems({ breaks: [['Recess', '10:30', '11:00'], ['Lunch', '10:45', '11:30']] }),
    ['"Recess" and "Lunch" overlap.']);
  assert.deepEqual(problems({ breaks: [['Late', '15:00', '15:30']] }), ['"Late" must be within the school day.']);
  assert.deepEqual(problems({ breaks: [['', '12:00', '12:10']] }), ['Break 1 needs a name.']);
  assert.equal(wizard.isOnSlotGrid('08:35'), true);
  assert.equal(wizard.isOnSlotGrid('08:32'), false);
});

test('class rules: unique, not a break name, cycle 1-10, at least one class', () => {
  assert.deepEqual(problems({ classes: [['K Blue', 'A', 'Kindergarten', 1], ['k blue', 'B', 'Kindergarten', 1]] }),
    ['Class "k blue" is used more than once.']);
  assert.deepEqual(problems({ classes: [['Recess', 'A', 'Year 1', 1]] }), ['Class "Recess" has the same name as a break.']);
  assert.deepEqual(problems({ classes: [['K Blue', 'A', 'Kindergarten', 11]] }),
    ['Class "K Blue" needs a cycle length of 1 to 10 weeks.']);
  assert.deepEqual(problems({ classes: [] }), ['Add at least one class.']);
  assert.deepEqual(problems({ classes: [['K Blue', '=x', 'Kindergarten', 1]] }),
    ['Class "K Blue": the teacher name can\'t start with =.']);
});

test('specialist rules: unique, KLA from the list, cycle 1-10', () => {
  assert.deepEqual(problems({ specialists: [['Mr Smith', 'PDHPE', 1], ['mr smith', 'PDHPE', 1]] }),
    ['Specialist "mr smith" is used more than once.']);
  assert.deepEqual(problems({ specialists: [['Mr Smith', 'Assembly', 1]] }), ['Specialist "Mr Smith" needs a KLA from the list.']);
  assert.deepEqual(problems({ specialists: [['Mr Smith', 'PDHPE', 0]] }),
    ['Specialist "Mr Smith" needs a cycle length of 1 to 10 weeks.']);
  assert.deepEqual(problems({ specialists: [['Template', 'PDHPE', 1]] }),
    ['Specialist "Template" is a name the timetable uses for its own tabs.']);
});

test('Amadeus is always added and only applies to Years 3-6', () => {
  const config = server.parsePayloadConfig(makePayload({ activities: ['Assembly'] }));
  assert.ok(config.activities.some(a => a.name === 'Amadeus - Tutorials'));
  assert.deepEqual(['Kindergarten', 'Year 2', 'Year 3', 'Year 6', 'Year 7'].map(g => server.isAmadeusGrade(g)),
    [false, false, true, true, false]);
});

test('wizard findDuplicates ignores case and spaces', () => {
  assert.deepEqual([...wizard.findDuplicates(['K Blue', 'k blue ', 'Year 1'])], ['k blue']);
});
