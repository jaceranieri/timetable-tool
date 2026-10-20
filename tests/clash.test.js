'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createWorld, makeTemplate, parseA1 } = require('./helpers/fakeSpreadsheet');
const { loadScript } = require('./helpers/loadScript');
const { makePayload } = require('./helpers/payload');
const { evaluateRule } = require('./helpers/formula');

const TEMPLATE_ID = '1x3pzraH_eG2nsQyiZPE9WGAzZ1wZIFGM2JUVF4W6FCk';

function generate(specialists) {
  const world = createWorld();
  makeTemplate(world, TEMPLATE_ID);
  loadScript('wizard/Code.gs', world.globals).generateTimetable(makePayload({ specialists }));
  return world.files[world.lastCopyId].spreadsheet;
}

/** Whether the red clash rule would colour this cell. */
function isRed(ss, sheetName, a1) {
  const sheet = ss.getSheetByName(sheetName);
  const rule = sheet.getConditionalFormatRules().find(r => r._formula && r._formula.includes('ADDRESS(ROW()'));
  const { row, column } = parseA1(a1);
  const covered = rule._ranges.some(range => range.getColumn() === column && row >= range.getRow() &&
    row < range.getRow() + range.getNumRows());
  if (!covered) return false;
  const anchor = rule._ranges[0].getCell(1, 1).getA1Notation();
  return evaluateRule(rule._formula, anchor, ss, sheet, row, column);
}

// Columns (row 11 = 09:00). A 1-week tab: J = Term 1 Monday, P = Term 2 Monday.
// A 2-week tab: J = Term 1 Week A Monday, P = Term 1 Week B Monday, V = Term 2 Week A Monday.

test('same cycle length: same cell on both tabs is a clash', () => {
  const ss = generate([['Mr Smith', 'PDHPE', 1], ['Ms Jones', 'Creative Arts', 1]]);
  ss.getSheetByName('SPEC_Mr Smith').getRange('J11').setValue('K Blue');
  ss.getSheetByName('SPEC_Ms Jones').getRange('J11').setValue('K Blue');
  assert.equal(isRed(ss, 'SPEC_Mr Smith', 'J11'), true);
  assert.equal(isRed(ss, 'SPEC_Ms Jones', 'J11'), true);
  ss.getSheetByName('SPEC_Ms Jones').getRange('J11').setValue('K Red');
  assert.equal(isRed(ss, 'SPEC_Mr Smith', 'J11'), false, 'different classes');
});

test('1-week and 2-week specialists: a weekly lesson clashes with week B (missed before)', () => {
  const ss = generate([['Mr Smith', 'PDHPE', 1], ['Ms Jones', 'Creative Arts', 2]]);
  ss.getSheetByName('SPEC_Mr Smith').getRange('J11').setValue('K Blue'); // every Monday, Term 1
  ss.getSheetByName('SPEC_Ms Jones').getRange('P11').setValue('K Blue'); // Term 1 week B Monday
  assert.equal(isRed(ss, 'SPEC_Mr Smith', 'J11'), true);
  assert.equal(isRed(ss, 'SPEC_Ms Jones', 'P11'), true);
});

test('1-week and 2-week specialists: same column in different terms is not a clash (flagged before)', () => {
  const ss = generate([['Mr Smith', 'PDHPE', 1], ['Ms Jones', 'Creative Arts', 2]]);
  ss.getSheetByName('SPEC_Mr Smith').getRange('P11').setValue('K Blue'); // Term 2 Monday
  ss.getSheetByName('SPEC_Ms Jones').getRange('P11').setValue('K Blue'); // Term 1 week B Monday
  assert.equal(isRed(ss, 'SPEC_Mr Smith', 'P11'), false);
  assert.equal(isRed(ss, 'SPEC_Ms Jones', 'P11'), false);
});

test('a lone lesson, break rows and empty cells are never red', () => {
  const ss = generate([['Mr Smith', 'PDHPE', 1], ['Ms Jones', 'Creative Arts', 2]]);
  ss.getSheetByName('SPEC_Mr Smith').getRange('K11').setValue('K Blue');
  assert.equal(isRed(ss, 'SPEC_Mr Smith', 'K11'), false);
  assert.equal(isRed(ss, 'SPEC_Mr Smith', 'K12'), false);
  // Break rows are identical on every tab; the grey break rule comes first.
  const rules = ss.getSheetByName('SPEC_Ms Jones').getConditionalFormatRules();
  assert.ok(rules[0]._formula.includes('BreakList'));
});

test('matches the real-week definition for every pairing of 1-4 week cycles', () => {
  // Week w of a C-week cycle and week x of a D-week cycle share a real week
  // when w and x are equal modulo gcd(C, D) (cycles restart each term).
  const gcd = (a, b) => (b ? gcd(b, a % b) : a);
  const column = (cycle, term, week, day) => 10 + term * 6 * cycle + week * 6 + day;
  const a1 = (col, row) => require('./helpers/fakeSpreadsheet').columnToLetters(col) + row;

  for (let a = 1; a <= 4; a++) {
    for (let b = 1; b <= 4; b++) {
      if (a === b) continue;
      const ss = generate([['Tab A', 'PDHPE', a], ['Tab B', 'Creative Arts', b]]);
      const tabA = ss.getSheetByName('SPEC_Tab A');
      const tabB = ss.getSheetByName('SPEC_Tab B');
      [[1, 0, 2], [3, a - 1, 4]].forEach(([term, weekA, day]) => {
        const cellA = a1(column(a, term, weekA, day), 11);
        tabA.getRange(cellA).setValue('K Blue');
        for (let weekB = 0; weekB < b; weekB++) {
          const cellB = a1(column(b, term, weekB, day), 11);
          tabB.getRange(cellB).setValue('K Blue');
          const expected = (weekA - weekB) % gcd(a, b) === 0;
          assert.equal(isRed(ss, 'SPEC_Tab A', cellA), expected, `A ${a}-week ${cellA} vs B ${b}-week ${cellB}`);
          assert.equal(isRed(ss, 'SPEC_Tab B', cellB), expected, `B ${b}-week ${cellB} vs A ${a}-week ${cellA}`);
          // Same week block, other terms: never a clash.
          const otherTerm = a1(column(b, (term + 1) % 4, weekB, day), 11);
          tabB.getRange(cellB).setValue('');
          tabB.getRange(otherTerm).setValue('K Blue');
          assert.equal(isRed(ss, 'SPEC_Tab A', cellA), false, `other term ${otherTerm}`);
          tabB.getRange(otherTerm).setValue('');
        }
        tabA.getRange(cellA).setValue('');
      });
    }
  }
});

test('specialist names with quotes still produce a working rule', () => {
  const ss = generate([["Mr O'Brien", 'PDHPE', 1], ['Ms "Jo" Jones', 'Creative Arts', 2]]);
  ss.getSheetByName("SPEC_Mr O'Brien").getRange('J11').setValue('K Blue');
  ss.getSheetByName('SPEC_Ms "Jo" Jones').getRange('J11').setValue('K Blue');
  assert.equal(isRed(ss, "SPEC_Mr O'Brien", 'J11'), true);
});
