'use strict';

const { parseA1, columnToLetters } = require('./fakeSpreadsheet');

/**
 * Evaluates a conditional-formatting custom formula for one cell, as Google
 * Sheets would, against the fake spreadsheet. Supports only what the
 * scripts' formulas use: AND, LEN, COUNTIF, INDIRECT, ADDRESS, ROW, COLUMN,
 * INT, MOD, arithmetic, comparisons and & for joining text.
 *
 * referenceCell is the rule's anchor (the top-left cell of its first
 * range); in the formula it stands for the cell being evaluated.
 */
function evaluateRule(formula, referenceCell, ss, sheet, row, column) {
  const js = toJavaScript(formula.replace(/^=/, ''), referenceCell);
  const helpers = {
    SELF: () => sheet.getRange(row, column).getValue(),
    ROW: () => row,
    COLUMN: () => column,
    INT: Math.floor,
    MOD: (a, b) => ((a % b) + b) % b,
    LEN: value => String(value === undefined || value === null ? '' : value).length,
    AND: (...values) => values.every(Boolean),
    ADDRESS: (r, c) => columnToLetters(c) + r,
    INDIRECT: reference => reference,
    COUNTIF: (reference, value) => {
      const match = String(reference).match(/^'((?:[^']|'')*)'!(.+)$/);
      if (!match) throw new Error('Unsupported INDIRECT reference in test: ' + reference);
      const target = ss.getSheetByName(match[1].replace(/''/g, "'"));
      if (!target) return 0;
      const cell = parseA1(match[2]);
      if (cell.row > target.getMaxRows() || cell.column > target.getMaxColumns()) return 0;
      return target.getRange(cell.row, cell.column).getValue() === value ? 1 : 0;
    }
  };
  const names = Object.keys(helpers);
  // eslint-disable-next-line no-new-func
  return Boolean(new Function(...names, 'return (' + js + ');')(...names.map(name => helpers[name])));
}

/** Formula text to a JavaScript expression: strings kept, & joins, anchor -> SELF(). */
function toJavaScript(formula, referenceCell) {
  let out = '';
  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i];
    if (ch === '"') {
      let text = '';
      i++;
      while (i < formula.length) {
        if (formula[i] === '"' && formula[i + 1] === '"') { text += '"'; i += 2; continue; }
        if (formula[i] === '"') break;
        text += formula[i++];
      }
      out += JSON.stringify(text);
    } else if (ch === '&') {
      out += '+';
    } else if (formula.startsWith(referenceCell, i) && !/[A-Z0-9]/.test(formula[i + referenceCell.length] || '') &&
        !/[A-Z0-9]/.test(formula[i - 1] || '')) {
      out += 'SELF()';
      i += referenceCell.length - 1;
    } else {
      out += ch;
    }
  }
  return out;
}

module.exports = { evaluateRule };
