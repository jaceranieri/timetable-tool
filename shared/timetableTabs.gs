/*
 * Shared by the wizard (wizard/Code.gs) and the template's own script
 * (template/TemplateSync.gs): "npm run build" copies this file into both.
 * Edit it here, run the build, then paste both files into Apps Script.
 */

// Notes and text the sync puts on class tabs. The wizard reads them too
// (the Start here tab counts conflicts; an update moves synced lessons when
// a specialist is renamed or archived).
const SYNC_NOTE_PREFIX = '[SPEC_SYNC]';
const CONFLICT_NOTE_PREFIX = '[SPEC_CONFLICT]';
const CONFLICT_VALUE_PREFIX = 'CONFLICT:';

/**
 * Generated specialist tabs are named
 * "SPEC_<name>"; SPEC_Template is the master copy and is never one of them.
 * Compared case-insensitively, like Google Sheets tab names.
 */
function isSpecialistSheet(sheet) {
  const name = sheet.getName().toUpperCase();
  return name.indexOf('SPEC_') === 0 && name !== 'SPEC_TEMPLATE';
}

/**
 * The wizard applies this when generating a timetable, and the Timetable
 * Admin menu re-applies it later.
 *
 * Adds two conditional-formatting rules to every specialist tab: break rows
 * in grey, and a live red clash warning when another specialist has the
 * same class at the same time in the same real week. Specialists can have
 * different cycle lengths, so their columns don't line up; the clash rule
 * works out which cells to compare (see buildClashFormula), the same way
 * "Sync Specialist Schedules" does.
 */
function applySpecialistClashFormatting(ss, allSpecialists) {
  const specialistSheets = ss.getSheets()
    .filter(isSpecialistSheet);

  if (!specialistSheets.length) return;

  const breakListRange = ss.getRangeByName('BreakList');
  // Every specialist tab's name and cycle length (from D3, written by the
  // wizard). The wizard passes the full list instead, because it formats the
  // first tab of each cycle length before the rest are copied from it.
  const cycleOf = sheet => Math.max(1, parseInt(sheet.getRange('D3').getValue(), 10) || 1);
  const specialists = allSpecialists || specialistSheets.map(sheet => ({
    name: sheet.getName(),
    cycle: cycleOf(sheet)
  }));

  specialistSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (lastRow < 5 || lastColumn < 10) return;

    // Day-name columns only (spacer columns have no heading in row 4).
    const dayHeadings = sheet.getRange(4, 10, 1, lastColumn - 9).getDisplayValues()[0];
    const scheduleRanges = [];
    dayHeadings.forEach((heading, index) => {
      if (String(heading).trim()) {
        scheduleRanges.push(sheet.getRange(5, index + 10, lastRow - 4, 1));
      }
    });

    if (!scheduleRanges.length) return;

    const referenceCell = scheduleRanges[0].getCell(1, 1).getA1Notation();

    // Keep any rules that aren't the break/clash rules this function manages
    // (e.g. the class-colour rules added elsewhere).
    const retainedRules = sheet.getConditionalFormatRules().filter(rule => {
      const condition = rule.getBooleanCondition();
      if (!condition) return true;
      const isManagedRule = condition.getCriteriaValues().some(value =>
        typeof value === 'string' &&
        ((value.indexOf('INDIRECT') !== -1 && value.indexOf('COUNTIF') !== -1) ||
          value.indexOf('BreakList') !== -1)
      );
      return !isManagedRule;
    });

    const newRules = [];

    // Break periods (Recess, Lunch, etc, as configured on the DATA tab) are
    // intentionally identical across every specialist at the same time slot,
    // so the clash rule below would otherwise flag them as false-positive
    // collisions. This rule sits above the clash rule so it always wins.
    // INDIRECT because conditional formatting can't refer to other tabs
    // (or named ranges on them) directly.
    if (breakListRange) {
      newRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=COUNTIF(INDIRECT("BreakList"),' + referenceCell + ')>0')
          .setBackground('#D9D9D9')
          .setFontColor('#434343')
          .setRanges(scheduleRanges)
          .build()
      );
    }

    if (specialists.length >= 2) {
      newRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(
            buildClashFormula(referenceCell, cycleOf(sheet), specialists)
          )
          .setBackground('#FF0000')
          .setFontColor('#FFFFFF')
          .setBold(true)
          .setRanges(scheduleRanges)
          .build()
      );
    }

    sheet.setConditionalFormatRules(newRules.concat(retainedRules));
  });
}

/**
 * The clash rule's formula for a specialist tab with the given cycle length.
 * It counts how many specialist tabs (this one included) have the same
 * class at the same time, on the same day, in the same term and in a week
 * that is the same real week, and is true when that count is above 1.
 *
 * Grid layout (from the wizard): columns start at J (10); each week block is
 * 5 day columns plus a spacer, so 6 columns, and a term is 6 x cycle
 * columns wide. For the cell being checked, COLUMN() gives its term, week
 * and day. Cycles restart each term, so week w of a C-week cycle falls on
 * real weeks w, w + C, w + 2C, ...; it shares a real week with week x of a
 * D-week cycle when w and x are equal modulo gcd(C, D). Those are the cells
 * compared on each other tab. With equal cycle lengths that is just the
 * cell in the same position.
 */
function buildClashFormula(referenceCell, cycle, specialists) {
  const gcd = (a, b) => (b ? gcd(b, a % b) : a);
  const offset = '(COLUMN()-10)';
  const term = 'INT(' + offset + '/' + (6 * cycle) + ')';
  const week = 'INT(MOD(' + offset + ',' + (6 * cycle) + ')/6)';
  const day = 'MOD(' + offset + ',6)';

  const counts = [];
  specialists.forEach(other => {
    const sheetName = "'" + other.name.replace(/'/g, "''").replace(/"/g, '""') + "'!";
    if (other.cycle === cycle) {
      counts.push('COUNTIF(INDIRECT("' + sheetName + '"&ADDRESS(ROW(),COLUMN(),4)),' + referenceCell + ')');
      return;
    }
    const shared = gcd(cycle, other.cycle);
    const firstWeek = shared === 1 ? '0' : 'MOD(' + week + ',' + shared + ')';
    for (let k = 0; k < other.cycle / shared; k++) {
      const otherWeek = firstWeek + '+' + (k * shared);
      const column = '10+' + term + '*' + (6 * other.cycle) + '+6*(' + otherWeek + ')+' + day;
      counts.push('COUNTIF(INDIRECT("' + sheetName + '"&ADDRESS(ROW(),' + column + ',4)),' + referenceCell + ')');
    }
  });

  return '=AND(LEN(' + referenceCell + ')>0,(' + counts.join('+') + ')>1)';
}

/** A tab name with the characters Google Sheets doesn't allow replaced. */
function sanitizeSheetName(name) {
  const sanitized = String(name || '')
    .replace(/[\/\\?*\[\]:]/g, '_')
    .trim();

  return sanitized || 'Untitled';
}
