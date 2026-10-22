'use strict';

/**
 * A small in-memory stand-in for the parts of SpreadsheetApp, DriveApp and
 * friends that the scripts use. It is strict where Apps Script is strict, so
 * mistakes show up in tests rather than in a school's timetable:
 *  - ranges outside the sheet throw,
 *  - setValues/setBackgrounds/... with the wrong dimensions throw,
 *  - methods Apps Script doesn't have don't exist here either.
 * Every method call on a spreadsheet, sheet, range or range list is counted
 * (world.calls), as a rough measure of how much work a script does.
 */

const DEFAULTS = { v: '', n: '', bg: '#ffffff', fc: '#000000', fw: 'normal', fmt: '', dv: null };
const LAYERS = Object.keys(DEFAULTS);

function columnToLetters(column) {
  let letters = '';
  while (column > 0) {
    const remainder = (column - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    column = Math.floor((column - 1) / 26);
  }
  return letters;
}

function lettersToColumn(letters) {
  let column = 0;
  for (const ch of letters) column = column * 26 + ch.charCodeAt(0) - 64;
  return column;
}

function parseA1(a1) {
  const match = String(a1).match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
  if (!match) throw new Error('Unsupported A1 notation in fake: ' + a1);
  const row = parseInt(match[2], 10);
  const column = lettersToColumn(match[1]);
  const lastRow = match[4] ? parseInt(match[4], 10) : row;
  const lastColumn = match[3] ? lettersToColumn(match[3]) : column;
  return { row, column, numRows: lastRow - row + 1, numColumns: lastColumn - column + 1 };
}

function displayValue(value) {
  if (value === true) return 'TRUE';
  if (value === false) return 'FALSE';
  if (value === null || value === undefined) return '';
  return String(value);
}

/** Wraps an object so every method call is counted and chaining stays wrapped. */
function counted(world, target) {
  const proxy = new Proxy(target, {
    get(obj, prop) {
      const value = obj[prop];
      if (typeof value !== 'function' || typeof prop !== 'string' || prop.startsWith('_') || prop === 'constructor') {
        return value;
      }
      return (...args) => {
        world.calls++;
        world.callsByMethod[prop] = (world.callsByMethod[prop] || 0) + 1;
        const result = value.apply(obj, args);
        return result === obj ? proxy : result;
      };
    }
  });
  target._proxy = proxy;
  return proxy;
}

class FakeRange {
  constructor(world, sheet, row, column, numRows, numColumns) {
    if (row < 1 || column < 1 || numRows < 1 || numColumns < 1 ||
        row + numRows - 1 > sheet._rows || column + numColumns - 1 > sheet._cols) {
      throw new Error('The coordinates of the range are outside the dimensions of the sheet: ' +
        sheet._name + '!' + columnToLetters(column) + row + ' (' + numRows + 'x' + numColumns + ')');
    }
    Object.assign(this, { _world: world, _sheet: sheet, _row: row, _column: column, _numRows: numRows, _numColumns: numColumns });
    return counted(world, this);
  }

  _get(layer) {
    return Array.from({ length: this._numRows }, (_, r) =>
      Array.from({ length: this._numColumns }, (_, c) => this._sheet._cells[layer][this._row - 1 + r][this._column - 1 + c]));
  }

  _set(layer, grid) {
    if (!Array.isArray(grid) || grid.length !== this._numRows || grid.some(row => row.length !== this._numColumns)) {
      throw new Error('The number of rows or columns in the data does not match the range (' +
        this._numRows + 'x' + this._numColumns + ')');
    }
    grid.forEach((row, r) => row.forEach((value, c) => {
      this._sheet._cells[layer][this._row - 1 + r][this._column - 1 + c] =
        value === null && layer !== 'v' ? DEFAULTS[layer] : value;
    }));
    if (layer === 'v') this._sheet._clearMergedNonAnchors();
    return this;
  }

  _fill(layer, value) {
    return this._set(layer, Array.from({ length: this._numRows }, () => new Array(this._numColumns).fill(value)));
  }

  getRow() { return this._row; }
  getColumn() { return this._column; }
  getNumRows() { return this._numRows; }
  getNumColumns() { return this._numColumns; }
  getSheet() { return this._sheet._proxy; }
  getA1Notation() {
    const start = columnToLetters(this._column) + this._row;
    if (this._numRows === 1 && this._numColumns === 1) return start;
    return start + ':' + columnToLetters(this._column + this._numColumns - 1) + (this._row + this._numRows - 1);
  }
  getCell(row, column) {
    return new FakeRange(this._world, this._sheet, this._row + row - 1, this._column + column - 1, 1, 1);
  }

  getValues() { return this._get('v'); }
  getValue() { return this._get('v')[0][0]; }
  getDisplayValues() { return this._get('v').map(row => row.map(displayValue)); }
  getDisplayValue() { return displayValue(this.getValue()); }
  setValues(grid) { return this._set('v', grid); }
  setValue(value) { return this._fill('v', value); }
  getNotes() { return this._get('n'); }
  setNotes(grid) { return this._set('n', grid); }
  setNote(note) { return this._fill('n', note); }
  getBackgrounds() { return this._get('bg'); }
  setBackgrounds(grid) { return this._set('bg', grid); }
  setBackground(color) { return this._fill('bg', color === null ? DEFAULTS.bg : color); }
  getFontColors() { return this._get('fc'); }
  setFontColors(grid) { return this._set('fc', grid); }
  setFontColor(color) { return this._fill('fc', color === null ? DEFAULTS.fc : color); }
  getFontWeights() { return this._get('fw'); }
  setFontWeights(grid) { return this._set('fw', grid); }
  setFontWeight(weight) { return this._fill('fw', weight === null ? DEFAULTS.fw : weight); }
  setNumberFormat(format) { return this._fill('fmt', format); }
  getNumberFormats() { return this._get('fmt'); }
  setFontSize() { return this; }
  setFontFamily() { return this; }
  setHorizontalAlignment() { return this; }
  setVerticalAlignment() { return this; }
  setWrap() { return this; }
  setBorder() { return this; }
  getFormulas() {
    return this._get('v').map(row => row.map(value => (typeof value === 'string' && value.charAt(0) === '=' ? value : '')));
  }
  clearContent() { return this._fill('v', ''); }
  clearDataValidations() { return this._fill('dv', null); }
  setDataValidation(rule) { return this._fill('dv', rule); }
  getDataValidations() { return this._get('dv'); }

  breakApart() {
    this._sheet._merges = this._sheet._merges.filter(merge => !this._overlaps(merge));
    return this;
  }
  merge() {
    if (this._sheet._merges.some(merge => this._overlaps(merge))) {
      throw new Error('You must select all cells in a merged range to merge or unmerge them.');
    }
    this._sheet._merges.push(this._box());
    this._sheet._clearMergedNonAnchors();
    return this;
  }
  _box() {
    return { row: this._row, column: this._column, numRows: this._numRows, numColumns: this._numColumns };
  }
  /** Whether this range overlaps a {row, column, numRows, numColumns} box. */
  _overlaps(other) {
    return this._row <= other.row + other.numRows - 1 && other.row <= this._row + this._numRows - 1 &&
      this._column <= other.column + other.numColumns - 1 && other.column <= this._column + this._numColumns - 1;
  }

  copyFormatToRange(sheet, column, columnEnd, row, rowEnd) {
    sheet.getRange(row, column, rowEnd - row + 1, columnEnd - column + 1); // bounds check
    return undefined;
  }

  copyTo(destination, copyPasteType) {
    if (copyPasteType !== 'PASTE_CONDITIONAL_FORMATTING') {
      throw new Error('Fake only supports PASTE_CONDITIONAL_FORMATTING');
    }
    const sourceRules = this._sheet._cfRules.filter(rule => rule._ranges.some(range => range._overlaps(this._box())));
    sourceRules.forEach(rule => {
      destination._sheet._cfRules.push(cloneRule(rule, [destination], { pastedFrom: this.getA1Notation() }));
    });
    return undefined;
  }

  protect() {
    const sheet = this._sheet;
    const range = this._proxy;
    const protection = {
      _description: '',
      setDescription(text) { this._description = text; return this; },
      getDescription() { return this._description; },
      addEditor() { return this; },
      removeEditors() { return this; },
      setDomainEdit() { return this; },
      canEdit() { return true; },
      getRange() { return range; },
      remove() { sheet._protections.splice(sheet._protections.indexOf(this), 1); }
    };
    sheet._protections.push(protection);
    return protection;
  }
}

class FakeRangeList {
  constructor(world, sheet, a1List) {
    this._ranges = a1List.map(a1 => sheet.getRange(a1));
    return counted(world, this);
  }
  _each(method, args) { this._ranges.forEach(range => range[method](...args)); return this; }
  getRanges() { return this._ranges.slice(); }
  setValue(...args) { return this._each('setValue', args); }
  setBackground(...args) { return this._each('setBackground', args); }
  setFontWeight(...args) { return this._each('setFontWeight', args); }
  setFontColor(...args) { return this._each('setFontColor', args); }
  setWrap(...args) { return this._each('setWrap', args); }
  setFontSize(...args) { return this._each('setFontSize', args); }
  setFontFamily(...args) { return this._each('setFontFamily', args); }
  setNumberFormat(...args) { return this._each('setNumberFormat', args); }
  setHorizontalAlignment(...args) { return this._each('setHorizontalAlignment', args); }
  setVerticalAlignment(...args) { return this._each('setVerticalAlignment', args); }
  setBorder(...args) { return this._each('setBorder', args); }
  clearDataValidations() { return this._each('clearDataValidations', []); }
  clearContent() { return this._each('clearContent', []); }
  breakApart() { return this._each('breakApart', []); }
}

/**
 * A conditional-format rule. formula is a custom formula, or null for other
 * kinds; extra can give criteriaType and values (e.g. TEXT_EQUAL_TO,
 * ['English']), background and fontColor.
 */
function makeRule(formula, ranges, extra) {
  const rule = Object.assign({
    _formula: formula,
    _ranges: ranges,
    getRanges() { return this._ranges; },
    getBooleanCondition() {
      const type = rule.criteriaType || (formula === null ? null : 'CUSTOM_FORMULA');
      if (!type) return null;
      return {
        getCriteriaType: () => type,
        getCriteriaValues: () => rule.values || [formula],
        getBackground: () => rule.background || null,
        getFontColor: () => rule.fontColor || null
      };
    }
  }, extra || {});
  return rule;
}

/** The same rule on other ranges (pasting, copying a sheet). */
function cloneRule(rule, ranges, extra) {
  return makeRule(rule._formula, ranges, Object.assign({
    criteriaType: rule.criteriaType, values: rule.values, background: rule.background,
    fontColor: rule.fontColor, pastedFrom: rule.pastedFrom
  }, extra || {}));
}

class FakeSheet {
  constructor(world, spreadsheet, name, rows, cols) {
    Object.assign(this, {
      _world: world, _spreadsheet: spreadsheet, _name: name, _rows: rows, _cols: cols,
      _hidden: false, _merges: [], _widths: {}, _cfRules: [], _protections: [],
      _id: (world.nextSheetId = (world.nextSheetId || 0) + 1)
    });
    this._cells = {};
    LAYERS.forEach(layer => {
      this._cells[layer] = Array.from({ length: rows }, () => new Array(cols).fill(DEFAULTS[layer]));
    });
    return counted(world, this);
  }

  /** Like Google Sheets, a merged range only holds a value in its top-left cell. */
  _clearMergedNonAnchors() {
    this._merges.forEach(merge => {
      for (let r = 0; r < merge.numRows; r++) {
        for (let c = 0; c < merge.numColumns; c++) {
          if (r || c) this._cells.v[merge.row - 1 + r][merge.column - 1 + c] = '';
        }
      }
    });
  }

  getName() { return this._name; }
  /** The selected cell: A1 unless a test picks another with setActiveSelection. */
  getActiveCell() { return this.getRange(this._activeCell || 'A1'); }
  setActiveSelection(a1) { this._activeCell = a1; return this.getRange(a1); }
  getSheetId() { return this._id; }
  setName(name) {
    if (this._spreadsheet._sheets.some(sheet => sheet !== this._proxy && sheet._name.toLowerCase() === String(name).toLowerCase())) {
      throw new Error('A sheet with the name "' + name + '" already exists.');
    }
    this._name = name;
    return this;
  }
  showSheet() { this._hidden = false; return this; }
  hideSheet() { this._hidden = true; return this; }
  setHiddenGridlines(hidden) { this._hiddenGridlines = hidden; return this; }
  setTabColor(color) { this._tabColor = color; return this; }
  setFrozenRows(rows) { this._frozenRows = rows; }
  isSheetHidden() { return this._hidden; }
  getMaxRows() { return this._rows; }
  getMaxColumns() { return this._cols; }
  insertRowsAfter(after, count) {
    LAYERS.forEach(layer => {
      for (let i = 0; i < count; i++) this._cells[layer].splice(after, 0, new Array(this._cols).fill(DEFAULTS[layer]));
    });
    this._rows += count;
    return this;
  }
  insertColumnsAfter(after, count) {
    LAYERS.forEach(layer => this._cells[layer].forEach(row => row.splice(after, 0, ...new Array(count).fill(DEFAULTS[layer]))));
    this._cols += count;
    return this;
  }
  getLastRow() {
    let last = 0;
    this._cells.v.forEach((row, index) => { if (row.some(value => value !== '')) last = index + 1; });
    return last;
  }
  getLastColumn() {
    let last = 0;
    this._cells.v.forEach(row => row.forEach((value, index) => { if (value !== '') last = Math.max(last, index + 1); }));
    return last;
  }
  getRange(a, b, c, d) {
    if (typeof a === 'string') {
      const parsed = parseA1(a);
      return new FakeRange(this._world, this, parsed.row, parsed.column, parsed.numRows, parsed.numColumns);
    }
    return new FakeRange(this._world, this, a, b, c === undefined ? 1 : c, d === undefined ? 1 : d);
  }
  getRangeList(a1List) { return new FakeRangeList(this._world, this, a1List); }
  setColumnWidth(column, width) { this._widths[column] = width; return this; }
  setRowHeights() { return this; }
  setColumnWidths(column, count, width) {
    for (let i = 0; i < count; i++) this._widths[column + i] = width;
    return this;
  }
  getColumnWidth(column) { return this._widths[column] || 100; }
  clearContents() { this._cells.v = this._cells.v.map(row => row.map(() => '')); return this; }
  clearConditionalFormatRules() { this._cfRules = []; }
  getConditionalFormatRules() { return this._cfRules.slice(); }
  setConditionalFormatRules(rules) { this._cfRules = rules.slice(); }
  getProtections() { return this._protections.slice(); }

  copyTo(spreadsheet) {
    const target = spreadsheet._spreadsheetTarget;
    const copy = new FakeSheet(this._world, target, 'Copy of ' + this._name, this._rows, this._cols);
    LAYERS.forEach(layer => { copy._cells[layer] = this._cells[layer].map(row => row.slice()); });
    copy._merges = this._merges.map(merge => Object.assign({}, merge));
    copy._widths = Object.assign({}, this._widths);
    copy._hidden = this._hidden;
    copy._cfRules = this._cfRules.map(rule => cloneRule(rule,
      rule._ranges.map(range => copy.getRange(range._row, range._column, range._numRows, range._numColumns))));
    target._sheets.push(copy);
    return copy;
  }

  // Test helpers (not Apps Script)
  _value(a1) { return this.getRange(a1).getValue(); }
}

class FakeSpreadsheet {
  constructor(world, id) {
    Object.assign(this, { _world: world, _id: id, _sheets: [], _named: {}, _owner: 'owner@school.example' });
    this._spreadsheetTarget = this;
    return counted(world, this);
  }
  _addSheet(name, rows, cols) {
    const sheet = new FakeSheet(this._world, this, name, rows, cols);
    this._sheets.push(sheet);
    return sheet;
  }
  getId() { return this._id; }
  getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this._id + '/edit'; }
  getName() { return this._world.files[this._id].handle.getName(); }
  rename(name) { this._world.files[this._id].handle.setName(name); }
  // Tabs added through the Sheets API aren't visible to SpreadsheetApp until
  // the next execution (world.newExecution()). That's the cautious
  // assumption the scripts must cope with.
  _visibleSheets() { return this._sheets.filter(sheet => !sheet._addedByApi); }
  getSheets() { return this._visibleSheets(); }
  getSheetByName(name) {
    return this._visibleSheets().find(sheet => sheet._name.toLowerCase() === String(name).toLowerCase()) || null;
  }
  setNamedRange(name, range) { this._named[name] = range; }
  getRangeByName(name) { return this._named[name] || null; }
  getSpreadsheetTimeZone() { return 'Australia/Sydney'; }
  insertSheet(name, index) {
    if (this.getSheetByName(name)) throw new Error('A sheet with the name "' + name + '" already exists.');
    const sheet = new FakeSheet(this._world, this, name, 1000, 26);
    this._sheets.splice(index === undefined ? this._sheets.length : index, 0, sheet);
    return sheet;
  }
  setActiveSheet(sheet) { this._active = sheet; return sheet; }
  getActiveSheet() { return this._active || this._sheets[0]; }
  deleteSheet(sheet) { this._sheets.splice(this._sheets.indexOf(sheet), 1); }
  moveActiveSheet(position) {
    const index = this._sheets.indexOf(this._active);
    this._sheets.splice(index, 1);
    this._sheets.splice(position - 1, 0, this._active);
  }

  /** Sheets API batchUpdate: only the request types the scripts send. */
  _batchUpdate(requests) {
    requests.forEach(request => {
      if (request.duplicateSheet) {
        const r = request.duplicateSheet;
        const source = this._sheets.find(sheet => sheet._id === r.sourceSheetId);
        if (!source) throw new Error('No sheet with id ' + r.sourceSheetId);
        if (this._sheets.some(sheet => sheet._id === r.newSheetId)) throw new Error('Sheet id in use: ' + r.newSheetId);
        if (this._sheets.some(sheet => sheet._name.toLowerCase() === r.newSheetName.toLowerCase())) {
          throw new Error('A sheet with the name "' + r.newSheetName + '" already exists.');
        }
        if (r.insertSheetIndex > this._sheets.length) throw new Error('insertSheetIndex out of range: ' + r.insertSheetIndex);
        const copy = source.copyTo(this);
        this._sheets.splice(this._sheets.indexOf(copy), 1);
        this._sheets.splice(r.insertSheetIndex, 0, copy);
        copy._id = r.newSheetId;
        copy._name = r.newSheetName;
        copy._addedByApi = true;
      } else if (request.updateCells) {
        const r = request.updateCells;
        if (r.fields !== 'userEnteredValue') throw new Error('Fake only supports fields: userEnteredValue');
        const sheet = this._sheets.find(item => item._id === r.start.sheetId);
        if (!sheet) throw new Error('No sheet with id ' + r.start.sheetId);
        const values = r.rows.map(row => row.values.map(cell => {
          const v = cell.userEnteredValue;
          if ('boolValue' in v) return v.boolValue;
          if ('numberValue' in v) return v.numberValue;
          return v.stringValue;
        }));
        sheet.getRange(r.start.rowIndex + 1, r.start.columnIndex + 1, values.length, values[0].length).setValues(values);
      } else if (request.setDataValidation) {
        const r = request.setDataValidation;
        const sheet = this._sheets.find(item => item._id === r.range.sheetId);
        if (!sheet) throw new Error('No sheet with id ' + r.range.sheetId);
        const range = sheet.getRange(r.range.startRowIndex + 1, r.range.startColumnIndex + 1,
          r.range.endRowIndex - r.range.startRowIndex, r.range.endColumnIndex - r.range.startColumnIndex);
        if (r.rule) {
          if (r.rule.condition.type !== 'ONE_OF_RANGE') throw new Error('Fake only supports ONE_OF_RANGE validation');
          range.setDataValidation({ kind: 'validation', list: r.rule.condition.values[0].userEnteredValue, allowInvalid: !r.rule.strict });
        } else {
          range.clearDataValidations();
        }
      } else {
        throw new Error('Fake does not support request: ' + Object.keys(request)[0]);
      }
    });
    return { replies: [] };
  }
  toast() {}
  getOwner() { return { getEmail: () => this._owner }; }

  _clone(world, id) {
    const copy = new FakeSpreadsheet(world, id);
    this._sheets.forEach(sheet => {
      const sheetCopy = sheet.copyTo(copy);
      sheetCopy._name = sheet._name;
    });
    return copy;
  }
}

/**
 * The globals Apps Script provides, wired to fake spreadsheets. The master
 * template is created by the caller with makeTemplate().
 */
function createWorld() {
  const world = { calls: 0, callsByMethod: {}, files: {}, cache: {}, alerts: [], properties: {}, userProperties: {} };
  world.resetCalls = () => { world.calls = 0; world.callsByMethod = {}; };
  // A later script run (e.g. a menu item): every tab is visible again.
  world.newExecution = () => {
    Object.values(world.files).forEach(file => file.spreadsheet._sheets.forEach(sheet => { sheet._addedByApi = false; }));
  };

  const builder = (kind) => {
    const state = { kind };
    const api = {
      requireValueInRange(range) { state.range = range; return api; },
      setAllowInvalid(allow) { state.allowInvalid = allow; return api; },
      whenFormulaSatisfied(formula) { state.formula = formula; return api; },
      setBackground(color) { state.background = color; return api; },
      setFontColor(color) { state.fontColor = color; return api; },
      setBold(bold) { state.bold = bold; return api; },
      setRanges(ranges) { state.ranges = ranges; return api; },
      build() {
        if (kind === 'validation') return Object.assign({}, state);
        return makeRule(state.formula, state.ranges || [], { background: state.background, fontColor: state.fontColor });
      }
    };
    return api;
  };

  world.globals = {
    console: { log() {}, warn() {}, error() {} },
    SpreadsheetApp: {
      openById: id => {
        if (!world.files[id]) throw new Error('No item with the given ID could be found.');
        return world.files[id].spreadsheet;
      },
      getActiveSpreadsheet: () => world.active,
      flush: () => { world.flushes = (world.flushes || 0) + 1; },
      getUi: () => {
        const menu = { items: [], addItem(label, fn) { this.items.push([label, fn]); return this; },
          addSeparator() { this.items.push(['-']); return this; }, addToUi() { world.menu = this.items; return this; } };
        return {
          // Answers come from world.answers (a queue), defaulting to YES.
          alert: (...args) => {
            world.alerts.push(args.length > 1 ? args[1] : args[0]);
            return world.answers && world.answers.length ? world.answers.shift() : 'YES';
          },
          prompt: (...args) => {
            world.prompts = (world.prompts || []).concat([args[1]]);
            const text = world.promptAnswers && world.promptAnswers.length ? world.promptAnswers.shift() : '';
            return { getSelectedButton: () => (text === null ? 'CANCEL' : 'OK'), getResponseText: () => text || '' };
          },
          showModalDialog: (output, title) => { world.dialogs = (world.dialogs || []).concat([{ title, html: output._html }]); },
          createMenu: () => menu,
          ButtonSet: { YES_NO: 'YES_NO', YES_NO_CANCEL: 'YES_NO_CANCEL', OK_CANCEL: 'OK_CANCEL' },
          Button: { YES: 'YES', NO: 'NO', CANCEL: 'CANCEL', OK: 'OK' }
        };
      },
      newDataValidation: () => builder('validation'),
      newConditionalFormatRule: () => builder('rule'),
      CopyPasteType: { PASTE_CONDITIONAL_FORMATTING: 'PASTE_CONDITIONAL_FORMATTING' },
      ProtectionType: { RANGE: 'RANGE' },
      BorderStyle: { SOLID: 'SOLID' }
    },
    DriveApp: {
      getRootFolder: () => world.root,
      getFileById: id => {
        const file = world.files[id];
        if (!file) {
          throw new Error('No item with the given ID could be found, or you do not have permission to access it.');
        }
        return file.handle;
      }
    },
    CacheService: {
      getUserCache: () => ({
        put: (key, value) => { world.cache[key] = value; },
        get: key => world.cache[key] || null
      })
    },
    PropertiesService: {
      getScriptProperties: () => ({ getProperty: key => world.properties[key] || null }),
      // Per-user storage; each value holds at most 9 KB, as in Apps Script.
      getUserProperties: () => ({
        getProperty: key => (key in world.userProperties ? world.userProperties[key] : null),
        setProperties: values => {
          Object.keys(values).forEach(key => {
            if (String(values[key]).length > 9000) throw new Error('Argument too large: value');
            world.userProperties[key] = String(values[key]);
          });
        },
        deleteProperty: key => { delete world.userProperties[key]; }
      })
    },
    Session: {
      getEffectiveUser: () => ({ getEmail: () => 'owner@school.example' }),
      getActiveUser: () => ({ getEmail: () => 'coordinator@school.example' })
    },
    // The Google Sheets API advanced service. Tests delete it to check the
    // one-at-a-time fallback.
    Sheets: {
      Spreadsheets: {
        batchUpdate: (body, spreadsheetId) => {
          world.calls++;
          world.batchUpdates = (world.batchUpdates || 0) + 1;
          return world.files[spreadsheetId].spreadsheet._batchUpdate(body.requests);
        }
      }
    },
    HtmlService: {
      createHtmlOutput: html => {
        const output = { _html: html, setWidth() { return output; }, setHeight() { return output; } };
        return output;
      }
    },
    // The Drive API advanced service (v3), for making Google Docs from HTML.
    Drive: {
      Files: {
        get: (id) => ({ parents: [world.files[id] && world.files[id].folder ? world.files[id].folder : 'My Drive'] }),
        create: (resource, blob) => {
          const id = 'doc' + ((world.docs || []).length + 1);
          world.docs = (world.docs || []).concat([{ id, name: resource.name, mimeType: resource.mimeType,
            parents: resource.parents, html: blob.getDataAsString(), page: {} }]);
          return { id };
        }
      }
    },
    DocumentApp: {
      ElementType: { PARAGRAPH: 'PARAGRAPH', TABLE: 'TABLE', PAGE_BREAK: 'PAGE_BREAK' },
      ParagraphHeading: { NORMAL: 'NORMAL', HEADING1: 'HEADING1', HEADING2: 'HEADING2' },
      openById: id => {
        const doc = world.docs.find(item => item.id === id);
        // The converted Doc's top-level elements, in order: headings and
        // paragraphs, and tables (each counted once, nested or not).
        if (!doc.children) {
          doc.children = [];
          const tokens = doc.html.match(/<(h1|h2|p|table)[\s>]|<\/table>/g) || [];
          let depth = 0;
          let position = 0;
          tokens.forEach(token => {
            const at = doc.html.indexOf(token, position);
            position = at + token.length;
            if (token === '</table>') { depth--; return; }
            const tag = token.slice(1, -1).trim();
            if (tag === 'table') {
              // Columns: the cells in the table's first row.
              const firstRow = doc.html.slice(at).match(/<tr[^>]*>([\s\S]*?)<\/tr>/);
              const columns = firstRow ? (firstRow[1].match(/<t[hd][\s>]/g) || []).length : 0;
              if (depth++ === 0) doc.children.push({ type: 'TABLE', columns: columns, widths: [] });
            } else if (depth === 0) {
              doc.children.push({ type: 'PARAGRAPH', heading: tag === 'h1' ? 'HEADING1' : tag === 'h2' ? 'HEADING2' : 'NORMAL' });
            }
          });
        }
        const element = item => ({
          getType: () => item.type,
          asParagraph: () => ({ getHeading: () => item.heading }),
          asTable: () => ({
            getNumRows: () => 1,
            getRow: () => ({ getNumCells: () => item.columns }),
            setColumnWidth: (column, width) => { item.widths[column] = width; }
          })
        });
        const body = {
          getNumChildren: () => doc.children.length,
          getChild: index => element(doc.children[index]),
          insertPageBreak: index => { doc.children.splice(index, 0, { type: 'PAGE_BREAK' }); }
        };
        ['PageWidth', 'PageHeight', 'MarginTop', 'MarginBottom', 'MarginLeft', 'MarginRight'].forEach(name => {
          body['set' + name] = value => { doc.page[name] = value; return body; };
        });
        return { getBody: () => body, saveAndClose: () => { doc.saved = true; } };
      }
    },
    Utilities: {
      newBlob: (data, contentType, name) => ({ getDataAsString: () => data, getContentType: () => contentType, getName: () => name }),
      // Times of day are real; full dates are fixed, so tests can check text that includes them.
      formatDate: (date, timeZone, pattern) => (pattern && pattern.indexOf('yyyy') !== -1
        ? '21 October 2026, 2:05 PM'
        : String(date.getHours()).padStart(2, '0') + ':' + String(date.getMinutes()).padStart(2, '0'))
    }
  };

  // My Drive: just enough folder handling for getOutputFolder().
  const makeFolder = name => {
    const folder = {
      _children: [],
      getName: () => name,
      getFoldersByName: childName => {
        const matches = folder._children.filter(child => child.getName() === childName);
        return { hasNext: () => matches.length > 0, next: () => matches.shift() };
      },
      createFolder: childName => {
        const child = makeFolder(childName);
        folder._children.push(child);
        world.foldersCreated = (world.foldersCreated || 0) + 1;
        return child;
      }
    };
    return folder;
  };
  world.root = makeFolder('My Drive');

  world.addFile = (id, spreadsheet) => {
    let name = 'Template';
    const file = {
      spreadsheet,
      handle: {
        getId: () => id,
        getName: () => name,
        setName: newName => { name = newName; return file.handle; },
        makeCopy: (newName, folder) => {
          const copyId = 'copy' + (Object.keys(world.files).length + 1);
          const copy = spreadsheet._clone(world, copyId);
          world.addFile(copyId, copy).handle.setName(newName);
          world.files[copyId].folder = folder ? folder.getName() : 'My Drive';
          world.lastCopyId = copyId;
          return world.files[copyId].handle;
        }
      }
    };
    world.files[id] = file;
    return file;
  };

  return world;
}

/**
 * A copy of the master template's layout, as far as the scripts care:
 * DATA, OVERVIEW, TEMPLATE (75 x 15, the first week only) and SPEC_Template.
 */
function makeTemplate(world, id) {
  const ss = new FakeSpreadsheet(world, id);
  const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];

  ss._addSheet('DATA', 1000, 17).getRange('A1:B1').setValues([['VariableName', 'Value']]);
  ss._addSheet('OVERVIEW', 19, 86).getRange('G1').setValue('School Compliance Overview');

  const template = ss._addSheet('TEMPLATE', 75, 15);
  template.getRange('J1').setValue('Week A - Term 1');
  template.getRange('I4:N4').setValues([['Time'].concat(days)]);
  template.getRange('B5:B12').setValues(['English', 'Mathematics', 'Religion', 'Science & Tech', 'HSIE', 'PDHPE', 'Creative Arts', 'Non-Curricular'].map(v => [v]));
  template.getRange('B21').setValue('Amadeus - Tutorials');
  template.getRange('B27:C30').setValues([['Teacher: ', 'Jason Ranieri'], ['Class Name', 'KGREEN'], ['Grade', 'Kindergarten'], ['Weeks per Cycle', '3 weeks per term']]);
  [
    ['English', '#D9EAD3', '#6AA84F'], ['Mathematics', '#FCE5CD', '#E69138'], ['Religion', '#FFF2CC', '#F1C232'],
    ['Science & Tech', '#C9DAF8', '#3D85C6'], ['HSIE', '#D9D2E9', '#674EA7'], ['PDHPE', '#D0E0E3', '#45818E'],
    ['Creative Arts', '#F4CCCC', '#CC0000']
  ].forEach(([kla, background, fontColor]) => {
    template._cfRules.push(makeRule(null, [template.getRange('A5:B12')],
      { criteriaType: 'TEXT_EQUAL_TO', values: [kla], background, fontColor }));
  });
  template._cfRules.push(makeRule('=LEN(TRIM(B5))>0', [template.getRange('B5')], { background: '#F3F3F3', fontColor: '#B7B7B7' }));

  const spec = ss._addSheet('SPEC_Template', 75, 15);
  spec.hideSheet();
  spec.getRange('B1').setValue('<TEACHER NAME>');
  spec.getRange('B4:D4').setValues([['Class', 'Mins', 'Hrs']]);
  spec.getRange('I4:N4').setValues([['Time'].concat(days)]);
  spec._cfRules.push(makeRule('=LEN(TRIM(J5))>0', [spec.getRange('J5:N75')]));

  world.addFile(id, ss);
  return ss;
}

module.exports = { createWorld, makeTemplate, parseA1, columnToLetters };
