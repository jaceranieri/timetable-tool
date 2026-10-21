'use strict';

// Shared code (PLAN step 34): shared/ is the one source; "npm run build"
// copies it into the scripts that are pasted into Apps Script.

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { build, blocksIn, TARGETS, ROOT } = require('../scripts/build');

const read = file => fs.readFileSync(path.join(ROOT, file), 'utf8');
const definedNames = text => [...text.matchAll(/^\s*(?:function|const) (\w+)\b/gm)].map(m => m[1]);

test('every shared block matches its file in shared/ (run "npm run build" after editing shared/)', () => {
  TARGETS.forEach(file => assert.equal(build(read(file), file), read(file), file + ' is out of date'));
});

test('each script has the shared files it needs, and every shared file is used', () => {
  assert.deepEqual(blocksIn(read('wizard/Code.gs')), ['shared/setupRules.js', 'shared/timetableTabs.gs']);
  assert.deepEqual(blocksIn(read('wizard/Index.html')), ['shared/setupRules.js']);
  assert.deepEqual(blocksIn(read('template/TemplateSync.gs')), ['shared/timetableTabs.gs']);
  const used = new Set(TARGETS.flatMap(file => blocksIn(read(file))));
  fs.readdirSync(path.join(ROOT, 'shared')).forEach(name => assert.ok(used.has('shared/' + name), 'shared/' + name + ' is not used'));
});

test('a shared function or constant is defined once in each script (only in its block)', () => {
  TARGETS.forEach(file => {
    const text = read(file);
    blocksIn(text).forEach(sharedFile => {
      // The shared file's top-level names (not variables inside its functions).
      [...read(sharedFile).matchAll(/^(?:function|const) (\w+)\b/gm)].map(m => m[1]).forEach(name => {
        const count = definedNames(text).filter(defined => defined === name).length;
        assert.equal(count, 1, file + ' defines ' + name + ' ' + count + ' times');
      });
    });
  });
});

test('the build fills a block with the shared file, indented like its markers', () => {
  const stale = [
    'before',
    '    // ---- shared/setupRules.js (copied here by "npm run build"; edit that file, not this block) ----',
    '    const OLD = 1;',
    '    // ---- end of shared/setupRules.js ----',
    'after'
  ].join('\n');
  const built = build(stale, 'test').split('\n');
  assert.equal(built[0], 'before');
  assert.equal(built[built.length - 1], 'after');
  assert.ok(!built.includes('    const OLD = 1;'));
  assert.ok(built.includes('    const SLOT_MINUTES = 5;'));
  assert.throws(() => build('// ---- shared/setupRules.js (copied here by "npm run build"; edit that file, not this block) ----\n', 'test'), /no end marker/);
});
