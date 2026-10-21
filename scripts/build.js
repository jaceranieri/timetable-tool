'use strict';

/*
 * Copies the files in shared/ into the scripts that use them, between
 * marker lines like these:
 *
 *   // ---- shared/setupRules.js (copied here by "npm run build"; edit that file, not this block) ----
 *   // ---- end of shared/setupRules.js ----
 *
 * so that each Apps Script file stays self-contained (they are pasted into
 * Apps Script one by one) while shared code has a single source.
 *
 *   node scripts/build.js           update the blocks
 *   node scripts/build.js --check   only report blocks that are out of date
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const TARGETS = ['wizard/Code.gs', 'wizard/Index.html', 'template/TemplateSync.gs'];
const START = /^([ \t]*)\/\/ ---- (shared\/[\w.\-]+) \(copied here by "npm run build"; edit that file, not this block\) ----$/;
const END = name => new RegExp('^[ \\t]*// ---- end of ' + name.replace(/[.\/]/g, '\\$&') + ' ----$');

/** A target file with its shared blocks filled in from shared/. */
function build(text, file) {
  const lines = text.split('\n');
  const out = [];
  for (let i = 0; i < lines.length; i++) {
    const start = lines[i].match(START);
    out.push(lines[i]);
    if (!start) continue;
    const indent = start[1];
    const name = start[2];
    const end = lines.findIndex((line, j) => j > i && END(name).test(line));
    if (end === -1) throw new Error(file + ': no end marker for ' + name);
    const source = fs.readFileSync(path.join(ROOT, name), 'utf8').replace(/\s+$/, '');
    source.split('\n').forEach(line => out.push(line ? indent + line : ''));
    out.push(lines[end]);
    i = end;
  }
  return out.join('\n');
}

/** The shared files each target uses. */
function blocksIn(text) {
  return text.split('\n').map(line => line.match(START)).filter(Boolean).map(match => match[2]);
}

function run(check) {
  const stale = [];
  TARGETS.forEach(file => {
    const fullPath = path.join(ROOT, file);
    const text = fs.readFileSync(fullPath, 'utf8');
    const built = build(text, file);
    if (built === text) return;
    stale.push(file);
    if (!check) fs.writeFileSync(fullPath, built);
  });
  return stale;
}

if (require.main === module) {
  const check = process.argv.includes('--check');
  const stale = run(check);
  if (check && stale.length) {
    console.error('Out of date with shared/: ' + stale.join(', ') + '. Run "npm run build".');
    process.exit(1);
  }
  console.log(check ? 'Shared blocks are up to date.' : (stale.length ? 'Updated ' + stale.join(', ') + '.' : 'Nothing to update.'));
}

module.exports = { build, blocksIn, TARGETS, ROOT };
