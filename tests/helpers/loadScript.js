'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..', '..');

/**
 * Runs Apps Script files in a fresh sandbox with the given globals and
 * returns the sandbox, so their top-level functions can be called. Each
 * Apps Script project gets its own sandbox, as it would in Google.
 */
function loadScript(relativePaths, globals) {
  const context = vm.createContext(Object.assign({}, globals));
  [].concat(relativePaths).forEach(relativePath => {
    const file = path.join(ROOT, relativePath);
    vm.runInContext(fs.readFileSync(file, 'utf8'), context, { filename: file });
  });
  return context;
}

/**
 * The plain-JavaScript rule helpers at the top of the wizard's script
 * (everything between SLOT_MINUTES and the App component), which need no
 * React or browser to run.
 */
function loadWizardRules() {
  const html = fs.readFileSync(path.join(ROOT, 'wizard', 'Index.html'), 'utf8');
  const match = html.match(/(const SLOT_MINUTES[\s\S]*?)\n\s*const App = /);
  if (!match) throw new Error('Could not find the wizard rule helpers in Index.html');
  const context = vm.createContext({});
  vm.runInContext(match[1] + '\nthis.rules = { getNameError, findDuplicates, isOnSlotGrid };', context);
  return context.rules;
}

module.exports = { loadScript, loadWizardRules };
