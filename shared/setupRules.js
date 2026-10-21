/*
 * Setup rules shared by the wizard page (wizard/Index.html), which checks
 * them as you type, and the server (wizard/Code.gs, validateConfig), which
 * checks them again before generating: "npm run build" copies this file into
 * both. Edit it here, run the build, then paste both files into Apps Script.
 */
const SLOT_MINUTES = 5;
const MAX_CYCLE_WEEKS = 10;
const MAX_NAME_LENGTH = 90;
const MAX_SCHOOL_NAME_LENGTH = 100;
// Tabs that already exist in the template (or are planned), compared
// case-insensitively because Google Sheets tab names are.
const RESERVED_TAB_NAMES = ['data', 'template', 'overview', 'spec_template', 'start here'];
// Amadeus runs in Years 3-6 in every school. The activity name must match
// TEMPLATE!B21 exactly, because C21 counts cells with this value.
const AMADEUS_ACTIVITY_NAME = 'Amadeus - Tutorials';
// Every class needs its teacher's name for compliance, so stand-ins are refused.
const TEACHER_PLACEHOLDER = /^(tbc|tba|tbd|n\/?a|none|unknown|teacher|[-?.\s]*)$/i;
