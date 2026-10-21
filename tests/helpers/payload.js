'use strict';

const KLAS = ['English', 'Mathematics', 'Religion', 'Science & Tech', 'HSIE', 'PDHPE', 'Creative Arts', 'Non-Curricular'];

/**
 * A wizard payload, in the same flat name/value form Index.html sends.
 * Anything not given uses a small, valid default school.
 */
function makePayload(options) {
  const o = options || {};
  const variables = [];
  const add = (name, value) => variables.push({ name, value });

  add('SchoolName', o.schoolName !== undefined ? o.schoolName : 'St Example Primary');
  add('SchoolStart', o.start !== undefined ? o.start : '08:30');
  add('SchoolEnd', o.end !== undefined ? o.end : '15:15');

  (o.breaks || [['Recess', '10:30', '11:00'], ['Lunch', '13:00', '13:45']]).forEach((item, i) => {
    add(`Break_${i + 1}_Name`, item[0]);
    add(`Break_${i + 1}_Start`, item[1]);
    add(`Break_${i + 1}_End`, item[2]);
  });

  KLAS.forEach((name, i) => add(`KLA_${i + 1}_Name`, name));

  (o.specialists || [['Mr Smith', 'PDHPE', 1], ['Ms Jones', 'Creative Arts', 2]]).forEach((item, i) => {
    add(`Specialist_${i + 1}_Name`, item[0]);
    add(`Specialist_${i + 1}_KLA`, item[1]);
    add(`Specialist_${i + 1}_Cycle`, item[2]);
  });

  (o.classes || [
    ['K Blue', 'Ms Adams', 'Kindergarten', 1],
    ['K Red', 'Mr Brown', 'Kindergarten', 1],
    ['Year 3 Gold', 'Ms Chen', 'Year 3', 2]
  ]).forEach((item, i) => {
    add(`Class_${i + 1}_Name`, item[0]);
    add(`Class_${i + 1}_Teacher`, item[1]);
    add(`Class_${i + 1}_Grade`, item[2]);
    add(`Class_${i + 1}_Cycle`, item[3]);
    add(`Class_${i + 1}_Color`, item[4] || '#FFADAD');
  });

  (o.activities || ['Assembly', 'Library', 'Amadeus - Tutorials']).forEach((name, i) => {
    add(`Activity_${i + 1}_Name`, name);
  });

  return { jobId: o.jobId || 'job_test', exportVariables: variables };
}

/** A large school: 4 classes per grade K-6, 3-week cycles, 6 specialists. */
function makeLargeSchoolPayload() {
  const grades = ['Kindergarten', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5', 'Year 6'];
  const classes = [];
  grades.forEach(grade => {
    ['Blue', 'Red', 'Gold', 'Green'].forEach(colour => {
      classes.push([grade + ' ' + colour, 'Teacher ' + (classes.length + 1), grade, 3]);
    });
  });
  const specialists = [
    ['PE Teacher', 'PDHPE', 1], ['Music Teacher', 'Creative Arts', 1], ['Art Teacher', 'Creative Arts', 2],
    ['Science Teacher', 'Science & Tech', 1], ['RE Teacher', 'Religion', 3], ['Library Teacher', 'English', 2]
  ];
  return makePayload({ classes, specialists });
}

module.exports = { makePayload, makeLargeSchoolPayload, KLAS };
