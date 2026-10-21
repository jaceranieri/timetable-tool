/**
 * Development only - leave this file out of the live wizard project.
 *
 * Generates a large test school (4 classes per grade K-6 on a 3-week cycle,
 * 6 specialists) and logs how long each step took, ending with a timing
 * summary. In the Apps Script editor, choose runLoadTest and click Run,
 * then copy the log from the execution panel (or under Executions). It
 * creates a real timetable in your Drive; delete it afterwards.
 */
function runLoadTest() {
  const grades = ['Kindergarten', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5', 'Year 6'];
  const colours = ['Blue', 'Red', 'Gold', 'Green'];
  const specialists = [
    ['PE Teacher', 'PDHPE', 1],
    ['Music Teacher', 'Creative Arts', 1],
    ['Art Teacher', 'Creative Arts', 2],
    ['Science Teacher', 'Science & Tech', 1],
    ['RE Teacher', 'Religion', 3],
    ['Library Teacher', 'English', 2]
  ];
  const klas = ['English', 'Mathematics', 'Religion', 'Science & Tech', 'HSIE', 'PDHPE', 'Creative Arts', 'Non-Curricular'];

  const variables = [];
  const add = (name, value) => variables.push({ name: name, value: value });
  add('SchoolName', 'Load Test School');
  add('SchoolStart', '08:30');
  add('SchoolEnd', '15:15');
  [['Recess', '10:30', '11:00'], ['Lunch', '13:00', '13:45']].forEach((item, i) => {
    add('Break_' + (i + 1) + '_Name', item[0]);
    add('Break_' + (i + 1) + '_Start', item[1]);
    add('Break_' + (i + 1) + '_End', item[2]);
  });
  klas.forEach((name, i) => add('KLA_' + (i + 1) + '_Name', name));
  specialists.forEach((item, i) => {
    add('Specialist_' + (i + 1) + '_Name', item[0]);
    add('Specialist_' + (i + 1) + '_KLA', item[1]);
    add('Specialist_' + (i + 1) + '_Cycle', item[2]);
  });
  let classNumber = 1;
  grades.forEach(grade => {
    colours.forEach(colour => {
      add('Class_' + classNumber + '_Name', grade + ' ' + colour);
      add('Class_' + classNumber + '_Teacher', 'Teacher ' + classNumber);
      add('Class_' + classNumber + '_Grade', grade);
      add('Class_' + classNumber + '_Cycle', 3);
      add('Class_' + classNumber + '_Color', '#FFADAD');
      classNumber++;
    });
  });
  ['Assembly', 'Library', AMADEUS_ACTIVITY_NAME].forEach((name, i) => add('Activity_' + (i + 1) + '_Name', name));

  const startedAt = Date.now();
  const url = generateTimetable(
    { jobId: 'load_test_' + startedAt, exportVariables: variables },
    { timing: true }
  );
  const seconds = Math.round((Date.now() - startedAt) / 1000);
  console.log('Load test: 28 classes and 6 specialists generated in ' + seconds + 's (limit: 360s). ' + url);
  return seconds;
}
