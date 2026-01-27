function summarizeTeacherStatusPerCourse() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Teacher Courses');
  const outputSheet = ss.getSheetByName('Teacher Courses Data') || ss.insertSheet('Teacher Courses Data');

  outputSheet.clearContents();

  const data = sourceSheet.getDataRange().getValues();

  const courseStartCol = 4; // E column (0-based)
  const courseEndCol = 69;  // BR column
  const courseNames = data[1].slice(courseStartCol, courseEndCol + 1);

  const teacherData = data.slice(2); // Skip first 2 rows
  const statusColIndex = 3; // D column

  // Headers
  const result = [['Course Name', 'Active', 'EWS', 'PIP', 'In Training', 'Actions']];

  for (let i = 0; i < courseNames.length; i++) {
    const courseName = courseNames[i];
    let activeCount = 0;
    let ewsCount = 0;
    let pipCount = 0;
    let trainingCount = 0;

    teacherData.forEach(row => {
      const status = row[statusColIndex];
      const onboardedValue = row[courseStartCol + i];

      const isOnboarded = onboardedValue &&
        onboardedValue.toString().toLowerCase().trim() !== 'not onboarded' &&
        onboardedValue.toString().trim() !== '';

      if (isOnboarded) {
        switch (status) {
          case 'Active':
            activeCount++;
            break;
          case 'EWS':
            ewsCount++;
            break;
          case 'PIP':
            pipCount++;
            break;
          case 'In Training':
            trainingCount++;
            break;
        }
      }
    });

    result.push([courseName, activeCount, ewsCount, pipCount, trainingCount, '']);
  }

  outputSheet.getRange(1, 1, result.length, result[0].length).setValues(result);
}
