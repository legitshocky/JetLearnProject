function getUpskilledTeachersAbove50() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Teacher Courses');
  const outputSheet = ss.getSheetByName('Teacher Courses Data');

  const data = sourceSheet.getDataRange().getValues(); // Full data
  const headers = data[1].slice(4); // Courses from E2:BR2
  const teacherRows = data.slice(2); // Starting from row 3

  const results = [];
  results.push(["Teacher Name", "Upskilled Courses (Above 50%)"]);

  teacherRows.forEach(row => {
    const teacherName = row[1]; // Column B = Teacher Name
    const scores = row.slice(4); // From Column E onwards

    const upskilledCourses = [];

    scores.forEach((score, index) => {
      if (typeof score === 'number' && score >= 50) {
        upskilledCourses.push(headers[index]); // Course header
      }
    });

    if (upskilledCourses.length > 0) {
      results.push([teacherName, upskilledCourses.join(', ')]);
    }
  });

  outputSheet.clearContents();

  if (results.length > 1) {
    outputSheet.getRange(1, 1, results.length, results[0].length).setValues(results);
  } else {
    outputSheet.getRange(1, 1).setValue("No teachers upskilled above 50% in any course.");
  }
}
