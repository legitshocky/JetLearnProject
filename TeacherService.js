function getTeacherData() {
  try {
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA);

    if (sheetData.length < 2) { 
      Logger.log("Teacher Data sheet is empty or only has headers.");
      return [];
    }

    const teachers = sheetData.slice(1).map(row => {
      // 1. Determine Prefix based on Gender (Column C / Index 2)
      const gender = String(row[2] || '').trim().toUpperCase();
      let autoPrefix = '';
      if (gender === 'M') {
        autoPrefix = 'Mr.';
      } else if (gender === 'F') {
        autoPrefix = 'Ms.';
      }

      return { 
        name: String(row[1] || '').trim(),
        email: String(row[8] || '').trim(),
        clsEmail: String(row[9] || '').trim(),
        tpManagerEmail: String(row[10] || '').trim(),
        manager: String(row[6] || '').trim(),
        clsManagerResponsible: String(row[7] || '').trim(),
        status: String(row[3] || 'Active').trim(),
        joinDate: row[2],
        
        // 2. Add the calculated prefix property
        prefix: autoPrefix 
      };
    }).filter(person => person.name !== ''); 

    Logger.log(`Found ${teachers.length} teachers.`);
    return teachers;
  } catch (error) {
    Logger.log('Error getting teacher data: ' + error.message);
    return [];
  }
}

function getTeacherCourses() {
  try {
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.TEACHER_COURSES);

    if (sheetData.length < 2) { 
      Logger.log('Teacher Courses sheet is empty or only has headers.');
      return {};
    }

    const coursesByTeacher = {};

    sheetData.slice(1).forEach(row => { 
      const teacher = String(row[0] || '').trim(); 
      const course = String(row[1] || '').trim(); 
      const status = String(row[2] || '').trim(); 
      const progress = String(row[3] || '').trim(); 

      if (!teacher || !course) return; 

      if (!coursesByTeacher[teacher]) {
        coursesByTeacher[teacher] = [];
      }

      coursesByTeacher[teacher].push({
        course: course,
        status: status,
        progress: progress
      });
    });

    return coursesByTeacher;
  } catch (error) {
    Logger.log('Error getting teacher courses: ' + error.message);
    return {};
  }
}
function getActiveTeachers() {
  Logger.log('getActiveTeachers called');

  try {
    const teachers = getTeacherData(); 
    return teachers
      .map(teacher => teacher.name);
  } catch (error) {
    Logger.log('Error getting active teachers: ' + error.message);
    return [];
  }
}

function getTeacherDetailsForTable() {
  Logger.log('getTeacherDetailsForTable called');

  try {
    const teacherData = getTeacherData(); 

    if (teacherData.length === 0) { 
      Logger.log("Teacher Data sheet is empty or only has headers.");
      return [];
    }

    const teacherCourses = getTeacherCourses(); 

    return teacherData.map(teacher => {
      if (!teacher.name) return null;

      const courses = teacherCourses[teacher.name] || [];
      const activeCoursesCount = courses.filter(c =>
        c.status === 'Active' && c.progress !== 'Not Onboarded' && c.progress !== '0%'
      ).length;
      const completedCoursesCount = courses.filter(c => c.status === 'Completed').length;

      return {
        name: teacher.name,
        email: teacher.email,
        clsEmail: teacher.clsEmail || 'N/A',
        status: teacher.status,
        joinDate: teacher.joinDate ? new Date(teacher.joinDate).toLocaleDateString('en-GB') : 'N/A', 
        activeCourses: activeCoursesCount,
        completedCourses: completedCoursesCount,
        lastActivity: getTeacherLastActivity(teacher.name) 
      };
    }).filter(teacher => teacher !== null); 
  } catch (error) {
    Logger.log('Error getting teacher details for table: ' + error.message);
    return [];
  }
}

function getTeacherLastActivity(teacherName) {
  try {
    const auditData = getAuditLog({ limit: 1000 }).data; 

    for (const row of auditData) {
      if ((String(row[4] || '').trim() === teacherName || String(row[5] || '').trim() === teacherName)
          && String(row[1]).includes('Migration')) { 
        return new Date(row[0]).toLocaleString('en-GB'); 
      }
    }

    return 'No recent activity';
  } catch (error) {
    Logger.log('Error getting teacher last activity: ' + error.message);
    return 'Unknown';
  }
}

function addNewTeacher(teacherData) {
  Logger.log('addNewTeacher called for: ' + teacherData.name);

  try {
    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.TEACHER_DATA);

    const existingTeachers = getTeacherData(); 
    if (existingTeachers.some(t => t.name.toLowerCase() === teacherData.name.toLowerCase())) {
      return { success: false, message: 'Teacher with this name already exists' };
    }
    if (teacherData.email && !isValidEmail(teacherData.email)) {
        return { success: false, message: 'Invalid teacher email address' };
    }
    if (teacherData.clsEmail && teacherData.clsEmail.trim() !== '' && !isValidEmail(teacherData.clsEmail)) {
        return { success: false, message: 'Invalid CLS email address' };
    }

    const newRowData = Array(11).fill(''); 
    newRowData[1] = teacherData.name;           
    newRowData[2] = teacherData.joinDate || new Date(); 
    newRowData[3] = teacherData.status || 'Active'; 
    newRowData[6] = teacherData.manager || '';  
    newRowData[7] = teacherData.clsManager || ''; 
    newRowData[8] = teacherData.email || '';    
    newRowData[9] = teacherData.clsEmail || ''; 
    newRowData[10] = teacherData.tpManagerEmail || ''; 

    sheet.appendRow(newRowData);

    logAction('Teacher Added', '', '', '', teacherData.name, '', 'Success', 'New teacher added to system');

    delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.TEACHER_DATA}`];

    return { success: true, message: 'Teacher added successfully' };
  } catch (error) {
    Logger.log('Error adding new teacher: ' + error.message);
    return { success: false, message: 'Error adding teacher: ' + error.message };
  }
}

function updateTeacherDetails(teacherData) {
  Logger.log('updateTeacherDetails called for: ' + teacherData.name);

  try {
    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.TEACHER_DATA);
    const data = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA); 

    let rowIndex = -1;
    for (let i = 0; i < data.length; i++) { 
      if (data[i][1] && String(data[i][1]).trim() === teacherData.name.trim()) {
        rowIndex = i + 1; 
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: 'Teacher not found' };
    }

    if (teacherData.email && !isValidEmail(teacherData.email)) {
        return { success: false, message: 'Invalid teacher email address' };
    }
    if (teacherData.clsEmail && teacherData.clsEmail.trim() !== '' && !isValidEmail(teacherData.clsEmail)) {
        return { success: false, message: 'Invalid CLS email address' };
    }

    sheet.getRange(rowIndex, 9).setValue(teacherData.email);
    sheet.getRange(rowIndex, 10).setValue(teacherData.clsEmail);
    sheet.getRange(rowIndex, 4).setValue(teacherData.status);

    logAction('Teacher Updated', '', '', '', teacherData.name, '', 'Success', `Teacher ${teacherData.name} details updated`);

    delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.TEACHER_DATA}`];

    Logger.log('Teacher details updated: ' + teacherData.name);
    return { success: true, message: 'Teacher details updated successfully' };
  } catch (error)
 {
    Logger.log('Error updating teacher details: ' + error.message);
    return { success: false, message: 'Error updating teacher details: ' + error.message };
  }
}

function getCourseNames() {
  Logger.log('getCourseNames called');
  try {
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.COURSE_NAME); 

    if (sheetData.length < 2) { 
        Logger.log('Course Name sheet is empty or only has headers.');
        return [];
    }

    const courses = sheetData.slice(1) 
      .map(row => (row[0] ? String(row[0]).trim() : '')) 
      .filter(name => name !== ''); 

    Logger.log(`getCourseNames found ${courses.length} courses.`);
    return courses;
  } catch (error) {
    Logger.log('getCourseNames error: ' + error.message);
    return [];
  }
}

function getCourseDetails() {
  Logger.log('getCourseDetails called for Courses Page');

  try {
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.TEACHER_COURSES);

    if (sheetData.length < 2) { 
      Logger.log('Teacher Courses sheet is empty or only has headers.');
      return {};
    }

    const coursesByTeacher = {};

    sheetData.slice(1).forEach(row => { 
      const teacher = String(row[0] || '').trim(); 
      const course = String(row[1] || '').trim(); 
      const status = String(row[2] || '').trim(); 
      const progress = String(row[3] || '').trim(); 

      if (!teacher || !course) return; 

      if (!coursesByTeacher[teacher]) {
        coursesByTeacher[teacher] = [];
      }

      coursesByTeacher[teacher].push({
        course: course,
        status: status,
        progress: progress
      });
    });

    return coursesByTeacher;
  } catch (error) {
    Logger.log('Error getting teacher courses: ' + error.message);
    return {};
  }
}

function getTeachersForCourse(courseName) {
  try {
    const teacherCourses = getTeacherCourses(); 
    const teachers = [];

    for (const teacher in teacherCourses) {
      if (teacherCourses[teacher].some(c => c.course === courseName)) {
        teachers.push(teacher);
      }
    }

    return teachers;
  } catch (error) {
    Logger.log('Error getting teachers for course: ' + error.message);
    return [];
  }
}

function addNewCourse(courseName) {
  Logger.log('addNewCourse called for: ' + courseName);

  try {
    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.COURSE_NAME);

    const existingCourses = getCourseNames(); 
    if (existingCourses.map(c => c.toLowerCase()).includes(courseName.toLowerCase())) {
      return { success: false, message: 'Course already exists' };
    }

    sheet.appendRow([courseName]);

    logAction('Course Added', '', '', '', '', courseName, 'Success', 'New course added to system');

    delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.COURSE_NAME}`];

    return { success: true, message: 'Course added successfully' };
  } catch (error) {
    Logger.log('Error adding new course: ' + error.message);
    return { success: false, message: 'Error adding course: ' + error.message };
  }
}

function getCourseProgressSummary(courseName = null) {
  Logger.log('getCourseProgressSummary called for: ' + (courseName || 'All Courses'));
  try {
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.COURSE_PROGRESS_SUMMARY); 

    if (!sheetData || sheetData.length < 1) { 
      Logger.log('Course Progress Summary sheet not found or is empty.');
      return { headers: [], data: [] };
    }

    const headers = sheetData[0];
    let filteredData = sheetData.slice(1);

    if (courseName) {
      filteredData = filteredData.filter(row => row[0] && String(row[0]).trim() === courseName.trim());
    }

    return { headers: headers, data: filteredData };

  }
   catch (error) {
    Logger.log('Error in getCourseProgressSummary: ' + error.message);
    return { headers: [], data: [], success: false, message: 'Failed to load course summary data. Please ensure the "Course Summary" sheet exists and is accessible.' };
  }
}

function getClsManagerEmailByName(managerName) {
  Logger.log("getClsManagerEmailByName is deprecated. Use findClsEmailByManagerName directly.");
  return findClsEmailByManagerName(managerName);
}




function getTPManagers() {
  try {
    const teacherData = getTeacherData(); 
    const tpManagerNames = new Set();

    teacherData.forEach(teacher => {
      if (teacher.manager && teacher.manager.trim() !== '') {
        tpManagerNames.add(teacher.manager.trim());
      }
    });

    const hardcodedManagers = [
        'Naureen Fatima',
        'Oorja M Srivastava',
        'Sangeeta Sarkar',
        'Sayani Chakraborty'
    ];
    hardcodedManagers.forEach(name => tpManagerNames.add(name));

    return Array.from(tpManagerNames).sort();
  } catch (error) {
    Logger.log('Error getting TP Managers: ' + error.message);
    return [];
  }
}

function getTeacherLabel(hubspotValue) {
  // If the value is empty or not provided, return it as is.
  if (!hubspotValue) {
    return hubspotValue;
  }

  const teacherHsData = _getCachedSheetData(CONFIG.SHEETS.TEACHER_HS_DATA);

  // 1. First, try to find a match assuming the hubspotValue is an INTERNAL ID.
  for (let i = 1; i < teacherHsData.length; i++) { 
    // The internal ID is in the second column (index 1).
    const internalId = teacherHsData[i][1]; 
    // The display name is in the third column (index 2).
    const displayName = teacherHsData[i][2];

    if (internalId === hubspotValue) { 
      // Found a match! Return the proper name.
      Logger.log(`Teacher ID '${hubspotValue}' was successfully mapped to name '${displayName}'.`);
      return displayName; 
    }
  }

  // 2. If no match was found, it means HubSpot likely sent a NAME directly.
  // In this case, we just return the original value because it's already the name we want.
  Logger.log(`Teacher lookup did not find an ID matching '${hubspotValue}'. Assuming this is already the correct name and returning it directly.`);
  return hubspotValue; 
}

function getCourseLabel(internalValue){
  const data = _getCachedSheetData(CONFIG.SHEETS.COURSE_HS_DATA);

  for (let i = 1; i < data.length; i++) { 
    if (data[i][0] === internalValue) {
      console.log(`Course internal: ${internalValue}, label: ${data[i][1]}`);
      return data[i][1]; 
    }
  }
  Logger.log(`Course label not found for internal value: ${internalValue}`);
  return internalValue; 
}

function getHSUserLabel(internalValue){
  const data = _getCachedSheetData(CONFIG.SHEETS.HS_USER_DATA);

  for (let i = 1; i < data.length; i++) { 
    if (String(data[i][0]) == String(internalValue)) {
      console.log(`HS User internal: ${internalValue}, label: ${data[i][1]}`);
      return data[i][1]; 
    }
  }
  Logger.log(`HS User label not found for internal value: ${internalValue}`);
  return internalValue; 
}

function getTeacherLoadData() {
  try {
    // Fetch data from the 'Teacher Courses' sheet (Wide Format)
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.TEACHER_COURSES);
    
    if (!sheetData || sheetData.length < 2) {
      return [];
    }

    const headers = sheetData[0];
    // Based on your example: Teacher(0), Email(1), Manager(2), Health(3). Courses start at 4.
    const COURSE_START_INDEX = 4; 

    const teacherLoads = sheetData.slice(1).map(row => {
      const name = row[0];
      const status = row[3] || 'Active'; // Assuming Health is column 3
      
      // Calculate Load: Count columns where value is NOT 'Not onboarded' and NOT empty
      let activeCount = 0;
      for (let i = COURSE_START_INDEX; i < row.length; i++) {
        const val = String(row[i] || '').trim();
        if (val && val.toLowerCase() !== 'not onboarded') {
          activeCount++;
        }
      }

      return {
        name: name,
        status: status,
        load: activeCount
      };
    });

    return teacherLoads;

  } catch (error) {
    Logger.log('Error getting teacher load data: ' + error.message);
    return [];
  }
}

function getTeacherListForDropdown() {
  try {
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA); // Or TEACHER_COURSES, works either way usually
    if (!sheetData || sheetData.length < 2) return [];

    // Filter out header row where name is 'Teacher' or empty
    const teachers = sheetData
      .map(row => String(row[1] || '').trim()) // Assuming Name is Col B (Index 1) in Teacher Data
      .filter(name => name !== '' && name.toLowerCase() !== 'teacher' && name.toLowerCase() !== 'teacher name')
      .sort();

    return teachers;
  } catch (e) {
    Logger.log("Error getting teacher list: " + e.message);
    return [];
  }
}

function getTeacherSpecificLoad(teacherName) {
  try {
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.TEACHER_COURSES);
    if (!sheetData || sheetData.length < 2) return { success: false, message: "No data found." };

    // 1. Header Detection (Fixes the issue where 'Teacher' showed as a person)
    let headerRowIndex = -1;
    for(let i = 0; i < Math.min(sheetData.length, 10); i++) {
      if(String(sheetData[i][0]).trim().toLowerCase() === 'teacher') {
        headerRowIndex = i;
        break;
      }
    }
    if (headerRowIndex === -1) headerRowIndex = 0;

    const headers = sheetData[headerRowIndex];
    
    // 2. Find specific teacher row
    const teacherRow = sheetData.slice(headerRowIndex + 1).find(row => 
        String(row[0]).trim().toLowerCase() === String(teacherName).trim().toLowerCase()
    );

    if (!teacherRow) return { success: false, message: "Teacher not found." };

    const courseDetails = [];
    const COURSE_START_INDEX = 4; // Teacher, Email, Manager, Health = 0,1,2,3

    // 3. Filter Logic
    for (let i = COURSE_START_INDEX; i < headers.length; i++) {
      const courseName = String(headers[i] || '').trim();
      const status = String(teacherRow[i] || '').trim();

      // Only show courses that are NOT "Not onboarded" and have a valid Header Name
      if (status && status.toLowerCase() !== 'not onboarded' && courseName !== '') {
        courseDetails.push({
          course: courseName,
          proficiency: status
        });
      }
    }

    // Sort 100% to top
    courseDetails.sort((a, b) => b.proficiency.localeCompare(a.proficiency));

    // 4. Get Last Activity (using your existing helper function)
    const lastActivity = getTeacherLastActivity(teacherName);

    return { 
      success: true, 
      teacherName: teacherRow[0], 
      status: teacherRow[3] || 'Active',
      courses: courseDetails,
      totalLoad: courseDetails.length,
      lastActivity: lastActivity
    };

  } catch (error) {
    return { success: false, message: error.message };
  }
}

function searchMatchingTeachers(requestData) {
  Logger.log('searchMatchingTeachers called with:', requestData);

  try {
    const mainData = _getCachedSheetData(CONFIG.SHEETS.PERSONA_DATA, CONFIG.PERSONA_SHEET_ID);

    if (mainData.length < 2) {
      Logger.log("Persona Data sheet is empty or has no headers.");
      return { success: true, results: [] };
    }
    const headers = mainData[1];

    const requestedDate = requestData.requestedDate;
    const requestedSlot = requestData.requestedSlot;
    const currentCourse = requestData.currentCourse;

    const futureCourses = [
      requestData.futureCourse1,
      requestData.futureCourse2,
      requestData.futureCourse3
    ].filter(Boolean);

    const normalize = arr => arr.map(t => String(t).trim()).filter(t => !!t);

    const mathTraits = normalize(requestData.mathTraits ? requestData.mathTraits.split(',') : []);
    const techTraits = normalize(requestData.techTraits ? requestData.techTraits.split(',') : []);

    const output = [];

    const headerMap = {};
    headers.forEach((h, idx) => {
      if (h) {
        headerMap[String(h).trim()] = idx;
      }
    });

    const teacherNameCol = headerMap['Teacher Name'];
    const teacherStatusCol = headerMap['Status'];
    const ageOrYearCol_Math = headerMap['Math Age/Year (preferred)'] || headerMap['Age/Year'];
    const ageOrYearCol_Tech = headerMap['Tech Age/Year (preferred)'] || headerMap['Age/Year'];
    const traitColsStart = headerMap['Trait 1'];
    const traitColsEnd = headerMap['Trait 9'];

    if ([teacherNameCol, teacherStatusCol, traitColsStart, traitColsEnd].includes(undefined)) {
      throw new Error("Missing required persona sheet columns (e.g., 'Teacher Name', 'Status', 'Trait 1').");
    }

    const progressOrder = ["100%", "91-99%", "81-90%", "71-80%", "61-70%", "51-60%", "41-50%", "31-40%", "21-30%", "11-20%", "1-10%", "0%", "Not Onboarded", "N/A"];
    const requestedDateObj = new Date(requestedDate);
    const requestedDateStr = requestedDateObj.toISOString().split('T')[0];

    for (let i = 2; i < mainData.length; i++) {
      const row = mainData[i];
      const teacherStatus = String(row[teacherStatusCol] || '').trim();
      if (teacherStatus !== "Active") continue;
      const teacherName = String(row[teacherNameCol] || '').trim();
      if (!teacherName) continue;

      let currentCourseProgress = 'N/A';
      if (currentCourse) {
        const currentCourseColIndex = headerMap[currentCourse];
        if (currentCourseColIndex !== undefined) {
          currentCourseProgress = String(row[currentCourseColIndex] || 'N/A').trim();
          const validStatuses = ["71-80%", "81-90%", "91-99%", "100%"];
          if (!validStatuses.includes(currentCourseProgress)) continue;
        }
      }

      const futureCourseStatuses = futureCourses.map(fc => {
        const fcColIndex = headerMap[fc];
        return fcColIndex !== undefined ? String(row[fcColIndex] || "N/A").trim() : "N/A";
      });

      const teacherTraitRaw = row.slice(traitColsStart, traitColsEnd + 1);
      const teacherTraits = normalize(teacherTraitRaw.flatMap(cell => String(cell).split(/\n|,/)));
      const normalizedTeacherTraits = new Set(teacherTraits.map(t => t.toLowerCase()));
      const isMathCourse = currentCourse && currentCourse.toLowerCase().includes("math");
      const targetTraits = isMathCourse ? mathTraits.map(t => t.toLowerCase()) : techTraits.map(t => t.toLowerCase());
      const traitMissing = targetTraits.filter(t => !normalizedTeacherTraits.has(t));
      const traitMatchesCount = targetTraits.length - traitMissing.length;
      const ageOrYearMatch = isMathCourse ? String(row[ageOrYearCol_Math] || 'N/A').trim() : String(row[ageOrYearCol_Tech] || 'N/A').trim();
      
      let slotMatch = '❌';
      let alternateSlots = [];
      const slotHeaderKey = Object.keys(headerMap).find(h => {
        try {
          return !isNaN(new Date(h).getTime()) && new Date(h).toISOString().split('T')[0] === requestedDateStr;
        } catch (e) { return false; }
      });
      const availability = slotHeaderKey ? String(row[headerMap[slotHeaderKey]] || '').trim() : "Date Column Not Found";
      if (availability.includes(requestedSlot)) slotMatch = '✔️';

      const formatDateForAltSlot = date => date.toLocaleDateString('en-GB', { weekday: 'short', day: '2-digit', month: 'short' });
      for (let d = 0; d < 3; d++) {
        const thisDate = new Date(requestedDateObj);
        thisDate.setDate(requestedDateObj.getDate() + d);
        const formattedDateForHeader = thisDate.toISOString().split('T')[0];
        const currentDayHeaderKey = Object.keys(headerMap).find(h => {
          try {
            return !isNaN(new Date(h).getTime()) && new Date(h).toISOString().split('T')[0] === formattedDateForHeader;
          } catch (e) { return false; }
        });
        if (currentDayHeaderKey !== undefined) {
          const slotVal = String(row[headerMap[currentDayHeaderKey]] || '').trim();
          if (slotVal && !["No Slots", "No Slots Available"].includes(slotVal)) {
            const slotList = slotVal.split(',').map(s => s.trim()).filter(Boolean);
            if (slotList.length > 0) alternateSlots.push(`${formatDateForAltSlot(thisDate)}: ${slotList.join(', ')}`);
          }
        }
      }
      
      output.push({
        teacherName, ageYear: ageOrYearMatch, slotMatch, alternateSlots: alternateSlots.join('<br>'),
        currentCourseProgress, futureCourse1Progress: futureCourseStatuses[0] || 'N/A',
        futureCourse2Progress: futureCourseStatuses[1] || 'N/A', futureCourse3Progress: futureCourseStatuses[2] || 'N/A',
        traitsMissing, _traitMatchesCount: traitMatchesCount, _currentCourseProgressOrder: progressOrder.indexOf(currentCourseProgress)
      });
    }

    output.sort((a, b) => {
      if (a.slotMatch !== b.slotMatch) return a.slotMatch === '✔️' ? -1 : 1;
      if (a._currentCourseProgressOrder !== b._currentCourseProgressOrder) return a._currentCourseProgressOrder - b._currentCourseProgressOrder;
      return b._traitMatchesCount - a._traitMatchesCount;
    });

    Logger.log('Found ' + output.length + ' matching teachers');
    return { success: true, results: output };

  } catch (error) {
    Logger.log('Error in searchMatchingTeachers: ' + error.message);
    return { success: false, message: 'Error searching teachers: ' + error.message };
  }
}

function updateTeacherPersona(teacherData) {
  Logger.log('updateTeacherPersona called for teacher: ' + teacherData['Teacher Name']);

  try {
    const spreadsheet = _getSpreadsheet(CONFIG.PERSONA_SHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.PERSONA_DATA);
    const data = _getCachedSheetData(CONFIG.SHEETS.PERSONA_DATA, CONFIG.PERSONA_SHEET_ID);
    const headers = data[1];

    const headerColMap = {};
    headers.forEach((h, idx) => {
      if (h) {
        headerColMap[String(h).trim()] = idx;
      }
    });

    const teacherNameColIndex = headerColMap['Teacher Name'];
    if (teacherNameColIndex === undefined) {
      throw new Error("'Teacher Name' column not found in Persona Sheet headers.");
    }
    
    let rowIndex = -1;
    for (let i = 2; i < data.length; i++) {
      if (data[i][teacherNameColIndex] && String(data[i][teacherNameColIndex]).trim() === teacherData['Teacher Name'].trim()) {
        rowIndex = i + 1;
        break;
      }
    }
    
    const rowToUpdate = rowIndex !== -1 ? data[rowIndex - 1].slice() : Array(headers.length).fill('');

    for (const key in teacherData) {
        const colIndex = headerColMap[key.trim()];
        if (colIndex !== undefined) {
            rowToUpdate[colIndex] = teacherData[key];
        } else {
            Logger.log(`Warning: Key '${key}' not found in Persona Sheet headers.`);
        }
    }

    if (rowIndex === -1) {
      sheet.appendRow(rowToUpdate);
      Logger.log('Added new teacher persona: ' + teacherData['Teacher Name']);
    } else {
      sheet.getRange(rowIndex, 1, 1, rowToUpdate.length).setValues([rowToUpdate]);
      Logger.log('Updated teacher persona: ' + teacherData['Teacher Name']);
    }

    delete _sheetDataCache[`${CONFIG.PERSONA_SHEET_ID}_${CONFIG.SHEETS.PERSONA_DATA}`];

    return { success: true, message: 'Teacher persona updated successfully' };
  } catch (error) {
    Logger.log('Error updating teacher persona: ' + error.message);
    return { success: false, message: 'Error updating teacher persona: ' + error.message };
  }
}

function searchTeacherPersonas(searchTerm) {
  Logger.log('searchTeacherPersonas called with term: ' + searchTerm);

  try {
    const allPersonas = getTeacherPersonaData(); 
    const term = searchTerm.toLowerCase();

    const results = allPersonas.filter(persona => {
      return Object.values(persona).some(value =>
        value && String(value).toLowerCase().includes(term)
      );
    });

    Logger.log('Found ' + results.length + ' matching teacher personas');
    return results;
  }
  catch (error) {
    Logger.log('Error searching teacher personas: ' + error.message);
    return [];
  }
}

function findClsEmailByManagerName(managerName) {
  if (!managerName || typeof managerName !== "string") {
    Logger.log("findClsEmailByManagerName: Manager name is empty or invalid.");
    return null;
  }

  try {
    Logger.log(`[DEBUG] findClsEmailByManagerName called for managerName: "${managerName}"`);
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA) || [];

    if (!Array.isArray(sheetData) || sheetData.length < 2) {
      Logger.log("findClsEmailByManagerName: Teacher Data sheet is empty or has only headers.");
      return null;
    }

    for (let i = 1; i < sheetData.length; i++) {
      const row = sheetData[i];
      const clsManagerName = row[7] ? String(row[7]).trim().toLowerCase() : "";
      const clsManagerEmail = row[9] ? String(row[9]).trim() : "";

      if (clsManagerName && managerName.trim().toLowerCase() === clsManagerName) {
        Logger.log(`[DEBUG] Match found for managerName: "${managerName}", returning email: "${clsManagerEmail}"`);
        return clsManagerEmail || null;
      }
    }

    Logger.log(`[WARN] No match found in Teacher Data for managerName: "${managerName}"`);
    return null;
  } catch (error) {
    Logger.log(`[ERROR] findClsEmailByManagerName failed for managerName "${managerName}": ${error.message}`);
    return null;
  }
}
