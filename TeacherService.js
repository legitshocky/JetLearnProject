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
  Logger.log('getTeacherDetailsForTable (V2 - Corrected) called');

  try {
    // \u2500\u2500 Step 1: Get the definitive list of all teachers from the main data sheet. \u2500
    // This is more reliable as it's the source of truth for all personnel.
    const allTeachersFromDataSheet = getTeacherData(); 
    if (!allTeachersFromDataSheet || allTeachersFromDataSheet.length === 0) {
      Logger.log('Main Teacher Data sheet is empty. Cannot build table.');
      return [];
    }

    // \u2500\u2500 Step 2: Fetch true active learner counts for all teachers from HubSpot \u2500
    const hubspotCounts = getActiveLearnersPerTeacher();
    Logger.log('Fetched HubSpot active learner counts.');

    // \u2500\u2500 Step 3: Build a Persona sheet manager lookup map \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    // Persona sheet has plain-text manager names (not emails)
    const personaManagerMap = {};
    try {
      const personaData = _getCachedSheetData(CONFIG.SHEETS.PERSONA_DATA, CONFIG.PERSONA_SHEET_ID);
      if (personaData && personaData.length > 1) {
        const pHeaders = personaData[1];
        const pHeaderMap = {};
        pHeaders.forEach((h, idx) => { if (h) pHeaderMap[String(h).trim()] = idx; });
        const pNameCol    = pHeaderMap['Teacher Name'];
        const pManagerCol = pHeaderMap['Manager'];
        const pClsCol     = pHeaderMap['CLS Manager'];
        const pHiddenCol  = pHeaderMap['Hidden In Search'];
        if (pNameCol !== undefined) {
          for (let r = 2; r < personaData.length; r++) {
            const pName = String(personaData[r][pNameCol] || '').trim();
            if (pName) {
              personaManagerMap[pName.toLowerCase()] = {
                tpManager:      pManagerCol !== undefined ? String(personaData[r][pManagerCol] || '').trim() : '',
                clsManager:     pClsCol     !== undefined ? String(personaData[r][pClsCol]     || '').trim() : '',
                hiddenInSearch: pHiddenCol  !== undefined ? String(personaData[r][pHiddenCol]  || '').trim() : ''
              };
            }
          }
        }
      }
    } catch(e) {
      Logger.log('[getTeacherDetailsForTable] Persona map error: ' + e.message);
    }

    // \u2500\u2500 Step 4: Build the final, enriched result array \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    const finalTeacherDetails = allTeachersFromDataSheet.map(teacher => {
      const teacherName = teacher.name;
      
      // Try all name variants: DB exact, DB normalized, HS canonical, HS normalized
      const teacherNorm  = teacherName.trim().toLowerCase().replace(/\s+/g, ' ');
      const resolvedName = resolveTeacherName(teacherName);
      const resolvedNorm = resolvedName.trim().toLowerCase().replace(/\s+/g, ' ');
      const hsData = hubspotCounts[teacherName]  ||
                     hubspotCounts[teacherNorm]   ||
                     hubspotCounts[resolvedName]  ||
                     hubspotCounts[resolvedNorm]  ||
                     { total: 0, coding: 0, math: 0 };
      if (hsData.total === 0) Logger.log('[ZERO] "' + teacherName + '" tried:"' + teacherNorm + '","' + resolvedName + '","' + resolvedNorm + '"');

      // Get manager names from Persona sheet (plain names, not emails)
      const personaEntry = personaManagerMap[teacherNorm] || personaManagerMap[resolvedNorm] || {};
      
      return {
        name:          teacherName,
        email:         teacher.email || 'N/A',
        clsEmail:      teacher.clsEmail || 'N/A',
        status:        teacher.status || 'Active',
        joinDate:      teacher.joinDate ? new Date(teacher.joinDate).toLocaleDateString('en-GB') : 'N/A',
        
        // Manager fields \u2014 from Persona sheet (names) with Teacher Data fallback (may be emails)
        manager:               personaEntry.tpManager  || teacher.manager               || '',
        clsManagerResponsible: personaEntry.clsManager || teacher.clsManagerResponsible || '',
        tpManagerEmail:        teacher.tpManagerEmail  || '',

        // Use live HubSpot data \u2014 but only for active/EWS teachers.
        // Attrited/On Leave teachers may still appear in old HubSpot deals; zero them out.
        activeCourses:  (['Active','EWS','Friendly'].includes(teacher.status) ? hsData.total  : 0),
        activeCoding:   (['Active','EWS','Friendly'].includes(teacher.status) ? hsData.coding : 0),
        activeMath:     (['Active','EWS','Friendly'].includes(teacher.status) ? hsData.math   : 0),
        hiddenInSearch: personaEntry.hiddenInSearch || '',

        // Get last activity from the audit log
        lastActivity:  getTeacherLastActivity(teacherName)
      };
    });

    Logger.log(`Successfully built details for ${finalTeacherDetails.length} teachers.`);
    return finalTeacherDetails;

  } catch (error) {
    Logger.log('FATAL Error in getTeacherDetailsForTable: ' + error.message);
    return []; // Return an empty array on failure to prevent frontend crashes.
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
    
    // 2. Find specific teacher row in TEACHER_COURSES
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

    // 4. Get Last Activity
    const lastActivity = getTeacherLastActivity(teacherName);

    // 5. Pull manager fields from Teacher Persona Mapping sheet
    // Persona sheet headers (row 1): Teacher Name, Manager, CLS Manager
    // Names are stored as plain text (not emails) \u2014 no resolution needed
    let tpManager = '', tpManagerEmail = '', clsManager = '', clsEmail = '';
    try {
      const personaData = _getCachedSheetData(CONFIG.SHEETS.PERSONA_DATA, CONFIG.PERSONA_SHEET_ID);
      if (personaData && personaData.length > 1) {
        const pHeaders = personaData[1];
        const pHeaderMap = {};
        pHeaders.forEach(function(h, idx) { if (h) pHeaderMap[String(h).trim()] = idx; });

        const pNameCol    = pHeaderMap['Teacher Name'];
        const pManagerCol = pHeaderMap['Manager'];
        const pClsCol     = pHeaderMap['CLS Manager'];

        if (pNameCol !== undefined) {
          const nameLower = String(teacherName).trim().toLowerCase();
          for (let r = 2; r < personaData.length; r++) {
            const rowName = String(personaData[r][pNameCol] || '').trim().toLowerCase();
            if (rowName === nameLower) {
              tpManager  = pManagerCol !== undefined ? String(personaData[r][pManagerCol] || '').trim() : '';
              clsManager = pClsCol     !== undefined ? String(personaData[r][pClsCol]     || '').trim() : '';
              // Emails not stored in Persona sheet \u2014 look up from Teacher Data as fallback
              break;
            }
          }
        }
      }
    } catch(e) {
      Logger.log('[getTeacherSpecificLoad] Persona manager lookup error: ' + e.message);
    }

    // Fallback: if Persona sheet had no data, try Teacher Data sheet for email-based lookup
    if (!tpManager && !clsManager) {
      try {
        const teacherDataSheet = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA);
        const nameLower = String(teacherName).trim().toLowerCase();
        for (let r = 1; r < teacherDataSheet.length; r++) {
          const rowName = String(teacherDataSheet[r][1] || '').trim().toLowerCase();
          if (rowName === nameLower) {
            tpManager      = String(teacherDataSheet[r][6]  || '').trim();
            clsManager     = String(teacherDataSheet[r][7]  || '').trim();
            clsEmail       = String(teacherDataSheet[r][9]  || '').trim();
            tpManagerEmail = String(teacherDataSheet[r][10] || '').trim();
            break;
          }
        }
      } catch(e) {
        Logger.log('[getTeacherSpecificLoad] Teacher Data fallback error: ' + e.message);
      }
    }

    return { 
      success: true, 
      teacherName: teacherRow[0], 
      status: teacherRow[3] || 'Active',
      courses: courseDetails,
      totalLoad: courseDetails.length,
      lastActivity: lastActivity,
      manager:               tpManager,
      clsManagerResponsible: clsManager,
      clsEmail:              clsEmail,
      tpManagerEmail:        tpManagerEmail
    };

  } catch (error) {
    return { success: false, message: error.message };
  }
}

function searchMatchingTeachers(requestData) {
  Logger.log('searchMatchingTeachers called: ' + JSON.stringify(requestData).substring(0, 500));

  try {
    // ── 1. Parse request ──
    var requestedSlots = requestData.requestedSlots || [];
    if (typeof requestedSlots === 'string') {
      try { requestedSlots = JSON.parse(requestedSlots); } catch(e) { requestedSlots = []; }
    }
    if (!Array.isArray(requestedSlots) || requestedSlots.length === 0) {
      if (requestData.requestedDate && requestData.requestedSlot) {
        requestedSlots = [{ date: requestData.requestedDate, slot: requestData.requestedSlot }];
      }
    }
    Logger.log('[SMT] slots: ' + JSON.stringify(requestedSlots));

    var currentCourse = String(requestData.currentCourse || '').trim();
    var futureCourses = [
      String(requestData.futureCourse1 || '').trim(),
      String(requestData.futureCourse2 || '').trim(),
      String(requestData.futureCourse3 || '').trim()
    ].filter(function(f) { return f && f !== 'None'; });
    var learnerAge   = String(requestData.learnerAge || '').trim();
    var isMathCourse = currentCourse.toLowerCase().indexOf('math') !== -1;

    // ── 2. Load Teacher Persona Mapping (same source as findSimilarTeachers) ──
    var personaData = _getCachedSheetData('Teacher Persona Mapping');
    if (!personaData || personaData.length < 2) {
      return { success: false, message: 'Teacher Persona Mapping sheet not found in migration spreadsheet.' };
    }
    var pHeaders = personaData[0].map(function(h) { return String(h).trim(); });
    var nameIdx = pHeaders.indexOf('Teacher Name');
    if (nameIdx === -1) return { success: false, message: "'Teacher Name' column not found in Teacher Persona Mapping." };

    var ageGroupIdx = pHeaders.indexOf('Preferred Age Group');
    if (ageGroupIdx === -1) ageGroupIdx = pHeaders.indexOf('Age Group');

    // Detect trait cols using same broad regex as findSimilarTeachers
    var traitCols = pHeaders.reduce(function(acc, h, i) {
      if (/trait|expertise|style|skill|strength|personality|subject|teaching/i.test(h)) acc.push(i);
      return acc;
    }, []);
    var hiddenIdx = pHeaders.indexOf('Hidden In Search');
    Logger.log('[SMT] Persona Mapping loaded. Trait cols: ' + traitCols.map(function(i) { return pHeaders[i]; }).join(', '));

    // ── 3. Build Teacher Courses map (wide format: col0=Teacher, col1=Email, col2=Manager, col3=Health, col4+=courses) ──
    var teacherCourseMap = {}; // normalizedName → { 'course name lower': progress }
    var upskillCountMap  = {}; // normalizedName → count of non-"Not Onboarded" courses
    try {
      var tcSheet = _getCachedSheetData(CONFIG.SHEETS.TEACHER_COURSES);
      if (tcSheet && tcSheet.length > 1) {
        // Find header row (first row where col 0 === 'teacher')
        var tcHeaderIdx = 0;
        for (var hi = 0; hi < Math.min(tcSheet.length, 10); hi++) {
          if (String(tcSheet[hi][0]).trim().toLowerCase() === 'teacher') { tcHeaderIdx = hi; break; }
        }
        var tcHeaders = tcSheet[tcHeaderIdx];
        var COURSE_START = 4; // Teacher | Email | Manager | Health → courses start at col 4
        tcSheet.slice(tcHeaderIdx + 1).forEach(function(row) {
          var rawName = String(row[0] || '').trim();
          if (!rawName || rawName.toLowerCase() === 'teacher') return;
          var key = normalizeTeacherName(rawName);
          var map = {};
          var count = 0;
          for (var ci = COURSE_START; ci < tcHeaders.length; ci++) {
            var cName = String(tcHeaders[ci] || '').trim();
            var prog  = String(row[ci] || '').trim();
            if (!cName) continue;
            map[cName.toLowerCase()] = prog || 'Not Onboarded';
            if (prog && prog.toLowerCase() !== 'not onboarded') count++;
          }
          teacherCourseMap[key] = map;
          upskillCountMap[key]  = count;
        });
        Logger.log('[SMT] teacherCourseMap: ' + Object.keys(teacherCourseMap).length + ' teachers');
        var sk = Object.keys(teacherCourseMap)[0];
        if (sk) Logger.log('[SMT] sample courses for "' + sk + '": ' + Object.keys(teacherCourseMap[sk]).slice(0,5).join(' | '));
      }
    } catch(tce) {
      Logger.log('[SMT] TEACHER_COURSES error: ' + tce.message);
    }

    // Fuzzy course progress lookup
    function getCourseProgress(tNorm, courseName) {
      if (!courseName) return null;
      var cLower = courseName.toLowerCase().trim();
      var courses = teacherCourseMap[tNorm];
      if (!courses) return null;
      if (courses[cLower] !== undefined) return courses[cLower];
      var courseKeys = Object.keys(courses);
      for (var ci = 0; ci < courseKeys.length; ci++) {
        var ck = courseKeys[ci];
        if (ck.indexOf(cLower) === 0 || cLower.indexOf(ck) === 0) return courses[ck];
        var prefix = cLower.substring(0, 15);
        if (prefix.length >= 10 && ck.indexOf(prefix) !== -1) return courses[ck];
      }
      return null;
    }

    // ── 4. Slot normaliser ──
    function normaliseSlot(s) {
      if (!s) return '';
      s = String(s).trim().toUpperCase().replace(/\s+/g, ' ').replace(/\s*-\s*/g, ' - ');
      s = s.replace(/\b(\d{1,2}):(\d{2})\b(?!\s*(AM|PM))/gi, function(_, h, m) {
        var hr = parseInt(h, 10);
        var ap = hr < 12 ? 'AM' : 'PM';
        var h12 = hr % 12 || 12;
        return String(h12).padStart(2, '0') + ':' + m + ' ' + ap;
      });
      s = s.replace(/\b(\d):(\d{2})\s*(AM|PM)/g, '0$1:$2 $3');
      return s;
    }

    // ── 5. Build availability maps for each requested date ──
    var uniqueDates = [];
    requestedSlots.forEach(function(rs) {
      if (rs.date && uniqueDates.indexOf(rs.date) === -1) uniqueDates.push(rs.date);
    });
    var availabilityByDate = {};
    uniqueDates.forEach(function(dateStr) {
      try {
        availabilityByDate[dateStr] = _getTeacherAvailabilityMap(dateStr);
        Logger.log('[SMT] avail ' + dateStr + ': ' + Object.keys(availabilityByDate[dateStr]).length + ' teachers');
      } catch(e) {
        Logger.log('[SMT] avail error ' + dateStr + ': ' + e.message);
        availabilityByDate[dateStr] = {};
      }
    });

    // ── 6. Audit score map (last 45 days) ──
    var auditScoreMap = {};
    try { auditScoreMap = _buildAuditScoreMapDays(45); } catch(ae) { Logger.log('[SMT] audit error: ' + ae.message); }

    // ── 7. Parse requested traits ──
    function splitTraits(str) {
      return str ? String(str).split(',').map(function(t){ return t.trim().toLowerCase(); }).filter(Boolean) : [];
    }
    var targetTraits = isMathCourse ? splitTraits(requestData.mathTraits) : splitTraits(requestData.techTraits);

    // ── 8. Learner age number (for age group matching) ──
    var learnerAgeNum = parseInt(learnerAge, 10);

    // ── 9. Progress → course score mapping (0-30pts) ──
    var PROG_SCORE = {
      '100%':30, '91-99%':27, '81-90%':24, '71-80%':21,
      '61-70%':18, '51-60%':15, '41-50%':12, '31-40%':9,
      '21-30%':6,  '11-20%':3,  '1-10%':1,   '0%':0
    };

    var output = [];
    var debugSlotFiltered = 0;

    // ── 10. Main loop — iterate over Teacher Persona Mapping (same as findSimilarTeachers) ──
    var candidates = personaData.slice(1).filter(function(r) {
      var rn = String(r[nameIdx] || '').trim();
      if (!rn) return false;
      if (hiddenIdx > -1 && String(r[hiddenIdx] || '').trim().toLowerCase() === 'yes') return false;
      return true;
    });
    Logger.log('[SMT] Scoring ' + candidates.length + ' candidates from Teacher Persona Mapping');

    candidates.forEach(function(row) {
      var rawName = String(row[nameIdx] || '').trim();
      var tNorm   = normalizeTeacherName(rawName);
      var tCanon  = normalizeTeacherName(resolveTeacherName(rawName));

      // ── SLOT MATCHING ──
      var slotsMatched    = 0;
      var allAlternate    = [];
      var teacherInAnyMap = false;
      for (var si = 0; si < requestedSlots.length; si++) {
        var req     = requestedSlots[si];
        var dateMap = availabilityByDate[req.date] || {};
        var sheetSlots = (dateMap[tNorm] !== undefined) ? dateMap[tNorm]
                       : (dateMap[tCanon] !== undefined) ? dateMap[tCanon]
                       : undefined;
        if (sheetSlots !== undefined) {
          teacherInAnyMap = true;
          if (sheetSlots.length > 0) {
            sheetSlots.forEach(function(s) { if (allAlternate.indexOf(s) === -1) allAlternate.push(s); });
          }
        }
        var reqNorm = normaliseSlot(req.slot);
        var isMatch = sheetSlots && sheetSlots.some(function(s) { return normaliseSlot(s) === reqNorm; });
        if (isMatch) slotsMatched++;
      }
      var totalSlots = requestedSlots.length;
      // Only drop if teacher IS in the map but has zero slots (confirmed unavailable)
      if (totalSlots > 0 && teacherInAnyMap && slotsMatched === 0 && allAlternate.length === 0) {
        debugSlotFiltered++;
        return;
      }

      // ── TRAITS (30pts) — same logic as findSimilarTeachers ──
      var teacherTraits = traitCols.reduce(function(acc, i) {
        String(row[i] || '').split(/[,\n]/).forEach(function(t) {
          var clean = t.trim().toLowerCase();
          if (clean) acc.push(clean);
        });
        return acc;
      }, []);
      var traitScore    = 0;
      var traitsMissing = [];
      if (targetTraits.length > 0) {
        var matchedT  = targetTraits.filter(function(t) { return teacherTraits.indexOf(t) > -1; });
        traitsMissing = targetTraits.filter(function(t) { return teacherTraits.indexOf(t) === -1; });
        traitScore    = Math.round((matchedT.length / targetTraits.length) * 30);
      } else {
        traitScore = 15; // neutral when no traits requested
      }

      // ── AGE (20pts) — same logic as findSimilarTeachers ──
      var candidateAgeGroups = [];
      var ageScore = 10; // neutral default
      if (ageGroupIdx > -1) {
        candidateAgeGroups = String(row[ageGroupIdx] || '').split(',')
          .map(function(a) { return a.trim(); }).filter(Boolean);
        if (!isNaN(learnerAgeNum) && candidateAgeGroups.length > 0) {
          var ageMatched = candidateAgeGroups.some(function(ag) {
            var agL = ag.toLowerCase();
            var rangeParts = agL.match(/(\d+)\s*[-\u2013]\s*(\d+)/);
            if (rangeParts) {
              return learnerAgeNum >= parseInt(rangeParts[1]) && learnerAgeNum <= parseInt(rangeParts[2]);
            }
            var plusMatch = agL.match(/(\d+)\+/);
            if (plusMatch) return learnerAgeNum >= parseInt(plusMatch[1]);
            return false;
          });
          ageScore = ageMatched ? 20 : 5;
        }
      }
      var ageMatchDisplay = candidateAgeGroups.join(', ') || 'N/A';

      // ── COURSE READINESS (30pts) — from Teacher Courses wide format ──
      var currentCourseProgress = 'Not Onboarded';
      var courseScore = 0;
      if (currentCourse) {
        var prog = getCourseProgress(tNorm, currentCourse);
        if (!prog && tCanon !== tNorm) prog = getCourseProgress(tCanon, currentCourse);
        if (prog) {
          currentCourseProgress = prog;
          courseScore = PROG_SCORE[prog] !== undefined ? PROG_SCORE[prog] : 5;
        }
      }

      // Future course progress lookup
      function fcProg(fc) {
        if (!fc) return 'N/A';
        return getCourseProgress(tNorm, fc) || getCourseProgress(tCanon, fc) || 'Not Onboarded';
      }

      // ── AUDIT (20pts) — same as findSimilarTeachers, last 45 days ──
      var auditData    = auditScoreMap[tNorm] || auditScoreMap[tCanon] || null;
      var auditGrade   = '\u2014';
      var redFlagCount = 0;
      var auditScore   = 10; // neutral when no data
      if (auditData) {
        redFlagCount = auditData.redFlags || 0;
        if (auditData.avgScore != null) {
          var sc = auditData.avgScore;
          auditGrade = sc >= 65 ? 'A' : sc >= 50 ? 'B' : sc >= 35 ? 'C' : 'D';
          auditScore = Math.max(0, Math.round((sc / 80) * 20) - Math.min(redFlagCount * 3, 10));
        }
      }

      // ── TOTAL SCORE (max 100: Traits30 + Age20 + Course30 + Audit20) ──
      var totalScore = traitScore + ageScore + courseScore + auditScore;

      // Slot display label
      var isFullSlotMatch = (totalSlots === 0) || (slotsMatched === totalSlots);
      var slotLabel = totalSlots === 0        ? '\u2714\uFE0F'
                    : isFullSlotMatch         ? '\u2714\uFE0F Match All'
                    : slotsMatched > 0        ? '\u26A0\uFE0F ' + slotsMatched + '/' + totalSlots
                    :                           '';

      output.push({
        teacherName           : rawName,
        ageYear               : ageMatchDisplay,
        slotMatch             : slotLabel,
        slotFullMatch         : isFullSlotMatch,
        alternateSlots        : allAlternate.join(', '),
        currentCourseProgress : currentCourseProgress,
        futureCourse1Progress : fcProg(futureCourses[0]),
        futureCourse2Progress : fcProg(futureCourses[1]),
        futureCourse3Progress : fcProg(futureCourses[2]),
        teacherTraits         : teacherTraits,
        traitsMissing         : traitsMissing,
        avgClassScore         : auditData && auditData.avgScore != null ? auditData.avgScore + '/80' : 'No data',
        auditGrade            : auditGrade,
        redFlagCount          : redFlagCount,
        auditCount45          : auditData ? (auditData.auditCount || 0) : 0,
        upskillCount          : upskillCountMap[tNorm] || upskillCountMap[tCanon] || 0,
        traitScore            : traitScore,
        ageScore              : ageScore,
        courseScore           : courseScore,
        auditScore            : auditScore,
        _rankScore            : totalScore,
        _traitMatchesCount    : targetTraits.length > 0 ? (targetTraits.length - traitsMissing.length) : 0,
        _currentCourseProgressOrder : ['100%','91-99%','81-90%','71-80%','61-70%','51-60%',
                                       '41-50%','31-40%','21-30%','11-20%','1-10%','0%',
                                       'Not Onboarded','N/A'].indexOf(currentCourseProgress)
      });
    });

    // Sort: full slot match first, then by total score descending
    output.sort(function(a, b) {
      if (b.slotFullMatch !== a.slotFullMatch) return b.slotFullMatch ? 1 : -1;
      return b._rankScore - a._rankScore;
    });
    Logger.log('[SMT] done: ' + output.length + ' results | slotFiltered=' + debugSlotFiltered);
    return { success: true, results: output };

  } catch (error) {
    Logger.log('Error in searchMatchingTeachers: ' + error.message + '\n' + error.stack);
    return { success: false, message: 'Search failed: ' + error.message };
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

function findSimilarTeachers(targetTeacherName) {
  try {
    // \u2500\u2500 1. Resolve canonical name \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    var resolvedTarget = resolveTeacherName(targetTeacherName);
    Logger.log('[findSimilarTeachers] Input: "' + targetTeacherName + '" \u2192 resolved: "' + resolvedTarget + '"');

    // \u2500\u2500 2. Load Persona Mapping sheet \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    var personaData = _getCachedSheetData('Teacher Persona Mapping');
    if (!personaData || personaData.length < 2) {
      return { success: false, message: 'Persona Mapping sheet not found. Please ensure "Teacher Persona Mapping" sheet exists.' };
    }

    var headers = personaData[0].map(function(h) { return String(h).trim(); });
    Logger.log('[findSimilarTeachers] Persona sheet headers: ' + headers.join(' | '));

    var nameIdx = headers.indexOf('Teacher Name');
    if (nameIdx === -1) return { success: false, message: 'Teacher Name column not found in Persona sheet.' };

    var ageGroupIdx = headers.indexOf('Preferred Age Group');
    if (ageGroupIdx === -1) ageGroupIdx = headers.indexOf('Age Group');

    // \u2705 FIX 1 (Bug 3): Broadened regex so columns like "Teaching Style",
    //    "Personality Trait", "Subject Expertise", "Key Skill" etc. are all detected.
    //    Also logs which columns were found so you can verify in the Apps Script log.
    var traitCols = headers.reduce(function(acc, h, i) {
      if (/trait|expertise|style|skill|strength|personality|subject|teaching/i.test(h)) acc.push(i);
      return acc;
    }, []);
    Logger.log('[findSimilarTeachers] Detected trait columns: '
      + (traitCols.length > 0 ? traitCols.map(function(i) { return headers[i]; }).join(', ') : 'NONE \u2014 check column names!'));

    // \u2500\u2500 3. Find target row \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    var targetRow = null;
    for (var i = 1; i < personaData.length; i++) {
      if (normalizeTeacherName(String(personaData[i][nameIdx])) === normalizeTeacherName(resolvedTarget)) {
        targetRow = personaData[i];
        break;
      }
    }

    // \u2705 FIX 2 (Bug 1): If the teacher is missing from the Persona sheet,
    //    fall back to a scoreless-but-valid result set using all other teachers
    //    instead of returning success:false (which silently shows nothing).
    if (!targetRow) {
      Logger.log('[findSimilarTeachers] WARNING: "' + resolvedTarget + '" not found in Persona Mapping sheet. Running fallback scoring.');

      // Build fallback list from Teacher Courses sheet only (upskill + escalation scoring)
      var upskillCountMapFallback = {};
      var tcSheetFallback = _getCachedSheetData(CONFIG.SHEETS.TEACHER_COURSES);
      if (tcSheetFallback && tcSheetFallback.length > 1) {
        var COURSE_START_FB = 4;
        tcSheetFallback.slice(1).forEach(function(row) {
          var rn = String(row[0] || '').trim();
          if (!rn || rn.toLowerCase() === 'teacher') return;
          var cn = resolveTeacherName(rn);
          var cnt = 0;
          for (var j = COURSE_START_FB; j < row.length; j++) {
            var v = String(row[j] || '').trim().toLowerCase();
            if (v && v !== 'not onboarded') cnt++;
          }
          upskillCountMapFallback[cn] = cnt;
        });
      }

      var escalationMapFallback = getEscalatedTeachersLast90Days();

      var fallbackResults = personaData.slice(1).map(function(r) {
        var rn   = String(r[nameIdx] || '').trim();
        var name = resolveTeacherName(rn);
        var esc  = escalationMapFallback[name] || escalationMapFallback[rn] || escalationMapFallback[normalizeTeacherName(name)] || escalationMapFallback[normalizeTeacherName(rn)] || 0;
        var usk  = upskillCountMapFallback[name] || upskillCountMapFallback[rn] || 0;
        var escalationScore = Math.max(0, 20 - (esc * 4));
        var courseScore     = usk > 0 ? 15 : 0;
        var totalScore      = courseScore + escalationScore;

        var escalationRisk, escalationColor;
        if      (esc === 0) { escalationRisk = 'No Escalations';              escalationColor = '#15803d'; }
        else if (esc <= 2)  { escalationRisk = esc + ' Tickets (Low)';        escalationColor = '#d97706'; }
        else if (esc <= 4)  { escalationRisk = esc + ' Tickets (Medium)';     escalationColor = '#d97706'; }
        else                { escalationRisk = esc + ' Tickets (High Risk)';  escalationColor = '#b91c1c'; }

        return {
          name:            name,
          matchScore:      totalScore,
          traitScore:      0,
          ageScore:        0,
          courseScore:     courseScore,
          escalationScore: escalationScore,
          loadScore:       0,
          courseOverlap:   usk + ' courses upskilled',
          overlapCount:    usk,
          activeLearners:  usk,
          upskillCount:    usk,
          upskillDiff:     usk,
          ageGroupMatch:   'N/A',
          escalations:     esc,
          escalationRisk:  escalationRisk,
          escalationColor: escalationColor,
          stability: { total: esc, risk: esc >= 5 ? 'High' : esc >= 3 ? 'Medium' : 'Stable' }
        };
      }).filter(function(r) { return normalizeTeacherName(r.name) !== normalizeTeacherName(resolvedTarget); })
        .sort(function(a, b) { return b.matchScore - a.matchScore; })
        .slice(0, 8);

      return {
        success:       true,
        data:          fallbackResults,
        aiSummary:     '\u26A0 "' + resolvedTarget + '" has no Persona Mapping entry \u2014 showing partial results based on course load & escalation history only. Add this teacher to the Persona Mapping sheet for full AI-scored replacements.',
        aiEnriched:    false,
        targetContext: {
          name:           resolvedTarget,
          courses:        ['No persona data'],
          activeLearners: 0,
          upskillCount:   0,
          ageGroups:      [],
          escalations:    getEscalatedTeachersLast90Days()[resolvedTarget] || 0
        }
      };
    }

    var targetTraits = traitCols.reduce(function(acc, i) {
      String(targetRow[i] || '').split(/[,\n]/).forEach(function(t) {
        var clean = t.trim().toLowerCase();
        if (clean) acc.push(clean);
      });
      return acc;
    }, []);

    var targetAgeGroups = [];
    if (ageGroupIdx > -1) {
      String(targetRow[ageGroupIdx] || '').split(',').forEach(function(a) {
        var clean = a.trim().toLowerCase();
        if (clean) targetAgeGroups.push(clean);
      });
    }

    Logger.log('[findSimilarTeachers] Target traits: ' + (targetTraits.length ? targetTraits.join(', ') : 'NONE'));
    Logger.log('[findSimilarTeachers] Target age groups: ' + (targetAgeGroups.length ? targetAgeGroups.join(', ') : 'NONE'));

    // \u2500\u2500 4. Build upskill count map \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    var upskillCountMap = {};
    var tcSheet = _getCachedSheetData(CONFIG.SHEETS.TEACHER_COURSES);
    if (tcSheet && tcSheet.length > 1) {
      var COURSE_START = 4;
      tcSheet.slice(1).forEach(function(row) {
        var rawName = String(row[0] || '').trim();
        if (!rawName || rawName.toLowerCase() === 'teacher') return;
        var canonName = resolveTeacherName(rawName);
        var count = 0;
        for (var i = COURSE_START; i < row.length; i++) {
          var val = String(row[i] || '').trim().toLowerCase();
          if (val && val !== 'not onboarded') count++;
        }
        upskillCountMap[canonName] = count;
      });
    }

    var targetUpskillCount = upskillCountMap[resolvedTarget] || upskillCountMap[normalizeTeacherName(resolvedTarget)] || 0;
    Logger.log('[findSimilarTeachers] Target upskill count: ' + targetUpskillCount);

    // \u2500\u2500 5. Get HubSpot escalations (last 90 days) \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    var escalationMap = getEscalatedTeachersLast90Days();
    Logger.log('[findSimilarTeachers] Escalation map loaded: ' + JSON.stringify(escalationMap));

    // Build audit score map once for all candidates
    var auditMap = {};
    try { auditMap = _buildAuditScoreMap(); } catch(ae) { Logger.log('[findSimilarTeachers] Audit map error: ' + ae.message); }
    Logger.log('[findSimilarTeachers] Audit map loaded for ' + Object.keys(auditMap).length + ' teachers.');

    // Detect Hidden In Search column index
    var hiddenIdx = headers.indexOf('Hidden In Search');

    // \u2500\u2500 6. Score all candidates \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    var candidates = personaData.slice(1).filter(function(r) {
      var rawName = String(r[nameIdx] || '').trim();
      if (!rawName || normalizeTeacherName(rawName) === normalizeTeacherName(resolvedTarget)) return false;
      // Exclude teachers hidden from search results
      if (hiddenIdx > -1 && String(r[hiddenIdx] || '').trim().toLowerCase() === 'yes') return false;
      return true;
    });

    Logger.log('[findSimilarTeachers] Scoring ' + candidates.length + ' candidates');

    var results = candidates.map(function(r) {
      var rawName = String(r[nameIdx]).trim();
      var name    = resolveTeacherName(rawName);

      var candidateTraits = traitCols.reduce(function(acc, i) {
        String(r[i] || '').split(/[,\n]/).forEach(function(t) {
          var clean = t.trim().toLowerCase();
          if (clean) acc.push(clean);
        });
        return acc;
      }, []);

      var candidateAgeGroups = [];
      if (ageGroupIdx > -1) {
        String(r[ageGroupIdx] || '').split(',').forEach(function(a) {
          var clean = a.trim().toLowerCase();
          if (clean) candidateAgeGroups.push(clean);
        });
      }

      var escalations  = escalationMap[name] || escalationMap[rawName] || escalationMap[normalizeTeacherName(name)] || escalationMap[normalizeTeacherName(rawName)] || 0;
      var upskillCount = upskillCountMap[name] || upskillCountMap[rawName] || upskillCountMap[normalizeTeacherName(name)] || upskillCountMap[normalizeTeacherName(rawName)] || 0;

      // Scoring: Traits 30 + Age Group 20 + Upskill Count 30 + Stability 20
      var traitScore = 0;
      if (targetTraits.length > 0) {
        var matched = candidateTraits.filter(function(t) { return targetTraits.indexOf(t) > -1; }).length;
        traitScore  = Math.round((matched / targetTraits.length) * 30);
      } else {
        // \u2705 FIX 3 (Bug 3 continued): If no trait columns were detected at all,
        //    award a neutral partial score (10/30) to all candidates so they
        //    aren't all penalised to 0 and still get meaningful ranking.
        traitScore = 10;
      }

      var ageScore = 0;
      if (targetAgeGroups.length > 0) {
        var ageOverlap = candidateAgeGroups.filter(function(a) { return targetAgeGroups.indexOf(a) > -1; }).length;
        ageScore       = Math.round((ageOverlap / targetAgeGroups.length) * 20);
      } else {
        // \u2705 FIX: If no age group data exists, award a neutral partial score (10/20)
        ageScore = 10;
      }

      var courseScore = 0;
      if (targetUpskillCount > 0) {
        var diff = Math.abs(targetUpskillCount - upskillCount);
        courseScore = Math.max(0, 30 - Math.floor(diff / 5) * 5);
      } else if (upskillCount > 0) {
        courseScore = 15;
      }

      var escalationScore = Math.max(0, 20 - (escalations * 4));

      // Audit score: class score out of 80 → 0-20pts. Red flags lower rank (not disqualify).
      var auditKey      = normalizeTeacherName(name);
      var auditData     = auditMap[auditKey] || auditMap[normalizeTeacherName(rawName)] || null;
      var auditBonus    = auditData && auditData.avgScore != null ? Math.round((auditData.avgScore / 80) * 20) : 10;
      var redFlagPenalty = auditData ? Math.min(auditData.redFlags * 3, 15) : 0;
      var auditScore    = Math.max(0, auditBonus - redFlagPenalty);
      var avgClassScore = auditData ? auditData.avgScore : null;
      var redFlagCount  = auditData ? auditData.redFlags : 0;

      var totalScore    = traitScore + ageScore + courseScore + escalationScore + auditScore;

      var escalationRisk, escalationColor;
      if      (escalations === 0) { escalationRisk = 'No Escalations';              escalationColor = '#15803d'; }
      else if (escalations <= 2)  { escalationRisk = escalations + ' Tickets (Low)';        escalationColor = '#d97706'; }
      else if (escalations <= 4)  { escalationRisk = escalations + ' Tickets (Medium)';     escalationColor = '#d97706'; }
      else                        { escalationRisk = escalations + ' Tickets (High Risk)';  escalationColor = '#b91c1c'; }

      var upskillDiff  = Math.abs(targetUpskillCount - upskillCount);
      var courseOverlap = upskillCount + ' courses upskilled'
        + (upskillDiff === 0 ? ' (exact match)' : ' (diff: ' + upskillDiff + ')');

      var ageGroupMatch = (candidateAgeGroups.length && targetAgeGroups.length)
        ? (candidateAgeGroups.filter(function(a) { return targetAgeGroups.indexOf(a) > -1; }).join(', ') || 'No overlap')
        : 'N/A';

      return {
        name:            name,
        matchScore:      totalScore,
        traitScore:      traitScore,
        ageScore:        ageScore,
        courseScore:     courseScore,
        escalationScore: escalationScore,
        loadScore:       0,
        courseOverlap:   courseOverlap,
        overlapCount:    upskillCount,
        activeLearners:  upskillCount,
        upskillCount:    upskillCount,
        upskillDiff:     upskillDiff,
        ageGroupMatch:   ageGroupMatch,
        escalations:     escalations,
        escalationRisk:  escalationRisk,
        escalationColor: escalationColor,
        auditScore:    auditScore,
        avgClassScore: avgClassScore,
        redFlagCount:  redFlagCount,
        stability: {
          total: escalations,
          risk:  escalations >= 5 ? 'High' : escalations >= 3 ? 'Medium' : 'Stable'
        }
      };
    });

    // Sort by score, take top 10 to send to AI (gives AI enough options)
    var top10 = results
      .sort(function(a, b) { return b.matchScore - a.matchScore; })
      .slice(0, 10);

    var targetContext = {
      name:           resolvedTarget,
      courses:        [targetUpskillCount + ' courses upskilled'],
      activeLearners: targetUpskillCount,
      upskillCount:   targetUpskillCount,
      ageGroups:      targetAgeGroups,
      escalations:    escalationMap[resolvedTarget] || escalationMap[normalizeTeacherName(resolvedTarget)] || 0
    };

    // \u2500\u2500 7. AI re-ranking \u2014 returns top 5 with reasoning \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    Logger.log('[findSimilarTeachers] Sending top 10 to AI for ranking...');
    var aiResult = rankReplacementTeachersWithAI(resolvedTarget, top10, targetContext);

    if (aiResult.success) {
      Logger.log('[findSimilarTeachers] AI ranking successful. Returning AI top 5.');
      return {
        success:       true,
        data:          aiResult.data,
        aiSummary:     aiResult.aiSummary,
        aiEnriched:    true,
        targetContext: targetContext
      };
    }

    // \u2500\u2500 8. Fallback \u2014 return algorithmic top 8 if AI fails \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    Logger.log('[findSimilarTeachers] AI ranking failed (' + aiResult.message + '). Falling back to algorithmic top 8.');
    return {
      success:       true,
      data:          top10.slice(0, 8),
      aiSummary:     '',
      aiEnriched:    false,
      targetContext: targetContext
    };

  } catch (e) {
    Logger.log('[findSimilarTeachers] Error: ' + e.message + '\nStack: ' + e.stack);
    return { success: false, message: e.message };
  }
}



function normalizeTeacherName(name) {
  return String(name || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

 
// \u2500\u2500 Alias map wrapped in a function (required for Google Apps Script) \u2500\u2500
//  Google Apps Script does NOT allow top-level const/let objects that
//  reference nothing \u2014 they throw "not defined" at runtime.
//  Wrapping in a function fixes this completely.
function getTeacherNameAliases() {
  return {
    // \u2500\u2500 Casing fixes (DB lowercase vs HS TitleCase) \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    'aditi chuhan'              : 'Aditi Chauhan',
    'aditi chauhan'             : 'Aditi Chauhan',
    'aditi chauahn'             : 'Aditi Chauhan',
    'aditi chahuan'             : 'Aditi Chauhan',

    // \u2500\u2500 DB name \u2192 different HS full name \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    'betty ann'                 : 'Betty Ann Lijo',        // HS: Betty Ann (exact match in HS sheet)
    'florence bogor'            : 'Florence',   // HS: Florence Bogor (exact match)
    'ramakant chandla'          : 'Ramakant Chandla', // DB all-caps
    'xavier kristeen ottilia'   : 'Ottilia Kristeen Xavier',

    // \u2500\u2500 DB short name \u2192 HS full name \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    'aarshi'                    : 'Aarshi Chaturvedi',
    'anjali'                    : 'Anjali Murali',    // DB "Anjali" = HS "Anjali Murali" (most likely)

    // \u2500\u2500 DB spelling \u2192 HS spelling \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    'akshay gunani'             : 'Akshay Gurnani',   // DB typo, HS correct
    'kim jeoffrey cuevas'       : 'Kim Jeoffrey Cuevass', // HS has double s
    'love sogarwal'             : 'Love Sogarwal',    // normalize casing
    'lovepreetkaur chadha'      : 'LovepreetKaur Chadha',
    'sakshi chillar'            : 'Sakshi Badgujjar', // confirm with team
    'saloni jain'               : 'Saloni Sharma',    // confirm with team
    'soni'                      : 'Akanksha Soni',    // confirm with team
    'komal'                     : 'Komal',
    'sakina jaorawala'          : 'Sakina Jaorawala', // normalize casing

    // \u2500\u2500 Add more as you find them from [ZERO] logs \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    // 'db name lowercase'      : 'Exact HubSpot Display Name',
  };
}
 
function resolveTeacherName(name) {
  const aliases = getTeacherNameAliases();
  const key = normalizeTeacherName(name);
  return aliases[key] || String(name || '').trim();
}
 
function namesMatch(a, b) {
  // Exact normalized match first
  if (normalizeTeacherName(a) === normalizeTeacherName(b)) return true;
  // Fuzzy: remove spaces entirely and compare (catches "chuhan" vs "chauhan" won't help,
  // but at minimum catches casing + spacing issues)
  return false;
}

function getTeacherProfileData(teacherName) {
  try {
    Logger.log('[getTeacherProfileData] Called for: ' + teacherName);

    // \u2500\u2500 1. Basic info from Teacher Data sheet \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    var allTeachers = getTeacherData();
    var teacherInfo = null;
    var nameLower = normalizeTeacherName(teacherName);
    for (var i = 0; i < allTeachers.length; i++) {
      if (normalizeTeacherName(allTeachers[i].name) === nameLower) {
        teacherInfo = allTeachers[i];
        break;
      }
    }

    // \u2500\u2500 2. Course data from Teacher Courses sheet \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    var loadData = getTeacherSpecificLoad(teacherName);
    var courses = (loadData && loadData.success) ? loadData.courses : [];

    // \u2500\u2500 3. Escalation history from HubSpot \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    var escalationData = getTeacherEscalationHistory(teacherName);

    // \u2500\u2500 3. Exact course name \u2192 category mapping \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    var CODING_BEGINNER = [
      'Introduction to Coding (code.org)',
      'Introduction to Coding II (code.org)',
      'Animation with Scratch Jr',
      'Science Adventures with Sprite Lab',
      'Tynker Animation Lab',
      'Minecraft with Tynker',
      'Building blocks of AI (TM)',
      'Game Dev and AI with Scratch',
      'Advanced Scratch',
      'Advanced Sprite Lab',
      'Summer Jam for Foundation',
      'Summer Jam for Pro',
      'Summer Jam for Advanced'
    ];

    var ROBOTICS_AI = [
      'Robotics with Microbit (Jr)',
      'Programming with Robotics (Microbit)',
      'Advanced Microbit',
      'AI with Pictoblox',
      'Immersive AR and VR Modeling (CoSpaces)',
      'App It Up',
      'Advanced Robotics using Makey Makey',
      'Machine Learning Wizkids',
      'ARDUINO'
    ];

    var MINECRAFT_ROBLOX = [
      'Minecraft Edu',
      'Advanced Minecraft with AI and Python',
      'Design with Roblox',
      'Design Pro with Roblox',
      'Design Code and create with Roblox',
      'Unity'
    ];

    var WEB_JS = [
      'Web 3.0',
      'Website Wizardry',
      'Immersive VR Experiences with Javascript',
      'Crack JavaScript Gaming',
      'Advanced Website Engineering'
    ];

    var PYTHON = [
      'Python EduBlocks',
      'Fundamentals of Python with AI',
      'Python 2.0: Beyond the Basics',
      'Python Game Developer',
      'Pro Game Developer in Python',
      'Build GUI with Python',
      'Open CV with Python',
      'Data Structures and Algorithm using Python',
      'Data Science with Python',
      'Machine Learning and Artificial Intelligence with Python',
      'Deep Learning with Python',
      'SQL',
      'GCSE Premium CS Pro',
      'PCEP Certification Prep'
    ];

    var MATHS = [
      'Maths Year 1','Maths Year 2','Maths Year 3','Maths Year 4',
      'Maths Year 5','Maths Year 6','Maths Year 7','Maths Year 8',
      'Maths UK Revision Yr 2','Maths UK Revision Yr 3','Maths UK Revision Yr 4',
      'Maths UK Revision Yr 5','Maths UK Revision Yr 6',
      'Maths UK Year 1','Maths UK Year 2','Maths UK Year 3','Maths UK Year 4',
      'Maths UK Year 5','Maths UK Year 6'
    ];

    var coursesByCategory = {};

    courses.forEach(function(c) {
      var entry = { name: c.course, progress: c.proficiency };
      var cat;
      if (CODING_BEGINNER.indexOf(c.course) > -1)   cat = '\u1F3AE Coding & Game Dev';
      else if (ROBOTICS_AI.indexOf(c.course) > -1)  cat = '\u1F916 Robotics & AI';
      else if (MINECRAFT_ROBLOX.indexOf(c.course) > -1) cat = '\u1F30D Minecraft, Roblox & Unity';
      else if (WEB_JS.indexOf(c.course) > -1)       cat = '\u1F310 Web & JavaScript';
      else if (PYTHON.indexOf(c.course) > -1)        cat = '\u1F40D Python & Data Science';
      else if (MATHS.indexOf(c.course) > -1)         cat = '\u1F4D0 Maths';
      else                                            cat = '\u1F4DA Other';

      if (!coursesByCategory[cat]) coursesByCategory[cat] = [];
      coursesByCategory[cat].push(entry);
    });

    return {
      success: true,
      profile: {
        name:              teacherName,
        email:             teacherInfo ? (teacherInfo.email   || 'N/A') : 'N/A',
        status:            teacherInfo ? (teacherInfo.status  || 'N/A') : (loadData && loadData.status ? loadData.status : 'N/A'),
        manager:           teacherInfo ? (teacherInfo.manager || 'N/A') : 'N/A',
        clsManager:        teacherInfo ? (teacherInfo.clsManagerResponsible || 'N/A') : 'N/A',
        joinDate:          teacherInfo && teacherInfo.joinDate ? new Date(teacherInfo.joinDate).toLocaleDateString('en-GB') : 'N/A',
        totalCourses:      courses.length,
        lastActivity:      loadData ? (loadData.lastActivity || 'N/A') : 'N/A',
        coursesByCategory: coursesByCategory,
        escalation: escalationData && escalationData.success ? {
          totalCount:         escalationData.totalCount,
          byReason:           escalationData.byReason,
          tickets:            escalationData.tickets,
          lastEscalationDate: escalationData.lastEscalationDate
        } : { totalCount: 0, byReason: {}, tickets: [], lastEscalationDate: null }
      }
    };

  } catch (e) {
    Logger.log('[getTeacherProfileData] Error: ' + e.message);
    return { success: false, message: e.message };
  }
}


// ── Audit Intelligence ────────────────────────────────────────────────────────

/**
 * Reads all 4 audit tabs from the audit spreadsheet and returns averaged
 * class score and red-flag count for the given teacher.
 * Class Score is out of 80 (already weighted on the sheet).
 * Red flags lower rank but do NOT disqualify.
 */
function getTeacherAuditData(teacherName) {
  try {
    var normalized = normalizeTeacherName(teacherName);
    var ss = SpreadsheetApp.openById(CONFIG.AUDIT_SHEET_ID);
    var auditTabs = ["Coding Audit'26", "Math Audit'26", "GCSE Audit'26"];
    var scores = [], redFlagTotal = 0, auditRows = [];

    auditTabs.forEach(function(tabName) {
      var sheet = ss.getSheetByName(tabName);
      if (!sheet) return;
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) return;
      var headers = data[0].map(function(h) { return String(h).trim(); });
      var teacherIdx = headers.indexOf('Teacher');
      var scoreIdx   = headers.indexOf('Class Score');
      var rfIdx      = headers.indexOf('Red Flags');
      if (teacherIdx === -1 || scoreIdx === -1) return;

      for (var i = 1; i < data.length; i++) {
        var rowTeacher = normalizeTeacherName(String(data[i][teacherIdx] || ''));
        if (!rowTeacher || rowTeacher !== normalized) continue;
        var score = parseFloat(data[i][scoreIdx]) || 0;
        if (score > 0) scores.push(score);
        if (rfIdx > -1) {
          var rf = String(data[i][rfIdx] || '').trim();
          if (rf && rf.toLowerCase() !== '' && rf.toLowerCase() !== 'none' && rf.toLowerCase() !== 'no' && rf !== '-') {
            redFlagTotal++;
          }
        }
        auditRows.push({ tab: tabName, score: score, redFlags: rfIdx > -1 ? String(data[i][rfIdx] || '').trim() : '' });
      }
    });

    var avgScore = scores.length > 0
      ? Math.round(scores.reduce(function(a, b) { return a + b; }, 0) / scores.length * 10) / 10
      : null;

    return {
      success:    true,
      hasData:    scores.length > 0,
      avgScore:   avgScore,
      redFlags:   redFlagTotal,
      auditCount: auditRows.length,
      auditRows:  auditRows
    };
  } catch (e) {
    Logger.log('[getTeacherAuditData] Error for "' + teacherName + '": ' + e.message);
    return { success: false, hasData: false, avgScore: null, redFlags: 0, auditCount: 0, auditRows: [] };
  }
}

/**
 * Builds a map of { normalizedTeacherName: { avgScore, redFlags } }
 * by reading all 4 audit tabs once — used inside findSimilarTeachers.
 */
function _buildAuditScoreMap() {
  var map = {};
  try {
    var ss = SpreadsheetApp.openById(CONFIG.AUDIT_SHEET_ID);
    var auditTabs = ["Coding Audit'26", "Math Audit'26", "GCSE Audit'26"];

    auditTabs.forEach(function(tabName) {
      var sheet = ss.getSheetByName(tabName);
      if (!sheet) return;
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) return;
      var headers = data[0].map(function(h) { return String(h).trim(); });
      var teacherIdx = headers.indexOf('Teacher');
      var scoreIdx   = headers.indexOf('Class Score');
      var rfIdx      = headers.indexOf('Red Flags');
      if (teacherIdx === -1 || scoreIdx === -1) return;

      for (var i = 1; i < data.length; i++) {
        var rawTeacher = String(data[i][teacherIdx] || '').trim();
        if (!rawTeacher) continue;
        var key   = normalizeTeacherName(rawTeacher);
        var score = parseFloat(data[i][scoreIdx]) || 0;
        if (score <= 0) continue;
        if (!map[key]) map[key] = { scores: [], redFlags: 0 };
        map[key].scores.push(score);
        if (rfIdx > -1) {
          var rf = String(data[i][rfIdx] || '').trim();
          if (rf && rf.toLowerCase() !== 'none' && rf.toLowerCase() !== 'no' && rf !== '-' && rf !== '') {
            map[key].redFlags++;
          }
        }
      }
    });

    // Collapse scores array to avgScore
    Object.keys(map).forEach(function(k) {
      var s = map[k].scores;
      map[k].avgScore = s.length > 0
        ? Math.round(s.reduce(function(a, b) { return a + b; }, 0) / s.length * 10) / 10
        : null;
      delete map[k].scores;
    });
  } catch (e) {
    Logger.log('[_buildAuditScoreMap] Error: ' + e.message);
  }
  return map;
}

/**
 * Toggle teacher visibility in search results.
 * Stores "Hidden In Search" = "Yes" (or blank) in the Persona sheet.
 */
function setTeacherVisibility(teacherName, isHidden) {
  return updateTeacherPersona({
    'Teacher Name':     teacherName,
    'Hidden In Search': isHidden ? 'Yes' : ''
  });
}

/**
 * Reads the "Migration Teacher" tab from AUDIT_SHEET_ID.
 * Returns a map: { normalizedTeacherName: [available time slot strings] }
 * for the specific requested date column.
 * Date columns are formatted like "Fri, 3-Apr-26", "Sat, 4-Apr-26".
 */
function _getTeacherAvailabilityMap(requestedDateStr) {
  var map = {};
  try {
    var ss = SpreadsheetApp.openById(CONFIG.AUDIT_SHEET_ID);
    var sheet = ss.getSheetByName('Migration Teacher');
    if (!sheet) {
      Logger.log('[_getTeacherAvailabilityMap] Migration Teacher tab not found in audit sheet');
      return map;
    }
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return map;

    var headers = data[0].map(function(h) { return String(h).trim(); });
    var teacherIdx = 0; // Column A = Teacher

    // Headers like "Fri, 3-Apr-26" — new Date() can't parse these reliably in GAS.
    var MONTH_MAP = { jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12 };
    function _parseSheetDateHeader(h) {
      var clean = h.replace(/^[A-Za-z]+,?\s*/, '').trim();
      var parts = clean.split('-');
      if (parts.length < 3) return null;
      var day = parseInt(parts[0], 10);
      var mon = MONTH_MAP[parts[1].toLowerCase().substring(0,3)];
      var yr  = parseInt(parts[2], 10);
      if (isNaN(day) || !mon || isNaN(yr)) return null;
      var fy = yr < 100 ? 2000 + yr : yr;
      return fy + '-' + String(mon).padStart(2,'0') + '-' + String(day).padStart(2,'0');
    }
    var targetColIdx = -1;
    for (var ci = 0; ci < headers.length; ci++) {
      var h = headers[ci];
      if (!h || h === 'Teacher' || h === 'Calendar ID') continue;
      if (_parseSheetDateHeader(String(h).trim()) === requestedDateStr) { targetColIdx = ci; break; }
    }
    Logger.log('[avail] date=' + requestedDateStr + ' col=' + targetColIdx + (targetColIdx > -1 ? ' (' + headers[targetColIdx] + ')' : ' NOT FOUND'));

    for (var ri = 1; ri < data.length; ri++) {
      var teacherName = String(data[ri][teacherIdx] || '').trim();
      if (!teacherName) continue;
      var key = normalizeTeacherName(teacherName);
      var slots = [];
      if (targetColIdx > -1) {
        var cell = String(data[ri][targetColIdx] || '').trim();
        if (cell && cell !== 'No Slots' && cell !== 'No Slots Available') {
          // Cells use newlines (Alt+Enter) and/or commas between multiple slots
          slots = cell.split(/[\n,]/).map(function(s) { return s.trim(); }).filter(Boolean);
        }
      }
      map[key] = slots;
    }
  } catch (e) {
    Logger.log('[_getTeacherAvailabilityMap] Error: ' + e.message);
  }
  return map;
}

/**
 * Builds audit score map using only rows within the last N days.
 * Uses date column (searches for "Date", "Audit Date", "Class Date" header).
 * Falls back to all rows if no date column found.
 */
function _buildAuditScoreMapDays(days) {
  var map = {};
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  try {
    var ss = SpreadsheetApp.openById(CONFIG.AUDIT_SHEET_ID);
    var auditTabs = ["Coding Audit'26", "Math Audit'26", "GCSE Audit'26"];

    auditTabs.forEach(function(tabName) {
      var sheet = ss.getSheetByName(tabName);
      if (!sheet) return;
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) return;
      var headers = data[0].map(function(h) { return String(h).trim(); });
      var teacherIdx = headers.indexOf('Teacher');
      var scoreIdx   = headers.indexOf('Class Score');
      var rfIdx      = headers.indexOf('Red Flags');
      // Find date column
      var dateIdx = -1;
      ['Date', 'Audit Date', 'Class Date', 'Session Date', 'Observation Date'].forEach(function(n) {
        if (dateIdx === -1 && headers.indexOf(n) > -1) dateIdx = headers.indexOf(n);
      });
      if (teacherIdx === -1 || scoreIdx === -1) return;

      for (var i = 1; i < data.length; i++) {
        // Date filter
        if (dateIdx > -1) {
          var rowDate = data[i][dateIdx];
          if (rowDate) {
            var d = new Date(rowDate);
            if (!isNaN(d.getTime()) && d < cutoff) continue;
          }
        }
        var rawTeacher = String(data[i][teacherIdx] || '').trim();
        if (!rawTeacher) continue;
        var key   = normalizeTeacherName(rawTeacher);
        var score = parseFloat(data[i][scoreIdx]) || 0;
        if (score <= 0) continue;
        if (!map[key]) map[key] = { scores: [], redFlags: 0 };
        map[key].scores.push(score);
        if (rfIdx > -1) {
          var rf = String(data[i][rfIdx] || '').trim();
          if (rf && rf.toLowerCase() !== 'none' && rf.toLowerCase() !== 'no' && rf !== '-' && rf !== '') {
            map[key].redFlags++;
          }
        }
      }
    });

    Object.keys(map).forEach(function(k) {
      var s = map[k].scores;
      map[k].avgScore = s.length > 0
        ? Math.round(s.reduce(function(a, b) { return a + b; }, 0) / s.length * 10) / 10
        : null;
      map[k].auditCount = s.length;
      delete map[k].scores;
    });
  } catch (e) {
    Logger.log('[_buildAuditScoreMapDays] Error: ' + e.message);
  }
  return map;
}