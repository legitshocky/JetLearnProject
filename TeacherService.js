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
        hiddenInSearch: String(row[11] || '').trim(),
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

function getCourseTeacherDetails(courseName) {
  Logger.log('getCourseTeacherDetails called for: ' + courseName);
  try {
    // ── 1. Read wide-format TEACHER_COURSES sheet ──────────────────────────
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.TEACHER_COURSES);
    if (!sheetData || sheetData.length < 2) {
      return { success: false, message: 'No teacher data found.', teachers: [] };
    }
 
    // Detect header row
    let headerRowIndex = 0;
    for (let i = 0; i < Math.min(sheetData.length, 10); i++) {
      if (String(sheetData[i][0]).trim().toLowerCase() === 'teacher') {
        headerRowIndex = i;
        break;
      }
    }
    const headers     = sheetData[headerRowIndex];
    const COURSE_START = 4;
 
    // Find the column index for the requested course (case-insensitive)
    let courseColIndex = -1;
    for (let i = COURSE_START; i < headers.length; i++) {
      if (String(headers[i] || '').trim().toLowerCase() === courseName.trim().toLowerCase()) {
        courseColIndex = i;
        break;
      }
    }
 
    if (courseColIndex === -1) {
      Logger.log('getCourseTeacherDetails: course column not found for "' + courseName + '"');
      // Return empty but successful — course exists but no one upskilled yet
      return { success: true, courseName: courseName, teachers: [] };
    }
 
    // Build proficiency map for all teachers upskilled on this course
    const proficiencyMap = {}; // teacherName -> proficiency string
    sheetData.slice(headerRowIndex + 1).forEach(function(row) {
      const teacher     = String(row[0] || '').trim();
      const proficiency = String(row[courseColIndex] || '').trim();
      if (!teacher || !proficiency || proficiency.toLowerCase() === 'not onboarded') return;
      proficiencyMap[teacher] = proficiency;
    });
 
    if (Object.keys(proficiencyMap).length === 0) {
      return { success: true, courseName: courseName, teachers: [] };
    }
 
    // ── 2. Fetch HubSpot active learner counts ─────────────────────────────
    //    One paginated call gives us both totalLearners and learnersOnThisCourse
    const token          = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    const searchUrl      = 'https://api.hubapi.com/crm/v3/objects/deals/search';
    const activeStatuses = ['Active Learner', 'Friendly Learner', 'VIP', 'Break & Return'];
 
    const teacherTotals   = {}; // normalised key -> total active learners
    const teacherOnCourse = {}; // normalised key -> learners on THIS course
    const teacherLearners = {}; // normalised key -> [ {name, daysOnCourse, status} ]
    const nowMs = new Date().getTime();

    let after = undefined, page = 0;
    do {
      var body = {
        filterGroups: [{
          filters: [{
            propertyName: 'learner_status',
            operator:     'IN',
            values:       activeStatuses
          }]
        }],
        properties: ['current_teacher', 'current_course', 'dealname', 'createdate', 'learner_status'],
        limit: 100
      };
      if (after) body.after = after;

      var resp = UrlFetchApp.fetch(searchUrl, {
        method:             'post',
        headers:            { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload:            JSON.stringify(body),
        muteHttpExceptions: true
      });
      var data = JSON.parse(resp.getContentText());
      if (!data.results) break;

      data.results.forEach(function(deal) {
        var rawTeacher = deal.properties.current_teacher || '';
        var rawCourse  = deal.properties.current_course  || '';
        var teacher    = getTeacherLabel(rawTeacher);
        if (!teacher) return;

        var teacherKey  = teacher.trim().toLowerCase().replace(/\s+/g, ' ');
        var courseLabel = getCourseLabel(rawCourse);

        // Total active learners for this teacher
        if (!teacherTotals[teacherKey]) teacherTotals[teacherKey] = 0;
        teacherTotals[teacherKey]++;

        // If this learner is on the selected course — collect details
        if (courseLabel && courseLabel.trim().toLowerCase() === courseName.trim().toLowerCase()) {
          if (!teacherOnCourse[teacherKey])  teacherOnCourse[teacherKey]  = 0;
          if (!teacherLearners[teacherKey])  teacherLearners[teacherKey]  = [];
          teacherOnCourse[teacherKey]++;

          // Days on course: use createdate as proxy
          var createdMs   = deal.properties.createdate ? new Date(deal.properties.createdate).getTime() : null;
          var daysOnCourse = createdMs ? Math.floor((nowMs - createdMs) / 86400000) : null;

          teacherLearners[teacherKey].push({
            name:         (deal.properties.dealname || 'Learner').trim(),
            daysOnCourse: daysOnCourse,
            status:       (deal.properties.learner_status || '').trim()
          });
        }
      });

      after = data.paging && data.paging.next ? data.paging.next.after : undefined;
      page++;
    } while (after && page < 30);

    // ── 3. Build result array ──────────────────────────────────────────────
    var teachers = Object.keys(proficiencyMap).map(function(teacherName) {
      var key      = teacherName.trim().toLowerCase().replace(/\s+/g, ' ');
      var total    = teacherTotals[key]   || 0;
      var onCourse = teacherOnCourse[key] || 0;
      var learners = teacherLearners[key] || [];
      // Sort learners: longest tenure first
      learners.sort(function(a, b) { return (b.daysOnCourse || 0) - (a.daysOnCourse || 0); });

      return {
        name:             teacherName,
        proficiency:      proficiencyMap[teacherName],
        totalLearners:    total,
        learnersOnCourse: onCourse,
        learners:         learners
      };
    });
 
    // Sort: most learners on this course first, then highest proficiency
    teachers.sort(function(a, b) {
      var diff = b.learnersOnCourse - a.learnersOnCourse;
      if (diff !== 0) return diff;
      return (parseFloat(b.proficiency) || 0) - (parseFloat(a.proficiency) || 0);
    });
 
    Logger.log('getCourseTeacherDetails: ' + teachers.length + ' teachers found for ' + courseName);
    return { success: true, courseName: courseName, teachers: teachers };
 
  } catch (e) {
    Logger.log('getCourseTeacherDetails error: ' + e.message);
    return { success: false, message: e.message, teachers: [] };
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

    // Load hidden teachers from Script Properties and merge with Persona sheet
    const hiddenTeachersMap = getHiddenTeachersMap();


    // ── Step 3b: Build Persona Traits map ──────────────────────────
    const traitsMap = {};
    try {
      const personaTraitData = _getCachedSheetData('Teacher Persona Mapping');
      if (personaTraitData && personaTraitData.length > 1) {
        const ptHeaders = personaTraitData[0].map(h => String(h).trim());
        const ptNameCol = ptHeaders.indexOf('Teacher Name');
        const traitCols = [];
        ptHeaders.forEach((h, i) => {
          if (/trait|expertise|style|skill|strength|personality|subject|teaching/i.test(h)) traitCols.push(i);
        });
        personaTraitData.slice(1).forEach(row => {
          const name = String(row[ptNameCol] || '').trim().toLowerCase();
          if (!name) return;
          const traits = [];
          traitCols.forEach(i => { const t = String(row[i] || '').trim(); if (t) traits.push(t); });
          traitsMap[name] = traits;
        });
      }
    } catch(e) { Logger.log('[getTeacherDetailsForTable] traitsMap error: ' + e.message); }

    // ── Step 3c: Build latest audit grade map (most-recent score, all-time) ──────────────────────
    // Using all-time latest instead of 45-day window so grades show for teachers
    // who haven't been audited recently (audits are periodic, not daily).
    let latestGradeMap = {};
    try {
      const allAuditRecs = _getAllAuditRecords();
      allAuditRecs.forEach(function(a) {
        if (!a.teacher || a.classScore === null || a.classScore === undefined) return;
        const key = normalizeTeacherName(a.teacher);
        if (!latestGradeMap[key] || (a.date && (!latestGradeMap[key].date || a.date > latestGradeMap[key].date))) {
          latestGradeMap[key] = { score: a.classScore, date: a.date };
        }
      });
    } catch(ae) {
      Logger.log('[getTeacherDetailsForTable] latestGradeMap error: ' + ae.message);
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
        hiddenInSearch: personaEntry.hiddenInSearch || teacher.hiddenInSearch || '',

        // Get last activity from the audit log
        lastActivity:  getTeacherLastActivity(teacherName),

        // Traits & Audit grade — fed into allTeacherDetailsCache on the frontend
        traits:     traitsMap[teacherNorm] || traitsMap[resolvedNorm] || [],
        auditGrade: (() => {
          const ad = latestGradeMap[teacherNorm] || latestGradeMap[resolvedNorm] || null;
          if (!ad || ad.score === null || ad.score === undefined) return '—';
          return _auditGrade(ad.score);
        })()
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
    // Get the master list of course names from the Course Name sheet
    const courseNames = getCourseNames();
 
    // Count how many teachers are upskilled on each course
    // using the wide-format TEACHER_COURSES sheet
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.TEACHER_COURSES);
    const courseTeacherCount = {};
 
    if (sheetData && sheetData.length > 1) {
      // Detect header row (same pattern used across persona engine)
      let headerRowIndex = 0;
      for (let i = 0; i < Math.min(sheetData.length, 10); i++) {
        if (String(sheetData[i][0]).trim().toLowerCase() === 'teacher') {
          headerRowIndex = i;
          break;
        }
      }
      const headers = sheetData[headerRowIndex];
      const COURSE_START = 4; // col0=Teacher, col1=Email, col2=Manager, col3=Health
 
      sheetData.slice(headerRowIndex + 1).forEach(function(row) {
        const teacherName = String(row[0] || '').trim();
        if (!teacherName) return;
        for (let i = COURSE_START; i < headers.length; i++) {
          const courseName = String(headers[i] || '').trim();
          const val        = String(row[i]   || '').trim();
          if (!courseName || !val || val.toLowerCase() === 'not onboarded') continue;
          courseTeacherCount[courseName] = (courseTeacherCount[courseName] || 0) + 1;
        }
      });
    }
 
    return courseNames.map(function(name) {
      return {
        name:               name,
        activeTeachersCount: courseTeacherCount[name] || 0
      };
    });
 
  } catch (e) {
    Logger.log('getCourseDetails error: ' + e.message);
    return [];
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
    var learnerAgeNum = parseInt(learnerAge, 10);
    var isMathCourse  = currentCourse.toLowerCase().indexOf('math') !== -1;

    // ── 2. Slot normaliser ──
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

    // ── 3. Load Migration Teacher sheet as PRIMARY teacher list + slot data ──
    var MONTH_MAP = { jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12 };
    function parseMigDateHeader(h) {
      var clean = h.replace(/^[A-Za-z]+,?\s*/, '').trim();
      var parts  = clean.split('-');
      if (parts.length < 3) return null;
      var day = parseInt(parts[0], 10);
      var mon = MONTH_MAP[parts[1].toLowerCase().substring(0,3)];
      var yr  = parseInt(parts[2], 10);
      if (isNaN(day) || !mon || isNaN(yr)) return null;
      var fy = yr < 100 ? 2000 + yr : yr;
      return fy + '-' + String(mon).padStart(2,'0') + '-' + String(day).padStart(2,'0');
    }

    var migSS    = SpreadsheetApp.openById(CONFIG.AUDIT_SHEET_ID);
    var migSheet = migSS.getSheetByName('Migration Teacher');
    if (!migSheet) return { success: false, message: 'Migration Teacher tab not found in audit spreadsheet.' };

    // Use getValues() for slot cell data, getDisplayValues() for headers (date headers
    // may be stored as actual Date objects in the sheet — getDisplayValues() returns
    // the visible text like "Mon, 6-Apr-26" regardless of underlying type)
    var migRange        = migSheet.getDataRange();
    var migData         = migRange.getValues();        // raw values for slot cells
    var migDataDisplay  = migRange.getDisplayValues(); // display strings for header parsing
    if (migData.length < 2) return { success: true, results: [] };

    // Find the header row (first row where col A display = "Teacher")
    var migHeaderRowIdx = 0;
    for (var hri = 0; hri < Math.min(migDataDisplay.length, 5); hri++) {
      if (String(migDataDisplay[hri][0] || '').trim().toLowerCase() === 'teacher') {
        migHeaderRowIdx = hri; break;
      }
    }
    var migHeaderDisplay = migDataDisplay[migHeaderRowIdx]; // display strings for all header cells
    var migDataRows      = migData.slice(migHeaderRowIdx + 1); // raw data rows
    Logger.log('[SMT] migHeaderRowIdx=' + migHeaderRowIdx
      + ' displayHeaders[0..6]=' + migHeaderDisplay.slice(0,7).join(' | '));

    // Map each requested date (YYYY-MM-DD) → column index
    // Handles both text headers like "Mon, 6-Apr-26" and Date-formatted headers
    function headerToYMD(h) {
      var s = String(h || '').trim();
      if (!s) return null;
      // Try MONTH_MAP text format first: "Mon, 6-Apr-26"
      var clean = s.replace(/^[A-Za-z]+,?\s*/, '').trim();
      var parts  = clean.split('-');
      if (parts.length === 3) {
        var day = parseInt(parts[0], 10);
        var mon = MONTH_MAP[parts[1].toLowerCase().substring(0, 3)];
        var yr  = parseInt(parts[2], 10);
        if (!isNaN(day) && mon && !isNaN(yr)) {
          var fy = yr < 100 ? 2000 + yr : yr;
          return fy + '-' + String(mon).padStart(2, '0') + '-' + String(day).padStart(2, '0');
        }
      }
      // Try "d/m/yyyy" or "m/d/yyyy" display format Google sometimes uses
      var slash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (slash) {
        // Ambiguous — assume d/m/yyyy (European) matching GAS locale
        var dd = parseInt(slash[1], 10), mm = parseInt(slash[2], 10), yyyy = parseInt(slash[3], 10);
        if (dd > 12) return yyyy + '-' + String(mm).padStart(2,'0') + '-' + String(dd).padStart(2,'0');
        return yyyy + '-' + String(dd).padStart(2,'0') + '-' + String(mm).padStart(2,'0');
      }
      return null;
    }

    var slotColMap = {}; // dateStr (YYYY-MM-DD) → colIdx
    requestedSlots.forEach(function(rs) {
      if (slotColMap[rs.date] !== undefined) return;
      for (var ci = 1; ci < migHeaderDisplay.length; ci++) {
        if (headerToYMD(migHeaderDisplay[ci]) === rs.date) {
          slotColMap[rs.date] = ci;
          break;
        }
      }
    });
    Logger.log('[SMT] slotColMap: ' + JSON.stringify(slotColMap)
      + ' | requested dates: ' + requestedSlots.map(function(r){ return r.date; }).join(', '));

    // ── Detect "Hidden In Search" column in Migration Teacher sheet ──
    // Add a column named "Hidden In Search" to the sheet and set "Yes" to exclude a teacher
    var hiddenColIdx = -1;
    for (var hci = 0; hci < migHeaderDisplay.length; hci++) {
      var hh = String(migHeaderDisplay[hci] || '').trim().toLowerCase();
      if (hh === 'hidden in search' || hh === 'hidden' || hh === 'exclude') {
        hiddenColIdx = hci; break;
      }
    }
    Logger.log('[SMT] hiddenColIdx=' + hiddenColIdx);

    // ── 3b. Build calendarId map from Migration Teacher col B ──
    var calendarIdMap = {}; // normalizedName → calendarId (email)
    migDataRows.forEach(function(row) {
      var n   = String(row[0] || '').trim();
      var cid = String(row[1] || '').trim();
      if (n && cid && cid.indexOf('@') > -1) {
        calendarIdMap[normalizeTeacherName(n)] = cid;
      }
    });
    Logger.log('[SMT] calendarIdMap: ' + Object.keys(calendarIdMap).length + ' teachers with calendar IDs');

    // ── 3c. Build name → jetlearn email map + hidden set from Teacher Data ──
    var teacherEmailMap = {}; // normalizedName → jetlearn email (row[8])
    var hiddenTeacherSet = {};
    try {
      var tdRows = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA);
      tdRows.slice(1).forEach(function(row) {
        var n  = String(row[1] || '').trim();
        var em = String(row[8] || '').trim().toLowerCase();
        if (n && em) teacherEmailMap[normalizeTeacherName(n)] = em;
        // Build hidden set from col L (index 11 = "Hide in Search")
        var hideFlag = String(row[11] || '').trim().toLowerCase();
        if (n && hideFlag === 'yes') hiddenTeacherSet[normalizeTeacherName(n)] = true;
      });
    } catch(tde) { Logger.log('[SMT] teacherEmailMap error: ' + tde.message); }
    Logger.log('[SMT] teacherEmailMap: ' + Object.keys(teacherEmailMap).length + ' teachers | hiddenSet: ' + Object.keys(hiddenTeacherSet).length);

    // ── Calendar slot helpers ──
    var _calSlotCache  = {}; // key: calId+'_'+dateStr → [{startMs,endMs}] or null (error)
    var AVAIL_KEYWORDS = ['availability', 'available hours', 'teaching hours'];
    var SLOT_HOUR_MS   = 60 * 60 * 1000;

    // Cache master calendar events per date (shared across all teacher lookups)
    var _masterEvtCache = {}; // dateStr → [{startMs,endMs,guests:[email,...]}]
    function getMasterEventsForDate(dateStr) {
      if (_masterEvtCache[dateStr] !== undefined) return _masterEvtCache[dateStr];
      try {
        var masterCal = CalendarApp.getCalendarById(CONFIG.CLASS_SCHEDULE_CALENDAR_ID);
        if (!masterCal) { _masterEvtCache[dateStr] = []; return []; }
        var dp = dateStr.split('-');
        var dayDate = new Date(parseInt(dp[0]), parseInt(dp[1]) - 1, parseInt(dp[2]));
        var evts = masterCal.getEventsForDay(dayDate);
        _masterEvtCache[dateStr] = evts.map(function(ev) {
          var guests = [];
          try { ev.getGuestList().forEach(function(g) { guests.push(g.getEmail().toLowerCase()); }); } catch(e) {}
          return { startMs: ev.getStartTime().getTime(), endMs: ev.getEndTime().getTime(), guests: guests, title: ev.getTitle().toLowerCase() };
        });
      } catch(e) {
        Logger.log('[masterCal] Error fetching ' + dateStr + ': ' + e.message);
        _masterEvtCache[dateStr] = [];
      }
      return _masterEvtCache[dateStr];
    }

    // Fetch and cache available 1-hour slots for a teacher's calendar on a given date.
    // Returns [] if no availability events; null if calendar inaccessible (don't penalise teacher).
    var _scriptCache = CacheService.getScriptCache();
    var CAL_CACHE_VER = 'v5'; // bump to invalidate stale calendar cache
    function fetchTeacherCalendarSlots(calId, dateStr, teacherName, teacherEmail) {
      var ck = CAL_CACHE_VER + '_' + calId + '_' + dateStr;
      if (_calSlotCache[ck] !== undefined) return _calSlotCache[ck];
      // Check persistent script cache first (avoids repeat API calls within 1 hour)
      try {
        var cached = _scriptCache.get(ck);
        if (cached !== null) {
          var parsed = cached === '__null__' ? null : JSON.parse(cached);
          _calSlotCache[ck] = parsed;
          return parsed;
        }
      } catch(e) {}
      try {
        var cal = CalendarApp.getCalendarById(calId);
        if (!cal) {
          _calSlotCache[ck] = null;
          try { _scriptCache.put(ck, '__null__', 3600); } catch(e) {}
          return null;
        }
        var dp      = dateStr.split('-');
        var dayDate = new Date(parseInt(dp[0]), parseInt(dp[1]) - 1, parseInt(dp[2]));
        var allEvts = cal.getEventsForDay(dayDate);
        // Availability events — timed only (skip all-day banners; they don't represent real open hours)
        var availEvts = allEvts.filter(function(e) {
          if (e.isAllDayEvent()) return false;
          var t = e.getTitle().toLowerCase();
          return AVAIL_KEYWORDS.some(function(kw) { return t.indexOf(kw) > -1; });
        });
        // All other events on teacher's own calendar = blocked time
        var blockedEvts = allEvts.filter(function(e) {
          var t = e.getTitle().toLowerCase();
          return !AVAIL_KEYWORDS.some(function(kw) { return t.indexOf(kw) > -1; });
        });
        // Also block any master-calendar events where this teacher is an attendee
        // OR where the teacher's name appears in the event title
        // (class bookings live on hello@jet-learn.com, teacher often referenced by name not email)
        var masterEvts = getMasterEventsForDate(dateStr);
        var calIdLower = calId.toLowerCase();
        // Block if: teacher's calendar ID is a guest, OR jetlearn email is a guest, OR full name in title
        var teacherEmailLower = (teacherEmail || '').toLowerCase();
        var nameParts = (teacherName || '').toLowerCase().split(/\s+/).filter(function(p) { return p.length >= 3; });
        masterEvts.forEach(function(mev) {
          var isGuest    = mev.guests.indexOf(calIdLower) > -1
                        || (teacherEmailLower && mev.guests.indexOf(teacherEmailLower) > -1);
          var isInTitle  = nameParts.length > 0 && nameParts.every(function(part) { return mev.title.indexOf(part) > -1; });
          if (isGuest || isInTitle) {
            blockedEvts.push({ getStartTime: function(){ return { getTime: function(){ return mev.startMs; } }; },
                               getEndTime:   function(){ return { getTime: function(){ return mev.endMs;   } }; } });
          }
        });
        // Split each availability block into 1-hour chunks, clipped to the requested date in CET.
        // This prevents multi-day timed events (e.g. Sat 5pm → Sun 5pm) from generating
        // slots outside the teacher's actual working hours on the requested date.
        var dp2 = dateStr.split('-');
        var yr2 = parseInt(dp2[0]), mo2 = parseInt(dp2[1]), dy2 = parseInt(dp2[2]);
        var cetOff2    = (mo2 >= 4 && mo2 <= 9) ? 2 : 1;
        var dayStartMs = Date.UTC(yr2, mo2 - 1, dy2,     -cetOff2, 0, 0); // midnight CET on dateStr
        var dayEndMs   = Date.UTC(yr2, mo2 - 1, dy2 + 1, -cetOff2, 0, 0); // midnight CET next day
        var freeSlots = [];
        availEvts.forEach(function(av) {
          var s = Math.max(av.getStartTime().getTime(), dayStartMs);
          var e = Math.min(av.getEndTime().getTime(),   dayEndMs);
          for (var cur = s; cur + SLOT_HOUR_MS <= e; cur += SLOT_HOUR_MS) {
            freeSlots.push({ startMs: cur, endMs: cur + SLOT_HOUR_MS });
          }
        });
        // Subtract all blocked events (teacher's own + master calendar classes)
        var available = freeSlots.filter(function(slot) {
          return !blockedEvts.some(function(ev) {
            return slot.startMs < ev.getEndTime().getTime() && slot.endMs > ev.getStartTime().getTime();
          });
        });
        Logger.log('[cal] ' + calId + ' on ' + dateStr + ': ' + availEvts.length + ' avail evts, ' + blockedEvts.length + ' blocked, ' + available.length + ' free slots');
        _calSlotCache[ck] = available;
        try { _scriptCache.put(ck, JSON.stringify(available), 3600); } catch(e) {}
        return available;
      } catch(ce) {
        Logger.log('[cal] Error for ' + calId + ': ' + ce.message);
        _calSlotCache[ck] = null;
        return null;
      }
    }

    // Batch-prefetch multiple teacher calendars in parallel via Calendar REST API.
    // Populates _calSlotCache before the main teacher loop so each fetchTeacherCalendarSlots
    // call is a fast in-memory hit instead of a sequential HTTP round-trip.
    function prefetchCalendarsBatch(pairs) {
      // pairs: [{calId, dateStr, teacherName}] — de-dupe and skip already-cached
      var seen = {};
      var toFetch = [];
      pairs.forEach(function(p) {
        var ck = CAL_CACHE_VER + '_' + p.calId + '_' + p.dateStr;
        var key = p.calId + '_' + p.dateStr;
        if (seen[key]) return;
        seen[key] = true;
        if (_calSlotCache[ck] !== undefined) return;
        try {
          var hit = _scriptCache.get(ck);
          if (hit !== null) { _calSlotCache[ck] = hit === '__null__' ? null : JSON.parse(hit); return; }
        } catch(e) {}
        toFetch.push(p);
      });
      if (toFetch.length === 0) return;

      // Prefetch master calendar for all unique dates first (single sequential call per date)
      var uniqueDates = [];
      toFetch.forEach(function(p) { if (uniqueDates.indexOf(p.dateStr) === -1) uniqueDates.push(p.dateStr); });
      uniqueDates.forEach(function(ds) { getMasterEventsForDate(ds); });

      // Build parallel Calendar REST API requests
      var token = ScriptApp.getOAuthToken();
      var requests = toFetch.map(function(p) {
        var dp = p.dateStr.split('-');
        var yr = parseInt(dp[0]), mo = parseInt(dp[1]), dy = parseInt(dp[2]);
        var cetOff = (mo >= 4 && mo <= 9) ? 2 : 1;
        var tMin = new Date(Date.UTC(yr, mo-1, dy,   -cetOff, 0, 0)).toISOString();
        var tMax = new Date(Date.UTC(yr, mo-1, dy+1, -cetOff, 0, 0)).toISOString();
        return {
          url: 'https://www.googleapis.com/calendar/v3/calendars/' + encodeURIComponent(p.calId)
             + '/events?singleEvents=true&timeMin=' + encodeURIComponent(tMin)
             + '&timeMax=' + encodeURIComponent(tMax),
          headers: { Authorization: 'Bearer ' + token },
          muteHttpExceptions: true
        };
      });

      var responses;
      try { responses = UrlFetchApp.fetchAll(requests); }
      catch(fe) { Logger.log('[prefetch] fetchAll error: ' + fe.message); return; }

      responses.forEach(function(resp, idx) {
        var p  = toFetch[idx];
        var ck = CAL_CACHE_VER + '_' + p.calId + '_' + p.dateStr;
        if (resp.getResponseCode() !== 200) {
          _calSlotCache[ck] = null;
          try { _scriptCache.put(ck, '__null__', 3600); } catch(e) {}
          return;
        }
        try {
          var items = (JSON.parse(resp.getContentText()).items || []);
          var availEvts = [], blockedEvts = [];
          items.forEach(function(item) {
            if (item.status === 'cancelled') return;
            var title   = (item.summary || '').toLowerCase();
            var isAllDay = !!(item.start && item.start.date && !item.start.dateTime);
            if (isAllDay) return; // skip all-day banners
            var sMs = new Date(item.start.dateTime).getTime();
            var eMs = new Date(item.end.dateTime).getTime();
            if (AVAIL_KEYWORDS.some(function(kw) { return title.indexOf(kw) > -1; })) {
              availEvts.push({ startMs: sMs, endMs: eMs });
            } else {
              blockedEvts.push({ startMs: sMs, endMs: eMs });
            }
          });
          // Block master-calendar events matching by guest email OR teacher name in title
          var masterEvts  = getMasterEventsForDate(p.dateStr);
          var calIdLower      = p.calId.toLowerCase();
          var tEmailLower     = (p.teacherEmail || '').toLowerCase();
          var nameParts       = (p.teacherName || '').toLowerCase().split(/\s+/).filter(function(pt) { return pt.length >= 3; });
          masterEvts.forEach(function(mev) {
            var isGuest   = mev.guests.indexOf(calIdLower) > -1
                         || (tEmailLower && mev.guests.indexOf(tEmailLower) > -1);
            var isInTitle = nameParts.length > 0 && nameParts.every(function(part) { return mev.title.indexOf(part) > -1; });
            if (isGuest || isInTitle) blockedEvts.push({ startMs: mev.startMs, endMs: mev.endMs });
          });
          // Clip to CET day and enumerate 1-hour slots
          var dp2 = p.dateStr.split('-');
          var yr2 = parseInt(dp2[0]), mo2 = parseInt(dp2[1]), dy2 = parseInt(dp2[2]);
          var co2 = (mo2 >= 4 && mo2 <= 9) ? 2 : 1;
          var dSMs = Date.UTC(yr2, mo2-1, dy2,   -co2, 0, 0);
          var dEMs = Date.UTC(yr2, mo2-1, dy2+1, -co2, 0, 0);
          var freeSlots = [];
          availEvts.forEach(function(av) {
            var s = Math.max(av.startMs, dSMs), e = Math.min(av.endMs, dEMs);
            for (var cur = s; cur + SLOT_HOUR_MS <= e; cur += SLOT_HOUR_MS) {
              freeSlots.push({ startMs: cur, endMs: cur + SLOT_HOUR_MS });
            }
          });
          var available = freeSlots.filter(function(slot) {
            return !blockedEvts.some(function(ev) { return slot.startMs < ev.endMs && slot.endMs > ev.startMs; });
          });
          Logger.log('[prefetch] ' + p.calId + ' ' + p.dateStr + ': ' + availEvts.length + ' avail, ' + blockedEvts.length + ' blocked, ' + available.length + ' free');
          _calSlotCache[ck] = available;
          try { _scriptCache.put(ck, JSON.stringify(available), 3600); } catch(e) {}
        } catch(pe) {
          Logger.log('[prefetch] parse error ' + p.calId + ': ' + pe.message);
          _calSlotCache[ck] = null;
        }
      });
    }

    // Convert "04:00 PM - 05:00 PM" (CET) + "2026-04-08" → [startMs, endMs] UTC
    function parseSlotToUtcMs(dateStr, slotStr) {
      var di = slotStr.indexOf(' - ');
      if (di === -1) return null;
      var startStr = slotStr.substring(0, di).trim();
      var endStr   = slotStr.substring(di + 3).trim();
      function p12(ts) {
        var m = ts.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
        if (!m) return null;
        var h = parseInt(m[1]), mn = parseInt(m[2]), ap = m[3].toUpperCase();
        if (ap === 'PM' && h !== 12) h += 12;
        if (ap === 'AM' && h === 12) h = 0;
        return { h: h, m: mn };
      }
      var st = p12(startStr), et = p12(endStr);
      if (!st || !et) return null;
      var dp = dateStr.split('-');
      var yr = parseInt(dp[0]), mo = parseInt(dp[1]), dy = parseInt(dp[2]);
      // Apr–Sep = CEST (UTC+2); Oct–Mar = CET (UTC+1)
      var cetOff = (mo >= 4 && mo <= 9) ? 2 : 1;
      return [
        Date.UTC(yr, mo - 1, dy, st.h - cetOff, st.m, 0),
        Date.UTC(yr, mo - 1, dy, et.h - cetOff, et.m, 0)
      ];
    }

    // Format a UTC ms range as "Wed, 9 Apr · 04:00 PM - 05:00 PM" in CET
    var DAY_NAMES   = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    var MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    function utcMsToCetLabel(startMs, endMs, dateStr) {
      function fmt(ms) {
        var d = new Date(ms);
        var mo = d.getUTCMonth() + 1;
        var cetOff = (mo >= 4 && mo <= 9) ? 2 : 1;
        var h  = (d.getUTCHours() + cetOff) % 24;
        var mn = d.getUTCMinutes();
        var ap = h >= 12 ? 'PM' : 'AM';
        var h12 = h % 12 || 12;
        return String(h12).padStart(2, '0') + ':' + String(mn).padStart(2, '0') + ' ' + ap;
      }
      var datePrefix = '';
      if (dateStr) {
        var dp  = dateStr.split('-');
        var dd  = new Date(parseInt(dp[0]), parseInt(dp[1]) - 1, parseInt(dp[2]));
        datePrefix = DAY_NAMES[dd.getDay()] + ', ' + dd.getDate() + ' ' + MONTH_NAMES[dd.getMonth()] + ' · ';
      }
      return datePrefix + fmt(startMs) + ' - ' + fmt(endMs);
    }

    // ── 4. Load Teacher Persona Mapping — traits + age (join by name) ──
    var personaMap = {}; // normalizedName → { traits[], ageGroups[] }
    Logger.log('[SMT] hiddenTeacherSet: ' + Object.keys(hiddenTeacherSet).length + ' teachers hidden');
    try {
      var personaData = _getCachedSheetData('Teacher Persona Mapping');
      if (personaData && personaData.length > 1) {
        var pHeaders   = personaData[0].map(function(h) { return String(h).trim(); });
        var pNameIdx   = pHeaders.indexOf('Teacher Name');
        var pAgeIdx    = pHeaders.indexOf('Preferred Age Group');
        if (pAgeIdx === -1) pAgeIdx = pHeaders.indexOf('Age Group');
        var pHiddenIdx = pHeaders.indexOf('Hidden In Search');
        var pTraitCols = pHeaders.reduce(function(acc, h, i) {
          if (/trait|expertise|style|skill|strength|personality|subject|teaching/i.test(h)) acc.push(i);
          return acc;
        }, []);
        personaData.slice(1).forEach(function(row) {
          var rn = String(row[pNameIdx] || '').trim();
          if (!rn) return;
          var key = normalizeTeacherName(rn);
          if (pHiddenIdx > -1 && String(row[pHiddenIdx] || '').trim().toLowerCase() === 'yes') return;
          var traits = [];
          pTraitCols.forEach(function(i) {
            String(row[i] || '').split(/[,\n]/).forEach(function(t) {
              var clean = t.trim(); if (clean) traits.push(clean);
            });
          });
          var ageGroups = pAgeIdx > -1
            ? String(row[pAgeIdx] || '').split(',').map(function(a) { return a.trim(); }).filter(Boolean)
            : [];
          personaMap[key] = { traits: traits, ageGroups: ageGroups };
        });
        Logger.log('[SMT] personaMap: ' + Object.keys(personaMap).length + ' teachers');
      }
    } catch(pe) { Logger.log('[SMT] personaMap error: ' + pe.message); }

    // ── 5. Load Teacher Courses (wide format: col0=Teacher, col4+=courses) ──
    var teacherCourseMap = {}; // normalizedName → { courseLower: progress }
    var upskillCountMap  = {}; // normalizedName → count
    try {
      var tcSheet = _getCachedSheetData(CONFIG.SHEETS.TEACHER_COURSES);
      if (tcSheet && tcSheet.length > 1) {
        var tcHeaderIdx = 0;
        for (var hi = 0; hi < Math.min(tcSheet.length, 10); hi++) {
          if (String(tcSheet[hi][0]).trim().toLowerCase() === 'teacher') { tcHeaderIdx = hi; break; }
        }
        var tcHeaders   = tcSheet[tcHeaderIdx];
        var COURSE_START = 4;
        tcSheet.slice(tcHeaderIdx + 1).forEach(function(row) {
          var rawName = String(row[0] || '').trim();
          if (!rawName || rawName.toLowerCase() === 'teacher') return;
          var key = normalizeTeacherName(rawName);
          var map = {}; var count = 0;
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
      }
    } catch(tce) { Logger.log('[SMT] teacherCourseMap error: ' + tce.message); }

    // Fuzzy course lookup — exact → prefix → 15-char substr → word overlap
    function getCourseProgress(tNorm, courseName) {
      if (!courseName) return null;
      var cLower  = courseName.toLowerCase().trim();
      var courses = teacherCourseMap[tNorm];
      if (!courses) return null;
      if (courses[cLower] !== undefined) return courses[cLower];
      var courseKeys = Object.keys(courses);
      for (var ci = 0; ci < courseKeys.length; ci++) {
        var ck = courseKeys[ci];
        if (ck.indexOf(cLower) === 0 || cLower.indexOf(ck) === 0) return courses[ck];
        var prefix = cLower.substring(0, 15);
        if (prefix.length >= 10 && ck.indexOf(prefix) !== -1) return courses[ck];
        // Word overlap: ≥60% of meaningful words match (handles "Design Pro with Roblox" ↔ "Design with Roblox")
        var appWords   = cLower.split(/\s+/).filter(function(w) { return w.length > 3; });
        var sheetWords = ck.split(/\s+/).filter(function(w) { return w.length > 3; });
        if (appWords.length > 0 && sheetWords.length > 0) {
          var overlap = appWords.filter(function(w) { return sheetWords.indexOf(w) > -1; }).length;
          var minLen  = Math.min(appWords.length, sheetWords.length);
          if (overlap >= Math.ceil(minLen * 0.6)) return courses[ck];
        }
      }
      return null;
    }

    // ── 5b. Normalize currentCourse/futureCourses to exact Teacher Courses sheet key ──
    // HubSpot labels (e.g. "Game development and AI with Scratch") may differ from
    // sheet names (e.g. "Game Dev and AI with Scratch") — find best fuzzy match.
    function normalizeCourseToSheetKey(courseName) {
      if (!courseName) return courseName;
      var cLower = courseName.toLowerCase().trim();
      // Collect all unique course keys across all teachers
      var allKeys = {};
      Object.keys(teacherCourseMap).forEach(function(t) {
        Object.keys(teacherCourseMap[t]).forEach(function(k) { allKeys[k] = true; });
      });
      if (allKeys[cLower] !== undefined) return courseName; // already exact
      var keys = Object.keys(allKeys);
      for (var ki = 0; ki < keys.length; ki++) {
        var ck = keys[ki];
        if (ck === cLower) return ck;
        if (ck.indexOf(cLower) === 0 || cLower.indexOf(ck) === 0) return ck;
        var prefix = cLower.substring(0, 15);
        if (prefix.length >= 10 && ck.indexOf(prefix) !== -1) return ck;
        var appW  = cLower.split(/\s+/).filter(function(w) { return w.length > 3; });
        var shW   = ck.split(/\s+/).filter(function(w) { return w.length > 3; });
        if (appW.length > 0 && shW.length > 0) {
          var ovlp = appW.filter(function(w) { return shW.indexOf(w) > -1; }).length;
          var minL = Math.min(appW.length, shW.length);
          if (ovlp >= Math.ceil(minL * 0.6)) return ck;
        }
      }
      return courseName; // no match found, keep original
    }
    currentCourse = normalizeCourseToSheetKey(currentCourse);
    futureCourses = futureCourses.map(normalizeCourseToSheetKey);
    Logger.log('[SMT] normalized currentCourse: ' + currentCourse);

    // ── 6. Audit map (last 45 days) ──
    var auditScoreMap = {};
    try { auditScoreMap = _buildAuditScoreMapDays(45); } catch(ae) { Logger.log('[SMT] audit error: ' + ae.message); }

    // ── 7. Traits / scoring helpers ──
    function splitTraits(str) {
      return str ? String(str).split(',').map(function(t){ return t.trim().toLowerCase(); }).filter(Boolean) : [];
    }
    var targetTraits = isMathCourse ? splitTraits(requestData.mathTraits) : splitTraits(requestData.techTraits);

    var PROG_SCORE = {
      '100%':30,'91-99%':27,'81-90%':24,'71-80%':21,'61-70%':18,
      '51-60%':15,'41-50%':12,'31-40%':9,'21-30%':6,'11-20%':3,'1-10%':1,'0%':0
    };

    var VALID_PROGRESS = ['71-80%','81-90%','91-99%','100%'];
    var output = [];
    var debugSlotFiltered = 0;

    // ── 7b. Parallel pre-fetch — batch all teacher calendars in one shot ──
    // This converts N sequential CalendarApp HTTP calls into 1 parallel UrlFetchApp.fetchAll,
    // reducing wall-clock time from ~N×500ms to ~500ms regardless of teacher count.
    if (requestedSlots.length > 0) {
      var prefetchPairs = [];
      migDataRows.forEach(function(mRow) {
        var rn = String(mRow[0] || '').trim();
        if (!rn || rn.toLowerCase() === 'teacher') return;
        var tn = normalizeTeacherName(rn);
        var tc = normalizeTeacherName(resolveTeacherName(rn));
        if (hiddenTeacherSet[tn] || hiddenTeacherSet[tc]) return;
        if (currentCourse) {
          var ep = getCourseProgress(tn, currentCourse);
          if (!ep && tc !== tn) ep = getCourseProgress(tc, currentCourse);
          if (VALID_PROGRESS.indexOf(ep || '') === -1) return;
        }
        var cid = calendarIdMap[tn] || calendarIdMap[tc] || null;
        if (!cid) return;
        requestedSlots.forEach(function(req) {
          var em = teacherEmailMap[tn] || teacherEmailMap[tc] || '';
          prefetchPairs.push({ calId: cid, dateStr: req.date, teacherName: rn, teacherEmail: em });
        });
      });
      prefetchCalendarsBatch(prefetchPairs);
    }

    // ── 8. Main loop — iterate over Migration Teacher data rows ──
    for (var ri = 0; ri < migDataRows.length; ri++) {
      var mRow    = migDataRows[ri];
      var rawName = String(mRow[0] || '').trim();
      if (!rawName || rawName.toLowerCase() === 'teacher') continue;

      var tNorm  = normalizeTeacherName(rawName);
      var tCanon = normalizeTeacherName(resolveTeacherName(rawName));

      // ── HIDDEN CHECK — skip teachers hidden via the "Hide" button on replacement page ──
      if (hiddenTeacherSet[tNorm] || hiddenTeacherSet[tCanon]) continue;

      // ── COURSE FILTER EARLY — skip unskilled teachers before any calendar API call ──
      if (currentCourse) {
        var earlyProg = getCourseProgress(tNorm, currentCourse);
        if (!earlyProg && tCanon !== tNorm) earlyProg = getCourseProgress(tCanon, currentCourse);
        if (VALID_PROGRESS.indexOf(earlyProg || '') === -1) continue;
      }

      // ── SLOT MATCHING — live Google Calendar (fallback: Migration Teacher sheet) ──
      var calId        = calendarIdMap[tNorm] || calendarIdMap[tCanon] || null;
      var slotsMatched = 0;
      var allAlternate = [];
      var teacherInAnyMap = false;

      for (var si = 0; si < requestedSlots.length; si++) {
        var req = requestedSlots[si];

        if (calId) {
          // ── Calendar path: get live free slots for this teacher on req.date ──
          var tEmail   = teacherEmailMap[tNorm] || teacherEmailMap[tCanon] || '';
          var calSlots = fetchTeacherCalendarSlots(calId, req.date, rawName, tEmail);
          if (calSlots !== null) {
            // null = calendar inaccessible (API error) → don't penalise teacher
            // [] or [slots] = calendar accessible → we know their real availability
            teacherInAnyMap = true;
            var utcRange = parseSlotToUtcMs(req.date, normaliseSlot(req.slot));
            if (utcRange) {
              var reqSMs = utcRange[0], reqEMs = utcRange[1];
              // Slot is FREE if availability block covers the entire requested window
              // AND no class/blocked event overlaps it (already subtracted in fetchTeacherCalendarSlots)
              if (calSlots.some(function(cs) { return cs.startMs <= reqSMs && cs.endMs >= reqEMs; })) {
                slotsMatched++;
              }
            }
            // Only collect alternates if requested slot was NOT matched (teacher unavailable at that time)
            var thisSlotMatched = utcRange && calSlots.some(function(cs) { return cs.startMs <= utcRange[0] && cs.endMs >= utcRange[1]; });
            if (!thisSlotMatched) {
              calSlots.forEach(function(cs) {
                var lbl = utcMsToCetLabel(cs.startMs, cs.endMs, req.date);
                var exists = allAlternate.some(function(x) { return (x.lbl || x) === lbl; });
                if (!exists) allAlternate.push({ lbl: lbl, startMs: cs.startMs });
              });
            }
          }
          // If calSlots === null (inaccessible): teacherInAnyMap stays false → teacher won't be dropped
        } else {
          // ── Sheet fallback: read slot text from Migration Teacher column ──
          var colIdx = slotColMap[req.date];
          if (colIdx !== undefined) {
            teacherInAnyMap = true;
            var cell    = String(mRow[colIdx] || '').trim();
            var slotArr = (cell && cell !== 'No Slots' && cell !== 'No Slots Available')
              ? cell.split(/[\n,]/).map(function(sv) { return sv.trim(); }).filter(Boolean)
              : [];
            var reqNorm = normaliseSlot(req.slot);
            var sheetSlotMatched = slotArr.some(function(sv) { return normaliseSlot(sv) === reqNorm; });
            if (sheetSlotMatched) {
              slotsMatched++;
            } else {
              // Only add alternates if this slot wasn't matched
              slotArr.forEach(function(sv) {
                var exists = allAlternate.some(function(x) { return (x.lbl || x) === sv; });
                if (!exists) allAlternate.push({ lbl: sv, startMs: 0 });
              });
            }
          }
        }
      }

      var totalSlots      = requestedSlots.length;
      var isFullSlotMatch = (totalSlots === 0) || (slotsMatched === totalSlots);
      // Drop teacher only if we have confirmed availability data and they have zero free slots
      if (totalSlots > 0 && teacherInAnyMap && slotsMatched === 0 && allAlternate.length === 0) {
        debugSlotFiltered++; continue;
      }

      // ── TRAITS + AGE — joined from Teacher Persona Mapping ──
      var personaEntry       = personaMap[tNorm] || personaMap[tCanon] || null;
      var teacherTraits      = personaEntry ? personaEntry.traits : [];
      var candidateAgeGroups = personaEntry ? personaEntry.ageGroups : [];

      var traitScore    = 15; // neutral when no traits requested
      var traitsMissing = [];
      if (targetTraits.length > 0) {
        var tLow    = teacherTraits.map(function(t) { return t.toLowerCase(); });
        var matched = targetTraits.filter(function(t) { return tLow.indexOf(t) > -1; });
        traitsMissing = targetTraits.filter(function(t) { return tLow.indexOf(t) === -1; });
        traitScore    = teacherTraits.length > 0 ? Math.round((matched.length / targetTraits.length) * 30) : 0;
      }

      var ageScore = 10; // neutral
      if (!isNaN(learnerAgeNum) && candidateAgeGroups.length > 0) {
        var ageMatched = candidateAgeGroups.some(function(ag) {
          var agL = ag.toLowerCase();
          var rng = agL.match(/(\d+)\s*[-\u2013]\s*(\d+)/);
          if (rng) return learnerAgeNum >= parseInt(rng[1]) && learnerAgeNum <= parseInt(rng[2]);
          var plus = agL.match(/(\d+)\+/);
          if (plus) return learnerAgeNum >= parseInt(plus[1]);
          return false;
        });
        ageScore = ageMatched ? 20 : 5;
      }

      // ── COURSE READINESS — from Teacher Courses wide format ──
      var currentCourseProgress = 'Not Onboarded';
      var courseScore = 0;
      if (currentCourse) {
        var prog = getCourseProgress(tNorm, currentCourse);
        if (!prog && tCanon !== tNorm) prog = getCourseProgress(tCanon, currentCourse);
        if (prog) { currentCourseProgress = prog; courseScore = PROG_SCORE[prog] !== undefined ? PROG_SCORE[prog] : 5; }
      }
      // ── DROP if not at least 71% on current course (already checked early, this is a safety net) ──
      if (currentCourse && VALID_PROGRESS.indexOf(currentCourseProgress) === -1) continue;
      function fcProg(fc) {
        if (!fc) return 'N/A';
        var p = getCourseProgress(tNorm, fc); if (p) return p;
        p = getCourseProgress(tCanon, fc); return p || 'Not Onboarded';
      }

      // ── AUDIT — last 45 days ──
      var auditData    = auditScoreMap[tNorm] || auditScoreMap[tCanon] || null;
      var auditGrade   = '\u2014';
      var redFlagCount = 0;
      var auditScore   = 10;
      if (auditData) {
        redFlagCount = auditData.redFlags || 0;
        if (auditData.avgScore != null) {
          var sc = auditData.avgScore;
          auditGrade  = sc >= 65 ? 'A' : sc >= 50 ? 'B' : sc >= 35 ? 'C' : 'D';
          auditScore  = Math.max(0, Math.round((sc / 80) * 20) - Math.min(redFlagCount * 3, 10));
        }
      }

      var totalScore = traitScore + ageScore + courseScore + auditScore;
      var slotLabel  = totalSlots === 0    ? '\u2714\uFE0F'
                     : isFullSlotMatch     ? '\u2714\uFE0F Match All'
                     : slotsMatched > 0   ? '\u26A0\uFE0F ' + slotsMatched + '/' + totalSlots
                     :                      '';

      output.push({
        teacherName           : rawName,
        ageYear               : candidateAgeGroups.join(', ') || 'N/A',
        slotMatch             : slotLabel,
        slotFullMatch         : isFullSlotMatch,
        alternateSlots        : allAlternate.sort(function(a,b){ return (a.startMs||0)-(b.startMs||0); }).map(function(x){ return x.lbl||x; }).join(', '),
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
        _rankScore            : totalScore,
        _traitMatchesCount    : targetTraits.length > 0 ? (targetTraits.length - traitsMissing.length) : 0
      });
    }

    // Sort: 1) slot match (✓ first)  2) course progress (100% > 91% > ...)  3) trait matches count
    var PROGRESS_RANK = {'100%':0,'91-99%':1,'81-90%':2,'71-80%':3};
    output.sort(function(a, b) {
      // Priority 1: slot match
      var sa = a.slotFullMatch ? 0 : 1;
      var sb = b.slotFullMatch ? 0 : 1;
      if (sa !== sb) return sa - sb;
      // Priority 2: course progress (higher % = lower rank number = earlier)
      var pa = PROGRESS_RANK[a.currentCourseProgress] !== undefined ? PROGRESS_RANK[a.currentCourseProgress] : 99;
      var pb = PROGRESS_RANK[b.currentCourseProgress] !== undefined ? PROGRESS_RANK[b.currentCourseProgress] : 99;
      if (pa !== pb) return pa - pb;
      // Priority 3: trait matches
      return (b._traitMatchesCount || 0) - (a._traitMatchesCount || 0);
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

    // ── Extra fields from Teacher Persona Mapping sheet ──────────────────────
    var gender = '', dateOfJoining = '', contactNumber = '';
    try {
      var personaMapData = _getCachedSheetData('Teacher Persona Mapping');
      if (personaMapData && personaMapData.length > 1) {
        var pmHeaders = personaMapData[0].map(function(h){ return String(h).trim(); });
        var pmNameIdx    = pmHeaders.indexOf('Teacher Name');
        var pmGenderIdx  = pmHeaders.indexOf('Gender');
        var pmDojIdx     = pmHeaders.findIndex(function(h){ return /date.*join/i.test(h); });
        if (pmDojIdx === -1) pmDojIdx = pmHeaders.indexOf('Date of Joining');
        var pmContactIdx = pmHeaders.findIndex(function(h){ return /contact|phone|mobile/i.test(h); });
        if (pmContactIdx === -1) pmContactIdx = pmHeaders.indexOf('Contact Number');

        if (pmNameIdx > -1) {
          var pmNameLower = String(teacherName).trim().toLowerCase();
          for (var pr = 1; pr < personaMapData.length; pr++) {
            var pmRow = personaMapData[pr];
            if (String(pmRow[pmNameIdx] || '').trim().toLowerCase() === pmNameLower) {
              if (pmGenderIdx  > -1) gender        = String(pmRow[pmGenderIdx]  || '').trim();
              if (pmDojIdx     > -1) {
                var doj = pmRow[pmDojIdx];
                if (doj) {
                  try { dateOfJoining = new Date(doj).toLocaleDateString('en-GB'); } catch(de) { dateOfJoining = String(doj).trim(); }
                }
              }
              if (pmContactIdx > -1) contactNumber = String(pmRow[pmContactIdx] || '').trim();
              break;
            }
          }
        }
      }
    } catch(pe) {
      Logger.log('[getTeacherProfileData] Persona enrichment error: ' + pe.message);
    }

    return {
      success: true,
      profile: {
        name:              teacherName,
        email:             teacherInfo ? (teacherInfo.email   || 'N/A') : 'N/A',
        status:            teacherInfo ? (teacherInfo.status  || 'N/A') : (loadData && loadData.status ? loadData.status : 'N/A'),
        manager:           teacherInfo ? (teacherInfo.manager || 'N/A') : 'N/A',
        clsManager:        teacherInfo ? (teacherInfo.clsManagerResponsible || 'N/A') : 'N/A',
        joinDate:          teacherInfo && teacherInfo.joinDate ? new Date(teacherInfo.joinDate).toLocaleDateString('en-GB') : (dateOfJoining || 'N/A'),
        gender:            gender || 'N/A',
        dateOfJoining:     dateOfJoining || 'N/A',
        contactNumber:     contactNumber || 'N/A',
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
  try {
    var ss    = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.TEACHER_DATA);
    if (!sheet) throw new Error('Teacher Data sheet not found');

    // Sheet has a header row at row 1; data starts at row 2
    // Columns: A=TeacherID, B=Name, ..., L=Hide in Search (index 11, col 12)
    var data   = sheet.getDataRange().getValues();
    var needle = teacherName.trim().toLowerCase();
    var updated = false;

    for (var r = 1; r < data.length; r++) { // row 0 = header
      if (String(data[r][1] || '').trim().toLowerCase() === needle) {
        sheet.getRange(r + 1, 12).setValue(isHidden ? 'Yes' : ''); // col L = 12
        updated = true;
        break;
      }
    }
    if (!updated) throw new Error('Teacher "' + teacherName + '" not found in Teacher Data sheet');

    Logger.log('[setTeacherVisibility] Teacher Data sheet updated: ' + teacherName + ' hidden=' + isHidden);
    return { success: true, message: teacherName + (isHidden ? ' hidden' : ' unhidden') + ' successfully.' };
  } catch(e) {
    Logger.log('[setTeacherVisibility] Error: ' + e.message);
    return { success: false, message: 'Could not update visibility: ' + e.message };
  }
}

// Helper: get all hidden teacher names from Teacher Data sheet (col L = Hide in Search)
function getHiddenTeachersMap() {
  try {
    var data = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA);
    var map  = {};
    for (var i = 1; i < data.length; i++) {
      var name   = String(data[i][1] || '').trim();
      var hidden = String(data[i][11] || '').trim().toLowerCase();
      if (name && hidden === 'yes') map[name] = true;
    }
    return map;
  } catch(e) {
    Logger.log('[getHiddenTeachersMap] Error: ' + e.message);
    return {};
  }
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

// =============================================
// GOOGLE CALENDAR TEST — run from GAS editor
// Verify teacher calendar availability before full implementation
// =============================================
function testCalendarSlotAvailability() {
  // --- CONFIG ---
  var TEST_DATE_STR = '2026-04-08';   // Change to any upcoming date you want to test
  var MAX_TEACHERS  = 5;              // How many teachers to sample
  var AVAIL_KW      = ['availability', 'available hours', 'teaching hours'];
  var SLOT_MS       = 60 * 60 * 1000; // 1 hour

  Logger.log('===== testCalendarSlotAvailability (v2 — own-calendar subtraction) =====');
  Logger.log('Test date: ' + TEST_DATE_STR);
  Logger.log('Logic: free slot = Teacher Availability Hour with NO other event overlapping it on their OWN calendar');

  var dateParts = TEST_DATE_STR.split('-');
  var testDate  = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));

  // Load Migration Teacher sheet
  var auditSS;
  try { auditSS = SpreadsheetApp.openById(CONFIG.AUDIT_SHEET_ID); }
  catch (e) { Logger.log('ERROR opening AUDIT sheet: ' + e.message); return; }

  var migSheet = null;
  var allSheets = auditSS.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName().toLowerCase().indexOf('migration teacher') > -1) {
      migSheet = allSheets[i]; break;
    }
  }
  if (!migSheet) { Logger.log('ERROR: Migration Teacher tab not found'); return; }
  Logger.log('Found sheet: ' + migSheet.getName());

  var migData = migSheet.getDataRange().getValues();
  var headerRowIdx = -1;
  for (var r = 0; r < Math.min(10, migData.length); r++) {
    if (String(migData[r][0]).trim().toLowerCase() === 'teacher') { headerRowIdx = r; break; }
  }
  if (headerRowIdx === -1) { Logger.log('ERROR: No "Teacher" header row found'); return; }

  // Collect up to MAX_TEACHERS teachers with a calendar ID in col B
  var teachers = [];
  for (var r = headerRowIdx + 1; r < migData.length && teachers.length < MAX_TEACHERS; r++) {
    var name = String(migData[r][0] || '').trim();
    var cid  = String(migData[r][1] || '').trim();
    if (!name || !cid || cid.indexOf('@') === -1) continue;
    teachers.push({ name: name, calendarId: cid });
  }
  Logger.log('Teachers to test (' + teachers.length + '): ' + teachers.map(function(t) { return t.name; }).join(', '));

  // CET label helper (Apr-Sep = CEST UTC+2, else CET UTC+1)
  function toCetLabel(startMs, endMs) {
    function fmt(ms) {
      var d = new Date(ms), mo = d.getUTCMonth() + 1, off = (mo >= 4 && mo <= 9) ? 2 : 1;
      var h = (d.getUTCHours() + off) % 24, mn = d.getUTCMinutes();
      var ap = h >= 12 ? 'PM' : 'AM', h12 = h % 12 || 12;
      return String(h12).padStart(2,'0') + ':' + String(mn).padStart(2,'0') + ' ' + ap;
    }
    return fmt(startMs) + ' - ' + fmt(endMs) + ' CET';
  }

  // --- Process each teacher ---
  teachers.forEach(function(teacher) {
    Logger.log('\n--- ' + teacher.name + ' (' + teacher.calendarId + ') ---');

    var cal;
    try { cal = CalendarApp.getCalendarById(teacher.calendarId); } catch(e) { Logger.log('  ERROR: ' + e.message); return; }
    if (!cal) { Logger.log('  ERROR: calendar null — not shared with hello@jet-learn.com'); return; }

    var allEvents = cal.getEventsForDay(testDate);
    Logger.log('  Total events on ' + TEST_DATE_STR + ': ' + allEvents.length);

    // ── Availability events (teacher-marked open blocks) ──
    var availEvents = allEvents.filter(function(ev) {
      var t = ev.getTitle().toLowerCase();
      return AVAIL_KW.some(function(kw) { return t.indexOf(kw) > -1; });
    });

    // ── Everything else on teacher's own calendar = blocked time ──
    var blockedEvents = allEvents.filter(function(ev) {
      var t = ev.getTitle().toLowerCase();
      return !AVAIL_KW.some(function(kw) { return t.indexOf(kw) > -1; });
    });

    if (availEvents.length === 0) {
      Logger.log('  No availability events found.');
      Logger.log('  All event titles: ' + allEvents.map(function(e) { return '"' + e.getTitle() + '"'; }).join(', '));
      return;
    }

    Logger.log('  Availability events: ' + availEvents.length);
    Logger.log('  Other (blocked) events: ' + blockedEvents.length + ' → ' +
      blockedEvents.map(function(e) { return '"' + e.getTitle() + '"'; }).join(', '));

    // Split each availability block into 1-hour chunks
    var rawSlots = [];
    availEvents.forEach(function(av) {
      var s = av.getStartTime().getTime(), e = av.getEndTime().getTime();
      for (var cur = s; cur + SLOT_MS <= e; cur += SLOT_MS) {
        rawSlots.push({ startMs: cur, endMs: cur + SLOT_MS });
      }
    });
    Logger.log('  Raw 1-hour slots: ' + rawSlots.length);

    // Subtract any blocked event that overlaps a slot
    var freeSlots = rawSlots.filter(function(slot) {
      var blocker = blockedEvents.filter(function(ev) {
        return slot.startMs < ev.getEndTime().getTime() && slot.endMs > ev.getStartTime().getTime();
      });
      if (blocker.length > 0) {
        Logger.log('  BLOCKED ' + toCetLabel(slot.startMs, slot.endMs) + ' by: ' +
          blocker.map(function(e) { return '"' + e.getTitle() + '"'; }).join(', '));
      }
      return blocker.length === 0;
    });

    Logger.log('  FREE slots (' + freeSlots.length + '):');
    if (freeSlots.length === 0) {
      Logger.log('    (none — teacher fully booked on this day)');
    } else {
      freeSlots.forEach(function(s) { Logger.log('    ✓ ' + toCetLabel(s.startMs, s.endMs)); });
    }
  });

  Logger.log('\n===== TEST COMPLETE =====');
}

// ─────────────────────────────────────────────────────────────────────────────
// TP MANAGER DASHBOARD
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Reads all 3 audit sheets and returns records as an array of objects.
 */
function _getAllAuditRecords() {
  var ss      = SpreadsheetApp.openById(CONFIG.AUDIT_SHEET_ID);
  var sheets  = ['Coding Audit\'26', 'Math Audit\'26', 'GCSE Audit\'26'];
  var records = [];
  sheets.forEach(function(sheetName) {
    try {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var teacher = String(row[2] || '').trim();
        if (!teacher) continue;
        var fmtDate = function(v) {
          if (!v) return '';
          try { return Utilities.formatDate(new Date(v), Session.getScriptTimeZone(), 'yyyy-MM-dd'); } catch(e) { return String(v); }
        };
        records.push({
          auditType       : sheetName,
          date            : fmtDate(row[1]),
          teacher         : teacher,
          learner         : String(row[3]  || '').trim(),
          course          : String(row[4]  || '').trim(),
          classType       : String(row[5]  || '').trim(),
          classDate       : fmtDate(row[6]),
          recording       : String(row[8]  || '').trim(),
          auditReason     : String(row[9]  || '').trim(),
          priority        : String(row[10] || '').trim(),
          auditor         : String(row[12] || '').trim(),
          classScore      : (row[13] !== '' && row[13] !== null && row[13] !== undefined) ? Number(row[13]) : null,
          warmUp          : (row[14] !== '' && row[14] !== null && row[14] !== undefined) ? Number(row[14]) : null,
          activity        : (row[15] !== '' && row[15] !== null && row[15] !== undefined) ? Number(row[15]) : null,
          wrapUp          : (row[16] !== '' && row[16] !== null && row[16] !== undefined) ? Number(row[16]) : null,
          overallHygiene  : (row[17] !== '' && row[17] !== null && row[17] !== undefined) ? Number(row[17]) : null,
          redFlags        : String(row[18] || '').trim(),
          completionDate  : fmtDate(row[19])
        });
      }
    } catch(e) {
      Logger.log('[_getAllAuditRecords] Error reading ' + sheetName + ': ' + e.message);
    }
  });
  return records;
}

/**
 * Returns grade letter based on score out of 80.
 */
function _auditGrade(score) {
  if (score === null || score === undefined || isNaN(score)) return '—';
  if (score >= 65) return 'A';
  if (score >= 50) return 'B';
  if (score >= 35) return 'C';
  return 'D';
}

/**
 * Main server function for TP Manager Dashboard.
 * Returns all teachers under the given TP manager with their audit summary.
 */
/**
 * @param {string} tpManagerName
 * @param {string} dateFrom  — ISO date string e.g. "2026-02-08" (optional, default 60 days ago)
 * @param {string} dateTo    — ISO date string e.g. "2026-04-09" (optional, default today)
 */
function getTPManagerDashboard(tpManagerName, dateFrom, dateTo) {
  try {
    if (!tpManagerName) return { success: false, message: 'TP Manager name required.' };

    // Date range — default last 60 days
    var now    = new Date();
    var toMs   = dateTo   ? new Date(dateTo).getTime()   : now.getTime();
    var fromMs = dateFrom ? new Date(dateFrom).getTime() : (now.getTime() - 60 * 24 * 60 * 60 * 1000);

    // 1. Get all teachers under this TP
    var allTeachers = getTeacherData();
    var myTeachers  = allTeachers.filter(function(t) {
      return t.manager && t.manager.trim().toLowerCase() === tpManagerName.trim().toLowerCase();
    });
    if (myTeachers.length === 0) return { success: true, teachers: [], summary: {} };

    // 2. Get all audit records then filter by date range
    var allAudits = _getAllAuditRecords();
    var rangedAudits = allAudits.filter(function(a) {
      if (!a.date) return false;
      var d = new Date(a.date).getTime();
      return d >= fromMs && d <= toMs;
    });

    // 3. Build per-teacher summary
    var teacherRows = myTeachers.map(function(t) {
      var tNameNorm = t.name.trim().toLowerCase();

      // All audits for this teacher (ALL TIME for latest score)
      var allMyAudits = allAudits.filter(function(a) {
        return a.teacher.trim().toLowerCase() === tNameNorm;
      }).sort(function(a, b) {
        return new Date(b.date).getTime() - new Date(a.date).getTime();
      });

      // Date-ranged audits for this teacher
      var rangedMyAudits = rangedAudits.filter(function(a) {
        return a.teacher.trim().toLowerCase() === tNameNorm;
      }).sort(function(a, b) {
        return new Date(b.date).getTime() - new Date(a.date).getTime();
      });

      var latest      = allMyAudits.length > 0 ? allMyAudits[0] : null;
      var lastScore   = latest ? latest.classScore : null;
      var lastGrade   = _auditGrade(lastScore);
      var lastDate    = latest ? latest.date : null;
      var lastAuditor = latest ? latest.auditor : '';

      // Red flag incidents within date range, per audit type
      var redFlagIncidents = [];
      rangedMyAudits.forEach(function(a) {
        if (a.redFlags) {
          redFlagIncidents.push({
            auditType : a.auditType,
            learner   : a.learner,
            course    : a.course,
            date      : a.date,
            flags     : a.redFlags,
            score     : a.classScore,
            grade     : _auditGrade(a.classScore),
            recording : a.recording,
            auditor   : a.auditor,
            warmUp    : a.warmUp,
            activity  : a.activity,
            wrapUp    : a.wrapUp,
            hygiene   : a.overallHygiene
          });
        }
      });

      // Group ranged audits by audit type (Coding / Math / GCSE)
      var byType = {};
      rangedMyAudits.forEach(function(a) {
        var type = a.auditType.replace(" Audit'26", '').trim(); // "Coding", "Math", "GCSE"
        if (!byType[type]) byType[type] = [];
        byType[type].push({
          date      : a.date,
          learner   : a.learner,
          course    : a.course,
          score     : a.classScore,
          grade     : _auditGrade(a.classScore),
          redFlags  : a.redFlags,
          auditor   : a.auditor,
          recording : a.recording,
          warmUp    : a.warmUp,
          activity  : a.activity,
          wrapUp    : a.wrapUp,
          hygiene   : a.overallHygiene
        });
      });

      // Due for audit: no audit in last 30 days
      var thirtyDaysAgo = now.getTime() - 30 * 24 * 60 * 60 * 1000;
      var dueForAudit = !lastDate || new Date(lastDate).getTime() < thirtyDaysAgo;

      // Risk score: High / Medium / Low / Stable
      var riskScore;
      if (t.status === 'EWS' || lastGrade === 'D') {
        riskScore = 'High';
      } else if (lastGrade === 'C' || redFlagIncidents.length >= 2) {
        riskScore = 'Medium';
      } else if (lastGrade === 'B' || redFlagIncidents.length === 1) {
        riskScore = 'Low';
      } else {
        riskScore = 'Stable';
      }

      return {
        name            : t.name,
        status          : t.status,
        email           : t.email,
        hiddenInSearch  : t.hiddenInSearch || '',
        lastAuditDate   : lastDate,
        lastScore       : lastScore,
        lastGrade       : lastGrade,
        lastAuditor     : lastAuditor,
        auditCountAll   : allMyAudits.length,
        auditCountRange : rangedMyAudits.length,
        redFlagCount    : redFlagIncidents.length,
        redFlagIncidents: redFlagIncidents,
        auditsByType    : byType,
        dueForAudit     : dueForAudit,
        riskScore       : riskScore,
        lastAuditDetail : latest ? {
          course    : latest.course,
          learner   : latest.learner,
          classType : latest.classType,
          recording : latest.recording,
          warmUp    : latest.warmUp,
          activity  : latest.activity,
          wrapUp    : latest.wrapUp,
          hygiene   : latest.overallHygiene,
          redFlags  : latest.redFlags,
          auditor   : latest.auditor
        } : null
      };
    });

    // 4. Summary stats (based on date range)
    var ewsCount     = teacherRows.filter(function(t) { return t.status === 'EWS'; }).length;
    var noAuditCount = teacherRows.filter(function(t) { return t.auditCountRange === 0; }).length;
    var redFlagTotal = teacherRows.reduce(function(s, t) { return s + t.redFlagCount; }, 0);
    var scores       = teacherRows.filter(function(t) { return t.lastScore !== null; }).map(function(t) { return t.lastScore; });
    var avgScore     = scores.length > 0 ? Math.round(scores.reduce(function(a,b){return a+b;},0) / scores.length) : null;

    return {
      success  : true,
      teachers : teacherRows,
      dateFrom : new Date(fromMs).toISOString().substring(0, 10),
      dateTo   : new Date(toMs).toISOString().substring(0, 10),
      summary  : { total: teacherRows.length, ewsCount: ewsCount, noAudit: noAuditCount, redFlags: redFlagTotal, avgScore: avgScore }
    };
  } catch(e) {
    Logger.log('[getTPManagerDashboard] Error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// TP NOTES
// ─────────────────────────────────────────────────────────────────────────────
function saveTPNote(teacherName, tpManager, noteText, createdBy) {
  try {
    if (!teacherName || !noteText) return { success: false, message: 'Teacher and note required.' };
    var sheet = getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.TP_NOTES);
    var id = Utilities.getUuid();
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([id, teacherName, tpManager || '', noteText, createdBy || Session.getActiveUser().getEmail(), now, now]);
    return { success: true, id: id };
  } catch(e) {
    Logger.log('[saveTPNote] ' + e.message);
    return { success: false, message: e.message };
  }
}

function getTPNotes(teacherName) {
  try {
    var data = _getAppDataCachedSheetData(CONFIG.APP_DATA_SHEETS.TP_NOTES);
    if (data.length < 2) return { success: true, notes: [] };
    var nameLow = (teacherName || '').toLowerCase().trim();
    var notes = data.slice(1).filter(function(r) {
      return String(r[1] || '').toLowerCase().trim() === nameLow;
    }).map(function(r) {
      return { id: r[0], teacher: r[1], tpManager: r[2], note: r[3], createdBy: r[4], createdAt: r[5] };
    }).reverse(); // newest first
    return { success: true, notes: notes };
  } catch(e) {
    Logger.log('[getTPNotes] ' + e.message);
    return { success: false, notes: [] };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// UPSKILL TRACKER
// ─────────────────────────────────────────────────────────────────────────────
function logUpskillTask(teacherName, tpManager, jlid, learnerName, gapCourses, hsTaskId) {
  try {
    var sheet = getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.UPSKILL_TRACKER);
    var id = Utilities.getUuid();
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([id, teacherName, tpManager || '', jlid || '', learnerName || '', gapCourses || '', hsTaskId || '', 'Pending', now, '', '']);
    _clearAppDataCache(CONFIG.APP_DATA_SHEETS.UPSKILL_TRACKER);
    Logger.log('[logUpskillTask] Logged upskill task for ' + teacherName);
    return { success: true, id: id };
  } catch(e) {
    Logger.log('[logUpskillTask] ' + e.message);
    return { success: false, message: e.message };
  }
}

function updateUpskillStatus(id, status, notes) {
  try {
    var sheet = getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.UPSKILL_TRACKER);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, 8).setValue(status); // Status col
        if (notes) sheet.getRange(i + 1, 11).setValue(notes);
        if (status === 'Done') sheet.getRange(i + 1, 10).setValue(
          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
        );
        _clearAppDataCache(CONFIG.APP_DATA_SHEETS.UPSKILL_TRACKER);
        return { success: true };
      }
    }
    return { success: false, message: 'Task not found.' };
  } catch(e) {
    Logger.log('[updateUpskillStatus] ' + e.message);
    return { success: false, message: e.message };
  }
}

function getUpskillTasksForTeacher(teacherName) {
  try {
    var data = _getAppDataCachedSheetData(CONFIG.APP_DATA_SHEETS.UPSKILL_TRACKER);
    if (data.length < 2) return { success: true, tasks: [] };
    var nameLow = (teacherName || '').toLowerCase().trim();
    var tasks = data.slice(1).filter(function(r) {
      return String(r[1] || '').toLowerCase().trim() === nameLow;
    }).map(function(r) {
      return { id: r[0], teacher: r[1], tpManager: r[2], jlid: r[3], learner: r[4], gapCourses: r[5], hsTaskId: r[6], status: r[7], createdAt: r[8], completedAt: r[9], notes: r[10] };
    }).reverse();
    return { success: true, tasks: tasks };
  } catch(e) {
    Logger.log('[getUpskillTasksForTeacher] ' + e.message);
    return { success: false, tasks: [] };
  }
}
// ─────────────────────────────────────────────────────────────────────────────
// RED FLAG ALERT EMAIL
// Runs on a daily trigger. Checks last 2 days of audits for red flags,
// sends one email per teacher-per-flag to their TP manager. Logs to Alert Log
// to prevent duplicates.
// ─────────────────────────────────────────────────────────────────────────────
function checkAndSendRedFlagAlerts() {
  try {
    var now      = new Date();
    var cutoff   = new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000); // last 2 days
    var allAudits = _getAllAuditRecords();

    // Load sent alert keys from Alert Log to avoid duplicates
    var alertSheet = getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.ALERT_LOG);
    var alertData  = alertSheet.getDataRange().getValues();
    var sentKeys   = {};
    alertData.slice(1).forEach(function(r) { if (r[0]) sentKeys[String(r[0])] = true; });

    var teacherMap = {};
    try {
      getTeacherData().forEach(function(t) {
        if (t.name) teacherMap[t.name.toLowerCase().trim()] = t;
      });
    } catch(e) {}

    var alertsSent = 0;
    allAudits.forEach(function(a) {
      if (!a.redFlags) return;
      if (!a.date) return;
      var auditDate = new Date(a.date);
      if (auditDate < cutoff) return;

      // Unique key: teacher + learner + date + auditType
      var key = [a.teacher, a.learner, a.date, a.auditType].join('||');
      if (sentKeys[key]) return; // already sent

      // Find TP manager
      var tMatch = teacherMap[(a.teacher || '').toLowerCase().trim()];
      var tpManagerEmail = tMatch ? tMatch.tpManagerEmail : '';
      var tpManagerName  = tMatch ? tMatch.manager        : '';
      if (!tpManagerEmail) {
        Logger.log('[RedFlagAlert] No TP manager email for ' + a.teacher + ', skipping.');
        return;
      }

      // Send alert email
      var subject = '⚑ Red Flag Alert — ' + a.teacher + ' (' + (a.auditType||'').replace(" Audit'26",'') + ' Audit)';
      var body = 'Hi ' + (tpManagerName || 'TP Manager') + ',\n\n'
        + 'A red flag was logged in a recent audit.\n\n'
        + 'Teacher     : ' + a.teacher + '\n'
        + 'Learner     : ' + a.learner + '\n'
        + 'Course      : ' + a.course  + '\n'
        + 'Audit Type  : ' + (a.auditType||'').replace(" Audit'26",'') + '\n'
        + 'Audit Date  : ' + a.date + '\n'
        + 'Score       : ' + (a.classScore !== null ? a.classScore + '/80' : 'N/A') + '\n'
        + 'Red Flags   : ' + a.redFlags + '\n\n'
        + 'Please review and take necessary action before the next session.\n\n'
        + '— JetLearn Automated Alert';

      try {
        GmailApp.sendEmail(tpManagerEmail, subject, body, {
          from: CONFIG.EMAIL.FROM,
          name: 'JetLearn Alerts'
        });
        // Log to Alert Log
        var logId = Utilities.getUuid();
        alertSheet.appendRow([
          key, 'Red Flag', a.teacher, tpManagerName, '', a.learner,
          a.redFlags, Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
          tpManagerEmail
        ]);
        sentKeys[key] = true;
        alertsSent++;
        Logger.log('[RedFlagAlert] Sent alert for ' + a.teacher + ' to ' + tpManagerEmail);
      } catch(mailErr) {
        Logger.log('[RedFlagAlert] Email failed for ' + a.teacher + ': ' + mailErr.message);
      }
    });

    Logger.log('[RedFlagAlert] Done. Alerts sent: ' + alertsSent);
    return { success: true, sent: alertsSent };
  } catch(e) {
    Logger.log('[checkAndSendRedFlagAlerts] Error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Teaching Hours Distribution
// Reads events from the class-schedule calendar (hello@jet-learn.com) for the
// past N weeks and aggregates teaching hours per day of week for a given teacher.
// ─────────────────────────────────────────────────────────────────────────────
function getTeacherCalendarHours(teacherName, weeksBack) {
  try {
    weeksBack = weeksBack || 4;
    var cal = CalendarApp.getCalendarById(CONFIG.CLASS_SCHEDULE_CALENDAR_ID);
    if (!cal) return { success: false, message: 'Calendar not accessible.' };

    var now     = new Date();
    var startDt = new Date(now.getTime() - weeksBack * 7 * 24 * 60 * 60 * 1000);

    var events = cal.getEvents(startDt, now);
    Logger.log('[CalHours] total events in range: ' + events.length);

    // Normalise teacher name parts for matching (need all name parts ≥3 chars to appear in title)
    var nameLower = String(teacherName || '').trim().toLowerCase();
    var nameParts = nameLower.split(/\s+/).filter(function(p) { return p.length >= 3; });

    // Also get teacher email for guest-list matching
    var teacherEmail = '';
    try {
      var tdRows = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA);
      var nl = nameLower;
      for (var ti = 1; ti < tdRows.length; ti++) {
        if (String(tdRows[ti][1] || '').trim().toLowerCase() === nl) {
          teacherEmail = String(tdRows[ti][8] || '').trim().toLowerCase();
          break;
        }
      }
    } catch(e) {}

    // hoursByDow[0]=Sun … [6]=Sat; we re-index to Mon=0 … Sun=6 at return time
    var hoursByDow  = [0, 0, 0, 0, 0, 0, 0];
    var eventsByDow = [0, 0, 0, 0, 0, 0, 0];
    var totalMinutes = 0;
    var matched = 0;

    events.forEach(function(ev) {
      if (ev.isAllDayEvent()) return;

      var title  = ev.getTitle().toLowerCase();
      var inTitle = nameParts.length > 0 && nameParts.every(function(p) { return title.indexOf(p) > -1; });

      var inGuests = false;
      if (!inTitle && teacherEmail) {
        try {
          ev.getGuestList().forEach(function(g) {
            if (g.getEmail().toLowerCase() === teacherEmail) inGuests = true;
          });
        } catch(e) {}
      }

      if (!inTitle && !inGuests) return;

      matched++;
      var durationMin = (ev.getEndTime().getTime() - ev.getStartTime().getTime()) / 60000;
      if (durationMin <= 0 || durationMin > 480) return; // ignore all-day-like durations

      var dow = ev.getStartTime().getDay(); // 0=Sun
      hoursByDow[dow]  += durationMin / 60;
      eventsByDow[dow] += 1;
      totalMinutes     += durationMin;
    });

    Logger.log('[CalHours] matched events for "' + teacherName + '": ' + matched);

    // Re-order to Mon(0)…Sun(6) for the chart
    var DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
    // JS dow: 0=Sun,1=Mon,...,6=Sat → our index: Mon=0,...Sun=6
    var dowMap = [6, 0, 1, 2, 3, 4, 5]; // dowMap[jsDow] → our index
    var hoursArr  = [0, 0, 0, 0, 0, 0, 0];
    var eventsArr = [0, 0, 0, 0, 0, 0, 0];
    for (var d = 0; d < 7; d++) {
      hoursArr[dowMap[d]]  = Math.round(hoursByDow[d]  * 10) / 10;
      eventsArr[dowMap[d]] = eventsByDow[d];
    }

    var totalHours = Math.round(totalMinutes / 60 * 10) / 10;
    var avgPerWeek = Math.round(totalHours / weeksBack * 10) / 10;

    // Peak day
    var peakIdx  = hoursArr.indexOf(Math.max.apply(null, hoursArr));
    var peakDay  = hoursArr[peakIdx] > 0 ? DAYS[peakIdx] : 'N/A';

    return {
      success:    true,
      days:       DAYS,
      hours:      hoursArr,
      eventCount: eventsArr,
      totalHours: totalHours,
      avgPerWeek: avgPerWeek,
      peakDay:    peakDay,
      weeksBack:  weeksBack,
      matched:    matched
    };
  } catch(e) {
    Logger.log('[getTeacherCalendarHours] Error: ' + e.message);
    return { success: false, message: e.message };
  }
}


// One-time setup: creates a daily trigger for red flag alerts.
// Run once from the GAS editor.
function setupRedFlagAlertTrigger() {
  // Remove existing triggers for this function first
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'checkAndSendRedFlagAlerts') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkAndSendRedFlagAlerts')
    .timeBased().everyDays(1).atHour(8).create();
  Logger.log('[setupRedFlagAlertTrigger] Daily trigger created at 8AM.');
}


// ─────────────────────────────────────────────────────────────────────────────
// Teacher Onboarding Tracker — CRUD
// Sheet: JetLearn App Data → "Onboarding Tracker"
// Columns: ID | Teacher Name | Join Date | TP Manager | Primary Course |
//          Training % | First Class Done | First Audit Done | 30-Day Review |
//          Notes | Status | Created At
// ─────────────────────────────────────────────────────────────────────────────
var OB_HEADERS = ['ID','Teacher Name','Join Date','TP Manager','Primary Course',
                  'Training %','First Class Done','First Audit Done','30-Day Review Done',
                  'Notes','Status','Created At'];

function getOnboardingTrackerData() {
  try {
    var sheet = getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.ONBOARDING_TRACKER);
    // Ensure headers
    if (sheet.getLastRow() === 0 || sheet.getRange('A1').getValue() !== 'ID') {
      sheet.clearContents();
      sheet.appendRow(OB_HEADERS);
    }
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, records: [] };

    var fmt = function(v) {
      if (!v) return '';
      if (v instanceof Date) {
        try { return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd'); } catch(e) { return String(v); }
      }
      return String(v);
    };

    var records = data.slice(1).map(function(row) {
      return {
        id:              String(row[0]  || ''),
        teacherName:     String(row[1]  || ''),
        joinDate:        fmt(row[2]),
        tpManager:       String(row[3]  || ''),
        primaryCourse:   String(row[4]  || ''),
        trainingPct:     Number(row[5]  || 0),
        firstClassDone:  String(row[6]  || 'No'),
        firstAuditDone:  String(row[7]  || 'No'),
        reviewDone:      String(row[8]  || 'No'),
        notes:           String(row[9]  || ''),
        status:          String(row[10] || 'Active'),
        createdAt:       fmt(row[11])
      };
    }).filter(function(r) { return r.teacherName; });

    return { success: true, records: records };
  } catch(e) {
    Logger.log('[getOnboardingTrackerData] Error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function addOnboardingRecord(data) {
  try {
    var sheet = getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.ONBOARDING_TRACKER);
    if (sheet.getLastRow() === 0 || sheet.getRange('A1').getValue() !== 'ID') {
      sheet.clearContents();
      sheet.appendRow(OB_HEADERS);
    }
    var id  = 'OB-' + Utilities.getUuid().split('-')[0].toUpperCase();
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([
      id,
      data.teacherName  || '',
      data.joinDate     || '',
      data.tpManager    || '',
      data.primaryCourse|| '',
      0,        // Training %
      'No',     // First Class Done
      'No',     // First Audit Done
      'No',     // 30-Day Review Done
      '',       // Notes
      'Active', // Status
      now
    ]);
    return { success: true, id: id };
  } catch(e) {
    Logger.log('[addOnboardingRecord] Error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function updateOnboardingRecord(id, field, value) {
  try {
    var sheet   = getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.ONBOARDING_TRACKER);
    var data    = sheet.getDataRange().getValues();
    var colMap  = {};
    OB_HEADERS.forEach(function(h, i) { colMap[h] = i; });
    var fieldColMap = {
      'trainingPct':    'Training %',
      'firstClassDone': 'First Class Done',
      'firstAuditDone': 'First Audit Done',
      'reviewDone':     '30-Day Review Done',
      'notes':          'Notes',
      'status':         'Status'
    };
    var col = colMap[fieldColMap[field]];
    if (col === undefined) return { success: false, message: 'Unknown field: ' + field };

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, col + 1).setValue(value);
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, message: 'Record not found.' };
  } catch(e) {
    Logger.log('[updateOnboardingRecord] Error: ' + e.message);
    return { success: false, message: e.message };
  }
}