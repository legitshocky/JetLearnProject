function getAuditLog(params = {}) {
  Logger.log('getAuditLog called with params: ' + JSON.stringify(params));

  // Initialize defaults to ensure consistent return structure
  let paginatedData = [];
  let totalItems = 0;
  let totalPages = 0;
  let currentPage = params.page || 1;

  try {
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.AUDIT_LOG);
    
    // Check if data exists and has more than just headers
    if (sheetData && sheetData.length > 1) { 
        let filteredData = sheetData.slice(1); 

        if (params.status && params.status !== 'all') {
          filteredData = filteredData.filter(row => row[7] && String(row[7]).toLowerCase() === params.status.toLowerCase());
        }

        if (params.fromDate || params.toDate) {
          const fromDate = params.fromDate ? new Date(params.fromDate) : null;
          if (fromDate) fromDate.setHours(0, 0, 0, 0);

          const toDate = params.toDate ? new Date(params.toDate) : null;
          if(toDate) toDate.setHours(23, 59, 59, 999);
          
          filteredData = filteredData.filter(row => {
            const rowDate = parseSheetDate(row[0]); 
            if (!rowDate) return false;
            const isAfterOrEqualFrom = fromDate ? rowDate >= fromDate : true;
            const isBeforeOrEqualTo = toDate ? rowDate <= toDate : true;
            return isAfterOrEqualFrom && isBeforeOrEqualTo;
          });
        }

        if (params.search) {
          const searchTerm = params.search.toLowerCase();
          filteredData = filteredData.filter(row =>
            row.some(cell => cell && String(cell).toLowerCase().includes(searchTerm))
          );
        }

        // Sort descending by date
        filteredData.sort((a, b) => {
            const dateA = parseSheetDate(a[0]);
            const dateB = parseSheetDate(b[0]);
            if (!dateA && !dateB) return 0;
            if (!dateA) return 1; 
            if (!dateB) return -1;
            return dateB.getTime() - dateA.getTime();
        });

        totalItems = filteredData.length;
        const limit = params.limit || CONFIG.PAGINATION_LIMIT;
        const startIndex = (currentPage - 1) * limit;
        const endIndex = startIndex + limit;
        paginatedData = filteredData.slice(startIndex, endIndex);
        totalPages = (totalItems > 0 && limit > 0) ? Math.ceil(totalItems / limit) : 0;
    }

    return {
      data: paginatedData,
      total: totalItems,
      page: currentPage,
      totalPages: totalPages
    };

  } catch (error) {
    Logger.log('getAuditLog error: ' + error.message);
    // CRITICAL FIX: Return the full structure even on error to prevent client crashes
    return { 
      data: [], 
      total: 0, 
      page: 1, 
      totalPages: 0 
    };
  }
}

function logAction(action, jlid, learner, oldTeacher, newTeacher, course, status, notes, reason = '', intervenedBy = '') {
  Logger.log(`logAction called: ${action} for JLID: ${jlid}`);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    
    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.AUDIT_LOG);
    const timestamp = new Date();
    const sessionId = Utilities.getUuid();
    sheet.appendRow([
      timestamp,
      action || '',
      jlid || '',
      learner || '',
      oldTeacher || '',
      newTeacher || '',
      course || '',
      status || 'Unknown',
      notes || '',
      sessionId,
      reason || '', 
      intervenedBy || ''  
    ]);
    SpreadsheetApp.flush();    
    if(typeof _sheetDataCache !== 'undefined') {
        delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.AUDIT_LOG}`];
    }
    Logger.log('Action logged successfully: ' + action);
  } catch (error) {
    Logger.log('Error logging action (Lock/Write failed): ' + error.message);
  } finally {
    lock.releaseLock();
  }
}


function logUserActivity(username, action, details) {
  try {
    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.USER_ACTIVITY_LOG);
    if (sheet.getLastRow() === 0 || sheet.getRange('A1').isBlank()) {
      sheet.appendRow(['Timestamp', 'Username', 'Action', 'Details', 'UserEmail']);
    }

    const timestamp = new Date();

    sheet.appendRow([
      timestamp,
      username,
      action,
      details || '',
      Session.getActiveUser().getEmail() 
    ]);
    delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.USER_ACTIVITY_LOG}`];
  } catch (error) {
    Logger.log('Error logging user activity: ' + error.message);
  }
}

function exportAuditData(startDate, endDate) {
  Logger.log('exportAuditData called');

  try {
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.AUDIT_LOG); 

    if (sheetData.length <= 1) { 
      return sheetData;
    }

    if (!startDate || !endDate) {
      return sheetData; 
    }

    const start = new Date(startDate);
    start.setHours(0,0,0,0);
    const end = new Date(endDate);
    end.setHours(23,59,59,999);

    const filteredData = sheetData.filter((row, index) => {
      if (index === 0) return true; 
      const rowDate = new Date(row[0]);
      return rowDate >= start && rowDate <= end;
    });

    Logger.log('Exported ' + filteredData.length + ' audit records');
    return filteredData;
  }
   catch (error) {
    Logger.log('Export failed: ' + error.message);
    return [];
  }
}

function cleanupOldAuditLogs(daysToKeep = 90) {
  Logger.log('cleanupOldAuditLogs called for ' + daysToKeep + ' days');

  try {
    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.AUDIT_LOG);
    const data = _getCachedSheetData(CONFIG.SHEETS.AUDIT_LOG); 

    if (data.length <= 1) return { success: true, message: 'No data to cleanup' };

    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);
    cutoffDate.setHours(0,0,0,0); 

    const header = data[0]; 
    const recordsToKeep = data.slice(1).filter(row => {
      const rowDate = new Date(row[0]);
      return rowDate >= cutoffDate;
    });

    sheet.clearContents();
    sheet.getRange(1, 1, 1, header.length).setValues([header]); 
    if (recordsToKeep.length > 0) {
      sheet.getRange(2, 1, recordsToKeep.length, recordsToKeep[0].length).setValues(recordsToKeep);
    }

    const deletedCount = data.length - 1 - recordsToKeep.length; 
    Logger.log('Cleaned up ' + deletedCount + ' old audit records');

    delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.AUDIT_LOG}`];

    return {
      success: true,
      message: `Cleaned up ${deletedCount} records older than ${daysToKeep} days`
    };
  } catch (error) {
    Logger.log('Cleanup failed: ' + error.message);
    return { success: false, message: 'Cleanup failed: ' + error.message };
  }
}

function getRawLearnerAuditLog(searchTerm) {
  Logger.log('getRawLearnerAuditLog called for searchTerm: ' + searchTerm);

  try {
    if (!searchTerm || searchTerm.trim() === '') {
      return { success: false, message: 'Search term cannot be empty.' };
    }

    const auditSheetData = _getCachedSheetData(CONFIG.SHEETS.AUDIT_LOG);

    if (auditSheetData.length <= 1) {
      return { success: true, data: [], message: 'No audit history available.' };
    }

    const headers = auditSheetData[0];
    const jlidColIndex = headers.indexOf('JLID');
    const learnerColIndex = headers.indexOf('Learner');

    if (jlidColIndex === -1 || learnerColIndex === -1) {
      return { success: false, message: 'Audit Log sheet is missing the "JLID" or "Learner" column.' };
    }

    const normalizedSearchTerm = searchTerm.toLowerCase().replace(/\s/g, '');
    const numericSearchTerm = searchTerm.replace(/\D/g, '');

    const history = auditSheetData.slice(1).filter(row => {
      if (row.length <= Math.max(jlidColIndex, learnerColIndex)) {
        return false;
      }

      const rowJlid = String(row[jlidColIndex] || '');
      const rowLearner = String(row[learnerColIndex] || '');

      const normalizedRowLearner = rowLearner.toLowerCase().replace(/\s/g, '');
      
      if (normalizedRowLearner.includes(normalizedSearchTerm)) {
        return true;
      }

      const numericRowJlid = rowJlid.replace(/\D/g, '');
      if (numericSearchTerm && numericRowJlid && numericRowJlid === numericSearchTerm) {
        return true;
      }
      
      return false; 
    });

    history.sort((a, b) => new Date(a[0]).getTime() - new Date(b[0]).getTime());

    return { success: true, data: history, headers: headers };

  } catch (error) {
    Logger.log('Error in getRawLearnerAuditLog: ' + error.message);
    return { success: false, message: 'Failed to retrieve raw learner audit log: ' + error.message };
  }
}

function summarizeLearnerHistory(learnerData, auditLogTimeline) {
  Logger.log('summarizeLearnerHistory called.');
  
  const GOOGLE_API_KEY = PropertiesService.getScriptProperties().getProperty('GOOGLE_GENERATIVE_AI_KEY');
  if (!GOOGLE_API_KEY) {
    Logger.log('GOOGLE_GENERATIVE_AI_KEY not configured in Script Properties.');
    return { success: false, message: 'AI summarization failed: API key not configured.' };
  }

  const model = 'gemini-2.5-flash'; 
  const endpoint = `https://generativelanguage.googleapis.com/v1beta/${model}:generateContent?key=${GOOGLE_API_KEY}`;

  let prompt = `Summarize the following learner's educational journey and key changes.
  Focus on identifying the learner's initial setup, any course changes, teacher changes, and migration reasons.
  The summary should be concise, clear, and easy to understand for an operations team member.
  
  Learner Information:
  - Learner Name: ${learnerData.learnerName || 'N/A'}
  - JLID: ${learnerData.jlid || 'N/A'}
  - Age: ${learnerData.age || 'N/A'}
  - Current Course: ${learnerData.course || 'N/A'}
  - Current Teacher: ${learnerData.currentTeacher || 'N/A'}
  - Onboarding Date: ${learnerData.startingDate || 'N/A'}

  Chronological History of Events:
  `;

  if (auditLogTimeline && auditLogTimeline.length > 0) {
    auditLogTimeline.forEach(entry => {
      let cleanedDescription = entry.description
        .replace(/Action: (.*?)\. /, '')
        .replace(/Status: (.*?)\./, '')
        .trim();
      if (cleanedDescription.includes("Reason:")) {
        cleanedDescription = cleanedDescription.substring(cleanedDescription.indexOf("Reason:"));
      }
      prompt += `- ${entry.timestamp}: ${cleanedDescription}\n`;
    });
  } else {
    prompt += `- No detailed historical events found in the audit log.\n`;
  }

  prompt += `\nBased on the above, provide a summary of ${learnerData.learnerName}'s educational journey.`;

  const requestBody = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: 500,
    },
  };

  try {
    const response = callGenerativeAIWithRetry(endpoint, requestBody); // Use the new retry wrapper
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      const textPart = jsonResponse?.candidates?.[0]?.content?.parts?.[0]?.text;
      if (textPart) {
        return { success: true, summary: textPart };
      } else {
        Logger.log('AI API response missing expected text content or malformed: ' + responseBody);
        return { success: false, message: 'AI response malformed or empty summary received.' };
      }
    } else {
      Logger.log(`AI API Error (${responseCode}): ${responseBody}`);
      return { success: false, message: `AI API Error (${responseCode}): ${responseBody}` };
    }
  } catch (error) {
    Logger.log('Error calling AI API: ' + error.message + ' Stack: ' + error.stack);
    return { success: false, message: `Failed to connect to AI service: ${error.message}` };
  }
}

function getComprehensiveLearnerHistory(jlid) {
  Logger.log('getComprehensiveLearnerHistory called for JLID: ' + jlid);

  let hubspotData = {}; 
  let auditLogTimeline = [];
  let migrationCount = 0;
  const migrationReasons = {};
  const teacherChangeEvents = [];
  const courseChangeEvents = [];
  let aiSummary = 'No summary available.';
  let overallSuccess = true;
  let overallMessage = 'Learner history fetched successfully.';
  let journeyAnalysis = null; // New Object for Stats

  try {
    if (!jlid || jlid.trim() === '') return { success: false, message: 'JLID is required.' };

    // 1. Fetch HubSpot Profile
    const hubspotResult = fetchHubspotByJlid(jlid);
    if (hubspotResult.success) {
      hubspotData = hubspotResult.data;
      
      // 1a. Fetch External Timeline (Notes/Tickets)
      if (hubspotData.dealId) {
        const externalHistory = fetchHubspotHistory(hubspotData.dealId);
        auditLogTimeline = auditLogTimeline.concat(externalHistory);
      }

      // 1b. Fetch Journey Stability Stats (NEW)
      const ticketStats = getMigrationHistoryStats(jlid);
      
      // Calculate Risk
      let riskLevel = "Low";
      let riskMessage = "Healthy Journey";
      let monthsActive = 1;
      
      if (hubspotData.startingDate) {
          const start = new Date(hubspotData.startingDate);
          const now = new Date();
          monthsActive = (now.getFullYear() - start.getFullYear()) * 12 + (now.getMonth() - start.getMonth());
          if (monthsActive < 1) monthsActive = 1;
      }
      
      const avgTenure = monthsActive / (ticketStats.total + 1);

      if (ticketStats.outbound >= 2 && monthsActive <= 6) {
          riskLevel = "Critical";
          riskMessage = "🚨 Frequent Disruptions: 2+ JetLearn-initiated changes in < 6 months.";
      } else if (ticketStats.total >= 3 && avgTenure < 4) {
          riskLevel = "High";
          riskMessage = "⚠️ High Churn Risk: Avg tenure is less than 4 months.";
      } else if (ticketStats.total >= 3) {
          riskLevel = "Medium";
          riskMessage = "Frequent changes detected.";
      }

      journeyAnalysis = {
          totalMigrations: ticketStats.total,
          inbound: ticketStats.inbound,
          outbound: ticketStats.outbound,
          riskLevel: riskLevel,
          riskMessage: riskMessage,
          ticketDetails: ticketStats.events
      };

    } else {
      hubspotData = { jlid: jlid, learnerName: 'N/A', isPartial: true };
      overallSuccess = false;
      overallMessage = `HubSpot data missing: ${hubspotResult.message}`;
    }

    // 2. Fetch Local Audit Log
    const rawAuditLogResult = getRawLearnerAuditLog(jlid);
    if (rawAuditLogResult.success && rawAuditLogResult.data.length > 0) {
      const headers = rawAuditLogResult.headers;
      const actionCol = headers.indexOf('Action');
      const timestampCol = headers.indexOf('Timestamp');
      const oldTeacherCol = headers.indexOf('Old Teacher');
      const newTeacherCol = headers.indexOf('New Teacher');
      const courseCol = headers.indexOf('Course');
      const notesCol = headers.indexOf('Notes');
      const statusCol = headers.indexOf('Status');
      const reasonCol = headers.indexOf('Reason for Migration');
      
      rawAuditLogResult.data.forEach(row => {
        const timestamp = parseSheetDate(row[timestampCol]);
        if (!timestamp) return;

        const action = String(row[actionCol] || '').trim();
        const oldTeacher = String(row[oldTeacherCol] || '').trim();
        const newTeacher = String(row[newTeacherCol] || '').trim();
        const course = String(row[courseCol] || '').trim();
        const reason = String(row[reasonCol] || '').trim();
        const notes = String(row[notesCol] || '').trim();
        const status = String(row[statusCol] || 'Unknown').trim();
        const finalReason = reason || notes;

        let entryDescription = `Action: ${action}. Notes: ${notes}. Status: ${status}.`;
        let eventType = 'general';

        if (action.includes('Migration')) {
          migrationCount++;
          eventType = 'migration';
          entryDescription = `Learner migrated from ${oldTeacher || 'N/A'} to ${newTeacher || 'N/A'}. Reason: ${finalReason || 'N/A'}. Status: ${status}.`;
          if (finalReason) migrationReasons[finalReason] = (migrationReasons[finalReason] || 0) + 1;
        } else if (action.includes('Onboarding')) {
          eventType = 'onboarding';
          entryDescription = `Learner onboarded with Teacher: ${newTeacher || 'N/A'}. Status: ${status}.`;
        }
        
        auditLogTimeline.push({
          timestamp: timestamp.toISOString(),
          type: eventType,
          description: entryDescription,
        });
      });
    }

    // 3. Sort Combined Timeline
    auditLogTimeline.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp)); // Newest first for UI logic, though timeline usually needs oldest first, we'll reverse in UI if needed. Actually standard is oldest first.
    auditLogTimeline.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp)); // Correct: Oldest First

    // 4. Generate AI Summary
    if (hubspotData && (hubspotData.learnerName !== 'N/A' || auditLogTimeline.length > 0)) {
      const aiResult = summarizeLearnerHistory(hubspotData, auditLogTimeline);
      aiSummary = aiResult.success ? aiResult.summary : `AI summary unavailable: ${aiResult.message}`;
    }

    return {
      success: overallSuccess,
      learnerProfile: hubspotData,
      auditLogTimeline: auditLogTimeline,
      journeyAnalysis: journeyAnalysis, // <--- Passing the new object
      migrationStats: {
        count: migrationCount,
        reasons: migrationReasons,
        teacherChanges: teacherChangeEvents,
        courseChanges: courseChangeEvents
      },
      aiSummary: aiSummary,
      message: overallMessage
    };

  } catch (error) {
    Logger.log('Error in getComprehensiveLearnerHistory: ' + error.message);
    return { success: false, message: 'Server error: ' + error.message };
  }
}

function getTasks() {
    Logger.log('getTasks called');
    try {
        const sheetData = _getCachedSheetData(CONFIG.SHEETS.TASKS); 
        if (sheetData.length === 0) { 
            const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.TASKS);
            sheet.appendRow(['Task ID', 'Created', 'Learner JLID', 'Learner Name', 'Task Description', 'Assigned To', 'Status', 'Due Date', 'Notes']);
            return []; 
        }
        if (sheetData.length <= 1) return []; 

        const headers = sheetData[0];
        const tasks = sheetData.slice(1).map(row => {
            const task = {};
            headers.forEach((header, i) => {
                task[header.toLowerCase().replace(/\s/g, '')] = row[i];
            });
            return {
                id: task.taskid,
                created: task.created,
                learnerJlid: task.learnerjlid,
                learner: task.learnername,
                description: task.taskdescription,
                assignedTo: task.assignedto,
                status: task.status,
                dueDate: task.duedate,
                notes: task.notes
            };
        });
        return tasks.sort((a,b) => new Date(b.created) - new Date(a.created)); 
    } catch (error) {
        Logger.log('Error getting tasks: ' + error.message);
        return [];
    }
}

function updateTaskStatus(taskId, newStatus) {
    Logger.log(`updateTaskStatus called for ${taskId} to status ${newStatus}`);
    try {
        const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.TASKS);
        const data = _getCachedSheetData(CONFIG.SHEETS.TASKS); 
        const headers = data[0].map(h => h.toLowerCase().replace(/\s/g, ''));
        const taskIdCol = headers.indexOf('taskid');
        const statusCol = headers.indexOf('status');

        if (taskIdCol === -1 || statusCol === -1) {
            throw new Error('Tasks sheet is missing required columns (Task ID, Status).');
        }

        let updated = false;
        for (let i = 0; i < data.length; i++) { 
            if (data[i][taskIdCol] && String(data[i][taskIdCol]).trim() === String(taskId).trim()) {
                sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
                updated = true;
                logAction('Task Updated', '', '', '', '', '', 'Success', `Task ${taskId} status changed to ${newStatus}`);
                delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.TASKS}`];
                break;
            }
        }
        if (!updated) {
            throw new Error(`Task with ID ${taskId} not found.`);
        }
        return { success: true, message: 'Task status updated.' };
    } catch (error) {
        Logger.log('Error updating task status: ' + error.message);
        logAction('Task Update Failed', '', '', '', '', '', 'Failed', `Failed to update task ${taskId}: ${error.message}`);
        return { success: false, message: 'Failed to update task: ' + error.message };
    }
}
