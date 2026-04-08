function calculateDashboardStats() {
  const stats = {
    migrations: { total: 0, successful: 0, failed: 0, today: 0, thisWeek: 0, thisMonth: 0, successRate: 0 },
    onboardings: { total: 0, today: 0, thisWeek: 0, thisMonth: 0 },
    recentActivities:[]
  };

  try {
    const rawSheetData = _getCachedSheetData(CONFIG.SHEETS.AUDIT_LOG);
    const auditData = (rawSheetData && Array.isArray(rawSheetData) && rawSheetData.length > 1) ? rawSheetData.slice(1) :[];

    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const thisWeekStart = new Date(now.getTime() - (now.getDay()) * 24 * 60 * 60 * 1000); 
    thisWeekStart.setHours(0,0,0,0);
    const thisMonthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    
    let combinedActivities =[];
    const seenOnboardJlids = new Set(); // deduplicate by JLID

    auditData.forEach(row => {
        const timestamp = parseSheetDate(row[0]);
        if (!timestamp || isNaN(timestamp.getTime())) return;

        const action = String(row[1]);
        const jlid   = String(row[2] || '').trim().toUpperCase();
        const status = String(row[7]);
        const learner = row[3];
        const newTeacher = row[5];
        const oldTeacher = row[4];

        const isOnboarding = action.includes('New Learner Onboarded') || action.includes('Email Sent (Onboarding');

        if (action.includes('Migration')) {
          stats.migrations.total++;
          if (status === 'Success') stats.migrations.successful++;
          else stats.migrations.failed++;
          
          if (timestamp >= today) stats.migrations.today++;
          if (timestamp >= thisWeekStart) stats.migrations.thisWeek++;
          if (timestamp >= thisMonthStart) stats.migrations.thisMonth++;
          
          combinedActivities.push({
            timestamp: timestamp,
            type: 'migration',
            title: `Migration ${status}`,
            description: `${learner} migrated from ${oldTeacher || 'N/A'} to ${newTeacher}`,
            status: status
          });
        } 
        else if (isOnboarding) {
          if (status === 'Success' || status === 'Partial Success') {
            if (!jlid || !seenOnboardJlids.has(jlid)) {
              if (jlid) seenOnboardJlids.add(jlid);
              stats.onboardings.total++;
              if (timestamp >= today) stats.onboardings.today++;
              if (timestamp >= thisWeekStart) stats.onboardings.thisWeek++;
              if (timestamp >= thisMonthStart) stats.onboardings.thisMonth++;
            }
          }
          combinedActivities.push({
            timestamp: timestamp,
            type: 'onboarding',
            title: 'New Learner Onboarded',
            description: `${learner} ${newTeacher ? 'assigned to ' + newTeacher : 'onboarded'}`,
            status: status 
          });
        }
    });

    if (stats.migrations.total > 0) {
      stats.migrations.successRate = Math.round((stats.migrations.successful / stats.migrations.total) * 100);
    }
    
    try {
      const emailLogs = _getCachedSheetData(CONFIG.SHEETS.EMAIL_LOGS);
      if (emailLogs && Array.isArray(emailLogs) && emailLogs.length > 1) {
          const headers = emailLogs[0];
          const trackingIdCol = headers.findIndex(h => h.trim() === 'Tracking ID'); 
          const statusCol = headers.findIndex(h => h.trim() === 'Status');
          const recipientCol = headers.findIndex(h => h.trim() === 'Recipient');
          const subjectCol = headers.findIndex(h => h.trim() === 'Subject');
          const sentAtCol = headers.findIndex(h => h.trim() === 'Sent At');
          const openedAtCol = headers.findIndex(h => h.trim() === 'Opened At');
          const repliedAtCol = headers.findIndex(h => h.trim() === 'Replied At');

          emailLogs.slice(1).forEach(row => {
              if (statusCol === -1 || trackingIdCol === -1) return;
              const status = row[statusCol];
              if (status === 'Sent' || status === 'Opened' || status === 'Replied') {
                  let timestamp;
                  if (status === 'Opened' && openedAtCol > -1 && row[openedAtCol]) timestamp = parseSheetDate(row[openedAtCol]);
                  else if (status === 'Replied' && repliedAtCol > -1 && row[repliedAtCol]) timestamp = parseSheetDate(row[repliedAtCol]);
                  else if (sentAtCol > -1) timestamp = parseSheetDate(row[sentAtCol]);
                  
                  if (!timestamp) return;

                  combinedActivities.push({
                      timestamp: timestamp,
                      type: 'email',
                      title: `Email ${status}`,
                      description: `To: ${recipientCol > -1 ? row[recipientCol] : 'Unknown'} | Subject: "${subjectCol > -1 ? row[subjectCol] : 'No Subject'}"`,
                      status: (status === 'Replied' ? 'Info' : (status === 'Sent' ? 'Skipped' : 'Success')),
                      trackingId: row[trackingIdCol]
                  });
              }
          });
      }
    } catch (e) { Logger.log('Error processing Email Logs: ' + e.message); }
    
    combinedActivities.sort((a, b) => b.timestamp - a.timestamp);
    
    stats.recentActivities = combinedActivities.slice(0, 7).map(activity => {
        let timeStr = null; 
        try { if (activity.timestamp instanceof Date && !isNaN(activity.timestamp)) timeStr = activity.timestamp.toISOString(); } catch (e) {}
        return { ...activity, timestamp: timeStr };
    });

    return stats;

  } catch (error) {
    Logger.log('CRITICAL Error in calculateDashboardStats: ' + error.message);
    return stats;
  }
}


// ── Public function called by the frontend ───────────────────────────────────
// activeLearners  = real HubSpot total (deals with active status)
// onboardings.today/thisWeek/thisMonth = audit log new-onboarding events
// These are intentionally separate: total is a live snapshot, cadence tracks new additions
function getDashboardStatistics() {
  try {
    const stats = calculateDashboardStats();

    try {
      const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
      if (!token) throw new Error('HUBSPOT_API_KEY not set');

      const body = {
        filterGroups: [{ filters: [{ propertyName: 'learner_status', operator: 'IN', values: ['Active Learner', 'Friendly Learner', 'VIP', 'Break & Return'] }] }],
        properties: ['hs_object_id'],
        limit: 1
      };
      const response = UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/objects/deals/search', {
        method: 'post',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify(body),
        muteHttpExceptions: true
      });
      const data = JSON.parse(response.getContentText());
      if (data.error || data.status === 'error') throw new Error(data.message || 'HubSpot error');
      const hsTotal = data.total || 0;
      Logger.log('[getDashboardStatistics] HubSpot active total: ' + hsTotal);
      stats.activeLearners = hsTotal > 0 ? hsTotal : stats.onboardings.total;
    } catch (e) {
      Logger.log('[getDashboardStatistics] HubSpot call failed: ' + e.message);
      stats.activeLearners = stats.onboardings.total;
    }

    return stats;
  } catch (e) {
    Logger.log('[getDashboardStatistics] Fatal error: ' + e.message);
    return { migrations: { total:0,successful:0,failed:0,today:0,thisWeek:0,thisMonth:0,successRate:0 }, onboardings: { total:0,today:0,thisWeek:0,thisMonth:0 }, recentActivities:[], activeLearners:0 };
  }
}

function getMigrationTrends(days = 30) {
  try {
    const auditData = getAuditLog({ limit: 1000 }).data || [];
    const trends = {};
    const now = new Date();

    for (let i = 0; i < days; i++) {
      const date = new Date(now);
      date.setDate(date.getDate() - i);
      const dateStr = date.toISOString().split('T')[0];
      trends[dateStr] = { successful: 0, failed: 0 };
    }

    auditData.forEach(row => {
      if (row[1] && String(row[1]).includes('Migration')) { 
        const date = new Date(row[0]);
        const dateStr = date.toISOString().split('T')[0];
        if (trends[dateStr]) {
          if (row[7] === 'Success') trends[dateStr].successful++;
          else trends[dateStr].failed++;
        }
      }
    });

    return Object.entries(trends)
      .sort(([a], [b]) => new Date(a) - new Date(b))
      .map(([date, data]) => ({ date, successful: data.successful, failed: data.failed, total: data.successful + data.failed }));
  } catch (error) { return []; }
}


function getMigrationTrends(days = 30) {
  Logger.log('getMigrationTrends called for ' + days + ' days');

  try {
    const auditData = getAuditLog({ limit: 1000 }).data; 
    const trends = {};
    const now = new Date();

    for (let i = 0; i < days; i++) {
      const date = new Date(now);
      date.setDate(date.getDate() - i);
      const dateStr = date.toISOString().split('T')[0];
      trends[dateStr] = { successful: 0, failed: 0 };
    }

    auditData.forEach(row => {
      if (row[1] && String(row[1]).includes('Migration')) { 
        const date = new Date(row[0]);
        const dateStr = date.toISOString().split('T')[0];

        if (trends[dateStr]) {
          if (row[7] === 'Success') {
            trends[dateStr].successful++;
          } else {
            trends[dateStr].failed++;
          }
        }
      }
    });

    const trendArray = Object.entries(trends)
      .sort(([a], [b]) => new Date(a) - new Date(b))
      .map(([date, data]) => ({
        date: date,
        successful: data.successful,
        failed: data.failed,
        total: data.successful + data.failed
      }));

    Logger.log('Migration trends calculated for ' + days + ' days');
    return trendArray;
  }
   catch (error) {
    Logger.log('Error calculating migration trends: ' + error.message);
    return [];
  }
}

//Monthly & AI Reports:
function getEnhancedMigrationReport(params) {
    try {
        const toDate = new Date(params.toDate || params.endDate);
        toDate.setHours(23, 59, 59, 999);
        const fromDate = new Date(params.fromDate || params.startDate);
        fromDate.setHours(0, 0, 0, 0);

        const MIGRATION_PIPELINE_ID = '66161281';
        const EXCLUDED_STAGE_IDS = ['133821818', '153457301']; 

        const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
        const searchUrl = 'https://api.hubapi.com/crm/v3/objects/tickets/search';

        const fetchExecutedTickets = (start, end) => {
            let allTickets = [];
            let after = undefined;
            do {
                const requestBody = {
                    filterGroups: [{
                        filters: [
                            { propertyName: "hs_pipeline", operator: "EQ", value: MIGRATION_PIPELINE_ID },
                            { propertyName: "migration_completed_date", operator: "BETWEEN", value: start.getTime(), highValue: end.getTime() },
                            { propertyName: "hs_pipeline_stage", operator: "NOT_IN", values: EXCLUDED_STAGE_IDS }
                        ]
                    }],
                    properties: ["reason_of_migration__t_", "migration_intervened_by", "current_teacher__t_", "new_teacher", "current_course"],
                    limit: 100,
                    after: after
                };
                const response = UrlFetchApp.fetch(searchUrl, { method: 'post', headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' }, payload: JSON.stringify(requestBody), muteHttpExceptions: true });
                const data = JSON.parse(response.getContentText());
                if (data && data.results) allTickets = allTickets.concat(data.results);
                after = data.paging && data.paging.next ? data.paging.next.after : undefined;
            } while (after);
            return allTickets;
        };

        const currentTickets = fetchExecutedTickets(fromDate, toDate);

        const processTickets = (tickets) => {
            const periodData = { totalMigrations: 0, clsInvolvement: 0, tpInvolvement: 0, opsInvolvement: 0, reasonBreakdown: {}, teamInvolvementByReason: {}, teacherMigrationsFrom: {}, teacherMigrationsTo: {}, courseMigrations: {} };
            tickets.forEach(ticket => {
                const props = ticket.properties;
                periodData.totalMigrations++;
                const reason = String(props.reason_of_migration__t_ || 'Unknown').trim();
                const teams = String(props.migration_intervened_by || '').toLowerCase();
                
                if (teams.includes('cls')) periodData.clsInvolvement++;
                if (teams.includes('tp')) periodData.tpInvolvement++;
                if (teams.includes('ops')) periodData.opsInvolvement++;

                periodData.reasonBreakdown[reason] = (periodData.reasonBreakdown[reason] || 0) + 1;
                
                // ======================= FIX IS HERE ==========================
                // This block was missing, causing the crash. It initializes and populates the team involvement data.
                if (!periodData.teamInvolvementByReason[reason]) {
                    periodData.teamInvolvementByReason[reason] = { cls: 0, tp: 0, ops: 0 };
                }
                if (teams.includes('cls')) periodData.teamInvolvementByReason[reason].cls++;
                if (teams.includes('tp')) periodData.teamInvolvementByReason[reason].tp++;
                if (teams.includes('ops')) periodData.teamInvolvementByReason[reason].ops++;
                // ===================== END OF FIX ===========================

                const oldTeacher = getTeacherLabel(props.current_teacher__t_);
                const newTeacher = getTeacherLabel(props.new_teacher);
                if (oldTeacher) periodData.teacherMigrationsFrom[oldTeacher] = (periodData.teacherMigrationsFrom[oldTeacher] || 0) + 1;
                if (newTeacher) periodData.teacherMigrationsTo[newTeacher] = (periodData.teacherMigrationsTo[newTeacher] || 0) + 1;
                
                const course = getCourseLabel(props.current_course);
                if (course) {
                   if (!periodData.courseMigrations[course]) periodData.courseMigrations[course] = { count: 0, reasons: {} };
                   periodData.courseMigrations[course].count++;
                   periodData.courseMigrations[course].reasons[reason] = (periodData.courseMigrations[course].reasons[reason] || 0) + 1;
                }
            });
            return periodData;
        };
        
        const currentPeriod = processTickets(currentTickets);
        
        const daysInPeriod = ((toDate.getTime() - fromDate.getTime()) / (1000 * 3600 * 24)) + 1;
        const kpis = {
            totalMigrations: { current: currentPeriod.totalMigrations, previous: 0, change: 0 },
            avgMigrationsPerDay: { current: (daysInPeriod > 0 ? (currentPeriod.totalMigrations / daysInPeriod) : 0), previous: 0, change: 0 },
            clsRate: { current: (currentPeriod.totalMigrations > 0 ? (currentPeriod.clsInvolvement / currentPeriod.totalMigrations) * 100 : 0), previous: 0, change: 0 },
            tpRate: { current: (currentPeriod.totalMigrations > 0 ? (currentPeriod.tpInvolvement / currentPeriod.totalMigrations) * 100 : 0), previous: 0, change: 0 },
        };
        // This logic remains the same
        const teacherImpact = Object.keys(currentPeriod.teacherMigrationsFrom).map(name => ({ name, from: currentPeriod.teacherMigrationsFrom[name] || 0, to: currentPeriod.teacherMigrationsTo[name] || 0, netFlow: (currentPeriod.teacherMigrationsTo[name] || 0) - (currentPeriod.teacherMigrationsFrom[name] || 0) })).sort((a,b) => b.from - a.from);
        const courseImpact = Object.entries(currentPeriod.courseMigrations).map(([name, data]) => ({ name, count: data.count, topReason: Object.keys(data.reasons).sort((a, b) => data.reasons[b] - data.reasons[a])[0] || 'N/A' })).sort((a, b) => b.count - a.count);
        
        const reportData = {
            kpis,
            reasonBreakdown: Object.entries(currentPeriod.reasonBreakdown).map(([name, count]) => ({ name, count })).sort((a, b) => b.count - a.count),
            teamInvolvementByReason: currentPeriod.teamInvolvementByReason, // This object is now correctly populated
            teacherImpact,
            courseImpact
        };
        
        const aiInsights = getEnhancedAIInsights(reportData, fromDate.toLocaleString('default', { month: 'long' }));

        return { success: true, data: reportData, aiInsights: aiInsights };

    } catch (error) {
        Logger.log('Error in getEnhancedMigrationReport (Ticket Version): ' + error.stack);
        return { success: false, message: 'Error generating report from tickets: ' + error.message };
    }
}

function getEnhancedAIInsights(reportData, monthName) {
    const GOOGLE_API_KEY = PropertiesService.getScriptProperties().getProperty('GOOGLE_GENERATIVE_AI_KEY');
    if (!GOOGLE_API_KEY) {
        return { executive: "AI insights unavailable: API key not configured.", rootCause: "", impact: "" };
    }

    const model = 'gemini-2.5-flash'; 
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GOOGLE_API_KEY}`;
    
    const prompt = `
      You are a top-tier Senior Operations Analyst at JetLearn. Provide a three-part analysis of migration data for ${monthName}.
      Return valid JSON with keys: "executive", "rootCause", "impact". No other text.

      Data:
      ${JSON.stringify(reportData)}

      1. executive: High-level summary for leadership.
      2. rootCause: Analysis of top migration reasons.
      3. impact: Actionable analysis on teachers/courses.
    `;

    try {
        const payload = { contents: [{ parts:[{ text: prompt }] }] };
        const response = callGenerativeAIWithRetry(endpoint, payload); 
        if (response.getResponseCode() === 200) {
            let textPart = JSON.parse(response.getContentText())?.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
            textPart = textPart.replace(/```json/g, '').replace(/```/g, '').trim();
            return JSON.parse(textPart);
        }
        return { executive: "AI Error", rootCause: "", impact: "" };
    } catch (error) {
        return { executive: `AI Error: ${error.message}`, rootCause: "", impact: "" };
    }
}



function getEnhancedAIInsights(reportData, monthName) {
    const GOOGLE_API_KEY = PropertiesService.getScriptProperties().getProperty('GOOGLE_GENERATIVE_AI_KEY');
    if (!GOOGLE_API_KEY) {
        return { executive: "AI insights unavailable: API key not configured.", rootCause: "", impact: "" };
    }

    const model = 'gemini-2.5-flash'; 
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GOOGLE_API_KEY}`;
    
    const prompt = `
      You are a top-tier Senior Operations Analyst at JetLearn. Provide a three-part analysis of migration data for ${monthName}.
      Return valid JSON with keys: "executive", "rootCause", "impact". No other text.

      Data:
      ${JSON.stringify(reportData)}

      1. executive: High-level summary for leadership.
      2. rootCause: Analysis of top migration reasons.
      3. impact: Actionable analysis on teachers/courses.
    `;

    try {
        const payload = { contents: [{ parts: [{ text: prompt }] }] };
        const response = callGenerativeAIWithRetry(endpoint, payload); 
        if (response.getResponseCode() === 200) {
            let textPart = JSON.parse(response.getContentText())?.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
            textPart = textPart.replace(/```json/g, '').replace(/```/g, '').trim();
            return JSON.parse(textPart);
        }
        return { executive: "AI Error", rootCause: "", impact: "" };
    } catch (error) {
        return { executive: `AI Error: ${error.message}`, rootCause: "", impact: "" };
    }
}



function generateMonthlyReport(month, year, perspective = 'All') {
    const start = new Date(year, month, 1);
    const end = new Date(year, month + 1, 0);
    return getEnhancedMigrationReport({ fromDate: start, toDate: end });
}


function emailMonthlyReport() {
    Logger.log('Starting monthly report email process.');
    const now = new Date();
    
    const lastMonthDate = new Date(now.getFullYear(), now.getMonth(), 0);
    const lastMonth = lastMonthDate.getMonth();
    const lastMonthYear = lastMonthDate.getFullYear();

    const prevMonthDate = new Date(lastMonthYear, lastMonth, 0);
    const prevMonth = prevMonthDate.getMonth();
    const prevMonthYear = prevMonthDate.getFullYear();

    const currentMonthReport = generateMonthlyReport(lastMonth, lastMonthYear, 'All');
    const previousMonthReport = generateMonthlyReport(prevMonth, prevMonthYear, 'All');
    
    const aiInsightsJsonString = getAIGeneratedInsights(currentMonthReport, previousMonthReport, 'All');
    let aiInsightsHtml = '<li>Could not generate AI insights.</li>';
    try {
      const insightsArray = JSON.parse(aiInsightsJsonString);
      if(Array.isArray(insightsArray)) {
        aiInsightsHtml = insightsArray.map(item => `<li>${item}</li>`).join('');
      }
    } catch(e) {
      Logger.log("Could not parse AI insights JSON for email: " + e.message);
    }

    const webAppUrl = ScriptApp.getService().getUrl();
    const reportMonthParam = `${lastMonthYear}-${String(lastMonth + 1).padStart(2, '0')}`; 
    const reportUrl = `${webAppUrl}?page=report&month=${reportMonthParam}`;

    const subject = `Migration Report & Insights - ${currentMonthReport.month} ${lastMonthYear}`;
    const htmlBody = `
        <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <h2>Migration Report - ${currentMonthReport.month} ${lastMonthYear}</h2>
            <p>The automated migration report for ${currentMonthReport.month} is now available.</p>
            
            <h3>AI-Powered Insights This Month (Overall):</h3>
            <ul style="padding-left: 20px;">
                ${aiInsightsHtml}
            </ul>
            
            <p style="margin-top: 25px; margin-bottom: 25px;">
                <a href="${reportUrl}" 
                   style="display: inline-block; padding: 12px 20px; background-color: #4a3c8a; color: white; text-decoration: none; border-radius: 5px; font-size: 16px; font-weight: bold;">
                   View Full Interactive Report
                </a>
            </p>
            <hr style="border: none; border-top: 1px solid #eee;">
            <p style="font-size: 12px; color: #777;">
                You are receiving this email as a designated recipient for monthly system reports. The link above provides a detailed, interactive view of the data with options to filter by team perspective.
            </p>
        </div>
    `;

    MailApp.sendEmail({
        to: CONFIG.EMAIL.REPORT_RECIPIENTS,
        subject: subject,
        htmlBody: htmlBody,
        name: CONFIG.EMAIL.FROM_NAME,
        from: CONFIG.EMAIL.FROM
    });

    Logger.log(`Monthly report notification email sent successfully to ${CONFIG.EMAIL.REPORT_RECIPIENTS}.`);
}

function createMonthlyTrigger() {
    const existingTriggers = ScriptApp.getProjectTriggers();
    existingTriggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'emailMonthlyReport') {
            ScriptApp.deleteTrigger(trigger);
            Logger.log('Deleted existing trigger for emailMonthlyReport.');
        }
    });

    ScriptApp.newTrigger('emailMonthlyReport')
        .timeBased()
        .onMonthDay(1)
        .atHour(8)
        .create();

    Logger.log('Monthly report trigger created successfully.');
}

//Exports & Subscriptions:
function exportReportAsCsv(params) {
  Logger.log(`exportReportAsCsv called with: ${JSON.stringify(params)}`);
  try {
    const reportResult = getEnhancedMigrationReport(params);
    if (!reportResult.success || !reportResult.data) {
      throw new Error(reportResult.message || "No data to export.");
    }
    
    const { kpis, reasonBreakdown, teacherImpact, courseImpact } = reportResult.data;
    const csvDataArray = [];

    csvDataArray.push(['Key Metrics']);
    csvDataArray.push(['Metric', 'Current Period', 'Previous Period', 'Change (%)']);
    csvDataArray.push(['Total Migrations', kpis.totalMigrations.current, kpis.totalMigrations.previous, kpis.totalMigrations.change.toFixed(2)]);
    csvDataArray.push(['CLS Involvement Rate', kpis.clsRate.current.toFixed(2), kpis.clsRate.previous.toFixed(2), kpis.clsRate.change.toFixed(2)]);
    csvDataArray.push(['TP Involvement Rate', kpis.tpRate.current.toFixed(2), kpis.tpRate.previous.toFixed(2), kpis.tpRate.change.toFixed(2)]);
    csvDataArray.push([]);

    csvDataArray.push(['Reason Breakdown']);
    csvDataArray.push(['Reason', 'Count']);
    reasonBreakdown.forEach(row => csvDataArray.push([row.name, row.count]));
    csvDataArray.push([]);

    csvDataArray.push(['Teacher Impact Analysis']);
    csvDataArray.push(['Teacher Name', 'Migrations From', 'Migrations To', 'Net Flow']);
    teacherImpact.forEach(row => csvDataArray.push([row.name, row.from, row.to, row.netFlow]));
    csvDataArray.push([]);

    csvDataArray.push(['Course Impact Analysis']);
    csvDataArray.push(['Course Name', 'Migration Count', 'Top Reason']);
    courseImpact.forEach(row => csvDataArray.push([row.name, row.count, row.topReason]));

    return { success: true, data: csvDataArray };
  } catch (error) {
    Logger.log('Error exporting report as CSV: ' + error.message);
    return { success: false, message: 'Failed to export as CSV: ' + error.message };
  }
}

function exportReportAsPdf(params) {
  Logger.log(`exportReportAsPdf called with: ${JSON.stringify(params)}`);
  try {
    const reportResult = getEnhancedMigrationReport(params);
     if (!reportResult.success || !reportResult.data) {
      throw new Error(reportResult.message || "No data for PDF export.");
    }
    
    const template = HtmlService.createTemplateFromFile('ReportExportTemplate');
    template.reportData = reportResult.data;
    template.aiInsights = reportResult.aiInsights;
    template.period = `${params.fromDate} to ${params.toDate}`;
    const html = template.evaluate().getContent();
    
    const pdfBlob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF);
    const base64Data = Utilities.base64Encode(pdfBlob.getBytes());
    
    const filename = `Migration_Report_${params.fromDate}_to_${params.toDate}.pdf`;
    
    return { success: true, data: base64Data, filename: filename };
  } catch (error) {
    Logger.log('Error exporting report as PDF: ' + error.message);
    return { success: false, message: 'Failed to export as PDF: ' + error.message };
  }
}

function emailReportNow(params) {
    Logger.log('emailReportNow called');
    try {
        const reportResult = getEnhancedMigrationReport(params);
        if (!reportResult.success || !reportResult.data) {
            throw new Error(reportResult.message || "Could not generate report data to email.");
        }

        const { kpis, reasonBreakdown } = reportResult.data;
        const { aiInsights } = reportResult; 
        const executiveSummary = aiInsights?.executive || "AI summary could not be generated for this period.";
        
        const topReason = reasonBreakdown[0] ? `${reasonBreakdown[0].name} (${reasonBreakdown[0].count} cases)` : 'N/A';

        const subject = `Migration Intelligence Report: ${params.fromDate} to ${params.toDate}`;
        const htmlBody = `
            <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 700px; margin: auto; border: 1px solid #ddd; padding: 20px;">
                <h2 style="color: #4a3c8a;">Migration Intelligence Summary</h2>
                <p><strong>Period:</strong> ${params.fromDate} to ${params.toDate}</p>
                <hr style="border: none; border-top: 1px solid #eee;">
                
                <h3 style="color: #4a3c8a;">AI Executive Summary:</h3>
                <p style="background-color: #f8f9ff; padding: 15px; border-left: 4px solid #6b5bae; margin: 0;">${executiveSummary}</p>

                <h3 style="color: #4a3c8a; margin-top: 25px;">Key Metrics:</h3>
                <table style="width: 100%; border-collapse: collapse;">
                    <tr style="background-color: #f2f2f2;">
                        <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Metric</th>
                        <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Value</th>
                        <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Change vs. Prev. Period</th>
                    </tr>
                    <tr>
                        <td style="padding: 8px; border: 1px solid #ddd;">Total Migrations</td>
                        <td style="padding: 8px; border: 1px solid #ddd;">${kpis.totalMigrations.current}</td>
                        <td style="padding: 8px; border: 1px solid #ddd; color: ${kpis.totalMigrations.change >= 0 ? 'red' : 'green'};">${kpis.totalMigrations.change.toFixed(1)}%</td>
                    </tr>
                    <tr>
                        <td style="padding: 8px; border: 1px solid #ddd;">Top Migration Reason</td>
                        <td style="padding: 8px; border: 1px solid #ddd;" colspan="2">${topReason}</td>
                    </tr>
                </table>

                <p style="text-align: center; margin-top: 25px;">
                    <a href="${ScriptApp.getService().getUrl()}" 
                       style="display: inline-block; padding: 12px 20px; background-color: #4a3c8a; color: white; text-decoration: none; border-radius: 5px; font-size: 16px;">
                       View Full Interactive Dashboard
                    </a>
                </p>
            </div>
        `;

        MailApp.sendEmail({
            to: CONFIG.EMAIL.REPORT_RECIPIENTS,
            subject: subject,
            htmlBody: htmlBody,
            name: CONFIG.EMAIL.FROM_NAME,
            from: CONFIG.EMAIL.FROM
        });

        Logger.log(`Report email sent successfully to ${CONFIG.EMAIL.REPORT_RECIPIENTS}.`);
        return { success: true, message: `Report successfully sent to ${CONFIG.EMAIL.REPORT_RECIPIENTS}.` };

    } catch (error) {
        Logger.log('Error in emailReportNow: ' + error.message);
        return { success: false, message: 'Failed to send report email: ' + error.message };
    }
}

function subscribeToWeeklyReport() {
    Logger.log('subscribeToWeeklyReport called');

    const userEmail = Session.getActiveUser().getEmail(); 

    Logger.log(`User ${userEmail} requested subscription to weekly report.`);
    return { success: true, message: `Subscription request received for ${userEmail}. (Feature not fully implemented)` };
}

function testApiCall() {
  const rates = getLiveCurrencyRates();
  console.log(rates);
}

function getAIGeneratedInsights(currentData, previousData) {
    Logger.log(`Getting AI insights for monthly report.`);
    const GOOGLE_API_KEY = PropertiesService.getScriptProperties().getProperty('GOOGLE_GENERATIVE_AI_KEY');
    if (!GOOGLE_API_KEY) {
        Logger.log('GOOGLE_GENERATIVE_AI_KEY not configured in Script Properties.');
        return '["AI insights are unavailable: API key not configured."]';
    }

    const model = 'gemini-2.5-flash';
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/${model}:generateContent?key=${GOOGLE_API_KEY}`;

    const prompt = `
      You are a senior operations analyst for JetLearn, providing an executive summary on monthly learner migration trends.
      Your response must be a valid JSON array of strings, where each string is a key insight. Do not include any text outside of the JSON array.
      Analyze the provided data which compares team involvement in migrations over the last two months.

      Your analysis must focus on:
      1.  **Intervention Change:** What is the most significant change in team involvement (CLS, TP, or Ops) this month compared to last month? Mention the percentage point change and the absolute numbers.
      2.  **Top Driver Analysis:** Identify the top migration driver for the current month and state which team was most involved in handling cases with that reason.
      3.  **Actionable Recommendation:** Based on the data, provide one specific, actionable recommendation for operational improvement. For example, "Given that TP involvement in 'Attrition' cases rose by 15%, a review of teacher-related attrition factors is recommended."

      Data for Current Month (${currentData.month} ${currentData.year}):
      ${JSON.stringify(currentData)}

      Data for Previous Month (${previousData.month} ${previousData.year}):
      ${JSON.stringify(previousData)}
    `;

    try {
        const payload = { contents: [{ parts: [{ text: prompt }] }] };
        const response = callGenerativeAIWithRetry(endpoint, payload); // Use the new retry wrapper
        const responseCode = response.getResponseCode();
        const responseBody = response.getContentText();

        if (responseCode === 200) {
            const jsonResponse = JSON.parse(responseBody);
            const textPart = jsonResponse?.candidates?.[0]?.content?.parts?.[0]?.text;
            if (textPart) {
                Logger.log('AI Insights Generated: ' + textPart);
                return textPart.trim();
            }
        }
        Logger.log(`AI API Error (${responseCode}): ${responseBody}`);
        return '["AI insights could not be generated due to an API error."]';
    } catch (error) {
        Logger.log('Error calling AI API: ' + error.message);
        return `["AI insights are unavailable due to a connection error: ${error.message}"]`;
    }
}

function getDoubleMigrationReport(params) {
  try {
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.AUDIT_LOG);
    
    if (!sheetData || !Array.isArray(sheetData) || sheetData.length < 2) {
        Logger.log("Audit Log empty or null.");
        return[];
    }

    const fromDate = new Date(params.fromDate);
    fromDate.setHours(0, 0, 0, 0);
    const toDate = new Date(params.toDate);
    toDate.setHours(23, 59, 59, 999);

    const learnerMap = {};

    // 1. Gather all valid migrations
    sheetData.slice(1).forEach(row => {
      const action = String(row[1] || "");
      const status = String(row[7] || "");
      
      // Ignore non-migrations and explicit failures
      if (!action.includes("Migration") || status.trim().toLowerCase() === 'failed') return;

      const date = parseSheetDate(row[0]);
      if (!date || isNaN(date.getTime()) || date < fromDate || date > toDate) return;

      const jlid = String(row[2]).trim();
      if (!jlid) return;

      if (!learnerMap[jlid]) {
          learnerMap[jlid] = { name: row[3] || "Unknown", events: [] };
      }

      learnerMap[jlid].events.push({
        date: date, // Native date object (needed for sorting)
        dateStr: formatDateDDMMYYYY(date),
        oldTeacher: String(row[4] || "").trim(),
        newTeacher: String(row[5] || "").trim(),
        reason: String(row[10] || "Unspecified").trim()
      });
    });

    const doubleMigrations =[];

    // 2. De-duplicate and Count
    Object.keys(learnerMap).forEach(jlid => {
      const data = learnerMap[jlid];
      
      // SORT: Oldest to Newest using the native Date object
      data.events.sort((a, b) => a.date.getTime() - b.date.getTime());

      const uniqueEvents =[];
      let lastTeacher = null;

      data.events.forEach(e => {
          // If the New Teacher is exactly the same as the previous New Teacher, skip it (Retry/Duplicate)
          if (e.newTeacher !== lastTeacher) {
              uniqueEvents.push(e);
              lastTeacher = e.newTeacher;
          }
      });

      // MUST have 2 or more unique teacher changes
      if (uniqueEvents.length >= 2) {
        const path = [];
        const reasons =[];
        
        uniqueEvents.forEach((e, index) => {
           if (index === 0) path.push(e.oldTeacher || "Original");
           path.push(e.newTeacher);
           reasons.push(e.reason);
        });

        // Clean up reasons string to remove obvious repeats
        const uniqueReasons = [...new Set(reasons)].filter(r => r !== "Unspecified").join(", ") || "No reason specified";

        doubleMigrations.push({
            jlid: jlid,
            name: data.name,
            count: uniqueEvents.length,
            path: path.join(" <i class='fas fa-arrow-right' style='font-size:0.8em; color:#bbb; margin:0 5px;'></i> "),
            reasons: uniqueReasons,
            lastDate: uniqueEvents[uniqueEvents.length - 1].dateStr,
            
            // --- CRITICAL FIX IS HERE ---
            // We map over the array to REMOVE the native 'date' property before sending to the browser.
            timeline: uniqueEvents.map(e => ({
                dateStr: e.dateStr,
                oldTeacher: e.oldTeacher,
                newTeacher: e.newTeacher,
                reason: e.reason
            }))
        });
      }
    });

    return doubleMigrations.sort((a, b) => b.count - a.count);

  } catch (e) {
    Logger.log("Error in getDoubleMigrationReport: " + e.stack);
    return[];
  }
}

function getMultiMigrationReport(params) {
  try {
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.AUDIT_LOG);
    
    if (!sheetData || !Array.isArray(sheetData) || sheetData.length < 2) {
        return [];
    }

    const fromDate = new Date(params.fromDate);
    fromDate.setHours(0, 0, 0, 0);
    const toDate = new Date(params.toDate);
    toDate.setHours(23, 59, 59, 999);

    const learnerMap = {};

    // 1. Gather all migrations from the audit log
    sheetData.slice(1).forEach(row => {
      const action = String(row[1] || "");
      if (!action.includes("Migration") || String(row[7] || "").trim().toLowerCase() === 'failed') return;

      const date = parseSheetDate(row[0]);
      if (!date || isNaN(date.getTime()) || date < fromDate || date > toDate) return;

      const jlid = String(row[2]).trim().toUpperCase();
      if (!jlid) return;

      if (!learnerMap[jlid]) {
          learnerMap[jlid] = { name: row[3] || "Unknown", events: [] };
      }

      learnerMap[jlid].events.push({
        date: date,
        dateStr: formatDateDDMMYYYY(date),
        oldTeacher: String(row[4] || "").trim(),
        newTeacher: String(row[5] || "").trim(),
        reason: String(row[10] || "Unspecified").trim()
      });
    });

    const multiMigrations = [];
    const learnerJlids = Object.keys(learnerMap);

    // 2. Process each learner and enrich with HubSpot data
    for (const jlid of learnerJlids) {
      const data = learnerMap[jlid];
      data.events.sort((a, b) => a.date.getTime() - b.date.getTime());

      const uniqueEvents = data.events.filter((e, index, arr) => {
        if (index === 0) return true;
        return e.newTeacher !== arr[index - 1].newTeacher;
      });

      if (uniqueEvents.length >= 1) {
        
        // --- THIS IS THE CORE FIX ---
        // We set default "error" values first.
        let learnerStatus = 'HubSpot Error';
        let subscriptionEndDate = null;
        let pauseDate = null;
        let pausedTiming = 'N/A';

        // Then, we TRY to update them with real HubSpot data.
        try {
          const hubspotResult = HubSpotService.fetchHubspotByJlid(jlid);
          if (hubspotResult.success) {
            const hsData = hubspotResult.data;
            learnerStatus = hsData.learnerStatus || 'Status Not Found';
            subscriptionEndDate = hsData.endDate ? new Date(hsData.endDate) : null;
            pauseDate = hsData.pauseDate ? new Date(hsData.pauseDate) : null;
            
            if (learnerStatus === 'Paused Learner' && pauseDate && subscriptionEndDate) {
                pausedTiming = pauseDate.getTime() < subscriptionEndDate.getTime() ? 'Before End of Subscription' : 'At End of Subscription';
            }
          }
        } catch (hsError) {
          Logger.log(`Graceful failure for HubSpot data on ${jlid}: ${hsError.message}`);
        }
        
        // NOW, we push the learner to the report REGARDLESS of the HubSpot result.
        // The user's request to "Include in it both Active and Paused Learners" is now
        // a filter on the data we successfully get, not a hard requirement to show the row.
        multiMigrations.push({
            jlid: jlid,
            name: data.name,
            count: uniqueEvents.length,
            timeline: uniqueEvents.map(e => ({
                dateStr: e.dateStr,
                oldTeacher: e.oldTeacher,
                newTeacher: e.newTeacher,
                reason: e.reason
            })),
            learnerStatus: learnerStatus,
            subscriptionEndDate: subscriptionEndDate ? subscriptionEndDate.toLocaleDateString('en-CA') : '',
            pauseDate: pauseDate ? pauseDate.toLocaleDateString('en-CA') : '',
            pausedTiming: pausedTiming
        });
      }
    }

    return multiMigrations.sort((a, b) => b.count - a.count);

  } catch (e) {
    Logger.log("FATAL Error in getMultiMigrationReport: " + e.stack);
    return []; // Return empty on catastrophic failure
  }
}

function getMultiMigrationReportFromTickets(params) {
  try {
    const fromDate = new Date(params.fromDate);
    const toDate = new Date(params.toDate);
    const MIGRATION_PIPELINE_ID = '66161281';

    // *** The correct IDs are now here ***
    const EXCLUDED_STAGE_IDS = ['133821818', '153457301']; 

    const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    if (!token) throw new Error("HubSpot API Key is not configured.");

    const searchUrl = 'https://api.hubapi.com/crm/v3/objects/tickets/search';
    
    let allTickets = [];
    let after = undefined;

    do {
      const requestBody = {
        filterGroups: [{
          filters: [
            { propertyName: "hs_pipeline", operator: "EQ", value: MIGRATION_PIPELINE_ID },
            { propertyName: "migration_completed_date", operator: "BETWEEN", value: fromDate.getTime(), highValue: toDate.getTime() },
            { propertyName: "hs_pipeline_stage", operator: "NOT_IN", values: EXCLUDED_STAGE_IDS }
          ]
        }],
        properties: ["learner_uid", "learner_full_name", "createdate", "migration_completed_date", "current_teacher__t_", "new_teacher", "reason_of_migration__t_"],
        sorts: [{ propertyName: "migration_completed_date", direction: "ASCENDING" }],
        limit: 100,
        after: after
      };

      const response = UrlFetchApp.fetch(searchUrl, { method: 'post', headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' }, payload: JSON.stringify(requestBody), muteHttpExceptions: true });
      const data = JSON.parse(response.getContentText());
      if (data && data.results) allTickets = allTickets.concat(data.results);
      after = data.paging && data.paging.next ? data.paging.next.after : undefined;

    } while (after);

    const learnerMap = {};
    allTickets.forEach(ticket => {
        const props = ticket.properties;
        const jlid = (props.learner_uid || '').trim().toUpperCase();
        if (!jlid) return;
        if (!learnerMap[jlid]) learnerMap[jlid] = { name: props.learner_full_name || "Unknown", events: [] };
        
        learnerMap[jlid].events.push({
            date: new Date(props.migration_completed_date || ticket.createdAt),
            dateStr: formatDateDDMMYYYY(new Date(props.migration_completed_date || ticket.createdAt)),
            oldTeacher: getTeacherLabel(props.current_teacher__t_) || 'N/A',
            newTeacher: getTeacherLabel(props.new_teacher) || '[Unassigned]',
            reason: props.reason_of_migration__t_ || "Unspecified"
        });
    });
    
    const multiMigrations = [];
    for (const jlid of Object.keys(learnerMap)) {
        const data = learnerMap[jlid];
        if (data.events.length < 1) continue;
        let learnerStatus = 'HubSpot Error';
        try {
            const hubspotResult = HubSpotService.fetchHubspotByJlid(jlid);
            if (hubspotResult.success) {
                learnerStatus = hubspotResult.data.learnerStatus || 'Status Not Found';
            } else {
                learnerStatus = hubspotResult.message || 'Deal Not Found';
            }
        } catch (hsError) { /* Graceful failure */ }

        multiMigrations.push({ jlid, name: data.name, count: data.events.length, timeline: data.events, learnerStatus });
    }
    return multiMigrations.sort((a, b) => b.count - a.count);
  } catch (e) {
    Logger.log("FATAL Error in getMultiMigrationReportFromTickets: " + e.stack);
    return [];
  }
}
