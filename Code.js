// Enhanced Migration System
// Version: V26 - Refactored Structure

// =============================================
// CONFIGURATION
// =============================================
const CONFIG = {
  MIGRATION_SHEET_ID: '1xzprj2U6NpJwoevBMvM1DVfIj76wVjAd0ZcMjVC1xMM',
  PERSONA_SHEET_ID: '1rSweVyLKEwb1xThFHMLoH4xWnrLs8wbRM_61VtRjGww', 
  DRIVE_FOLDER_ID: '1K-Zb9BO2dm_dPg2AWTDT5t-ghkPoRNSW', 
  HUBSPOT_API_KEY: 'pat-na1-840cfb1a-acb3-45d6-8b0d-31f8c3f7cb34', 
  CLASS_SCHEDULE_CALENDAR_ID: 'hello@jet-learn.com',
  EXCHANGE_RATE_API_URL: 'https://v6.exchangerate-api.com/v6/YOUR_API_KEY/latest/EUR', // Ensure this is set if using live rates

  SHEETS: {
    TEACHER_DATA: 'Teacher Data',
    COURSE_NAME: 'Course Name',
    AUDIT_LOG: 'Audit Log',
    USER_PROFILES: 'User Profiles',
    TEACHER_COURSES: 'Teacher Courses',
    COURSE_PROGRESS_SUMMARY: 'Course Summary',
    USER_ACTIVITY_LOG: 'User Activity Log', 
    TASKS: 'Tasks', 
    INVOICE_PRODUCTS: 'Invoice Products',
    TEACHER_HS_DATA:'Teacher HS values', 
    COURSE_HS_DATA : 'Course HS values', 
    HS_USER_DATA : 'HS User Values', 
    PERSONA_DATA: 'Main Sheet',
    EMAIL_LOGS: 'Email Logs' 
  },
  EMAIL: {
    FROM: 'hello@jet-learn.com',
    FROM_NAME: 'JetLearn',
    MAIN_MANAGER: 'sankalita.mitra@jet-learn.com', 
    MAIN_MANAGER_NAME: 'Sankalita Mitra',
    REPORT_RECIPIENTS: 'sourav.pal@jet-learn.com',
    AUDIT_REPORT_RECIPIENTS: 'sourav.pal@jet-learn.com'
  },
  RANGES: { 
    TEACHER_DATA: 'A2:K',
    COURSE_DATA: 'A2:A',
    PERSONA_DATA: 'A1:BR',
    TEACHER_COURSES: 'A2:D',
    COURSE_PROGRESS_DATA: 'A1:L'
  },
  PAGINATION_LIMIT: 50
};

const ROLES = {
  SUPER_ADMIN: 'Super Admin',
  ADMIN: 'Admin',
  USER: 'User',
  GUEST: 'Guest'
};

const PERMISSIONS = {
  [ROLES.SUPER_ADMIN]: [
    'view_dashboard', 'send_emails', 'view_audit', 'manage_users',
    'view_reports', 'use_persona_tool', 'manage_settings',
    'send_generic_emails', 'manage_invoices', 'run_audit_center',
    'manage_agentic_audit', 'use_ai_pm',
    'create_users',        // NEW
    'send_welcome_email',  // NEW
    'send_reset_email', // NEW
  ],
  [ROLES.ADMIN]: [
    'view_dashboard', 'send_emails', 'view_audit',
    'use_persona_tool', 'view_reports', 'send_generic_emails',
    'manage_invoices', 'run_audit_center', 'manage_settings', 'use_ai_pm'
    // NOTE: No manage_users, no create_users
  ],
  [ROLES.USER]: [
    'view_dashboard', 'send_emails', 'view_audit',
    'use_persona_tool', 'view_reports', 'send_generic_emails',
    'manage_invoices', 'run_audit_center', 'use_ai_pm'
  ],
  [ROLES.GUEST]: ['view_dashboard']
};

// =============================================
// CORE SYSTEM FUNCTIONS
// =============================================
function doGet(e) {
  Logger.log('doGet called with parameters: ' + JSON.stringify(e?.parameter));

  // --- Tracking Pixel Route ---
  if (e && e.parameter && e.parameter.page === 'track') {
    return handleTrackingPixel(e.parameter.id);
  }

  // --- Standard Page Serving Route ---
  // Reset Caches on new request
  if (typeof _sheetDataCache !== 'undefined') _sheetDataCache = {};
  if (typeof _spreadsheetCache !== 'undefined') _spreadsheetCache = {};

  try {
    const page = e?.parameter?.page || 'index';
    const monthParam = e?.parameter?.month || null;
    const perspectiveParam = e?.parameter?.perspective || 'All';

    if (page === 'report') {
      let reportDate;
      if (monthParam && /^\d{4}-\d{2}$/.test(monthParam)) {
          const [year, month] = monthParam.split('-').map(Number);
          reportDate = new Date(year, month - 1, 1);
      } else {
          const today = new Date();
          reportDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
      }
      const currentMonth = reportDate.getMonth();
      const currentYear = reportDate.getFullYear();
      const previousMonthDate = new Date(currentYear, currentMonth, 0);
      const previousMonth = previousMonthDate.getMonth();
      const previousMonthYear = previousMonthDate.getFullYear();

      const reportData = generateMonthlyReport(currentMonth, currentYear, perspectiveParam);
      const previousMonthReportData = generateMonthlyReport(previousMonth, previousMonthYear, perspectiveParam);
      reportData.previousMonth = previousMonthReportData;
      
      const aiInsights = getAIGeneratedInsights(reportData, previousMonthReportData, perspectiveParam);
      reportData.aiInsights = aiInsights;
      
      const template = HtmlService.createTemplateFromFile('Report');
      template.reportData = reportData;
      template.currentPerspective = perspectiveParam;

      return template.evaluate()
          .setTitle(`Monthly Migration Report (${perspectiveParam}) - ${reportData.month} ${reportData.year}`)
          .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    
    const template = HtmlService.createTemplateFromFile('Index');
    template.resetToken = e?.parameter?.resetToken || ''; 
    
    // getLiveCurrencyRates is now in InvoiceService.js
    const rates = getLiveCurrencyRates();
    template.currencyRates = rates; 
    template.currencyRatesJson = JSON.stringify(rates);
    
    return template.evaluate()
      .setTitle('JetLearn Operation System')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  } catch (error) {
    Logger.log('Error in doGet: ' + error.message);
    return HtmlService.createHtmlOutput('<h1>Error loading application: ' + error.message + '</h1>');
  }
}

// =============================================
// SYSTEM INITIALIZATION & DATA FETCHING
// =============================================

function initializeSystem() {
  Logger.log('Initializing Migration System');

  try {
    getOrCreateSheet(CONFIG.SHEETS.TEACHER_DATA);
    getOrCreateSheet(CONFIG.SHEETS.COURSE_NAME);
    const auditSheet = getOrCreateSheet(CONFIG.SHEETS.AUDIT_LOG);
    getOrCreateSheet(CONFIG.SHEETS.USER_PROFILES);
    getOrCreateSheet(CONFIG.SHEETS.TEACHER_COURSES);
    getOrCreateSheet(CONFIG.SHEETS.COURSE_PROGRESS_SUMMARY);
    getOrCreateSheet(CONFIG.SHEETS.USER_ACTIVITY_LOG);
    getOrCreateSheet(CONFIG.SHEETS.TASKS);
    const invoiceProductsSheet = getOrCreateSheet(CONFIG.SHEETS.INVOICE_PRODUCTS); 

    const userProfilesData = _getCachedSheetData(CONFIG.SHEETS.USER_PROFILES);
    if (userProfilesData.length <= 1) { 
      createDefaultUsers(); 
    }

    if (auditSheet.getLastRow() === 0 || auditSheet.getRange('A1').isBlank()) {
      auditSheet.appendRow([
        'Timestamp', 'Action', 'JLID', 'Learner', 'Old Teacher', 'New Teacher', 'Course', 'Status', 'Notes', 'Session ID', 'Reason for Migration', 'Intervened By'
      ]);
    }

    const courseProgressSheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.COURSE_PROGRESS_SUMMARY);
    if (courseProgressSheet.getLastRow() === 0 || courseProgressSheet.getRange('A1').isBlank()) {
      courseProgressSheet.appendRow(['Course Name', 'Not Onboarded', '1-10%', '11-20%', '21-30%', '31-40%', '41-50%', '51-60%', '61-70%', '71-80%', '81-90%', '91-99%', '100%']);
    }

    const userActivitySheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.USER_ACTIVITY_LOG);
    if (userActivitySheet.getLastRow() === 0 || userActivitySheet.getRange('A1').isBlank()) {
        userActivitySheet.appendRow(['Timestamp', 'Username', 'Action', 'Details', 'UserEmail']);
    }

    const tasksSheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.TASKS);
    if (tasksSheet.getLastRow() === 0 || tasksSheet.getRange('A1').isBlank()) {
        tasksSheet.appendRow(['Task ID', 'Created', 'Learner JLID', 'Learner Name', 'Task Description', 'Assigned To', 'Status', 'Due Date', 'Notes']);
    }

    if (invoiceProductsSheet.getLastRow() === 0 || invoiceProductsSheet.getRange('A1').isBlank()) {
        invoiceProductsSheet.appendRow(['Plan Name', 'Base Price EUR', 'Base Price GBP', 'Base Price USD', 'Base Price INR', 'Months Tenure', 'Default Sessions', 'Installment Count', 'Fixed Classes']); 
        invoiceProductsSheet.appendRow(['Credit Transfer', 0.00, 0.00, 0.00, 0, 1, 1, 1, 0]); 
        invoiceProductsSheet.appendRow(['GCSE Custom Revision', 0.00, 0.00, 0.00, 0, 1, 1, 1, 0]);
        invoiceProductsSheet.appendRow(['GCSE Programming + Algorithms', 0.00, 0.00, 0.00, 0, 1, 1, 1, 0]);
        invoiceProductsSheet.appendRow(['GCSE AC', 1788.00, 1500.00, 2000.00, 166667, 12, 1, 1, 48]); 
        invoiceProductsSheet.appendRow(['GCSE NC', 1788.00, 1500.00, 2000.00, 166667, 12, 1, 1, 48]);
        invoiceProductsSheet.appendRow(['PRM', 149.00, 125.00, 167.00, 13889, 1, 1, 1, 0]);

        invoiceProductsSheet.appendRow(['3 Years', 5364.00, 4500.00, 6000.00, 500000, 36, 1, 1, 144]); 
        invoiceProductsSheet.appendRow(['2 Years', 3576.00, 3000.00, 4000.00, 333333, 24, 1, 1, 96]); 
        invoiceProductsSheet.appendRow(['Annual', 1788.00, 1500.00, 2000.00, 166667, 12, 1, 1, 48]); 
        invoiceProductsSheet.appendRow(['Half-Yearly', 894.00, 750.00, 1000.00, 83333, 6, 1, 1, 24]); 
        invoiceProductsSheet.appendRow(['Quarterly', 447.00, 375.00, 500.00, 41667, 3, 1, 3, 12]); 
        invoiceProductsSheet.appendRow(['Monthly', 149.00, 125.00, 167.00, 13889, 1, 1, 1, 0]); 
    }

    Logger.log('System initialization completed successfully');
    return { success: true, message: 'System initialized successfully' };
  } catch (error) {
    Logger.log('System initialization failed: ' + error.message);
    return { success: false, message: 'Initialization failed: ' + error.message };
  }
}

function getCommunicationPageData() {
  Logger.log('getCommunicationPageData called');
  try {
    const teachers = getActiveTeachers();
    const courses = getCourseNames();
    const tpManagers = getTPManagers();
    const invoiceProducts = getInvoiceProductsData(); // Fetch invoice products

    const allTeacherData = getTeacherData(); 
    const clsManagers = new Set(allTeacherData.map(t => String(t.clsManagerResponsible || '').trim()).filter(name => name !== ''));
    
    // NEW: Comprehensive Timezone List
    const timezones = [
      '(GMT-12:00) International Date Line West', '(GMT-11:00) Coordinated Universal Time-11', '(GMT-10:00) Hawaii', '(GMT-09:00) Alaska', 
      '(GMT-08:00) Baja California', '(GMT-08:00) Pacific Time (US & Canada)', '(GMT-07:00) Arizona', '(GMT-07:00) Chihuahua, La Paz, Mazatlan', 
      '(GMT-07:00) Mountain Time (US & Canada)', '(GMT-06:00) Central America', '(GMT-06:00) Central Time (US & Canada)', 
      '(GMT-06:00) Guadalajara, Mexico City, Monterrey', '(GMT-06:00) Saskatchewan', '(GMT-05:00) Bogota, Lima, Quito', 
      '(GMT-05:00) Eastern Time (US & Canada)', '(GMT-05:00) Indiana (East)', '(GMT-04:30) Caracas', '(GMT-04:00) Asuncion', 
      '(GMT-04:00) Atlantic Time (Canada)', '(GMT-04:00) Cuiaba', '(GMT-04:00) Georgetown, La Paz, Manaus, San Juan', '(GMT-04:00) Santiago', 
      '(GMT-03:30) Newfoundland', '(GMT-03:00) Brasilia', '(GMT-03:00) Buenos Aires', '(GMT-03:00) Cayenne, Fortaleza', '(GMT-03:00) Greenland', 
      '(GMT-03:00) Montevideo', '(GMT-03:00) Salvador', '(GMT-02:00) Coordinated Universal Time-02', '(GMT-01:00) Azores', '(GMT-01:00) Cape Verde Is.', 
      '(GMT+00:00) Casablanca', '(GMT+00:00) Coordinated Universal Time', '(GMT+00:00) Dublin, Edinburgh, Lisbon, London', 
      '(GMT+00:00) Monrovia, Reykjavik', '(GMT+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna', 
      '(GMT+01:00) Belgrade, Bratislava, Budapest, Ljubljana, Prague', '(GMT+01:00) Brussels, Copenhagen, Madrid, Paris', 
      '(GMT+01:00) Sarajevo, Skopje, Warsaw, Zagreb', '(GMT+01:00) West Central Africa', '(GMT+01:00) Windhoek', '(GMT+02:00) Amman', 
      '(GMT+02:00) Athens, Bucharest', '(GMT+02:00) Beirut', '(GMT+02:00) Cairo', '(GMT+02:00) Damascus', '(GMT+02:00) E. Europe', 
      '(GMT+02:00) Harare, Pretoria', '(GMT+02:00) Helsinki, Kyiv, Riga, Sofia, Tallinn, Vilnius', '(GMT+02:00) Istanbul', 
      '(GMT+02:00) Jerusalem', '(GMT+02:00) Kaliningrad', '(GMT+02:00) Tripoli', '(GMT+03:00) Baghdad', '(GMT+03:00) Kuwait, Riyadh', '(GMT+03:00) Minsk', 
      '(GMT+03:00) Moscow, St. Petersburg, Volgograd', '(GMT+03:00) Nairobi', '(GMT+03:30) Tehran', '(GMT+04:00) Abu Dhabi, Muscat', '(GMT+04:00) Baku', 
      '(GMT+04:00) Izhevsk, Samara', '(GMT+04:00) Port Louis', '(GMT+04:00) Tbilisi', '(GMT+04:00) Yerevan', '(GMT+04:30) Kabul', 
      '(GMT+05:00) Ashgabat, Tashkent', '(GMT+05:00) Ekaterinburg', '(GMT+05:00) Islamabad, Karachi', '(GMT+05:30) Chennai, Kolkata, Mumbai, New Delhi', 
      '(GMT+05:30) Sri Jayawardenepura', '(GMT+05:45) Kathmandu', '(GMT+06:00) Astana', '(GMT+06:00) Dhaka', '(GMT+06:00) Novosibirsk', 
      '(GMT+06:30) Yangon (Rangoon)', '(GMT+07:00) Bangkok, Hanoi, Jakarta', '(GMT+07:00) Krasnoyarsk', 
      '(GMT+08:00) Beijing, Chongqing, Hong Kong, Urumqi', '(GMT+08:00) Irkutsk', '(GMT+08:00) Kuala Lumpur, Singapore', '(GMT+08:00) Perth', 
      '(GMT+08:00) Taipei', '(GMT+08:00) Ulaanbaatar', '(GMT+09:00) Osaka, Sapporo, Tokyo', '(GMT+09:00) Seoul', '(GMT+09:00) Yakutsk', 
      '(GMT+09:30) Adelaide', '(GMT+09:30) Darwin', '(GMT+10:00) Brisbane', '(GMT+10:00) Canberra, Melbourne, Sydney', 
      '(GMT+10:00) Guam, Port Moresby', '(GMT+10:00) Hobart', '(GMT+10:00) Vladivostok', '(GMT+11:00) Chokurdakh', '(GMT+11:00) Magadan', 
      '(GMT+11:00) Solomon Is., New Caledonia', '(GMT+12:00) Anadyr, Petropavlovsk-Kamchatsky', '(GMT+12:00) Auckland, Wellington', 
      '(GMT+12:00) Coordinated Universal Time+12', '(GMT+12:00) Fiji', '(GMT+13:00) Nuku\'alofa', '(GMT+13:00) Samoa'
    ];

const TIMEZONE_IANA_MAP = {
  // UK & Europe
  '(GMT+00:00) Dublin, Edinburgh, Lisbon, London': 'Europe/London',
  '(GMT+00:00) Monrovia, Reykjavik': 'Atlantic/Reykjavik',
  '(GMT+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna': 'Europe/Berlin',
  '(GMT+01:00) Brussels, Copenhagen, Madrid, Paris': 'Europe/Paris',
  
  // US & Americas
  '(GMT-05:00) Eastern Time (US & Canada)': 'America/New_York',
  '(GMT-05:00) Indiana (East)': 'America/Indiana/Indianapolis',
  '(GMT-06:00) Central Time (US & Canada)': 'America/Chicago',
  '(GMT-07:00) Mountain Time (US & Canada)': 'America/Denver',
  '(GMT-07:00) Arizona': 'America/Phoenix',
  '(GMT-08:00) Pacific Time (US & Canada)': 'America/Los_Angeles',
  '(GMT-08:00) Baja California': 'America/Tijuana',
  
  // Asia / Pacific
  '(GMT+05:30) Chennai, Kolkata, Mumbai, New Delhi': 'Asia/Kolkata',
  '(GMT+04:00) Abu Dhabi, Muscat': 'Asia/Dubai',
  '(GMT+08:00) Kuala Lumpur, Singapore': 'Asia/Singapore',
  '(GMT+10:00) Canberra, Melbourne, Sydney': 'Australia/Sydney',
  '(GMT+12:00) Auckland, Wellington': 'Pacific/Auckland'
};


    return {
      success: true,
      teachers: teachers,
      courses: courses,
      tpManagers: tpManagers,
      clsManagers: Array.from(clsManagers).sort(),
      jetGuides: ['Abhishek Nayak', 'Aishwarya Jain', 'Anamika Parmar', 'Molishka Rai', 'Spreha Jain', 'Satyam Mehra', 'Sunil Amarnath', ],
      invoiceProducts: invoiceProducts,
      timezones: timezones
    };
  } catch (error) {
    Logger.log('Error in getCommunicationPageData: ' + error.message);
    return { success: false, message: 'Failed to load form data: ' + error.message };
  }
}

function getSystemHealth() {
  Logger.log('getSystemHealth called');

  try {
    const health = {
      spreadsheetAccess: false,
      emailService: false,
      sheets: {
        teacherData: false,
        courseName: false,
        auditLog: false,
        userProfiles: false,
        teacherCourses: false,
        personaData: false,
        courseProgressSummary: false,
        userActivityLog: false,
        tasks: false, 
        invoiceProducts: false 
      },
      emailQuota: 'N/A', 
      lastCheck: new Date().toISOString()
    };

    try {
      const migrationSpreadsheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID);
      health.spreadsheetAccess = true;

      if (migrationSpreadsheet.getSheetByName(CONFIG.SHEETS.TEACHER_DATA)) health.sheets.teacherData = true;
      if (migrationSpreadsheet.getSheetByName(CONFIG.SHEETS.COURSE_NAME)) health.sheets.courseName = true;
      if (migrationSpreadsheet.getSheetByName(CONFIG.SHEETS.AUDIT_LOG)) health.sheets.auditLog = true;
      if (migrationSpreadsheet.getSheetByName(CONFIG.SHEETS.USER_PROFILES)) health.sheets.userProfiles = true;
      if (migrationSpreadsheet.getSheetByName(CONFIG.SHEETS.TEACHER_COURSES)) health.sheets.teacherCourses = true;
      if (migrationSpreadsheet.getSheetByName(CONFIG.SHEETS.COURSE_PROGRESS_SUMMARY)) health.sheets.courseProgressSummary = true;
      if (migrationSpreadsheet.getSheetByName(CONFIG.SHEETS.USER_ACTIVITY_LOG)) health.sheets.userActivityLog = true;
      if (migrationSpreadsheet.getSheetByName(CONFIG.SHEETS.TASKS)) health.sheets.tasks = true;
      if (migrationSpreadsheet.getSheetByName(CONFIG.SHEETS.INVOICE_PRODUCTS)) health.sheets.invoiceProducts = true; 

      try {
        const personaSpreadsheet = _getSpreadsheet(CONFIG.PERSONA_SHEET_ID);
        health.sheets.personaData = personaSpreadsheet.getSheetByName(CONFIG.SHEETS.PERSONA_DATA) !== null;
      } catch (e) {
        Logger.log('Persona spreadsheet/sheet access failed: ' + e.message);
      }

    } catch (error) {
      Logger.log('Migration Spreadsheet access failed: ' + error.message);
    }

    try {
      const quota = MailApp.getRemainingDailyQuota();
      health.emailService = quota > 0;
      health.emailQuota = quota;
    } catch (error) {
      Logger.log('Email service check failed: ' + error.message);
    }

    return health;
  } catch (error) {
    Logger.log('System health check failed: ' + error.message);
    return { error: error.message };
  }
}
const APP_VERSION = "340"; 

function getAppVersion() {
  return APP_VERSION;
}

function setupCheckRepliesTrigger() {
  // Delete existing triggers first to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkReplies') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Run every 2 hours
  ScriptApp.newTrigger('checkReplies')
    .timeBased()
    .everyHours(2)
    .create();

  Logger.log('checkReplies trigger set up successfully — runs every 2 hours.');
}

// ── AI Teacher Dashboard Insights ────────────────────────────────────────────
// Called from the frontend AI Dashboard tab — takes a prompt string and
// returns a JSON array string using the existing Gemini setup in AIService.js
function getAITeacherInsights(prompt) {
  Logger.log('[getAITeacherInsights] called, prompt length: ' + prompt.length);
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty(GEMINI_API_KEY_PROP);
    if (!apiKey) throw new Error('GEMINI_API_KEY not set in Script Properties.');

    const model = PropertiesService.getScriptProperties().getProperty('AI_SELECTED_MODEL') || 'gemini-2.5-flash';
    const url   = GEMINI_BASE_URL + model + ':generateContent?key=' + apiKey;

    const response = UrlFetchApp.fetch(url, {
      method:  'post',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: {
          temperature:     0.4,
          maxOutputTokens: 1024
        }
      }),
      muteHttpExceptions: true
    });

    const json = JSON.parse(response.getContentText());
    if (json.error) throw new Error('Gemini error: ' + json.error.message);

    const text = json.candidates[0].content.parts[0].text || '';

    // Strip markdown fences and return raw JSON string
    const clean = text.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();

    // Validate it's a JSON array before returning
    const parsed = JSON.parse(clean);
    if (!Array.isArray(parsed)) throw new Error('Gemini returned non-array response.');

    Logger.log('[getAITeacherInsights] returned ' + parsed.length + ' insights via ' + model);
    return clean;

  } catch (e) {
    Logger.log('[getAITeacherInsights] error: ' + e.message);
    // Return a single error insight so the frontend still renders gracefully
    return JSON.stringify([{ type: 'info', text: 'AI analysis unavailable: ' + e.message }]);
  }
}