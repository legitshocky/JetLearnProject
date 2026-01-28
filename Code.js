// Enhanced Migration System
// Version: V25 - Refactored Structure

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
  ADMIN: 'Admin',
  USER: 'User',
  GUEST: 'Guest'
};

const PERMISSIONS = {
  [ROLES.ADMIN]: ['view_dashboard', 'send_emails', 'view_audit', 'manage_users', 'view_reports', 'use_persona_tool', 'manage_settings', 'send_generic_emails', 'manage_invoices', 'run_audit_center', 'manage_agentic_audit'], 
  [ROLES.USER]: ['view_dashboard', 'send_emails', 'view_audit', 'use_persona_tool', 'view_reports', 'send_generic_emails', 'manage_invoices', 'run_audit_center'], 
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
    template.resetToken = e?.parameter?.resetToken || null;
    
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

    // User creation logic is now in UserService.js, called here if needed
    const userProfilesData = _getCachedSheetData(CONFIG.SHEETS.USER_PROFILES);
    if (userProfilesData.length <= 1) { 
      createDefaultUsers(); 
    }

    if (auditSheet.getLastRow() === 0 || auditSheet.getRange('A1').isBlank()) {
      auditSheet.appendRow([
        'Timestamp', 'Action', 'JLID', 'Learner', 'Old Teacher', 'New Teacher', 'Course', 'Status', 'Notes', 'Session ID', 'Reason for Migration', 'Intervened By'
      ]);
    }

    const courseProgressSheet = getOrCreateSheet(CONFIG.SHEETS.COURSE_PROGRESS_SUMMARY);
    if (courseProgressSheet.getLastRow() === 0 || courseProgressSheet.getRange('A1').isBlank()) {
      courseProgressSheet.appendRow(['Course Name', 'Not Onboarded', '1-10%', '11-20%', '21-30%', '31-40%', '41-50%', '51-60%', '61-70%', '71-80%', '81-90%', '91-99%', '100%']);
    }

    const userActivitySheet = getOrCreateSheet(CONFIG.SHEETS.USER_ACTIVITY_LOG);
    if (userActivitySheet.getLastRow() === 0 || userActivitySheet.getRange('A1').isBlank()) {
        userActivitySheet.appendRow(['Timestamp', 'Username', 'Action', 'Details', 'UserEmail']);
    }

    const tasksSheet = getOrCreateSheet(CONFIG.SHEETS.TASKS);
    if (tasksSheet.getLastRow() === 0 || tasksSheet.getRange('A1').isBlank()) {
        tasksSheet.appendRow(['Task ID', 'Created', 'Learner JLID', 'Learner Name', 'Task Description', 'Assigned To', 'Status', 'Due Date', 'Notes']);
    }

    if (invoiceProductsSheet.getLastRow() === 0 || invoiceProductsSheet.getRange('A1').isBlank()) {
        invoiceProductsSheet.appendRow(['Plan Name', 'Base Price EUR', 'Base Price GBP', 'Base Price USD', 'Base Price INR', 'Months Tenure', 'Default Sessions', 'Installment Count', 'Fixed Classes']); 
        // ... (Default products can be added here or manually)
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
    const teachers = getActiveTeachers(); // TeacherService.js
    const courses = getCourseNames();     // TeacherService.js
    const tpManagers = getTPManagers();   // TeacherService.js
    const invoiceProducts = getInvoiceProductsData(); // InvoiceService.js

    const allTeacherData = getTeacherData(); // TeacherService.js
    const clsManagers = new Set(allTeacherData.map(t => String(t.clsManagerResponsible || '').trim()).filter(name => name !== ''));
    
    // Timezones list (Shortened for brevity, can remain here or move to Utils)
    const timezones = [
      '(GMT-12:00) International Date Line West', '(GMT-11:00) Coordinated Universal Time-11', '(GMT-10:00) Hawaii', 
      '(GMT-09:00) Alaska', '(GMT-08:00) Pacific Time (US & Canada)', '(GMT-07:00) Mountain Time (US & Canada)', 
      '(GMT-06:00) Central Time (US & Canada)', '(GMT-05:00) Eastern Time (US & Canada)', '(GMT-04:00) Atlantic Time (Canada)', 
      '(GMT+00:00) Dublin, Edinburgh, Lisbon, London', '(GMT+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna', 
      '(GMT+02:00) Helsinki, Kyiv, Riga, Sofia, Tallinn, Vilnius', '(GMT+03:00) Moscow, St. Petersburg, Volgograd', 
      '(GMT+04:00) Abu Dhabi, Muscat', '(GMT+05:30) Chennai, Kolkata, Mumbai, New Delhi', '(GMT+08:00) Kuala Lumpur, Singapore', 
      '(GMT+09:00) Osaka, Sapporo, Tokyo', '(GMT+10:00) Canberra, Melbourne, Sydney', '(GMT+12:00) Auckland, Wellington'
    ];

    return {
      success: true,
      teachers: teachers,
      courses: courses,
      tpManagers: tpManagers,
      clsManagers: Array.from(clsManagers).sort(),
      jetGuides: ['Abhishek Nayak', 'Aishwarya Jain', 'Anamika Parmar', 'Manish Singh', 'Molishka Rai', 'Sana Rais', 'Satyam Mehra', 'Sunil Amarnath', 'Uday Kanika', 'Ishita Pahwa'],
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
