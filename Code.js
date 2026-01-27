// Enhanced Migration System with Dashboard, Statistics, and Teacher Persona Tool
// Author: Assistant
// Version: V25 - Wati Access!
// Major changes:

// =============================================
// CONFIGURATION
// =============================================
const CONFIG = {
  MIGRATION_SHEET_ID: '1xzprj2U6NpJwoevBMvM1DVfIj76wVjAd0ZcMjVC1xMM',
  PERSONA_SHEET_ID: '1rSweVyLKEwb1xThFHMLoH4xWnrLs8wbRM_61VtRjGww', 
  DRIVE_FOLDER_ID: '1K-Zb9BO2dm_dPg2AWTDT5t-ghkPoRNSW', 
  HUBSPOT_API_KEY: 'pat-na1-840cfb1a-acb3-45d6-8b0d-31f8c3f7cb34', 
  CLASS_SCHEDULE_CALENDAR_ID: 'hello@jet-learn.com', 

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
    EMAIL_LOGS: 'Email Logs' // <-- NEW

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


// --- Global Cache for current script execution ---
let _sheetDataCache = {};
let _spreadsheetCache = {};

function _getSpreadsheet(id) {
  if (!_spreadsheetCache[id]) {
    _spreadsheetCache[id] = SpreadsheetApp.openById(id);
  }
  return _spreadsheetCache[id];
}

/**
 * Fetches data from a specified sheet, leveraging an in-memory cache for the current execution.
 * @param {string} sheetName The name of the sheet to fetch.
 * @param {string} spreadsheetId The ID of the spreadsheet containing the sheet.
 * @returns {Array<Array<any>>} The sheet data.
 */
function _getCachedSheetData(sheetName, spreadsheetId = CONFIG.MIGRATION_SHEET_ID) {
  const cacheKey = `${spreadsheetId}_${sheetName}`;
  if (!_sheetDataCache[cacheKey]) {
    try {
      const spreadsheet = _getSpreadsheet(spreadsheetId);
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        Logger.log(`Warning: Sheet '${sheetName}' not found in spreadsheet ID: ${spreadsheetId}. Returning empty array.`);
        _sheetDataCache[cacheKey] = []; // Cache empty result to avoid repeated lookups
        return [];
      }
      _sheetDataCache[cacheKey] = sheet.getDataRange().getValues();
    } catch (e) {
      Logger.log(`Error fetching cached data for ${sheetName}: ${e.message}`);
      _sheetDataCache[cacheKey] = []; 
      return []; // Return empty array to prevent "No data returned" crash
    }
  }
  return _sheetDataCache[cacheKey];
}




// =============================================
// invoice function
// =============================================
function safeToFixed(value, decimals) {
  if (value === null || value === undefined || isNaN(value)) {
    return '0.00';
  }
  if (decimals === undefined) decimals = 2;
  return Number(value).toFixed(decimals);
}

/**
 * Fetches live currency conversion rates from an external API.
 * Falls back to a hardcoded list if the API call fails.
 * @returns {object} An object containing currency conversion rates against EUR.
 */
function getLiveCurrencyRates() {
  // This is our reliable fallback in case the API fails
  const fallbackRates = {
    'EUR': 1.0, 'USD': 1.1700, 'GBP': 0.8665, 'INR': 103.16, 'CHF': 0.9348,
    'AED': 4.273, 'CAD': 1.6224, 'AUD': 1.7682, 'JPY': 172.50, 'SGD': 1.50,
    'HKD': 9.1187, 'ZAR': 20.57, 'CNY': 8.3387, 'NZD': 1.9704, 'SEK': 10.951,
    'NOK': 11.6195, 'DKK': 7.45, 'MXN': 21.8069, 'BRL': 6.3207
  };

  try {
    const API_KEY = PropertiesService.getScriptProperties().getProperty('EXCHANGERATE_API_KEY');
    if (!API_KEY) {
      Logger.log('ExchangeRate API Key not found in Script Properties. Using fallback rates.');
      return fallbackRates;
    }

    const apiUrl = `https://v6.exchangerate-api.com/v6/${API_KEY}/latest/EUR`;
    const response = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data.result === 'success' && data.conversion_rates) {
        Logger.log('Successfully fetched live currency rates.');
        return data.conversion_rates;
      }
    }
    
    Logger.log('API call for currency rates failed. Response code: ' + response.getResponseCode() + '. Using fallback rates.');
    return fallbackRates;

  } catch (error) {
    Logger.log('Error fetching live currency rates: ' + error.message + '. Using fallback rates.');
    return fallbackRates;
  }
}

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
  _sheetDataCache = {};
  _spreadsheetCache = {};

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

/**
 * NEW: A robust wrapper for calling the Gemini API with a retry mechanism.
 * @param {string} endpoint The full API endpoint URL.
 * @param {object} payload The JSON payload for the request.
 * @returns {GoogleAppsScript.URL_Fetch.HTTPResponse} The successful HTTP response.
 * @throws {Error} If the API call fails after all retries.
 */
function callGenerativeAIWithRetry(endpoint, payload) {
  const MAX_RETRIES = 3;
  let lastError = null;

  for (let i = 0; i < MAX_RETRIES; i++) {
    try {
      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(endpoint, options);
      const responseCode = response.getResponseCode();

      if (responseCode === 200) {
        return response; // Success
      } else if (responseCode === 429) {
        Logger.log(`[AI API] Rate limit hit (429). Retry ${i + 1}/${MAX_RETRIES}...`);
        lastError = new Error(`API rate limit exceeded. Response: ${response.getContentText()}`);
        Utilities.sleep((i + 1) * 1000); // Wait 1s, then 2s, etc.
      } else {
        throw new Error(`AI API Error (${responseCode}): ${response.getContentText()}`);
      }
    } catch (e) {
      lastError = e;
      Logger.log(`[AI API] Connection error on retry ${i + 1}/${MAX_RETRIES}: ${e.message}`);
      Utilities.sleep((i + 1) * 1000); // Also wait on connection errors
    }
  }
  
  Logger.log(`[AI API] Final failure after ${MAX_RETRIES} retries.`);
  throw lastError; // Throw the last captured error after all retries fail
}


/**
 * Uses Google Gemini to analyze and generate insights by comparing two monthly reports.
 * @param {object} currentData - The report object for the current month.
 * @param {object} previousData - The report object for the previous month.
 * @returns {string} A JSON string representing an array of insight strings.
 */
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
// =============================================
// ONBOARDING EMAIL FUNCTIONS
// =============================================
// function sendOnboardingEmail(data, attachments = []) {
//   try {
//     const teacherData = getTeacherData();
    
//     // FIXED: Case-insensitive and trimmed comparison
//     const teacherInfo = teacherData.find(t => 
//       String(t.name).trim().toLowerCase() === String(data.teacherName).trim().toLowerCase()
//     );

//     if (!teacherInfo || !isValidEmail(teacherInfo.email)) {
//       throw new Error(`Teacher '${data.teacherName}' not found or has invalid email.`);
//     }

//     const clsManagerEmail = findClsEmailByManagerName(data.clsManager);
//     const ccList = new Set();
//     if(clsManagerEmail) ccList.add(clsManagerEmail);
//     if(teacherInfo.tpManagerEmail) ccList.add(teacherInfo.tpManagerEmail);

//     const subject = `New Learner Onboarded || ${data.learnerName} (${data.jlid || 'N/A'})`;
//     const htmlBody = getOnboardingEmailHTML(data);
    
//     // Use the central tracked email service
//     const result = sendTrackedEmail({
//       to: teacherInfo.email,
//       cc: Array.from(ccList).join(','),
//       subject: subject,
//       htmlBody: htmlBody,
//       jlid: data.jlid,
//       attachments: attachments
//     });
    
//     // NOTE: We no longer call logAction here. The parent function does it.
    
//     return { success: true, message: `Onboarding email sent to teacher. TID: ${result.trackingId}` };

//   } catch (error) {
//     Logger.log(`Error in helper sendOnboardingEmail: ${error.message}`);
//     // Re-throw the error so the main function knows it failed
//     throw new Error(`Failed to send teacher onboarding email: ${error.message}`);
//   }
// }


// =============================================
// MIGRATION MANAGEMENT
// =============================================
// function sendMigrationEmail(data, attachments = []) {
//   let finalStatus = 'Partial Success';
//   let notes = [];

//   try {
//     // ==========================================
//     // 1. TEACHER EMAIL (Always CET)
//     // ==========================================
//     if (data.sendEmailToTeacher) {
//         if (!data.jlid || !data.newTeacher || !data.course) throw new Error('Missing fields for Email.');
        
//         const teacherData = getTeacherData();
//         // Case-insensitive lookup for new teacher
//         const newTeacherInfo = teacherData.find(t => String(t.name).trim().toLowerCase() === String(data.newTeacher).trim().toLowerCase());
        
//         // Lookup old teacher if provided
//         let oldTeacherInfo = null;
//         if (data.oldTeacher) {
//             oldTeacherInfo = teacherData.find(t => String(t.name).trim().toLowerCase() === String(data.oldTeacher).trim().toLowerCase());
//         }

//         if (!newTeacherInfo || !isValidEmail(newTeacherInfo.email)) {
//             throw new Error(`New teacher email invalid or not found: ${data.newTeacher}`);
//         }
        
//         const finalClsEmailForCC = findClsEmailByManagerName(data.clsManager);
        
//         // Send to New Teacher
//         const newTeacherResult = sendTrackedEmail({
//           to: newTeacherInfo.email, 
//           cc: finalClsEmailForCC, 
//           subject: `Migration Notice - ${data.learner} Assigned`,
//           htmlBody: getNewTeacherEmailHTML(data, data.newTeacherComments || ''), 
//           jlid: data.jlid, 
//           attachments: attachments
//         });
//         notes.push(`Teacher Email Sent (TID: ${newTeacherResult.trackingId})`);
        
//         // Optional: Send to Old Teacher
//         if (oldTeacherInfo && isValidEmail(oldTeacherInfo.email)) {
//            sendTrackedEmail({
//             to: oldTeacherInfo.email, 
//             cc: finalClsEmailForCC, 
//             subject: `${data.learner} - Migration`,
//             htmlBody: getOldTeacherEmailHTML(data, ''), 
//             jlid: data.jlid
//           });
//           notes.push("Old Teacher Email Sent.");
//         }
//     } else {
//         notes.push("Teacher Email Skipped.");
//     }

//     // ==========================================
//     // 2. WATI WHATSAPP (Parent's Timezone)
//     // ==========================================
//     if (data.sendWhatsappToParent) {
        
//         // A. Fetch Hubspot Data to get Parent Phone & Name
//         const hubspotResult = fetchHubspotByJlid(data.jlid);
//         if (!hubspotResult.success) {
//              throw new Error(`WATI Failed: Could not fetch data for JLID ${data.jlid}.`);
//         }
        
//         const hsData = hubspotResult.data;
//         const parentPhone = hsData.parentContact;
        
//         if (!parentPhone) throw new Error("WATI Failed: Parent phone number missing in HubSpot.");

//         // B. Calculate Local Time for the message
//         const firstSession = data.classSessions && data.classSessions.length > 0 ? data.classSessions[0] : { day: "TBD", time: "TBD" };
//         const cetTime = firstSession.time; 
//         const parentTimezone = data.manualTimezone || hsData.timezone;
//         // Inject into data object for helper to find
//         data.calculatedLocalTime = convertCetToLocal(cetTime, parentTimezone); 

//         // C. DETERMINE TEMPLATE ID
//         // Priority 1: Use the specific template ID sent from the dropdown (data.watiTemplateName)
//         // Priority 2: Look up default for the reason
//         // Priority 3: Generic fallback
//         let templateId = data.watiTemplateName;
        
//         if (!templateId || templateId === "") {
//             const mapping = WATI_REASON_MAPPING[data.reasonOfMigration];
//             if (mapping && mapping.length > 0) {
//                 templateId = mapping[0].id; // Use the first mapped template
//             } else {
//                 templateId = "migration_generic_update"; // Final fallback
//             }
//         }

//         // D. Build Parameters
//         const watiParameters = getWatiParameters(templateId, data, hsData);

//         // Debug Log
//         Logger.log(`Sending WATI Template: ${templateId}`);
//         Logger.log(`Params: ${JSON.stringify(watiParameters)}`);

//         // E. Send Message
//         const watiResult = sendWatiMessage(parentPhone, templateId, watiParameters);
        
//         if(watiResult.success) {
//             notes.push(`WATI Sent (${templateId}) to ${parentPhone}`);
//         } else {
//             notes.push(`WATI Failed: ${JSON.stringify(watiResult)}`);
//         }
//     } else {
//         notes.push("WhatsApp Skipped.");
//     }

//     finalStatus = 'Success';
//     return { success: true, message: 'Migration process completed successfully.' };

//   } catch (error) {
//     Logger.log(`Error in Migration Process: ${error.message}`);
//     notes.push(`Error: ${error.message}`);
//     return { success: false, message: `Failed: ${error.message}` };
//   } finally {
//     // Log the outcome
//     logAction('Migration Process', data.jlid, data.learner, data.oldTeacher, data.newTeacher, data.course, finalStatus, notes.join('; '), data.reasonOfMigration);
//   }
// }


// =============================================
// NEW: GENERIC EMAIL SENDER & TEMPLATES
// =============================================

function sendGenericEmail(emailType, formData, recipientEmail, attachmentsBase64, comments) {
  try {
    if (!isValidEmail(recipientEmail)) throw new Error('Valid Recipient Email is required.');

    let subject = '';
    let htmlBody = '';

    switch (emailType) {
      case 'Minecraft Install':
        subject = `JetLearn: Minecraft Installation Guide for ${formData.learnerName || 'Learner'}`;
        htmlBody = getMinecraftInstallEmailHTML(formData, comments);
        break;
      case 'Roblox Install':
        subject = `JetLearn: Roblox Installation Guide for ${formData.learnerName || 'Learner'}`;
        htmlBody = getRobloxInstallEmailHTML(formData, comments);
        break;
      default: throw new Error(`Unknown email type: ${emailType}.`);
    }

    const attachments = (attachmentsBase64 && attachmentsBase64.length > 0) ? uploadAttachments(attachmentsBase64) : [];
    
    const result = sendTrackedEmail({ to: recipientEmail, subject, htmlBody, jlid: formData.jlid, attachments });
    
    logAction(`Email Sent (${emailType})`, formData.jlid, formData.learnerName, '', '', formData.course, 'Success', `TID: ${result.trackingId}`);
    return { success: true, message: `${emailType} email sent successfully!` };
  } catch (error) {
    logAction(`Email Failed (${emailType})`, formData.jlid, formData.learnerName, '', '', formData.course, 'Failed', error.message);
    return { success: false, message: `Failed to send ${emailType} email: ${error.message}` };
  }
}


/**
 * NEW: Generates a combined HTML document showing both the email and invoice preview.
 * This is called specifically when previewing an onboarding email with an invoice.
 * @param {object} formData - The complete form data.
 * @returns {object} A result object with the combined HTML.
 */
function renderCombinedPreview(formData) {
  try {
    // First, validate the data to ensure both parts can be generated.
    const validationErrors = validateInvoiceData(formData);
    if (validationErrors.length > 0) {
      throw new Error('Invoice data is invalid: ' + validationErrors.join('; '));
    }

    // Generate both HTML components.
    const pricingDetails = calculateInvoicePricing(formData, true); // true for previewOnly
    
    // FIX: Inject sessions into formData
    formData.sessions = pricingDetails.displayTotalSessions;

    const invoiceHtml = getInvoiceHTML(formData, pricingDetails);
    const emailHtml = getParentOnboardingEmailHTML(formData);

    // Create a simple wrapper HTML to display both parts, one after the other.
    const combinedHtml = `
      <html>
        <head>
          <style>
            body { 
              font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif; 
              background-color: #f0f2f5; 
              margin: 0; 
              padding: 20px; 
              /* START OF CORRECTION: This line fixes the text color issue */
              color: #1a1a1a !important; 
              /* END OF CORRECTION */
            }
            .preview-section { background: white; border: 1px solid #dee2e6; margin: 0 auto 20px auto; box-shadow: 0 4px 6px rgba(0,0,0,0.05); max-width: 850px; }
            .preview-header { background-color: #4a3c8a; color: white; padding: 12px 20px; font-size: 1.1em; font-weight: 600; border-bottom: 1px solid #dee2e6; }
            .preview-content { transform: scale(0.95); transform-origin: top left; padding-left: 2.5%; padding-right: 2.5%; }
          </style>
        </head>
        <body>
          <div class="preview-section">
            <div class="preview-header">Email Preview</div>
            <div class="preview-content">
              ${emailHtml}
            </div>
          </div>
          <div class="preview-section">
            <div class="preview-header">Attached Invoice Preview</div>
            <div class="preview-content">
              ${invoiceHtml}
            </div>
          </div>
        </body>
      </html>
    `;

    return { success: true, html: combinedHtml };

  } catch (error) {
    Logger.log('Error in renderCombinedPreview: ' + error.message);
    return { success: false, message: error.message };
  }
}

// =============================================
// RENEWAL COMMUNICATION LOGIC
// =============================================

/**
 * Fetches Deal Data AND Line Items for Renewal Page
 */
function fetchHubspotRenewalData(jlid) {
  const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  
  // 1. Search for potential deals
  const searchUrl = 'https://api.hubapi.com/crm/v3/objects/deals/search';
  const searchPayload = {
    filterGroups: [{ filters: [{ propertyName: 'jetlearner_id', operator: 'EQ', value: jlid }] }],
    sorts: [{ propertyName: 'createdate', direction: 'DESCENDING' }],
    properties: ['dealname', 'amount', 'deal_currency_code', 'parent_name', 'parent_email', 'jetlearner_id'],
    limit: 5
  };
  
  try {
    const searchRes = UrlFetchApp.fetch(searchUrl, {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(searchPayload),
      muteHttpExceptions: true
    });
    
    const searchJson = JSON.parse(searchRes.getContentText());
    if (!searchJson.results || searchJson.results.length === 0) {
        return { success: false, message: "No Deal found for this JLID." };
    }

    // 2. Find the deal that actually has line items
    let targetDeal = null;
    let lineItemIds = [];

    for (const deal of searchJson.results) {
        const assocUrl = `https://api.hubapi.com/crm/v4/objects/deals/${deal.id}/associations/line_items`;
        const assocRes = UrlFetchApp.fetch(assocUrl, {
            method: 'get',
            headers: { 'Authorization': 'Bearer ' + token },
            muteHttpExceptions: true
        });
        
        const assocJson = JSON.parse(assocRes.getContentText());
        if (assocJson.results && assocJson.results.length > 0) {
            targetDeal = deal;
            lineItemIds = assocJson.results.map(r => ({ id: r.toObjectId }));
            break; 
        }
    }

    if (!targetDeal) {
        return { success: false, message: "Deal found, but 0 line items attached." };
    }

    // 3. Fetch details for those line items
    const batchUrl = `https://api.hubapi.com/crm/v3/objects/line_items/batch/read`;
    const batchPayload = {
        properties: [
            'name', 'price', 'discount', 'hs_total_discount', 'net_price', 'currency', 'quantity', 'hs_createdate',
            // VALIDATED PROPERTIES FROM DEBUG LOGS
            'payment_received_date___cloned_', 
            'renewal__payment_type__cloned_', 
            'renewal__payment_term__cloned_',
            'full_payment_received__y_n___cloned_',
            'hs_recurring_billing_number_of_payments'
        ],
        inputs: lineItemIds
    };
    
    const batchRes = UrlFetchApp.fetch(batchUrl, {
        method: 'post',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify(batchPayload),
        muteHttpExceptions: true
    });
    
    const batchJson = JSON.parse(batchRes.getContentText());
    let processedItems = [];

    if (batchJson.results) {
        // Sort by ID Descending (Newest First) since createdate was null in logs
        const sortedResults = batchJson.results.sort((a, b) => parseInt(b.id) - parseInt(a.id));

        processedItems = sortedResults.map(item => {
            const p = item.properties;
            const qty = parseFloat(p.quantity) || 1;
            const price = parseFloat(p.price) || 0;
            const discount = parseFloat(p.hs_total_discount) || parseFloat(p.discount) || 0;
            const net = p.net_price ? parseFloat(p.net_price) : ((price * qty) - discount);

            // Date Parsing (Handle '2025-12-21')
            let finalDate = '';
            let rawDate = p.payment_received_date___cloned_;
            if (rawDate) {
                // If it's already YYYY-MM-DD, use it. If timestamp, convert.
                if (rawDate.includes('-')) finalDate = rawDate;
                else if (!isNaN(rawDate)) finalDate = new Date(parseInt(rawDate)).toISOString().split('T')[0];
            }

            return {
                id: item.id,
                name: p.name || 'Unknown Item',
                createdDate: p.hs_createdate || "N/A", // Display only
                unitPrice: (price * qty).toFixed(2),
                discount: discount.toFixed(2),
                netPrice: net.toFixed(2),
                currency: p.currency,
                
                // Mapped Custom Fields
                paymentDate: finalDate, 
                paymentType: p.renewal__payment_type__cloned_ || 'Upfront',
                frequency: p.renewal__payment_term__cloned_ || 'Monthly',
                installments: p.hs_recurring_billing_number_of_payments || '1',
                isFullPayment: p.full_payment_received__y_n___cloned_
            };
        });
    }

    return { 
        success: true, 
        data: {
            deal: {
                dealId: targetDeal.id,
                learnerName: targetDeal.properties.dealname || '',
                parentName: targetDeal.properties.parent_name || '',
                parentEmail: targetDeal.properties.parent_email || '',
                currency: targetDeal.properties.deal_currency_code || 'EUR'
            },
            lineItems: processedItems
        }
    };

  } catch (e) {
    Logger.log("Error: " + e.message);
    return { success: false, message: "Server Error: " + e.message };
  }
}


/**
 * Generates Renewal PDF and Sends Email
 */
function sendRenewalCommunication(formData) {
  try {
    // 1. Generate PDF
    const template = HtmlService.createTemplateFromFile('RenewalInvoiceTemplate');
    template.data = formData;
    const html = template.evaluate().getContent();
    const pdfBlob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF)
                             .setName(`Renewal_Invoice_${formData.learnerName}.pdf`);

    // 2. Prepare Email HTML
    const emailTemplate = HtmlService.createTemplateFromFile('RenewalEmailTemplate');
    emailTemplate.data = formData;
    const emailBody = emailTemplate.evaluate().getContent();

    // 3. Send via central tracker
    sendTrackedEmail({
        to: formData.parentEmail,
        subject: `Renewal Confirmation - ${formData.learnerName}`,
        htmlBody: emailBody,
        jlid: formData.jlid,
        attachments: [pdfBlob]
    });

    logAction('Renewal Sent', formData.jlid, formData.learnerName, '', '', formData.planName, 'Success', `Renewal Amount: ${formData.netPrice}`);

    return { success: true };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Helpers for Preview Generation
function getRenewalEmailHTML(formData) {
  const template = HtmlService.createTemplateFromFile('RenewalEmailTemplate');
  template.data = formData;
  return template.evaluate().getContent();
}

function getRenewalInvoiceHTML(formData) {
  const template = HtmlService.createTemplateFromFile('RenewalInvoiceTemplate');
  template.data = formData;
  return template.evaluate().getContent();
}

/**
 * Renders the HTML content for an email preview based on the email type and form data.
 * This function does NOT send an email.
 * @param {string} emailType The type of email template to render (e.g., 'Onboarding', 'Migration', 'Onboarding (Parent)', 'Invoice').
 * @param {object} formData The form data containing all necessary details for the email.
 * @param {string} comments Additional comments to include in the email.
 * @param {Array<string>} attachmentNames List of names of files that would be attached (for display in preview).
 * @returns {object} An object containing { success: boolean, html: string, message: string }.
 */
function renderEmailPreview(emailType, formData, comments, attachmentNames) { 
  Logger.log(`renderEmailPreview called for type: ${emailType}`);
  
  try {
    // --- NEW: RENEWAL CASE ---
    if (emailType === 'Renewal') {
        const emailHtml = getRenewalEmailHTML(formData);
        const invoiceHtml = getRenewalInvoiceHTML(formData);
        
        // Combined HTML for side-by-side preview
        const combinedHtml = `
          <html>
            <head>
              <style>
                body { background-color: #eef1f5; padding: 20px; font-family: sans-serif; }
                .preview-label { background: #333; color: #fff; padding: 8px 15px; border-radius: 4px 4px 0 0; font-weight: bold; font-size: 14px; margin-top: 20px; width: fit-content; }
                .preview-card { background: white; border: 1px solid #ddd; box-shadow: 0 2px 10px rgba(0,0,0,0.1); margin-bottom: 40px; border-radius: 0 4px 4px 4px; overflow: hidden; }
              </style>
            </head>
            <body>
              <div class="preview-label"><i class="fas fa-envelope"></i> Email Body</div>
              <div class="preview-card">${emailHtml}</div>

              <div class="preview-label"><i class="fas fa-file-pdf"></i> Attached PDF Invoice</div>
              <div class="preview-card" style="padding: 20px;">
                <div style="transform: scale(0.95); transform-origin: top left; border: 1px solid #eee;">
                   ${invoiceHtml}
                </div>
              </div>
            </body>
          </html>`;
        return { success: true, html: combinedHtml, message: 'Renewal preview generated.' };
    }

    // --- EXISTING CASES (Kept exactly as they were) ---
    if (emailType === 'Invoice') {
        const validationErrors = validateInvoiceData(formData); 
        if (validationErrors.length > 0) throw new Error('Validation failed: ' + validationErrors.join(', '));
        const pricingDetailsPreview = calculateInvoicePricing(formData, true); 
        const invoiceHtml = getInvoiceHTML(formData, pricingDetailsPreview);
        return { success: true, html: invoiceHtml, message: 'Invoice preview generated.' };
    }

    let htmlBody = '';
    let previewMessage = 'Preview generated successfully.';

    if (emailType === 'Onboarding (Parent)' && formData.generateAndAttachInvoice) {
      return renderCombinedPreview(formData);
    }

    switch (emailType) {
        case 'Onboarding (Parent)':
            if (!formData.sessions) {
                try { const pricing = calculateInvoicePricing(formData, true); formData.sessions = pricing.displayTotalSessions; } catch (e) { formData.sessions = 0; }
            }
            const parentTemplate = HtmlService.createTemplateFromFile('ParentOnboardingTemplate');
            parentTemplate.data = formData; 
            htmlBody = parentTemplate.evaluate().getContent();
            break;
        case 'Minecraft Install':
            const mcTemplate = HtmlService.createTemplateFromFile('MinecraftInstallTemplate');
            mcTemplate.data = formData; mcTemplate.comments = comments;
            htmlBody = mcTemplate.evaluate().getContent();
            break;
        case 'Roblox Install':
            const rbTemplate = HtmlService.createTemplateFromFile('RobloxInstallTemplate');
            rbTemplate.data = formData; rbTemplate.comments = comments;
            htmlBody = rbTemplate.evaluate().getContent();
            break;
        default:
            throw new Error(`Unknown email type selected for preview: ${emailType}`);
    }

    let attachmentPreview = '';
    if (attachmentNames && attachmentNames.length > 0) {
        attachmentPreview = `<p style="margin-top: 20px; font-style: italic; color: #555;"><strong>Attachments:</strong> ${attachmentNames.join(', ')}</p>`;
    }
    let commentsPreview = '';
    if (comments) {
        commentsPreview = `<p style="margin-top: 10px; font-style: italic; color: #555;"><strong>Additional Comments:</strong> ${comments}</p>`;
    }

    const finalHtml = htmlBody.replace('</body>', `${commentsPreview}${attachmentPreview}</body>`);
    return { success: true, html: finalHtml, message: previewMessage };

  } catch (error) {
    Logger.log('Error in renderEmailPreview: ' + error.message);
    return { success: false, message: 'Failed to generate preview: ' + error.message };
  }
}


// Get HTML content from Apps Script HTML files
function getOnboardingEmailHTML(data) { 
  const template = HtmlService.createTemplateFromFile('OnboardingTemplate');
  template.data = data;
  return template.evaluate().getContent();
}

function getNewTeacherEmailHTML(data, comments) {
  const template = HtmlService.createTemplateFromFile('NewTeacherTemplate');
  template.data = data;
  template.comments = comments;
  return template.evaluate().getContent();
}

function getOldTeacherEmailHTML(data, comments) {
  const template = HtmlService.createTemplateFromFile('OldTeacherTemplate');
  template.data = data;
  template.comments = comments;
  return template.evaluate().getContent();
}

function getParentOnboardingEmailHTML(data) { 
  const template = HtmlService.createTemplateFromFile('ParentOnboardingTemplate');
  template.data = data; 
  return template.evaluate().getContent();
}

function getMinecraftInstallEmailHTML(data, comments) {
  const template = HtmlService.createTemplateFromFile('MinecraftInstallTemplate');
  template.data = data;
  template.comments = comments;
  return template.evaluate().getContent();
}

function getRobloxInstallEmailHTML(data, comments) {
  const template = HtmlService.createTemplateFromFile('RobloxInstallTemplate');
  template.data = data;
  template.comments = comments;
  return template.evaluate().getContent();
}

function getInvoiceHTML(formData, pricingDetails) {
  const jlid = (formData.jlid || '').trim().toUpperCase();
  let planDescription = 'Comprehensive Learning Program'; 

  if (jlid.endsWith('C')) {
    planDescription = 'Comprehensive AI Coding Program';
  } else if (jlid.endsWith('M')) {
    planDescription = 'Comprehensive Math Program';
  }

  formData.planDescription = planDescription;

  const template = HtmlService.createTemplateFromFile('InvoiceTemplate');
  template.data = formData;
  template.pricing = pricingDetails; 
  return template.evaluate().getContent();
}

/**
 * Uploads base64 encoded files to Google Drive and returns them as BlobSource objects.
 * @param {Array<Object>} base64Files - Array of objects with { fileName: string, mimeType: string, base64Data: string }
 * @returns {Array<GoogleAppsScript.Base.BlobSource>} Array of BlobSource objects.
 */
function uploadAttachments(base64Files) {
  if (!CONFIG.DRIVE_FOLDER_ID || CONFIG.DRIVE_FOLDER_ID === '1_exampleFolderID1234567890abcdef') {
    Logger.log("DRIVE_FOLDER_ID is not configured. Attachments will not be saved to Drive.");
    throw new Error("Drive folder ID is not configured. Please set CONFIG.DRIVE_FOLDER_ID in code.js.");
  }

  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const attachments = [];

  base64Files.forEach(fileData => {
    try {
      const decodedData = Utilities.base64Decode(fileData.base64Data);
      const blob = Utilities.newBlob(decodedData, fileData.mimeType, fileData.fileName);
      folder.createFile(blob); 
      attachments.push(blob); 
      Logger.log(`Uploaded attachment: ${fileData.fileName}`);
    } catch (e) {
      Logger.log(`Failed to upload or create blob for ${fileData.fileName}: ${e.message}`);
      throw new Error(`Failed to process attachment ${fileData.fileName}: ${e.message}`); 
    }
  });
  return attachments;
}

// =============================================
// NEW: HUBSPOT INTEGRATION
// =============================================

/**
 * Fetches learner and parent details from HubSpot using JLID.
 * Assumes 'jetlearner_id' is a custom property in HubSpot deals.
 * @param {string} jlid - The JetLearn ID.
 * @returns {object} Clean JSON object with learner/parent details or success: false.
 */
function fetchHubspotByJlid(jlid) {
  Logger.log('fetchHubspotByJlid called for JLID: ' + jlid);
  if (!jlid) {
    return { success: false, message: 'JLID is required for HubSpot lookup.' };
  }

  const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  if (!token) {
    Logger.log('HubSpot API token not configured.');
    return { success: false, message: 'HubSpot API token not configured. Please set it in Script Properties.' };
  }

  const hubspotApiUrl = 'https://api.hubapi.com/crm/v3/objects/deals/search';

  const properties = [
    'dealname', 'jetlearner_id', 'amount', 'deal_currency_code', 'hs_object_id', 'age', 'learner_status',
    'module_start_date', 'module_end_date', 'total_classes_committed_through_learner_s_journey',
    'current_teacher', 'current_course', 'time_zone', 'regular_class_day', 'frequency_of_classes',
    'payment_type', 'subscription', 'subscription_tenure', 'payment_term', 'class_timings',
    'learner_practice_document_link', 'installment_type', 'installment_terms_final', 
    'installment_months', 'installment_received_months__cloned_', 'payment_due_date',
    'full_payment_received__y_n_', 'jet_guide', 'cls_manager', 'teacher_manager',
    'parent_email', 'parent_name', 'phone_number_deal_',
    'stage____payment_trigger_date', 'zoom_masked_link'
  ];

  const requestBody = {
    filterGroups: [{ filters: [{ propertyName: 'jetlearner_id', operator: 'EQ', value: jlid }] }],
    properties: properties,
    limit: 1
  };

  try {
    const options = {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(hubspotApiUrl, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    const jsonResponse = JSON.parse(responseBody);

    if (responseCode !== 200) {
      Logger.log(`HubSpot API Error (${responseCode}): ${responseBody}`);
      const errorMessage = jsonResponse.message || (jsonResponse.errors && jsonResponse.errors[0] && jsonResponse.errors[0].message) || 'Unknown error';
      return { success: false, message: `HubSpot API Error: ${errorMessage}` };
    }

    if (jsonResponse.results && jsonResponse.results.length > 0) {
      const contactProperties = jsonResponse.results[0].properties; 
      if (!contactProperties) {
          Logger.log('HubSpot API: Deal found but properties object is null for JLID: ' + jlid);
          return { success: false, message: 'HubSpot data found but properties are empty.' };
      }
      
      // ================== FIX START: AUTOMATIC DISCOUNT CALCULATION ==================
      const tenure = parseInt(contactProperties.subscription_tenure || '0');
      const dealAmount = parseFloat(contactProperties.amount || '0');
      const currencyCode = contactProperties.deal_currency_code || 'EUR'; // Get the deal currency (e.g., CHF)
      let calculatedDiscount = 0;

      if (tenure > 0 && dealAmount > 0) {
          // 1. Calculate Standard Price in EUR (Base is €149/mo)
          const standardPriceEur = tenure * 149; 
          
          // 2. Get conversion rate (e.g., 1 EUR = ~0.93 CHF)
          // Uses your existing helper function to get the rate for the specific currency
          const conversionRate = getConversionRate(currencyCode); 
          
          // 3. Convert Standard Price to the Deal's Currency
          const standardPriceLocal = standardPriceEur * conversionRate;

          // 4. Calculate Discount (Standard Local Price - Actual Deal Amount)
          // Example: (1666.06) - 829 = 837.06
          if (standardPriceLocal > dealAmount) {
              calculatedDiscount = standardPriceLocal - dealAmount;
          }
      }

      // =================== FIX END: AUTOMATIC DISCOUNT CALCULATION ===================


      const parseClassTimings = (timingsString) => {
          if (!timingsString) return [];
          if (timingsString.includes(' at ')) {
            return timingsString.split(';').map(sessionStr => {
                const parts = sessionStr.trim().match(/(\w+)\s+at\s+(\d{1,2}:\d{2}\s(?:AM|PM))/i);
                if (parts && parts.length === 3) {
                    return { day: parts[1], time: parts[2] };
                }
                return null;
            }).filter(Boolean);
          }
          return timingsString.split(/[,;]/).map(dayStr => {
              if (dayStr.trim()) {
                  return { day: dayStr.trim(), time: '' };
              }
              return null;
          }).filter(Boolean);
      };

      const parsePaymentPlan = (hubspotPlan) => {
        if (!hubspotPlan) return { paymentPlanType: 'Upfront', installmentFrequency: '', customPlanDetails: '' };
        hubspotPlan = hubspotPlan.toLowerCase();
        if (hubspotPlan.includes('upfront') || hubspotPlan.includes('fully paid')) {
            return { paymentPlanType: 'Upfront', installmentFrequency: '', customPlanDetails: '' };
        } else if (hubspotPlan.includes('installment')) {
            let frequency = 'Monthly';
            if (hubspotPlan.includes('bi-monthly') || hubspotPlan.includes('alternate')) frequency = 'Alternate';
            else if (hubspotPlan.includes('quarterly')) frequency = 'Quarterly';
            return { paymentPlanType: 'Installment', installmentFrequency: frequency, customPlanDetails: '' };
        } else {
            return { paymentPlanType: 'Custom', installmentFrequency: '', customPlanDetails: hubspotPlan };
        }
      };
      
      let sessionsPerWeekString = '';
      const rawFrequency = contactProperties.frequency_of_classes;
      if (rawFrequency && typeof rawFrequency === 'string') {
          const numMatch = rawFrequency.match(/\d+/);
          if (numMatch) {
              const num = parseInt(numMatch[0], 10);
              if (rawFrequency.toLowerCase().includes('week')) {
                  sessionsPerWeekString = `${num} Session${num === 1 ? '' : 's'}/week`;
              } else if (rawFrequency.toLowerCase().includes('month')) {
                  sessionsPerWeekString = `${num} Session${num === 1 ? '' : 's'}/month`;
              } else {
                   sessionsPerWeekString = rawFrequency;
              }
          } else {
              sessionsPerWeekString = rawFrequency;
          }
      }

      const paymentPlanParsed = parsePaymentPlan(contactProperties.payment_type);

      const data = {
        dealId: contactProperties.hs_object_id || null,
        jlid: contactProperties.jetlearner_id || jlid,
        learnerName: `${contactProperties.dealname || ''}`.trim(),
        parentName: contactProperties.parent_name || '',
        parentEmail: contactProperties.parent_email || '',
        parentContact: contactProperties.phone_number_deal_ || contactProperties.phone || '',
        course: getCourseLabel(contactProperties.current_course) || '',
        subscriptionTenureMonths: parseInt(contactProperties.subscription_tenure || '0') || 0,
        age: contactProperties.age || '', 
        currentTeacher: getTeacherLabel(contactProperties.current_teacher) || '',
        newTeacher: '', 
        startingDate: contactProperties.module_start_date || '', 
        endDate: contactProperties.module_end_date || '',
        subscriptionStartDate: contactProperties.current_subscription_start_date || '', 
        planName: contactProperties.subscription || '',
        classSessions: parseClassTimings(contactProperties.class_timings || contactProperties.regular_class_day),
        paymentType: paymentPlanParsed.paymentPlanType,
        installmentFrequency: paymentPlanParsed.installmentFrequency,
        customPlanDetails: paymentPlanParsed.customPlanDetails,
        zoomLink: contactProperties.zoom_masked_link || '',
        practiceDocumentLink: contactProperties.learner_practice_document_link || '',
        jetGuideName: getHSUserLabel(contactProperties.jet_guide) || '',
        clsManagerName: getHSUserLabel(contactProperties.cls_manager) || '',
        tpManagerName: getHSUserLabel(contactProperties.teacher_manager) || '',
        currency: contactProperties.deal_currency_code || 'EUR', 
        sessionsPerWeek: sessionsPerWeekString,
        timezone: contactProperties.time_zone || '',
        dealAmount: parseFloat(contactProperties.amount || '0'),
        paymentReceivedDate: contactProperties.stage____payment_trigger_date || null,
        installmentTerms: contactProperties.installment_terms_final || '',
        discount: calculatedDiscount 
      };
      Logger.log('HubSpot data fetched successfully for JLID: ' + jlid);
      Logger.log(data);
      return { success: true, data: data };
    } else {
      Logger.log('No contact found for JLID: ' + jlid);
      return { success: false, message: 'No learner found with this JLID in HubSpot.' };
    }

  } catch (error) {
    Logger.log('Error fetching from HubSpot: ' + error.message + ' Stack: ' + error.stack);
    return { success: false, message: 'Failed to connect to HubSpot: ' + error.message };
  }
}

/**
 * Logs a sent email as an engagement on a HubSpot deal timeline.
 * @param {string} dealId The HubSpot object ID for the deal.
 * @param {string} subject The subject of the email.
 * @param {string} htmlBody The full HTML content of the sent email.
 * @returns {void}
 */
function logEmailToHubspot(dealId, subject, htmlBody) {
  if (!dealId) {
    Logger.log('[HubSpot Logging] Skipped: No Deal ID provided.');
    return;
  }

  Logger.log(`[HubSpot Logging] Logging email for Deal ID: ${dealId}`);
  const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  if (!token) {
    Logger.log('[HubSpot Logging] Failed: HubSpot API token not configured.');
    return;
  }

  const hubspotApiUrl = 'https://api.hubapi.com/crm/v3/objects/emails';

  const requestBody = {
    properties: {
      hs_timestamp: new Date().toISOString(),
      hs_email_subject: subject,
      hs_email_html_body: htmlBody,
      hs_email_direction: 'EMAIL', // Indicates an email sent from your system
      hs_email_status: 'SENT'
    },
    associations: [
      {
        to: {
          id: dealId
        },
        types: [
          {
            // This is the standard HubSpot association type for an email to a deal
            associationCategory: "HUBSPOT_DEFINED",
            associationTypeId: 214 
          }
        ]
      }
    ]
  };

  const options = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(hubspotApiUrl, options);
    const responseCode = response.getResponseCode();
    if (responseCode === 201) { // 201 Created is the success code for this API call
      Logger.log(`[HubSpot Logging] Successfully logged email to deal ${dealId}.`);
    } else {
      Logger.log(`[HubSpot Logging] Error logging email. Status: ${responseCode}, Response: ${response.getContentText()}`);
    }
  } catch (error) {
    Logger.log(`[HubSpot Logging] FATAL ERROR logging email: ${error.message}`);
  }
}


// =============================================
// NEW: HUBSPOT INTEGRATION - Ticket
// =============================================

/**
 * 1. FETCH LATEST TICKET (Production Version)
 * Searches by 'learner_uid' to find specific migration details.
 */
function fetchLatestMigrationTicket(jlid) {
  const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  const hubspotApiUrl = 'https://api.hubapi.com/crm/v3/objects/tickets/search';

  const PIPELINE_ID = '66161281';
  
  const properties = [
    'current_teacher__t_',       
    'new_teacher',               
    'reason_of_migration__t_',   
    'current_course__t_',        
    'current_course',            
    'regular_class_day__t_',     
    'regular_class_time__in_cet_', 
    'subject',
    'createdate' // Asking for it, but will fallback to root
  ];

  const requestBody = {
    filterGroups: [{
      filters: [
        { propertyName: "hs_pipeline", operator: "EQ", value: PIPELINE_ID },
        { propertyName: "learner_uid", operator: "EQ", value: jlid }
      ]
    }],
    properties: properties,
    limit: 10 
  };

  try {
    const options = {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(hubspotApiUrl, options);
    const data = JSON.parse(response.getContentText());

    if (data.results && data.results.length > 0) {
      Logger.log(`Found ${data.results.length} tickets for JLID: ${jlid}`);

      // --- MANUAL SORTING LOGIC (FIXED) ---
      // We use 'ticket.createdAt' (Root Level) because 'properties.createdate' was null
      const sortedTickets = data.results.sort((a, b) => {
        const dateA = new Date(a.createdAt); 
        const dateB = new Date(b.createdAt);
        return dateB - dateA; // Descending (Newest first)
      });

      const latestTicket = sortedTickets[0];
      const props = latestTicket.properties;
      
      Logger.log(`✅ Selected Newest Ticket:`);
      Logger.log(`   ID: ${latestTicket.id}`);
      Logger.log(`   Root CreatedAt: ${latestTicket.createdAt}`); // This should show the real date
      Logger.log(`   Subject: ${props.subject}`);

      return {
        found: true,
        oldTeacher: getTeacherLabel(props.current_teacher__t_) || '', 
        newTeacher: getTeacherLabel(props.new_teacher) || '',
        reason: props.reason_of_migration__t_ || '',
        ticketCourse: getCourseLabel(props.current_course__t_ || props.current_course) || '', 
        classDay: props.regular_class_day__t_ || '',
        classTime: props.regular_class_time__in_cet_ || ''
      };
    }
    
    Logger.log(`No tickets found for JLID: ${jlid}`);
    return { found: false };

  } catch (e) {
    Logger.log("Error fetching ticket: " + e.message);
    return { found: false };
  }
}


/**
 * 2. HYBRID FETCH CONTROLLER
 * Called by the frontend. Combines Deal Profile + Ticket Migration Data.
 */
function fetchMigrationHybridData(jlid) {
  // A. Fetch Deal Data
  const dealResult = fetchHubspotByJlid(jlid);
  if (!dealResult.success) return dealResult;

  const finalData = dealResult.data;

  // B. Fetch Ticket Data
  const ticketResult = fetchLatestMigrationTicket(jlid);

  if (ticketResult.found) {
    // 1. Teachers
    if (ticketResult.oldTeacher) finalData.currentTeacher = ticketResult.oldTeacher; 
    if (ticketResult.newTeacher) finalData.newTeacher = ticketResult.newTeacher; 
    
    // 2. Reason
    if (ticketResult.reason) finalData.migrationReason = ticketResult.reason;

    // 3. Course (CRITICAL FIX: Ticket course overrides Deal course)
    if (ticketResult.ticketCourse) {
        finalData.course = ticketResult.ticketCourse;
    }

    // 4. Schedule
    if (ticketResult.classDay || ticketResult.classTime) {
        finalData.ticketSchedule = {
            day: ticketResult.classDay,
            time: ticketResult.classTime
        };
    }

    finalData.source = "Hybrid (Deal + Ticket)";
  } else {
    finalData.source = "Deal Only";
  }

  return { success: true, data: finalData };
}



/**
 * Converts an internal course value from HubSpot to a user-friendly label using data from 'Course HS values' sheet.
 * @param {string} internalValue The internal HubSpot course value.
 * @returns {string|null} The user-friendly course label, or null if not found.
 */
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

/**
 * Converts an internal HubSpot user ID to a user-friendly name using data from 'HS User Values' sheet.
 * @param {string} internalValue The internal HubSpot user ID.
 * @returns {string|null} The user-friendly name, or null if not found.
 */
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

/**
 * Converts an internal HubSpot teacher value to a user-friendly name using data from 'Teacher HS values' sheet.
 * @param {string} internalValue The internal HubSpot teacher value.
 * @returns {string|null} The user-friendly teacher name, or null if not found.
 */
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

// =============================================
// NEW: INVOICE MANAGEMENT FUNCTIONS
// =============================================

/**
 * Fetches invoice product data from the 'Invoice Products' sheet.
 * @returns {Array<Object>} An array of objects, each representing an invoice product plan.
 */
function getInvoiceProductsData() {
  Logger.log('getInvoiceProductsData called');
  try {
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.INVOICE_PRODUCTS);
    if (sheetData.length < 2) { 
      Logger.log('Invoice Products sheet is empty or only has headers.');
      return [];
    }
    const headers = sheetData[0];
    const products = sheetData.slice(1).map(row => {
      const product = {};
      headers.forEach((header, i) => {
        product[header.trim()] = row[i];
      });
      return product;
    }).filter(p => p['Plan Name'] !== '');
    Logger.log(`Found ${products.length} invoice products.`);
    return products;
  } catch (error) {
    Logger.log('Error getting invoice products data: ' + error.message);
    return [];
  }
}

/**
 * NEW: Fetches live exchange rates from an external API.
 * Uses a robust error handling and fallback mechanism.
 * @returns {object} An object of currency codes and their rates relative to EUR, or {} on failure.
 */
function getLiveExchangeRates() {
  Logger.log('Fetching live exchange rates from API.');
  try {
    const response = UrlFetchApp.fetch(CONFIG.EXCHANGE_RATE_API_URL, {
      muteHttpExceptions: true 
    });

    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      Logger.log(`Exchange rate API returned non-200 status: ${responseCode}. Response: ${response.getContentText()}`);
      return {}; 
    }

    const jsonResponse = JSON.parse(response.getContentText());

    if (jsonResponse && jsonResponse.result === 'success' && jsonResponse.conversion_rates) {
      const rates = jsonResponse.conversion_rates;
      
      delete rates.EUR;
      delete rates.GBP;
      delete rates.USD;

      Logger.log(`Successfully fetched and filtered ${Object.keys(rates).length} live exchange rates.`);
      return rates;
    } else {
      Logger.log(`Exchange rate API response was invalid or unsuccessful. Response: ${JSON.stringify(jsonResponse)}`);
      return {}; 
    }
  } catch (error) {
    Logger.log(`Error fetching or parsing live exchange rates: ${error.message}. Stack: ${error.stack}`);
    return {}; 
  }
}


/**
 * Fetches a conversion rate from EUR to a target currency.
 * It prioritizes a live API call and uses a hardcoded list as a reliable fallback.
 * @param {string} toCurrency The currency to convert to (e.g., 'INR', 'USD').
 * @returns {number} The conversion rate (e.g., 1 EUR = X of the target currency).
 */
function getConversionRate(toCurrency) {
  toCurrency = toCurrency.toUpperCase();
  if (toCurrency === 'EUR') {
    return 1.0;
  }

  if (!_sheetDataCache['liveRates']) {
    Logger.log('Fetching live currency rates for this execution...');
    _sheetDataCache['liveRates'] = getLiveCurrencyRates(); 
  }
  const liveRates = _sheetDataCache['liveRates'];

  if (liveRates && liveRates[toCurrency]) {
    Logger.log(`Using LIVE rate for EUR to ${toCurrency}: ${liveRates[toCurrency]}`);
    return liveRates[toCurrency];
  }
  
  const fallbackRates = {
    'EUR': 1.0, 'USD': 1.1700, 'GBP': 0.8665, 'INR': 103.16, 'CHF': 0.9348,
    'AED': 4.273, 'CAD': 1.6224, 'AUD': 1.7682, 'JPY': 172.50, 'SGD': 1.50,
    'HKD': 9.1187, 'ZAR': 20.57, 'CNY': 8.3387, 'NZD': 1.9704, 'SEK': 10.951,
    'NOK': 11.6195, 'DKK': 7.45, 'MXN': 21.8069, 'BRL': 6.3207
  };
  
  const fallbackRate = fallbackRates[toCurrency];
  if (fallbackRate !== undefined) {
    Logger.log(`Using FALLBACK hardcoded rate for EUR to ${toCurrency}: ${fallbackRate}`);
    return fallbackRate;
  }

  Logger.log(`No live or hardcoded rate found for EUR to ${toCurrency}. Returning 1 as a safe default.`);
  return 1; 
}

/**
 * Calculates all pricing details for an invoice (now only for general invoices).
 * @param {object} formData - The data from the invoice form.
 * @param {boolean} previewOnly - If true, only calculates and returns, does not apply discount capping messages.
 * @returns {object} An object containing all calculated pricing details.
 */
function calculateInvoicePricing(formData, previewOnly = false) {
    let effectiveBasePrice = 0;
    let discount = parseFloat(formData.discount || '0');
    let customCurrencyExtraDiscountPercentage = parseFloat(formData.customCurrencyExtraDiscountPercentage || '0');
    let finalCurrencySymbol = getCurrencySymbol(formData.currency); 

    const userSelectedTenureMonths = parseInt(formData.subscriptionTenure || '0');
    const freeClasses = parseInt(formData.freeClasses || '0');
    const sessionsPerWeekNum = parseInt((formData.sessionsPerWeek && String(formData.sessionsPerWeek).split(' ')[0]) || '0');
    const customPaidAmount = parseFloat(formData.customPaidAmount || '0');
    const invoiceProducts = getInvoiceProductsData();
    const selectedPlan = invoiceProducts.find(p => p['Plan Name'] === formData.planName);

    if (!selectedPlan) {
        throw new Error(`Invoice plan '${formData.planName}' not found.`);
    }
    
    // FIX: Force non-numbers to 0
    let fixedClassesPerPlan = parseInt(selectedPlan['Fixed Classes']);
    if (isNaN(fixedClassesPerPlan)) fixedClassesPerPlan = 0;

    let targetCurrencyCode = formData.currency; 
    let finalConversionRate = 1.0; 

    if (targetCurrencyCode === 'CUSTOM') {
        targetCurrencyCode = (formData.customCurrencyCode || 'EUR').toUpperCase(); 
        finalConversionRate = (formData.customCurrencyRate && parseFloat(formData.customCurrencyRate) > 0) 
            ? parseFloat(formData.customCurrencyRate) 
            : getConversionRate(targetCurrencyCode);
        finalCurrencySymbol = getCurrencySymbol(targetCurrencyCode);
    } else if (!['EUR', 'USD', 'GBP'].includes(targetCurrencyCode)) {
        finalConversionRate = getConversionRate(targetCurrencyCode);
    }
    
    let planBasePriceForDefaultTenure = 0;
    if (formData.currency === 'USD') {
        planBasePriceForDefaultTenure = parseFloat(String(selectedPlan['Base Price USD'] || '0').replace(/[^0-9.-]/g, '')) || 0;
    } else if (formData.currency === 'GBP') {
        planBasePriceForDefaultTenure = parseFloat(String(selectedPlan['Base Price GBP'] || '0').replace(/[^0-9.-]/g, '')) || 0;
    } else {
        planBasePriceForDefaultTenure = parseFloat(String(selectedPlan['Base Price EUR'] || '0').replace(/[^0-9.-]/g, '')) || 0;
    }

    const selectedPlanDefaultMonthsTenure = parseInt(selectedPlan['Months Tenure'] || '1');
    const monthlyRate = selectedPlanDefaultMonthsTenure > 0 ? (planBasePriceForDefaultTenure / selectedPlanDefaultMonthsTenure) : 0;
    
    effectiveBasePrice = monthlyRate * userSelectedTenureMonths * finalConversionRate;
    if (isNaN(effectiveBasePrice)) effectiveBasePrice = 0;

    if (formData.currency === 'CUSTOM' && effectiveBasePrice > 0 && customCurrencyExtraDiscountPercentage > 0) {
        effectiveBasePrice *= (1 - customCurrencyExtraDiscountPercentage / 100);
    }

    if (discount > effectiveBasePrice) discount = effectiveBasePrice;
    
    let finalTotal = effectiveBasePrice - discount;
    if (finalTotal < 0) finalTotal = 0;

    const numInstallments = parseInt(formData.numberOfInstallments || '1');
    const isInstallment = formData.paymentType === 'Installment' && numInstallments > 0;

    let amountPaid = parseFloat(formData.customPaidAmount) || 0;
    if (amountPaid === 0) {
        if (isInstallment && numInstallments > 0) {
            amountPaid = finalTotal / numInstallments;
        } else {
            amountPaid = finalTotal;
        }
    }
    
    const balanceDue = finalTotal - amountPaid;

    let totalClasses = (fixedClassesPerPlan > 0) ? fixedClassesPerPlan + freeClasses : (userSelectedTenureMonths * 4 * sessionsPerWeekNum) + freeClasses;
    totalClasses = Math.max(0, totalClasses);
    
    const unitPrice = totalClasses > 0 ? (effectiveBasePrice / totalClasses) : 0;

    let weeksRequired = (sessionsPerWeekNum > 0) ? Math.ceil(totalClasses / sessionsPerWeekNum) : 0;
    
    let finalEndDate;
    if (formData.endDate && formData.endDate.trim() !== '') {
        finalEndDate = new Date(formData.endDate);
    } else {
        const startDate = new Date(formData.startDate);
        finalEndDate = new Date(startDate);
        if (weeksRequired > 0) {
            finalEndDate.setDate(startDate.getDate() + (weeksRequired * 7) - 1);
        }
    }

    const installments = [];
    if (isInstallment && formData.installmentType && formData.dueDayToPay) {
        const billingAnchorDate = formData.firstPaymentDate ? new Date(formData.firstPaymentDate) : new Date(formData.startDate);
        const dueDay = parseInt(formData.dueDayToPay);
        let lastDueDate = new Date(billingAnchorDate);

        let monthIncrement = 1;
            switch (formData.installmentType) {
            case 'Alternate': monthIncrement = 2; break;
            case 'Quarterly': monthIncrement = 3; break;
            case '4 Months': monthIncrement = 4; break;
            case '5 Months': monthIncrement = 5; break;
            case '6 Months': monthIncrement = 6; break;
            case '7 Months': monthIncrement = 7; break;
            case '8 Months': monthIncrement = 8; break;
            case '9 Months': monthIncrement = 9; break;
            case '10 Months': monthIncrement = 10; break;
            case '11 Months': monthIncrement = 11; break;
            case '12 Months': monthIncrement = 12; break;
        }
        if (customPaidAmount > 0) {
            const installmentAmount = balanceDue > 0 && numInstallments > 0 ? balanceDue / numInstallments : 0;
            for (let i = 0; i < numInstallments; i++) {
                let nextDueDate = new Date(lastDueDate);
                if (i > 0) {
                    nextDueDate.setMonth(nextDueDate.getMonth() + monthIncrement);
                }
                nextDueDate.setDate(dueDay);
                installments.push({ number: i + 1, amount: installmentAmount, isPaid: false, dueDate: nextDueDate, dueDateFormatted: formatDateDDMMYYYY(nextDueDate) });
                lastDueDate = nextDueDate;
            }
        } else {
            const installmentAmount = finalTotal / numInstallments;
            installments.push({ number: 1, amount: installmentAmount, isPaid: true, dueDate: billingAnchorDate, dueDateFormatted: formatDateDDMMYYYY(billingAnchorDate) });
            let lastDueDate = new Date(billingAnchorDate);
            for (let i = 1; i < numInstallments; i++) {
                let nextDueDate = new Date(lastDueDate);
                nextDueDate.setMonth(nextDueDate.getMonth() + monthIncrement);
                nextDueDate.setDate(dueDay);
                installments.push({ number: i + 1, amount: installmentAmount, isPaid: false, dueDate: nextDueDate, dueDateFormatted: formatDateDDMMYYYY(nextDueDate) });
                lastDueDate = nextDueDate;
            }
        }
    }

    return {
        unitPrice: unitPrice,
        effectiveBasePrice: effectiveBasePrice,
        discount: discount,
        finalTotal: finalTotal,
        amountPaid: amountPaid,
        balanceDue: balanceDue,
        currencySymbol: finalCurrencySymbol,
        subscriptionTenureMonths: userSelectedTenureMonths,
        startDateFormatted: formatDate(new Date(formData.startDate)),
        endDateFormatted: formatDate(finalEndDate),
        paymentType: formData.paymentType,
        numberOfInstallments: numInstallments,
        displayTotalSessions: totalClasses,
        planDescription: formData.planName,
        installments: installments,
        upfrontDueDate: formData.upfrontDueDate ? formatDate(new Date(formData.upfrontDueDate)) : null
    };
}

/**
 * Generates an invoice PDF, saves it to Drive, and emails it to the parent.
 * Can also be used to generate a preview HTML.
 * @param {object} formData - The data submitted from the invoice form.
 * @param {boolean} previewOnly - If true, only returns HTML for preview, does not save/email.
 * @returns {object} An object with success status and message, or HTML for preview.
 */
function generateInvoicePDFAndEmail(formData) { 
  Logger.log(`generateInvoicePDFAndEmail called for Learner: ${formData.learnerName}`);
  let trackingId = null; // To capture the ID for logging, even on failure
  
  try {
    const validationErrors = validateInvoiceData(formData); 
    if (validationErrors.length > 0) {
      throw new Error('Validation failed: ' + validationErrors.join(', '));
    }

    const pricingDetails = calculateInvoicePricing(formData); 
    const invoiceHtml = getInvoiceHTML(formData, pricingDetails);

    const pdfName = `Invoice-${formData.learnerName.replace(/\s/g, '_')}-${formData.jlid || 'N_A'}-${new Date().toISOString().split('T')[0]}.pdf`;
    const blob = Utilities.newBlob(invoiceHtml, 'text/html')
                             .getAs(MimeType.PDF)
                             .setName(pdfName);

    // Save to Google Drive
    if (CONFIG.DRIVE_FOLDER_ID) {
        try {
            DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID).createFile(blob);
            Logger.log(`Invoice PDF saved to Drive: ${pdfName}`);
        } catch (e) {
            Logger.log(`Warning: Could not save invoice PDF to Drive: ${e.message}`);
            // Non-fatal error, we can still try to email it.
        }
    }

    // Prepare the simple wrapper email body for tracking
    const emailHtmlBody = `
     <!DOCTYPE html>
      <html>
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>JetLearn Invoice</title>
      </head>
      <body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f5f5f5;">
          <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f5f5f5; padding: 20px;">
              <tr>
                  <td align="center">
                      <table width="600" cellpadding="0" cellspacing="0" style="background-color: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
                          
                          <!-- Header -->
                          <tr>
                              <td style="background: linear-gradient(135deg, #FFD700 0%, #FFA500 100%); padding: 30px 40px; text-align: center;">
                                  <h1 style="margin: 0; color: #000; font-size: 32px; font-weight: 700; letter-spacing: 0.5px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
                                      JetLearn
                                  </h1>
                                  <p style="margin: 8px 0 0; color: #333; font-size: 14px; font-weight: 500;">
                                      World's Top Online AI Academy
                                  </p>
                              </td>
                          </tr>
                          
                          <!-- Content -->
                          <tr>
                              <td style="padding: 40px;">
                                  <p style="color: #666; line-height: 1.6; margin: 0 0 20px; font-size: 16px;">
                                      Dear ${formData.parentName},
                                  </p>
                                  
                                  <p style="color: #666; line-height: 1.6; margin: 0 0 20px; font-size: 16px;">
                                      Please find attached the invoice for <strong style="color: #000;">${formData.learnerName}</strong>'s reference.
                                  </p>
                                  
                                  <p style="color: #666; line-height: 1.6; margin: 0 0 20px; font-size: 16px;">
                                      We appreciate you giving your child the best opportunity to learn and grow. We are honored to be part of this journey!
                                  </p>
                                  
                                  <p style="color: #666; line-height: 1.6; margin: 25px 0 0; font-size: 16px;">
                                      Best regards,<br>
                                      <strong style="color: #000;">The JetLearn Team</strong>
                                  </p>
                              </td>
                          </tr>
                          
                          <!-- Footer -->
                          <tr>
                              <td style="background-color: #000; padding: 25px 40px; text-align: center;">
                                  <p style="margin: 0; color: #FFD700; font-size: 12px;">
                                      © 2025 JetLearn. Empowering kids to lead in the age of AI.
                                  </p>
                              </td>
                          </tr>
                          
                      </table>
                  </td>
              </tr>
          </table>
      </body>
      </html>
    `;

    // Use the central tracked email service
    const emailResult = sendTrackedEmail({
      to: formData.parentEmail,
      subject: `JetLearn Invoice for ${formData.learnerName}`,
      htmlBody: emailHtmlBody,
      jlid: formData.jlid,
      attachments: [blob]
    });
    trackingId = emailResult.trackingId;

    // Log the successful action to the main audit log
    logAction('Invoice Sent', formData.jlid || '', formData.learnerName, '', '', formData.planName, 'Success', `Invoice PDF sent to ${formData.parentEmail}. TID: ${trackingId}`);
    
    return { success: true, message: 'Invoice generated and emailed successfully!' };

  } catch (error) {
    Logger.log('Error in generateInvoicePDFAndEmail: ' + error.message);
    // Log the failed action to the main audit log
    logAction('Invoice Failed', formData.jlid || '', formData.learnerName, '', '', formData.planName, 'Failed', `Error: ${error.message}. TID Attempt: ${trackingId}`);
    return { success: false, message: 'Failed to send invoice email: ' + error.message };
  } 
}


/**
 * Generates an invoice PDF, saves it to Drive, and returns it as a base64 encoded string for client download.
 * @param {object} formData - The data submitted from the invoice form.
 * @returns {object} An object with success status, message, and base64 PDF data.
 */
function generateInvoicePDFForDownload(formData) { 
  Logger.log(`generateInvoicePDFForDownload called for Learner: ${formData.learnerName}`);

  const validationErrors = validateInvoiceData(formData); 
  if (validationErrors.length > 0) {
    return { success: false, message: 'Validation failed: ' + validationErrors.join(', ')} ;
  }

  if (!formData.learnerEmail || formData.learnerEmail.trim() === '') {
      formData.learnerEmail = formData.parentEmail;
  }

  const pricingDetails = calculateInvoicePricing(formData, false); 

  const htmlBody = getInvoiceHTML(formData, pricingDetails);

  const pdfName = `Invoice-${formData.learnerName.replace(/\s/g, '_')}-${formData.jlid || 'N_A'}-${new Date().toISOString().split('T')[0]}_${Date.now()}.pdf`;
  const blob = Utilities.newBlob(htmlBody, 'text/html')
                           .getAs(MimeType.PDF)
                           .setName(pdfName);

  if (!CONFIG.DRIVE_FOLDER_ID || CONFIG.DRIVE_FOLDER_ID === '1_exampleFolderID1234567890abcdef') {
      Logger.log("DRIVE_FOLDER_ID is not configured. Invoice PDF will not be saved to Drive.");
  } else {
      try {
          const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
          folder.createFile(blob);
          Logger.log(`Invoice PDF saved to Drive for download: ${pdfName}`);
      } catch (e) {
          Logger.log(`Error saving invoice PDF to Drive for download: ${e.message}`);
      }
  }

  try {
      const base64Data = Utilities.base64Encode(blob.getBytes());
      return { success: true, message: 'PDF generated successfully for download.', filename: pdfName, data: base64Data };
  } catch (error) {
      Logger.log('Error encoding PDF to base64: ' + error.message);
      return { success: false, message: 'Failed to encode PDF for download: ' + error.message };
  }
}

/**
 * Validates the data for invoice generation (now only for general invoices).
 * @param {object} formData - The data from the invoice form.
 * @returns {Array<string>} An array of error messages.
 */
function validateInvoiceData(formData) { 
  const errors = [];

  if (!formData.learnerName || formData.learnerName.trim() === '') errors.push('Learner Name is required');
  if (formData.learnerEmail && !isValidEmail(formData.learnerEmail)) errors.push('Learner Email, if provided, must be valid.');

  if (!formData.parentName || formData.parentName.trim() === '') errors.push('Parent Name is required');
  if (!formData.parentEmail || !isValidEmail(formData.parentEmail)) errors.push('Valid Parent Email is required');
  if (!formData.parentContact || formData.parentContact.trim() === '') errors.push('Parent Contact is required');
  if (!formData.planName || formData.planName.trim() === '') errors.push('Plan Name is required');
  if (!formData.currency || formData.currency.trim() === '') errors.push('Currency is required');
  if (!formData.sessionsPerWeek || formData.sessionsPerWeek.trim() === '') errors.push('Sessions per Week is required');
  if (parseInt(formData.subscriptionTenure || '0') < 0) errors.push('Subscription Tenure must be a non-negative number.');
  if (!formData.startDate || formData.startDate.trim() === '') errors.push('Start Date is required');
  if (!formData.endDate || formData.endDate.trim() === '') errors.push('End Date is required'); 
  if (new Date(formData.endDate) < new Date(formData.startDate)) errors.push('End Date cannot be before Start Date.');
  if (!formData.paymentType || formData.paymentType.trim() === '') errors.push('Payment Type is required');

  let discount = parseFloat(formData.discount || '0'); 
  const customPaidAmount = parseFloat(formData.customPaidAmount || '0'); 
  const freeClasses = parseInt(formData.freeClasses || '0'); 

  if (discount < 0) errors.push('Discount cannot be negative.');
  if (customPaidAmount < 0) errors.push('Partial Payment Received cannot be negative.');
  if (freeClasses < 0) errors.push('Free Classes cannot be negative.'); 

  const invoiceProducts = getInvoiceProductsData();
  const selectedPlan = invoiceProducts.find(p => p['Plan Name'] === formData.planName);

  if (selectedPlan && formData.currency && (parseInt(formData.subscriptionTenure || '0') >= 0)) {
      let targetCurrencyCode = formData.currency;
      let finalConversionRateFromEUR = 1;
      let customCurrencyExtraDiscountPercentage = parseFloat(formData.customCurrencyExtraDiscountPercentage || '0');

      if (targetCurrencyCode === 'CUSTOM') {
          if (!formData.customCurrencyCode || formData.customCurrencyCode.trim() === '') errors.push('Custom Currency Code is required for custom currency.');
          targetCurrencyCode = (formData.customCurrencyCode || 'EUR').toUpperCase(); 
          finalConversionRateFromEUR = (formData.customCurrencyRate && parseFloat(formData.customCurrencyRate) > 0)
            ? parseFloat(formData.customCurrencyRate)
            : getConversionRate('EUR', targetCurrencyCode);

          if (finalConversionRateFromEUR <= 0) {
              errors.push(`Invalid conversion rate from EUR to ${targetCurrencyCode}.`);
          }
          if (customCurrencyExtraDiscountPercentage < 0 || customCurrencyExtraDiscountPercentage > 100) {
              errors.push('Custom currency discount percentage must be between 0 and 100.');
          }
      } else {
          finalConversionRateFromEUR = getConversionRate('EUR', targetCurrencyCode);
      }

      const planBasePriceForDefaultTenure = parseFloat(String(selectedPlan['Base Price EUR'] || '0').replace(/[^0-9.-]/g, '')) || 0;
      const selectedPlanDefaultMonthsTenure = parseInt(selectedPlan['Months Tenure'] || '1'); 
      const userSelectedTenureMonths = parseInt(formData.subscriptionTenureMonths || '0');

      const monthlyRateEUR = selectedPlanDefaultMonthsTenure > 0 ? (planBasePriceForDefaultTenure / selectedPlanDefaultMonthsTenure) : 0;
      let effectiveBasePrice = monthlyRateEUR * userSelectedTenureMonths;
      if (isNaN(effectiveBasePrice)) effectiveBasePrice = 0;

      if (formData.currency === 'CUSTOM' && effectiveBasePrice > 0 && customCurrencyExtraDiscountPercentage > 0) {
          effectiveBasePrice *= (1 - customCurrencyExtraDiscountPercentage / 100);
      }
      effectiveBasePrice *= finalConversionRateFromEUR;

      // if (discount > effectiveBasePrice) {
      //     errors.push('Discount cannot exceed the calculated base price. It will be capped automatically.');
      // }
  }

  if (formData.paymentType === 'Installment') {
      const numInstallments = parseInt(formData.numberOfInstallments || '0');
      if (isNaN(numInstallments) || numInstallments <= 0) {
          errors.push('Number of Installments must be a positive number for installment plans.');
      }
      if (!formData.installmentType || formData.installmentType.trim() === '') {
          errors.push('Installment Type is required for installment plans.');
      }
      const dueDay = parseInt(formData.dueDayToPay || '0');
      if (isNaN(dueDay) || dueDay < 1 || dueDay > 31) {
          errors.push('Due Day to Pay must be a number between 1 and 31 for installment plans.');
      }
  }

  return errors;
}

/**
 * NEW: Sends the parent onboarding email and, if requested, generates and attaches an invoice in one step.
 * This version prioritizes the zoomLink from HubSpot.
 * @param {object} formData - The complete form data from the client, including email and invoice details.
 * @param {Array<Object>} attachmentsBase64 - Any user-uploaded files in base64 format.
 * @returns {object} A result object with success status and message.
 */
function sendParentOnboardingWithInvoice(formData, attachmentsBase64) {
  Logger.log(`Starting combined parent/teacher/WhatsApp onboarding for learner: ${formData.learnerName}`);
  
  let parentTrackingId = null;
  let teacherTrackingId = null;
  let watiStatus = []; 
  let overallStatus = 'Failed';
  let logNotes = '';

  try {
    // --- 1. VALIDATE & PREPARE DATA ---
    if (!formData.parentEmail || !isValidEmail(formData.parentEmail)) throw new Error('Parent Email is required.');
    if (!formData.teacherName) throw new Error('A teacher must be assigned.');

    // --- 2. LOOKUP TEACHER & APPLY PREFIX ---
    const teacherData = getTeacherData();
    const teacherInfo = teacherData.find(t => 
      String(t.name).trim().toLowerCase() === String(formData.teacherName).trim().toLowerCase()
    );

    if (!teacherInfo || !isValidEmail(teacherInfo.email)) {
        throw new Error(`Teacher '${formData.teacherName}' not found or has an invalid email.`);
    }

    if (teacherInfo.prefix) {
        formData.teacherName = `${teacherInfo.prefix} ${teacherInfo.name}`;
    }

    // --- 3. PROCESS ATTACHMENTS & CALCULATIONS ---
    if (!formData.zoomLink || !formData.zoomLink.startsWith('http')) {
      formData.zoomLink = `https://live.jetlearn.com/join/${cleanJlidForZoom(formData.jlid)}`;
    }

    let allAttachments = [];
    if (attachmentsBase64 && attachmentsBase64.length > 0) {
      allAttachments.push(...uploadAttachments(attachmentsBase64));
    }

    const pricingDetails = calculateInvoicePricing(formData);
    formData.sessions = pricingDetails.displayTotalSessions;

    // >>> REMOVED THE BREAKING DATE FORMATTING LINES HERE <<< 
    // We keep formData.startDate as YYYY-MM-DD so the Invoice and Email Template don't crash.

    // Generate Invoice PDF
    if (formData.generateAndAttachInvoice) {
      const validationErrors = validateInvoiceData(formData);
      if (validationErrors.length > 0) throw new Error('Invoice data is invalid: ' + validationErrors.join('; '));
      
      const invoiceHtml = getInvoiceHTML(formData, pricingDetails);
      const pdfName = `Invoice-${formData.learnerName.replace(/\s/g, '_')}-${formData.jlid || 'NA'}.pdf`;
      const invoiceBlob = Utilities.newBlob(invoiceHtml, 'text/html').getAs(MimeType.PDF).setName(pdfName);
      allAttachments.push(invoiceBlob);
    }

    // --- 4. SEND TRACKED EMAIL TO PARENT ---
    const parentSubject = `Welcome to JetLearn - Enrollment Details for ${formData.learnerName}`;
    const parentHtmlBody = getParentOnboardingEmailHTML(formData);
    
    const parentEmailResult = sendTrackedEmail({
      to: formData.parentEmail, 
      subject: parentSubject, 
      htmlBody: parentHtmlBody, 
      jlid: formData.jlid, 
      attachments: allAttachments
    });
    parentTrackingId = parentEmailResult.trackingId;
    
    if (formData.dealId) logEmailToHubspot(formData.dealId, parentSubject, parentHtmlBody);

    // --- 5. SEND WHATSAPP SEQUENCE ---
    if (formData.sendWhatsapp) { 
        watiStatus = sendOnboardingWhatsAppSequence(formData);
        Logger.log("WATI Onboarding Status: " + JSON.stringify(watiStatus));
    } else {
        watiStatus = ["Skipped (Checkbox unchecked)"];
    }

    // --- 6. SEND TRACKED EMAIL TO TEACHER ---
    const clsManagerEmail = findClsEmailByManagerName(formData.clsManager);
    const ccList = new Set();
    if(clsManagerEmail) ccList.add(clsManagerEmail);
    if(teacherInfo.tpManagerEmail) ccList.add(teacherInfo.tpManagerEmail);

    const teacherSubject = `New Learner Onboarded || ${formData.learnerName} (${formData.jlid || 'N/A'})`;
    const teacherHtmlBody = getOnboardingEmailHTML(formData);
    
    const teacherEmailResult = sendTrackedEmail({
      to: teacherInfo.email,
      cc: Array.from(ccList).join(','),
      subject: teacherSubject,
      htmlBody: teacherHtmlBody,
      jlid: formData.jlid
    });
    teacherTrackingId = teacherEmailResult.trackingId;

    // --- 7. FINALIZE ---
    overallStatus = 'Success';
    logNotes = `Parent Email Sent (TID: ${parentTrackingId}). Teacher Email Sent (TID: ${teacherTrackingId}). WhatsApp: [${watiStatus.join(', ')}].`;
    
    return { success: true, message: 'Onboarding emails and WhatsApp messages sent successfully.' };

  } catch (error) {
    Logger.log(`Error in sendParentOnboardingWithInvoice: ${error.message}\nStack: ${error.stack}`);
    logNotes = `Failed during onboarding process: ${error.message}.`;
    return { success: false, message: `Failed to complete onboarding: ${error.message}` };

  } finally {
    logAction('New Learner Onboarded', formData.jlid, formData.learnerName, '', formData.teacherName, formData.course, overallStatus, logNotes);
  }
}

// =============================================
// ================== NEW: AGENTIC AI AUDIT ==================
// =============================================

/**
 * Uses Gemini AI to act as a QA agent, comparing data from three sources.
 * @param {object} hubspotProps - The properties object from a HubSpot deal.
 * @param {object} noteData - The data object parsed from a sales note.
 * @param {object} calendarData - The verification results from Google Calendar.
 * @param {string} rawNote - The raw text of the sales note for context.
 * @returns {object} An object with `isMismatch: boolean` and `summary: string`.
 */
function getAIAgentAnalysis(hubspotProps, noteData, calendarData, rawNote) {
  Logger.log(`[AI Agent] Analyzing JLID: ${hubspotProps.jetlearner_id}`);
  const GOOGLE_API_KEY = PropertiesService.getScriptProperties().getProperty('GOOGLE_GENERATIVE_AI_KEY');
  if (!GOOGLE_API_KEY) {
    return { isMismatch: true, summary: "AI Agent could not run: API key not configured." };
  }

  const model = 'gemini-2.5-flash';
  const endpoint = `https://generativelanguage.googleapis.com/v1beta/${model}:generateContent?key=${GOOGLE_API_KEY}`;

  const prompt = `
    You are an automated Quality Assurance Agent for JetLearn. Your task is to verify a new learner's onboarding data by comparing three sources: HubSpot (the system of record), the Sales Note (human-entered data), and Google Calendar (scheduling confirmation).

    Your response MUST be a single, valid JSON object with two keys: "isMismatch" (a boolean) and "summary" (a string).
    - "isMismatch" should be \`true\` if you find ANY discrepancy, no matter how small. Otherwise, it should be \`false\`.
    - "summary" should be a concise, bulleted list of your findings. If there are mismatches, list ONLY the discrepancies. If everything matches, state "All key data points align across HubSpot, Sales Note, and Calendar."

    Analyze the following data for learner ${hubspotProps.dealname} (${hubspotProps.jetlearner_id}):

    1. HubSpot Data:
    ${JSON.stringify(hubspotProps, null, 2)}

    2. Parsed Sales Note Data:
    ${JSON.stringify(noteData, null, 2)}

    3. Google Calendar Verification:
    ${JSON.stringify(calendarData, null, 2)}

    4. Full Sales Note Text (for context, especially for teacher preferences):
    ---
    ${rawNote || "No note provided."}
    ---

    Your verification checklist:
    - **Amount & Currency:** Does the amount and currency match between HubSpot and the Sales Note?
    - **Subscription Term:** Does the subscription tenure (in months) match?
    - **Committed Classes:** Does the number of classes match between HubSpot, the note, and the calendar events found?
    - **Dates:** Does the subscription start/end date in HubSpot align with the first/last class in the calendar?
    - **Teacher Preference (CRITICAL):** Read the full sales note. Is there a specific teacher request or a note about a teacher the parent *dislikes*? Does the assigned 'current_teacher' in HubSpot respect this request? This is a high-priority check.

    Now, provide your analysis in the specified JSON format.
  `;

  try {
    const payload = { contents: [{ parts: [{ text: prompt }] }] };
    const response = callGenerativeAIWithRetry(endpoint, payload); // Use the new retry wrapper
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      let textPart = jsonResponse?.candidates?.[0]?.content?.parts?.[0]?.text || '{"isMismatch": true, "summary": "AI response was empty or malformed."}';
      textPart = textPart.replace(/```json/g, '').replace(/```/g, '').trim();
      Logger.log(`[AI Agent] Analysis for ${hubspotProps.jetlearner_id}: ${textPart}`);
      return JSON.parse(textPart);
    }
    
    Logger.log(`[AI Agent] API Error (${responseCode}): ${responseBody}`);
    return { isMismatch: true, summary: `AI Agent API Error (${responseCode}).` };
  } catch (error) {
    Logger.log('[AI Agent] Error calling or parsing AI API: ' + error.message);
    return { isMismatch: true, summary: `AI Agent connection error: ${error.message}` };
  }
}

/**
 * Uses an AI agent to analyze a HubSpot deal and its sales notes to generate a structured invoice plan.
 * @param {string} jlid The JetLearn ID of the deal to analyze.
 * @returns {object} A result object with success status and the structured invoice plan.
 */
function generateSmartInvoicePlan(jlid) {
  try {
    // 1. Fetch All Necessary Data from HubSpot
    const hubspotResult = fetchHubspotByJlid(jlid);
    if (!hubspotResult.success) {
      throw new Error("Could not fetch HubSpot data for this JLID. Please check if the JLID is correct and exists in HubSpot.");
    }
    const dealId = hubspotResult.data.dealId;
    const noteText = fetchLatestSalesNoteForDeal(dealId);
    
    // 2. Construct the AI Prompt (The "Training")
    const prompt = `
      You are an expert financial assistant for JetLearn, tasked with creating a precise JSON object for invoicing based on raw data.

      **Your Core Directives:**
      1.  **Source of Truth:** The human-written 'Sales Note Text' is the absolute source of truth. If it mentions specific payment amounts, schedules, or discounts, those values MUST be used, overriding any conflicting 'HubSpot Properties'.
      2.  **Calculate Implied Discount:** The standard price is 149 per month. Calculate the 'Total Standard Price' as (subscription_tenure * 149). The 'discount' is always (Total Standard Price - HubSpot Deal Amount). For an Annual plan (12 months), the standard price is 1788. If the deal amount is 600, you MUST calculate the discount as 1188.
      3.  **Handle Uneven Installments:** If the sales note describes an uneven payment plan (e.g., '250 paid, 350 due next month'), you MUST create an 'installments' array that exactly matches this schedule. Mark payments already made as 'isPaid: true'.
      4.  **Strict JSON Output:** Your entire response MUST be ONLY the JSON object. Do not include any explanatory text, comments, or markdown like \`\`\`json.

      **Analyze the following data:**
      
      **HubSpot Properties (System Data):**
      ${JSON.stringify(hubspotResult.data, null, 2)}

      **Sales Note Text (Human Input):**
      ---
      ${noteText || "No sales note was found for this deal."}
      ---

      **Generate the JSON object using this exact format:**
      {
        "learnerName": "string",
        "parentName": "string",
        "parentEmail": "string",
        "parentContact": "string",
        "planName": "string",
        "currency": "string (e.g., EUR, GBP, USD)",
        "sessionsPerWeek": "string (e.g., 1 Session/week)",
        "discount": "float",
        "subscriptionTenure": "integer (months)",
        "startDate": "YYYY-MM-DD",
        "endDate": "YYYY-MM-DD",
        "paymentType": "'Upfront' or 'Installment'",
        "totalAmount": "float",
        "finalPayable": "float",
        "installments": [
          { "installmentNumber": 1, "amount": "float", "dueDate": "YYYY-MM-DD", "isPaid": "boolean" }
        ],
        "aiReasoning": "A brief explanation of how you derived the plan, especially for complex cases like uneven payments or implied discounts."
      }
    `;

    // 3. Call the Gemini API
    const GOOGLE_API_KEY = PropertiesService.getScriptProperties().getProperty('GOOGLE_GENERATIVE_AI_KEY');
    if (!GOOGLE_API_KEY) {
        throw new Error("AI Agent is offline: API key is not configured on the server.");
    }
    
    const model = 'gemini-2.5-flash'; 
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GOOGLE_API_KEY}`;

    
    const payload = { contents: [{ parts: [{ text: prompt }] }] };
    const response = callGenerativeAIWithRetry(endpoint, payload); // Use the new retry wrapper

    const responseBody = response.getContentText();
    if (response.getResponseCode() !== 200) {
        Logger.log(`AI API Error Response: ${responseBody}`);
        throw new Error(`The AI Agent returned an error. Please check the logs for details.`);
    }
    
    const jsonResponse = JSON.parse(responseBody);
    const textPart = jsonResponse?.candidates?.[0]?.content?.parts?.[0]?.text;

    if (!textPart) {
      throw new Error("The AI returned an empty response. It might be unable to process this specific deal's data.");
    }

    // Clean the response to ensure it's valid JSON
    const cleanedText = textPart.replace(/```json/g, '').replace(/```/g, '').trim();
    const aiPlan = JSON.parse(cleanedText);
    
    // 4. Final Validation and Return
    if (!aiPlan.planName || !aiPlan.hasOwnProperty('finalPayable')) {
        throw new Error("The AI returned an incomplete or invalid plan. Please try again or enter the details manually.");
    }
    
    return { success: true, data: aiPlan };

  } catch (error) {
    Logger.log(`[FATAL] Error in generateSmartInvoicePlan for JLID ${jlid}: ${error.message}\nStack: ${error.stack}`);
    return { success: false, message: error.message };
  }
}

/**
 * Runs the daily AI-powered audit and sends a report if discrepancies are found.
 * Can be run by a time-based trigger or manually from the UI.
 * @param {boolean} manualTrigger - If true, indicates the run was started from the UI.
 * @returns {object} A result object for the client-side call.
 */
function runDailyAIAudit(manualTrigger = false) {
  Logger.log('Starting Daily AI Onboarding Audit...');
  
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const toDate = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const fromDate = toDate;
  
  try {
    const deals = fetchDealsByOnboardingCompletionDate(fromDate, toDate);
    if (deals.length === 0) {
      const message = `No onboardings found for ${fromDate}. No audit report generated.`;
      Logger.log(message);
      if (manualTrigger) return { success: true, message: message };
      return;
    }

    let mismatches = [];

    deals.forEach(deal => {
      try {
        const props = deal.properties;
        const noteText = fetchLatestSalesNoteForDeal(deal.id);
        const noteData = parseSalesNote(noteText);
        const calendarResult = verifySubscriptionWithCalendar(
          props.jetlearner_id, 
          props.module_start_date, 
          props.module_end_date, 
          parseInt(props.total_classes_committed_through_learner_s_journey)
        );

        const aiAnalysis = getAIAgentAnalysis(props, noteData, calendarResult, noteText);

        if (aiAnalysis.isMismatch) {
          mismatches.push({
            learnerName: props.dealname,
            jlid: props.jetlearner_id,
            dealId: deal.id,
            summary: aiAnalysis.summary
          });
        }
      } catch(e) {
        Logger.log(`Error processing deal ${deal.id} in daily audit: ${e.message}`);
        mismatches.push({
          learnerName: deal.properties.dealname || 'Unknown',
          jlid: deal.properties.jetlearner_id || 'N/A',
          dealId: deal.id,
          summary: `Audit failed due to a system error: ${e.message}`
        });
      }
    });

    if (mismatches.length > 0) {
      Logger.log(`Found ${mismatches.length} mismatches. Sending report.`);
      
      let htmlBody = `
        <h2>Daily Onboarding Audit Report - ${fromDate}</h2>
        <p>The automated AI agent found the following ${mismatches.length} potential discrepancies in yesterday's onboardings. Please review and take action.</p>
      `;

      mismatches.forEach(m => {
        const dealUrl = `https://app.hubspot.com/contacts/19972323/deal/${m.dealId}`;
        htmlBody += `
          <div style="border: 1px solid #ccc; border-left: 4px solid #f44336; padding: 10px; margin-bottom: 15px;">
            <h3><a href="${dealUrl}">${m.learnerName} (${m.jlid || 'No JLID'})</a></h3>
            <p><strong>AI Agent Findings:</strong></p>
            <ul>${m.summary.replace(/-/g, '<li>')}</ul>
          </div>
        `;
      });
      
      MailApp.sendEmail({
        to: CONFIG.EMAIL.AUDIT_REPORT_RECIPIENTS,
        subject: `[ACTION REQUIRED] Daily Onboarding Audit Found ${mismatches.length} Mismatches`,
        htmlBody: htmlBody,
        name: `${CONFIG.EMAIL.FROM_NAME} (AI Agent)`,
        from: CONFIG.EMAIL.FROM
      });
      if (manualTrigger) return { success: true, message: `Audit complete. Found ${mismatches.length} mismatches. Report sent to ${CONFIG.EMAIL.AUDIT_REPORT_RECIPIENTS}.`};
    } else {
      const message = `Audit complete for ${fromDate}. All ${deals.length} onboardings are compliant.`;
      Logger.log(message);
      if (manualTrigger) return { success: true, message: message };
    }

  } catch (error) {
    Logger.log(`FATAL ERROR in Daily AI Audit: ${error.message}`);
    MailApp.sendEmail({
        to: CONFIG.EMAIL.AUDIT_REPORT_RECIPIENTS,
        subject: `[ERROR] Daily Onboarding AI Audit Failed to Run`,
        body: `The automated daily audit encountered a critical error and could not complete. Please check the script logs.\n\nError: ${error.message}`,
        name: `${CONFIG.EMAIL.FROM_NAME} (System Alert)`,
        from: CONFIG.EMAIL.FROM
      });
    if (manualTrigger) return { success: false, message: `Audit failed to run: ${error.message}` };
  }
}


/**
 * Creates a time-driven trigger to run the daily AI audit.
 * Should be run once manually by an admin to set up the automation.
 */
function createDailyAuditTrigger() {
  // Delete existing triggers to prevent duplicates
  const existingTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of existingTriggers) {
    if (trigger.getHandlerFunction() === 'runDailyAIAudit') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Deleted existing daily audit trigger.');
    }
  }

  // Create a new trigger to run every day between 8 AM and 9 AM
  ScriptApp.newTrigger('runDailyAIAudit')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  
  const message = 'Successfully created trigger for the Daily AI Audit. It will run every morning.';
  Logger.log(message);
  return { success: true, message: message };
}

// =============================================
// UTILITY FUNCTIONS (FINAL FIXES)
// =============================================

/**
 * NEW HELPER FUNCTION
 * Looks up the primary email address for a CLS Manager by their name.
 * Assumes CLS Managers exist in the 'Teacher Data' sheet with their name in Column H and email in Column J.
 * This is the refined version ensuring correct column lookup as per the "FINAL FIX" requirements.
 * @param {string} managerName The name of the CLS Manager to find.
 * @returns {string|null} The email address from Column J, or null if not found or invalid.
 */
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


/**
 * Reliably removes a trailing 'C' or 'M' from a JLID to create a clean Zoom link ID.
 * @param {string} jlid The JetLearn ID.
 * @returns {string} The cleaned JLID.
 */
function cleanJlidForZoom(jlid) {
  if (!jlid || typeof jlid !== 'string') return '';
  // This regex now removes a trailing 'C' or 'M' followed by any number of digits.
  return jlid.trim().toUpperCase().replace(/(C|M)\d*$/, '');
}


/**
 * Gets a unique list of TP Managers from the Teacher Data sheet.
 * @returns {Array<string>} An array of TP Manager names.
 */
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

// =============================================
// AUTHENTICATION & USER MANAGEMENT
// =============================================

function authenticateUser(username, password) {
  Logger.log('authenticateUser called for username: ' + username);

  if (!username || !password) {
    Logger.log('Missing credentials');
    return { success: false, role: ROLES.GUEST, message: 'Missing credentials' };
  }

  try {
    const userProfiles = getUserProfiles(); 
    const user = userProfiles.find(u => u.username === username);

    if (!user) {
      Logger.log('User not found: ' + username);
      return { success: false, role: ROLES.GUEST, message: 'Invalid credentials' };
    }

    if (user.password !== password) {
      Logger.log('Invalid password for user: ' + username);
      logUserActivity(username, 'Failed Login', 'Invalid credentials');
      return { success: false, role: ROLES.GUEST, message: 'Invalid credentials' };
    }

    if (!user.isActive) {
      Logger.log('Inactive user attempted login: ' + username);
      return { success: false, role: ROLES.GUEST, message: 'Account inactive' };
    }

    updateUserLastLogin(username);
    logUserActivity(username, 'Successful Login', 'User logged in');

    Logger.log('Authentication successful for user: ' + username + ', role: ' + user.role);
    return {
      success: true,
      role: user.role,
      username: username,
      permissions: PERMISSIONS[user.role] || []
    };
  } catch (error) {
    Logger.log('Error in authenticateUser: ' + error.message);
    return { success: false, role: ROLES.GUEST, message: 'Authentication error' };
  }
}

function verifyUserSession(username) {
  Logger.log('verifyUserSession called for: ' + username);
  try {
      const userProfiles = getUserProfiles(); 
      const user = userProfiles.find(u => u.username === username);

      if (!user || !user.isActive) {
        Logger.log('Session verification failed for user: ' + username);
        return { success: false, message: 'Invalid or inactive session.' };
      }

      Logger.log('Session verification successful for user: ' + username);
      return {
        success: true,
        role: user.role,
        username: user.username,
        permissions: PERMISSIONS[user.role] || []
      };
  } catch (error) {
    Logger.log('Error in verifyUserSession: ' + error.message);
    return { success: false, message: 'Session verification error.' };
  }
}

function getUserProfiles() {
  try {
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.USER_PROFILES);

    if (sheetData.length <= 1) { 
      createDefaultUsers(); 
      return getUserProfiles();
    }

    const headers = sheetData[0];
    return sheetData.slice(1).map(row => {
      const user = {};
      headers.forEach((header, i) => {
        const key = header.toLowerCase().replace(/\s/g, ''); 
        user[key] = row[i];
      });
      return {
        username: user.username,
        password: user.password,
        role: user.role,
        email: user.email,
        isActive: user.isactive,
        lastLogin: user.lastlogin,
        createdDate: user.createddate
      };
    });
  } catch (error) {
    Logger.log('Error getting user profiles: ' + error.message);
    return [];
  }
}

function createDefaultUsers() {
  try {
    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.USER_PROFILES);

    if (sheet.getLastRow() === 0 || sheet.getRange('A1').isBlank()) { 
      sheet.getRange(1, 1, 9, 9).setValues([
        ['Username', 'Password', 'Role', 'Email', 'IsActive', 'LastLogin', 'CreatedDate', 'ResetToken', 'TokenExpiry']
      ]);
    }

    const defaultUsers = [
      ['Admin', 'JetLearn2025$', ROLES.ADMIN, 'admin@jet-learn.com', true, '', new Date(), '', ''],
      ['Ops_team', 'Opsteam@2025$', ROLES.USER, 'ops@jet-learn.com', true, '', new Date(), '', '']
    ];

    sheet.getRange(sheet.getLastRow() + 1, 1, defaultUsers.length, 9).setValues(defaultUsers);
    Logger.log('Default users created');

    delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.USER_PROFILES}`];

  } catch (error) {
    Logger.log('Error creating default users: ' + error.message);
  }
}

function updateUserLastLogin(username) {
  try {
    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.USER_PROFILES);
    const data = _getCachedSheetData(CONFIG.SHEETS.USER_PROFILES); 
    const headers = data[0].map(h => h.toLowerCase().replace(/\s/g, ''));
    const usernameColIndex = headers.indexOf('username');
    const lastLoginColIndex = headers.indexOf('lastlogin');

    if (usernameColIndex === -1 || lastLoginColIndex === -1) {
      Logger.log('User Profiles sheet missing Username or LastLogin column.');
      return;
    }

    for (let i = 1; i < data.length; i++) { 
      if (data[i][usernameColIndex] === username) {
        sheet.getRange(i + 1, lastLoginColIndex + 1).setValue(new Date()); 
        delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.USER_PROFILES}`];
        break;
      }
    }
  } catch (error) {
    Logger.log('Error updating last login: ' + error.message);
  }
}

function requestPasswordReset(email) {
  Logger.log('requestPasswordReset called for email: ' + email);
  try {
    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.USER_PROFILES);
    const data = _getCachedSheetData(CONFIG.SHEETS.USER_PROFILES); 

    if (data.length < 1) {
        Logger.log("User Profiles sheet is empty or does not have headers. Cannot process password reset.");
        return { success: false, message: "Server configuration error: User profiles not set up." };
    }

    const headers = data[0].map(h => h.toLowerCase().replace(/\s/g, '')); 
    const emailCol = headers.indexOf('email');
    const usernameCol = headers.indexOf('username');
    const tokenCol = headers.indexOf('resettoken');
    const expiryCol = headers.indexOf('tokenexpiry');

    if (emailCol === -1 || tokenCol === -1 || expiryCol === -1 || usernameCol === -1) {
      Logger.log("User Profiles sheet missing one or more required columns for password reset (Email, ResetToken, TokenExpiry, Username).");
      return { success: false, message: "Server configuration error: Required columns for password reset not found. Please check 'User Profiles' sheet headers." };
    }

    let userRowDataIndex = -1; 
    for (let i = 1; i < data.length; i++) { 
        if (data[i][emailCol] && String(data[i][emailCol]).trim().toLowerCase() === email.trim().toLowerCase()) {
            userRowDataIndex = i; 
            break;
        }
    }

    if (userRowDataIndex === -1) {
      Logger.log('Email address not found in user profiles: ' + email);
      return { success: false, message: 'Email address not found.' };
    }

    const userSheetRowIndex = userRowDataIndex + 1;

    const token = Utilities.getUuid();
    const expiry = new Date(new Date().getTime() + 60 * 60 * 1000);

    sheet.getRange(userSheetRowIndex, tokenCol + 1).setValue(token);
    sheet.getRange(userSheetRowIndex, expiryCol + 1).setValue(expiry);
    Logger.log(`Generated token for ${email}: ${token}, expires: ${expiry.toLocaleString()}`);

    const webAppUrl = ScriptApp.getService().getUrl();
    const resetUrl = `${webAppUrl}?resetToken=${token}`;
    Logger.log('Generated reset URL: ' + resetUrl);

    const username = data[userRowDataIndex][usernameCol];

    const emailBody = `
      <p>Hello ${username},</p>
      <p>A password reset was requested for your JetLearn System account. Please click the link below to reset your password. This link is valid for 1 hour.</p>
      <p><a href="${resetUrl}" style="display: inline-block; padding: 10px 15px; background-color: #4a3c8a; color: white; text-decoration: none; border-radius: 5px;">Reset Password</a></p>
      <p>If you did not request this, please ignore this email.</p>
      <p>Thanks,<br>The JetLearn Team</p>
    `;

    MailApp.sendEmail({
      to: email,
      subject: 'JetLearn System - Password Reset Request',
      htmlBody: emailBody,
      name: CONFIG.EMAIL.FROM_NAME,
      from: CONFIG.EMAIL.FROM 
    });
    Logger.log('Password reset email sent to: ' + email);

    delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.USER_PROFILES}`];

    return { success: true, message: 'A password reset link has been sent to your email.' };
  } catch (error) {
    Logger.log('Error in requestPasswordReset: ' + error.message);
    if (error.stack) {
        Logger.log('Stack trace: ' + error.stack);
    }
    return { success: false, message: 'An error occurred. Please try again later.' };
  }
}

/**
 * Resets a user's password using a secure, single-use token.
 * @param {string} token The password reset token.
 * @param {string} newPassword The new password for the user.
 * @returns {object} A result object with success status and message.
 */
function resetPassword(token, newPassword) {
  Logger.log('resetPassword called with token: ' + token);
  try {
    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.USER_PROFILES);
    const data = _getCachedSheetData(CONFIG.SHEETS.USER_PROFILES); 

    if (data.length < 1) {
        Logger.log("User Profiles sheet is empty or does not have headers. Cannot process password reset.");
        return { success: false, message: "Server configuration error: User profiles not set up." };
    }

    const headers = data[0].map(h => h.toLowerCase().replace(/\s/g, ''));
    const tokenCol = headers.indexOf('resettoken');
    const expiryCol = headers.indexOf('tokenexpiry');
    const passwordCol = headers.indexOf('password');

    if (tokenCol === -1 || expiryCol === -1 || passwordCol === -1) {
      Logger.log("User Profiles sheet missing one or more required columns for password reset (ResetToken, TokenExpiry, Password).");
      return { success: false, message: "Server configuration error: Required columns for password reset not found. Please check 'User Profiles' sheet headers." };
    }

    let userRowDataIndex = -1; 
    for (let i = 1; i < data.length; i++) { 
      if (data[i][tokenCol] && String(data[i][tokenCol]).trim() === String(token).trim()) {
        userRowDataIndex = i; 
        break;
      }
    }

    if (userRowDataIndex === -1) {
      Logger.log('Invalid or non-existent reset token provided: ' + token);
      return { success: false, message: 'Invalid or expired reset token.' };
    }

    const userSheetRowIndex = userRowDataIndex + 1;

    const expiryDate = new Date(data[userRowDataIndex][expiryCol]); 
    if (isNaN(expiryDate.getTime())) {
        Logger.log(`Invalid expiry date for token ${token}: ${data[userRowDataIndex][expiryCol]}`);
        return { success: false, message: 'Invalid token expiry date. Please request a new link.' };
    }

    if (expiryDate < new Date()) {
      Logger.log(`Token ${token} has expired.`);
      sheet.getRange(userSheetRowIndex, tokenCol + 1).setValue('');
      sheet.getRange(userSheetRowIndex, expiryCol + 1).setValue('');
      delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.USER_PROFILES}`];
      return { success: false, message: 'Password reset token has expired.' };
    }

    if (!newPassword || newPassword.length < 6) { 
        return { success: false, message: 'New password must be at least 6 characters long.' };
    }

    sheet.getRange(userSheetRowIndex, passwordCol + 1).setValue(newPassword);
    Logger.log(`Password updated for user at row ${userSheetRowIndex}.`);

    sheet.getRange(userSheetRowIndex, tokenCol + 1).setValue('');
    sheet.getRange(userSheetRowIndex, expiryCol + 1).setValue('');
    Logger.log(`Token ${token} invalidated after use.`);

    delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.USER_PROFILES}`];

    return { success: true, message: 'Your password has been reset successfully.' };
  } catch (error) {
    Logger.log('Error in resetPassword: ' + error.message);
    if (error.stack) {
        Logger.log('Stack trace: ' + error.stack);
    }
    return { success: false, message: 'An error occurred while resetting the password.' };
  }
}

// =============================================
// DASHBOARD & STATISTICS
// =============================================

function getDashboardStatistics() {
  Logger.log('getDashboardStatistics called');

  const stats = {
    migrations: { total: 0, successful: 0, failed: 0, today: 0, thisWeek: 0, thisMonth: 0, successRate: 0 },
    onboardings: { total: 0, today: 0, thisWeek: 0, thisMonth: 0 },
    recentActivities: []
  };

  try {
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const thisWeekStart = new Date(now.getTime() - (now.getDay()) * 24 * 60 * 60 * 1000); 
    thisWeekStart.setHours(0,0,0,0);
    const thisMonthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    
    let combinedActivities = [];

    // --- 1. Process Migrations & Onboardings from Audit Log ---
    try {
      const rawSheetData = _getCachedSheetData(CONFIG.SHEETS.AUDIT_LOG);
      const auditData = rawSheetData.length > 1 ? rawSheetData.slice(1) : [];

      auditData.forEach(row => {
        const timestamp = parseSheetDate(row[0]);
        if (!timestamp || isNaN(timestamp.getTime())) return;

        const action = String(row[1]);
        const status = String(row[7]);
        const learner = row[3];
        const newTeacher = row[5];
        const oldTeacher = row[4];

        // --- LOGIC UPDATE: Combine all Onboarding variations ---
        const isOnboarding = action.includes('New Learner Onboarded') || 
                             action.includes('Email Sent (Onboarding');

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
        // --- Check for the expanded Onboarding definition ---
        else if (isOnboarding) {
          if (status === 'Success' || status === 'Partial Success') {
            stats.onboardings.total++;
            if (timestamp >= today) stats.onboardings.today++;
            if (timestamp >= thisWeekStart) stats.onboardings.thisWeek++;
            if (timestamp >= thisMonthStart) stats.onboardings.thisMonth++;
          }

          combinedActivities.push({
            timestamp: timestamp,
            type: 'onboarding',
            title: 'New Learner Onboarded',
            description: `${learner} ${newTeacher ? 'assigned to ' + newTeacher : 'onboarded'}`,
            // FIX APPLIED HERE: Pass actual status instead of forcing 'Warning'
            status: status 
          });
        }
      });
    } catch (e) {
      Logger.log('Error processing Audit Log stats: ' + e.message);
    }

    if (stats.migrations.total > 0) {
      stats.migrations.successRate = Math.round((stats.migrations.successful / stats.migrations.total) * 100);
    }
    
    // --- 2. Process Email Activities ---
    try {
      const emailLogs = _getCachedSheetData(CONFIG.SHEETS.EMAIL_LOGS);
      if (emailLogs.length > 1) {
          const headers = emailLogs[0];
          const trackingIdCol = headers.findIndex(h => h.trim() === 'Tracking ID'); 
          const statusCol = headers.findIndex(h => h.trim() === 'Status');
          const recipientCol = headers.findIndex(h => h.trim() === 'Recipient');
          const subjectCol = headers.findIndex(h => h.trim() === 'Subject');
          const sentAtCol = headers.findIndex(h => h.trim() === 'Sent At');
          const openedAtCol = headers.findIndex(h => h.trim() === 'Opened At');
          const repliedAtCol = headers.findIndex(h => h.trim() === 'Replied At');

          const recentEmailRows = emailLogs.slice(1); 

          recentEmailRows.forEach(row => {
              if (statusCol === -1 || trackingIdCol === -1) return;

              const status = row[statusCol];
              if (status === 'Sent' || status === 'Opened' || status === 'Replied') {
                  let timestamp;
                  let displayTitle = `Email ${status}`;
                  
                  if (status === 'Opened' && openedAtCol > -1 && row[openedAtCol]) {
                      timestamp = parseSheetDate(row[openedAtCol]);
                  } else if (status === 'Replied' && repliedAtCol > -1 && row[repliedAtCol]) {
                      timestamp = parseSheetDate(row[repliedAtCol]);
                  } else if (sentAtCol > -1) {
                      timestamp = parseSheetDate(row[sentAtCol]);
                  }
                  
                  if (!timestamp) return;

                  combinedActivities.push({
                      timestamp: timestamp,
                      type: 'email',
                      title: displayTitle,
                      description: `To: ${recipientCol > -1 ? row[recipientCol] : 'Unknown'} | Subject: "${subjectCol > -1 ? row[subjectCol] : 'No Subject'}"`,
                      status: (status === 'Replied' ? 'Info' : (status === 'Sent' ? 'Skipped' : 'Success')),
                      trackingId: row[trackingIdCol]
                  });
              }
          });
      }
    } catch (e) {
        Logger.log('Error processing Email Logs stats: ' + e.message);
    }
    
    // Sort descending
    combinedActivities.sort((a, b) => b.timestamp - a.timestamp);
    
    // --- DATE FIX IS HERE ---
    // Convert Date objects to ISO Strings so the browser (Moment.js) can read them perfectly
    stats.recentActivities = combinedActivities.slice(0, 7).map(activity => {
        let timeStr = null; // Send ISO string or null
        try {
            if (activity.timestamp instanceof Date && !isNaN(activity.timestamp)) {
                timeStr = activity.timestamp.toISOString(); 
            }
        } catch (e) { Logger.log('Date conversion error: ' + e.message); }

        return {
            ...activity,
            timestamp: timeStr 
        };
    });

    return stats;

  } catch (error) {
    Logger.log('CRITICAL Error in getDashboardStatistics: ' + error.message);
    return stats;
  }
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

// =============================================
// TEACHER MANAGEMENT FUNCTIONS
// =============================================

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




// =============================================
// COURSE MANAGEMENT FUNCTIONS
// =============================================

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

// =============================================
// ### MODIFIED ### REPORT GENERATION FUNCTIONS V3.0
// =============================================

/**
 * =================================================================
 * ENHANCED REPORTING ENGINE V3.0
 * =================================================================
 * This function generates a comprehensive, analyst-grade migration report.
 * It calculates KPIs, MoM changes, and drills down into teacher/course impact.
 *
 * @param {object} params - Contains fromDate and toDate for the report period.
 * @returns {object} A rich JSON object containing all data for the new dashboard.
 */
function getEnhancedMigrationReport(params) {
    Logger.log(`getEnhancedMigrationReport (V3) called with: ${JSON.stringify(params)}`);
    try {
        const auditData = getAuditLog({ limit: 20000 }).data; 
        const migrationLogs = auditData.filter(row => String(row[1]).includes('Migration'));
        if (migrationLogs.length === 0) {
            return { success: true, message: "No migration data found.", data: null };
        }

        const headers = _getCachedSheetData(CONFIG.SHEETS.AUDIT_LOG)[0] || [];
        const timestampCol = headers.indexOf('Timestamp');
        const reasonCol = headers.indexOf('Reason for Migration');
        const intervenedCol = headers.indexOf('Intervened By');
        const oldTeacherCol = headers.indexOf('Old Teacher');
        const newTeacherCol = headers.indexOf('New Teacher');
        const courseCol = headers.indexOf('Course');

        const toDate = new Date(params.toDate);
        const fromDate = new Date(params.fromDate);
        const periodDuration = toDate.getTime() - fromDate.getTime();
        const prevToDate = new Date(fromDate.getTime() - (24 * 60 * 60 * 1000)); 
        const prevFromDate = new Date(prevToDate.getTime() - periodDuration);

        const processPeriod = (start, end) => {
            const periodData = {
                totalMigrations: 0,
                clsInvolvement: 0,
                tpInvolvement: 0,
                opsInvolvement: 0,
                reasonBreakdown: {},
                teamInvolvementByReason: {},
                teacherMigrationsFrom: {},
                teacherMigrationsTo: {},
                courseMigrations: {}
            };

            migrationLogs.forEach(row => {
                const timestamp = parseSheetDate(row[timestampCol]);
                if (!timestamp || timestamp < start || timestamp > end) return;

                periodData.totalMigrations++;
                const reason = String(row[reasonCol] || 'Unknown').trim();
                const teams = String(row[intervenedCol] || '').toLowerCase();

                if (teams.includes('cls')) periodData.clsInvolvement++;
                if (teams.includes('tp')) periodData.tpInvolvement++;
                if (teams.includes('ops')) periodData.opsInvolvement++;

                periodData.reasonBreakdown[reason] = (periodData.reasonBreakdown[reason] || 0) + 1;
                
                if (!periodData.teamInvolvementByReason[reason]) {
                    periodData.teamInvolvementByReason[reason] = { cls: 0, tp: 0, ops: 0 };
                }
                if (teams.includes('cls')) periodData.teamInvolvementByReason[reason].cls++;
                if (teams.includes('tp')) periodData.teamInvolvementByReason[reason].tp++;
                if (teams.includes('ops')) periodData.teamInvolvementByReason[reason].ops++;


                const oldTeacher = String(row[oldTeacherCol] || '').trim();
                const newTeacher = String(row[newTeacherCol] || '').trim();
                if (oldTeacher) periodData.teacherMigrationsFrom[oldTeacher] = (periodData.teacherMigrationsFrom[oldTeacher] || 0) + 1;
                if (newTeacher) periodData.teacherMigrationsTo[newTeacher] = (periodData.teacherMigrationsTo[newTeacher] || 0) + 1;

                const course = String(row[courseCol] || 'Unknown').trim();
                if (!periodData.courseMigrations[course]) {
                    periodData.courseMigrations[course] = { count: 0, reasons: {} };
                }
                periodData.courseMigrations[course].count++;
                periodData.courseMigrations[course].reasons[reason] = (periodData.courseMigrations[course].reasons[reason] || 0) + 1;
            });
            return periodData;
        };

        const currentPeriod = processPeriod(fromDate, toDate);
        const previousPeriod = processPeriod(prevFromDate, prevToDate);

        const calculateChange = (current, previous) => {
            if (previous === 0) return current > 0 ? 10000 : 0; 
            return ((current - previous) / previous) * 100;
        };
        
        const daysInPeriod = (toDate.getTime() - fromDate.getTime()) / (1000 * 3600 * 24) + 1;
        const daysInPrevPeriod = (prevToDate.getTime() - prevFromDate.getTime()) / (1000 * 3600 * 24) + 1;

        const currentAvg = daysInPeriod > 0 ? (currentPeriod.totalMigrations / daysInPeriod) : 0;
        const prevAvg = daysInPrevPeriod > 0 ? (previousPeriod.totalMigrations / daysInPrevPeriod) : 0;

        const kpis = {
            totalMigrations: { current: currentPeriod.totalMigrations, previous: previousPeriod.totalMigrations, change: calculateChange(currentPeriod.totalMigrations, previousPeriod.totalMigrations) },
            avgMigrationsPerDay: { current: currentAvg, previous: prevAvg, change: calculateChange(currentAvg, prevAvg) },
            clsRate: { current: (currentPeriod.totalMigrations > 0 ? (currentPeriod.clsInvolvement / currentPeriod.totalMigrations) * 100 : 0), previous: (previousPeriod.totalMigrations > 0 ? (previousPeriod.clsInvolvement / previousPeriod.totalMigrations) * 100 : 0) },
            tpRate: { current: (currentPeriod.totalMigrations > 0 ? (currentPeriod.tpInvolvement / currentPeriod.totalMigrations) * 100 : 0), previous: (previousPeriod.totalMigrations > 0 ? (previousPeriod.tpInvolvement / previousPeriod.totalMigrations) * 100 : 0) },
        };
        
        kpis.clsRate.change = kpis.clsRate.current - kpis.clsRate.previous;
        kpis.tpRate.change = kpis.tpRate.current - kpis.tpRate.previous;


        const allTeachers = new Set([...Object.keys(currentPeriod.teacherMigrationsFrom), ...Object.keys(currentPeriod.teacherMigrationsTo)]);
        const teacherImpact = Array.from(allTeachers).map(name => {
            const from = currentPeriod.teacherMigrationsFrom[name] || 0;
            const to = currentPeriod.teacherMigrationsTo[name] || 0;
            return { name, from, to, netFlow: to - from };
        }).sort((a, b) => a.netFlow - b.netFlow);

        const courseImpact = Object.entries(currentPeriod.courseMigrations).map(([name, data]) => {
            const topReason = Object.keys(data.reasons).length > 0 ? Object.entries(data.reasons).sort((a, b) => b[1] - a[1])[0][0] : 'N/A';
            return { name, count: data.count, topReason };
        }).sort((a, b) => b.count - a.count);

        const reportData = {
            kpis,
            reasonBreakdown: Object.entries(currentPeriod.reasonBreakdown).map(([name, count]) => ({ name, count })).sort((a, b) => b.count - a.count),
            teamInvolvementByReason: currentPeriod.teamInvolvementByReason,
            teacherImpact,
            courseImpact
        };

        const aiInsights = getEnhancedAIInsights(reportData, fromDate.toLocaleString('default', { month: 'long' }));

        return { success: true, data: reportData, aiInsights: aiInsights };
    } catch (error) {
        Logger.log('Error in getEnhancedMigrationReport: ' + error.message + ' Stack: ' + error.stack);
        return { success: false, message: 'Error generating enhanced report: ' + error.message, data: null };
    }
}


/**
 * ### NEW ### Generates a multi-layered AI analysis based on the enhanced report data.
 * @param {object} reportData The rich data object from getEnhancedMigrationReport.
 * @param {string} monthName The name of the current month being analyzed.
 * @returns {object} An object containing executive, rootCause, and impact summaries.
 */
function getEnhancedAIInsights(reportData, monthName) {
    Logger.log(`Getting Enhanced AI insights for ${monthName}.`);
    const GOOGLE_API_KEY = PropertiesService.getScriptProperties().getProperty('GOOGLE_GENERATIVE_AI_KEY');
    if (!GOOGLE_API_KEY) {
        return { executive: "AI insights unavailable: API key not configured.", rootCause: "", impact: "" };
    }

    const model = 'gemini-2.5-flash';
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/${model}:generateContent?key=${GOOGLE_API_KEY}`;
    
    const prompt = `
      You are a top-tier Senior Operations Analyst at JetLearn. Your task is to provide a three-part analysis of the monthly learner migration data.
      Your response MUST be a single, valid JSON object with three keys: "executive", "rootCause", and "impact". Each key's value should be a concise, analytical string using markdown for bolding. Do not include any text outside this JSON object.

      Analyze the following data for the month of ${monthName}:
      ${JSON.stringify(reportData, null, 2)}

      Based on this data, generate the following three summaries:

      1.  **executive**: A high-level summary for leadership. Focus on the most significant KPI change (e.g., total migrations percentage change) and its primary business implication. Mention the most impactful team shift. Be direct and start with the most critical finding.

      2.  **rootCause**: An analysis for operations managers. Identify the top migration driver from the 'reasonBreakdown'. If it's 'Unknown', highlight the data integrity issue. For the top *known* reason, explain which team is handling it and what that implies (e.g., "**Attrition** is a high-touch issue handled by **CLS**").

      3.  **impact**: An actionable analysis for team leads. Identify the teacher with the most negative 'netFlow' and recommend a specific action (e.g., "warrants a performance review or support session"). Also, identify the course with the highest migration 'count' and recommend an action (e.g., "recommend a review of the curriculum or teacher cohort for this course").
    `;

    try {
        const payload = { contents: [{ parts: [{ text: prompt }] }] };
        const response = callGenerativeAIWithRetry(endpoint, payload); // Use the new retry wrapper
        const responseCode = response.getResponseCode();
        const responseBody = response.getContentText();

        if (responseCode === 200) {
            const jsonResponse = JSON.parse(responseBody);
            let textPart = jsonResponse?.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
            textPart = textPart.replace(/```json/g, '').replace(/```/g, '').trim();
            Logger.log('AI Insights Generated: ' + textPart);
            return JSON.parse(textPart);
        }
        
        Logger.log(`AI API Error (${responseCode}): ${responseBody}`);
        return { executive: "AI insights could not be generated due to an API error.", rootCause: "", impact: "" };
    } catch (error) {
        Logger.log('Error calling or parsing AI API: ' + error.message);
        return { executive: `AI insights are unavailable due to a connection or parsing error: ${error.message}`, rootCause: "", impact: "" };
    }
}

// =============================================
// ### NEW & UPDATED ### REPORT EXPORT & EMAIL FUNCTIONS
// =============================================

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

/**
 * ### NEW ### Sends the generated report to designated recipients.
 * @param {object} params - Contains fromDate and toDate.
 * @returns {object} A result object with success status and message.
 */
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

// =============================================
// USER MANAGEMENT FUNCTIONS (ADMIN)
// =============================================

function getActiveUsers() {
  Logger.log('getActiveUsers called');

  try {
    const users = getUserProfiles(); 
    return users.map(user => ({
      username: user.username,
      role: user.role,
      email: user.email,
      lastLogin: user.lastLogin ? new Date(user.lastLogin).toLocaleString() : 'Never',
      createdDate: user.createdDate ? new Date(user.createdDate).toLocaleDateString() : 'N/A',
      isActive: user.isActive
    }));
  } catch (error) {
    Logger.log('Error getting active users: ' + error.message);
    return [];
  }
}

function addNewUser(userData) {
  Logger.log('addNewUser called for: ' + userData.username);

  try {
    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.USER_PROFILES);

    const existingUsers = getUserProfiles(); 
    if (existingUsers.some(u => u.username.toLowerCase() === userData.username.toLowerCase())) {
      return { success: false, message: 'Username already exists' };
    }

    if (!isValidEmail(userData.email)) {
      return { success: false, message: 'Invalid email address' };
    }

    if (!Object.values(ROLES).includes(userData.role)) {
      return { success: false, message: 'Invalid user role' };
    }

    sheet.appendRow([
      userData.username,
      userData.password,
      userData.role,
      userData.email,
      true, 
      '',   
      new Date(), 
      '',   
      ''    
    ]);

    logAction('User Added', '', '', '', '', '', 'Success', `New ${userData.role} user added: ${userData.username}`);

    delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.USER_PROFILES}`];

    return { success: true, message: 'User added successfully' };
  } catch (error) {
    Logger.log('Error adding new user: ' + error.message);
    return { success: false, message: 'Error adding user: ' + error.message };
  }
}

/**
 * [ADMIN-ONLY] Updates a user's details, including role, email, password, and active status.
 * @param {object} userData - The user data to update. Requires 'username'. Can include 'role', 'email', 'password', 'isActive'.
 * @param {object} currentUser - The user object of the person performing the action. Requires 'username' and 'role'.
 * @returns {object} A result object with success status and message.
 */
function updateUser(userData, currentUser) {
  Logger.log(`updateUser called for '${userData.username}' by user '${currentUser.username}'`);

  try {
    if (!currentUser || currentUser.role !== ROLES.ADMIN || !hasPermission(currentUser.role, 'manage_users')) {
      logUserActivity(currentUser.username, 'Update User Failed', `Permission denied to update ${userData.username}.`);
      return { success: false, message: 'Permission denied. Only Admins can manage users.' };
    }

    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.USER_PROFILES);
    const data = _getCachedSheetData(CONFIG.SHEETS.USER_PROFILES); 

    let rowIndex = -1;
    for (let i = 0; i < data.length; i++) { 
      if (data[i][0] === userData.username) {
        rowIndex = i + 1; 
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: `User '${userData.username}' not found.` };
    }

    const headers = data[0].map(h => h.toLowerCase().replace(/\s/g, ''));
    const changes = [];

    if (userData.password) {
      const colIndex = headers.indexOf('password');
      if (colIndex > -1) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(userData.password);
        changes.push('password');
      }
    }

    if (userData.role) {
      if (!Object.values(ROLES).includes(userData.role)) {
        return { success: false, message: `Invalid role: ${userData.role}` };
      }
      const colIndex = headers.indexOf('role');
      if (colIndex > -1) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(userData.role);
        changes.push(`role to '${userData.role}'`);
      }
    }

    if (userData.email) {
       if (!isValidEmail(userData.email)) {
        return { success: false, message: `Invalid email address: ${userData.email}` };
      }
      const colIndex = headers.indexOf('email');
      if (colIndex > -1) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(userData.email);
        changes.push(`email to '${userData.email}'`);
      }
    }

    if (userData.hasOwnProperty('isActive')) {
      const colIndex = headers.indexOf('isactive');
      if (colIndex > -1) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(userData.isActive);
        changes.push(`status to '${userData.isActive ? 'Active' : 'Inactive'}'`);
      }
    }

    if (changes.length > 0) {
      const logDetails = `Admin '${currentUser.username}' updated user '${userData.username}': changed ${changes.join(', ')}.`;
      logAction('User Updated', '', '', '', '', '', 'Success', logDetails);
      Logger.log(logDetails);
      delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.USER_PROFILES}`];
      return { success: true, message: 'User updated successfully.' };
    } else {
      return { success: true, message: 'No changes were applied.' };
    }

  } catch (error) {
    Logger.log('Error updating user: ' + error.message);
    return { success: false, message: 'An unexpected error occurred while updating the user.' };
  }
}


// =============================================
// SETTINGS MANAGEMENT
// =============================================

function updateSystemConfig(newConfig) {
  Logger.log('updateSystemConfig called');

  try {
    if (newConfig.emailFrom && !isValidEmail(newConfig.emailFrom)) {
      return { success: false, message: 'Invalid email address' };
    }

    if (newConfig.paginationLimit && (isNaN(newConfig.paginationLimit) || newConfig.paginationLimit < 10 || newConfig.paginationLimit > 100)) {
      return { success: false, message: 'Pagination limit must be between 10 and 100' };
    }

    Logger.log('System config (pseudo) updated: ' + JSON.stringify(newConfig));
    return { success: true, message: 'System configuration updated successfully' };
  } catch (error) {
    Logger.log('Error updating system config: ' + error.message);
    return { success: false, message: 'Error updating system config: ' + error.message };
  }
}


// =============================================
// AUDIT LOG MANAGEMENT
// =============================================
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

/**
 * Logs an action to the Audit Log sheet with new dedicated columns for reporting.
 * @param {string} action The action being performed.
 * @param {string} jlid The learner's JLID.
 * @param {string} learner The learner's name.
 * @param {string} oldTeacher The original teacher.
 * @param {string} newTeacher The new teacher.
 * @param {string} course The course involved.
 * @param {string} status The result (e.g., 'Success', 'Failed').
 * @param {string} notes General notes or comments.
 * @param {string} reason The specific reason for a migration.
 * @param {string} intervenedBy Comma-separated list of teams that intervened.
 */
function logAction(action, jlid, learner, oldTeacher, newTeacher, course, status, notes, reason = '', intervenedBy = '') {
  Logger.log(`logAction called: ${action} for JLID: ${jlid}`);

  try {
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

    delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.AUDIT_LOG}`];

    Logger.log('Action logged successfully: ' + action);
  } catch (error) {
    Logger.log('Error logging action: ' + error.message);
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

// =============================================
// TEACHER & COURSE DATA ACCESS
// =============================================

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

// =============================================
// TEACHER PERSONA TOOL FUNCTIONS
// =============================================

/**
 * Searches for teachers matching specified criteria including availability, course proficiency, and traits.
 * FINAL FIX: Handles cases where a header in the Persona sheet might not be a string (e.g., a number or date).
 *
 * @param {object} requestData The search criteria submitted from the web app's form.
 * @returns {object} An object containing the success status and an array of matching teacher results.
 */
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

/**
 * Adds or updates a teacher's record in the Persona Tool sheet.
 * This function is robust against column reordering and non-string header values.
 *
 * @param {object} teacherData An object where keys are header names and values are the data to be written.
 * @returns {object} A result object with success status and a message.
 */

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

const TIMEZONE_FRIENDLY_LABELS = {
  'Europe/London': 'UK Time',
  'Europe/Berlin': 'CET',
  'Europe/Paris': 'CET',
  'America/New_York': 'EST/EDT',
  'America/Chicago': 'CST/CDT',
  'America/Los_Angeles': 'PST/PDT',
  'Asia/Kolkata': 'IST',
  'Asia/Dubai': 'GST',
  'Australia/Sydney': 'AEST',
  'Asia/Singapore': 'SGT'
};

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


// =============================================
// SYSTEM MANAGEMENT
// =============================================

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

function getSystemSettings() {
  return {
    emailFrom: CONFIG.EMAIL.FROM,
    emailFromName: CONFIG.EMAIL.FROM_NAME,
    paginationLimit: CONFIG.PAGINATION_LIMIT,
    auditRetentionDays: 90 
  };
}

// =============================================
// UTILITY FUNCTIONS
// =============================================

/**
 * NEW HELPER FUNCTION
 * Looks up the primary email address for a CLS Manager by their name.
 * Assumes CLS Managers exist in the 'Teacher Data' sheet with their name in Column H and email in Column J.
 * This is the refined version ensuring correct column lookup as per the "FINAL FIX" requirements.
 * @param {string} managerName The name of the CLS Manager to find.
 * @returns {string|null} The email address from Column J, or null if not found or invalid.
 */
function getClsManagerEmailByName(managerName) {
  Logger.log("getClsManagerEmailByName is deprecated. Use findClsEmailByManagerName directly.");
  return findClsEmailByManagerName(managerName);
}


function getOrCreateSheet(sheetName) {
  try {
    const spreadsheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID);
    let sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      Logger.log('Created new sheet: ' + sheetName);
    }

    return sheet;
  } catch (error) {
    Logger.log('Error getting/creating sheet: ' + error.message);
    throw error;
  }
}

function hasPermission(userRole, permission) {
  const userPermissions = PERMISSIONS[userRole] || [];
  return userPermissions.includes(permission);
}

/**
 * Finds all images in a specified Google Drive folder and converts them to Base64 data URIs.
 * @param {string} folderId The ID of the Google Drive folder.
 * @returns {object} An object mapping filenames to their Base64 data URI strings.
 *                   e.g., { "Logo.png": "data:image/png;base64,iVBORw0KGgo...", ... }
 */
function getImagesAsBase64FromDrive(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const imageData = {};

    while (files.hasNext()) {
      const file = files.next();
      const blob = file.getBlob();
      const contentType = blob.getContentType();
      const base64Data = Utilities.base64Encode(blob.getBytes());
      
      // Construct the full Data URI for embedding in HTML
      imageData[file.getName()] = `data:${contentType};base64,${base64Data}`;
    }
    
    Logger.log(`Successfully processed ${Object.keys(imageData).length} images from Drive.`);
    return imageData;

  } catch (e) {
    Logger.log(`Error accessing Drive folder or processing images: ${e.message}`);
    return {}; // Return empty object on error
  }
}


function testEmailConfiguration() {
  try {
    MailApp.sendEmail({
      to: CONFIG.EMAIL.MAIN_MANAGER,
      subject: 'JetLearn Migration System - Test Email',
      body: 'This is a test email from the JetLearn Migration System. If you received this, email sending is working correctly.',
      name: CONFIG.EMAIL.FROM_NAME,
      from: CONFIG.EMAIL.FROM 
    });
    return { success: true, message: 'Test email sent successfully to ' + CONFIG.EMAIL.MAIN_MANAGER };
  } catch (error) {
    Logger.log('Test email failed: ' + error.message);
    return { success: false, message: 'Test email failed: ' + error.message };
  }
}

/**
 * Reliably formats a date object or a date string (e.g., "2025-09-21") into DD/MM/YYYY format.
 * This version is robust against timezone issues.
 * @param {Date|string} dateInput The Date object or string to format.
 * @returns {string} The formatted date string or 'N/A' if invalid.
 */
function formatDateDDMMYYYY(dateInput) {
  try {
    let dateToFormat;
    
    if (dateInput instanceof Date && !isNaN(dateInput)) {
      dateToFormat = dateInput;
    } 
    else if (typeof dateInput === 'string') {
      dateToFormat = new Date(dateInput);
      if (isNaN(dateToFormat.getTime()) && /^\d{4}-\d{2}-\d{2}/.test(dateInput)) {
        const parts = dateInput.split('T')[0].split('-');
        dateToFormat = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
      }
    }

    if (!dateToFormat || isNaN(dateToFormat.getTime())) {
      Logger.log('Could not format invalid dateInput: ' + dateInput);
      return 'N/A';
    }

    const day = String(dateToFormat.getUTCDate()).padStart(2, '0');
    const month = String(dateToFormat.getUTCMonth() + 1).padStart(2, '0');
    const year = dateToFormat.getUTCFullYear();

    return `${day}/${month}/${year}`;
  } catch (e) {
    Logger.log('Error formatting date: ' + e.message);
    return 'N/A';
  }
}

/**
 * Formats a Date object into a more readable string like "13 August 2025".
 * @param {Date} dateObj The Date object to format.
 * @returns {string} The formatted date string.
 */
function formatDate(dateObj) {
  if (!dateObj || isNaN(dateObj.getTime())) return 'N/A';
  const options = { year: 'numeric', month: 'long', day: 'numeric' };
  return dateObj.toLocaleDateString('en-GB', options);
}

/**
 * Returns the currency symbol for a given currency code.
 * @param {string} currencyCode E.g., "GBP", "EUR", "USD", "INR", or a custom code.
 * @returns {string} The currency symbol.
 */
function getCurrencySymbol(currencyCode) {
    switch (currencyCode) {
        case 'GBP': return '£';
        case 'EUR': return '€';
        case 'USD': return '$';
        case 'INR': return '₹';
        case 'JPY': return '¥';
        case 'AUD': return 'A$';
        case 'CAD': return 'C$';
        case 'CHF': return 'CHF';
        case 'CNY': return '¥';
        case 'SEK': return 'kr';
        case 'NZD': return 'NZ$';
        case 'AED': return 'د.إ';
        case 'HKD': return 'HK$';
        case 'ZAR': return 'R';
        case 'SGD': return 'S$';
        case 'NOK': return 'kr';
        case 'DKK': return 'kr';
        case 'MXN': return 'Mex$';
        case 'BRL': return 'R$';
        default: return currencyCode || ''; 
    }
}

function isValidEmail(email) {
  if (!email || typeof email !== 'string') return false;
  const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  return emailRegex.test(email.trim());
}

function validateMigrationData(data) {
  const errors = [];

  if (!data.jlid || data.jlid.trim() === '') {
    errors.push('JLID is required');
  }

  if (!data.learner || data.learner.trim() === '') {
    errors.push('Learner name is required');
  }

  if (!data.newTeacher || data.newTeacher.trim() === '') {
    errors.push('New teacher is required');
  }

  if (!data.course || data.course.trim() === '') {
    errors.push('Course is required');
  }

  if (!data.classSessions || data.classSessions.length === 0) {
    errors.push('At least one class session (Day and Time) is required');
  } else {
    data.classSessions.forEach((session, index) => {
      if (!session.day || session.day.trim() === '') {
        errors.push(`Class Day for session ${index + 1} is required`);
      }
      if (!session.time || !session.time.match(/^\d{1,2}:\d{2}\s(AM|PM)$/i)) {
        errors.push(`Class Time for session ${index + 1} is invalid or missing. Expected format HH:MM AM/PM`);
      }
    });
  }

  if (!data.clsManager || data.clsManager.trim() === '') {
    errors.push('CLS Manager is required');
  }

  if (!data.jetGuide || data.jetGuide.trim() === '') {
    errors.push('JetGuide is required');
  }

  if (!data.startDate || data.startDate.trim() === '') {
    errors.push('Start Date is required');
  }

  if (!data.migrationType || !['Mid-Course', 'New Assignment'].includes(data.migrationType)) {
    errors.push('Valid migration type is required');
  }

  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

function processBatchUpload(csvContent) {
  Logger.log('processBatchUpload called');

  const lines = csvContent.split('\n');
  if (lines.length < 2) {
    return { total: 0, success: 0, failed: 0, results: [], message: 'CSV is empty or has only headers.' };
  }

  const headers = lines[0].split(',').map(h => h.trim());
  const requiredHeaders = ['jlid', 'learnerName', 'oldTeacher', 'newTeacher', 'course', 'reason', 'clsManager', 'jetGuide', 'startDate', 'day', 'time'];

  const missingHeaders = requiredHeaders.filter(rh => !headers.includes(rh));
  if (missingHeaders.length > 0) {
    return { success: false, message: `Missing required CSV headers: ${missingHeaders.join(', ')}.`, total: 0, results: [] };
  }

  const results = [];
  const migrationsToProcess = [];

  for (let i = 1; i < lines.length; i++) {
    if (lines[i].trim() === '') continue; 

    const row = lines[i].split(',');
    const migration = {};
    headers.forEach((header, index) => {
      migration[header] = row[index] ? row[index].trim() : '';
    });

    migration.classSessions = [{ day: migration.day, time: migration.time }];

    migration.migrationType = migration.migrationType || 'New Assignment';

    migrationsToProcess.push({ data: migration, originalRow: i + 1 });
  }

  if (migrationsToProcess.length === 0) {
    return { total: 0, success: 0, failed: 0, results: [], message: 'No valid migration rows found in CSV.' };
  }

  const batchResults = [];
  let successCount = 0;
  let failedCount = 0;

  migrationsToProcess.forEach(item => {
    const migrationData = item.data;
    const rowNum = item.originalRow;

    try {
      const validation = validateMigrationData(migrationData);

      if (!validation.isValid) {
        batchResults.push({
          row: rowNum,
          jlid: migrationData.jlid,
          learnerName: migrationData.learnerName,
          success: false,
          message: 'Validation failed: ' + validation.errors.join(', ')
        });
        failedCount++;
        return;
      }

      const sendResult = sendMigrationEmail(migrationData); 

      if (sendResult.success) {
        batchResults.push({
          row: rowNum,
          jlid: migrationData.jlid,
          learnerName: migrationData.learnerName,
          success: true,
          message: sendResult.message
        });
        successCount++;
      } else {
        batchResults.push({
          row: rowNum,
          jlid: migrationData.jlid,
          learnerName: migrationData.learnerName,
          success: false,
          message: sendResult.message
        });
        failedCount++;
      }

    } catch (error) {
      Logger.log(`Error processing batch row ${rowNum} for JLID ${migrationData.jlid}: ${error.message}`);
      batchResults.push({
        row: rowNum,
        jlid: migrationData.jlid,
        learnerName: migrationData.learnerName,
        success: false,
        message: 'Internal server error during processing: ' + error.message
      });
      failedCount++;
    }
  });

  logAction(
    'Batch Migration Processed',
    '', '', '', '', '',
    successCount === migrationsToProcess.length ? 'Success' :
      (failedCount === migrationsToProcess.length ? 'Failed' : 'Partial Success'),
    `Processed ${migrationsToProcess.length} migrations (${successCount} success, ${failedCount} failed)`
  );

  Logger.log('Batch migration processing completed');
  return {
    total: migrationsToProcess.length,
    success: successCount,
    failed: failedCount,
    results: batchResults
  };
}


// Renaming the existing `getLearnerHistory` to a more specific name
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

// AI Summarization Function (new)
/**
 * Summarize learner history using Google Gemini API
 */
 /**
 * Summarize learner history using Google Gemini API
 */
// === AI Summarization Function ===
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

/**
 * Retrieves and processes a comprehensive history for a given learner,
 * including HubSpot profile data, audit log events, and an AI-generated summary.
 * This function is designed to be robust against partial data availability or API failures.
 * @param {string} jlid - The JetLearn ID or learner name to search for.
 * @returns {object} A structured object containing learner profile, timeline, stats, AI summary, and success status/message.
 */
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

  try {
    if (!jlid || jlid.trim() === '') {
      return { success: false, message: 'Learner ID (JLID) or Name is required for history lookup.' };
    }

    const hubspotResult = fetchHubspotByJlid(jlid);
    if (hubspotResult.success) {
      hubspotData = hubspotResult.data;
    } else {
      Logger.log(`Warning: HubSpot data retrieval failed for ${jlid}: ${hubspotResult.message}`);
      hubspotData = { jlid: jlid, learnerName: jlid.toUpperCase().startsWith('JL') ? 'N/A' : jlid, isPartial: true };
      overallSuccess = false;
      overallMessage = `HubSpot data could not be retrieved: ${hubspotResult.message}. Audit log data may be available.`;
    }

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

      const requiredCols = [actionCol, timestampCol, oldTeacherCol, newTeacherCol, courseCol, notesCol, statusCol, reasonCol];
      if (requiredCols.some(col => col === -1)) {
        throw new Error("Audit Log sheet is missing one or more required columns for history generation.");
      }
      
      const maxIndexRequired = Math.max(...requiredCols);
      
      rawAuditLogResult.data.forEach(row => {
        if (!row || row.length <= maxIndexRequired) return;

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
          if (oldTeacher && newTeacher && oldTeacher !== newTeacher) {
            teacherChangeEvents.push({ eventDate: timestamp.toISOString(), fromTeacher: oldTeacher, toTeacher: newTeacher, reason: finalReason || action });
          }
          if (course) { 
            courseChangeEvents.push({ eventDate: timestamp.toISOString(), course: course, reason: finalReason || action });
          }
        } else if (action.includes('Onboarding')) {
          eventType = 'onboarding';
          entryDescription = `Learner onboarded with Teacher: ${newTeacher || 'N/A'} for Course: ${course || 'N/A'}. Status: ${status}.`;
          if (newTeacher) {
            teacherChangeEvents.push({ eventDate: timestamp.toISOString(), fromTeacher: 'Initial', toTeacher: newTeacher, reason: action });
          }
        }
        
        auditLogTimeline.push({
          timestamp: timestamp.toISOString(),
          type: eventType,
          description: entryDescription,
        });
      });
    }

    auditLogTimeline.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));

    if (hubspotData && (hubspotData.learnerName !== 'N/A' || auditLogTimeline.length > 0)) {
      const aiResult = summarizeLearnerHistory(hubspotData, auditLogTimeline);
      aiSummary = aiResult.success ? aiResult.summary : `AI summary unavailable: ${aiResult.message}`;
    } else {
      aiSummary = 'Not enough information for an AI summary.';
    }

    return {
      success: overallSuccess,
      learnerProfile: hubspotData,
      auditLogTimeline: auditLogTimeline,
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
    Logger.log('FATAL Error in getComprehensiveLearnerHistory: ' + error.message + ' Stack: ' + error.stack);
    return { success: false, message: 'An unexpected server error occurred while retrieving learner history: ' + error.message };
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

/**
 * 1. Gets a simple list of teachers for the dropdown.
 * Removes the 'Teacher' header row to fix your visual bug.
 */
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


function subscribeToWeeklyReport() {
    Logger.log('subscribeToWeeklyReport called');

    const userEmail = Session.getActiveUser().getEmail(); 

    Logger.log(`User ${userEmail} requested subscription to weekly report.`);
    return { success: true, message: `Subscription request received for ${userEmail}. (Feature not fully implemented)` };
}

// Dummy/Placeholder for getTeacherProfileData (if not already implemented)
function getTeacherProfileData(teacherName) {
  Logger.log('getTeacherProfileData called for: ' + teacherName);
  return {
    success: true,
    profile: {
      name: teacherName,
      email: `${teacherName.toLowerCase().replace(/\s/g, '.')}@example.com`,
      status: 'Active',
      manager: 'Jane Doe', 
      joinDate: '2022-01-15', 
      coursesByCategory: {
        'Active Courses': [{ name: 'GCSE Programming', progress: '75%' }, { name: 'AI Fundamentals', progress: '40%' }],
        'Completed Courses': [{ name: 'Beginner Python', progress: '100%' }]
      }
    },
    message: `Profile data for ${teacherName} (dummy data).`
  };
}

// =============================================
// SYSTEM INITIALIZATION
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

// =============================================
// NEW: MONTHLY REPORTING AND AI INSIGHTS
// =============================================

/**
 * Generates an aggregated report of migration data for a specific month and year.
 * @param {number} month - The month to report on (0-11, where 0 is January).
 * @param {number} year - The year to report on.
 * @param {string} perspective - The team perspective ('CLS', 'TP', 'Ops', or 'All'). Defaults to 'All'.
 * @returns {object} A JSON object containing the aggregated report data.
 */
function generateMonthlyReport(month, year, perspective = 'All') {
    Logger.log(`Generating monthly report for month: ${month + 1}, year: ${year}, Perspective: ${perspective}`);
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.AUDIT_LOG);
    const headers = sheetData[0] || [];
    const timestampCol = 0;
    const actionCol = headers.indexOf('Action');
    const reasonCol = headers.indexOf('Reason for Migration');
    const intervenedCol = headers.indexOf('Intervened By');

    if (actionCol === -1 || reasonCol === -1 || intervenedCol === -1) {
        Logger.log("Audit Log sheet is missing required columns for reporting: Action, Reason for Migration, Intervened By.");
        return { month: "Error", year: year, total: 0, CLS_Involvement: 0, TP_Involvement: 0, Ops_Involvement: 0, reasons: { "Error": "Missing Columns" } };
    }
    
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sept", "Oct", "Nov", "Dec"];

    const report = {
        month: monthNames[month],
        year: year,
        perspective: perspective,
        total: 0,
        CLS_Involvement: 0,
        TP_Involvement: 0,
        Ops_Involvement: 0,
        reasons: {}
    };

    const perspectiveFilter = perspective.toLowerCase();

    const migrationLogs = sheetData.slice(1).filter(row => {
        if (!row[actionCol] || !row[actionCol].includes('Migration')) return false;

        const timestamp = new Date(row[timestampCol]);
        if (timestamp.getFullYear() !== year || timestamp.getMonth() !== month) return false;

        const teams = (row[intervenedCol] || '').toLowerCase();
        if (perspective !== 'All' && !teams.includes(perspectiveFilter)) {
            return false;
        }
        return true;
    });

    report.total = migrationLogs.length;
    if (report.total === 0) {
        Logger.log(`No migration data found for perspective ${perspective} in the specified month.`);
        return report;
    }

    migrationLogs.forEach(row => {
        const teams = (row[intervenedCol] || '').toLowerCase();
        if (teams.includes('cls')) report.CLS_Involvement++;
        if (teams.includes('tp')) report.TP_Involvement++;
        if (teams.includes('ops')) report.Ops_Involvement++;

        const reason = (row[reasonCol] || 'Unknown').trim();
        if (reason) {
            report.reasons[reason] = (report.reasons[reason] || 0) + 1;
        }
    });

    Logger.log(`Monthly report generated successfully for ${perspective}: ` + JSON.stringify(report));
    return report;
}

/**
 * Generates and emails a notification with a link to the hosted monthly report.
 * This function is designed to be run automatically by a time-based trigger.
 */
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

/**
 * Creates a time-driven trigger to run the emailMonthlyReport function monthly.
 * This function should be run MANUALLY once from the Apps Script editor to set up automation.
 */
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

function testApiCall() {
  const rates = getLiveCurrencyRates();
  console.log(rates);
}
/**
 * Reliably parses a date from the Google Sheet, supporting multiple formats.
 * @param {string|Date} sheetDate The value from the sheet's date column.
 * @returns {Date|null} A valid Date object or null if parsing fails.
 */
function parseSheetDate(sheetDate) {
  if (sheetDate instanceof Date && !isNaN(sheetDate)) return sheetDate;

  const directDate = new Date(sheetDate);
  if (directDate instanceof Date && !isNaN(directDate)) return directDate;
  
  if (typeof sheetDate === 'string') {
    const datePart = sheetDate.split(' ')[0].replace(/,/g, '');
    const parts = datePart.split(/[\/-]/); 
    if (parts.length === 3) {
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1; 
      let year = parseInt(parts[2], 10);

      if (year < 100) {
        year += 2000;
      }
      
      if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
          const customDate = new Date(Date.UTC(year, month, day)); 
          if (sheetDate.includes(' ')) {
              const timeString = sheetDate.split(' ')[1];
              if (timeString && timeString.includes(':')) {
                  const timeParts = timeString.split(':');
                  if (timeParts.length >= 2) {
                      customDate.setUTCHours(parseInt(timeParts[0], 10) || 0);
                      customDate.setUTCMinutes(parseInt(timeParts[1], 10) || 0);
                      if (timeParts.length === 3) {
                        customDate.setUTCSeconds(parseInt(timeParts[2], 10) || 0);
                      }
                  }
              }
          }
          if (!isNaN(customDate)) return customDate;
      }
    }
  }
  return null;
}
// =============================================
// ### NEW ### EMAIL TRACKING & LOGGING MODULE
// =============================================

/**
 * Handles the request for the tracking pixel, logs the 'open' event, and returns a transparent image.
 * This is called by the pixel in the recipient's email client.
 * @param {string} trackingId The unique ID for the email being tracked.
 * @returns {GoogleAppsScript.Content.TextOutput} A 1x1 transparent PNG image.
 */
// function sendTrackedEmail(payload) {
//   const { to, cc, subject, htmlBody, jlid, attachments } = payload;
//   const trackingId = Utilities.getUuid();

//   try {
//     if (!to || !subject || !htmlBody) {
//       throw new Error("Missing required fields: 'to', 'subject', and 'htmlBody' are mandatory.");
//     }

//     const webAppUrl = ScriptApp.getService().getUrl();
//     const trackingPixelUrl = `${webAppUrl}?page=track&id=${trackingId}`;
//     const trackingPixelImg = `<img src="${trackingPixelUrl}" width="1" height="1" style="display:none;"/>`;
//     const finalHtmlBody = htmlBody.replace('</body>', `${trackingPixelImg}</body>`);

//     const mailOptions = {
//       to: to,
//       subject: subject,
//       htmlBody: finalHtmlBody,
//       name: CONFIG.EMAIL.FROM_NAME,
//       from: CONFIG.EMAIL.FROM
//     };
//     if (cc) mailOptions.cc = cc;
//     if (attachments) mailOptions.attachments = attachments;

//     MailApp.sendEmail(mailOptions);

//     logEmail({
//       trackingId: trackingId,
//       recipient: to,
//       subject: subject,
//       jlid: jlid,
//       status: 'Sent',
//       sentAt: new Date(),
//       htmlBody: finalHtmlBody,
//       attachments: attachments
//     });

//     Logger.log(`Successfully sent and logged tracked email to ${to} with ID ${trackingId}.`);
//     return { success: true, trackingId: trackingId };

//   } catch (error) {
//     Logger.log(`Failed to send tracked email to ${to}. Error: ${error.message}`);
//     logEmail({
//       trackingId: trackingId,
//       recipient: to,
//       subject: `FAILED: ${subject}`,
//       jlid: jlid,
//       status: 'Failed',
//       sentAt: new Date(),
//       htmlBody: htmlBody,
//       rawPayload: error.message
//     });
//     throw error; // Re-throw the error so the calling function knows it failed
//   }
// }

function handleTrackingPixel(trackingId) {
  const transparentPngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";
  try {
    if (trackingId) {
      const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.EMAIL_LOGS);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const trackingIdCol = headers.indexOf('Tracking ID') + 1;
      const openedAtCol = headers.indexOf('Opened At') + 1;
      const statusCol = headers.indexOf('Status') + 1;

      if (trackingIdCol > 0 && openedAtCol > 0 && statusCol > 0) {
        const textFinder = sheet.createTextFinder(trackingId).matchEntireCell(true);
        const foundRange = textFinder.findNext();

        if (foundRange) {
          const row = foundRange.getRow();
          if (sheet.getRange(row, openedAtCol).isBlank()) {
            sheet.getRange(row, openedAtCol).setValue(new Date());
            const currentStatus = sheet.getRange(row, statusCol).getValue();
            if (currentStatus === 'Sent') {
              sheet.getRange(row, statusCol).setValue('Opened');
            }
          }
        }
      }
    }
  } catch (error) {
    Logger.log(`Error processing tracking pixel for ID ${trackingId}: ${error.message}`);
  }
  return ContentService.createTextOutput(Utilities.base64Decode(transparentPngBase64)).setMimeType(ContentService.MimeType.PNG);
}

/**
 * Logs email metadata to the 'Email Logs' sheet.
 * @param {object} logData The data to log.
 */
function logEmail(logData) {
  try {
    const sheet = getOrCreateSheet(CONFIG.SHEETS.EMAIL_LOGS);
    if (sheet.getLastRow() === 0) {
        const headers = ['Timestamp', 'Tracking ID', 'Recipient', 'Subject', 'Related JLID', 'Status', 'Sent At', 'Opened At', 'Replied At', 'HTML File URL', 'Headers (JSON)', 'Attachments (JSON)', 'Reply Content', 'Raw Payload'];
        sheet.appendRow(headers);
    }

    let htmlFileUrl = '';
    if (logData.htmlBody) {
        try {
            const now = new Date();
            const folderName = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM");
            const parentFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
            let monthFolder;
            const folders = parentFolder.getFoldersByName(folderName);
            monthFolder = folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
            
            const fileName = `email_${logData.trackingId}.html`;
            const file = monthFolder.createFile(fileName, logData.htmlBody, MimeType.HTML);
            htmlFileUrl = file.getUrl();
        } catch (e) {
            Logger.log(`Could not save email body to Drive: ${e.message}`);
            htmlFileUrl = `Error: ${e.message}`;
        }
    }

    const newRow = [
      new Date(),
      logData.trackingId || '',
      logData.recipient || '',
      logData.subject || '',
      logData.jlid || '',
      logData.status || 'Unknown',
      logData.sentAt || '',
      '', // Opened At
      '', // Replied At
      htmlFileUrl,
      logData.headers ? JSON.stringify(logData.headers) : '',
      logData.attachments ? JSON.stringify(logData.attachments.map(a => a.getName ? a.getName() : a.fileName)) : '[]',
      '', // Reply Content
      logData.rawPayload || ''
    ];
    sheet.appendRow(newRow);
  } catch (error) {
    Logger.log(`Failed to log email data. Error: ${error.message}`);
  }
}

/**
 * A function to be run on a time-based trigger to check for replies to sent emails.
 */
function checkReplies() {
  Logger.log("Starting reply check job...");
  try {
    const sheet = getOrCreateSheet(CONFIG.SHEETS.EMAIL_LOGS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const statusCol = headers.indexOf('Status');
    const recipientCol = headers.indexOf('Recipient');
    const subjectCol = headers.indexOf('Subject');
    const sentAtCol = headers.indexOf('Sent At');
    const repliedAtCol = headers.indexOf('Replied At');
    const replyContentCol = headers.indexOf('Reply Content');

    for (let i = data.length - 1; i > 0; i--) {
      const row = data[i];
      const currentStatus = row[statusCol];

      if (currentStatus === 'Sent' || currentStatus === 'Opened') {
        const recipient = row[recipientCol];
        const subject = row[subjectCol];
        const sentAt = new Date(row[sentAtCol]);
        if (!recipient || !subject || isNaN(sentAt.getTime())) continue;

        const searchQuery = `subject:("${subject}") from:(${recipient}) after:${Utilities.formatDate(sentAt, Session.getScriptTimeZone(), "yyyy/MM/dd")}`;
        const threads = GmailApp.search(searchQuery, 0, 1);

        if (threads.length > 0) {
          const messages = threads[0].getMessages();
          for (const message of messages) {
            const messageDate = message.getDate();
            if (message.getFrom().includes(recipient) && messageDate > sentAt) {
              sheet.getRange(i + 1, statusCol + 1).setValue('Replied');
              sheet.getRange(i + 1, repliedAtCol + 1).setValue(messageDate);
              sheet.getRange(i + 1, replyContentCol + 1).setValue(message.getPlainBody().substring(0, 500));
              break; 
            }
          }
        }
      }
    }
    Logger.log("Reply check job finished.");
  } catch (error) {
    Logger.log(`Error during reply check job: ${error.message}`);
  }
}

/**
 * Fetches email logs for the frontend dashboard.
 * @param {object} params Parameters for fetching data (searchTerm, page, pageSize).
 * @returns {object} An object containing the log data and pagination info.
 */
function getEmailActivities(params = {}) {
  Logger.log('getEmailActivities called with params: ' + JSON.stringify(params));

  try {
    const sheet = getOrCreateSheet(CONFIG.SHEETS.EMAIL_LOGS);
    // Safety check for empty sheet
    if (sheet.getLastRow() < 2) {
      return { success: true, data: [], pagination: { currentPage: 1, pageSize: 25, totalItems: 0, totalPages: 0 } };
    }
    
    const allData = sheet.getDataRange().getValues();
    const headers = allData.shift(); // Remove headers
    const headerMap = headers.map(h => String(h).replace(/\s/g, '')); 

    let rows = allData;

    // Search Filter
    if (params.searchTerm) {
      const term = params.searchTerm.toLowerCase();
      rows = rows.filter(row => row.some(cell => String(cell).toLowerCase().includes(term)));
    }

    // Sort by Date (Assuming Timestamp is column 0) - Newest First
    rows.sort((a, b) => {
      // SAFE DATE PARSING
      const dateA = a[0] ? new Date(a[0]) : new Date(0); // Default to epoch if empty
      const dateB = b[0] ? new Date(b[0]) : new Date(0);
      
      // If date is invalid, treat as old
      const timeA = isNaN(dateA.getTime()) ? 0 : dateA.getTime();
      const timeB = isNaN(dateB.getTime()) ? 0 : dateB.getTime();
      
      return timeB - timeA;
    });

    // Pagination
    const page = params.page || 1;
    const pageSize = params.pageSize || 25;
    const totalItems = rows.length;
    const totalPages = Math.ceil(totalItems / pageSize);
    const startIdx = (page - 1) * pageSize;
    const paginatedRows = rows.slice(startIdx, startIdx + pageSize);

    // Map to Object and SANITIZE DATES
    const results = paginatedRows.map(row => {
      const obj = {};
      headerMap.forEach((key, index) => {
        let value = row[index];
        // CRITICAL FIX: Convert Date objects to Strings so GAS doesn't crash
        if (value instanceof Date) {
          try {
            value = value.toISOString();
          } catch(e) {
            value = 'Invalid Date';
          }
        }
        obj[key] = value;
      });
      return obj;
    });

    return {
      success: true,
      data: results,
      pagination: { currentPage: page, pageSize, totalItems, totalPages }
    };

  } catch (error) {
    Logger.log('Error in getEmailActivities: ' + error.message);
    return { success: false, message: "Server error: " + error.message };
  }
}


/**
 * Fetches the full details of a single logged email by its tracking ID.
 * @param {string} trackingId The unique ID of the email log.
 * @returns {object} The full log entry object or an error object.
 */
function getEmailLogDetails(trackingId) {
  if (!trackingId) return { success: false, message: "Invalid ID." };

  try {
    const sheet = getOrCreateSheet(CONFIG.SHEETS.EMAIL_LOGS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const headerMap = headers.map(h => String(h).replace(/\s/g, ''));
    
    // Find ID Column (usually 'Tracking ID' -> 'TrackingID')
    const idIndex = headerMap.indexOf('TrackingID');
    if (idIndex === -1) return { success: false, message: "Tracking ID column not found." };

    // Find Row
    const row = data.slice(1).find(r => String(r[idIndex]).trim() === String(trackingId).trim());
    
    if (!row) return { success: false, message: "Log entry not found." };

    const logObject = {};
    headerMap.forEach((key, index) => {
      let value = row[index];
      
      // CRITICAL FIX: Convert Date objects to Strings
      if (value instanceof Date) {
        value = value.toISOString();
      }
      logObject[key] = value;
    });

    // Handle HTML Content from Drive
    if (logObject['HTMLFileURL'] && String(logObject['HTMLFileURL']).includes('drive.google.com')) {
        try {
            const match = logObject['HTMLFileURL'].match(/[-\w]{25,}/);
            if (match) {
                logObject['EmailHTMLBody'] = DriveApp.getFileById(match[0]).getBlob().getDataAsString();
            }
        } catch (e) {
            logObject['EmailHTMLBody'] = "Error loading content: " + e.message;
        }
    }

    return { success: true, data: logObject };

  } catch (error) {
    Logger.log('Error in getEmailLogDetails: ' + error.message);
    return { success: false, message: error.message };
  }
}

// =============================================
// ### NEW ### ONBOARDING AUDIT CENTER (VERSION 2.0)
// =============================================

/**
 * Fetches deals from HubSpot where the onboarding completion date falls within a specified range.
 * @param {string} fromDate - The start date in YYYY-MM-DD format.
 * @param {string} toDate - The end date in YYYY-MM-DD format.
 * @returns {Array<Object>} An array of HubSpot deal objects.
 */
function fetchDealsByOnboardingCompletionDate(fromDate, toDate) {
  const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  if (!token) throw new Error('HubSpot API token not configured.');

  const hubspotApiUrl = 'https://api.hubapi.com/crm/v3/objects/deals/search';
  const onboardingDateProperty = 'onboarding_completion_date';

  // Using the exact internal property names provided by the user
  const propertiesToFetch = [
    'dealname', 'jetlearner_id', 'amount', 'deal_currency_code', 'hs_object_id', 'age', 'learner_status',
    'module_start_date', 'module_end_date', 'total_classes_committed_through_learner_s_journey',
    'current_teacher', 'current_course', 'time_zone', 'regular_class_day', 'frequency_of_classes',
    'payment_type', 'subscription', 'subscription_tenure', 'payment_term',
    'learner_practice_document_link', onboardingDateProperty
  ];

  const requestBody = {
    filterGroups: [{
      filters: [{
        propertyName: onboardingDateProperty,
        operator: 'BETWEEN',
        highValue: new Date(`${toDate}T23:59:59Z`).getTime(),
        value: new Date(`${fromDate}T00:00:00Z`).getTime()
      }]
    }],
    properties: propertiesToFetch,
    limit: 100 
  };
  
  const options = {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(hubspotApiUrl, options);
  const jsonResponse = JSON.parse(response.getContentText());

  if (response.getResponseCode() !== 200) {
    throw new Error(`HubSpot API Error (${response.getResponseCode()}): ${jsonResponse.message || 'Unknown error'}`);
  }
  return jsonResponse.results || [];
}


/**
 * Fetches the most recent note associated with a HubSpot deal.
 * @param {string} dealId - The HubSpot ID of the deal.
 * @returns {string|null} The body of the latest note, or null if no note is found.
 */
function fetchLatestSalesNoteForDeal(dealId) {
  const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  if (!token) throw new Error('HubSpot API key not configured.');

  // 1. Get associations to the 5 most recent notes
  const associationUrl = `https://api.hubapi.com/crm/v4/objects/deal/${dealId}/associations/note?limit=5`;
  const assocOptions = { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true };
  const assocResponse = UrlFetchApp.fetch(associationUrl, assocOptions);

  if (assocResponse.getResponseCode() !== 200) {
      Logger.log(`Could not fetch note associations for deal ${dealId}. Status: ${assocResponse.getResponseCode()}`);
      return null;
  }
  
  const assocData = JSON.parse(assocResponse.getContentText());
  if (!assocData.results || assocData.results.length === 0) {
      Logger.log(`No notes found for deal ${dealId}.`);
      return null;
  }
  
  const noteIds = assocData.results.map(r => r.toObjectId);

  // 2. Fetch the content of those 5 notes
  const notesUrl = `https://api.hubapi.com/crm/v3/objects/notes/batch/read`;
  const notesOptions = {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify({
      properties: ["hs_note_body", "hs_createdate"],
      inputs: noteIds.map(id => ({ id }))
    }),
    muteHttpExceptions: true
  };

  const notesResponse = UrlFetchApp.fetch(notesUrl, notesOptions);
  if (notesResponse.getResponseCode() !== 200) {
      Logger.log(`Could not fetch note details for deal ${dealId}. Status: ${notesResponse.getResponseCode()}`);
      return null;
  }

  const notesData = JSON.parse(notesResponse.getContentText());
  if (!notesData.results || notesData.results.length === 0) return null;

  // Sort notes by creation date, most recent first
  const sortedNotes = notesData.results.sort((a, b) => new Date(b.properties.hs_createdate) - new Date(a.properties.hs_createdate));
  const noteBodies = sortedNotes.map(n => n.properties.hs_note_body).filter(Boolean);

  // 3. Intelligently find the best note
  // Priority 1: Find the detailed 15-point Sales Note.
  const salesNote = noteBodies.find(note => /^\s*1:\s*Learner Name/im.test(note.replace(/<[^>]*>/g, '')));
  if (salesNote) {
    Logger.log(`Found detailed 15-point Sales Note for deal ${dealId}.`);
    return salesNote;
  }

  // Priority 2: Find the shorter Onboarding Note as a fallback.
  const onboardingNote = noteBodies.find(note => /payment received/i.test(note) && /athena checked/i.test(note));
  if (onboardingNote) {
    Logger.log(`Sales Note not found. Falling back to Onboarding Note for deal ${dealId}.`);
    return onboardingNote;
  }

  // Priority 3: Return the most recent note if no specific format is found.
  Logger.log(`No specific note format found. Returning most recent note for deal ${dealId}.`);
  return noteBodies[0] || null;
}

/**
 * [REWRITTEN] Parses different formats of sales/onboarding notes into a standardized object.
 * This version is robust and handles both the 15-point Sales Note and the shorter Onboarding Note.
 * @param {string} noteText - The raw text of the note.
 * @returns {Object} A structured object with the parsed data.
 */
function parseSalesNote(noteText) {
  if (!noteText) return {};
  const data = {};
  const cleanText = noteText.replace(/<[^>]*>/g, '\n'); // Strip HTML tags and replace with newlines for better regex matching

  // Helper function for robust extraction
  const extract = (regex, text = cleanText) => (text.match(regex) || [])[1]?.trim() || null;

  // --- Detection Logic ---
  const isSalesNote = /^\s*1:\s*Learner Name/im.test(cleanText);
  const isOnboardingNote = /payment received/i.test(cleanText) && /athena checked/i.test(cleanText);

  if (isSalesNote) {
    Logger.log("Parsing as a detailed 15-point Sales Note.");
    data.learnerName = extract(/1:\s*Learner Name:\s*([^\n]+)/i);
    
    // Improved regex to handle currency symbol before or after the amount
    const dealAmountMatch = cleanText.match(/3:\s*Total Deal Amount with currency:\s*([€$£A-Z]+)?\s*([\d.,]+)\s*([€$£A-Z]+)?/i);
    if (dealAmountMatch) {
        data.totalDealCurrency = (dealAmountMatch[1] || dealAmountMatch[3] || '').trim().toUpperCase();
        data.totalDealAmount = parseFloat(dealAmountMatch[2].replace(/[.,]$/g, '').replace(/,/g, ''));
    }
    
    data.paymentType = extract(/4:\s*Payment Type:\s*([^\n]+)/i);
    data.subscriptionDuration = parseInt(extract(/8:\s*Subscription Duration:\s*(\d+)/i), 10) || null;
    data.courseEnrolled = extract(/11:\s*Course enrolled on:\s*([^\n]+)/i);
    data.committedClasses = parseInt(extract(/12:\s*Number of committed class:\s*(\d+)/i), 10) || null;
    data.classFrequency = extract(/15:\s*Class Frequency:\s*(.+)/i);
    data.teacherPreference = extract(/10:\s*Pref teacher:\s*([^\n]+)/i);

  } else if (isOnboardingNote) {
    Logger.log("Parsing as a short Onboarding Note.");
    
    const paymentMatch = cleanText.match(/Payment Received\s*-\s*([\d.,]+)\s*([A-Z]{3})/i);
    if (paymentMatch) {
      data.totalDealAmount = parseFloat(paymentMatch[1].replace(/,/g, ''));
      data.totalDealCurrency = paymentMatch[2].toUpperCase();
    }
    
    data.courseEnrolled = extract(/Athena Checked\s*-\s*(.+)/i);
    data.teacherQualified = extract(/Teacher Qualified\s*-\s*(.+)/i);
    data.timeZone = extract(/TZ\s*-\s*(.+)/i);

  } else {
    Logger.log("Note format not recognized. No data parsed.");
  }

  // Sanitize null values and empty strings for consistency
  for (const key in data) {
    if (data[key] === null || data[key] === undefined || (typeof data[key] === 'number' && isNaN(data[key])) || data[key] === 'na') {
        data[key] = null;
    }
  }
  return data;
}

/**
 * NEW/ENHANCED: Verifies subscription details against Google Calendar.
 * Checks both the start/end dates and the total number of classes.
 * @param {string} jlid - The learner's JLID.
 * @param {string} expectedStartDate - The subscription start date from HubSpot.
 * @param {string} expectedEndDate - The subscription end date from HubSpot.
 * @param {number} expectedClassCount - The total committed classes from HubSpot.
 * @returns {object} An object with verification results for dates and class count.
 */
function verifySubscriptionWithCalendar(jlid, expectedStartDate, expectedEndDate, expectedClassCount) {
    Logger.log(`Starting calendar verification for JLID: ${jlid}`);
    const calendarId = CONFIG.CLASS_SCHEDULE_CALENDAR_ID;
    const results = {
        dateCheck: { status: 'Skipped', message: 'Calendar ID not configured.' },
        countCheck: { status: 'Skipped', message: 'Calendar ID not configured.' }
    };

    if (!calendarId || calendarId === 'YOUR_GOOGLE_CALENDAR_ID_HERE') {
        Logger.log('Calendar verification skipped: CALENDAR_ID is not configured.');
        return results;
    }

    // This requires the Calendar API to be enabled in your Apps Script project.
    // Go to "Services" > "+" and add the "Google Calendar API".
    try {
        if (!jlid) {
            Logger.log('Calendar verification warning: JLID is missing.');
            results.dateCheck = { status: 'Warning', message: 'JLID is missing, cannot search calendar.' };
            results.countCheck = { status: 'Warning', message: 'JLID is missing, cannot search calendar.' };
            return results;
        }
        
        // Define a reasonable search window to avoid timeouts.
        const searchStartDate = expectedStartDate ? new Date(expectedStartDate) : new Date();
        searchStartDate.setDate(searchStartDate.getDate() - 7); // Start search 7 days prior to be safe
        const searchEndDate = new Date(searchStartDate.getFullYear() + 4, searchStartDate.getMonth(), searchStartDate.getDate()); // Search 4 years into the future

        Logger.log(`Searching calendar '${calendarId}' for events with "${jlid}" between ${searchStartDate} and ${searchEndDate}`);
        
        // Isolate the Calendar API call
        let events;
        try {
            events = Calendar.Events.list(calendarId, {
                q: jlid,
                timeMin: searchStartDate.toISOString(),
                timeMax: searchEndDate.toISOString(),
                singleEvents: true,
                orderBy: 'startTime'
            }).items;
             Logger.log(`Calendar API call successful. Found ${events.length} potential events.`);
        } catch(apiError) {
            Logger.log(`CRITICAL CALENDAR API ERROR: ${apiError.message}. This might be a permissions issue. Ensure Calendar API is enabled.`);
            throw new Error(`Google Calendar API failed: ${apiError.message}. Please ensure the API service is enabled and permissions are correct.`);
        }


        if (!events || events.length === 0) {
            Logger.log('No classes found in calendar for this JLID.');
            results.dateCheck = { status: 'Warning', message: 'No classes found in the calendar for this JLID.' };
            results.countCheck = { status: 'Warning', message: `Expected ${expectedClassCount || 'N/A'} classes, but found 0 in calendar.` };
            return results;
        }

        // 1. Verify Class Count
        Logger.log(`Verifying class count. HubSpot: ${expectedClassCount}, Calendar: ${events.length}`);
        if (expectedClassCount !== null && !isNaN(expectedClassCount)) {
            if (events.length !== expectedClassCount) {
                results.countCheck = { status: 'Mismatch', message: `HubSpot expects ${expectedClassCount} classes, but calendar has ${events.length}.` };
            } else {
                results.countCheck = { status: 'Match', message: `Count matches: ${events.length} classes.` };
            }
        } else {
            results.countCheck = { status: 'Info', message: 'Committed classes not specified in HubSpot for comparison.' };
        }

        // 2. Verify Dates
        const firstEvent = events[0];
        const lastEvent = events[events.length - 1];

        const firstClassDate = new Date(firstEvent.start.dateTime || firstEvent.start.date);
        const lastClassDate = new Date(lastEvent.end.dateTime || lastEvent.end.date);
        const hsStartDate = expectedStartDate ? new Date(expectedStartDate) : null;
        const hsEndDate = expectedEndDate ? new Date(expectedEndDate) : null;
        
        Logger.log(`Verifying dates. HubSpot Start: ${hsStartDate}, Calendar First Class: ${firstClassDate}`);
        Logger.log(`HubSpot End: ${hsEndDate}, Calendar Last Class: ${lastClassDate}`);

        let dateMismatches = [];
        if (hsStartDate) {
            hsStartDate.setHours(0,0,0,0);
            firstClassDate.setHours(0,0,0,0);
            if (hsStartDate.getTime() !== firstClassDate.getTime()) {
                dateMismatches.push(`Start Date Mismatch (HS: ${formatDate(hsStartDate)}, Calendar: ${formatDate(firstClassDate)})`);
            }
        } else {
            dateMismatches.push("HubSpot Start Date is missing.");
        }

        if (hsEndDate) {
            hsEndDate.setHours(0,0,0,0);
            lastClassDate.setHours(0,0,0,0);
            if (hsEndDate.getTime() !== lastClassDate.getTime()) {
                dateMismatches.push(`End Date Mismatch (HS: ${formatDate(hsEndDate)}, Calendar: ${formatDate(lastClassDate)})`);
            }
        } else {
            dateMismatches.push("HubSpot End Date is missing.");
        }

        if (dateMismatches.length > 0) {
            results.dateCheck = { status: 'Mismatch', message: dateMismatches.join('; ') };
        } else {
            results.dateCheck = { status: 'Match', message: `Dates align (Start: ${formatDate(hsStartDate)}, End: ${formatDate(hsEndDate)}).` };
        }
        
        Logger.log("Calendar verification finished successfully.");
        return results;

    } catch (e) {
        Logger.log(`FATAL ERROR during calendar verification for JLID ${jlid}: ${e.message} \nStack: ${e.stack}`);
        const errorMsg = `Calendar verification failed: ${e.message}`;
        return {
            dateCheck: { status: 'Error', message: errorMsg },
            countCheck: { status: 'Error', message: errorMsg }
        };
    }
}

/**
 * NEW: Fetches and compares all data points for a single deal for the "View Details" popup.
 * @param {string} dealId - The HubSpot ID of the deal.
 * @returns {Object} A detailed comparison object.
 */
function getAuditDetails(dealId) {
    Logger.log(`Fetching audit details for deal ID: ${dealId}`);
    try {
        const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
        const properties = [
            'dealname', 'jetlearner_id', 'amount', 'deal_currency_code', 'age', 'learner_status',
            'module_start_date', 'module_end_date', 'total_classes_committed_through_learner_s_journey', 
            'current_teacher', 'current_course', 'time_zone', 'regular_class_day', 'frequency_of_classes', 
            'payment_type', 'subscription', 'subscription_tenure', 'payment_term', 
            'learner_practice_document_link'
        ];
        const dealUrl = `https://api.hubapi.com/crm/v3/objects/deals/${dealId}?properties=${properties.join(',')}`;
        
        const dealOptions = { 
            headers: { 'Authorization': 'Bearer ' + token },
            muteHttpExceptions: true
        };
        
        Logger.log("Fetching deal from HubSpot...");
        const dealResponse = UrlFetchApp.fetch(dealUrl, dealOptions);

        if (dealResponse.getResponseCode() !== 200) {
            const errorBody = dealResponse.getContentText();
            Logger.log(`HubSpot API Error fetching deal ${dealId}: ${errorBody}`);
            throw new Error(`Failed to fetch HubSpot deal. Status: ${dealResponse.getResponseCode()}. Please check if the Deal ID is valid.`);
        }

        const dealData = JSON.parse(dealResponse.getContentText());
        const props = dealData.properties;
        Logger.log("Deal data fetched successfully.");

        // Use the NEW intelligent note finding function
        const noteText = fetchLatestSalesNoteForDeal(dealId);
        // Use the NEW multi-format parser
        const noteData = parseSalesNote(noteText);
        Logger.log("Sales note parsed.", noteData);

        const calendarResult = verifySubscriptionWithCalendar(
            props.jetlearner_id, 
            props.module_start_date, 
            props.module_end_date, 
            parseInt(props.total_classes_committed_through_learner_s_journey)
        );

        Logger.log("Building comparison details...");
        const details = [
            getComparisonResult('Amount', parseFloat(props.amount), noteData.totalDealAmount, (hs, note) => hs !== null && note !== null && Math.abs(hs - note) < 0.01),
            getComparisonResult('Currency', props.deal_currency_code, noteData.totalDealCurrency, (hs, note) => hs && note && hs.toUpperCase() === note.toUpperCase()),
            getComparisonResult('Payment Type', props.payment_type, noteData.paymentType, (hs, note) => hs && note && hs.toLowerCase().includes(note.toLowerCase())),
            getComparisonResult('Subscription Tenure (Months)', parseInt(props.subscription_tenure), noteData.subscriptionDuration),
            getComparisonResult('Committed Classes', parseInt(props.total_classes_committed_through_learner_s_journey), noteData.committedClasses),
            getComparisonResult('Current Course', getCourseLabel(props.current_course), noteData.courseEnrolled, (hs, note) => hs && note && (hs.toLowerCase().includes(note.toLowerCase()) || note.toLowerCase().includes(hs.toLowerCase()))),
            getComparisonResult('Class Frequency', props.frequency_of_classes, noteData.classFrequency, (hs, note) => hs && note && hs.toLowerCase().includes(note.toLowerCase().replace(" per week",""))),
            getComparisonResult('Time Zone', props.time_zone, noteData.timeZone, (hs, note) => hs && note && hs.toLowerCase().includes(note.toLowerCase())),
            getComparisonResult('Assigned Teacher', getTeacherLabel(props.current_teacher), noteData.teacherQualified),
            getComparisonResult('Age', props.age, null),
            getComparisonResult('Learner Status', props.learner_status, null)
        ];
        Logger.log("Comparison details built successfully.");

        const returnData = { 
            success: true, 
            details: details, 
            calendar: calendarResult, 
            rawNote: noteText || "No relevant sales or onboarding note found." 
        };

        // Return a stringified JSON to prevent Apps Script from converting it, which can cause issues client-side.
        return JSON.stringify(returnData);

    } catch (e) {
        Logger.log(`FATAL ERROR in getAuditDetails for deal ${dealId}: ${e.message}\nStack: ${e.stack}`);
        return JSON.stringify({ success: false, message: `A critical server error occurred: ${e.message}. Check Apps Script logs.` });
    }
}


/**
 * Compares a property from HubSpot and a sales note, providing a detailed status.
 * @param {string} propertyName - The user-friendly name of the property.
 * @param {any} hsValue - The value from the HubSpot deal property.
 * @param {any} noteValue - The value parsed from the sales note.
 * @param {Function} comparisonFn - A function to compare the two values. Should return true if they match.
 * @returns {object} A structured comparison result object.
 */
function getComparisonResult(propertyName, hsValue, noteValue, comparisonFn = (a, b) => String(a) === String(b)) {
    const formatValue = (val) => (val === null || val === undefined) ? 'Not Found' : val;
    const result = { field: propertyName, hsValue: formatValue(hsValue), noteValue: formatValue(noteValue), status: 'Mismatch' };
    
    if (hsValue === null || hsValue === undefined) {
        result.status = 'Warning'; // HubSpot data is missing
        return result;
    }
    if (noteValue === null || noteValue === undefined) {
        result.status = 'Info'; // Note data is missing, can't compare
        return result;
    }

    if (comparisonFn(hsValue, noteValue)) {
        result.status = 'Match';
    }
    return result;
}

/**
 * Main function to run the onboarding audit process.
 * @param {Object} params - Contains fromDate and toDate.
 * @returns {Object} A result object with success status and audit data.
 */
function runOnboardingAudit(params) {
  // CORRECTED: Added a validation block to prevent crashes
  if (!params || params.fromDate === undefined || params.toDate === undefined) {
    const errorMessage = "Audit function was called without the required date range. The 'params' object was not received correctly.";
    Logger.log(`FATAL ERROR in runOnboardingAudit: ${errorMessage}`);
    // Return a proper error object to the client instead of crashing
    return { 
      success: false, 
      message: errorMessage 
    };
  }

  Logger.log(`Running onboarding audit from ${params.fromDate} to ${params.toDate}`);
  try { // MASTER "SAFETY NET" CATCH BLOCK STARTS HERE
    
    const deals = fetchDealsByOnboardingCompletionDate(params.fromDate, params.toDate);
    if (deals.length === 0) {
      return { success: true, data: [], message: 'No deals found with an onboarding completion date in this range.' };
    }

    const auditResults = deals.map(deal => {
      // This inner try-catch ensures that if one deal fails, it doesn't stop the entire audit.
      try {
        const result = {
          dealId: deal.id,
          jlid: deal.properties.jetlearner_id || 'N/A',
          learnerName: deal.properties.dealname,
          onboardingDate: deal.properties.onboarding_completion_date ? new Date(deal.properties.onboarding_completion_date).toLocaleDateString('en-GB') : 'N/A',
          discrepancyCount: 0,
          status: 'Compliant'
        };

        const noteText = fetchLatestSalesNoteForDeal(deal.id);
        if (!noteText) {
          result.status = 'Warning';
          result.discrepancyCount = 1; // Count "no note" as a discrepancy
          return result;
        }
        
        const noteData = parseSalesNote(noteText);
        const props = deal.properties;

        const checks = [
          getComparisonResult('Amount', parseFloat(props.amount), noteData.totalDealAmount, (hs, note) => Math.abs(hs - note) < 1),
          getComparisonResult('Payment Type', props.payment_type, noteData.paymentType, (hs, note) => hs?.toLowerCase() === note?.toLowerCase()),
          getComparisonResult('Subscription Tenure', parseInt(props.subscription_tenure), noteData.subscriptionDuration),
          getComparisonResult('Committed Classes', parseInt(props.total_classes_committed_through_learner_s_journey), noteData.committedClasses),
          getComparisonResult('Current Course', props.current_course, noteData.courseEnrolled, (hs, note) => hs && note && (hs.toLowerCase().includes(note.toLowerCase()) || note.toLowerCase().includes(hs.toLowerCase()))),
        ];
        
        const calendarVerification = verifySubscriptionWithCalendar(props.jetlearner_id, props.module_start_date, props.module_end_date, parseInt(props.total_classes_committed_through_learner_s_journey));
        if(calendarVerification.dateCheck.status === 'Mismatch' || calendarVerification.countCheck.status === 'Mismatch') {
          result.discrepancyCount++;
        }

        result.discrepancyCount += checks.filter(c => c.status === 'Mismatch').length;
        if (checks.some(c => c.status === 'Warning') || calendarVerification.dateCheck.status === 'Warning' || calendarVerification.countCheck.status === 'Warning') {
          result.status = 'Warning';
        }
        if (result.discrepancyCount > 0) {
          result.status = 'Mismatch';
        }

        return result;

      } catch (loopError) {
          Logger.log(`Error processing a single deal (${deal.id || 'Unknown ID'}) inside audit loop: ${loopError.message}`);
          // Return an error object for this specific row so the UI can show it.
          return {
              dealId: deal.id || 'Unknown',
              jlid: deal.properties.jetlearner_id || 'N/A',
              learnerName: deal.properties.dealname || 'Unknown',
              onboardingDate: 'N/A',
              discrepancyCount: 1,
              status: 'Error',
          };
      }
    });

    Logger.log(`Audit completed. Processed ${auditResults.length} deals.`);
    return { success: true, data: auditResults };

  } catch (error) {
    // THIS IS THE CRITICAL PART.
    // It catches any catastrophic error from the functions called above.
    Logger.log(`!!!!!! FATAL ERROR in runOnboardingAudit !!!!!!\nMessage: ${error.message}\nStack: ${error.stack}`);
    
    // It then returns a PROPER error object to the client, preventing the crash.
    return { 
      success: false, 
      message: `A critical server error occurred: "${error.message}". Please check the Apps Script "Executions" log for a "FATAL ERROR" entry to see the full details.` 
    };
  }
}

// ==========================================
// 🟢 WATI CONFIGURATION & MAPPING
// ==========================================

// 1. Map "Migration Reasons" to your specific "WATI Template Names"
// UPDATE THIS LIST with your exact template names from WATI Dashboard
const WATI_REASON_MAPPING = {
  // --- ONE-TO-ONE MAPPINGS (Automatic) ---
  "Teacher on Leaves": [
    { id: "migration_teacher_on_leaves_studies", label: "Higher Studies" },
    { id: "migration_teacher_on_leave_matenity", label: "Maternity Leave" },
    { id: "teacher_leave_migration", label: "Medical/Family Emergency" }
  ],
  "Teacher Change after PRM": [
    { id: "migration_teacher_change_after_prm", label: "Teacher Change Post-PRM" }
  ],
  "Teacher change after PRM": [ // Handling case sensitivity
    { id: "migration_teacher_change_after_prm", label: "Teacher Change Post-PRM" }
  ],
  "Teacher Performance Issue": [
    { id: "migration_teacher_performance_issue", label: "Performance Issue" }
  ],
  "Teacher Affinity": [
    { id: "migration_same_teacher_request", label: "Teacher Affinity" }
  ],
  "Boomerang": [
    { id: "migration_boomerang", label: "Boomerang Return" }
  ],
  "Course change after PRM": [
    { id: "migration_course_change_after_prm", label: "Course Change Post-PRM" }
  ],
  "Slot change - Learner request": [
    { id: "migration_slot_change_lr_request", label: "Slot Change (Learner)" }
  ],
  "Slot change - Teacher request": [
    { id: "migration_releasing_teacher_bandwith_", label: "Teacher Bandwidth/Schedule Adjustment" }
  ],
  "Course Change Teacher Change": [
    { id: "course_change_teacher_change_migration", label: "Course & Teacher Change" }
  ],

  // --- ONE-TO-MANY (Ambiguous - User choice required) ---
  "Attrition": [
    { id: "migration_attrition_tr_change", label: "Standard Attrition" },
    { id: "migration_attrition_tr_promoted_new", label: "Teacher Promoted" }
  ],

  // --- MANY-TO-ONE (Optimization Group) ---
  "GCSE Optimization": [{ id: "migration_releasing_teacher_bandwith_", label: "Optimization / Bandwidth" }],
  "Math Optimization": [{ id: "migration_releasing_teacher_bandwith_", label: "Optimization / Bandwidth" }],
  "Paid BnR Optimization": [{ id: "migration_releasing_teacher_bandwith_", label: "Optimization / Bandwidth" }],
  "Releasing Vintage teacher bandwidth": [{ id: "migration_releasing_teacher_bandwith_", label: "Optimization / Bandwidth" }],

  // --- ESCALATIONS ---
  "Escalation on Teacher": [
    { id: "migration_escalation_on_teacher", label: "Escalation Update" }
  ],
  "Escalation on Teacher Post Migration": [
    { id: "migration_escalation_on_teacher", label: "Escalation Update" }
  ],

  // --- PAUSE / PSW ---
  "Pause Request": [
    { id: "migration_psw", label: "Pause/PSW Message" }
  ],
  "PSW": [
    { id: "migration_psw", label: "Pause/PSW Message" }
  ],

  // --- FALLBACK ---
  "Default": [
    { id: "migration_generic_update", label: "Generic Update" },
    { id: "migration_link_change_infromation", label: "Link Change Only" }
  ]
};


function getWatiParameters(templateName, migrationData, hubspotData) {
  
  // A. Prepare Data Variables
  const session = (migrationData.classSessions && migrationData.classSessions.length > 0) 
                  ? migrationData.classSessions[0] 
                  : { day: "TBD", time: "TBD" };

  let dateStr = "TBD";
  try {
    if (migrationData.startDate) {
        const d = new Date(migrationData.startDate);
        dateStr = `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
    }
  } catch(e) {}

  let timeStr = migrationData.calculatedLocalTime;
  if (!timeStr) {
      const rawTime = session.time || "TBD";
      const tz = migrationData.manualTimezone || hubspotData.timezone || "Europe/London";
      timeStr = convertCetToLocal(rawTime, tz);
  }

  const parentName = hubspotData.parentName || "Parent";
  const learnerName = migrationData.learner || "Student";
  const teacherName = migrationData.newTeacher || "New Teacher"; 
  const oldTeacherName = migrationData.oldTeacher || "Previous Teacher"; 
  const courseName = migrationData.course || "Course";
  const classLink = migrationData.classLink || "https://live.jetlearn.com/login";
  const weekdayStr = session.day || "Day";

  // B. Construct Parameter Array based on Template ID
  let requiredParams = [];

  switch (templateName) {
    
    // ----------------------------------------------------
    // NEW: TEACHER ON LEAVES - HIGHER STUDIES
    // ----------------------------------------------------
    case "migration_teacher_on_leaves_studies":
      requiredParams = [
        { name: "Parent", value: parentName },       // {{Parent}}
        { name: "Learner", value: learnerName },     // {{Learner}}
        { name: "teacher", value: oldTeacherName },  // {{teacher}} (Old Teacher)
        { name: "new_teacher", value: teacherName }, // {{new_teacher}} (New Teacher)
        { name: "Weekday", value: weekdayStr },      // {{Weekday}}
        { name: "Time", value: timeStr },            // {{Time}}
        { name: "Meeting_Link", value: classLink }   // {{Meeting_Link}}
      ];
      break;

    // ----------------------------------------------------
    // NEW: TEACHER ON LEAVES - MATERNITY
    // ----------------------------------------------------
    case "migration_teacher_on_leave_matenity":
      requiredParams = [
        { name: "name", value: parentName },         // {{name}}
        { name: "Learner", value: learnerName },     // {{Learner}}
        { name: "teacher", value: oldTeacherName },  // {{teacher}} (Old Teacher)
        { name: "new_teacher", value: teacherName }, // {{new_teacher}} (New Teacher)
        { name: "Weekday", value: weekdayStr },      // {{Weekday}}
        { name: "Time", value: timeStr },            // {{Time}}
        { name: "Link", value: classLink }           // {{Link}}
      ];
      break;

    // ----------------------------------------------------
    // NEW: TEACHER ON LEAVES - MEDICAL/FAMILY (Utility)
    // ----------------------------------------------------
    case "teacher_leave_migration":
      requiredParams = [
        { name: "name", value: parentName },         // {{name}}
        { name: "Learner", value: learnerName },     // {{Learner}}
        { name: "teacher", value: oldTeacherName },  // {{teacher}} (Old Teacher)
        { name: "new_teacher", value: teacherName }, // {{new_teacher}} (New Teacher)
        { name: "Weekday", value: weekdayStr },      // {{Weekday}}
        { name: "Time", value: timeStr },            // {{Time}}
        { name: "Link", value: classLink }           // {{Link}}
      ];
      break;
    
    // ----------------------------------------------------
    // 1. ATTRITION: STANDARD (Teacher Left)
    // ----------------------------------------------------
    case "migration_attrition_tr_change":
      requiredParams = [
        { name: "Parent", value: parentName },
        { name: "Teacher", value: oldTeacherName }, // "Teacher X is leaving"
        { name: "Learner", value: learnerName },
        { name: "New_Teacher", value: teacherName }, // "New instructor Y"
        { name: "Weekday", value: weekdayStr },
        { name: "Time", value: timeStr },
        { name: "Link", value: classLink }
      ];
      break;

    // ----------------------------------------------------
    // 2. ATTRITION: PROMOTED (Teacher Promoted)
    // ----------------------------------------------------
    case "migration_attrition_tr_promoted_new":
      requiredParams = [
        { name: "Parent", value: parentName },
        { name: "Teacher", value: oldTeacherName }, // "Teacher X promoted"
        { name: "new_teacher", value: teacherName }, // "New teacher Y" (lowercase variable in template)
        { name: "Learner", value: learnerName },
        { name: "Weekday", value: weekdayStr },
        { name: "Time", value: timeStr },
        { name: "Link", value: classLink }
      ];
      break;

    // ----------------------------------------------------
    // 3. ESCALATION
    // ----------------------------------------------------
    case "migration_escalation_on_teacher":
      requiredParams = [
        { name: "parent", value: parentName },
        { name: "teacher", value: teacherName },
        { name: "learner", value: learnerName },
        { name: "coures_type", value: courseName },
        { name: "date", value: dateStr },
        { name: "weekday", value: weekdayStr },
        { name: "time", value: timeStr },
        { name: "link", value: classLink }
      ];
      break;

    // ----------------------------------------------------
    // 4. COURSE CHANGE
    // ----------------------------------------------------
   case "course_change_teacher_change_migration":
      requiredParams = [
        { name: "parent", value: parentName },       // Screenshot shows {{parent}} (lowercase)
        { name: "Student", value: learnerName },     // Screenshot shows {{Student}} (Capitalized)
        { name: "Course", value: courseName },       // Screenshot shows {{Course}} (Capitalized)
        { name: "Teacher", value: teacherName },     // Screenshot shows {{Teacher}} (Capitalized) - WAS MISSING
        { name: "Weekday", value: weekdayStr },      // Screenshot shows {{Weekday}}
        { name: "Time", value: timeStr },            // Screenshot shows {{Time}}
        { name: "Link", value: classLink }           // Screenshot shows {{Link}}
      ];
      break;

    // ----------------------------------------------------
    // 5. BOOMERANG
    // ----------------------------------------------------
    case "migration_boomerang":
      requiredParams = [
        { name: "parent", value: parentName },
        { name: "learner", value: learnerName },
        { name: "new_teacher", value: teacherName },
        { name: "date", value: dateStr }
      ];
      break;

    // ----------------------------------------------------
    // 6. SLOT CHANGE
    // ----------------------------------------------------
    case "migration_slot_change_lr_request":
      requiredParams = [
        { name: "Parent", value: parentName },
        { name: "Teacher", value: teacherName},
        { name: "Learner", value: learnerName },
        { name: "Weekday", value: weekdayStr },
        { name: "Time", value: timeStr },
        { name: "link", value: classLink }
      ];
      break;
      
    // ----------------------------------------------------
    // 7. PSW / PAUSE / OPTIMIZATION
    // ----------------------------------------------------
    case "migration_psw":
    case "migration_releasing_teacher_bandwith_":
      requiredParams = [
        { name: "Parent", value: parentName },       // {{Parent}} (Capital P)
        { name: "Learner", value: learnerName },     // {{Learner}} (Capital L)
        { name: "teacher", value: oldTeacherName },  // {{teacher}} (Old teacher)
        { name: "new_teacher", value: teacherName }, // {{new_teacher}} (New teacher)
        { name: "Weekday", value: weekdayStr },      // {{Weekday}} (Capital W)
        { name: "Time", value: timeStr },            // {{Time}} (Capital T)
        { name: "Link", value: classLink }           // {{Link}} (Capital L)
      ];
      break;
      
    case "migration_link_change_infromation":
      requiredParams = [
        { name: "parent", value: parentName },
        { name: "learner", value: learnerName },
        { name: "link", value: classLink }
      ];
      break;
    //----------------------------------------------------
    // 8. COURSE CHANGE AFTER PRM
    //----------------------------------------------------
    case "migration_course_change_after_prm":
      requiredParams = [
        { name: "Parent", value: parentName },
        { name: "Learner", value: learnerName },
        { name: "Course", value: courseName},
        { name: "teacher", value: teacherName },
        { name: "Weekday", value: weekdayStr },
        { name: "date", value: dateStr },
        { name: "time", value: timeStr },
        { name: "link", value: classLink }
      ];
      break;
      
    // ----------------------------------------------------
    // 9. DEFAULT FALLBACK
    // ----------------------------------------------------
    default:
      requiredParams = [
        { name: "parent", value: parentName },
        { name: "learner", value: learnerName },
        { name: "teacher", value: teacherName },
        { name: "date", value: dateStr },
        { name: "time", value: timeStr },
        { name: "link", value: classLink }
      ];
  }

  return requiredParams;
}

function getWatiPreviewDetails(migrationData) {
  try {
    // 1. Fetch HubSpot Data
    const hubspotResult = fetchHubspotByJlid(migrationData.jlid);
    if (!hubspotResult.success) {
      throw new Error("Could not fetch HubSpot data: " + hubspotResult.message);
    }
    const hsData = hubspotResult.data;

    // 2. Determine Template ID
    let templateId = migrationData.watiTemplateName;
    if (!templateId) {
      const mapping = WATI_REASON_MAPPING[migrationData.reasonOfMigration];
      templateId = (mapping && mapping.length > 0) ? mapping[0].id : "migration_generic_update";
    }

    // 3. Calculate Local Time
    const firstSession = (migrationData.classSessions && migrationData.classSessions.length > 0) 
                         ? migrationData.classSessions[0] : { day: "TBD", time: "TBD" };
    
    const tz = migrationData.manualTimezone || hsData.timezone || "Europe/London";
    migrationData.calculatedLocalTime = convertCetToLocal(firstSession.time, tz);

    // 4. Generate Parameters
    const params = getWatiParameters(templateId, migrationData, hsData);

    // --- NEW: Generate Direct WATI Link ---
    let directLink = null;
    if (hsData.parentContact) {
        // Reuse your existing helper logic to find the link
        const linkResult = fetchWatiDirectLink(hsData.parentContact); 
        if (linkResult.success) {
            directLink = linkResult.link;
        }
    }

    return {
      success: true,
      templateName: templateId,
      parentPhone: hsData.parentContact,
      parameters: params,
      watiLink: directLink // Sending this back to frontend
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getTemplatesForReason(reason) {
  // Normalize input
  const normalizedReason = String(reason).trim();
  // Return specific mapping or default
  return WATI_REASON_MAPPING[normalizedReason] || WATI_REASON_MAPPING["Default"];
}

// =============================================
// ONBOARDING WHATSAPP CONFIGURATION
// =============================================
const WATI_ONBOARDING_CONFIG = {
  WELCOME_TEMPLATE_NAME: "onboarding_link", 
  REMINDER_TEMPLATE_NAME: "onboarding_reminder",  
};

/**
 * Helper: Sends the 2-step Onboarding WhatsApp sequence
 */
function sendOnboardingWhatsAppSequence(data) {
  const results = [];
  
  if (!data.parentContact) return ["Skipped: No Parent Phone"];
  const phone = String(data.parentContact).replace(/\D/g, '');

  let timeSlotStr = data.classTimings || "Check Email for Schedule"; 

  // *** FIX: Format date strictly for WhatsApp here ***
  let whatsAppDate = data.startDate;
  if (typeof formatDateDDMMYYYY === 'function') {
      whatsAppDate = formatDateDDMMYYYY(data.startDate).replace(/\//g, '-');
  }
  // **************************************************

  const welcomeParams = [
    { name: "learner", value: `*${data.learnerName}*` }, 
    { name: "current_teacher", value: `*${data.teacherName}*` }, 
    { name: "timeslot", value: `*${timeSlotStr}*` }, 
    { name: "start_date", value: `*${whatsAppDate}*` }, // Use the formatted variable
    { name: "Zoom_Link", value: data.zoomLink }, 
    { name: "text", value: data.practiceDocumentLink || "Link will be shared shortly" },
    { name: "parent_email", value: `*${data.parentEmail}*` } 
  ];

  try {
    const res1 = sendWatiMessage(phone, WATI_ONBOARDING_CONFIG.WELCOME_TEMPLATE_NAME, welcomeParams);
    results.push(res1.success ? "Msg 1: Sent" : "Msg 1: Failed");
  } catch (e) {
    results.push("Msg 1 Error: " + e.message);
  }

  try {
    Utilities.sleep(1500); 
    const res2 = sendWatiMessage(phone, WATI_ONBOARDING_CONFIG.REMINDER_TEMPLATE_NAME, []);
    results.push(res2.success ? "Msg 2: Sent" : "Msg 2: Failed");
  } catch (e) {
    results.push("Msg 2 Error: " + e.message);
  }

  return results;
}

/**
 * Sends the actual message to WATI
 */
function sendWatiMessage(phoneNumber, templateName, parameters) {
  const scriptProperties = PropertiesService.getScriptProperties();
  let API_ENDPOINT_BASE = (scriptProperties.getProperty('WATI_API_ENDPOINT') || "").trim();
  let ACCESS_TOKEN = (scriptProperties.getProperty('WATI_ACCESS_TOKEN') || "").trim();

  if (!API_ENDPOINT_BASE || !ACCESS_TOKEN) {
    throw new Error("WATI Configuration Error: Missing Script Properties.");
  }

  // 1. Ensure Bearer prefix is present
  if (!ACCESS_TOKEN.startsWith("Bearer ")) {
    ACCESS_TOKEN = "Bearer " + ACCESS_TOKEN;
  }

  // 2. Clean Base URL
  if (API_ENDPOINT_BASE.endsWith("/")) {
    API_ENDPOINT_BASE = API_ENDPOINT_BASE.slice(0, -1);
  }

  // 3. Sanitize Phone Number
  const cleanPhone = String(phoneNumber).replace(/\D/g, '');

  // 4. Construct URL (Singular Endpoint + Query Parameter)
  // NOTE: For the singular endpoint, the phone number MUST be in the URL.
  const FULL_API_ENDPOINT = `${API_ENDPOINT_BASE}/api/v1/sendTemplateMessage?whatsappNumber=${cleanPhone}`;

  // 5. Construct Payload (No 'template_messages' array, simple object)
  const payload = {
    "template_name": templateName,
    "broadcast_name": "JetLearn_Notification",
    "parameters": parameters
  };

  const options = {
    "method": "post",
    "headers": {
      "Authorization": ACCESS_TOKEN,
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    Logger.log(`[WATI] Sending to: ${FULL_API_ENDPOINT}`);
    Logger.log(`[WATI] Payload: ${JSON.stringify(payload)}`);

    const response = UrlFetchApp.fetch(FULL_API_ENDPOINT, options);
    const responseCode = response.getResponseCode();
    const content = response.getContentText();

    Logger.log(`[WATI] Response Code: ${responseCode}`);
    
    // Check for success (200)
    if (responseCode !== 200) {
      Logger.log(`[WATI] Error Body: ${content}`);
      // Common WATI error: 400 usually means template name mismatch or param mismatch
      throw new Error(`WATI API Error (${responseCode}): ${content || 'No Content returned'}`);
    }

    const result = JSON.parse(content);

    // WATI sometimes returns 200 but with result: false
    if (result.result === false || result.status === 'error') {
       const detail = (result.messages && result.messages.length > 0) ? result.messages[0].message : JSON.stringify(result);
       throw new Error(`WATI Application Error: ${detail}`);
    }

    return { success: true, result: result };

  } catch (e) {
    Logger.log(`[WATI] Exception: ${e.message}`);
    throw e; 
  }
}

// 3. Update your main Migration Handler to support the "Both" logic
function processMigrationSubmission(data, sendEmail, sendWhatsapp) {
  const results = { email: 'Skipped', whatsapp: 'Skipped' };
  
  // 1. Send Email (Teacher)
  if (sendEmail) {
     try {
      const dealData = fetchHubspotByJlid(data.jlid);
      const parentPhone = dealData.data.parentContact;
       sendMigrationEmail(data); 
       results.email = 'Success';
     } catch(e) {
       results.email = 'Failed: ' + e.message;
     }
  }

  // 2. Send WhatsApp (Parent)
  if (sendWhatsapp) {
    try {
      // Re-fetch parent phone from Deal to be safe
      const dealData = fetchHubspotByJlid(data.jlid);
      const parentPhone = dealData.data.parentContact;
      const parentName = dealData.data.parentName; // Get parent name from deal
      
      if(parentPhone) {
        // Inject fetched parent data into `data` object for the generator
        data.parentPhone = parentPhone;
        data.parentName = dealData.data.parentName; 

        const watiConfig = generateWatiPreview(data);
        const watiResult = sendWatiMessage(parentPhone, watiConfig.templateName, watiConfig.parameters);
        
        results.whatsapp = watiResult.success ? 'Success' : 'Failed: ' + watiResult.message;
      } else {
        results.whatsapp = 'Failed: No Parent Phone in HubSpot';
      }
    } catch(e) {
      results.whatsapp = 'Failed: ' + e.message;
    }
  }

  // Log action to sheet
  logAction(
    'Migration', data.jlid, data.learner, data.oldTeacher, data.newTeacher, data.course, 
    'Completed', `Email: ${results.email}, WhatsApp: ${results.whatsapp}`
  );

  return { 
    success: true, 
    message: `Email: ${results.email} | WhatsApp: ${results.whatsapp}` 
  };
}

function generateWatiPreviewWrapper(data) {
  try {
    // 1. Determine the Template ID
    // If the user selected a specific template in the dropdown, use it.
    // Otherwise, default to the first template mapped to this Reason.
    let templateId = data.watiTemplateName;
    
    if (!templateId) {
        const mapping = WATI_REASON_MAPPING[data.reasonOfMigration] || WATI_REASON_MAPPING["Default"];
        templateId = mapping[0].id;
    }

    // 2. Prepare Mock HubSpot Data for Preview
    // (We might not have the full HS fetch result here, so we use placeholders or passed data)
    const mockHsData = {
        parentName: data.parentName || "Parent Name",
        // Pass other known fields if available in 'data'
    };

    // 3. Generate Parameters using the NEW central function
    const calculatedParams = getWatiParameters(templateId, data, mockHsData);

    // 4. Build HTML List
    let varsListHtml = calculatedParams.map(p => `<li><strong>{{${p.name}}}</strong>: ${p.value}</li>`).join("");

    const htmlPreview = `
      <div style="font-family: Helvetica, Arial, sans-serif; background-color: #E5DDD5; padding: 20px; border-radius: 8px;">
        <div style="background-color: white; padding: 10px 15px; border-radius: 8px; box-shadow: 0 1px 1px rgba(0,0,0,0.1); max-width: 400px;">
          <div style="color: #075E54; font-weight: bold; font-size: 13px; margin-bottom: 5px;">Template: ${templateId}</div>
          <div style="color: #333; font-size: 14px; line-height: 1.5;">
            <p><em>(Template text is stored in WATI. We will inject these values:)</em></p>
            <ul style="padding-left: 20px; margin: 5px 0;">${varsListHtml}</ul>
          </div>
        </div>
      </div>
    `;

    return {
      success: true,
      html: htmlPreview,
      templateName: templateId,
      parameters: calculatedParams
    };

  } catch (e) {
    Logger.log("Preview Error: " + e.message);
    return { success: false, message: "Preview Error: " + e.message };
  }
}


/**
 * Converts a CET time string to the Learner's Local Time.
 * Uses smart detection to support ALL timezones including MST, ADT, ACST, etc.
 * @param {string} timeStr - Time in "HH:mm" or "HH:mm AM/PM" format (assumed CET).
 * @param {string} targetTzString - The timezone string from HubSpot (e.g., "(GMT-07:00) Mountain Time...").
 * @returns {string} Formatted local time string (e.g., "5:30 PM MST").
 */
function convertCetToLocal(timeStr, targetTzString) {
  try {
    if (!timeStr || timeStr === "TBD") return "TBD";

    // 1. Parse Input Time (Assumed CET)
    let hours = 0, minutes = 0;
    const timeMatch = timeStr.match(/(\d+):(\d+)\s*(AM|PM)?/i);
    
    if (timeMatch) {
      hours = parseInt(timeMatch[1]);
      minutes = parseInt(timeMatch[2]);
      const meridiem = timeMatch[3] ? timeMatch[3].toUpperCase() : null;
      
      // Convert to 24h format
      if (meridiem === 'PM' && hours < 12) hours += 12;
      if (meridiem === 'AM' && hours === 12) hours = 0;
    } else {
      return timeStr + " (CET)"; // Fallback
    }

    // 2. Identify Target IANA Timezone dynamically
    // Default fallback
    let ianaZone = "Europe/Berlin"; 

    if (targetTzString) {
      const tz = targetTzString; 

      // --- SPECIAL OFFSETS & NO-DST REGIONS (Check these first!) ---
      if (tz.includes("Newfoundland")) ianaZone = "America/St_Johns"; // -03:30
      else if (tz.includes("Adelaide")) ianaZone = "Australia/Adelaide"; // +09:30
      else if (tz.includes("Darwin")) ianaZone = "Australia/Darwin"; // +09:30 (No DST)
      else if (tz.includes("Kathmandu")) ianaZone = "Asia/Kathmandu"; // +05:45
      else if (tz.includes("Yangon") || tz.includes("Rangoon")) ianaZone = "Asia/Yangon"; // +06:30
      else if (tz.includes("Saskatchewan")) ianaZone = "America/Regina"; // -06:00 (No DST)
      else if (tz.includes("Arizona")) ianaZone = "America/Phoenix"; // -07:00 (No DST)
      else if (tz.includes("Brisbane")) ianaZone = "Australia/Brisbane"; // +10:00 (No DST)
      else if (tz.includes("Central America")) ianaZone = "America/Guatemala"; // -06:00 (No DST)
      else if (tz.includes("Hawaii")) ianaZone = "Pacific/Honolulu"; // -10:00 (No DST)
      else if (tz.includes("International Date Line West")) ianaZone = "Etc/GMT+12";

      // --- NORTH & SOUTH AMERICA ---
      else if (tz.includes("Eastern")) ianaZone = "America/New_York"; // EST/EDT
      else if (tz.includes("Central Time")) ianaZone = "America/Chicago"; // CST/CDT
      else if (tz.includes("Mountain")) ianaZone = "America/Denver"; // MST/MDT
      else if (tz.includes("Pacific")) ianaZone = "America/Los_Angeles"; // PST/PDT
      else if (tz.includes("Alaska")) ianaZone = "America/Anchorage"; // AKST/AKDT
      else if (tz.includes("Atlantic")) ianaZone = "America/Halifax"; // AST/ADT
      else if (tz.includes("Indiana")) ianaZone = "America/Indiana/Indianapolis";
      else if (tz.includes("Bogota") || tz.includes("Lima") || tz.includes("Quito")) ianaZone = "America/Bogota";
      else if (tz.includes("Caracas")) ianaZone = "America/Caracas";
      else if (tz.includes("Santiago")) ianaZone = "America/Santiago";
      else if (tz.includes("Brasilia")) ianaZone = "America/Sao_Paulo";
      else if (tz.includes("Buenos Aires")) ianaZone = "America/Argentina/Buenos_Aires";
      else if (tz.includes("Georgetown")) ianaZone = "America/Manaus";
      else if (tz.includes("Mexico City")) ianaZone = "America/Mexico_City";

      // --- EUROPE & AFRICA ---
      else if (tz.includes("London") || tz.includes("Dublin") || tz.includes("Edinburgh")) ianaZone = "Europe/London"; // GMT/BST
      else if (tz.includes("Amsterdam") || tz.includes("Berlin") || tz.includes("Rome") || tz.includes("Paris") || tz.includes("Stockholm")) ianaZone = "Europe/Berlin"; // CET/CEST
      else if (tz.includes("Athens") || tz.includes("Bucharest") || tz.includes("Cairo")) ianaZone = "Europe/Athens"; // EET/EEST
      else if (tz.includes("Moscow")) ianaZone = "Europe/Moscow";
      else if (tz.includes("Johannesburg") || tz.includes("Pretoria") || tz.includes("Harare")) ianaZone = "Africa/Johannesburg";
      else if (tz.includes("Casablanca")) ianaZone = "Africa/Casablanca";
      else if (tz.includes("Nairobi")) ianaZone = "Africa/Nairobi";

      // --- ASIA ---
      else if (tz.includes("India") || tz.includes("Kolkata") || tz.includes("New Delhi") || tz.includes("Chennai")) ianaZone = "Asia/Kolkata"; // IST
      else if (tz.includes("Islamabad") || tz.includes("Karachi")) ianaZone = "Asia/Karachi";
      else if (tz.includes("Dubai") || tz.includes("Muscat") || tz.includes("Abu Dhabi")) ianaZone = "Asia/Dubai";
      else if (tz.includes("Singapore") || tz.includes("Kuala Lumpur")) ianaZone = "Asia/Singapore";
      else if (tz.includes("Hong Kong") || tz.includes("Beijing")) ianaZone = "Asia/Hong_Kong";
      else if (tz.includes("Tokyo") || tz.includes("Osaka") || tz.includes("Sapporo")) ianaZone = "Asia/Tokyo";
      else if (tz.includes("Seoul")) ianaZone = "Asia/Seoul";
      else if (tz.includes("Bangkok") || tz.includes("Hanoi") || tz.includes("Jakarta")) ianaZone = "Asia/Bangkok";
      
      // --- AUSTRALIA / PACIFIC ---
      else if (tz.includes("Sydney") || tz.includes("Melbourne") || tz.includes("Canberra")) ianaZone = "Australia/Sydney"; // AEST/AEDT
      else if (tz.includes("Perth")) ianaZone = "Australia/Perth";
      else if (tz.includes("Auckland") || tz.includes("Wellington")) ianaZone = "Pacific/Auckland";
    }

    // 3. Create a CET Date Object
    // We strictly define the "source" time as Europe/Berlin
    const now = new Date();
    const cetString = now.toLocaleString("en-US", {timeZone: "Europe/Berlin"});
    const cetDate = new Date(cetString);
    
    // Set the specific class time onto that CET date
    cetDate.setHours(hours);
    cetDate.setMinutes(minutes);
    cetDate.setSeconds(0);

    // 4. Convert to Target Zone
    const options = {
        hour: 'numeric',
        minute: '2-digit',
        hour12: true,
        timeZone: ianaZone,
        timeZoneName: 'short' 
    };

    // This handles all the math, including Daylight Savings automatically
    const targetTimeStr = cetDate.toLocaleString("en-US", options);
    
    return targetTimeStr;

  } catch (e) {
    Logger.log("Time Conversion Error: " + e.message);
    return timeStr + " (CET)";
  }
}


/**
 * Fetches the WATI Contact ID for a phone number to generate a direct chat link.
 */
function fetchWatiDirectLink(phoneNumber) {
  const scriptProperties = PropertiesService.getScriptProperties();
  let API_ENDPOINT_BASE = (scriptProperties.getProperty('WATI_API_ENDPOINT') || "").trim();
  let ACCESS_TOKEN = (scriptProperties.getProperty('WATI_ACCESS_TOKEN') || "").trim();

  if (!API_ENDPOINT_BASE || !ACCESS_TOKEN) return { success: false, message: "Config missing" };

  if (!ACCESS_TOKEN.startsWith("Bearer ")) ACCESS_TOKEN = "Bearer " + ACCESS_TOKEN;
  if (API_ENDPOINT_BASE.endsWith("/")) API_ENDPOINT_BASE = API_ENDPOINT_BASE.slice(0, -1);

  const tenantId = API_ENDPOINT_BASE.split('/').pop();
  const cleanPhone = String(phoneNumber).replace(/\D/g, '');
  
  const options = {
    "headers": { "Authorization": ACCESS_TOKEN, "Content-Type": "application/json" },
    "muteHttpExceptions": true
  };

  try {
    // 1. Ensure Contact Exists (addContact)
    // We still need this to handle the "New Contact" case gracefully and get a fallback ID
    const addContactUrl = `${API_ENDPOINT_BASE}/api/v1/addContact/${cleanPhone}`;
    const contactRes = UrlFetchApp.fetch(addContactUrl, { ...options, "method": "post", "payload": JSON.stringify({ "name": "Unknown Parent" }) });
    
    let contactId = null;
    // Extract contact ID just in case we don't find a conversation history
    if (contactRes.getResponseCode() === 200 || contactRes.getResponseCode() === 201) {
        const cData = JSON.parse(contactRes.getContentText());
        contactId = cData.contact?.id || cData.id;
    }

    // 2. Fetch Messages to get the actual Conversation ID (Chat Thread)
    const messagesUrl = `${API_ENDPOINT_BASE}/api/v1/getMessages/${cleanPhone}?pageSize=1`;
    const msgRes = UrlFetchApp.fetch(messagesUrl, { ...options, "method": "get" });
    
    let finalId = null;

    if (msgRes.getResponseCode() === 200) {
      const msgData = JSON.parse(msgRes.getContentText());
      if (msgData.messages && msgData.messages.items && msgData.messages.items.length > 0) {
        const latestMsg = msgData.messages.items[0];
        
        // *** THE FIX: Use conversationId based on your logs ***
        finalId = latestMsg.conversationId || latestMsg.ticketId;
      }
    }

    // 3. Fallback to Contact ID if no conversation history exists yet
    if (!finalId) {
        finalId = contactId;
    }

    if (finalId) {
      const directLink = `https://live.wati.io/${tenantId}/teamInbox/${finalId}`;
      return { success: true, link: directLink };
    }
    
    return { success: false, message: "Could not find Conversation ID." };

  } catch (e) {
    Logger.log("WATI Link Error: " + e.message);
    return { success: false, message: e.message };
  }
}

function testGemini25() {
  const key = PropertiesService.getScriptProperties().getProperty('GOOGLE_GENERATIVE_AI_KEY');
  const model = 'gemini-2.5-flash';
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`;
  const payload = {
    contents: [{ parts: [{ text: "Generate a one-line test response" }] }]
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  Logger.log(res.getResponseCode());
  Logger.log(res.getContentText());
}

function debugTicketSorting() {
  const TEST_JLID = 'JL11834214568C'; // The JLID from your screenshot
  
  Logger.log(`🔍 Debugging Ticket Sort for: ${TEST_JLID}`);
  const result = fetchLatestMigrationTicket(TEST_JLID);
  
  Logger.log("--- FINAL RESULT ---");
  Logger.log(JSON.stringify(result, null, 2));
}

function debugWatiStructure() {
  const TEST_PHONE = "918583831888"; 
  const scriptProperties = PropertiesService.getScriptProperties();
  let API_ENDPOINT_BASE = (scriptProperties.getProperty('WATI_API_ENDPOINT') || "").trim();
  let ACCESS_TOKEN = (scriptProperties.getProperty('WATI_ACCESS_TOKEN') || "").trim();

  if (!ACCESS_TOKEN.startsWith("Bearer ")) ACCESS_TOKEN = "Bearer " + ACCESS_TOKEN;
  if (API_ENDPOINT_BASE.endsWith("/")) API_ENDPOINT_BASE = API_ENDPOINT_BASE.slice(0, -1);

  // Fetch 1 message
  const url = `${API_ENDPOINT_BASE}/api/v1/getMessages/${TEST_PHONE}?pageSize=1`;
  
  Logger.log(`Fetching Messages from: ${url}`);

  const options = {
    "method": "get",
    "headers": { "Authorization": ACCESS_TOKEN },
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(url, options);
  const content = response.getContentText();
  
  try {
    const json = JSON.parse(content);
    if (json.messages && json.messages.items && json.messages.items.length > 0) {
      // Log the KEYS of the message object to find the ID
      const msg = json.messages.items[0];
      Logger.log("--- MESSAGE OBJECT KEYS ---");
      Logger.log(Object.keys(msg));
      
      Logger.log("--- LOOKING FOR '115' ID ---");
      // Search values for the specific ID ending in 115
      for (const [key, value] of Object.entries(msg)) {
        Logger.log(`${key}: ${value}`);
      }
    } else {
      Logger.log("No messages found in response.");
    }
  } catch (e) {
    Logger.log("JSON Parse Error: " + e.message);
  }
}