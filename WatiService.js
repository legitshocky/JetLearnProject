function getWatiParameters(templateName, migrationData, hubspotData) {
    let weekdayStr = "TBD";
  let timeStr = "TBD";
  if (migrationData.classSessions && migrationData.classSessions.length > 0) {
    weekdayStr = migrationData.classSessions.map(s => s.day).join(' & ');
    const parentTimezone = migrationData.manualTimezone || hubspotData.timezone || "Europe/London";    
    timeStr = migrationData.classSessions.map(s => {
       return convertCetToLocal(s.time, parentTimezone);
    }).join(' & ');
  }
  let dateStr = "TBD";
  try {
    if (migrationData.startDate) {
        const d = new Date(migrationData.startDate);
        dateStr = `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
    }
  } catch(e) {}

  const parentName = hubspotData.parentName || "Parent";
  const learnerName = migrationData.learner || "Student";
  const teacherName = migrationData.newTeacher || "New Teacher"; 
  const oldTeacherName = migrationData.oldTeacher || "Previous Teacher"; 
  const courseName = migrationData.course || "Course";
  const classLink = migrationData.classLink || "https://live.jetlearn.com/login";
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
        { name: "Parent", value: parentName },
        { name: "Learner", value: learnerName },
        { name: "Teacher", value: teacherName },     // New Teacher
        { name: "Course", value: courseName },       // Added based on screenshot
        { name: "Date", value: dateStr },
        { name: "Weekday", value: weekdayStr },      // Added based on screenshot
        { name: "Time", value: timeStr },            // Added based on screenshot
        { name: "Link", value: classLink }           // Added based on screenshot
      ];
      break;


    // ----------------------------------------------------
    // 6. SLOT CHANGE
    // ----------------------------------------------------
    case "migration_slot_change_lr_request":
      requiredParams = [
        { name: "Parent", value: parentName },
        { name: "Learner", value: learnerName },
        { name: "Teacher", value: teacherName},
        { name: "Date", value: dateStr },        // <--- CRITICAL: MUST BE HERE
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

    //----------------------------------------------------
    // 9. TEACHER PERFORMANCE ISSUE
    //----------------------------------------------------
      
    case "migration_teacher_performance_issue":
      requiredParams = [
        { name: "Parent", value: parentName },       // Was 'parent'
        { name: "Learner", value: learnerName },     // Was 'learner'
        { name: "Teacher", value: teacherName },     // Was 'teacher'
        { name: "course_type", value: courseName },  // Kept as is, assuming template uses snake_case here?
        { name: "Date", value: dateStr },            // Was 'date'
        { name: "Weekday", value: weekdayStr },
        { name: "Time", value: timeStr },            // Was 'time'
        { name: "Link", value: classLink }           // Was 'link'
      ];
      break;

    //----------------------------------------------------
    // 10. TEACHER CHANGE AFTER PRM
    //----------------------------------------------------  

    case "migration_teacher_change_after_prm":
      requiredParams = [
        { name: "Parent", value: parentName },       // Was 'parent'
        { name: "Learner", value: learnerName },     // Was 'learner'
        { name: "Teacher", value: teacherName },     // Was 'teacher'
        { name: "Course", value: courseName},
        { name: "Weekday", value: weekdayStr },
        { name: "Time", value: timeStr },            // Was 'time'
        { name: "Link", value: classLink }           // Was 'link'
      ];
      break;

      case "migration_teacher_affinity":
        requiredParams = [
        { name: "Parent", value: parentName },       // Was 'parent'
        { name: "learner", value: learnerName },     // Was 'learner'
        { name: "teacher", value: teacherName },     // Was 'teacher'
        { name: "Weekday", value: weekdayStr },
        { name: "Time", value: timeStr },            // Was 'time'
        { name: "Meeting_Link", value: classLink }           // Was 'link'
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

function normalizeReason(str) {
  if (!str) return "";
  return String(str).toLowerCase().replace(/[^a-z0-9]/g, "");
}


function getTemplatesForReason(reason) {
  // LOAD CONFIG FROM SHEET
  const dynamicMapping = getWatiConfigFromSheet(); 
  const defaultTemplates = dynamicMapping["Default"] || [];
  
  if (!reason) return defaultTemplates;

  // 1. Try Exact Match
  if (dynamicMapping[reason]) {
    return dynamicMapping[reason];
  }

  // 2. Try Normalized Match (Fuzzy)
  const targetNormalized = normalizeReason(reason);
  const configKeys = Object.keys(dynamicMapping);
  
  for (let i = 0; i < configKeys.length; i++) {
    const key = configKeys[i];
    if (normalizeReason(key) === targetNormalized) {
      return dynamicMapping[key];
    }
  }

  // 3. Fallback
  return defaultTemplates;
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

  if (!ACCESS_TOKEN.startsWith("Bearer ")) ACCESS_TOKEN = "Bearer " + ACCESS_TOKEN;
  if (API_ENDPOINT_BASE.endsWith("/")) API_ENDPOINT_BASE = API_ENDPOINT_BASE.slice(0, -1);

  const cleanPhone = String(phoneNumber).replace(/\D/g, '');
  const FULL_API_ENDPOINT = `${API_ENDPOINT_BASE}/api/v1/sendTemplateMessage?whatsappNumber=${cleanPhone}`;

  // 1. Sanitize Parameters (Prevent empty strings)
  const safeParameters = parameters.map(p => ({
    name: p.name,
    value: (p.value === null || p.value === undefined || String(p.value).trim() === '') ? "N/A" : String(p.value)
  }));

  const payload = {
    "template_name": templateName,
    "broadcast_name": "JetLearn_Notification",
    "parameters": safeParameters
  };

  const options = {
    "method": "post",
    "headers": { "Authorization": ACCESS_TOKEN, "Content-Type": "application/json" },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    Logger.log(`[WATI] Sending to: ${cleanPhone}`);
    Logger.log(`[WATI PAYLOAD] ${JSON.stringify(payload)}`);
    
    const response = UrlFetchApp.fetch(FULL_API_ENDPOINT, options);
    const responseCode = response.getResponseCode();
    const content = response.getContentText();

    // ============================================================
    // 👇 THIS IS THE CRITICAL FIX YOU MIGHT BE MISSING 👇
    // ============================================================
    if (content.trim().startsWith("<")) {
      Logger.log(`[WATI CRITICAL ERROR] Server returned HTML: \n${content}`);
      throw new Error(`WATI API returned HTML error (Code ${responseCode}). Check Executions log for details.`);
    }
    // ============================================================

    const result = JSON.parse(content);

    if (result.result === false || result.status === 'error') {
       const detail = (result.messages && result.messages.length > 0) ? result.messages[0].message : (result.info || JSON.stringify(result));
       throw new Error(`WATI Rejected: ${detail}`);
    }

    return { success: true, result: result };

  } catch (e) {
    Logger.log(`[WATI] Exception: ${e.message}`);
    throw e;
  }
}

// ── Send free-text session message (within 24hr window after parent reply) ──
function sendWatiSessionMessage(phoneNumber, messageText) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var API_ENDPOINT_BASE = (scriptProperties.getProperty('WATI_API_ENDPOINT') || '').trim();
  var ACCESS_TOKEN      = (scriptProperties.getProperty('WATI_ACCESS_TOKEN')  || '').trim();

  if (!API_ENDPOINT_BASE || !ACCESS_TOKEN) {
    Logger.log('[WATI] sendWatiSessionMessage: missing config.');
    return { success: false, message: 'Missing WATI config' };
  }
  if (!ACCESS_TOKEN.startsWith('Bearer ')) ACCESS_TOKEN = 'Bearer ' + ACCESS_TOKEN;
  if (API_ENDPOINT_BASE.endsWith('/')) API_ENDPOINT_BASE = API_ENDPOINT_BASE.slice(0, -1);

  var cleanPhone = String(phoneNumber).replace(/\D/g, '');
  var url = API_ENDPOINT_BASE + '/api/v1/sendSessionMessage/' + cleanPhone;

  try {
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: { 'Authorization': ACCESS_TOKEN, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ messageText: messageText }),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    var body = resp.getContentText();
    Logger.log('[WATI] sendWatiSessionMessage to ' + cleanPhone + ' → HTTP ' + code);
    if (code === 200) return { success: true };
    return { success: false, message: 'HTTP ' + code + ': ' + body.substring(0, 200) };
  } catch (e) {
    Logger.log('[WATI] sendWatiSessionMessage ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

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
    { id: "migration_teacher_affinity", label: "Teacher Affinity" }
  ],
  "Boomerang": [
    { id: "migration_boomerang", label: "Boomerang Return" }
  ],
  "Course change after PRM": [
    { id: "migration_course_change_after_prm", label: "Course Change Post-PRM" }
  ],
  "Slot change -Learner request": [
    { id: "migration_slot_change_lr_request", label: "Slot Change (Learner)" }
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

function getWatiConfigFromSheet() {
  const cacheKey = 'WATI_CONFIG_MAP';
  
  // 1. Try global cache first (fastest)
  if (typeof _sheetDataCache !== 'undefined' && _sheetDataCache[cacheKey]) {
    return _sheetDataCache[cacheKey];
  }

  try {
    const SETTINGS_SHEET_NAME = 'Settings_Wati'; 
    // Uses your existing helper to get data
    const sheetData = _getCachedSheetData(SETTINGS_SHEET_NAME);

    // If sheet is missing or empty, fallback to hardcoded default
    if (!sheetData || sheetData.length < 2) {
      Logger.log("Settings_Wati sheet not found or empty. Using default fallback.");
      return WATI_REASON_MAPPING; 
    }

    const mapping = {};

    // 2. Loop through rows (skip header)
    for (let i = 1; i < sheetData.length; i++) {
      const row = sheetData[i];
      const reason = String(row[0]).trim();     // Col A
      const templateId = String(row[1]).trim(); // Col B
      const label = String(row[2]).trim();      // Col C

      if (reason && templateId) {
        if (!mapping[reason]) mapping[reason] = [];
        mapping[reason].push({ id: templateId, label: label || templateId });
      }
    }

    // 3. Cache the result
    if(typeof _sheetDataCache !== 'undefined') {
        _sheetDataCache[cacheKey] = mapping;
    }
    
    return mapping;

  } catch (e) {
    Logger.log("Error loading WATI config: " + e.message);
    return WATI_REASON_MAPPING; // Emergency fallback
  }
}

function sendWatiTemplate(parentContact, templateId, params, manualOverridePhone = null) {
  // 1. Determine the target phone
  // If manual override exists, use it. Otherwise use the HubSpot contact.
  let targetPhone = manualOverridePhone || parentContact;

  if (!targetPhone) return { success: false, message: "Skipped: No Phone Number Found" };
  
  const phone = String(targetPhone).replace(/\D/g, '');
  
  try {
    const res = sendWatiMessage(phone, templateId, params);
    return { success: res.success, message: "Sent to " + phone };
  } catch (e) {
    return { success: false, message: "Error: " + e.message };
  }
}


