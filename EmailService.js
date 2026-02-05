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

    // Call InvoiceService logic (Global scope allows this)
    const pricingDetails = calculateInvoicePricing(formData);
    formData.sessions = pricingDetails.displayTotalSessions;

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
        // Call WatiService logic
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
    // --- UPDATED ERROR LOGGING ---
    logError('EmailService: sendParentOnboardingWithInvoice', error);
    
    logNotes = `Failed during onboarding process: ${error.message}.`;
    return { success: false, message: `Failed to complete onboarding: ${error.message}` };

  } finally {
    logAction('New Learner Onboarded', formData.jlid, formData.learnerName, '', formData.teacherName, formData.course, overallStatus, logNotes);
  }
}

function sendTrackedEmail(payload) {
  const { to, cc, subject, htmlBody, jlid, attachments } = payload;
  const trackingId = Utilities.getUuid();

  try {
    if (!to || !subject || !htmlBody) {
      throw new Error("Missing required fields: 'to', 'subject', and 'htmlBody' are mandatory.");
    }

    const webAppUrl = ScriptApp.getService().getUrl();
    const trackingPixelUrl = `${webAppUrl}?page=track&id=${trackingId}`;
    const trackingPixelImg = `<img src="${trackingPixelUrl}" width="1" height="1" style="display:none;"/>`;
    const finalHtmlBody = htmlBody.replace('</body>', `${trackingPixelImg}</body>`);

    const mailOptions = {
      to: to,
      subject: subject,
      htmlBody: finalHtmlBody,
      name: CONFIG.EMAIL.FROM_NAME,
      from: CONFIG.EMAIL.FROM
    };
    if (cc) mailOptions.cc = cc;
    if (attachments) mailOptions.attachments = attachments;

    MailApp.sendEmail(mailOptions);

    logEmail({
      trackingId: trackingId,
      recipient: to,
      subject: subject,
      jlid: jlid,
      status: 'Sent',
      sentAt: new Date(),
      htmlBody: finalHtmlBody,
      attachments: attachments
    });

    Logger.log(`Successfully sent and logged tracked email to ${to} with ID ${trackingId}.`);
    return { success: true, trackingId: trackingId };

  } catch (error) {
    Logger.log(`Failed to send tracked email to ${to}. Error: ${error.message}`);
    logEmail({
      trackingId: trackingId,
      recipient: to,
      subject: `FAILED: ${subject}`,
      jlid: jlid,
      status: 'Failed',
      sentAt: new Date(),
      htmlBody: htmlBody,
      rawPayload: error.message
    });
    throw error; // Re-throw the error so the calling function knows it failed
  }
}

function sendOnboardingEmail(data, attachments = []) {
     const validation = validateInput(data, ['jlid', 'learnerName', 'teacherName', 'course', 'clsManager']);
    if (!validation.isValid) {
      throw new Error(validation.message);
    }

  try {
    const teacherData = getTeacherData();
    
    // FIXED: Case-insensitive and trimmed comparison
    const teacherInfo = teacherData.find(t => 
      String(t.name).trim().toLowerCase() === String(data.teacherName).trim().toLowerCase()
    );

    if (!teacherInfo || !isValidEmail(teacherInfo.email)) {
      throw new Error(`Teacher '${data.teacherName}' not found or has invalid email.`);
    }

    const clsManagerEmail = findClsEmailByManagerName(data.clsManager);
    const ccList = new Set();
    if(clsManagerEmail) ccList.add(clsManagerEmail);
    if(teacherInfo.tpManagerEmail) ccList.add(teacherInfo.tpManagerEmail);

    const subject = `New Learner Onboarded || ${data.learnerName} (${data.jlid || 'N/A'})`;
    const htmlBody = getOnboardingEmailHTML(data);
    
    // Use the central tracked email service
    const result = sendTrackedEmail({
      to: teacherInfo.email,
      cc: Array.from(ccList).join(','),
      subject: subject,
      htmlBody: htmlBody,
      jlid: data.jlid,
      attachments: attachments
    });
    
    // NOTE: We no longer call logAction here. The parent function does it.
    
    return { success: true, message: `Onboarding email sent to teacher. TID: ${result.trackingId}` };

  } catch (error) {
    Logger.log(`Error in helper sendOnboardingEmail: ${error.message}`);
    // Re-throw the error so the main function knows it failed
    throw new Error(`Failed to send teacher onboarding email: ${error.message}`);
  }
}


function sendMigrationEmail(data, attachments = []) {
  const requiredFields = ['jlid', 'newTeacher', 'course', 'reasonOfMigration'];
  const validationError = validateRequiredFields(data, requiredFields);

  if (validationError) {
    Logger.log(`Migration Email Aborted: ${validationError}`);
    return { success: false, message: validationError };
  }

  let finalStatus = 'Partial Success';
  let notes = [];
  let watiSuccess = true; 

  try {
    // ==========================================
    // 1. SEND EMAILS (TEACHERS)
    // ==========================================
    if (data.sendEmailToTeacher) {
        const teacherData = getTeacherData();
        const newTeacherInfo = teacherData.find(t => String(t.name).trim().toLowerCase() === String(data.newTeacher).trim().toLowerCase());
        
        let oldTeacherInfo = null;
        if (data.oldTeacher) {
            oldTeacherInfo = teacherData.find(t => String(t.name).trim().toLowerCase() === String(data.oldTeacher).trim().toLowerCase());
        }

        // --- A. NEW TEACHER ---
        if (!newTeacherInfo || !isValidEmail(newTeacherInfo.email)) {
            throw new Error(`New teacher email invalid: ${data.newTeacher}`);
        }
        
        const finalClsEmailForCC = findClsEmailByManagerName(data.clsManager);
        
        const newTeacherResult = sendTrackedEmail({
          to: newTeacherInfo.email, 
          cc: finalClsEmailForCC, 
          subject: `Migration Notice - ${data.learner} Assigned`,
          htmlBody: getNewTeacherEmailHTML(data, data.newTeacherComments || ''), 
          jlid: data.jlid, 
          attachments: attachments
        });
        notes.push(`Teacher Email Sent (TID: ${newTeacherResult.trackingId})`);
        
        // --- B. OLD TEACHER (Moved Here to ensure it runs) ---
        if (oldTeacherInfo && isValidEmail(oldTeacherInfo.email)) {
           try {
             sendTrackedEmail({
              to: oldTeacherInfo.email, 
              cc: finalClsEmailForCC, 
              subject: `${data.learner} - Migration`,
              htmlBody: getOldTeacherEmailHTML(data, ''), 
              jlid: data.jlid
            });
            notes.push("Old Teacher Email Sent.");
           } catch(e) {
             Logger.log("Old Teacher Email Failed: " + e.message);
             notes.push("Old Teacher Email Failed.");
           }
        } else if (data.oldTeacher) {
           notes.push("Old Teacher Email Skipped (Invalid Email/Not Found).");
        }
    } else {
        notes.push("Teacher Email Skipped.");
    }

    // Safety Pause
    if (data.sendEmailToTeacher && data.sendWhatsappToParent) {
       Utilities.sleep(2000); 
    }

    // ==========================================
    // 2. WATI WHATSAPP
    // ==========================================
    if (data.sendWhatsappToParent) {
      // We wrap the ENTIRE WATI block in its own try/catch
      // This ensures that even if HubSpot/WATI crashes, the Emails (already sent above) are recorded.
      try {
        
        // SAFE HUBSPOT FETCH
        let hsData = {};
        try {
           const hubspotResult = fetchHubspotByJlid(data.jlid);
           if (hubspotResult.success) {
               hsData = hubspotResult.data;
           } else {
               Logger.log("HubSpot Warning: " + hubspotResult.message);
           }
        } catch (hsError) {
           Logger.log("HubSpot Critical Failure: " + hsError.message);
           // We continue, but parent phone might be missing
        }

        const parentPhone = hsData.parentContact;
        if (!parentPhone) throw new Error("Parent phone number missing in HubSpot.");

        let templateId = data.watiTemplateName;
        if (!templateId) {
            const templates = getTemplatesForReason(data.reasonOfMigration);
            templateId = (templates && templates.length > 0) ? templates[0].id : "migration_generic_update";
        }

        const watiParameters = getWatiParameters(templateId, data, hsData);

        // LOGGING ENHANCEMENT
        const logDetail = watiParameters.map(p => `${p.name}: ${p.value}`).join(' | ');

        // SEND
        const watiRes = sendWatiTemplate(parentPhone, templateId, watiParameters);
        
        if (watiRes.success) {
            notes.push(`WATI Sent (${templateId}) -> [DATA: ${logDetail}]`);
        } else {
            watiSuccess = false;
            notes.push(`WATI Failed: ${watiRes.message}`);
        }

      } catch(e) {
        watiSuccess = false;
        Logger.log("WATI Block Error: " + e.message);
        const safeError = e.message.substring(0, 150);
        notes.push("Error: " + safeError);
      }
    } else {
        notes.push("WhatsApp Skipped.");
    }

    // --- STATUS CHECK ---
    if (watiSuccess) {
        finalStatus = 'Success';
        return { success: true, message: 'Migration process completed successfully.' };
    } else {
        finalStatus = 'Partial Success';
        return { success: true, message: 'Email sent, but WhatsApp failed. Check Logs.' };
    }

  } catch (error) {
    Logger.log(`Error in Migration Process: ${error.message}`);
    notes.push(`Critical Error: ${error.message}`);
    return { success: false, message: `Failed: ${error.message}` };
  } finally {
    logAction('Migration Process', data.jlid, data.learner, data.oldTeacher, data.newTeacher, data.course, finalStatus, notes.join('; '), data.reasonOfMigration);
  }
}



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

