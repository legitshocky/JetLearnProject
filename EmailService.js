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
    logAction('New Learner Onboarded', formData.jlid, formData.learnerName, '', formData.teacherName, formData.course, overallStatus, logNotes, '', formData.performedBy || '');
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
    
    logAction('New Learner Onboarded', data.jlid, data.learnerName, '', data.teacherName, data.course, 'Success', `Teacher email TID: ${result.trackingId}`, '', data.performedBy || '');

    return { success: true, message: `Onboarding email sent to teacher. TID: ${result.trackingId}` };

  } catch (error) {
    Logger.log(`Error in helper sendOnboardingEmail: ${error.message}`);
    logAction('New Learner Onboarded', data.jlid, data.learnerName, '', data.teacherName, data.course, 'Failed', error.message, '', data.performedBy || '');
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

        // Enrich with HubSpot learner data + future courses from ticket
        let hsLearnerData = {};
        let futureCourseLabels = [];
        let upskillGaps = [];
        try {
          const hsResult = fetchHubspotByJlid(data.jlid);
          if (hsResult.success) hsLearnerData = hsResult.data;
        } catch(he) { Logger.log('[sendMigrationEmail] HS enrich failed: ' + he.message); }
        try {
          const tcResult = fetchLatestMigrationTicket(data.jlid);
          if (tcResult.found) {
            const rp = tcResult.rawProperties || {};
            [rp.future_course_1, rp.future_course_2, rp.future_course_3].forEach(function(raw) {
              if (!raw) return;
              try { futureCourseLabels.push(getCourseLabel(raw) || raw); } catch(e) { futureCourseLabels.push(raw); }
            });
            // Check upskill gaps for TP Manager email
            const loadResult = getTeacherSpecificLoad(data.newTeacher);
            const teacherCourses = (loadResult && loadResult.success) ? (loadResult.courses || []) : [];
            const upskilledNames = teacherCourses.map(c => (c.course || '').toLowerCase().trim());
            futureCourseLabels.forEach(function(label) {
              if (upskilledNames.indexOf(label.toLowerCase().trim()) === -1) upskillGaps.push(label);
            });
          }
        } catch(te) { Logger.log('[sendMigrationEmail] Ticket enrich failed: ' + te.message); }

        const enrichedData = Object.assign({}, data, {
          learnerHealth:           hsLearnerData.learnerHealth || '',
          learnerHealthReasonCode: hsLearnerData.learnerHealthReasonCode || '',
          currentSubscriptionTakenClasses: hsLearnerData.currentSubscriptionTakenClasses || '',
          moduleStartDate:         hsLearnerData.moduleStartDate || '',
          totalClassesJourney:     hsLearnerData.totalClassesJourney || '',
          practiceDocumentLink:    hsLearnerData.practiceDocumentLink || '',
          age:                     hsLearnerData.age || data.age || '',
          futureCourses:           futureCourseLabels
        });

        const finalClsEmailForCC = findClsEmailByManagerName(data.clsManager);

        const newTeacherResult = sendTrackedEmail({
          to: newTeacherInfo.email,
          cc: finalClsEmailForCC,
          subject: `Migration Notice - ${data.learner} Assigned`,
          htmlBody: getNewTeacherEmailHTML(enrichedData, data.newTeacherComments || ''),
          jlid: data.jlid,
          attachments: attachments
        });
        notes.push(`Teacher Email Sent (TID: ${newTeacherResult.trackingId})`);

        // --- TP MANAGER UPSKILLING GAP EMAIL + HUBSPOT TASK ---
        if (upskillGaps.length > 0) {
          const tpEmail = newTeacherInfo.tpManagerEmail || '';
          // Email
          try {
            if (tpEmail && isValidEmail(tpEmail)) {
              sendTrackedEmail({
                to: tpEmail,
                cc: finalClsEmailForCC,
                subject: `Upskilling Required — ${data.newTeacher} assigned to ${data.learner}`,
                htmlBody: getTPUpskillEmailHTML(enrichedData, upskillGaps),
                jlid: data.jlid
              });
              notes.push(`TP Upskilling Email Sent to ${tpEmail}`);
            } else {
              notes.push('TP Upskilling Email Skipped — no TP Manager email found.');
            }
          } catch(tpe) {
            Logger.log('[sendMigrationEmail] TP upskill email failed: ' + tpe.message);
            notes.push('TP Upskilling Email Failed: ' + tpe.message);
          }
          // HubSpot Task
          try {
            const dealId = hsLearnerData.dealId || '';
            // Resolve TP Manager HubSpot owner ID by the teacher's manager name
            const tpManagerName  = newTeacherInfo.manager || '';
            const tpManagerHsId  = getHubSpotOwnerIdByName(tpManagerName) || '';
            Logger.log('[sendMigrationEmail] TP Manager: ' + tpManagerName + ' → HsId: ' + tpManagerHsId);
            createUpskillTaskOnHubSpot(
              data.jlid,
              data.newTeacher,
              data.learner,
              tpManagerHsId,
              upskillGaps,
              dealId
            );
            notes.push('HubSpot Upskill Task Created for TP Manager: ' + (tpManagerName || 'unknown'));
          } catch(hte) {
            Logger.log('[sendMigrationEmail] HubSpot task failed: ' + hte.message);
            notes.push('HubSpot Upskill Task Failed: ' + hte.message);
          }
        }
        
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
    let watiErrorMessage = '';
    let watiParentEmail  = '';
    let watiSentCount    = 0;

    if (data.sendWhatsappToParent) {
      try {
        // SAFE HUBSPOT FETCH
        let hsData = {};
        try {
           const hubspotResult = fetchHubspotByJlid(data.jlid);
           if (hubspotResult.success) {
               hsData = hubspotResult.data;
               watiParentEmail = hsData.parentEmail || '';
           } else {
               Logger.log("HubSpot Warning: " + hubspotResult.message);
           }
        } catch (hsError) {
           Logger.log("HubSpot Critical Failure: " + hsError.message);
        }

        let templateId = data.watiTemplateName;
        if (!templateId) {
            const templates = getTemplatesForReason(data.reasonOfMigration);
            templateId = (templates && templates.length > 0) ? templates[0].id : "migration_generic_update";
        }

        const watiParameters = getWatiParameters(templateId, data, hsData);
        const logDetail = watiParameters.map(p => `${p.name}: ${p.value}`).join(' | ');

        // Determine target phones: use explicitly selected list or fall back to auto-fetch
        let targetPhones = [];
        if (data.watiPhoneTargets && data.watiPhoneTargets.length > 0) {
          // User explicitly selected contacts from the picker
          targetPhones = data.watiPhoneTargets.filter(Boolean);
        } else if (data.manualPhone) {
          targetPhones = [data.manualPhone];
        } else {
          // Auto-fetch best number
          const phone = hsData.parentContact;
          if (!phone) throw new Error("Parent phone number missing in HubSpot. No manual override provided.");
          targetPhones = [phone];
        }

        // Send to each selected target
        const failedTargets = [];
        targetPhones.forEach(function(phone) {
          try {
            const watiRes = sendWatiTemplate(phone, templateId, watiParameters, '');
            if (watiRes.success) {
              watiSentCount++;
              notes.push(`WATI Sent → ${phone} (${templateId})`);
            } else {
              failedTargets.push({ phone, reason: watiRes.message });
              notes.push(`WATI Failed → ${phone}: ${watiRes.message}`);
            }
          } catch(we) {
            failedTargets.push({ phone, reason: we.message });
            notes.push(`WATI Error → ${phone}: ${we.message}`);
          }
        });

        if (failedTargets.length > 0 && watiSentCount === 0) {
          watiSuccess = false;
          watiErrorMessage = failedTargets.map(f => f.phone + ': ' + f.reason).join('; ');
        } else if (failedTargets.length > 0) {
          // Partial — some sent, some failed
          notes.push(`WATI Partial: ${failedTargets.length} of ${targetPhones.length} failed`);
        } else {
          notes.push(`WATI OK → all ${watiSentCount} sent [DATA: ${logDetail}]`);
        }

      } catch(e) {
        watiSuccess = false;
        watiErrorMessage = e.message.substring(0, 200);
        Logger.log("WATI Block Error: " + e.message);
        notes.push("WATI Error: " + watiErrorMessage);
      }
    } else {
        notes.push("WhatsApp Skipped.");
    }

    // --- STATUS CHECK ---
    if (watiSuccess) {
        finalStatus = 'Success';
    } else {
        finalStatus = 'Partial Success';
    }

    // ── Write upskill note to HubSpot ticket ──────────────────────────────────
    try {
      checkAndWriteUpskillNote(data.jlid, data.newTeacher, data.confirmedFutureCourses || []);
    } catch(upErr) {
      Logger.log('[sendMigrationEmail] Upskill note failed: ' + upErr.message);
    }

    // ── Send course completion certificate to parent ───────────────────────────
    if (data.sendCertificate && data.course) {
      try {
        // Fetch parent email if not already in data
        let certParentEmail = data.parentEmail || watiParentEmail || '';
        let certParentName  = data.parentName  || '';
        if (!certParentEmail) {
          try {
            const hsRes = fetchHubspotByJlid(data.jlid);
            if (hsRes.success) {
              certParentEmail = hsRes.data.parentEmail  || '';
              certParentName  = hsRes.data.parentName   || '';
            }
          } catch(hse) {}
        }

        if (certParentEmail) {
          const certResult = sendCourseCertificateEmail(
            data.jlid, data.learner,
            data.completedCourseForCert || data.course,  // always old (completed) course, not new
            certParentEmail, certParentName,
            data.performedBy || ''
          );

          if (certResult.success) {
            notes.push('Certificate Sent to ' + certParentEmail);

            // ── HubSpot note: certificate sent ─────────────────────────────
            try {
              const tcRes = fetchLatestMigrationTicket(data.jlid);
              if (tcRes.found && tcRes.ticketId) {
                const driveUrl = certResult.driveUrl ? '\nDrive URL: ' + certResult.driveUrl : '';
                const noteText = 'Certificate Sent\n'
                  + '--------------------\n'
                  + 'Learner : ' + data.learner + '\n'
                  + 'Course  : ' + data.course + '\n'
                  + 'Emailed : ' + certParentEmail + '\n'
                  + 'Sent by : ' + (data.performedBy || 'CLS') + '\n'
                  + 'Date    : ' + new Date().toDateString()
                  + driveUrl;
                addNoteToHubSpotTicket(tcRes.ticketId, noteText);
                Logger.log('[sendMigrationEmail] HubSpot cert note written to ticket ' + tcRes.ticketId);
              } else {
                Logger.log('[sendMigrationEmail] No ticket found for JLID ' + data.jlid + ' — cert note skipped');
              }
            } catch(nte) {
              Logger.log('[sendMigrationEmail] Cert HubSpot note failed: ' + nte.message);
            }
          } else {
            notes.push('Certificate Failed: ' + certResult.message);
          }
        } else {
          notes.push('Certificate Skipped — no parent email found.');
        }
      } catch(certErr) {
        Logger.log('[sendMigrationEmail] Certificate block error: ' + certErr.message);
        notes.push('Certificate Error: ' + certErr.message);
      }
    }

    if (watiSuccess) {
        return { success: true, message: 'Migration process completed successfully.', watiError: false };
    } else {
        return {
          success: true,
          message: 'Emails sent. WhatsApp message failed — see details below.',
          watiError: true,
          watiErrorMessage: watiErrorMessage,
          parentEmail: watiParentEmail,
          learner: data.learner || '',
          newTeacher: data.newTeacher || '',
          course: data.course || '',
          jlid: data.jlid || ''
        };
    }

  } catch (error) {
    Logger.log(`Error in Migration Process: ${error.message}`);
    notes.push(`Critical Error: ${error.message}`);
    return { success: false, message: `Failed: ${error.message}` };
  } finally {
    // Include performedBy (actual username) so Impact Score is attributed correctly
    const intervenedParts = Array.isArray(data.migrationIntervenedBy)
        ? data.migrationIntervenedBy
        : (data.migrationIntervenedBy ? [data.migrationIntervenedBy] : []);
    if (data.performedBy && intervenedParts.indexOf(data.performedBy) === -1) {
      intervenedParts.push(data.performedBy);
    }
    const intervenedStr = intervenedParts.filter(Boolean).join(', ');

    logAction(
        'Migration Process', 
        data.jlid, 
        data.learner, 
        data.oldTeacher, 
        data.newTeacher, 
        data.course, 
        finalStatus, 
        notes.join('; '), 
        data.reasonOfMigration, 
        intervenedStr // <--- THIS WAS MISSING!
    );
  }
}




// ─────────────────────────────────────────────────────────────────────────────
// getWatiContactsForJlid(jlid)
// Returns all phone numbers (contacts) associated with this learner's deal.
// Used by the migration form to let CLS pick which contact(s) to WhatsApp.
// ─────────────────────────────────────────────────────────────────────────────
function getWatiContactsForJlid(jlid) {
  try {
    var hsResult = fetchHubspotByJlid(jlid);
    if (!hsResult.success) return { success: false, message: hsResult.message, contacts: [], best: '', parentEmail: '' };
    var hs = hsResult.data;
    var parentEmail = hs.parentEmail || '';
    var dealId = hs.dealId || '';

    if (!dealId) {
      // No deal ID — fall back to the single phone on the deal
      var fallback = hs.parentContact || '';
      return {
        success: true,
        contacts: fallback ? [{ number: fallback, label: 'Deal Phone', type: 'phone' }] : [],
        best: fallback,
        parentEmail: parentEmail
      };
    }

    var phoneData = getPhoneNumbersForDeal(dealId);
    var contacts = (phoneData.all || []).map(function(entry) {
      // Entry format: "+91xxxx (WhatsApp)" or "+91xxxx (Mobile)" or "+91xxxx (Phone)"
      var match = String(entry).match(/^(.+?)\s*\((.+?)\)$/);
      if (match) return { number: match[1].trim(), label: match[2].trim(), type: match[2].toLowerCase() };
      return { number: String(entry).trim(), label: 'Phone', type: 'phone' };
    });

    return {
      success: true,
      contacts: contacts,
      best: phoneData.best || '',
      parentEmail: parentEmail
    };
  } catch(e) {
    Logger.log('[getWatiContactsForJlid] Error: ' + e.message);
    return { success: false, message: e.message, contacts: [], best: '', parentEmail: '' };
  }
}


// ─────────────────────────────────────────────────────────────────────────────
// _fetchWatiTemplateBody(templateName)
// Fetches the actual message body of a WATI template via API.
// Returns the body string with {{VariableName}} placeholders, or null on failure.
// ─────────────────────────────────────────────────────────────────────────────
function _fetchWatiTemplateBody(templateName) {
  try {
    var props = PropertiesService.getScriptProperties();
    var base  = (props.getProperty('WATI_API_ENDPOINT') || '').trim();
    var token = (props.getProperty('WATI_ACCESS_TOKEN') || '').trim();
    if (!base || !token) return null;
    if (!token.startsWith('Bearer ')) token = 'Bearer ' + token;
    if (base.endsWith('/')) base = base.slice(0, -1);

    var headers = { Authorization: token };

    // Direct lookup by name — single API call (verify name matches, WATI filter is not strict)
    var url = base + '/api/v1/getMessageTemplates?pageSize=1&templateName=' + encodeURIComponent(templateName);
    var res = UrlFetchApp.fetch(url, { method: 'get', headers: headers, muteHttpExceptions: true });

    if (res.getResponseCode() === 200) {
      var data = JSON.parse(res.getContentText());
      var templates = data.messageTemplates || data.templates || data.items || data.data || [];
      if (templates.length > 0) {
        var t = templates[0];
        var returnedName = (t.elementName || t.name || '').toLowerCase();
        if (returnedName === templateName.toLowerCase()) {
          var body = t.body || t.content || null;
          if (body) return body;
        }
      }
    }

    // Fallback: paginate if direct lookup returned nothing
    var pageNumber = 1;
    while (pageNumber <= 10) {
      url = base + '/api/v1/getMessageTemplates?pageSize=200&pageNumber=' + pageNumber;
      res = UrlFetchApp.fetch(url, { method: 'get', headers: headers, muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) break;

      data = JSON.parse(res.getContentText());
      templates = data.messageTemplates || data.templates || data.items || data.data || [];
      if (!templates || templates.length === 0) break;

      for (var i = 0; i < templates.length; i++) {
        var t = templates[i];
        if ((t.name || t.elementName || '').toLowerCase() === templateName.toLowerCase()) {
          return t.body || t.content || null;
        }
      }

      if (templates.length < 200) break;
      pageNumber++;
    }

    return null;
  } catch(e) {
    Logger.log('[_fetchWatiTemplateBody] Error: ' + e.message);
    return null;
  }
}


// ─────────────────────────────────────────────────────────────────────────────
// _substituteWatiParams(body, params)
// Replaces {{VariableName}} and {{1}}/{{2}} positional placeholders with values.
// params = [{name, value}] array from getWatiParameters()
// ─────────────────────────────────────────────────────────────────────────────
function _substituteWatiParams(body, params) {
  var result = body;

  // Named substitution: {{Parent}}, {{Learner}}, {{new_teacher}} etc. (case-insensitive)
  params.forEach(function(p) {
    var regex = new RegExp('\\{\\{' + p.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\}\\}', 'gi');
    result = result.replace(regex, p.value || '');
  });

  // Positional fallback: {{1}}, {{2}} ... in order of params array
  params.forEach(function(p, idx) {
    var regex = new RegExp('\\{\\{' + (idx + 1) + '\\}\\}', 'g');
    result = result.replace(regex, p.value || '');
  });

  return result;
}


// ─────────────────────────────────────────────────────────────────────────────
// sendMigrationParentFallbackEmail(jlid, migrationContext, performedBy)
// Called when WATI fails — sends parent the exact same content as the WhatsApp.
// migrationContext = full migrationData object from the form
// ─────────────────────────────────────────────────────────────────────────────
function sendMigrationParentFallbackEmail(jlid, migrationContext, performedBy) {
  try {
    // 1. Fetch HubSpot data for parent email + enrich params
    var hsData = {};
    var parentEmail = migrationContext.parentEmail || '';
    try {
      var hsRes = fetchHubspotByJlid(jlid);
      if (hsRes.success) {
        hsData = hsRes.data;
        if (!parentEmail || !isValidEmail(parentEmail)) parentEmail = hsData.parentEmail || '';
      }
    } catch(he) { Logger.log('[FallbackEmail] HS fetch: ' + he.message); }

    if (!parentEmail || !isValidEmail(parentEmail)) {
      return { success: false, message: 'No valid parent email found for JLID: ' + jlid };
    }

    // 2. Resolve template name
    var templateId = migrationContext.watiTemplateName || '';
    if (!templateId) {
      try {
        var tpls = getTemplatesForReason(migrationContext.reasonOfMigration || '');
        templateId = (tpls && tpls.length > 0) ? tpls[0].id : 'migration_generic_update';
      } catch(e) { templateId = 'migration_generic_update'; }
    }

    // 3. Build same params WATI would have used
    var paramsList = [];
    try { paramsList = getWatiParameters(templateId, migrationContext, hsData); } catch(e) {}

    // 4. Fetch the actual template body from WATI API and substitute params
    var messageBody = null;
    try { messageBody = _fetchWatiTemplateBody(templateId); } catch(e) {}

    var learnerName = migrationContext.learner || jlid;
    var newTeacher  = migrationContext.newTeacher || '';
    var oldTeacher  = migrationContext.oldTeacher || '';
    var subject     = 'JetLearn - ' + learnerName + ' - Class Update';
    var htmlBody;

    if (messageBody) {
      // Substitute all {{params}} with real values — exact WATI content
      var filledText = _substituteWatiParams(messageBody, paramsList);
      htmlBody = getMigrationParentFallbackEmailHTML({ messageText: filledText, learnerName: learnerName });
    } else {
      // WATI API unreachable — fall back to structured summary (all key fields)
      Logger.log('[FallbackEmail] Could not fetch WATI template body, using structured fallback.');
      var paramsMap = {};
      paramsList.forEach(function(p) { paramsMap[p.name.toLowerCase()] = p.value; });
      htmlBody = getMigrationParentFallbackEmailHTML({
        messageText: null,
        learnerName:  learnerName,
        newTeacher:   paramsMap['new_teacher'] || paramsMap['teacher'] || newTeacher,
        oldTeacher:   oldTeacher,
        course:       paramsMap['course'] || paramsMap['coures_type'] || migrationContext.course || '',
        weekday:      paramsMap['weekday'] || '',
        time:         paramsMap['time'] || '',
        classLink:    paramsMap['link'] || paramsMap['meeting_link'] || migrationContext.classLink || '',
        startDate:    paramsMap['date'] || ''
      });
    }

    var sendResult = sendTrackedEmail({ to: parentEmail, subject: subject, htmlBody: htmlBody, jlid: jlid });

    logAction('Migration Email Sent', jlid, learnerName, oldTeacher, newTeacher,
      migrationContext.course || '', 'Success',
      'Parent fallback email (WhatsApp failed). TID: ' + sendResult.trackingId,
      migrationContext.reasonOfMigration || '', performedBy || '');

    return { success: true, message: 'Parent notified via email.', trackingId: sendResult.trackingId };
  } catch(e) {
    Logger.log('[sendMigrationParentFallbackEmail] Error: ' + e.message);
    return { success: false, message: e.message };
  }
}


// ─────────────────────────────────────────────────────────────────────────────
// getMigrationParentFallbackEmailHTML(ctx)
// ctx.messageText  = filled WATI template body (preferred — exact content)
// ctx.*            = structured fields used only when messageText is null
// ─────────────────────────────────────────────────────────────────────────────
function getMigrationParentFallbackEmailHTML(ctx) {
  var learner   = ctx.learnerName || 'Learner';
  var classLink = ctx.classLink || '';

  // ── PRIMARY PATH: exact WATI template body ────────────────────────────────
  if (ctx.messageText) {
    // Convert WhatsApp line breaks / *bold* to basic HTML
    var bodyHtml = ctx.messageText
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/\*(.*?)\*/g, '<strong>$1</strong>')   // *bold*
      .replace(/_(.*?)_/g, '<em>$1</em>')             // _italic_
      .replace(/\n/g, '<br/>');

    return `<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>JetLearn - ${learner} - Class Update</title>
</head>
<body style="margin:0;padding:0;background:#f0ede8;font-family:Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f0ede8;padding:32px 16px;">
  <tr><td align="center">
    <table width="100%" cellpadding="0" cellspacing="0"
      style="max-width:520px;background:#ffffff;border-radius:14px;border:1px solid #e4e1dc;overflow:hidden;">

      <!-- Minimal header -->
      <tr>
        <td style="background:linear-gradient(135deg,#2d2a6e,#4c3a9e);padding:18px 28px;">
          <span style="color:#fff;font-size:1.1rem;font-weight:700;">✈ JetLearn</span>
        </td>
      </tr>

      <!-- Exact template body -->
      <tr>
        <td style="padding:28px 28px 24px;font-size:0.93rem;color:#222;line-height:1.75;">
          ${bodyHtml}
        </td>
      </tr>

      ${classLink ? `<!-- CTA -->
      <tr>
        <td style="padding:0 28px 28px;text-align:center;">
          <a href="${classLink}"
            style="display:inline-block;background:linear-gradient(135deg,#2d2a6e,#4c3a9e);
              color:#ffffff;text-decoration:none;padding:11px 28px;border-radius:9px;
              font-weight:600;font-size:0.88rem;">
            Join Class →
          </a>
        </td>
      </tr>` : ''}

    </table>
  </td></tr>
</table>
</body></html>`;
  }

  // ── FALLBACK PATH: WATI API unreachable, render key fields ────────────────
  var newTeacher = ctx.newTeacher  || '';
  var oldTeacher = ctx.oldTeacher  || '';
  var course     = ctx.course      || '';
  var weekday    = ctx.weekday     || '';
  var time       = ctx.time        || '';
  var startDate  = ctx.startDate   || '';

  var rows = '';
  if (newTeacher) rows += `<tr><td style="padding:8px 0;border-bottom:1px solid #f0ede8;color:#888;font-size:0.83rem;width:120px;">New Teacher</td><td style="padding:8px 0;border-bottom:1px solid #f0ede8;color:#111;font-weight:600;">${newTeacher}</td></tr>`;
  if (oldTeacher) rows += `<tr><td style="padding:8px 0;border-bottom:1px solid #f0ede8;color:#888;font-size:0.83rem;">Previous Teacher</td><td style="padding:8px 0;border-bottom:1px solid #f0ede8;color:#555;">${oldTeacher}</td></tr>`;
  if (course)     rows += `<tr><td style="padding:8px 0;border-bottom:1px solid #f0ede8;color:#888;font-size:0.83rem;">Course</td><td style="padding:8px 0;border-bottom:1px solid #f0ede8;color:#111;">${course}</td></tr>`;
  if (weekday||time) rows += `<tr><td style="padding:8px 0;border-bottom:1px solid #f0ede8;color:#888;font-size:0.83rem;">Schedule</td><td style="padding:8px 0;border-bottom:1px solid #f0ede8;color:#111;font-weight:600;">${weekday}${weekday&&time?' · ':''}${time}</td></tr>`;
  if (startDate)  rows += `<tr><td style="padding:8px 0;border-bottom:1px solid #f0ede8;color:#888;font-size:0.83rem;">Effective From</td><td style="padding:8px 0;border-bottom:1px solid #f0ede8;color:#111;">${startDate}</td></tr>`;
  if (classLink)  rows += `<tr><td style="padding:8px 0;color:#888;font-size:0.83rem;">Class Link</td><td style="padding:8px 0;"><a href="${classLink}" style="color:#5546d4;">${classLink}</a></td></tr>`;

  return `<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"/>
<title>JetLearn - ${learner} - Class Update</title>
</head>
<body style="margin:0;padding:0;background:#f0ede8;font-family:Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f0ede8;padding:32px 16px;">
  <tr><td align="center">
    <table width="100%" cellpadding="0" cellspacing="0"
      style="max-width:520px;background:#fff;border-radius:14px;border:1px solid #e4e1dc;overflow:hidden;">
      <tr><td style="background:linear-gradient(135deg,#2d2a6e,#4c3a9e);padding:18px 28px;">
        <span style="color:#fff;font-size:1.1rem;font-weight:700;">✈ JetLearn</span>
      </td></tr>
      <tr><td style="padding:24px 28px 8px;font-size:0.93rem;color:#333;">
        Here is an update regarding <strong>${learner}</strong>'s upcoming classes.
      </td></tr>
      <tr><td style="padding:12px 28px 24px;">
        <table width="100%" cellpadding="0" cellspacing="0"
          style="background:#f8f7ff;border:1px solid #ddd8fa;border-radius:8px;padding:4px 14px;">
          ${rows}
        </table>
      </td></tr>
    </table>
  </td></tr>
</table>
</body></html>`;
}


function handleTrackingPixel(trackingId) {
  const transparentPngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";
  try {
    if (trackingId) {
      const sheet = getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.EMAIL_LOGS);
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
    const sheet = getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.EMAIL_LOGS);
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
  Logger.log('Starting reply check job...');
  try {
    const sheet = getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.EMAIL_LOGS);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    const headers = data[0];
    const statusCol      = headers.indexOf('Status');
    const recipientCol   = headers.indexOf('Recipient');
    const subjectCol     = headers.indexOf('Subject');
    const sentAtCol      = headers.indexOf('Sent At');
    const timestampCol   = headers.indexOf('Timestamp');   // fallback
    const repliedAtCol   = headers.indexOf('Replied At');
    const replyContentCol = headers.indexOf('Reply Content');

    // Only check last 100 rows to avoid timeout
    const startRow = Math.max(1, data.length - 100);

    for (let i = data.length - 1; i >= startRow; i--) {
      const row = data[i];
      const currentStatus = String(row[statusCol] || '');

      if (currentStatus !== 'Sent' && currentStatus !== 'Opened') continue;

      const recipient = String(row[recipientCol] || '').trim();
      const subject   = String(row[subjectCol]   || '').trim();

      // Try 'Sent At' first, fall back to 'Timestamp'
      let sentAt = new Date(row[sentAtCol]);
      if (isNaN(sentAt.getTime())) sentAt = new Date(row[timestampCol]);
      if (isNaN(sentAt.getTime())) continue;

      if (!recipient || !subject) continue;

      // Only check emails sent in the last 30 days
      const thirtyDaysAgo = new Date();
      thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
      if (sentAt < thirtyDaysAgo) continue;

      try {
        const dateStr = Utilities.formatDate(sentAt, Session.getScriptTimeZone(), 'yyyy/MM/dd');
        const searchQuery = `subject:("${subject}") from:(${recipient}) after:${dateStr}`;
        const threads = GmailApp.search(searchQuery, 0, 1);

        if (threads.length > 0) {
          const messages = threads[0].getMessages();
          for (const message of messages) {
            const messageDate = message.getDate();
            const fromEmail   = message.getFrom();

            if (fromEmail.includes(recipient) && messageDate > sentAt) {
              sheet.getRange(i + 1, statusCol + 1).setValue('Replied');
              sheet.getRange(i + 1, repliedAtCol + 1).setValue(messageDate);
              sheet.getRange(i + 1, replyContentCol + 1).setValue(
                message.getPlainBody().substring(0, 500)
              );
              Logger.log(`Reply detected from ${recipient} for subject: ${subject}`);
              break;
            }
          }
        }
      } catch (rowError) {
        Logger.log(`Skipping row ${i} due to error: ${rowError.message}`);
        continue;
      }
    }

    Logger.log('Reply check job finished.');
  } catch (error) {
    Logger.log(`Error during reply check job: ${error.message}`);
  }
}


function getEmailActivities(params = {}) {
  Logger.log('getEmailActivities called with params: ' + JSON.stringify(params));

  try {
    const sheet = getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.EMAIL_LOGS);
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
    const sheet = getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.EMAIL_LOGS);
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

function getTPUpskillEmailHTML(data, gapCourses) {
  const template = HtmlService.createTemplateFromFile('TPUpskillTemplate');
  template.data = data;
  template.gapCourses = gapCourses;
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

function _getBlobFromDriveFolder(folderId, fallbackName) {
  // Grabs the first file found in a Drive folder and returns its blob
  try {
    var folder = DriveApp.getFolderById(folderId);
    var files   = folder.getFiles();
    if (!files.hasNext()) {
      Logger.log('_getBlobFromDriveFolder: folder ' + folderId + ' is empty.');
      return null;
    }
    var file = files.next();
    var blob = file.getBlob();
    // Preserve original filename; fallbackName used only if Drive returns blank
    var name = file.getName() || fallbackName;
    blob.setName(name);
    Logger.log('_getBlobFromDriveFolder: attached "' + name + '" from folder ' + folderId);
    return blob;
  } catch(e) {
    Logger.log('_getBlobFromDriveFolder error (' + folderId + '): ' + e.message);
    return null;
  }
}

function _getMinecraftDefaultAttachments() {
  var blobs = [];
  var sources = [
    { folder: CONFIG.MINECRAFT_VIDEO_MAC_FOLDER, name: 'Minecraft Steps - Mac.mp4'     },
    { folder: CONFIG.MINECRAFT_VIDEO_WIN_FOLDER, name: 'Minecraft Steps - Windows.mp4' }
  ];
  sources.forEach(function(src) {
    if (!src.folder) return;
    var blob = _getBlobFromDriveFolder(src.folder, src.name);
    if (blob) blobs.push(blob);
  });
  return blobs;
}

function _getRobloxDefaultAttachments() {
  var blobs = [];
  if (!CONFIG.ROBLOX_SETUP_PDF_FOLDER) return blobs;
  var blob = _getBlobFromDriveFolder(CONFIG.ROBLOX_SETUP_PDF_FOLDER, 'Roblox Studio Setup Instructions.pdf');
  if (blob) blobs.push(blob);
  return blobs;
}

function sendGenericEmail(emailType, formData, recipientEmail, attachmentsBase64, comments) {
  try {
    if (!isValidEmail(recipientEmail)) throw new Error('Valid Recipient Email is required.');

    let subject = '';
    let htmlBody = '';

    switch (emailType) {
      case 'Minecraft Install':
        subject = 'Minecraft Course Download Link';
        htmlBody = getMinecraftInstallEmailHTML(formData, comments);
        break;
      case 'Roblox Install':
        subject = 'Download & Set-up for Roblox Studio';
        htmlBody = getRobloxInstallEmailHTML(formData, comments);
        break;
      default: throw new Error(`Unknown email type: ${emailType}.`);
    }

    const attachments = (attachmentsBase64 && attachmentsBase64.length > 0) ? uploadAttachments(attachmentsBase64) : [];

    // Auto-append default Drive attachments for course emails
    if (emailType === 'Minecraft Install') {
      const mcAuto = _getMinecraftDefaultAttachments();
      attachments.push(...mcAuto);
    }
    if (emailType === 'Roblox Install') {
      const rbAuto = _getRobloxDefaultAttachments();
      attachments.push(...rbAuto);
    }

    const result = sendTrackedEmail({ to: recipientEmail, subject, htmlBody, jlid: formData.jlid, attachments });
    
    logAction(`Email Sent (${emailType})`, formData.jlid, formData.learnerName, '', '', formData.course, 'Success', `TID: ${result.trackingId}`, '', formData.performedBy || '');
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

function _parseCsvRow(line) {
  const fields = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQuotes) {
      if (ch === '"' && i + 1 < line.length && line[i + 1] === '"') {
        current += '"';
        i++;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        current += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === ',') {
        fields.push(current);
        current = '';
      } else {
        current += ch;
      }
    }
  }
  fields.push(current);
  return fields;
}

function processBatchUpload(csvContent) {
  Logger.log('processBatchUpload called');

  const lines = csvContent.split('\n');
  if (lines.length < 2) {
    return { total: 0, success: 0, failed: 0, results: [], message: 'CSV is empty or has only headers.' };
  }

  const headers = _parseCsvRow(lines[0]).map(h => h.trim());
  const requiredHeaders = ['jlid', 'learnerName', 'oldTeacher', 'newTeacher', 'course', 'reason', 'clsManager', 'jetGuide', 'startDate', 'day', 'time'];

  const missingHeaders = requiredHeaders.filter(rh => !headers.includes(rh));
  if (missingHeaders.length > 0) {
    return { success: false, message: `Missing required CSV headers: ${missingHeaders.join(', ')}.`, total: 0, results: [] };
  }

  const results = [];
  const migrationsToProcess = [];

  for (let i = 1; i < lines.length; i++) {
    if (lines[i].trim() === '') continue;

    const row = _parseCsvRow(lines[i]);
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

