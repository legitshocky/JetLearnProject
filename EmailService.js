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
  let finalStatus = 'Partial Success';
  let notes = [];

  try {
    // ==========================================
    // 1. TEACHER EMAIL (Always CET)
    // ==========================================
    if (data.sendEmailToTeacher) {
        if (!data.jlid || !data.newTeacher || !data.course) throw new Error('Missing fields for Email.');
        
        const teacherData = getTeacherData();
        // Case-insensitive lookup for new teacher
        const newTeacherInfo = teacherData.find(t => String(t.name).trim().toLowerCase() === String(data.newTeacher).trim().toLowerCase());
        
        // Lookup old teacher if provided
        let oldTeacherInfo = null;
        if (data.oldTeacher) {
            oldTeacherInfo = teacherData.find(t => String(t.name).trim().toLowerCase() === String(data.oldTeacher).trim().toLowerCase());
        }

        if (!newTeacherInfo || !isValidEmail(newTeacherInfo.email)) {
            throw new Error(`New teacher email invalid or not found: ${data.newTeacher}`);
        }
        
        const finalClsEmailForCC = findClsEmailByManagerName(data.clsManager);
        
        // Send to New Teacher
        const newTeacherResult = sendTrackedEmail({
          to: newTeacherInfo.email, 
          cc: finalClsEmailForCC, 
          subject: `Migration Notice - ${data.learner} Assigned`,
          htmlBody: getNewTeacherEmailHTML(data, data.newTeacherComments || ''), 
          jlid: data.jlid, 
          attachments: attachments
        });
        notes.push(`Teacher Email Sent (TID: ${newTeacherResult.trackingId})`);
        
        // Optional: Send to Old Teacher
        if (oldTeacherInfo && isValidEmail(oldTeacherInfo.email)) {
           sendTrackedEmail({
            to: oldTeacherInfo.email, 
            cc: finalClsEmailForCC, 
            subject: `${data.learner} - Migration`,
            htmlBody: getOldTeacherEmailHTML(data, ''), 
            jlid: data.jlid
          });
          notes.push("Old Teacher Email Sent.");
        }
    } else {
        notes.push("Teacher Email Skipped.");
    }

    // ==========================================
    // 2. WATI WHATSAPP (Parent's Timezone)
    // ==========================================
    if (data.sendWhatsappToParent) {
        
        // A. Fetch Hubspot Data to get Parent Phone & Name
        const hubspotResult = fetchHubspotByJlid(data.jlid);
        if (!hubspotResult.success) {
             throw new Error(`WATI Failed: Could not fetch data for JLID ${data.jlid}.`);
        }
        
        const hsData = hubspotResult.data;
        const parentPhone = hsData.parentContact;
        
        if (!parentPhone) throw new Error("WATI Failed: Parent phone number missing in HubSpot.");

        // B. Calculate Local Time for the message
        const firstSession = data.classSessions && data.classSessions.length > 0 ? data.classSessions[0] : { day: "TBD", time: "TBD" };
        const cetTime = firstSession.time; 
        const parentTimezone = data.manualTimezone || hsData.timezone;
        // Inject into data object for helper to find
        data.calculatedLocalTime = convertCetToLocal(cetTime, parentTimezone); 

        // C. DETERMINE TEMPLATE ID
        // Priority 1: Use the specific template ID sent from the dropdown (data.watiTemplateName)
        // Priority 2: Look up default for the reason
        // Priority 3: Generic fallback
        let templateId = data.watiTemplateName;
        
        if (!templateId || templateId === "") {
            const mapping = WATI_REASON_MAPPING[data.reasonOfMigration];
            if (mapping && mapping.length > 0) {
                templateId = mapping[0].id; // Use the first mapped template
            } else {
                templateId = "migration_generic_update"; // Final fallback
            }
        }

        // D. Build Parameters
        const watiParameters = getWatiParameters(templateId, data, hsData);

        // Debug Log
        Logger.log(`Sending WATI Template: ${templateId}`);
        Logger.log(`Params: ${JSON.stringify(watiParameters)}`);

        // E. Send Message
        const watiResult = sendWatiMessage(parentPhone, templateId, watiParameters);
        
        if(watiResult.success) {
            notes.push(`WATI Sent (${templateId}) to ${parentPhone}`);
        } else {
            notes.push(`WATI Failed: ${JSON.stringify(watiResult)}`);
        }
    } else {
        notes.push("WhatsApp Skipped.");
    }

    finalStatus = 'Success';
    return { success: true, message: 'Migration process completed successfully.' };

  } catch (error) {
    Logger.log(`Error in Migration Process: ${error.message}`);
    notes.push(`Error: ${error.message}`);
    return { success: false, message: `Failed: ${error.message}` };
  } finally {
    // Log the outcome
    logAction('Migration Process', data.jlid, data.learner, data.oldTeacher, data.newTeacher, data.course, finalStatus, notes.join('; '), data.reasonOfMigration);
  }
}
