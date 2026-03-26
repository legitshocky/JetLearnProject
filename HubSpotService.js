function safeParseHubspotNumber(value, defaultValue = 0) {
  if (value === null || value === undefined) return defaultValue;
  if (typeof value === 'number') return value;
  // Remove everything that isn't a digit, a decimal point, or a minus sign
  const cleaned = String(value).replace(/[^0-9.-]/g, '');
  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? defaultValue : parsed;
}



function fetchHubspotByJlid(jlid) {
  Logger.log('fetchHubspotByJlid called for JLID: ' + jlid);
  
  if (!jlid) return { success: false, message: 'JLID is required.' };

  const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
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

  const options = {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(hubspotApiUrl, options);
    const responseBody = response.getContentText();

    if (responseBody.trim().startsWith("<")) return { success: false, message: `HubSpot API Error: Connection Failed` };

    const jsonResponse = JSON.parse(responseBody);

    if (jsonResponse.results && jsonResponse.results.length > 0) {
      const contactProperties = jsonResponse.results[0].properties; 
      
      const smartPhone = getBestPhoneNumberForDeal(contactProperties.hs_object_id);
      const finalParentPhone = smartPhone || contactProperties.phone_number_deal_ || contactProperties.phone || '';

      const tenure = safeParseHubspotNumber(contactProperties.subscription_tenure);
      const dealAmount = safeParseHubspotNumber(contactProperties.amount);
      const currencyCode = contactProperties.deal_currency_code || 'EUR'; 
      let calculatedDiscount = 0;

      if (tenure > 0 && dealAmount > 0) {
          const standardPriceEur = tenure * 149; 
          const conversionRate = getConversionRate(currencyCode); 
          const standardPriceLocal = standardPriceEur * conversionRate;
          if (standardPriceLocal > dealAmount) {
              const rawDiscount = standardPriceLocal - dealAmount;
              calculatedDiscount = parseFloat(rawDiscount.toFixed(2));
          }
      }

      // --- NEW: CHURN RISK CALCULATION ---
      let churnAlert = null;
      try {
          const ticketStats = getMigrationHistoryStats(jlid); 
          const today = new Date();
          const threeMonthsAgo = new Date();
          threeMonthsAgo.setMonth(today.getMonth() - 3);

          // Filter tickets: Only count migrations in the last 90 days
          const recentMigrations = ticketStats.events.filter(t => {
              const tDate = new Date(t.date);
              return tDate >= threeMonthsAgo;
          });

          if (recentMigrations.length >= 2) {
              churnAlert = {
                  level: 'Critical',
                  count: recentMigrations.length,
                  message: `⚠️ HIGH RISK: This learner has moved ${recentMigrations.length} times in the last 3 months!`
              };
          } else if (recentMigrations.length === 1) {
              churnAlert = {
                  level: 'Warning',
                  count: 1,
                  message: `Note: Learner moved 1 time recently.`
              };
          }
      } catch (statsErr) {
          Logger.log("Error calculating churn risk: " + statsErr.message);
      }
      // -----------------------------------

      // Helper: Parse Class Timings
      const parseClassTimings = (t) => { 
          if(!t) return []; 
          if(t.includes(' at ')) return t.split(';').map(s=>{const p=s.trim().match(/(\w+)\s+at\s+(\d{1,2}:\d{2}\s(?:AM|PM))/i);return p?{day:p[1],time:p[2]}:null}).filter(Boolean); 
          return t.split(/[,;]/).map(d=>d.trim()?{day:d.trim(),time:''}:null).filter(Boolean); 
      };

      // Helper: Parse Payment Plan
      const parsePaymentPlan = (h) => { 
          if(!h) return {paymentPlanType:'Upfront',installmentFrequency:'',customPlanDetails:''}; 
          h=h.toLowerCase(); 
          if(h.includes('installment')) return {paymentPlanType:'Installment',installmentFrequency:h.includes('quarterly')?'Quarterly':'Monthly',customPlanDetails:''}; 
          return {paymentPlanType:'Upfront',installmentFrequency:'',customPlanDetails:''}; 
      };

      const paymentPlanParsed = parsePaymentPlan(contactProperties.payment_type);

      const data = {
        dealId: contactProperties.hs_object_id || null,
        jlid: contactProperties.jetlearner_id || jlid,
        learnerName: `${contactProperties.dealname || ''}`.trim(),
        parentName: contactProperties.parent_name || '',
        parentEmail: contactProperties.parent_email || '',
        parentContact: finalParentPhone, 
        course: getCourseLabel(contactProperties.current_course) || '',
        subscriptionTenureMonths: tenure,
        dealAmount: dealAmount,
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
        sessionsPerWeek: contactProperties.frequency_of_classes || '',
        timezone: contactProperties.time_zone || '',
        paymentReceivedDate: contactProperties.stage____payment_trigger_date || null,
        installmentTerms: contactProperties.installment_terms_final || '',
        discount: calculatedDiscount,
        
        // Pass the alert object
        churnAlert: churnAlert 
      };
      
      return { success: true, data: data };
    } else {
      return { success: false, message: 'No learner found with this JLID.' };
    }
  } catch (error) {
    return { success: false, message: 'HubSpot Connection Error: ' + error.message };
  }
}



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

function getBestPhoneNumberForDeal(dealId) {
  const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  if (!dealId || !token) return null;

  try {
    // 1. Get IDs of Contacts associated with this Deal
    const assocUrl = `https://api.hubapi.com/crm/v4/objects/deals/${dealId}/associations/contacts`;
    const assocOptions = {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    };
    
    const assocRes = UrlFetchApp.fetch(assocUrl, assocOptions);
    const assocData = JSON.parse(assocRes.getContentText());

    if (!assocData.results || assocData.results.length === 0) {
      // No contacts found, fallback to deal property
      return null;
    }

    // Get all Contact IDs
    const contactIds = assocData.results.map(r => ({ id: r.toObjectId }));

    // 2. Fetch Phone Properties for these Contacts
    const contactsUrl = `https://api.hubapi.com/crm/v3/objects/contacts/batch/read`;
    const contactsPayload = {
      properties: ["mobilephone", "phone", "hs_whatsapp_phone_number"],
      inputs: contactIds
    };
    
    const contactsRes = UrlFetchApp.fetch(contactsUrl, {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(contactsPayload),
      muteHttpExceptions: true
    });
    
    const contactsData = JSON.parse(contactsRes.getContentText());

    // 3. Logic to find the BEST number
    let bestNumber = null;

    if (contactsData.results && contactsData.results.length > 0) {
        // Loop through all parents/contacts attached to this deal
        for (const contact of contactsData.results) {
            const props = contact.properties;
            
            // Priority 1: Specific WhatsApp Field
            if (props.hs_whatsapp_phone_number) {
                return props.hs_whatsapp_phone_number;
            }
            
            // Priority 2: Mobile Phone
            if (props.mobilephone && !bestNumber) {
                bestNumber = props.mobilephone;
            }
            
            // Priority 3: Standard Phone (Fallback)
            if (props.phone && !bestNumber) {
                bestNumber = props.phone;
            }
        }
    }
    
    return bestNumber;

  } catch (e) {
    Logger.log(`[HubSpot Phone Fetch Error]: ${e.message}`);
    return null; // Fallback to the original deal property if this fails
  }
}

function getPhoneNumbersForDeal(dealId) {
  const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  if (!dealId || !token) return { best: null, all: [] };

  try {
    // 1. Get IDs of Contacts
    const assocUrl = `https://api.hubapi.com/crm/v4/objects/deals/${dealId}/associations/contacts`;
    const assocOptions = { method: 'get', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true };
    const assocRes = UrlFetchApp.fetch(assocUrl, assocOptions);
    const assocData = JSON.parse(assocRes.getContentText());

    if (!assocData.results || assocData.results.length === 0) return { best: null, all: [] };

    const contactIds = assocData.results.map(r => ({ id: r.toObjectId }));

    // 2. Fetch Phone Properties
    const contactsUrl = `https://api.hubapi.com/crm/v3/objects/contacts/batch/read`;
    const contactsPayload = {
      properties: ["mobilephone", "phone", "hs_whatsapp_phone_number"],
      inputs: contactIds
    };
    
    const contactsRes = UrlFetchApp.fetch(contactsUrl, {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(contactsPayload),
      muteHttpExceptions: true
    });
    
    const contactsData = JSON.parse(contactsRes.getContentText());
    
    // 3. Collect ALL numbers and pick the BEST one
    let bestNumber = null;
    let allNumbers = new Set();

    if (contactsData.results) {
        for (const contact of contactsData.results) {
            const p = contact.properties;
            
            if (p.hs_whatsapp_phone_number) {
                bestNumber = p.hs_whatsapp_phone_number;
                allNumbers.add(p.hs_whatsapp_phone_number + " (WhatsApp)");
            }
            if (p.mobilephone) {
                if (!bestNumber) bestNumber = p.mobilephone;
                allNumbers.add(p.mobilephone + " (Mobile)");
            }
            if (p.phone) {
                if (!bestNumber) bestNumber = p.phone;
                allNumbers.add(p.phone + " (Phone)");
            }
        }
    }
    
    return { best: bestNumber, all: Array.from(allNumbers) };

  } catch (e) {
    Logger.log(`[HubSpot Phone Fetch Error]: ${e.message}`);
    return { best: null, all: [] }; 
  }
}

function fetchHubspotHistory(dealId) {
  if (!dealId) return [];

  const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  const headers = { 'Authorization': 'Bearer ' + token };
  const historyEvents = [];

  try {
    // 1. Fetch Associated TICKETS
    const ticketUrl = `https://api.hubapi.com/crm/v4/objects/deals/${dealId}/associations/tickets`;
    const ticketRes = UrlFetchApp.fetch(ticketUrl, { headers: headers, muteHttpExceptions: true });
    const ticketAssoc = JSON.parse(ticketRes.getContentText()).results || [];

    if (ticketAssoc.length > 0) {
      const batchBody = {
        properties: ["subject", "content", "createdate", "hs_pipeline_stage"],
        inputs: ticketAssoc.map(t => ({ id: t.toObjectId }))
      };
      const detailsRes = UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/objects/tickets/batch/read', {
        method: 'post',
        headers: { ...headers, 'Content-Type': 'application/json' },
        payload: JSON.stringify(batchBody),
        muteHttpExceptions: true
      });
      
      const tickets = JSON.parse(detailsRes.getContentText()).results || [];
      
      // --- FILTERING LOGIC ---
      // We only want to show tickets that impact the Teacher/Course
      const relevantKeywords = ["Migration", "Pause", "Escalation", "Teacher Change", "Slot Change", "Onboarding"];
      const ignoreKeywords = ["PRM", "Renewal", "Kit", "Device", "Laptop", "Feedback"];

      tickets.forEach(t => {
        const subject = t.properties.subject || "";
        
        // A. Exclude Noise (PRMs, Renewals)
        if (ignoreKeywords.some(kw => subject.includes(kw))) return;

        // B. (Optional) Only Include Specific Topics
        // If you want to be very strict, uncomment the next line:
        // if (!relevantKeywords.some(kw => subject.includes(kw))) return;

        historyEvents.push({
          timestamp: t.properties.createdate,
          type: 'hubspot-ticket',
          description: `[HubSpot Ticket] ${subject}`,
          source: 'HubSpot'
        });
      });
    }

    // 2. Fetch Associated NOTES
    // --- REMOVED ENTIRELY TO REDUCE NOISE ---
    // Notes are usually sales calls ("Called mom, no answer") which we don't need here.

    return historyEvents;

  } catch (e) {
    Logger.log("Error fetching HubSpot History: " + e.message);
    return []; 
  }
}


/**
 * getComprehensiveLearnerHistory(jlid)
 * Powers the Learner Migration Timeline page.
 * Returns: learnerProfile, migrationTimeline, auditLogTimeline, journeyAnalysis, aiSummary.
 */
function getComprehensiveLearnerHistory(jlid) {
  try {
    if (!jlid) return { success: false, message: 'JLID is required.' };
    Logger.log('[getLearnerTimeline] JLID: ' + jlid);

    var PORTAL_ID          = '7729491';
    var MIGRATION_PIPELINE = '66161281';
    var token              = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    var headers            = { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' };

    var STAGE_LABELS = {
      '128913747': 'Migration Triggered',
      '128913748': 'WIP',
      '128913750': 'WIP - TP Approval Pending',
      '128913752': 'WIP - CLS Approval Pending',
      '1030980247': 'WIP - Rejected by CLS',
      '133755411': 'WIP - Approved by CLS',
      '1065336836': 'Execution Pending',
      '128913749': 'WIP - PR Approval Pending',
      '128913753': 'Migration Completed'
    };
    var EXCLUDED_STAGES  = ['133821818', '153457301'];
    var INBOUND_REASONS  = ['Slot change - Learner request', 'Slot change -Learner request', 'Teacher Affinity', 'Pause Request', 'Special Learning Needs', 'Course Change'];
    var IGNORE_KEYWORDS  = ['PRM', 'Renewal', 'Feedback', 'Review', 'Kit', 'Laptop', 'Device', 'Tab'];

    // ── 1. Learner profile from HubSpot deal ─────────────────────────────
    var learnerProfile = { learnerName: jlid, jlid: jlid };
    var dealId = null;
    try {
      var dealResult = fetchHubspotByJlid(jlid);
      if (dealResult && dealResult.success && dealResult.data) {
        var d = dealResult.data;
        dealId = d.dealId || null;
        learnerProfile = {
          learnerName:             d.learnerName           || jlid,
          jlid:                    d.jlid                  || jlid,
          age:                     d.age                   || 'N/A',
          course:                  d.course                || 'N/A',
          currentTeacher:          d.currentTeacher        || 'N/A',
          currentSubscriptionType: d.paymentType           || 'N/A',
          startingDate:            d.startingDate          || null,
          subscriptionStartDate:   d.subscriptionStartDate || null,
          dealAmount:              d.dealAmount            || 0,
          currency:                d.currency              || 'EUR',
          tenure:                  d.subscriptionTenureMonths || 0,
          jetGuide:                d.jetGuideName          || 'N/A',
          hubspotLink:             dealId ? 'https://app.hubspot.com/contacts/' + PORTAL_ID + '/deal/' + dealId : null
        };
      }
    } catch(de) { Logger.log('[getLearnerTimeline] Deal fetch error: ' + de.message); }

    // ── 2. Migration tickets for this JLID ───────────────────────────────
    var migrationTimeline = [];
    var inboundCount = 0, outboundCount = 0;
    try {
      var ticketBody = {
        filterGroups: [{ filters: [
          { propertyName: 'learner_uid',  operator: 'EQ', value: jlid },
          { propertyName: 'hs_pipeline', operator: 'EQ', value: MIGRATION_PIPELINE }
        ]}],
        properties: [
          'subject', 'createdate', 'reason_of_migration__t_',
          'new_teacher', 'current_teacher__t_', 'hs_pipeline_stage',
          'migration_completed_date', 'hs_ticket_id', 'migration_intervened_by'
        ],
        sorts: [{ propertyName: 'createdate', direction: 'ASCENDING' }],
        limit: 100
      };
      var ticketRes  = UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/objects/tickets/search', {
        method: 'post', headers: headers, payload: JSON.stringify(ticketBody), muteHttpExceptions: true
      });
      var ticketData = JSON.parse(ticketRes.getContentText());
      var rawTickets = (ticketData && ticketData.results) ? ticketData.results : [];

      rawTickets.forEach(function(t) {
        var props   = t.properties || {};
        var subject = String(props.subject                    || '').trim();
        var reason  = String(props.reason_of_migration__t_   || '').trim();
        var stage   = String(props.hs_pipeline_stage         || '').trim();

        if (EXCLUDED_STAGES.indexOf(stage) !== -1) return;
        if (IGNORE_KEYWORDS.some(function(kw){ return subject.indexOf(kw) !== -1; })) return;
        if (!reason && subject.indexOf('Migration') === -1) return;

        var isInbound     = INBOUND_REASONS.indexOf(reason) !== -1;
        var stageLabel    = STAGE_LABELS[stage] || (stage ? 'Stage ' + stage : 'Unknown');
        var isCompleted   = stageLabel === 'Migration Completed';
        var triggeredDate = t.createdAt ? new Date(t.createdAt) : (props.createdate ? new Date(props.createdate) : null);
        var completedDate = props.migration_completed_date ? new Date(props.migration_completed_date) : null;
        var daysToResolve = (triggeredDate && completedDate && isCompleted)
          ? Math.round((completedDate - triggeredDate) / 86400000) : null;

        if (isInbound) inboundCount++; else outboundCount++;

        migrationTimeline.push({
          id:            t.id,
          ticketId:      props.hs_ticket_id || t.id,
          hubspotLink:   'https://app.hubspot.com/contacts/' + PORTAL_ID + '/ticket/' + t.id,
          date:          triggeredDate ? triggeredDate.toISOString() : null,
          completedDate: completedDate ? completedDate.toISOString() : null,
          daysToResolve: daysToResolve,
          subject:       subject || 'Migration',
          reason:        reason  || 'Unspecified',
          fromTeacher:   getTeacherLabel(props.current_teacher__t_) || 'Unknown',
          toTeacher:     getTeacherLabel(props.new_teacher)         || 'Not assigned',
          stage:         stageLabel,
          isCompleted:   isCompleted,
          type:          isInbound ? 'inbound' : 'outbound',
          intervenedBy:  props.migration_intervened_by || ''
        });
      });
    } catch(te) { Logger.log('[getLearnerTimeline] Ticket fetch error: ' + te.message); }

    // ── 3. Audit log rows for this JLID ──────────────────────────────────
    var auditLogTimeline = [];
    try {
      var auditResult = getAuditLog({ limit: 500 });
      var auditRows   = (auditResult && auditResult.data) ? auditResult.data : [];
      auditRows.forEach(function(row) {
        if (String(row[2] || '').trim() !== jlid) return;
        var action = String(row[1] || '').trim();
        if (!action) return;
        auditLogTimeline.push({
          timestamp:   row[0]  || null,
          action:      action,
          fromTeacher: String(row[4] || '').trim(),
          toTeacher:   String(row[5] || '').trim(),
          course:      String(row[6] || '').trim(),
          status:      String(row[7] || '').trim(),
          notes:       String(row[8] || '').trim()
        });
      });
      auditLogTimeline.sort(function(a, b){ return new Date(a.timestamp||0) - new Date(b.timestamp||0); });
    } catch(ae) { Logger.log('[getLearnerTimeline] Audit log error: ' + ae.message); }

    // ── 4. Journey analysis ───────────────────────────────────────────────
    var total = migrationTimeline.length;
    var riskLevel, riskMessage;
    if      (outboundCount >= 3) { riskLevel = 'Critical'; riskMessage = 'Learner has been moved 3+ times by JetLearn. High churn risk — CLS review recommended immediately.'; }
    else if (outboundCount === 2) { riskLevel = 'High';     riskMessage = 'Two JetLearn-initiated moves on record. Monitor closely and ensure the current teacher is a strong fit.'; }
    else if (outboundCount === 1) { riskLevel = 'Medium';   riskMessage = 'One JetLearn-initiated move on record. Journey is mostly stable.'; }
    else if (inboundCount >= 2)   { riskLevel = 'Watch';    riskMessage = 'Parent has requested schedule or teacher changes more than once. Check if the current slot is working well.'; }
    else                          { riskLevel = 'Stable';   riskMessage = 'No major disruptions detected. Learner journey looks healthy.'; }

    // ── 5. Plain-English summary ──────────────────────────────────────────
    var aiSummary = learnerProfile.learnerName + ' has no migration history on record.';
    if (total > 0) {
      var first = migrationTimeline[0];
      var last  = migrationTimeline[total - 1];
      var fd    = first.date ? new Date(first.date).toLocaleDateString('en-GB', {day:'2-digit',month:'short',year:'numeric'}) : 'unknown';
      var ld    = last.date  ? new Date(last.date).toLocaleDateString('en-GB',  {day:'2-digit',month:'short',year:'numeric'}) : 'unknown';
      aiSummary = learnerProfile.learnerName
        + ' has had ' + total + ' migration event' + (total > 1 ? 's' : '')
        + ' (' + inboundCount + ' parent-requested, ' + outboundCount + ' JetLearn-initiated)'
        + ', first recorded on ' + fd + ' and most recently on ' + ld + '.'
        + ' Current teacher: ' + learnerProfile.currentTeacher + '.'
        + ' Overall journey stability: ' + riskLevel + '.';
    }

    return {
      success:           true,
      learnerProfile:    learnerProfile,
      migrationTimeline: migrationTimeline,
      auditLogTimeline:  auditLogTimeline,
      journeyAnalysis:   {
        totalMigrations: total,
        inbound:         inboundCount,
        outbound:        outboundCount,
        riskLevel:       riskLevel,
        riskMessage:     riskMessage,
        ticketDetails:   migrationTimeline
      },
      aiSummary: aiSummary
    };

  } catch(e) {
    Logger.log('[getLearnerTimeline] FATAL: ' + e.message + '\n' + e.stack);
    return { success: false, message: e.message };
  }
}

/**
 * Fetches ALL migration tickets to analyze churn risk (Inbound vs Outbound).
 */
/**
 * UPDATED: Fetches Migration tickets.
 * Filters: Pipeline, Keywords (Kits/PRM), and Cancelled Stages.
 */
function getMigrationHistoryStats(jlid) {
  const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  const searchUrl = 'https://api.hubapi.com/crm/v3/objects/tickets/search';
  
  const MIGRATION_PIPELINE_ID = '66161281'; 

  // --- CONFIGURATION: Exclude specific Ticket Stages ---
  // You can find these IDs in your HubSpot URL when viewing the pipeline settings
  // or by inspecting a ticket in that column.
  const EXCLUDED_STAGE_IDS = [
      "133821818", // Example: ID for "Cancelled"
      "153457301"  // Example: ID for "Rejected"
  ];

  const requestBody = {
    filterGroups: [{
      filters: [
        { propertyName: "learner_uid", operator: "EQ", value: jlid },
        { propertyName: "hs_pipeline", operator: "EQ", value: MIGRATION_PIPELINE_ID }
      ]
    }],
    limit: 100, 
    sorts: [{ propertyName: "createdate", direction: "DESCENDING" }],
    properties: [ 
      "subject", 
      "createdate", 
      "reason_of_migration__t_", 
      "new_teacher", 
      "current_teacher__t_",
      "hs_pipeline_stage" // Fetch the status/stage
    ]
  };

  try {
    const response = UrlFetchApp.fetch(searchUrl, {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    });
    
    const data = JSON.parse(response.getContentText());
    if (!data.results) return { total: 0, inbound: 0, outbound: 0, events: [] };

    const events = [];
    let inboundCount = 0;
    let outboundCount = 0;

    const inboundReasons = [
      "Slot change - Learner request", "Slot change -Learner request", 
      "Teacher Affinity", "Pause Request", "Special Learning Needs"
    ];
    
    // Keywords to ignore (Logistics, Reviews, etc.)
    const ignoreKeywords = ["PRM", "Renewal", "Feedback", "Review", "Kit", "Laptop", "Device", "Tab"]; 

    data.results.forEach(t => {
      const subject = t.properties.subject || "";
      const reason = t.properties.reason_of_migration__t_ || "";
      const stage = t.properties.hs_pipeline_stage;

      // 1. FILTER: Exclude specific Stages (Cancelled/Rejected)
      if (EXCLUDED_STAGE_IDS.includes(stage)) {
          return; 
      }

      // 2. FILTER: Skip non-migration subjects (Logistics/PRMs)
      if (ignoreKeywords.some(kw => subject.includes(kw))) {
          return; 
      }

      // 3. FILTER: Must have a Migration Reason OR explicitly say "Migration"
      // This filters out empty "placeholder" tickets
      if (!reason && !subject.includes("Migration")) {
          return; 
      }

      const is_inbound = inboundReasons.includes(reason);
      if (is_inbound) inboundCount++; else outboundCount++;

      events.push({
        id: t.id,
        date: t.properties.createdate,
        reason: reason || "Unspecified Migration",
        type: is_inbound ? 'Inbound (Parent)' : 'Outbound (JetLearn)',
        from: t.properties.current_teacher__t_,
        to: t.properties.new_teacher
      });
    });

    return { total: events.length, inbound: inboundCount, outbound: outboundCount, events: events };

  } catch (e) {
    Logger.log("Error analyzing migration stats: " + e.message);
    return { total: 0, inbound:0, outbound:0, events: [] };
  }
}

function getTeacherAttritionReport(teacherName) {
  var token     = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  var searchUrl = 'https://api.hubapi.com/crm/v3/objects/deals/search';
  var PORTAL_ID = '7729491';
 
  var resolvedName = resolveTeacherName(teacherName);
  Logger.log('[getTeacherAttritionReport] "' + teacherName + '" → "' + resolvedName + '"');

  // current_teacher stores a HubSpot internal ID, not a name.
  // Look up the ID from Teacher HS values sheet first.
  var hsId = null;
  try {
    var teacherHsData = _getCachedSheetData(CONFIG.SHEETS.TEACHER_HS_DATA);
    var resolvedLower = resolvedName.trim().toLowerCase();
    for (var hi = 1; hi < teacherHsData.length; hi++) {
      var rowName = String(teacherHsData[hi][2] || '').trim().toLowerCase();
      if (rowName === resolvedLower) {
        hsId = String(teacherHsData[hi][1] || '').trim();
        Logger.log('[getTeacherAttritionReport] Found HS ID: "' + hsId + '" for "' + resolvedName + '"');
        break;
      }
    }
    if (!hsId) Logger.log('[getTeacherAttritionReport] No HS ID for "' + resolvedName + '" — falling back to name');
  } catch (e) {
    Logger.log('[getTeacherAttritionReport] HS ID lookup error: ' + e.message);
  }

  // Use EQ on ID if found, else CONTAINS_TOKEN on name as fallback
  var teacherFilter = hsId
    ? { propertyName: 'current_teacher', operator: 'EQ',             value: hsId }
    : { propertyName: 'current_teacher', operator: 'CONTAINS_TOKEN', value: resolvedName };
 
  var requestBody = {
    filterGroups: [{
      filters: [
        teacherFilter,
        { propertyName: 'learner_status', operator: 'IN', values: ['Active Learner', 'Friendly Learner', 'VIP', 'Break & Return'] }
      ]
    }],
    limit: 100,
    properties: [
      'dealname', 'jetlearner_id', 'current_course', 'module_start_date',
      'learner_status', 'dealstage', 'amount', 'subscription_tenure',
      'deal_currency_code', 'payment_type', 'installment_type', 'subscription'
    ]
  };
 
  try {
    var response = UrlFetchApp.fetch(searchUrl, {
      method:             'post',
      headers:            { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload:            JSON.stringify(requestBody),
      muteHttpExceptions: true
    });
 
    var data = JSON.parse(response.getContentText());
    if (!data.results) return { success: false, message: 'No students found.' };
 
    Logger.log('[getTeacherAttritionReport] Found ' + data.results.length + ' learners for "' + resolvedName + '"');
 
    var auditData = _getCachedSheetData(CONFIG.SHEETS.AUDIT_LOG);
 
    var students = data.results.map(function(deal) {
      var jlid           = deal.properties.jetlearner_id;
      var amount         = safeParseHubspotNumber(deal.properties.amount);
      var tenure         = safeParseHubspotNumber(deal.properties.subscription_tenure);
      var currency       = deal.properties.deal_currency_code || 'EUR';
      var planName       = (deal.properties.subscription     || '').toLowerCase();
      var rawPaymentType = (deal.properties.payment_type     || '').toLowerCase();
      var rawInstType    = (deal.properties.installment_type || '').toLowerCase();
 
      Logger.log('[AttritionReport] JLID: ' + jlid + ' | ' + deal.properties.dealname + ' | payment_type: "' + deal.properties.payment_type + '" | tenure: ' + tenure + ' | amount: ' + amount + ' | currency: ' + currency);
 
      var isInstallment = rawPaymentType.includes('installment')
                       || rawPaymentType.includes('emi')
                       || rawPaymentType.includes('recurring')
                       || rawInstType.includes('installment')
                       || rawInstType.includes('emi')
                       || rawInstType.includes('monthly')
                       || rawInstType.includes('quarterly');
      var paymentTag = isInstallment ? 'Installment' : 'Upfront';
 
      var rate = getConversionRate(currency) || 1;
      var amountEur = (currency === 'EUR') ? amount : amount / rate;
      if (amountEur > 50000) amountEur = amount * rate;
      amountEur = Math.round(amountEur);
 
      // Deal value classification
      var dealValueLabel = 'Low Value', dealValueColor = '#b91c1c', dealValueBg = '#fee2e2';
      var setHigh    = function() { dealValueLabel = 'High Value'; dealValueColor = '#15803d'; dealValueBg = '#dcfce7'; };
      var setMid     = function() { dealValueLabel = 'Mid Value';  dealValueColor = '#b45309'; dealValueBg = '#fef3c7'; };
      var setUnknown = function() { dealValueLabel = 'Unknown';    dealValueColor = '#4a5568'; dealValueBg = '#e2e8f0'; };
 
      var applyTier = function(highFloor, midFloor) {
        if      (amountEur >= highFloor) setHigh();
        else if (amountEur >= midFloor)  setMid();
      };
      var applyPerMonthTier = function() {
        if (tenure <= 0) { setUnknown(); return; }
        var pmv = amountEur / tenure;
        if      (pmv >= 119) setHigh();
        else if (pmv >= 61)  setMid();
      };
 
      if (amountEur === 0 || tenure === 0)      setUnknown();
      else if (planName.includes('gcse'))        setHigh();
      else if (isInstallment)                    applyPerMonthTier();
      else if (tenure >= 36)                     setHigh();
      else if (tenure >= 24)                     applyTier(1400, 900);
      else if (tenure >= 12)                     applyTier(899,  600);
      else if (tenure >= 6)                      applyTier(499,  300);
      else if (tenure >= 3)                      applyTier(357,  183);
      else                                       applyPerMonthTier();
 
      // Migration history from audit log
      var recentMoves = 0, lastMoveDate = null, prevTeacher = 'N/A', moveReason = 'N/A';
      if (auditData && auditData.length > 1) {
        var now = new Date();
        var threeMonthsAgo = new Date(); threeMonthsAgo.setMonth(now.getMonth() - 3);
        auditData.forEach(function(row) {
          if (String(row[2]) === jlid && String(row[1]).includes('Migration')) {
            var d = parseSheetDate(row[0]);
            if (d && d > threeMonthsAgo) recentMoves++;
            if (d && (!lastMoveDate || d > lastMoveDate)) {
              lastMoveDate = d;
              prevTeacher  = row[4]  || 'Unknown';
              moveReason   = row[10] || 'Unspecified';
            }
          }
        });
      }
 
      return {
        name:             deal.properties.dealname,
        jlid:             jlid,
        course:           getCourseLabel(deal.properties.current_course),
        status:           deal.properties.learner_status,
        hubspotLink:      'https://app.hubspot.com/contacts/' + PORTAL_ID + '/deal/' + deal.id,
        dealValueLabel:   dealValueLabel,
        dealValueColor:   dealValueColor,
        dealValueBg:      dealValueBg,
        dealAmountLocal:  amount,
        dealAmountEur:    amountEur,
        dealCurrency:     currency,
        dealTenureMonths: tenure,
        paymentTag:       paymentTag,
        recentMoves:      recentMoves,
        lastMoveDate:     lastMoveDate ? lastMoveDate.toLocaleDateString('en-GB') : 'No Record',
        prevTeacher:      prevTeacher,
        moveReason:       moveReason
      };
    });
 
    students.sort(function(a, b) { return b.recentMoves - a.recentMoves; });
    return { success: true, students: students, teacher: resolvedName };
 
  } catch (e) {
    Logger.log('Error in getTeacherAttritionReport: ' + e.message);
    return { success: false, message: e.message };
  }
}



function getMigrationHistoryStatsByTeacher(teacherName) {
  var token     = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  var searchUrl = 'https://api.hubapi.com/crm/v3/objects/tickets/search';
 
  var resolvedName = resolveTeacherName(teacherName);
  Logger.log('[getMigrationHistoryStatsByTeacher] "' + teacherName + '" → "' + resolvedName + '"');

  // current_teacher__t_ stores HS internal ID — look it up first
  var hsId = null;
  try {
    var teacherHsData = _getCachedSheetData(CONFIG.SHEETS.TEACHER_HS_DATA);
    var resolvedLower = resolvedName.trim().toLowerCase();
    for (var hi = 1; hi < teacherHsData.length; hi++) {
      if (String(teacherHsData[hi][2] || '').trim().toLowerCase() === resolvedLower) {
        hsId = String(teacherHsData[hi][1] || '').trim();
        break;
      }
    }
  } catch(e) { Logger.log('[getMigrationHistoryStatsByTeacher] ID lookup error: ' + e.message); }

  var filterValue = hsId || resolvedName;
 
  var requestBody = {
    filterGroups: [{
      filters: [{ propertyName: 'current_teacher__t_', operator: 'EQ', value: filterValue }]
    }],
    properties: ['hs_pipeline_stage', 'createdate'],
    limit: 100
  };
 
  try {
    var response = UrlFetchApp.fetch(searchUrl, {
      method:             'post',
      headers:            { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload:            JSON.stringify(requestBody),
      muteHttpExceptions: true
    });
    var data = JSON.parse(response.getContentText());
    return { total: data.total || 0 };
  } catch (e) {
    Logger.log('[getMigrationHistoryStatsByTeacher] Error: ' + e.message);
    return { total: 0 };
  }
}


function getActiveLearnersPerTeacher() {
  Logger.log('[HubSpot] getActiveLearnersPerTeacher started');

  const token      = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  const searchUrl  = 'https://api.hubapi.com/crm/v3/objects/deals/search';
  const counts     = {};
  const activeStatuses = ['Active Learner', 'Friendly Learner', 'VIP', 'Break & Return'];

  let after  = undefined;
  let page   = 0;
  const MAX_PAGES = 30;  // increased: 30×100 = 3000 max, covers ~2546 active learners

  try {
    do {
      const body = {
        filterGroups: [{
          filters: [{
            propertyName: 'learner_status',
            operator:     'IN',
            values:       activeStatuses
          }]
        }],
        properties: ['current_teacher', 'current_course'],
        limit: 100
      };
      if (after) body.after = after;

      const response = UrlFetchApp.fetch(searchUrl, {
        method:           'post',
        headers:          { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload:          JSON.stringify(body),
        muteHttpExceptions: true
      });

      const data = JSON.parse(response.getContentText());
      if (!data.results) break;

      data.results.forEach(function(deal) {
        const rawTeacher = deal.properties.current_teacher || '';
        const course     = String(deal.properties.current_course || '').toLowerCase();
        const teacher    = getTeacherLabel(rawTeacher);
        if (!teacher) return;

        const teacherNorm = teacher.trim().toLowerCase().replace(/\s+/g, ' ');
        [teacher, teacherNorm].forEach(function(key) {
          if (!counts[key]) counts[key] = { total: 0, coding: 0, math: 0 };
          counts[key].total++;
          if (course.includes('math')) { counts[key].math++; }
          else                         { counts[key].coding++; }
        });
      });

      after = data.paging && data.paging.next ? data.paging.next.after : undefined;
      page++;

    } while (after && page < MAX_PAGES);

    Logger.log('[HubSpot] getActiveLearnersPerTeacher done. Teachers found: ' + Object.keys(counts).length);
    return counts;

  } catch (e) {
    Logger.log('[HubSpot] getActiveLearnersPerTeacher error: ' + e.message);
    return {};
  }
}

function getEscalatedTeachersLast90Days() {
  var token     = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  var searchUrl = 'https://api.hubapi.com/crm/v3/objects/tickets/search';
  var MIGRATION_PIPELINE_ID = '66161281';
  var ESCALATION_REASONS = [
    'Teacher Performance Issue', 'Escalation On Teacher',
    'Escalation on Teacher', 'Escalation on Teacher Post Migration'
  ];
  var escalationMap = {}, after = undefined, page = 0, MAX_PAGES = 10;
  try {
    do {
      var requestBody = {
        filterGroups: [{ filters: [
          { propertyName: 'hs_pipeline',             operator: 'EQ', value: MIGRATION_PIPELINE_ID },
          { propertyName: 'reason_of_migration__t_', operator: 'IN', values: ESCALATION_REASONS }
        ]}],
        properties: ['current_teacher__t_', 'reason_of_migration__t_', 'createdate'],
        limit: 200, sorts: [{ propertyName: 'createdate', direction: 'DESCENDING' }]
      };
      if (after) requestBody.after = after;
      var response = UrlFetchApp.fetch(searchUrl, {
        method: 'post',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify(requestBody), muteHttpExceptions: true
      });
      var data = JSON.parse(response.getContentText());
      if (!data.results) { Logger.log('[getEscalatedTeachersLast90Days] No results page ' + page); break; }
      Logger.log('[getEscalatedTeachersLast90Days] Page ' + page + ': ' + data.results.length + ' tickets');
      data.results.forEach(function(ticket) {
        var rawTeacher    = ticket.properties.current_teacher__t_ || '';
        var teacherLabel  = getTeacherLabel(rawTeacher);
        var canonicalName = resolveTeacherName(teacherLabel);
        if (!canonicalName) return;
        var normalizedName = canonicalName.trim().toLowerCase().replace(/\s+/g, ' ');
        escalationMap[canonicalName]  = (escalationMap[canonicalName]  || 0) + 1;
        escalationMap[normalizedName] = (escalationMap[normalizedName] || 0) + 1;
        Logger.log('[getEscalatedTeachersLast90Days] +1 "' + canonicalName + '" id:' + rawTeacher);
      });
      after = data.paging && data.paging.next ? data.paging.next.after : undefined;
      page++;
    } while (after && page < MAX_PAGES);
    Logger.log('[getEscalatedTeachersLast90Days] Final map: ' + JSON.stringify(escalationMap));
    return escalationMap;
  } catch (e) {
    Logger.log('[getEscalatedTeachersLast90Days] Error: ' + e.message);
    return {};
  }
}

function getTeacherEscalationHistory(teacherName) {
  var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  var searchUrl = 'https://api.hubapi.com/crm/v3/objects/tickets/search';
  var MIGRATION_PIPELINE_ID = '66161281';
  var ESCALATION_REASONS = ['Escalation on Teacher','Escalation On Teacher','Teacher Performance Issue','Escalation on Teacher Post Migration'];
  var STAGE_LABELS = {
    '128913747':'Migration Triggered','128913748':'WIP','128913750':'WIP - TP Approval Pending',
    '128913752':'WIP - CLS Approval Pending','1030980247':'WIP - Rejected by CLS',
    '133755411':'WIP - Approved by CLS','1065336836':'Execution Pending',
    '128913749':'WIP - Pr Approval Pending','128913753':'Migration Completed'
  };
  var resolvedName = resolveTeacherName(teacherName);
  Logger.log('[getTeacherEscalationHistory] "' + teacherName + '" → "' + resolvedName + '"');
  var hsId = null;
  try {
    var teacherHsData = _getCachedSheetData(CONFIG.SHEETS.TEACHER_HS_DATA);
    var resolvedLower = resolvedName.trim().toLowerCase();
    for (var hi = 1; hi < teacherHsData.length; hi++) {
      if (String(teacherHsData[hi][2] || '').trim().toLowerCase() === resolvedLower) {
        hsId = String(teacherHsData[hi][1] || '').trim();
        Logger.log('[getTeacherEscalationHistory] HS ID: "' + hsId + '"');
        break;
      }
    }
  } catch(e) { Logger.log('[getTeacherEscalationHistory] ID lookup error: ' + e.message); }

  var teacherFilter = hsId
    ? { propertyName: 'current_teacher__t_', operator: 'EQ', value: hsId }
    : { propertyName: 'current_teacher__t_', operator: 'EQ', value: resolvedName };

  var requestBody = {
    filterGroups: [{ filters: [
      { propertyName: 'hs_pipeline', operator: 'EQ', value: MIGRATION_PIPELINE_ID },
      teacherFilter
    ]}],
    properties: ['subject','current_teacher__t_','reason_of_migration__t_','createdate',
                 'migration_completed_date','hs_pipeline_stage','learner_full_name','learner_uid','hs_ticket_id'],
    limit: 200, sorts: [{ propertyName: 'createdate', direction: 'DESCENDING' }]
  };
  try {
    var response = UrlFetchApp.fetch(searchUrl, {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(requestBody), muteHttpExceptions: true
    });
    var data = JSON.parse(response.getContentText());
    if (!data.results) return { success: true, totalCount: 0, byReason: {}, tickets: [], lastEscalationDate: null };
    Logger.log('[getTeacherEscalationHistory] Total tickets: ' + data.results.length);
    var escalationTickets = data.results.filter(function(t) {
      var r = String(t.properties.reason_of_migration__t_ || '').trim();
      return ESCALATION_REASONS.some(function(c) { return r.toLowerCase() === c.toLowerCase(); });
    });
    var byReason = {}, tickets = [], lastEscalationDate = null;
    escalationTickets.forEach(function(ticket) {
      var props = ticket.properties;
      var reason = String(props.reason_of_migration__t_ || '').trim();
      var triggeredDate = ticket.createdAt ? new Date(ticket.createdAt) : (props.createdate ? new Date(props.createdate) : null);
      var completedDate = props.migration_completed_date ? new Date(props.migration_completed_date) : null;
      var stageId = String(props.hs_pipeline_stage || '').trim();
      var stageLabel = STAGE_LABELS[stageId] || ('Stage ' + stageId);
      var isCompleted = stageLabel === 'Migration Completed';
      var daysToResolve = (triggeredDate && completedDate && isCompleted)
        ? Math.round((completedDate - triggeredDate) / 86400000) : null;
      if (triggeredDate && (!lastEscalationDate || triggeredDate > lastEscalationDate)) lastEscalationDate = triggeredDate;
      byReason[reason] = (byReason[reason] || 0) + 1;
      tickets.push({
        ticketId: ticket.id || props.hs_ticket_id || 'N/A',
        ticketName: props.subject || 'N/A', learnerName: props.learner_full_name || 'N/A',
        learnerUid: props.learner_uid || 'N/A', reason: reason, stageId: stageId,
        status: stageLabel, isCompleted: isCompleted,
        triggeredDate: triggeredDate ? triggeredDate.toLocaleDateString('en-GB') : 'N/A',
        completedDate: completedDate ? completedDate.toLocaleDateString('en-GB') : 'N/A',
        daysToResolve: daysToResolve !== null ? daysToResolve : 'N/A'
      });
    });
    tickets.sort(function(a,b) {
      var da = a.triggeredDate!=='N/A'?new Date(a.triggeredDate.split('/').reverse().join('-')):new Date(0);
      var db = b.triggeredDate!=='N/A'?new Date(b.triggeredDate.split('/').reverse().join('-')):new Date(0);
      return db-da;
    });
    return { success:true, totalCount:escalationTickets.length, byReason:byReason, tickets:tickets,
             lastEscalationDate: lastEscalationDate ? lastEscalationDate.toLocaleDateString('en-GB') : null };
  } catch(e) {
    Logger.log('[getTeacherEscalationHistory] Error: ' + e.message);
    return { success: false, message: e.message };
  }
}
// ── Fast accurate total active learner count from HubSpot ──────────────────
// Uses HubSpot's own total count field — no pagination needed, always accurate
// regardless of how many learners there are
function getTotalActiveLearnerCount() {
  const token      = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  const searchUrl  = 'https://api.hubapi.com/crm/v3/objects/deals/search';
  const activeStatuses = ['Active Learner', 'Friendly Learner', 'VIP', 'Break & Return'];
  try {
    const body = {
      filterGroups: [{ filters: [{ propertyName: 'learner_status', operator: 'IN', values: activeStatuses }] }],
      properties: ['hs_object_id'],
      limit: 1  // We only need the total count, not the actual records
    };
    const response = UrlFetchApp.fetch(searchUrl, {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    const data = JSON.parse(response.getContentText());
    if (data.error) throw new Error(data.message || 'HubSpot error');
    const total = data.total || 0;
    Logger.log('[getTotalActiveLearnerCount] Total: ' + total);
    return { success: true, total: total };
  } catch (e) {
    Logger.log('[getTotalActiveLearnerCount] error: ' + e.message);
    return { success: false, total: 0 };
  }
}