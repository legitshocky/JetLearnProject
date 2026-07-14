// ============================================================
// ONBOARDING CHECKLIST SERVICE
// Payment Received deals → send WATI welcome + post HubSpot note
// + update timelines_chat_link + learner_practice_document_link
// Stages: closedwon (Coding) | 208827055 (Maths)
// ============================================================

var OBC_STAGES   = ['closedwon', '208827055'];
var OBC_TEMPLATE = 'welcome_ob_u';

// ── Subject from JLID suffix ──────────────────────────────
function _obcSubject(jlid) {
  var m = String(jlid || '').match(/JL\d+([A-Z0-9]+)$/i);
  var suffix = m ? m[1].toUpperCase() : '';
  var map = { 'C': 'Coding', 'C2': 'Coding', 'M': 'Maths', 'F': 'Financial Literacy', 'FL': 'Financial Literacy' };
  return map[suffix] || 'Coding';
}

// ── Format HubSpot date value → DD-MM-YYYY ────────────────
function _obcFmtDate(val) {
  if (!val) return '';
  var d;
  if (val instanceof Date) {
    d = val;
  } else if (/^\d{10,}$/.test(String(val).trim())) {
    d = new Date(parseInt(val, 10));
  } else {
    d = new Date(val);
  }
  if (isNaN(d.getTime())) return String(val);
  var dd = String(d.getDate()).padStart(2, '0');
  var mm = String(d.getMonth() + 1).padStart(2, '0');
  return dd + '-' + mm + '-' + d.getFullYear();
}

// ── Clean phone (digits only) ─────────────────────────────
function _obcCleanPhone(p) {
  return String(p || '').replace(/[^\d]/g, '');
}

// ── Shorten HubSpot timezone string → IST / CET / PST etc ─
function _obcShortTz(tzString) {
  if (!tzString) return '';
  var t = tzString.toLowerCase();
  if (t.indexOf('bombay') !== -1 || t.indexOf('calcutta') !== -1 || t.indexOf('madras') !== -1 ||
      t.indexOf('new delhi') !== -1 || t.indexOf('kolkata') !== -1 || t.indexOf('mumbai') !== -1) return 'IST';
  if (t.indexOf('karachi') !== -1 || t.indexOf('islamabad') !== -1 || t.indexOf('lahore') !== -1)  return 'PKT';
  if (t.indexOf('dhaka') !== -1)                                                                    return 'BST';
  if (t.indexOf('colombo') !== -1)                                                                  return 'SLST';
  if (t.indexOf('london') !== -1 || t.indexOf('dublin') !== -1 || t.indexOf('edinburgh') !== -1)   return 'UKT';
  if (t.indexOf('paris') !== -1 || t.indexOf('brussels') !== -1 || t.indexOf('berlin') !== -1 ||
      t.indexOf('amsterdam') !== -1 || t.indexOf('madrid') !== -1 || t.indexOf('rome') !== -1 ||
      t.indexOf('vienna') !== -1 || t.indexOf('stockholm') !== -1 || t.indexOf('copenhagen') !== -1) return 'CET';
  if (t.indexOf('athens') !== -1 || t.indexOf('bucharest') !== -1 || t.indexOf('helsinki') !== -1) return 'EET';
  if (t.indexOf('abu dhabi') !== -1 || t.indexOf('dubai') !== -1 || t.indexOf('muscat') !== -1)    return 'GST';
  if (t.indexOf('riyadh') !== -1 || t.indexOf('kuwait') !== -1)                                    return 'AST';
  if (t.indexOf('singapore') !== -1 || t.indexOf('kuala lumpur') !== -1)                           return 'SGT';
  if (t.indexOf('sydney') !== -1 || t.indexOf('melbourne') !== -1 || t.indexOf('canberra') !== -1 ||
      t.indexOf('brisbane') !== -1)                                                                  return 'AEST';
  if (t.indexOf('darwin') !== -1 || t.indexOf('adelaide') !== -1)                                  return 'ACST';
  if (t.indexOf('eastern time') !== -1 || t.indexOf('new york') !== -1 || t.indexOf('toronto') !== -1) return 'EST';
  if (t.indexOf('central time') !== -1 || t.indexOf('chicago') !== -1)                             return 'CST';
  if (t.indexOf('mountain time') !== -1 || t.indexOf('denver') !== -1 || t.indexOf('phoenix') !== -1) return 'MST';
  if (t.indexOf('pacific time') !== -1 || t.indexOf('los angeles') !== -1 || t.indexOf('seattle') !== -1) return 'PST';
  if (t.indexOf('alaska') !== -1)                                                                   return 'AKST';
  if (t.indexOf('hawaii') !== -1)                                                                   return 'HST';
  if (t.indexOf('tokyo') !== -1 || t.indexOf('osaka') !== -1)                                      return 'JST';
  if (t.indexOf('beijing') !== -1 || t.indexOf('hong kong') !== -1 || t.indexOf('taipei') !== -1) return 'HKT';
  if (t.indexOf('auckland') !== -1 || t.indexOf('wellington') !== -1)                              return 'NZST';
  if (t.indexOf('nairobi') !== -1)                                                                  return 'EAT';
  if (t.indexOf('johannesburg') !== -1 || t.indexOf('pretoria') !== -1)                            return 'SAST';
  // Fallback: extract UTC offset from string
  var m = tzString.match(/GMT\s*([+-]?\d+(?::\d+)?)/i);
  if (m) return 'UTC' + m[1];
  return tzString;
}

// ── Check if teacher is upskilled on a course ────────────
function _obcCheckTeacherUpskilled(teacherName, courseName) {
  if (!teacherName || !courseName) return 'Unknown';
  try {
    var sheetData = _getCachedSheetData(CONFIG.SHEETS.TEACHER_COURSES);
    if (!sheetData || sheetData.length < 2) return 'Unknown';

    // Find header row (first row where col[0] === 'teacher')
    var headerRowIndex = 0;
    for (var i = 0; i < Math.min(sheetData.length, 10); i++) {
      if (String(sheetData[i][0]).trim().toLowerCase() === 'teacher') {
        headerRowIndex = i;
        break;
      }
    }
    var headers = sheetData[headerRowIndex];
    var COURSE_START = 4;

    // Find course column (case-insensitive)
    var courseColIndex = -1;
    for (var j = COURSE_START; j < headers.length; j++) {
      if (String(headers[j] || '').trim().toLowerCase() === courseName.trim().toLowerCase()) {
        courseColIndex = j;
        break;
      }
    }
    if (courseColIndex === -1) return 'Course not in sheet';

    // Find teacher row (case-insensitive)
    var teacherLow = teacherName.trim().toLowerCase();
    for (var r = headerRowIndex + 1; r < sheetData.length; r++) {
      if (String(sheetData[r][0] || '').trim().toLowerCase() === teacherLow) {
        var val = String(sheetData[r][courseColIndex] || '').trim();
        if (!val || val.toLowerCase() === 'not onboarded' || val === '') return 'No';
        return 'Yes (' + val + ')';
      }
    }
    return 'Teacher not in sheet';
  } catch(e) {
    Logger.log('[OBC] _obcCheckTeacherUpskilled error: ' + e.message);
    return 'Unknown';
  }
}

// ── PATCH HubSpot deal properties ─────────────────────────
function _obcPatchDeal(dealId, props) {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    var resp  = monitoredFetch('https://api.hubapi.com/crm/v3/objects/deals/' + dealId, {
      method:  'PATCH',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ properties: props }),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    Logger.log('[OBC] PATCH deal ' + dealId + ' → HTTP ' + code + ' props=' + JSON.stringify(props));
    if (code !== 200 && code !== 204) {
      Logger.log('[OBC] PATCH error body: ' + resp.getContentText().substring(0, 400));
    }
    return (code === 200 || code === 204);
  } catch(e) {
    Logger.log('[OBC] _obcPatchDeal ERROR: ' + e.message);
    return false;
  }
}

// ── Classes offered = SPW × tenure(months) × 4.33 ────────
function _obcClassesOffered(spwRaw, tenureRaw) {
  var spw    = parseInt(String(spwRaw || '1').match(/\d+/) || ['1']);
  var tenure = parseInt(tenureRaw || '0');
  if (spw > 0 && tenure > 0) return String(Math.round(spw * tenure * 4.33));
  return '';
}

// ── Fetch WATI contact ID → build proper teamInbox chat URL ─
// Expected format: https://live.wati.io/400399/teamInbox/<contactId>
function _obcGetWatiChatLink(cleanPhone) {
  try {
    var props   = PropertiesService.getScriptProperties();
    var base    = (props.getProperty('WATI_API_ENDPOINT') || '').trim().replace(/\/$/, '');
    var token   = (props.getProperty('WATI_ACCESS_TOKEN')  || '').trim();
    if (!base || !token) { Logger.log('[OBC] WATI config missing for chat link'); return ''; }
    if (!token.startsWith('Bearer ')) token = 'Bearer ' + token;

    // Try lookup with phone as-is, then with leading + stripped, then with + added
    var phonesToTry = [cleanPhone];
    if (cleanPhone.startsWith('+')) {
      phonesToTry.push(cleanPhone.replace(/^\+/, ''));
    } else {
      phonesToTry.push('+' + cleanPhone);
    }

    for (var pi = 0; pi < phonesToTry.length; pi++) {
      var phone = phonesToTry[pi];
      var resp = monitoredFetch(base + '/api/v1/contact/' + encodeURIComponent(phone), {
        method:  'get',
        headers: { 'Authorization': token },
        muteHttpExceptions: true
      });
      var httpCode = resp.getResponseCode();
      Logger.log('[OBC] WATI contact lookup phone="' + phone + '" HTTP ' + httpCode);
      if (httpCode !== 200) continue;

      var body = resp.getContentText();
      Logger.log('[OBC] WATI contact response: ' + body.substring(0, 300));
      var json = JSON.parse(body);
      // WATI API response: { result: true, contact: { id: "...", wAid: "...", ... } }
      var contactId = (json.contact && json.contact.id) ? json.contact.id
                    : (json.id)                          ? json.id
                    : '';
      if (!contactId) {
        Logger.log('[OBC] WATI contact id not found in response for phone="' + phone + '"');
        continue;
      }
      // base already contains account number, e.g. https://live.wati.io/400399
      var chatLink = base + '/teamInbox/' + contactId;
      Logger.log('[OBC] WATI chat link built: ' + chatLink);
      return chatLink;
    }
    Logger.log('[OBC] WATI contact not found for any phone variant of: ' + cleanPhone);
    return '';
  } catch(e) {
    Logger.log('[OBC] _obcGetWatiChatLink ERROR: ' + e.message);
    return '';
  }
}

// ── Public: get all Payment Received deals ────────────────
function getPaymentReceivedDeals() {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    var body  = {
      filterGroups: [{
        filters: [{
          propertyName: 'dealstage',
          operator:     'IN',
          values:       OBC_STAGES
        }]
      }],
      properties: [
        'dealname', 'jetlearner_id', 'amount', 'deal_currency_code', 'hs_object_id',
        'current_course__t_', 'current_teacher__t_',
        'current_course',     'current_teacher',
        'parent_name', 'parent_email', 'phone_number_deal_',
        'class_timings', 'time_zone', 'frequency_of_classes',
        'subscription_tenure', 'stage____payment_trigger_date', 'dealstage',
        'timelines_chat_link', 'learner_practice_document_link'
      ],
      sorts: [{ propertyName: 'stage____payment_trigger_date', direction: 'DESCENDING' }],
      limit: 100
    };

    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/deals/search', {
      method:  'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      return { success: false, error: 'HubSpot error ' + resp.getResponseCode() };
    }

    var json  = JSON.parse(resp.getContentText());
    var deals = (json.results || []).map(function(r) {
      var p    = r.properties;
      var jlid = p.jetlearner_id || '';
      // Use whichever teacher/course property has data (__t_ = pipeline-specific, plain = standard)
      var rawTeacher = p.current_teacher__t_ || p.current_teacher || '';
      var rawCourse  = p.current_course__t_  || p.current_course  || '';
      // getTeacherLabel resolves HubSpot user ID → display name; falls back to name as-is
      var teacherResolved = rawTeacher ? (getTeacherLabel(rawTeacher) || rawTeacher) : '';
      // getCourseLabel resolves internal course value → display label
      var courseResolved  = rawCourse  ? (getCourseLabel(rawCourse)  || rawCourse)  : '';
      var shortTz = _obcShortTz(p.time_zone || '');
      return {
        dealId:          p.hs_object_id        || '',
        jlid:            jlid,
        learnerName:     p.dealname            || '',
        parentName:      p.parent_name         || '',
        parentEmail:     p.parent_email        || '',
        phone:           p.phone_number_deal_  || '',
        course:          courseResolved,
        teacher:         teacherResolved,
        teacherRaw:      rawTeacher,
        classTimings:    p.class_timings       || '',
        timezone:        shortTz || p.time_zone || '',
        timezoneRaw:     p.time_zone           || '',
        amount:          p.amount              || '',
        currency:        p.deal_currency_code  || '',
        paymentDate:     _obcFmtDate(p.stage____payment_trigger_date),
        classesOffered:  _obcClassesOffered(p.frequency_of_classes, p.subscription_tenure),
        subject:         _obcSubject(jlid),
        watiChatLink:    p.timelines_chat_link            || '',
        practiceDocLink: p.learner_practice_document_link || '',
        dealstage:       p.dealstage                      || ''
      };
    });

    Logger.log('[OBC] getPaymentReceivedDeals → ' + deals.length + ' deals');
    return { success: true, deals: deals };

  } catch(e) {
    Logger.log('[OBC] getPaymentReceivedDeals ERROR: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── Public: run full onboarding checklist for one deal ────
// extraData = { classesOffered: '4', classTimings: 'Monday 10:00 AM' }
// Returns { success, watiSent, chatLinkUpdated, noteSent, practiceDocUrl, errors[] }
function runOnboardingChecklist(dealId, extraData) {
  var errors = [];
  var watiSent = false, chatLinkUpdated = false, noteSent = false, practiceDocUrl = '';
  extraData = extraData || {};

  try {
    // ── Fetch deal ────────────────────────────────────────
    var token  = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    var params = 'properties=dealname,jetlearner_id,amount,deal_currency_code,' +
                 'current_course__t_,current_teacher__t_,parent_name,parent_email,' +
                 'phone_number_deal_,class_timings,time_zone,frequency_of_classes,' +
                 'subscription_tenure,stage____payment_trigger_date,learner_practice_document_link';

    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/deals/' + dealId + '?' + params, {
      method:  'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      return { success: false, errors: ['Cannot fetch deal: HTTP ' + resp.getResponseCode()] };
    }

    var p = JSON.parse(resp.getContentText()).properties;

    var jlid         = p.jetlearner_id      || '';
    var learnerName  = p.dealname           || '';
    var parentName   = p.parent_name        || '';
    var parentEmail  = p.parent_email       || '';
    var phone        = _obcCleanPhone(p.phone_number_deal_);
    var subject      = _obcSubject(jlid);

    // UI-editable values take priority over HubSpot stored values
    var course   = extraData.course   || p.current_course__t_  || '';
    var teacher  = extraData.teacherName || p.current_teacher__t_ || '';
    var timezone = p.time_zone || '';

    var classesOffered = extraData.classesOffered
      ? String(extraData.classesOffered)
      : _obcClassesOffered(p.frequency_of_classes, p.subscription_tenure);

    var classTimings = extraData.classTimings
      ? String(extraData.classTimings)
      : (p.class_timings || 'TBD');

    // Payment date
    var payDateFmt = p.stage____payment_trigger_date
      ? _obcFmtDate(p.stage____payment_trigger_date)
      : _obcFmtDate(new Date());

    // Payment amount: distinguish token vs full
    var dealTotal    = String(p.amount || '').replace(/[^0-9.]/g, '');
    var currency     = p.deal_currency_code || '';
    var receivedRaw  = extraData.amountReceived ? String(extraData.amountReceived).replace(/[^0-9.]/g, '') : dealTotal;
    var paymentType  = extraData.paymentType || 'full';
    // Auto-detect token: received < total
    var isToken = paymentType === 'token' ||
                  (receivedRaw && dealTotal && parseFloat(receivedRaw) < parseFloat(dealTotal));
    var pendingAmt   = (isToken && dealTotal && receivedRaw)
      ? (parseFloat(dealTotal) - parseFloat(receivedRaw)).toFixed(2).replace(/\.00$/, '') + ' ' + currency + ' Pending'
      : '';
    var amountLabel  = receivedRaw + ' ' + currency +
                       (isToken ? ' (Token — ' + dealTotal + ' ' + currency + ' Total' + (pendingAmt ? ', ' + pendingAmt : '') + ')' : ' (Full Payment)');

    // Teacher label: "TJL1280 - Name"
    var teacherLabel = _pdTeacherIdName(teacher) || teacher;

    // Timezone: UI override → auto-detect → raw
    var shortTz = extraData.timezone || _obcShortTz(timezone) || timezone;

    // Teacher upskilled check
    var teacherUpskilled = _obcCheckTeacherUpskilled(teacher, course);

    // ── 1. Send WATI welcome ──────────────────────────────
    if (phone) {
      try {
        var wRes = sendWatiMessage(phone, OBC_TEMPLATE, [
          { name: 'Parent',  value: parentName  },
          { name: 'Learner', value: learnerName },
          { name: 'text',    value: subject     }
        ]);
        watiSent = !!(wRes && wRes.success);
        if (!watiSent) {
          errors.push('WATI failed: ' + (wRes && wRes.error ? wRes.error : 'Unknown'));
        }
      } catch(we) {
        errors.push('WATI error: ' + we.message);
      }

      // ── 2. Fetch WATI conversation ID → proper teamInbox chat URL ──
      // Wait 3s so WATI registers the new conversation before we fetch the link
      try {
        Utilities.sleep(3000);
        var linkRes = fetchWatiDirectLink(phone);
        // Retry once if first attempt fails (WATI sometimes needs extra time)
        if (!linkRes || !linkRes.success) {
          Utilities.sleep(3000);
          linkRes = fetchWatiDirectLink(phone);
        }
        if (linkRes && linkRes.success && linkRes.link) {
          Logger.log('[OBC] WATI chat link: ' + linkRes.link);
          chatLinkUpdated = _obcPatchDeal(dealId, { timelines_chat_link: linkRes.link });
          if (!chatLinkUpdated) errors.push('Chat link PATCH failed');
        } else {
          var linkErr = (linkRes && linkRes.message) ? linkRes.message : 'no conversationId found';
          Logger.log('[OBC] Chat link skipped: ' + linkErr);
          errors.push('Chat link skipped: ' + linkErr);
        }
      } catch(cle) {
        errors.push('Chat link error: ' + cle.message);
      }
    } else {
      errors.push('No phone on deal — WATI skipped');
    }

    // ── 3. Create practice doc + save link to HubSpot ────
    var existingDocLink = p.learner_practice_document_link || '';
    // Resolve teacher to display name for createPracticeDoc (internal HS ID won't match sheet lookup)
    var teacherForDoc = teacherLabel || teacher;
    // Strip "TJL1280 - " prefix if present — _pdTeacherEmail needs just the name
    var teacherNameOnly = teacherForDoc.replace(/^TJL\d+\s*-\s*/i, '').trim() || teacher;

    if (!existingDocLink && jlid && learnerName) {
      try {
        var pdRes = createPracticeDoc(jlid, learnerName, teacherNameOnly, parentEmail);
        if (pdRes && pdRes.success && pdRes.url) {
          practiceDocUrl = pdRes.url;
          // Save link to HubSpot deal property
          var docPatchOk = _obcPatchDeal(dealId, { learner_practice_document_link: practiceDocUrl });
          Logger.log('[OBC] Practice doc created: ' + practiceDocUrl + ' | HS patch=' + docPatchOk);
          if (!docPatchOk) errors.push('Practice doc link PATCH to HubSpot failed');
        } else {
          errors.push('Practice doc create failed: ' + (pdRes && pdRes.error ? pdRes.error : 'Unknown'));
        }
      } catch(pde) {
        errors.push('Practice doc error: ' + pde.message);
      }
    } else if (existingDocLink) {
      practiceDocUrl = existingDocLink;
      // Update teacher on existing doc (teacher may differ from when doc was first created)
      if (teacherNameOnly) {
        _updateExistingPracticeDocTeacher(existingDocLink, teacherNameOnly);
      }
      // Re-patch to ensure HubSpot property is in sync
      _obcPatchDeal(dealId, { learner_practice_document_link: practiceDocUrl });
    }

    // ── 4. Post onboarding note ───────────────────────────
    try {
      var noteLines = [
        'Payment Received : '    + amountLabel,
        'Date : '                + payDateFmt,
        'Course Onboarded on : ' + course,
        'Current Teacher : '     + teacherLabel,
        'Teacher Upskilled : '   + teacherUpskilled,
        'Time Zone : '           + shortTz,
        'Classes Offered : '     + classesOffered,
        'Classes day & Time : '  + classTimings
      ];
      if (practiceDocUrl) {
        noteLines.push('Practice Doc : ' + practiceDocUrl);
      }
      addNoteToHubSpotDeal(dealId, noteLines.join('\n'));
      noteSent = true;
    } catch(ne) {
      errors.push('Note error: ' + ne.message);
    }

    Logger.log('[OBC] deal=' + dealId + ' wati=' + watiSent + ' note=' + noteSent + ' chatLink=' + chatLinkUpdated + ' practiceDoc=' + practiceDocUrl);
    return {
      success:          true,
      watiSent:         watiSent,
      chatLinkUpdated:  chatLinkUpdated,
      noteSent:         noteSent,
      practiceDocUrl:   practiceDocUrl,
      errors:           errors
    };

  } catch(e) {
    Logger.log('[OBC] runOnboardingChecklist ERROR: ' + e.message);
    return { success: false, errors: [e.message] };
  }
}

// ── Public: update practice doc link property on deal ─────
function updatePracticeDocOnDeal(dealId, url) {
  try {
    if (!dealId || !url) return { success: false, error: 'dealId and url required' };
    var ok = _obcPatchDeal(dealId, { learner_practice_document_link: url });
    Logger.log('[OBC] updatePracticeDocOnDeal deal=' + dealId + ' ok=' + ok);
    return { success: ok };
  } catch(e) {
    Logger.log('[OBC] updatePracticeDocOnDeal ERROR: ' + e.message);
    return { success: false, error: e.message };
  }
}
