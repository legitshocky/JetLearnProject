// =============================================
// FILE: ParentWillBuyService.js
// Parent Will Buy Kit — automated follow-up system
// =============================================

// ── Column map (1-based, matches "Parent_will_buy" sheet tab) ──────────────
var PWB_COL = {
  SR_NO:             1,   // A
  DATE:              2,   // B — auto-filled at entry
  MONTH:             3,   // C — auto-filled at entry
  JLID:              4,   // D — team fills
  LEARNER_NAME:      5,   // E — auto from HubSpot
  PARENT_NAME:       6,   // F — auto from HubSpot
  PARENT_PHONE:      7,   // G — auto from HubSpot
  COURSE_NAME:       8,   // H — team fills
  KIT:               9,   // I — auto from course mapping
  COURSE_START_DATE: 10,  // J — team fills (DD/MM/YYYY or Date)
  AMAZON_LINK:       11,  // K — team fills
  STATUS:            12,  // L — system updates
  INITIAL_SENT_AT:   13,  // M
  FUP1_SENT_AT:      14,  // N
  FUP2_SENT_AT:      15,  // O
  FINAL_FUP_SENT_AT: 16,  // P
  PARENT_RESPONSE:   17,  // Q
  ESCALATED:         18,  // R
  ESCALATED_AT:      19,  // S
  INTERVAL:          20,  // T — locked at initial send
  ENTRY_BY:          21,  // U — team member who added the entry
  EMAIL_LOG:         22,  // V — "TO:email | CC:email,email" written when escalation email sent
  PROMISED_DATE:     23,  // W — AI-extracted date parent promised to buy (YYYY-MM-DD)
  AI_INTENT:         24,  // X — AI classification of last parent reply
  RESPONSE_AT:       25   // Y — timestamp of first parent response (for avg response-time KPI)
};

// ── Kit → Course mapping ───────────────────────────────────────────────────
var PWB_COURSE_KIT_MAP = {
  'Immersive AR and VR Modeling (CoSpaces)':       'VR Headset',
  'Immersive VR Experiences with Javascript':       'VR Headset',
  'Robotics with Microbit (Jr)':                   'Microbit',
  'Programming with Robotics (Microbit)':          'Microbit',
  'Advanced Programming with Robotics (Microbit)': 'Microbit',
  'Advanced Robotics using Makey Makey':           'Makey Makey'
};

var PWB_TERMINAL_STATUSES = ['Order Placed', 'Kit Received', "Parent Didn't Buy - Roadmap Changed"];

// ── HubSpot status values (must match exact enum values in HubSpot) ─────────
// Set script property PWB_HS_STATUS_PROP to override the property internal name.
// Run discoverPWBHubspotProperty() once to find the correct internal name.
var PWB_HS_STATUSES = {
  REMINDER_1:    'Reminder 1 sent',
  REMINDER_2:    'Reminder 2 sent',
  FINAL:         'Final reminder sent',
  BOUGHT:        'Parent bought it',
  ESCALATED:     'Escalated to CLS',
  ROADMAP:       "Parent didn't buy - Roadmap changed"
};

// ── Kit → HubSpot property map (confirmed from discoverPWBHubspotProperty) ──
var PWB_KIT_HS_PROP = {
  'VR Headset':  'vr_headset__oculus_status',
  'Microbit':    'microbit_kit_status',
  'Makey-Makey': 'makey_makey_kit_status',
  'Makey Makey': 'makey_makey_kit_status',
  'Arduino':     'arduino_kit_status'
};

// HubSpot's roadmap-changed enum internal value differs per property:
// microbit/VR have no space before the hyphen, makey has one, arduino lacks the option.
var PWB_ROADMAP_HS_VALUE = {
  'vr_headset__oculus_status': "Parent didn't buy- Roadmap changed",
  'microbit_kit_status':       "Parent didn't buy- Roadmap changed",
  'makey_makey_kit_status':    "Parent didn't buy - Roadmap changed"
};

// ── AI reply analysis — classify parent's free-text WhatsApp reply ─────────
// Returns { intent: 'BOUGHT'|'WILL_BUY'|'WONT_BUY'|'UNCLEAR', promisedDate: 'YYYY-MM-DD'|null }
function _pwbAnalyzeReply(replyText, courseStartDate) {
  try {
    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var prompt =
      'You classify a parent\'s WhatsApp reply about buying a learning kit for their child\'s course.\n' +
      'Today is ' + todayStr + '. The course starts on ' +
      (courseStartDate ? Utilities.formatDate(courseStartDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'unknown') + '.\n\n' +
      'Parent\'s reply: "' + String(replyText).replace(/"/g, "'") + '"\n\n' +
      'Classify the intent:\n' +
      '- BOUGHT: parent says they already ordered/bought/received the kit\n' +
      '- WILL_BUY: parent intends to buy (extract the promised date if any — "this weekend", "after salary", "next Monday" → resolve to an actual date; if intent but no date, promisedDate is null)\n' +
      '- WONT_BUY: parent clearly declines, cannot afford, or wants a different course\n' +
      '- UNCLEAR: anything else (questions, unrelated, ambiguous)\n\n' +
      'Reply with ONLY this JSON: {"intent":"...","promisedDate":"YYYY-MM-DD" or null}';
    var raw = _callGemini([{ role: 'user', parts: [{ text: prompt }] }], null, 512);
    var parsed = JSON.parse(raw);
    var intent = String(parsed.intent || 'UNCLEAR').toUpperCase();
    if (['BOUGHT','WILL_BUY','WONT_BUY','UNCLEAR'].indexOf(intent) === -1) intent = 'UNCLEAR';
    var pd = parsed.promisedDate && /^\d{4}-\d{2}-\d{2}$/.test(String(parsed.promisedDate)) ? String(parsed.promisedDate) : null;
    return { intent: intent, promisedDate: pd };
  } catch(e) {
    Logger.log('[PWB] _pwbAnalyzeReply AI error: ' + e.message);
    return { intent: 'UNCLEAR', promisedDate: null };
  }
}

// ── PATCH HubSpot deal property for PWB status ─────────────────────────────
function _updateHubspotPWBStatus(dealId, statusValue, kitName) {
  if (!dealId || !statusValue) return;
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY') || '';
    if (!token) { Logger.log('[PWB] No HUBSPOT_API_KEY — skipping HS status update'); return; }

    var propName = kitName ? (PWB_KIT_HS_PROP[kitName] || '') : '';
    if (!propName) {
      Logger.log('[PWB] No HS property mapped for kit "' + kitName + '" — skipping');
      return;
    }

    var payload = { properties: {} };
    payload.properties[propName] = statusValue;

    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/deals/' + dealId, {
      method:             'PATCH',
      headers:            { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code === 200) {
      Logger.log('[PWB] HubSpot deal ' + dealId + ' → ' + propName + ' = "' + statusValue + '" (kit: ' + kitName + ')');
    } else {
      Logger.log('[PWB] HubSpot PATCH failed (' + code + '): ' + resp.getContentText());
    }
  } catch(e) {
    Logger.log('[PWB] _updateHubspotPWBStatus ERROR: ' + e.message);
  }
}

// ── Discovery helper — run once in GAS editor to find the correct property name
// Look for a property whose enum values include "Reminder 1 sent" ─────────────
function discoverPWBHubspotProperty() {
  var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY') || '';
  if (!token) { Logger.log('[PWB] No HUBSPOT_API_KEY'); return; }

  var resp = UrlFetchApp.fetch(
    'https://api.hubapi.com/crm/v3/properties/deals?limit=1000',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  var data = JSON.parse(resp.getContentText());
  var results = (data.results || []).filter(function(p) {
    var opts = (p.options || []).map(function(o){ return o.label || o.value || ''; });
    return opts.some(function(v){ return v.indexOf('Reminder') > -1 || v.indexOf('Parent will buy') > -1; });
  });
  results.forEach(function(p) {
    Logger.log('[PWB] Candidate property: "' + p.name + '" label="' + p.label + '" options=' +
      JSON.stringify((p.options||[]).map(function(o){return o.value;})));
  });
  if (!results.length) Logger.log('[PWB] No matching properties found. Check HubSpot property names manually.');
}

// ── Sheet accessor ─────────────────────────────────────────────────────────
function _getPWBSheet() {
  var sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID_KIT_TRACKING')
                || '17Jsa2Kl2AkI5SgtlITYGqb-Q-PxfzkNFzzG5HwBJp_Q';
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('Parent_will_buy');
  if (!sheet) throw new Error('[PWB] Sheet "Parent_will_buy" not found in spreadsheet ' + sheetId);
  return sheet;
}

// ═══════════════════════════════════════════════════════════════════════════
// KIT LINKS — per-country Amazon (or local store) purchase links
// Sheet "Kit Links": Kit | Marketplace | Countries | Link
// Fill the Link column once; PWB auto-picks by learner country.
// "Other" marketplace rows cover countries with no Amazon — set Countries to a
// comma-separated country list and Link to the local store URL.
// ═══════════════════════════════════════════════════════════════════════════

// Amazon marketplace → countries it serves (lowercase, comma-matched loosely)
var PWB_AMAZON_MARKETS = [
  { market: 'amazon.com',     countries: 'usa, united states, us' },
  { market: 'amazon.ca',      countries: 'canada' },
  { market: 'amazon.com.mx',  countries: 'mexico' },
  { market: 'amazon.com.br',  countries: 'brazil' },
  { market: 'amazon.co.uk',   countries: 'uk, united kingdom, england, scotland, wales, great britain' },
  { market: 'amazon.ie',      countries: 'ireland' },
  { market: 'amazon.de',      countries: 'germany, austria, switzerland, liechtenstein' },
  { market: 'amazon.fr',      countries: 'france, monaco' },
  { market: 'amazon.it',      countries: 'italy' },
  { market: 'amazon.es',      countries: 'spain, portugal' },
  { market: 'amazon.nl',      countries: 'netherlands, holland' },
  { market: 'amazon.com.be',  countries: 'belgium, luxembourg' },
  { market: 'amazon.se',      countries: 'sweden, norway, denmark, finland' },
  { market: 'amazon.pl',      countries: 'poland' },
  { market: 'amazon.com.tr',  countries: 'turkey' },
  { market: 'amazon.ae',      countries: 'uae, united arab emirates, dubai, abu dhabi, oman, qatar, bahrain, kuwait' },
  { market: 'amazon.sa',      countries: 'saudi arabia, ksa' },
  { market: 'amazon.eg',      countries: 'egypt' },
  { market: 'amazon.in',      countries: 'india' },
  { market: 'amazon.co.jp',   countries: 'japan' },
  { market: 'amazon.sg',      countries: 'singapore, malaysia' },
  { market: 'amazon.com.au',  countries: 'australia, new zealand' }
];

var PWB_ALL_KITS = ['VR Headset', 'Microbit', 'Makey Makey', 'Arduino'];

function _getKitLinksSheet() {
  var sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID_KIT_TRACKING')
                || '17Jsa2Kl2AkI5SgtlITYGqb-Q-PxfzkNFzzG5HwBJp_Q';
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('Kit Links');
  if (!sheet) {
    sheet = ss.insertSheet('Kit Links');
    sheet.appendRow(['Kit', 'Marketplace', 'Countries', 'Link']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.setFrozenRows(1);
    // Seed: every kit × every Amazon marketplace + one "Other" template row per kit
    var rows = [];
    PWB_ALL_KITS.forEach(function(kit) {
      PWB_AMAZON_MARKETS.forEach(function(m) {
        rows.push([kit, m.market, m.countries, '']);
      });
      rows.push([kit, 'Other', '<country names here>', '<local store link here>']);
    });
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    Logger.log('[PWB] Kit Links sheet created with ' + rows.length + ' seed rows.');
  }
  return sheet;
}

// Run once from the editor to create/seed the Kit Links sheet.
function setupKitLinksSheet() {
  _getKitLinksSheet();
  return 'Kit Links sheet ready — fill in the Link column.';
}

// Resolves the purchase link for a kit + learner country.
// Priority: exact marketplace country match → "Other" row country match → ''.
function getKitPurchaseLink(kitName, country) {
  try {
    if (!kitName || !country) return '';
    var c = String(country).toLowerCase().trim();
    var kitLow = String(kitName).toLowerCase().replace(/[-\s]/g, '');
    var rows = _getKitLinksSheet().getDataRange().getValues();
    var otherFallback = '';
    for (var i = 1; i < rows.length; i++) {
      var rKit  = String(rows[i][0] || '').toLowerCase().replace(/[-\s]/g, '');
      var rMkt  = String(rows[i][1] || '').trim();
      var rCtry = String(rows[i][2] || '').toLowerCase();
      var rLink = String(rows[i][3] || '').trim();
      if (rKit !== kitLow || !rLink || rLink.indexOf('<') === 0) continue;
      var matched = rCtry.split(',').some(function(cc) {
        cc = cc.trim();
        return cc && (cc === c || c.indexOf(cc) > -1 || cc.indexOf(c) > -1);
      });
      if (!matched) continue;
      if (rMkt.toLowerCase() === 'other') { otherFallback = rLink; continue; }
      return rLink; // Amazon marketplace match wins
    }
    return otherFallback;
  } catch(e) {
    Logger.log('[PWB] getKitPurchaseLink error: ' + e.message);
    return '';
  }
}

// ── Course → Kit lookup ────────────────────────────────────────────────────
function _getKitForCourse(courseName) {
  if (!courseName) return null;
  var trimmed = String(courseName).trim();
  // Exact match first
  if (PWB_COURSE_KIT_MAP[trimmed]) return PWB_COURSE_KIT_MAP[trimmed];
  // Case-insensitive partial match
  var lower = trimmed.toLowerCase();
  for (var key in PWB_COURSE_KIT_MAP) {
    if (key.toLowerCase().indexOf(lower) > -1 || lower.indexOf(key.toLowerCase()) > -1) {
      return PWB_COURSE_KIT_MAP[key];
    }
  }
  return null;
}

// ── Days until a date ──────────────────────────────────────────────────────
function _pwbDaysUntil(targetDate, today) {
  var t = new Date(targetDate); t.setHours(0,0,0,0);
  var n = new Date(today);      n.setHours(0,0,0,0);
  return Math.round((t - n) / 86400000);
}

// ── Build WATI params for each template ───────────────────────────────────
function _buildParentWillBuyParams(templateName, rowData) {
  switch (templateName) {
    case 'migration_parent_will_buy_kit':
      return [
        { name: 'Parent',    value: rowData.parentName  || 'Parent' },
        { name: 'Learner',   value: rowData.learnerName || 'Learner' },
        { name: 'Course',    value: rowData.courseName  || 'Course' },
        { name: 'Kit_name',  value: rowData.kitName     || 'Kit' },
        { name: 'kit_link',  value: rowData.amazonLink  || 'N/A' }
      ];
    case 'kits_parent_will_buy_fup_1':
      return [
        { name: 'Parent',      value: rowData.parentName  || 'Parent' },
        { name: 'kit_name',    value: rowData.kitName     || 'Kit' },
        { name: 'Learner',     value: rowData.learnerName || 'Learner' },
        { name: 'Course_name', value: rowData.courseName  || 'Course' },
        { name: 'Kit_link',    value: rowData.amazonLink  || 'N/A' }
      ];
    case 'migration_parent_will_buy_final_fup':
      return [
        { name: 'Parent',      value: rowData.parentName  || 'Parent' },
        { name: 'kit_name',    value: rowData.kitName     || 'Kit' },
        { name: 'Learner',     value: rowData.learnerName || 'Learner' },
        { name: 'Course_name', value: rowData.courseName  || 'Course' },
        { name: 'Kit_link',    value: rowData.amazonLink  || 'N/A' }
      ];
    default:
      return [];
  }
}

// ── Auto-fill HubSpot data into blank sheet cells ─────────────────────────
function _pwbAutoFillHSData(sheet, sheetRow, row, learnerName, parentName, phone, kitName) {
  if (!String(row[PWB_COL.LEARNER_NAME - 1] || '').trim() && learnerName)
    sheet.getRange(sheetRow, PWB_COL.LEARNER_NAME).setValue(learnerName);
  if (!String(row[PWB_COL.PARENT_NAME - 1] || '').trim() && parentName)
    sheet.getRange(sheetRow, PWB_COL.PARENT_NAME).setValue(parentName);
  if (!String(row[PWB_COL.PARENT_PHONE - 1] || '').trim() && phone)
    sheet.getRange(sheetRow, PWB_COL.PARENT_PHONE).setValue(phone);
  if (!String(row[PWB_COL.KIT - 1] || '').trim() && kitName)
    sheet.getRange(sheetRow, PWB_COL.KIT).setValue(kitName);
}

// ── Send one WATI message and stamp the timestamp column ──────────────────
function _sendPWBMessage(sheet, sheetRow, templateName, sentAtCol, phone, rowData) {
  Logger.log('[PWB] Sending ' + templateName + ' to ' + phone + ' (row ' + sheetRow + ')');
  var params = _buildParentWillBuyParams(templateName, rowData);
  var result = sendWatiMessage(phone, templateName, params);
  Logger.log('[PWB] WATI result: ' + JSON.stringify(result));
  sheet.getRange(sheetRow, sentAtCol).setValue(new Date());
}

// ── Shared escalation email sender (used by manual trigger only) ─────────
// Auto-escalation no longer sends email — manual button in UI does.
// CC always includes sourav.pal@jet-learn.com + TP manager if available.
function _sendPWBEscalationEmail(rowData, reason, isUrgent) {
  var ccList = ['sourav.pal@jet-learn.com'];
  if (rowData.tpManagerEmail && rowData.tpManagerEmail !== 'sourav.pal@jet-learn.com') {
    ccList.push(rowData.tpManagerEmail);
  }
  var ccStr = ccList.join(',');

  var triggeredBy = rowData.triggeredBy || 'Ops Team';

  try {
    var urgentBanner = isUrgent
      ? '<p style="color:#c0392b;font-weight:bold;font-size:15px;font-family:sans-serif;">⚠️ Urgent — course starts in ' + (rowData.daysLeft || '?') + ' day(s)</p>'
      : '';

    var htmlBody = urgentBanner +
      '<p style="font-family:sans-serif;font-size:14px;line-height:1.6;">Hi,</p>' +
      '<p style="font-family:sans-serif;font-size:14px;line-height:1.6;">' +
      'We\'ve been following up with ' + (rowData.parentName || 'the parent') + ' regarding the <strong>' + (rowData.kitName || 'kit') + '</strong> ' +
      'required for <strong>' + (rowData.learnerName || 'the learner') + '</strong>\'s upcoming course — ' +
      '<strong>' + (rowData.courseName || 'as listed below') + '</strong>' +
      (rowData.courseStartDate ? ', starting <strong>' + _formatDMY(rowData.courseStartDate) + '</strong>' : '') +
      '. Despite multiple follow-ups over WhatsApp, we haven\'t received any response or confirmation that the kit has been purchased.' +
      '</p>' +
      '<p style="font-family:sans-serif;font-size:14px;line-height:1.6;">Here are the details:</p>' +
      '<table style="border-collapse:collapse;font-family:sans-serif;font-size:14px;margin-bottom:16px;">' +
      '<tr><td style="padding:4px 16px 4px 0;color:#64748b;">JLID</td><td><strong>' + (rowData.jlid || '—') + '</strong></td></tr>' +
      '<tr><td style="padding:4px 16px 4px 0;color:#64748b;">Learner</td><td>' + (rowData.learnerName || '—') + '</td></tr>' +
      '<tr><td style="padding:4px 16px 4px 0;color:#64748b;">Parent</td><td>' + (rowData.parentName || '—') + '</td></tr>' +
      '<tr><td style="padding:4px 16px 4px 0;color:#64748b;">Kit Required</td><td>' + (rowData.kitName || '—') + '</td></tr>' +
      '<tr><td style="padding:4px 16px 4px 0;color:#64748b;">Course</td><td>' + (rowData.courseName || '—') + '</td></tr>' +
      (rowData.courseStartDate ? '<tr><td style="padding:4px 16px 4px 0;color:#64748b;">Course Start</td><td>' + _formatDMY(rowData.courseStartDate) + '</td></tr>' : '') +
      '</table>' +
      '<p style="font-family:sans-serif;font-size:14px;line-height:1.6;">' +
      'Since we\'ve exhausted our follow-ups with no success, could you please take the following steps:' +
      '</p>' +
      '<ol style="font-family:sans-serif;font-size:14px;line-height:1.8;">' +
      '<li>Please <strong>skip <em>' + (rowData.courseName || 'this course') + '</em></strong> from ' + (rowData.learnerName || 'the learner') + '\'s roadmap in HubSpot.</li>' +
      '<li>Let the teacher know to <strong>move on to the next course</strong> in the roadmap.</li>' +
      '</ol>' +
      '<p style="font-family:sans-serif;font-size:14px;line-height:1.6;">Thanks for handling this!</p>' +
      '<p style="font-family:sans-serif;font-size:13px;color:#94a3b8;">— ' + triggeredBy + ' via JetLearn Ops</p>';

    var subject = isUrgent
      ? '⚠️ URGENT [' + (rowData.daysLeft || '?') + 'd to course] Kit Not Confirmed — ' + rowData.learnerName + ' (' + rowData.jlid + ')'
      : '[Action Required] Kit Not Purchased — ' + rowData.learnerName + ' (' + rowData.jlid + ')';

    MailApp.sendEmail({
      to:       rowData.clsManagerEmail || CONFIG.EMAIL.MAIN_MANAGER,
      cc:       ccStr,
      subject:  subject,
      htmlBody: htmlBody,
      name:     CONFIG.EMAIL.FROM_NAME,
      from:     CONFIG.EMAIL.FROM
    });
    Logger.log('[PWB] Escalation email → ' + (rowData.clsManagerEmail || CONFIG.EMAIL.MAIN_MANAGER) + ' CC: ' + ccStr);
    return true;
  } catch(e) {
    Logger.log('[PWB] _sendPWBEscalationEmail ERROR: ' + e.message);
    return false;
  }
}

// ── HubSpot task + sheet update (NO auto email — manual button sends it) ──
function _escalateToCLS(sheet, sheetRow, rowData, reason) {
  Logger.log('[PWB] Escalating row ' + sheetRow + ' — ' + reason);

  // 1. HubSpot task
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY') || '';
    if (token && rowData.dealId) {
      var dueDate = new Date(); dueDate.setDate(dueDate.getDate() + 1);
      var subject = '[Kit Not Purchased] ' + rowData.learnerName + ' — ' + rowData.kitName;
      var body    = 'Parent has not purchased the ' + rowData.kitName + ' kit required for "' +
        rowData.courseName + '" (starts ' + _formatDMY(rowData.courseStartDate) + ').\n\n' +
        'Reason: ' + reason + '\n\nJLID: ' + rowData.jlid +
        '\nParent: ' + rowData.parentName + '\n\nAction required: contact parent and update roadmap.';
      var taskProps = {
        hs_task_subject:  subject,
        hs_task_body:     body,
        hs_task_status:   'NOT_STARTED',
        hs_task_priority: 'HIGH',
        hs_task_type:     'TODO',
        hs_timestamp:     dueDate.getTime()
      };
      var clsOwnerId = rowData.clsManagerEmail
        ? _resolveHubSpotOwnerIdByEmail(rowData.clsManagerEmail, token) : null;
      if (clsOwnerId) taskProps.hubspot_owner_id = clsOwnerId;
      var taskPayload = {
        properties: taskProps,
        associations: [{
          to: { id: String(rowData.dealId) },
          types: [{ associationCategory: 'HUBSPOT_DEFINED', associationTypeId: 216 }]
        }]
      };
      monitoredFetch('https://api.hubapi.com/crm/v3/objects/tasks', {
        method: 'post',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify(taskPayload),
        muteHttpExceptions: true
      });
      Logger.log('[PWB] HubSpot task created for row ' + sheetRow + ' owner=' + (clsOwnerId || 'unassigned'));
    }
  } catch(taskErr) {
    Logger.log('[PWB] HubSpot task creation failed: ' + taskErr.message);
  }

  // NOTE: Email is NOT sent automatically — use sendPWBManualEscalation from UI

  // 3. Update sheet
  sheet.getRange(sheetRow, PWB_COL.STATUS).setValue("Parent Didn't Buy - Roadmap Changed");
  sheet.getRange(sheetRow, PWB_COL.ESCALATED).setValue('TRUE');
  sheet.getRange(sheetRow, PWB_COL.ESCALATED_AT).setValue(new Date());

  // 4. Update HubSpot deal status — roadmap-changed (per-kit enum value) when the
  // purchase window is over; plain "Escalated to CLS" otherwise
  if (rowData.dealId) {
    var hsValue = PWB_HS_STATUSES.ESCALATED;
    if (rowData.useRoadmapStatus) {
      var kitProp = PWB_KIT_HS_PROP[rowData.kitName];
      hsValue = (kitProp && PWB_ROADMAP_HS_VALUE[kitProp]) || PWB_HS_STATUSES.ESCALATED;
    }
    _updateHubspotPWBStatus(rowData.dealId, hsValue, rowData.kitName);
  }
}

// Before escalating: if the deal's current course has changed away from the
// course this kit was for (and the new course doesn't need the same kit),
// close the row quietly — no CLS escalation needed.
// Returns true if the row was closed as course-changed.
function _pwbCheckCourseChanged(sheet, sheetRow, row) {
  try {
    var jlid       = String(row[PWB_COL.JLID - 1] || '').trim();
    var rowCourse  = String(row[PWB_COL.COURSE_NAME - 1] || '').trim();
    var rowKit     = String(row[PWB_COL.KIT - 1] || '').trim();
    if (!jlid || !rowCourse) return false;

    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success || !hs.data || !hs.data.course) return false;
    var dealCourse = String(hs.data.course).trim();
    if (!dealCourse || dealCourse === 'N/A') return false;

    // Same course → nothing changed
    if (dealCourse.toLowerCase() === rowCourse.toLowerCase()) return false;

    // Course changed — does the NEW course still need this kit?
    var newKit = _getKitForCourse(dealCourse);
    if (newKit && rowKit && newKit === rowKit) return false; // same kit still needed → proceed normally

    // Course changed and kit no longer needed → close row, no escalation
    Logger.log('[PWB] Row ' + sheetRow + ': course changed "' + rowCourse + '" → "' + dealCourse + '" — kit not needed, closing without escalation.');
    sheet.getRange(sheetRow, PWB_COL.STATUS).setValue('Course Changed - Kit Not Needed');
    sheet.getRange(sheetRow, PWB_COL.ESCALATED).setValue('TRUE');
    sheet.getRange(sheetRow, PWB_COL.ESCALATED_AT).setValue(new Date());
    if (hs.data.dealId) {
      try {
        _addNoteToDeal(hs.data.dealId,
          '[Parent Will Buy Kit] Follow-up closed automatically: course changed from "' + rowCourse +
          '" to "' + dealCourse + '" and the ' + (rowKit || 'kit') + ' is no longer required. No CLS escalation raised.');
      } catch(ne) { Logger.log('[PWB] course-change note error: ' + ne.message); }
    }
    return true;
  } catch(e) {
    Logger.log('[PWB] _pwbCheckCourseChanged error: ' + e.message);
    return false; // on error, fall through to normal escalation
  }
}

// ── Extract rowData from raw row array then escalate ──────────────────────
function _runPWBEscalation(sheet, sheetRow, row, reason, useRoadmapStatus) {
  // Guard: course changed on the deal → kit may no longer be needed
  if (_pwbCheckCourseChanged(sheet, sheetRow, row)) return;
  var jlid           = String(row[PWB_COL.JLID - 1]             || '').trim();
  var learnerName    = String(row[PWB_COL.LEARNER_NAME - 1]      || '').trim();
  var parentName     = String(row[PWB_COL.PARENT_NAME - 1]       || '').trim();
  var courseName     = String(row[PWB_COL.COURSE_NAME - 1]       || '').trim();
  var kitName        = String(row[PWB_COL.KIT - 1]               || '').trim() || _getKitForCourse(courseName) || courseName;
  var courseStartRaw = row[PWB_COL.COURSE_START_DATE - 1];
  var courseStart    = (courseStartRaw instanceof Date) ? courseStartRaw : _parseDMY(String(courseStartRaw || ''));
  var dealId         = '';

  var clsEmail = CONFIG.EMAIL.MAIN_MANAGER;
  if (jlid) {
    try {
      var hs = fetchHubspotByJlid(jlid);
      if (hs && hs.success && hs.data) {
        dealId      = hs.data.dealId      || '';
        learnerName = learnerName || hs.data.learnerName || '';
        parentName  = parentName  || hs.data.parentName  || '';
        if (hs.data.clsManagerName) {
          clsEmail = findClsEmailByManagerName(hs.data.clsManagerName) || clsEmail;
        }
      }
    } catch(e) { Logger.log('[PWB] _runPWBEscalation HS lookup failed: ' + e.message); }
  }

  _escalateToCLS(sheet, sheetRow, {
    jlid: jlid, learnerName: learnerName, parentName: parentName,
    courseName: courseName, kitName: kitName,
    courseStartDate: courseStart, dealId: dealId,
    clsManagerEmail: clsEmail, useRoadmapStatus: !!useRoadmapStatus
  }, reason);
}

// ── Adaptive FUP interval — timeline drives speed, never skips steps ─────────
// >21d = weekly (7d)  |  15–21d = 5d  |  8–14d = 3d  |  ≤7d = 2d  |  ≤3d = 1d
function _getAdaptiveInterval(daysUntilStart) {
  if (daysUntilStart > 21) return 7;
  if (daysUntilStart > 14) return 5;
  if (daysUntilStart > 7)  return 3;
  if (daysUntilStart > 3)  return 2;
  return 1;  // last resort — daily
}

// ── CLS heads-up: warn without fully escalating (fires on day 5 of urgency) ──
function _notifyCLSUrgent(sheet, sheetRow, rowData, daysLeft) {
  Logger.log('[PWB] CLS urgent heads-up: row ' + sheetRow + ', ' + daysLeft + 'd left');

  // Sheet status
  sheet.getRange(sheetRow, PWB_COL.STATUS).setValue('⚠️ CLS Notified - Awaiting Response');

  // HubSpot note
  if (rowData.dealId) {
    try {
      _addNoteToDeal(rowData.dealId,
        '[URGENT — Parent Will Buy Kit] Course starts in ' + daysLeft + ' day(s). ' +
        'Parent has not confirmed kit purchase despite all follow-ups. ' +
        'Learner: ' + rowData.learnerName + ' | Kit: ' + rowData.kitName +
        ' | Course: ' + rowData.courseName + '. Immediate action required.');
    } catch(ne) { Logger.log('[PWB] CLS note error: ' + ne.message); }
    _updateHubspotPWBStatus(rowData.dealId, PWB_HS_STATUSES.ESCALATED, rowData.kitName);
  }

  // NOTE: Email NOT sent automatically — dashboard shows "Escalate to CLS" button.
  // Row status set to ⚠️ CLS Notified so it's visible in the UI.
  Logger.log('[PWB] CLS urgent flagged (no auto email) for ' + rowData.jlid + ', ' + daysLeft + 'd left');
}

// ── Per-row scheduling logic ───────────────────────────────────────────────
// RULE: Never skip steps. Timeline shrinks the interval between them.
// ≤7 days = URGENT mode (2-day interval, CLS notified on day 5 = 2d before start)
// >7 days = normal mode (interval scales with days available)
function _processPWBRow(sheet, sheetRow, row, today) {
  var jlid         = String(row[PWB_COL.JLID - 1]             || '').trim();
  var courseName   = String(row[PWB_COL.COURSE_NAME - 1]       || '').trim();
  var startDateRaw = row[PWB_COL.COURSE_START_DATE - 1];
  var amazonLink   = String(row[PWB_COL.AMAZON_LINK - 1]       || '').trim();
  var status       = String(row[PWB_COL.STATUS - 1]            || '').trim();
  var escalated    = String(row[PWB_COL.ESCALATED - 1]         || '').trim();

  // Skip terminal or already-escalated rows
  if (PWB_TERMINAL_STATUSES.indexOf(status) > -1) return;
  if (escalated === 'TRUE') return;
  // CLS already notified (heads-up sent) — only full escalation remains (handled at course start)
  if (status === '⚠️ CLS Notified - Awaiting Response') return;

  // Skip incomplete rows
  if (!jlid || !courseName || !startDateRaw) return;

  // Parse course start date
  var courseStartDate = (startDateRaw instanceof Date) ? startDateRaw : _parseDMY(String(startDateRaw));
  if (!courseStartDate) {
    Logger.log('[PWB] Row ' + sheetRow + ': cannot parse start date "' + startDateRaw + '"');
    return;
  }

  var daysUntilStart = _pwbDaysUntil(courseStartDate, today);
  var interval       = _getAdaptiveInterval(daysUntilStart);
  var isUrgent       = daysUntilStart <= 7;

  Logger.log('[PWB] Row ' + sheetRow + ' JLID=' + jlid +
             ' daysUntil=' + daysUntilStart + ' interval=' + interval + ' urgent=' + isUrgent);

  // Parent promised a purchase date (AI-extracted) — wait for it (+1 day grace)
  var promisedRaw = String(row[PWB_COL.PROMISED_DATE - 1] || '').trim();
  var promisedDate = null;
  if (promisedRaw) {
    var pp = promisedRaw.split('-');
    if (pp.length === 3) { promisedDate = new Date(parseInt(pp[0]), parseInt(pp[1])-1, parseInt(pp[2])); promisedDate.setHours(0,0,0,0); }
  }

  // Promised date passed without purchase — auto roadmap change
  if (promisedDate) {
    var graceEnd = new Date(promisedDate); graceEnd.setDate(graceEnd.getDate() + 1);
    if (today > graceEnd) {
      Logger.log('[PWB] Row ' + sheetRow + ': promised date ' + promisedRaw + ' passed without purchase — roadmap change.');
      _runPWBEscalation(sheet, sheetRow, row, 'Parent promised to buy by ' + promisedRaw + ' but did not — roadmap change needed', true);
      return;
    }
    // Still within promise window — pause reminders and escalation
    Logger.log('[PWB] Row ' + sheetRow + ': waiting on promised date ' + promisedRaw + ' — paused.');
    return;
  }

  // Course already started — full escalation (roadmap change needed)
  if (daysUntilStart < 0) {
    Logger.log('[PWB] Row ' + sheetRow + ': course started, escalating.');
    _runPWBEscalation(sheet, sheetRow, row, 'Course started without kit purchase confirmation', true);
    return;
  }

  // Fetch HubSpot data
  var hs = fetchHubspotByJlid(jlid);
  if (!hs || !hs.success) {
    Logger.log('[PWB] Row ' + sheetRow + ': HubSpot lookup failed for ' + jlid);
    return;
  }

  var phone       = _normalisePhone(hs.data.parentContact || '');
  var parentName  = hs.data.parentName  || String(row[PWB_COL.PARENT_NAME - 1]  || '').trim() || 'Parent';
  var learnerName = hs.data.learnerName || String(row[PWB_COL.LEARNER_NAME - 1] || '').trim() || 'Learner';
  var dealId      = hs.data.dealId || '';
  var kitName     = _getKitForCourse(courseName);

  if (!kitName) {
    Logger.log('[PWB] Row ' + sheetRow + ': course "' + courseName + '" not in kit mapping — skip.');
    return;
  }
  if (!phone) {
    Logger.log('[PWB] Row ' + sheetRow + ': no phone for ' + jlid + ' — skip.');
    return;
  }

  _pwbAutoFillHSData(sheet, sheetRow, row, learnerName, parentName, phone, kitName);

  // Auto-fill purchase link from Kit Links sheet by learner country if not set
  if (!amazonLink) {
    var autoLink = getKitPurchaseLink(kitName, hs.data.country || '');
    if (autoLink) {
      amazonLink = autoLink;
      sheet.getRange(sheetRow, PWB_COL.AMAZON_LINK).setValue(autoLink);
      Logger.log('[PWB] Row ' + sheetRow + ': auto-filled link for ' + kitName + ' / ' + (hs.data.country || '?') + ' → ' + autoLink);
    }
  }

  // Resolve learner's CLS manager email from HubSpot deal
  var clsManagerName  = hs.data.clsManagerName || '';
  var clsManagerEmail = '';
  if (clsManagerName) {
    try { clsManagerEmail = findClsEmailByManagerName(clsManagerName) || ''; } catch(e) {}
  }
  if (!clsManagerEmail) clsManagerEmail = CONFIG.EMAIL.MAIN_MANAGER; // fallback

  var rowData = {
    jlid: jlid, parentName: parentName, learnerName: learnerName,
    courseName: courseName, kitName: kitName,
    amazonLink: amazonLink, courseStartDate: courseStartDate,
    dealId: dealId, clsManagerEmail: clsManagerEmail
  };

  var initialSentAt = row[PWB_COL.INITIAL_SENT_AT - 1];
  var fup1SentAt    = row[PWB_COL.FUP1_SENT_AT - 1];
  var fup2SentAt    = row[PWB_COL.FUP2_SENT_AT - 1];
  var finalSentAt   = row[PWB_COL.FINAL_FUP_SENT_AT - 1];

  var initialSent = !!initialSentAt;
  var fup1Sent    = !!fup1SentAt;
  var fup2Sent    = !!fup2SentAt;
  var finalSent   = !!finalSentAt;

  var statusLabel = isUrgent ? 'In Progress - URGENT 🔴' : 'In Progress';

  // ── 1. INITIAL ─────────────────────────────────────────────────────────────
  // Send immediately. Lock interval into sheet so it never drifts on future runs.
  if (!initialSent) {
    _sendPWBMessage(sheet, sheetRow, 'migration_parent_will_buy_kit',
                    PWB_COL.INITIAL_SENT_AT, phone, rowData);
    sheet.getRange(sheetRow, PWB_COL.INTERVAL).setValue(interval); // lock it
    sheet.getRange(sheetRow, PWB_COL.STATUS).setValue(statusLabel);
    _updateHubspotPWBStatus(dealId, PWB_HS_STATUSES.REMINDER_1, kitName);
    return;
  }

  // Use locked interval from sheet — immune to date drift
  var lockedInterval = Number(row[PWB_COL.INTERVAL - 1]) || interval;

  var initialDate = new Date(initialSentAt instanceof Date ? initialSentAt : initialSentAt);
  initialDate.setHours(0,0,0,0);

  // ── 2. FUP 1 ──────────────────────────────────────────────────────────────
  if (!fup1Sent) {
    var fup1Due = new Date(initialDate); fup1Due.setDate(fup1Due.getDate() + lockedInterval);
    if (today >= fup1Due) {
      _sendPWBMessage(sheet, sheetRow, 'kits_parent_will_buy_fup_1',
                      PWB_COL.FUP1_SENT_AT, phone, rowData);
      sheet.getRange(sheetRow, PWB_COL.STATUS).setValue(statusLabel);
      _updateHubspotPWBStatus(dealId, PWB_HS_STATUSES.REMINDER_2, kitName);
    }
    return;
  }

  var fup1Date = new Date(fup1SentAt instanceof Date ? fup1SentAt : fup1SentAt);
  fup1Date.setHours(0,0,0,0);

  // ── 3. FUP 2 ──────────────────────────────────────────────────────────────
  if (!fup2Sent) {
    var fup2Due = new Date(fup1Date); fup2Due.setDate(fup2Due.getDate() + lockedInterval);
    if (today >= fup2Due) {
      _sendPWBMessage(sheet, sheetRow, 'kits_parent_will_buy_fup_1',
                      PWB_COL.FUP2_SENT_AT, phone, rowData);
      sheet.getRange(sheetRow, PWB_COL.STATUS).setValue(statusLabel);
      _updateHubspotPWBStatus(dealId, PWB_HS_STATUSES.REMINDER_2, kitName);
    }
    return;
  }

  var fup2Date = new Date(fup2SentAt instanceof Date ? fup2SentAt : fup2SentAt);
  fup2Date.setHours(0,0,0,0);

  // ── 4. FINAL FUP ──────────────────────────────────────────────────────────
  if (!finalSent) {
    var finalDue = new Date(fup2Date); finalDue.setDate(finalDue.getDate() + lockedInterval);
    if (today >= finalDue) {
      _sendPWBMessage(sheet, sheetRow, 'migration_parent_will_buy_final_fup',
                      PWB_COL.FINAL_FUP_SENT_AT, phone, rowData);
      _updateHubspotPWBStatus(dealId, PWB_HS_STATUSES.FINAL, kitName);
      if (daysUntilStart <= 2) {
        _notifyCLSUrgent(sheet, sheetRow, rowData, daysUntilStart);
      } else {
        sheet.getRange(sheetRow, PWB_COL.STATUS).setValue(statusLabel);
      }
    }
    return;
  }

  // ── 5. POST-FINAL: all messages sent, still no reply, course approaching ───
  // Non-urgent cases: final was sent >2 days ago but course now within 2 days
  if (daysUntilStart <= 2) {
    _notifyCLSUrgent(sheet, sheetRow, rowData, daysUntilStart);
  }
}

// ── Daily trigger entry point ──────────────────────────────────────────────
// ── Manual escalation: triggered from UI button after user review ─────────
// Sends email to CLS + CC TP manager + sourav.pal@jet-learn.com
function sendPWBManualEscalation(jlid) {
  if (!jlid) return { success: false, message: 'No JLID' };
  try {
    var sheet   = _getPWBSheet();
    var lastRow = sheet.getLastRow();
    var data    = sheet.getRange(2, 1, lastRow - 1, PWB_COL.ENTRY_BY).getValues();
    var rowIdx  = -1;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][PWB_COL.JLID - 1] || '').trim() === jlid.trim()) { rowIdx = i; break; }
    }
    if (rowIdx === -1) return { success: false, message: 'JLID not found: ' + jlid };

    var row      = data[rowIdx];
    var sheetRow = rowIdx + 2;

    // Resolve deal info + manager emails from HubSpot
    var dealId = '', clsEmail = CONFIG.EMAIL.MAIN_MANAGER, tpManagerEmail = '';
    try {
      var hs = fetchHubspotByJlid(jlid);
      if (hs && hs.success && hs.data) {
        dealId       = hs.data.dealId || '';
        if (hs.data.clsManagerName) clsEmail = findClsEmailByManagerName(hs.data.clsManagerName) || clsEmail;
        // Try to get TP manager email from teacher data
        var teacherName = (hs.data.currentTeacher || '').replace(/^TJL\d+\s*-\s*/i, '').trim();
        if (teacherName) {
          try {
            var td = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA);
            for (var ti = 1; ti < td.length; ti++) {
              if (String(td[ti][1] || '').toLowerCase().trim() === teacherName.toLowerCase()) {
                tpManagerEmail = String(td[ti][10] || '').trim(); // col K = TP manager email
                break;
              }
            }
          } catch(te) { Logger.log('[PWB] TP manager lookup: ' + te.message); }
        }
      }
    } catch(he) { Logger.log('[PWB] escalation HS lookup: ' + he.message); }

    var courseStartRaw = row[PWB_COL.COURSE_START_DATE - 1];
    var courseStart    = (courseStartRaw instanceof Date) ? courseStartRaw : _parseDMY(String(courseStartRaw || ''));
    var daysLeft       = courseStart ? _pwbDaysUntil(courseStart, new Date()) : null;

    var activeUser = '';
    try { activeUser = Session.getActiveUser().getEmail() || ''; } catch(ue) {}

    var rowData = {
      jlid:           jlid,
      learnerName:    String(row[PWB_COL.LEARNER_NAME - 1] || '').trim() || 'Learner',
      parentName:     String(row[PWB_COL.PARENT_NAME  - 1] || '').trim() || 'Parent',
      courseName:     String(row[PWB_COL.COURSE_NAME  - 1] || '').trim(),
      kitName:        String(row[PWB_COL.KIT          - 1] || '').trim(),
      courseStartDate:courseStart,
      daysLeft:       daysLeft,
      dealId:         dealId,
      clsManagerEmail:clsEmail,
      tpManagerEmail: tpManagerEmail,
      triggeredBy:    activeUser || 'Ops Team'
    };

    var isUrgent = daysLeft !== null && daysLeft <= 7;
    var reason   = 'Manual escalation by ops team — all FUPs sent, no kit confirmation';

    // Send email
    var sent = _sendPWBEscalationEmail(rowData, reason, isUrgent);

    // Update sheet: mark escalated
    sheet.getRange(sheetRow, PWB_COL.STATUS).setValue("⚠️ CLS Notified - Awaiting Response");
    sheet.getRange(sheetRow, PWB_COL.ESCALATED).setValue('TRUE');
    sheet.getRange(sheetRow, PWB_COL.ESCALATED_AT).setValue(new Date());
    if (sent) {
      var ccParts = ['sourav.pal@jet-learn.com'];
      if (tpManagerEmail) ccParts.push(tpManagerEmail);
      sheet.getRange(sheetRow, PWB_COL.EMAIL_LOG).setValue('TO:' + clsEmail + ' | CC:' + ccParts.join(', '));
    }

    // HubSpot note + status
    if (dealId) {
      try { _addNoteToDeal(dealId, '[Kit Escalation] CLS notified manually. Reason: ' + reason); } catch(ne) {}
      _updateHubspotPWBStatus(dealId, PWB_HS_STATUSES.ESCALATED, rowData.kitName);
    }

    Logger.log('[PWB] Manual escalation done for ' + jlid + ' emailSent=' + sent);
    return { success: true, emailSent: sent, clsEmail: clsEmail };

  } catch(e) {
    Logger.log('[PWB] sendPWBManualEscalation ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Override: reset wrong Order Placed / Kit Received back to In Progress ─
// Use when parent response was incorrectly matched or status wrongly set.
function sendPWBStatusOverride(jlid) {
  if (!jlid) return { success: false, message: 'No JLID' };
  try {
    var sheet   = _getPWBSheet();
    var lastRow = sheet.getLastRow();
    var data    = sheet.getRange(2, 1, lastRow - 1, PWB_COL.ENTRY_BY).getValues();
    var rowIdx  = -1;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][PWB_COL.JLID - 1] || '').trim() === jlid.trim()) { rowIdx = i; break; }
    }
    if (rowIdx === -1) return { success: false, message: 'JLID not found: ' + jlid };

    var row      = data[rowIdx];
    var sheetRow = rowIdx + 2;
    var oldStatus = String(row[PWB_COL.STATUS - 1] || '').trim();

    // Reset status + clear response
    sheet.getRange(sheetRow, PWB_COL.STATUS).setValue('In Progress');
    sheet.getRange(sheetRow, PWB_COL.PARENT_RESPONSE).setValue('');
    // Clear escalation flags if they were set due to wrong status
    if (oldStatus === "Parent Didn't Buy - Roadmap Changed") {
      sheet.getRange(sheetRow, PWB_COL.ESCALATED).setValue('');
      sheet.getRange(sheetRow, PWB_COL.ESCALATED_AT).setValue('');
    }

    // Update HubSpot status back to Reminder 2 sent (most recent real FUP stage)
    var kitName = String(row[PWB_COL.KIT - 1] || '').trim();
    try {
      var hs = fetchHubspotByJlid(jlid);
      if (hs && hs.success && hs.data && hs.data.dealId) {
        _updateHubspotPWBStatus(hs.data.dealId, PWB_HS_STATUSES.REMINDER_2, kitName);
      }
    } catch(he) { Logger.log('[PWB] override HS update: ' + he.message); }

    Logger.log('[PWB] Status override: ' + jlid + ' was "' + oldStatus + '" → In Progress');
    return { success: true, oldStatus: oldStatus };

  } catch(e) {
    Logger.log('[PWB] sendPWBStatusOverride ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Manual FUP: send next pending FUP for a specific JLID ────────────────
// Called from the PWB dashboard "Send FUP" button.
// Figures out which step is next, sends the right template, stamps sheet + HS.
function sendPWBManualFup(jlid) {
  if (!jlid) return { success: false, message: 'No JLID provided' };
  try {
    var sheet   = _getPWBSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'PWB sheet empty' };

    var data = sheet.getRange(2, 1, lastRow - 1, PWB_COL.ENTRY_BY).getValues();
    var rowIdx = -1;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][PWB_COL.JLID - 1] || '').trim() === jlid.trim()) {
        rowIdx = i; break;
      }
    }
    if (rowIdx === -1) return { success: false, message: 'JLID not found in PWB sheet: ' + jlid };

    var row      = data[rowIdx];
    var sheetRow = rowIdx + 2;

    // Skip terminal rows
    var status   = String(row[PWB_COL.STATUS - 1] || '').trim();
    if (PWB_TERMINAL_STATUSES.indexOf(status) > -1) {
      return { success: false, message: 'Row is in terminal status: ' + status };
    }

    // Determine which FUP to send next
    var initialSent = !!row[PWB_COL.INITIAL_SENT_AT   - 1];
    var fup1Sent    = !!row[PWB_COL.FUP1_SENT_AT      - 1];
    var fup2Sent    = !!row[PWB_COL.FUP2_SENT_AT      - 1];
    var finalSent   = !!row[PWB_COL.FINAL_FUP_SENT_AT - 1];

    var templateName, sentAtCol, hsStatus, fupLabel;
    if (!initialSent) {
      templateName = 'migration_parent_will_buy_kit';
      sentAtCol    = PWB_COL.INITIAL_SENT_AT;
      hsStatus     = PWB_HS_STATUSES.REMINDER_1;
      fupLabel     = 'Initial';
    } else if (!fup1Sent) {
      templateName = 'kits_parent_will_buy_fup_1';
      sentAtCol    = PWB_COL.FUP1_SENT_AT;
      hsStatus     = PWB_HS_STATUSES.REMINDER_2;
      fupLabel     = 'FUP 1';
    } else if (!fup2Sent) {
      templateName = 'kits_parent_will_buy_fup_1';
      sentAtCol    = PWB_COL.FUP2_SENT_AT;
      hsStatus     = PWB_HS_STATUSES.REMINDER_2;
      fupLabel     = 'FUP 2';
    } else if (!finalSent) {
      templateName = 'migration_parent_will_buy_final_fup';
      sentAtCol    = PWB_COL.FINAL_FUP_SENT_AT;
      hsStatus     = PWB_HS_STATUSES.FINAL;
      fupLabel     = 'Final FUP';
    } else {
      return { success: false, message: 'All FUPs already sent for ' + jlid };
    }

    // Resolve phone
    var phone = String(row[PWB_COL.PARENT_PHONE - 1] || '').replace(/\D/g, '');
    if (!phone) {
      try {
        var hs = fetchHubspotByJlid(jlid);
        if (hs && hs.success && hs.data) phone = String(hs.data.parentContact || '').replace(/\D/g, '');
      } catch(pe) { Logger.log('[PWB] phone lookup error: ' + pe.message); }
    }
    if (!phone) return { success: false, message: 'No phone number found for ' + jlid };

    // Build rowData for template params
    var rowData = {
      parentName:  String(row[PWB_COL.PARENT_NAME  - 1] || '').trim() || 'Parent',
      learnerName: String(row[PWB_COL.LEARNER_NAME - 1] || '').trim() || 'Learner',
      courseName:  String(row[PWB_COL.COURSE_NAME  - 1] || '').trim() || 'Course',
      kitName:     String(row[PWB_COL.KIT          - 1] || '').trim() || 'Kit',
      amazonLink:  String(row[PWB_COL.AMAZON_LINK  - 1] || '').trim() || ''
    };

    // Send WATI message
    _sendPWBMessage(sheet, sheetRow, templateName, sentAtCol, phone, rowData);

    // Update sheet status
    sheet.getRange(sheetRow, PWB_COL.STATUS).setValue('In Progress');

    // Update HubSpot status
    var dealId = '';
    try {
      var hsData = fetchHubspotByJlid(jlid);
      if (hsData && hsData.success && hsData.data) dealId = hsData.data.dealId || '';
    } catch(he) { Logger.log('[PWB] manual FUP HS lookup error: ' + he.message); }
    if (dealId) _updateHubspotPWBStatus(dealId, hsStatus, rowData.kitName);

    Logger.log('[PWB] Manual FUP sent: ' + fupLabel + ' → ' + jlid + ' (' + phone + ')');
    return { success: true, fupLabel: fupLabel, templateName: templateName, phone: phone };

  } catch(e) {
    Logger.log('[PWB] sendPWBManualFup ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

function sendParentWillBuyFollowUps() {
  Logger.log('[PWB] sendParentWillBuyFollowUps started');
  try {
    var sheet   = _getPWBSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) { Logger.log('[PWB] No data rows.'); return; }

    var today = new Date(); today.setHours(0,0,0,0);
    var data  = sheet.getRange(2, 1, lastRow - 1, PWB_COL.ENTRY_BY).getValues();

    data.forEach(function(row, idx) {
      var sheetRow = idx + 2;
      try {
        _processPWBRow(sheet, sheetRow, row, today);
        Utilities.sleep(500); // gentle rate limiting between rows
      } catch(rowErr) {
        Logger.log('[PWB] Error on row ' + sheetRow + ': ' + rowErr.message);
      }
    });

    Logger.log('[PWB] sendParentWillBuyFollowUps done.');
  } catch(e) {
    Logger.log('[PWB] FATAL: ' + e.message + '\n' + e.stack);
  }
}

// ── Webhook handler — called from Code.js _processWatiPWBReply ────────────
function handleParentWillBuyReply(waId, buttonText, isFreeText) {
  Logger.log('[PWB] handleParentWillBuyReply waId=' + waId + ' btn="' + buttonText + '" free=' + (isFreeText || false));
  try {
    var normPhone = _normalisePhone(waId);
    if (!normPhone) return;

    var sheet   = _getPWBSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var rows = sheet.getRange(2, 1, lastRow - 1, PWB_COL.ENTRY_BY).getValues();

    // Match: phone + initial sent + not terminal
    // If siblings share phone → pick row with latest activity (most recent timestamp)
    var matchRow      = -1;
    var matchActivity = null;

    rows.forEach(function(row, idx) {
      var rowPhone   = _normalisePhone(String(row[PWB_COL.PARENT_PHONE - 1] || ''));
      var rowStatus  = String(row[PWB_COL.STATUS - 1] || '').trim();
      var initialSent = !!row[PWB_COL.INITIAL_SENT_AT - 1]; // match from first message onward
      if (rowPhone !== normPhone) return;
      if (!initialSent) return;
      if (PWB_TERMINAL_STATUSES.indexOf(rowStatus) > -1) return;
      if (rowStatus === '⚠️ CLS Notified - Awaiting Response') return;

      // Most recent timestamp across any FUP column = most active row
      var lastActivity = Math.max(
        row[PWB_COL.INITIAL_SENT_AT - 1]   ? new Date(row[PWB_COL.INITIAL_SENT_AT - 1]).getTime()   : 0,
        row[PWB_COL.FUP1_SENT_AT - 1]      ? new Date(row[PWB_COL.FUP1_SENT_AT - 1]).getTime()      : 0,
        row[PWB_COL.FUP2_SENT_AT - 1]      ? new Date(row[PWB_COL.FUP2_SENT_AT - 1]).getTime()      : 0,
        row[PWB_COL.FINAL_FUP_SENT_AT - 1] ? new Date(row[PWB_COL.FINAL_FUP_SENT_AT - 1]).getTime() : 0
      );

      if (matchRow === -1 || lastActivity > matchActivity) {
        matchRow      = idx + 2;
        matchActivity = lastActivity;
      }
    });

    if (matchRow === -1) {
      Logger.log('[PWB] No matching active row for phone ' + normPhone);
      return;
    }

    var dataRow      = rows[matchRow - 2];
    var replyJlid    = String(dataRow[PWB_COL.JLID - 1]  || '').trim();
    var replyKitName = String(dataRow[PWB_COL.KIT - 1]   || '').trim();
    var replyDealId  = '';

    // Resolve dealId for HubSpot updates
    if (replyJlid) {
      try {
        var hsR = fetchHubspotByJlid(replyJlid);
        if (hsR && hsR.success && hsR.data) replyDealId = hsR.data.dealId || '';
      } catch(he) { Logger.log('[PWB] handleReply HS lookup failed: ' + he.message); }
    }

    sheet.getRange(matchRow, PWB_COL.PARENT_RESPONSE).setValue(buttonText);
    // Only stamp the FIRST response — used for avg response-time KPI, shouldn't reset on later replies
    if (!dataRow[PWB_COL.RESPONSE_AT - 1]) {
      sheet.getRange(matchRow, PWB_COL.RESPONSE_AT).setValue(new Date());
    }

    // ── Free text reply — AI classify, then act ─────────────────────────────
    if (isFreeText) {
      Logger.log('[PWB] Free text reply from row ' + matchRow + ': "' + buttonText + '"');

      var ftStartRaw = dataRow[PWB_COL.COURSE_START_DATE - 1];
      var ftStart = (ftStartRaw instanceof Date) ? ftStartRaw : _parseDMY(String(ftStartRaw || ''));
      var ai = _pwbAnalyzeReply(buttonText, ftStart);
      Logger.log('[PWB] AI classification: ' + JSON.stringify(ai));
      sheet.getRange(matchRow, PWB_COL.AI_INTENT).setValue(ai.intent + (ai.promisedDate ? ' (' + ai.promisedDate + ')' : ''));

      if (ai.intent === 'BOUGHT') {
        sheet.getRange(matchRow, PWB_COL.STATUS).setValue('Order Placed');
        _updateHubspotPWBStatus(replyDealId, PWB_HS_STATUSES.BOUGHT, replyKitName);
        if (replyDealId) {
          try { _addNoteToDeal(replyDealId, '[Parent Will Buy Kit — AI] Parent indicated the kit is bought: "' + buttonText + '" (' + _formatDMY(new Date()) + ').'); } catch(ne) {}
        }
        return;
      }

      if (ai.intent === 'WILL_BUY') {
        if (ai.promisedDate) sheet.getRange(matchRow, PWB_COL.PROMISED_DATE).setValue(ai.promisedDate);
        _updateHubspotPWBStatus(replyDealId, 'Parent will buy', replyKitName);
        if (replyDealId) {
          try { _addNoteToDeal(replyDealId, '[Parent Will Buy Kit — AI] Parent will buy' + (ai.promisedDate ? ' by ' + ai.promisedDate : '') + ': "' + buttonText + '" (' + _formatDMY(new Date()) + '). Escalation paused until then.'); } catch(ne) {}
        }
        return; // no CLS email — schedule continues, escalation waits for promised date
      }

      // WONT_BUY and UNCLEAR both go to CLS; WONT_BUY flagged as such in the note
      if (replyDealId) {
        try {
          _addNoteToDeal(replyDealId,
            '[Parent Will Buy Kit — Parent Message' + (ai.intent === 'WONT_BUY' ? ' — DECLINED' : '') + '] Parent replied: "' + buttonText + '" on ' +
            _formatDMY(new Date()) + '. Manual follow-up may be needed.');
        } catch(ne) { Logger.log('[PWB] Note error: ' + ne.message); }
      }
      // Email CLS so nothing slips through
      try {
        var ftRow    = rows[matchRow - 2];
        var ftJlid   = String(ftRow[PWB_COL.JLID - 1] || '');
        var ftName   = String(ftRow[PWB_COL.LEARNER_NAME - 1] || '');
        var ftParent = String(ftRow[PWB_COL.PARENT_NAME - 1] || '');
        var ftCourse = String(ftRow[PWB_COL.COURSE_NAME - 1] || '');
        var ftCLS    = (replyDealId ? '' : ''); // already resolved above as replyDealId lookup
        // Resolve CLS email
        var ftClsEmail = CONFIG.EMAIL.MAIN_MANAGER;
        if (ftJlid) {
          try {
            var ftHs = fetchHubspotByJlid(ftJlid);
            if (ftHs && ftHs.success && ftHs.data && ftHs.data.clsManagerName) {
              ftClsEmail = findClsEmailByManagerName(ftHs.data.clsManagerName) || ftClsEmail;
            }
          } catch(e2) {}
        }
        MailApp.sendEmail({
          to:       ftClsEmail,
          subject:  '💬 Parent Replied (Free Text) — ' + ftName + ' (' + ftJlid + ')',
          htmlBody: '<p>A parent sent a free-text WhatsApp reply to a <strong>Parent Will Buy Kit</strong> follow-up.</p>' +
            '<table style="border-collapse:collapse;font-family:sans-serif;font-size:14px;">' +
            '<tr><td style="padding:4px 16px 4px 0"><strong>JLID</strong></td><td>' + ftJlid + '</td></tr>' +
            '<tr><td style="padding:4px 16px 4px 0"><strong>Learner</strong></td><td>' + ftName + '</td></tr>' +
            '<tr><td style="padding:4px 16px 4px 0"><strong>Parent</strong></td><td>' + ftParent + '</td></tr>' +
            '<tr><td style="padding:4px 16px 4px 0"><strong>Course</strong></td><td>' + ftCourse + '</td></tr>' +
            '<tr><td style="padding:4px 16px 4px 0"><strong>Message</strong></td><td style="color:#c0392b;font-weight:bold;">"' + buttonText + '"</td></tr>' +
            '</table>' +
            '<p>Please follow up with the parent directly.</p>' +
            '<p>— JetLearn Automation</p>',
          name: CONFIG.EMAIL.FROM_NAME, from: CONFIG.EMAIL.FROM
        });
        Logger.log('[PWB] Free text CLS email sent to ' + ftClsEmail);
      } catch(emailErr) { Logger.log('[PWB] Free text email error: ' + emailErr.message); }
      return; // don't process as button
    }

    if (buttonText === 'Order Placed') {
      sheet.getRange(matchRow, PWB_COL.STATUS).setValue('Order Placed');
      _updateHubspotPWBStatus(replyDealId, PWB_HS_STATUSES.BOUGHT, replyKitName);
      Logger.log('[PWB] Row ' + matchRow + ': Order Placed → HS updated');

    } else if (buttonText === 'Kit Received') {
      sheet.getRange(matchRow, PWB_COL.STATUS).setValue('Kit Received');
      _updateHubspotPWBStatus(replyDealId, PWB_HS_STATUSES.BOUGHT, replyKitName);
      Logger.log('[PWB] Row ' + matchRow + ': Kit Received → HS updated');

    } else if (buttonText === 'Yet to place an order') {
      // Response logged — schedule continues; add HubSpot note for visibility
      Logger.log('[PWB] Row ' + matchRow + ': "Yet to place an order" — logged, schedule continues');
      if (replyDealId) {
        try {
          _addNoteToDeal(replyDealId,
            '[Parent Will Buy Kit] Parent replied "Yet to place an order" on ' +
            _formatDMY(new Date()) + '. Follow up urgently — course starts soon.');
        } catch(ne) { Logger.log('[PWB] HubSpot note error: ' + ne.message); }
      }
    }

  } catch(e) {
    Logger.log('[PWB] handleParentWillBuyReply ERROR: ' + e.message + '\n' + e.stack);
  }
}

// ── Trigger setup ──────────────────────────────────────────────────────────
function setupParentWillBuyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendParentWillBuyFollowUps') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendParentWillBuyFollowUps')
    .timeBased().everyDays(1).atHour(9).create();
  Logger.log('[PWB] Daily trigger registered for sendParentWillBuyFollowUps at 9am.');
}
