/**
 * KitTrackingService.js
 * â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
 * Automates kit delivery follow-up via WATI WhatsApp.
 *
 * Flow:
 *   Daily trigger â†’ sendKitFollowUps()
 *     â†’ reads Kit Tracking sheet
 *     â†’ sends WATI "migration_kit_fup_sent_by_us" template for overdue rows
 *     â†’ marks Follow-up Sent + stores phone
 *
 *   WATI webhook â†’ handleKitReply(waId, buttonText)
 *     â†’ matches phone to sheet row
 *     â†’ "Kit Received"     â†’ fills Delivery Date, updates HubSpot kit status
 *     â†’ "Not Received yet" â†’ flags sheet, adds HubSpot deal note
 *     â†’ "Need To Check"    â†’ flags sheet for manual review
 *
 * Sheet: "Kit Tracking" (SHEET_ID_KIT_TRACKING)
 * Columns (1-indexed):
 *   A(1)  Sr No         B(2)  Learner Name   C(3)  Kit
 *   D(4)  Country       E(5)  Price EUR       F(6)  Site
 *   G(7)  Date of Order H(8)  Timestamp Month I(9)  ETA
 *   J(10) Delivery Date K(11) Time Taken      L(12) Reason
 *   M(13) Subscription  N(14) Roadmap         O(15) Col 15
 *   P(16) JLID          Q(17) Follow-up Sent  R(18) Follow-up Sent At
 *   S(19) Parent Response  T(20) Phone Sent To
 */

// â”€â”€ Column indices (1-based for getRange) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
var KIT_COL = {
  SR_NO:              1,
  LEARNER_NAME:       2,
  KIT:                3,
  DATE_OF_ORDER:      7,
  ETA:                9,
  DELIVERY_DATE:      10,
  TIME_TAKEN:         11,
  JLID:               16,
  FOLLOWUP_SENT:      17,
  FOLLOWUP_SENT_AT:   18,
  PARENT_RESPONSE:    19,
  PHONE_SENT_TO:      20,
  FOLLOWUP2_SENT:     21,   // T — "TRUE" when 2nd reminder sent
  FOLLOWUP2_SENT_AT:  22    // U — timestamp of 2nd reminder
};

// â”€â”€ HubSpot kit property map â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function _kitPropertyForType(kitName) {
  if (!kitName) return null;
  var k = kitName.toLowerCase().trim();
  if (k.indexOf('vr')       > -1 || k.indexOf('oculus') > -1 || k.indexOf('headset') > -1) return 'vr_headset__oculus_status';
  if (k.indexOf('microbit') > -1)                                                            return 'microbit_kit_status';
  if (k.indexOf('makey')    > -1)                                                            return 'makey_makey_kit_status';
  if (k.indexOf('arduino')  > -1)                                                            return 'arduino_kit_status';
  return null;
}

// â”€â”€ Get Kit Tracking sheet â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function _getKitSheet() {
  var ssId = PropertiesService.getScriptProperties().getProperty('SHEET_ID_KIT_TRACKING');
  if (!ssId) {
    // Fallback: hardcoded ID from plan
    ssId = '17Jsa2Kl2AkI5SgtlITYGqb-Q-PxfzkNFzzG5HwBJp_Q';
  }
  var ss = SpreadsheetApp.openById(ssId);
  // Try tab named "Kit Tracking", fall back to first sheet
  var sheet = ss.getSheetByName('Kits') || ss.getSheetByName('Kit Tracking') || ss.getSheets()[0];
  return sheet;
}

// â”€â”€ Parse DD/MM/YYYY â†’ Date â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function _parseDMY(str) {
  if (!str) return null;
  var s = String(str).trim();
  // Handle Date objects returned by Sheets
  if (s.indexOf('/') === -1 && s.indexOf('-') === -1) return null;
  var parts = s.indexOf('/') > -1 ? s.split('/') : s.split('-');
  if (parts.length !== 3) return null;
  // DD/MM/YYYY
  var d = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10) - 1;
  var y = parseInt(parts[2], 10);
  var dt = new Date(y, m, d);
  return isNaN(dt.getTime()) ? null : dt;
}

// â”€â”€ Format Date â†’ DD/MM/YYYY â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function _formatDMY(date) {
  var d = date.getDate();
  var m = date.getMonth() + 1;
  var y = date.getFullYear();
  return (d < 10 ? '0' + d : d) + '/' + (m < 10 ? '0' + m : m) + '/' + y;
}

// â”€â”€ Normalise phone â†’ digits only, no leading + â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function _normalisePhone(raw) {
  if (!raw) return '';
  return String(raw).replace(/[^0-9]/g, '');
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// ALL KIT PROPERTIES  (used for multi-kit HubSpot search)
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
var KIT_HS_PROPS = [
  'vr_headset__oculus_status',
  'microbit_kit_status',
  'makey_makey_kit_status',
  'arduino_kit_status'
];

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// AUTO-FIND JLID  â€” searches HubSpot deals where ANY kit property = "Sent by Us"
// then matches by learner name. kitName used only as tiebreak log.
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function _findJlidByKitStatus(learnerName, kitName) {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    if (!token) return null;

    // Each filterGroup = OR condition â€” match deals where ANY kit = "Sent by Us"
    var filterGroups = KIT_HS_PROPS.map(function(prop) {
      return { filters: [{ propertyName: prop, operator: 'EQ', value: 'Sent by Us' }] };
    });

    var body = {
      filterGroups: filterGroups,
      properties: ['jetlearner_id', 'dealname'].concat(KIT_HS_PROPS),
      limit: 200
    };

    var resp = UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/objects/deals/search', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });

    var respCode = resp.getResponseCode();
    var respText = resp.getContentText();
    Logger.log('[KitTracking] _findJlidByKitStatus: HTTP ' + respCode + ' body=' + respText.substring(0, 300));
    var data    = JSON.parse(respText);
    var results = (data && data.results) || [];

    Logger.log('[KitTracking] _findJlidByKitStatus: ' + results.length + ' deals found with any kit="Sent by Us"');

    if (!results.length) return null;

    var nameLower = learnerName.toLowerCase().trim();
    var match = null;

    // Pass 1 â€” full name match
    results.forEach(function(deal) {
      if (match) return;
      var dealName = String((deal.properties && deal.properties.dealname) || '').toLowerCase();
      var jlid     = String((deal.properties && deal.properties.jetlearner_id) || '').trim();
      if (jlid && dealName.indexOf(nameLower) > -1) match = jlid;
    });

    // Pass 2 â€” first name only (fallback)
    if (!match) {
      var firstName = nameLower.split(' ')[0];
      if (firstName.length > 2) {
        results.forEach(function(deal) {
          if (match) return;
          var dealName = String((deal.properties && deal.properties.dealname) || '').toLowerCase();
          var jlid     = String((deal.properties && deal.properties.jetlearner_id) || '').trim();
          if (jlid && dealName.indexOf(firstName) > -1) match = jlid;
        });
      }
    }

    if (match) {
      Logger.log('[KitTracking] _findJlidByKitStatus: matched JLID=' + match + ' for "' + learnerName + '" (kit: ' + kitName + ')');
    } else {
      Logger.log('[KitTracking] _findJlidByKitStatus: no name match for "' + learnerName + '" in ' + results.length + ' results');
    }
    return match;

  } catch (e) {
    Logger.log('[KitTracking] _findJlidByKitStatus ERROR: ' + e.message);
    return null;
  }
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// SEND KIT FOLLOW-UPS  (daily trigger at 8am)
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function sendKitFollowUps() {
  Logger.log('[KitTracking] sendKitFollowUps started');
  try {
    var sheet = _getKitSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('[KitTracking] No data rows found.');
      return;
    }

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var rows = sheet.getRange(2, 1, lastRow - 1, KIT_COL.PHONE_SENT_TO).getValues();
    var sent = 0;

    rows.forEach(function(row, idx) {
      var sheetRow = idx + 2;

      var jlid          = String(row[KIT_COL.JLID - 1]          || '').trim();
      var deliveryDate  = String(row[KIT_COL.DELIVERY_DATE - 1]  || '').trim();
      var followupSent  = String(row[KIT_COL.FOLLOWUP_SENT - 1]  || '').trim();
      var etaRaw        = row[KIT_COL.ETA - 1];
      var kitName       = String(row[KIT_COL.KIT - 1]            || '').trim();
      var learnerName   = String(row[KIT_COL.LEARNER_NAME - 1]   || '').trim();

      if (deliveryDate) return;
      if (followupSent === 'TRUE' || followupSent === 'true' || followupSent === true) return;
      if (!learnerName || !kitName) return;

      // Only process 2026+ rows — skip if no order date or order date < 2026
      var orderRawCheck  = row[KIT_COL.DATE_OF_ORDER - 1];
      var orderYearCheck = null;
      if (orderRawCheck instanceof Date) orderYearCheck = orderRawCheck.getFullYear();
      else { var od = _parseDMY(String(orderRawCheck || '')); if (od) orderYearCheck = od.getFullYear(); }
      if (!orderYearCheck || orderYearCheck < 2026) return;

      // Auto-fill JLID from HubSpot if missing
      if (!jlid) {
        Logger.log('[KitTracking] Row ' + sheetRow + ': no JLID â€” auto-searching HubSpot for "' + learnerName + '" / ' + kitName);
        jlid = _findJlidByKitStatus(learnerName, kitName) || '';
        if (jlid) {
          sheet.getRange(sheetRow, KIT_COL.JLID).setValue(jlid);
          Logger.log('[KitTracking] Row ' + sheetRow + ': auto-filled JLID=' + jlid);
        } else {
          Logger.log('[KitTracking] Row ' + sheetRow + ': could not auto-find JLID, skipping.');
          return;
        }
      }

      // Parse ETA â€” skip if future
      var etaDate = null;
      if (etaRaw instanceof Date) {
        etaDate = etaRaw;
      } else {
        etaDate = _parseDMY(String(etaRaw));
      }
      if (!etaDate) {
        Logger.log('[KitTracking] Row ' + sheetRow + ': cannot parse ETA "' + etaRaw + '", skipping.');
        return;
      }
      etaDate.setHours(0, 0, 0, 0);
      if (etaDate > today) {
        Logger.log('[KitTracking] Row ' + sheetRow + ': ETA ' + _formatDMY(etaDate) + ' is future, skipping.');
        return;
      }

      // Fetch parent phone + name from HubSpot
      var hs = fetchHubspotByJlid(jlid);
      if (!hs || !hs.success || !hs.data.parentContact) {
        Logger.log('[KitTracking] Row ' + sheetRow + ': HubSpot lookup failed for ' + jlid + ' â€” ' + (hs && hs.message));
        return;
      }

      var phone      = _normalisePhone(hs.data.parentContact);
      var parentName = hs.data.parentName || learnerName || 'Parent';

      if (!phone) {
        Logger.log('[KitTracking] Row ' + sheetRow + ': no phone for ' + jlid + ', skipping.');
        return;
      }

      // Send WATI template
      Logger.log('[KitTracking] Sending follow-up for row ' + sheetRow + ' JLID=' + jlid + ' phone=' + phone + ' kit=' + kitName);
      var watiResult = sendWatiMessage(phone, 'migration_kit_fup_sent_by_us', [
        { name: 'Parent',   value: parentName  },
        { name: 'Kit_name', value: kitName     },
        { name: 'Learner',  value: learnerName }
      ]);
      Logger.log('[KitTracking] WATI result: ' + JSON.stringify(watiResult));

      // Mark sheet
      sheet.getRange(sheetRow, KIT_COL.FOLLOWUP_SENT).setValue('TRUE');
      sheet.getRange(sheetRow, KIT_COL.FOLLOWUP_SENT_AT).setValue(new Date());
      sheet.getRange(sheetRow, KIT_COL.PHONE_SENT_TO).setValue(phone);
      sent++;
    });

    Logger.log('[KitTracking] sendKitFollowUps done. Sent ' + sent + ' first reminders.');

    // ── 2nd reminder pass ─────────────────────────────────────────────
    // Rows where: 1st follow-up sent 2+ days ago, no response, no 2nd follow-up sent
    var sent2 = 0;
    var twoDaysAgo = new Date(today);
    twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);

    // Re-read sheet (may have been updated above)
    var rows2 = sheet.getRange(2, 1, lastRow - 1, KIT_COL.FOLLOWUP2_SENT_AT).getValues();

    rows2.forEach(function(row, idx) {
      var sheetRow = idx + 2;

      var deliveryDate   = String(row[KIT_COL.DELIVERY_DATE - 1]   || '').trim();
      var followupSent   = String(row[KIT_COL.FOLLOWUP_SENT - 1]   || '').trim();
      var followup2Sent  = String(row[KIT_COL.FOLLOWUP2_SENT - 1]  || '').trim();
      var parentResponse = String(row[KIT_COL.PARENT_RESPONSE - 1] || '').trim();
      var sentAtRaw      = row[KIT_COL.FOLLOWUP_SENT_AT - 1];
      var learnerName    = String(row[KIT_COL.LEARNER_NAME - 1]    || '').trim();
      var kitName        = String(row[KIT_COL.KIT - 1]             || '').trim();
      var jlid           = String(row[KIT_COL.JLID - 1]            || '').trim();
      var phone          = String(row[KIT_COL.PHONE_SENT_TO - 1]   || '').trim();

      if (deliveryDate)   return; // already delivered
      if (parentResponse) return; // already replied
      if (followup2Sent === 'TRUE' || followup2Sent === 'true' || followup2Sent === true) return; // already 2nd sent
      if (followupSent  !== 'TRUE' && followupSent !== 'true'  && followupSent  !== true)  return; // 1st not sent yet

      // Parse 1st follow-up timestamp
      var sentAt = (sentAtRaw instanceof Date) ? sentAtRaw : (sentAtRaw ? new Date(sentAtRaw) : null);
      if (!sentAt || isNaN(sentAt.getTime())) return;
      sentAt.setHours(0, 0, 0, 0);
      if (sentAt > twoDaysAgo) return; // less than 2 days old

      // Need phone to send
      if (!phone) {
        Logger.log('[KitTracking] 2nd reminder row ' + sheetRow + ': no phone, skipping.');
        return;
      }

      // Need HubSpot for parent name
      var parentName = learnerName || 'Parent';
      if (jlid) {
        try {
          var hs2 = fetchHubspotByJlid(jlid);
          if (hs2 && hs2.success && hs2.data.parentName) parentName = hs2.data.parentName;
        } catch(he) {}
      }

      Logger.log('[KitTracking] Sending 2nd reminder row ' + sheetRow + ' JLID=' + jlid + ' phone=' + phone);
      sendWatiMessage(phone, 'migration_kit_fup_sent_by_us', [
        { name: 'Parent',   value: parentName  },
        { name: 'Kit_name', value: kitName     },
        { name: 'Learner',  value: learnerName }
      ]);

      sheet.getRange(sheetRow, KIT_COL.FOLLOWUP2_SENT).setValue('TRUE');
      sheet.getRange(sheetRow, KIT_COL.FOLLOWUP2_SENT_AT).setValue(new Date());
      sent2++;
    });

    Logger.log('[KitTracking] sendKitFollowUps done. 2nd reminders sent: ' + sent2);
  } catch (e) {
    Logger.log('[KitTracking] sendKitFollowUps ERROR: ' + e.message + '\n' + e.stack);
  }
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// HANDLE KIT REPLY  (called from doPost WATI webhook)
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function handleKitReply(waId, buttonText) {
  Logger.log('[KitTracking] handleKitReply waId=' + waId + ' btn="' + buttonText + '"');
  try {
    var normPhone = _normalisePhone(waId);
    if (!normPhone) {
      Logger.log('[KitTracking] No phone in waId, aborting.');
      return;
    }

    var sheet   = _getKitSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var rows = sheet.getRange(2, 1, lastRow - 1, KIT_COL.PHONE_SENT_TO).getValues();
    var matchRow = -1;

    // Find matching row by phone AND no delivery date yet
    rows.forEach(function(row, idx) {
      if (matchRow > -1) return;
      var phone        = _normalisePhone(String(row[KIT_COL.PHONE_SENT_TO - 1] || ''));
      var deliveryDate = String(row[KIT_COL.DELIVERY_DATE - 1] || '').trim();
      if (phone && phone === normPhone && !deliveryDate) {
        matchRow = idx + 2; // 1-based
      }
    });

    if (matchRow === -1) {
      Logger.log('[KitTracking] No matching pending row for phone ' + normPhone);
      return;
    }

    var dataRow      = rows[matchRow - 2];
    var jlid         = String(dataRow[KIT_COL.JLID - 1]         || '').trim();
    var kitName      = String(dataRow[KIT_COL.KIT - 1]          || '').trim();
    var learnerName  = String(dataRow[KIT_COL.LEARNER_NAME - 1] || '').trim();
    var orderDateRaw = dataRow[KIT_COL.DATE_OF_ORDER - 1];

    // Normalise JLID — strip trailing non-alphanumeric chars (e.g. JL39611449152C2 → JL39611449152C)
    var normJlid = jlid.replace(/[^A-Z0-9]$/i, '').trim();
    if (normJlid !== jlid) Logger.log('[KitTracking] JLID normalised: “' + jlid + '” → “' + normJlid + '”');

    Logger.log('[KitTracking] Matched row ' + matchRow + ' JLID=' + normJlid + ' kit=' + kitName);

    // Write parent response regardless
    sheet.getRange(matchRow, KIT_COL.PARENT_RESPONSE).setValue(buttonText);

    if (buttonText === 'Kit Received') {
      var today = new Date();

      // Fill Delivery Date + Time Taken
      sheet.getRange(matchRow, KIT_COL.DELIVERY_DATE).setValue(_formatDMY(today));
      var orderDate = (orderDateRaw instanceof Date) ? orderDateRaw : _parseDMY(String(orderDateRaw || ''));
      if (orderDate) {
        var diffDays = Math.round((today - orderDate) / (1000 * 60 * 60 * 24));
        sheet.getRange(matchRow, KIT_COL.TIME_TAKEN).setValue(diffDays + ' days');
      }

      // Update HubSpot kit status
      if (normJlid) _updateHubspotKitStatus(normJlid, kitName, 'Received by the Parents');

      // Auto-reply to parent — confirm we've noted receipt
      try {
        sendWatiSessionMessage(normPhone,
          '✅ Thank you for confirming! We\'ve updated our records — the ' + kitName + ' for ' +
          learnerName + ' is marked as delivered. 🎉 Get ready for an amazing learning experience!\n\n— Team JetLearn ✨'
        );
      } catch(re) { Logger.log('[KitTracking] Auto-reply error: ' + re.message); }

      Logger.log('[KitTracking] Kit confirmed received — row ' + matchRow + ' updated.');

    } else if (buttonText === 'Not Received yet') {
      // Add HubSpot note
      if (normJlid) {
        var hs1 = fetchHubspotByJlid(normJlid);
        if (hs1 && hs1.success && hs1.data.dealId) {
          _addNoteToDeal(hs1.data.dealId,
            '[Kit Follow-up] Parent replied “Not Received yet” for kit: ' + kitName +
            ' on ' + _formatDMY(new Date()) + '. Verify with logistics and update parent.');
        }
      }
      // Auto-reply to parent — we're checking with logistics
      try {
        sendWatiSessionMessage(normPhone,
          '😟 We\'re sorry to hear that! We\'re checking with our logistics team right away.\n\n' +
          'We\'ll update you as soon as we have more information. Thank you for letting us know!\n\n— Team JetLearn ✨'
        );
      } catch(re) { Logger.log('[KitTracking] Auto-reply error: ' + re.message); }

      Logger.log('[KitTracking] Not received — row ' + matchRow + ' flagged, HS note added.');

    } else if (buttonText === 'Need To check') {
      // Add HubSpot note
      if (normJlid) {
        var hs2 = fetchHubspotByJlid(normJlid);
        if (hs2 && hs2.success && hs2.data.dealId) {
          _addNoteToDeal(hs2.data.dealId,
            '[Kit Follow-up] Parent replied “Need To Check” for kit: ' + kitName +
            ' on ' + _formatDMY(new Date()) + '. Awaiting parent confirmation in 12-24 hrs.');
        }
      }
      // Auto-reply to parent — we'll follow up in 12-24 hrs
      try {
        sendWatiSessionMessage(normPhone,
          '👍 No problem! Please check when you get a chance.\n\n' +
          'We\'ll follow up with you in 12-24 hours to confirm. Thank you!\n\n— Team JetLearn ✨'
        );
      } catch(re) { Logger.log('[KitTracking] Auto-reply error: ' + re.message); }

      Logger.log('[KitTracking] Need to check — row ' + matchRow + ' flagged, HS note added.');
    }

  } catch (e) {
    Logger.log('[KitTracking] handleKitReply ERROR: ' + e.message + '\n' + e.stack);
  }
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// UPDATE HUBSPOT KIT STATUS  (PATCH deal property)
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function _updateHubspotKitStatus(jlid, kitName, statusValue) {
  try {
    var prop = _kitPropertyForType(kitName);
    if (!prop) {
      Logger.log('[KitTracking] Unknown kit type "' + kitName + '" â€” no HubSpot property to update.');
      return;
    }

    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success || !hs.data.dealId) {
      Logger.log('[KitTracking] Cannot update HubSpot â€” no dealId for ' + jlid);
      return;
    }

    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    if (!token) {
      Logger.log('[KitTracking] HUBSPOT_API_KEY not set.');
      return;
    }

    var url = 'https://api.hubapi.com/crm/v3/objects/deals/' + hs.data.dealId;
    var resp = UrlFetchApp.fetch(url, {
      method: 'PATCH',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ properties: { [prop]: statusValue } }),
      muteHttpExceptions: true
    });

    var code = resp.getResponseCode();
    Logger.log('[KitTracking] HubSpot PATCH ' + prop + '="' + statusValue + '" â†’ HTTP ' + code);
    if (code !== 200) {
      Logger.log('[KitTracking] HubSpot error body: ' + resp.getContentText().substring(0, 300));
    }
  } catch (e) {
    Logger.log('[KitTracking] _updateHubspotKitStatus ERROR: ' + e.message);
  }
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// ADD NOTE TO HUBSPOT DEAL
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function _addNoteToDeal(dealId, noteBody) {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    if (!token || !dealId || !noteBody) return;

    // Create note via CRM v3 engagements (Notes)
    var payload = {
      properties: {
        hs_note_body:      noteBody,
        hs_timestamp:      String(new Date().getTime())
      }
    };
    var resp = UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/objects/notes', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var noteId = null;
    try { noteId = JSON.parse(resp.getContentText()).id; } catch(e2) {}

    // Associate note to deal
    if (noteId) {
      UrlFetchApp.fetch(
        'https://api.hubapi.com/crm/v3/objects/notes/' + noteId + '/associations/deals/' + dealId + '/note_to_deal',
        {
          method: 'put',
          headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
          payload: '{}',
          muteHttpExceptions: true
        }
      );
    }
    Logger.log('[KitTracking] Note added to deal ' + dealId + ' (noteId=' + noteId + ')');
  } catch (e) {
    Logger.log('[KitTracking] _addNoteToDeal ERROR: ' + e.message);
  }
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// GET KIT TRACKING DATA  (called from client dashboard)
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function getKitTrackingData() {
  try {
    var sheet   = _getKitSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, rows: [], stats: { total: 0, delivered: 0, awaiting: 0, notReceived: 0, overdue: 0 } };

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var raw  = sheet.getRange(2, 1, lastRow - 1, KIT_COL.FOLLOWUP2_SENT_AT).getValues();
    var rows = [];

    var cutoff = new Date(2026, 0, 1); // Jan 1 2026 â€” ignore older rows

    raw.forEach(function(r, idx) {
      var srNo         = r[KIT_COL.SR_NO - 1];
      var learnerName  = String(r[KIT_COL.LEARNER_NAME - 1]   || '').trim();
      var kit          = String(r[KIT_COL.KIT - 1]            || '').trim();
      var orderRaw     = r[KIT_COL.DATE_OF_ORDER - 1];
      var tsMonth      = String(r[7]                           || '').trim(); // col H = Timestamp month
      var etaRaw       = r[KIT_COL.ETA - 1];
      // Delivery date â€” handle both Date objects and DD/MM/YYYY strings
      var deliveryRaw  = r[KIT_COL.DELIVERY_DATE - 1];
      var deliveryDate = '';
      if (deliveryRaw instanceof Date && !isNaN(deliveryRaw.getTime())) {
        deliveryDate = _formatDMY(deliveryRaw);
      } else if (deliveryRaw) {
        var dStr = String(deliveryRaw).trim();
        // Strip full timestamp if present (e.g. "Fri Apr 03 2026 00:00:00 GMT+0530...")
        if (dStr.length > 10 && dStr.indexOf('/') === -1) {
          var parsed = new Date(dStr);
          deliveryDate = isNaN(parsed.getTime()) ? dStr : _formatDMY(parsed);
        } else {
          deliveryDate = dStr;
        }
      }
      var timeTaken    = String(r[KIT_COL.TIME_TAKEN - 1]     || '').trim();
      var followupSent  = String(r[KIT_COL.FOLLOWUP_SENT - 1]   || '').trim();
      var sentAt        = r[KIT_COL.FOLLOWUP_SENT_AT - 1];
      var response      = String(r[KIT_COL.PARENT_RESPONSE - 1] || '').trim();
      var jlid          = String(r[KIT_COL.JLID - 1]            || '').trim();
      var followup2Sent = String(r[KIT_COL.FOLLOWUP2_SENT - 1]  || '').trim();
      var sentAt2       = r[KIT_COL.FOLLOWUP2_SENT_AT - 1];

      if (!learnerName && !kit) return; // blank row

      // Parse order date â€” skip rows before Jan 2026
      var orderDate = null;
      if (orderRaw instanceof Date) {
        orderDate = orderRaw;
      } else if (orderRaw) {
        orderDate = _parseDMY(String(orderRaw));
      }
      if (!orderDate || orderDate < cutoff) return;

      // Build orderMonth label e.g. "January 2026"
      var orderMonth = '';
      if (tsMonth) {
        // Sheet col H already has "March-26", "April-26" etc â€” normalise to "March 2026"
        orderMonth = tsMonth.replace('-', ' 20').replace('-26', ' 2026').replace('-25', ' 2025');
        // Handle "March-26" â†’ "March 2026"
        if (orderMonth.match(/\w+ \d{2}$/)) {
          orderMonth = orderMonth.replace(/(\w+ )(\d{2})$/, '$120$2');
        }
      } else if (orderDate) {
        var MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
        orderMonth = MONTHS[orderDate.getMonth()] + ' ' + orderDate.getFullYear();
      }

      // Parse ETA
      var etaStr = '';
      var etaDate = null;
      if (etaRaw instanceof Date) {
        etaDate = etaRaw;
        etaStr  = _formatDMY(etaRaw);
      } else if (etaRaw) {
        etaDate = _parseDMY(String(etaRaw));
        etaStr  = String(etaRaw).trim();
      }
      if (etaDate) etaDate.setHours(0, 0, 0, 0);

      var fupSentBool  = (followupSent  === 'TRUE' || followupSent  === 'true');
      var fup2SentBool = (followup2Sent === 'TRUE' || followup2Sent === 'true');
      var sentAtStr    = '';
      if (sentAt instanceof Date) sentAtStr = _formatDMY(sentAt);
      else if (sentAt) sentAtStr = String(sentAt).trim();
      var sentAt2Str = '';
      if (sentAt2 instanceof Date) sentAt2Str = _formatDMY(sentAt2);
      else if (sentAt2) sentAt2Str = String(sentAt2).trim();

      // Days since 2nd follow-up
      var daysSince2nd = null;
      if (fup2SentBool && sentAt2) {
        var t2 = (sentAt2 instanceof Date) ? sentAt2 : new Date(sentAt2);
        t2.setHours(0,0,0,0);
        daysSince2nd = Math.floor((today - t2) / 86400000);
      }

      // Compute status
      var status = 'pending';
      if (deliveryDate || response === 'Kit Received') {
        status = 'delivered';
      } else if (response === 'Not Received yet') {
        status = 'not_received';
      } else if (response === 'Need To check') {
        status = 'need_check';
      } else if (fup2SentBool && !response) {
        // 2nd reminder sent, still no reply → escalated
        status = 'escalated';
      } else if (fupSentBool && !response) {
        status = 'awaiting';
      } else if (!fupSentBool && etaDate && etaDate <= today) {
        status = 'overdue';
      }

      rows.push({
        rowIndex:      idx + 2,
        srNo:          srNo || (idx + 1),
        learnerName:   learnerName,
        kit:           kit,
        orderMonth:    orderMonth,
        eta:           etaStr,
        deliveryDate:  deliveryDate,
        timeTaken:     timeTaken,
        followupSent:  fupSentBool,
        sentAt:        sentAtStr,
        followup2Sent: fup2SentBool,
        sentAt2:       sentAt2Str,
        daysSince2nd:  daysSince2nd,
        response:      response,
        jlid:          jlid,
        status:        status,
        // Extra detail fields
        country:      String(r[3]  || '').trim(),   // D: Country
        price:        String(r[4]  || '').trim(),   // E: Price EUR
        site:         String(r[5]  || '').trim(),   // F: Site
        orderDate:    (function() {
          var v = r[6];
          if (v instanceof Date) return _formatDMY(v);
          return String(v || '').trim();
        })(),                                        // G: Date of Order
        reason:       String(r[11] || '').trim(),   // L: Reason
        subscription: String(r[12] || '').trim(),   // M: Subscription
        roadmap:      String(r[13] || '').trim(),   // N: Roadmap
        sentBy:       String(r[14] || '').trim()    // O: Sent By
      });
    });

    // Stats
    var stats = {
      total:       rows.length,
      delivered:   rows.filter(function(r) { return r.status === 'delivered'; }).length,
      awaiting:    rows.filter(function(r) { return r.status === 'awaiting'; }).length,
      notReceived: rows.filter(function(r) { return r.status === 'not_received' || r.status === 'need_check'; }).length,
      overdue:     rows.filter(function(r) { return r.status === 'overdue'; }).length,
      escalated:   rows.filter(function(r) { return r.status === 'escalated'; }).length
    };

    Logger.log('[KitTracking] getKitTrackingData: ' + rows.length + ' rows, stats=' + JSON.stringify(stats));
    return { success: true, rows: rows, stats: stats };

  } catch (e) {
    Logger.log('[KitTracking] getKitTrackingData ERROR: ' + e.message);
    return { success: false, message: e.message, rows: [], stats: {} };
  }
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// RESEND KIT FOLLOW-UP  (manual resend from dashboard)
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// FETCH LEARNER DETAILS FOR KIT FORM  (JLID lookup in Add Kit modal)
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function fetchKitLearnerDetails(jlid) {
  try {
    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success) return { success: false, message: (hs && hs.message) || 'Learner not found for JLID: ' + jlid };
    var d = hs.data;
    return {
      success:      true,
      learnerName:  d.learnerName   || '',
      subscription: d.planName      || '',   // subscription property
      country:      d.country       || '',
      kitCostSoFar: d.learningKitCost || 0,  // existing learning_kit_cost
      dealId:       d.dealId        || ''
    };
  } catch (e) {
    Logger.log('[KitTracking] fetchKitLearnerDetails ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// ADD KIT ENTRY  (called from client Add Kit modal)
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function addKitEntry(data) {
  try {
    var sheet   = _getKitSheet();
    var sheetLastRow = sheet.getLastRow();

    // Find actual last data row — scan col A from bottom, skip empty/formatted rows
    var lastDataRow = 1;
    if (sheetLastRow >= 2) {
      var colA = sheet.getRange(2, 1, sheetLastRow - 1, 1).getValues();
      for (var i = colA.length - 1; i >= 0; i--) {
        if (String(colA[i][0]).trim() !== '') { lastDataRow = i + 2; break; }
      }
    }

    // Auto Sr No = highest existing Sr No + 1
    var lastSrNo = 0;
    if (sheetLastRow >= 2) {
      var srValues = sheet.getRange(2, KIT_COL.SR_NO, sheetLastRow - 1, 1).getValues();
      srValues.forEach(function(r) {
        var v = parseInt(r[0], 10);
        if (!isNaN(v) && v > lastSrNo) lastSrNo = v;
      });
    }
    var srNo = lastSrNo + 1;

    // Auto Timestamp Month from order date  e.g. "April-26"
    var MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    var SHORT_MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var tsMonth = '';
    var orderDate = null;
    if (data.orderDate) {
      // data.orderDate arrives as "YYYY-MM-DD" from HTML date input
      var parts = data.orderDate.split('-');
      if (parts.length === 3) {
        orderDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
        tsMonth = MONTHS[orderDate.getMonth()] + '-' + String(orderDate.getFullYear()).slice(2);
        // e.g. "April-26"
      }
    }

    // Format dates DD/MM/YYYY
    var fmtDate = function(iso) {
      if (!iso) return '';
      var p = iso.split('-');
      if (p.length !== 3) return iso;
      return p[2] + '/' + p[1] + '/' + p[0];
    };

    var jlid = String(data.jlid || '').trim();

    // Write per-column to skip col H (formulated in sheet — never overwrite)
    var writeRow = lastDataRow + 1;
    var writeMap = [
      [1,  srNo],                        // A: Sr No
      [2,  data.learnerName || ''],      // B: Learner's name
      [3,  data.kit         || ''],      // C: Kit
      [4,  data.country     || ''],      // D: Country
      [5,  data.price       || ''],      // E: Price (EUR)
      [6,  data.site        || ''],      // F: Site
      [7,  fmtDate(data.orderDate)],     // G: Date of Order
      // col 8 (H) = SKIP — formula in sheet
      [9,  fmtDate(data.eta)],           // I: ETA
      [10, fmtDate(data.deliveryDate)],  // J: Delivery Date
      [12, data.reason       || ''],     // L: Reason
      [13, data.subscription || ''],     // M: Current Subscription
      [14, data.roadmap      || ''],     // N: Roadmap
      [15, data.sentBy       || ''],     // O: Name
      [16, jlid]                         // P: JLID
    ];
    writeMap.forEach(function(pair) {
      sheet.getRange(writeRow, pair[0]).setValue(pair[1]);
    });
    Logger.log('[KitTracking] addKitEntry: wrote to row ' + writeRow + ' sr=' + srNo);
    Logger.log('[KitTracking] addKitEntry: row added sr=' + srNo + ' learner=' + data.learnerName + ' jlid=' + jlid + ' row=' + writeRow);

    // â”€â”€ Accumulate learning_kit_cost in HubSpot â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    var price = parseFloat(data.price) || 0;
    if (jlid && price > 0) {
      try {
        var hs = fetchHubspotByJlid(jlid);
        if (hs && hs.success && hs.data && hs.data.dealId) {
          var existing  = parseFloat(hs.data.learningKitCost) || 0;
          var newTotal  = existing + price;
          var token     = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
          UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/objects/deals/' + hs.data.dealId, {
            method: 'PATCH',
            headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
            payload: JSON.stringify({ properties: { learning_kit_cost: String(newTotal) } }),
            muteHttpExceptions: true
          });
          Logger.log('[KitTracking] learning_kit_cost updated: ' + existing + ' + ' + price + ' = ' + newTotal + ' for dealId=' + hs.data.dealId);
        }
      } catch (hsErr) {
        Logger.log('[KitTracking] learning_kit_cost update ERROR (non-fatal): ' + hsErr.message);
      }
    }

    return { success: true, srNo: srNo };

  } catch (e) {
    Logger.log('[KitTracking] addKitEntry ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// SEND FOLLOW-UP BY ROW INDEX  (manual send â€” handles missing JLID via auto-lookup)
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function sendKitFollowUpByRow(rowIndex, jlidOverride) {
  try {
    var sheet = _getKitSheet();
    var row   = sheet.getRange(rowIndex, 1, 1, KIT_COL.PHONE_SENT_TO).getValues()[0];

    var learnerName = String(row[KIT_COL.LEARNER_NAME - 1] || '').trim();
    var kitName     = String(row[KIT_COL.KIT - 1]          || '').trim();
    var jlid        = jlidOverride || String(row[KIT_COL.JLID - 1] || '').trim();

    // Auto-find JLID if missing
    if (!jlid) {
      jlid = _findJlidByKitStatus(learnerName, kitName) || '';
      if (!jlid) return { success: false, needJlid: true, message: 'Could not auto-find JLID for "' + learnerName + '". Please enter JLID manually.' };
      // Save it to sheet
      sheet.getRange(rowIndex, KIT_COL.JLID).setValue(jlid);
      Logger.log('[KitTracking] sendKitFollowUpByRow: auto-filled JLID=' + jlid + ' for row ' + rowIndex);
    }

    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success || !hs.data.parentContact) return { success: false, message: 'HubSpot lookup failed for ' + jlid + ': ' + (hs && hs.message) };

    var phone      = _normalisePhone(hs.data.parentContact);
    var parentName = hs.data.parentName || hs.data && hs.data.parentName || learnerName;
    if (!phone) return { success: false, message: 'No phone number found for ' + jlid };

    var watiResult = sendWatiMessage(phone, 'migration_kit_fup_sent_by_us', [
      { name: 'Parent',   value: parentName  },
      { name: 'Kit_name', value: kitName     },
      { name: 'Learner',  value: learnerName }
    ]);
    Logger.log('[KitTracking] sendKitFollowUpByRow: WATI=' + JSON.stringify(watiResult));

    sheet.getRange(rowIndex, KIT_COL.FOLLOWUP_SENT).setValue('TRUE');
    sheet.getRange(rowIndex, KIT_COL.FOLLOWUP_SENT_AT).setValue(new Date());
    sheet.getRange(rowIndex, KIT_COL.PHONE_SENT_TO).setValue(phone);

    return { success: true, message: 'Follow-up sent to ' + phone, jlid: jlid };
  } catch (e) {
    Logger.log('[KitTracking] sendKitFollowUpByRow ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// RESEND KIT FOLLOW-UP  (manual resend from dashboard)
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function resendKitFollowUp(jlid) {
  try {
    if (!jlid) return { success: false, message: 'No JLID provided.' };
    var sheet   = _getKitSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Sheet empty.' };

    var rows = sheet.getRange(2, 1, lastRow - 1, KIT_COL.PHONE_SENT_TO).getValues();
    var matchRow = -1;
    rows.forEach(function(r, idx) {
      if (matchRow > -1) return;
      if (String(r[KIT_COL.JLID - 1] || '').trim() === jlid) matchRow = idx + 2;
    });

    if (matchRow === -1) return { success: false, message: 'JLID not found in sheet: ' + jlid };

    var dataRow     = rows[matchRow - 2];
    var learnerName = String(dataRow[KIT_COL.LEARNER_NAME - 1] || '').trim();
    var kitName     = String(dataRow[KIT_COL.KIT - 1]          || '').trim();

    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success || !hs.data.parentContact) {
      return { success: false, message: 'HubSpot lookup failed for ' + jlid };
    }

    var phone      = _normalisePhone(hs.data.parentContact);
    var parentName = hs.data.parentName || learnerName || 'Parent';
    if (!phone) return { success: false, message: 'No phone found for ' + jlid };

    var watiResult = sendWatiMessage(phone, 'migration_kit_fup_sent_by_us', [
      { name: 'Parent',   value: parentName  },
      { name: 'Kit_name', value: kitName     },
      { name: 'Learner',  value: learnerName }
    ]);
    Logger.log('[KitTracking] Resend result: ' + JSON.stringify(watiResult));

    // Update sheet
    sheet.getRange(matchRow, KIT_COL.FOLLOWUP_SENT).setValue('TRUE');
    sheet.getRange(matchRow, KIT_COL.FOLLOWUP_SENT_AT).setValue(new Date());
    sheet.getRange(matchRow, KIT_COL.PHONE_SENT_TO).setValue(phone);

    return { success: true, message: 'Follow-up resent to ' + phone };
  } catch (e) {
    Logger.log('[KitTracking] resendKitFollowUp ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// UPDATE KIT ROW  (manual edit from dashboard — pencil button)
// ─────────────────────────────────────────────────────────────────────────────
function updateKitRow(data) {
  try {
    var rowIndex     = parseInt(data.rowIndex, 10);
    var jlid         = String(data.jlid         || '').trim();
    var deliveryDate = String(data.deliveryDate || '').trim();
    var response     = String(data.response     || '').trim();

    if (!rowIndex || rowIndex < 2) return { success: false, message: 'Invalid row index.' };

    var sheet = _getKitSheet();
    var row   = sheet.getRange(rowIndex, 1, 1, KIT_COL.PHONE_SENT_TO).getValues()[0];
    var kitName      = String(row[KIT_COL.KIT - 1]          || '').trim();
    var learnerName  = String(row[KIT_COL.LEARNER_NAME - 1] || '').trim();
    var existingJlid = String(row[KIT_COL.JLID - 1]         || '').trim();

    // Update JLID if provided and different
    if (jlid && jlid !== existingJlid) {
      sheet.getRange(rowIndex, KIT_COL.JLID).setValue(jlid);
      Logger.log('[KitTracking] updateKitRow: JLID=' + jlid + ' row=' + rowIndex);
    }
    var effectiveJlid = jlid || existingJlid;

    // Update delivery date + compute time taken
    if (deliveryDate) {
      sheet.getRange(rowIndex, KIT_COL.DELIVERY_DATE).setValue(deliveryDate);
      var orderRaw  = row[KIT_COL.DATE_OF_ORDER - 1];
      var orderDate = (orderRaw instanceof Date) ? orderRaw : _parseDMY(String(orderRaw || ''));
      var delivDate = _parseDMY(deliveryDate);
      if (orderDate && delivDate) {
        var days = Math.round((delivDate - orderDate) / 86400000);
        sheet.getRange(rowIndex, KIT_COL.TIME_TAKEN).setValue(days + ' days');
      }
      Logger.log('[KitTracking] updateKitRow: DeliveryDate=' + deliveryDate + ' row=' + rowIndex);
    }

    // Update parent response
    if (response) {
      sheet.getRange(rowIndex, KIT_COL.PARENT_RESPONSE).setValue(response);
      Logger.log('[KitTracking] updateKitRow: Response=”' + response + '” row=' + rowIndex);
    }

    var hsStatus = 'skipped';

    // Kit Received → update HubSpot kit status
    if (response === 'Kit Received') {
      if (!effectiveJlid) {
        hsStatus = 'no_jlid';
        Logger.log('[KitTracking] updateKitRow: Kit Received but no JLID — HubSpot NOT updated.');
      } else {
        var prop = _kitPropertyForType(kitName);
        if (!prop) {
          hsStatus = 'unknown_kit';
          Logger.log('[KitTracking] updateKitRow: unknown kit type "' + kitName + '" — cannot map to HubSpot property.');
        } else {
          try {
            var hsLookup = fetchHubspotByJlid(effectiveJlid);
            if (!hsLookup || !hsLookup.success || !hsLookup.data.dealId) {
              hsStatus = 'no_deal';
              Logger.log('[KitTracking] updateKitRow: no dealId for ' + effectiveJlid);
            } else {
              var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
              var url   = 'https://api.hubapi.com/crm/v3/objects/deals/' + hsLookup.data.dealId;
              var patchBody = {};
              patchBody[prop] = 'Received by the Parents';
              var resp = UrlFetchApp.fetch(url, {
                method: 'PATCH',
                headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
                payload: JSON.stringify({ properties: patchBody }),
                muteHttpExceptions: true
              });
              var code = resp.getResponseCode();
              Logger.log('[KitTracking] updateKitRow: HubSpot PATCH ' + prop + ' HTTP ' + code + ' body=' + resp.getContentText().substring(0, 200));
              hsStatus = (code === 200) ? 'updated' : ('http_' + code + ': ' + resp.getContentText().substring(0, 100));
            }
          } catch (hsErr) {
            hsStatus = 'error: ' + hsErr.message;
            Logger.log('[KitTracking] updateKitRow: HubSpot error: ' + hsErr.message);
          }
        }
      }
    }

    // Not Received / Need To Check → add HubSpot note
    if ((response === 'Not Received yet' || response === 'Need To check') && effectiveJlid) {
      try {
        var hsN = fetchHubspotByJlid(effectiveJlid);
        if (hsN && hsN.success && hsN.data.dealId) {
          _addNoteToDeal(hsN.data.dealId, '[Kit Tracking] Manual update: ' + response + ' for ' + kitName + ' — ' + learnerName);
        }
      } catch (noteErr) {
        Logger.log('[KitTracking] updateKitRow: note failed: ' + noteErr.message);
      }
    }

    // Build user-facing message
    var msg = 'Sheet updated.';
    if (response === 'Kit Received') {
      if (hsStatus === 'updated')       msg = 'Sheet updated ✓  HubSpot ' + prop + ' → Received by the Parents ✓';
      else if (hsStatus === 'no_jlid')  msg = 'Sheet updated ✓  HubSpot SKIPPED — JLID missing. Add JLID and save again.';
      else if (hsStatus === 'no_deal')  msg = 'Sheet updated ✓  HubSpot SKIPPED — deal not found for ' + effectiveJlid;
      else if (hsStatus === 'unknown_kit') msg = 'Sheet updated ✓  HubSpot SKIPPED — kit type "' + kitName + '" not mapped.';
      else                              msg = 'Sheet updated ✓  HubSpot failed: ' + hsStatus;
    }

    return { success: true, message: msg, hsStatus: hsStatus };
  } catch (e) {
    Logger.log('[KitTracking] updateKitRow ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// REGISTER DAILY TRIGGER  (call once from initializeSystem or manually)
function setupKitTrackingTrigger() {
  // Remove any existing duplicate triggers first
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendKitFollowUps') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('sendKitFollowUps')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  Logger.log('[KitTracking] Daily trigger registered for sendKitFollowUps at 8am.');
}

