// ============================================================
// EMAIL QUEUE SERVICE
// Stores scheduled emails in "Email Queue" sheet.
// Daily trigger: processEmailQueue() at 8am
// ============================================================

var EQ_SHEET_NAME = 'Email Queue';
var EQ_COL = {
  QUEUE_ID:       1,  // A
  SCHEDULED_DATE: 2,  // B  (DD/MM/YYYY)
  EMAIL_TYPE:     3,  // C
  JLID:           4,  // D
  RECIPIENT:      5,  // E
  LEARNER_NAME:   6,  // F
  FORM_DATA_JSON: 7,  // G
  STATUS:         8,  // H  Pending / Sent / Failed / Cancelled
  CREATED_AT:     9,  // I
  CREATED_BY:     10, // J
  SENT_AT:        11, // K
  ERROR:          12  // L
};

// ── Get or create Email Queue sheet ──────────────────────────
function _getEmailQueueSheet() {
  var ssId = PropertiesService.getScriptProperties().getProperty('SHEET_ID_KIT_TRACKING')
             || '17Jsa2Kl2AkI5SgtlITYGqb-Q-PxfzkNFzzG5HwBJp_Q';
  var ss    = SpreadsheetApp.openById(ssId);
  var sheet = ss.getSheetByName(EQ_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(EQ_SHEET_NAME);
    sheet.appendRow(['Queue ID','Scheduled Date','Email Type','JLID','Recipient Email',
                     'Learner Name','Form Data (JSON)','Status','Created At','Created By','Sent At','Error']);
    sheet.setFrozenRows(1);
    Logger.log('[EmailQueue] Created "Email Queue" sheet.');
  }
  return sheet;
}

// ── Format Date → DD/MM/YYYY ──────────────────────────────────
function _eqFmtDate(date) {
  if (!date) return '';
  var d = date instanceof Date ? date : new Date(date);
  if (isNaN(d.getTime())) return String(date);
  var dd = d.getDate(); var mm = d.getMonth() + 1; var yyyy = d.getFullYear();
  return (dd < 10 ? '0'+dd : dd) + '/' + (mm < 10 ? '0'+mm : mm) + '/' + yyyy;
}

// ── Parse DD/MM/YYYY → Date ───────────────────────────────────
function _eqParseDate(str) {
  if (!str) return null;
  var s = String(str).trim();
  if (s instanceof Date) return s;
  var parts = s.indexOf('/') > -1 ? s.split('/') : s.split('-');
  if (parts.length !== 3) return null;
  // DD/MM/YYYY
  var d = parseInt(parts[0],10), m = parseInt(parts[1],10)-1, y = parseInt(parts[2],10);
  // OR YYYY-MM-DD (from HTML date input)
  if (parts[0].length === 4) { y = parseInt(parts[0],10); m = parseInt(parts[1],10)-1; d = parseInt(parts[2],10); }
  var dt = new Date(y, m, d);
  return isNaN(dt.getTime()) ? null : dt;
}

// ── Public: Add email to queue ────────────────────────────────
/**
 * @param {object} payload {
 *   scheduledDate  {string}  YYYY-MM-DD from HTML date input
 *   emailType      {string}  'minecraft'|'roblox'|'onboardingParent'|'renewal'|'invoice'
 *   jlid           {string}
 *   recipientEmail {string}
 *   learnerName    {string}
 *   formData       {object}  full form data object
 *   comments       {string}
 *   createdBy      {string}
 * }
 */
function scheduleEmail(payload) {
  try {
    if (!payload || !payload.scheduledDate) return { success: false, error: 'scheduledDate required' };
    if (!payload.emailType)                 return { success: false, error: 'emailType required' };
    if (!payload.recipientEmail && payload.emailType !== 'onboardingParent')
                                            return { success: false, error: 'recipientEmail required' };

    var sheet   = _getEmailQueueSheet();
    var queueId = Utilities.getUuid();

    // Parse scheduled date — accept YYYY-MM-DD or DD/MM/YYYY
    var schedParts = (payload.scheduledDate || '').split('-');
    var schedFmt   = '';
    if (schedParts.length === 3 && schedParts[0].length === 4) {
      // YYYY-MM-DD → DD/MM/YYYY
      schedFmt = schedParts[2] + '/' + schedParts[1] + '/' + schedParts[0];
    } else {
      schedFmt = payload.scheduledDate; // already DD/MM/YYYY
    }

    sheet.appendRow([
      queueId,
      schedFmt,
      payload.emailType,
      payload.jlid           || '',
      payload.recipientEmail || '',
      payload.learnerName    || '',
      JSON.stringify(payload.formData || {}),
      'Pending',
      new Date(),
      payload.createdBy      || '',
      '',  // Sent At
      ''   // Error
    ]);

    Logger.log('[EmailQueue] Scheduled ' + payload.emailType + ' for ' + schedFmt + ' JLID=' + payload.jlid);
    return { success: true, queueId: queueId, scheduledDate: schedFmt };

  } catch(e) {
    Logger.log('[EmailQueue] scheduleEmail ERROR: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── Public: Cancel a queued email ─────────────────────────────
function cancelQueuedEmail(queueId) {
  try {
    if (!queueId) return { success: false, error: 'queueId required' };
    var sheet   = _getEmailQueueSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Queue empty' };

    var data = sheet.getRange(2, 1, lastRow - 1, EQ_COL.STATUS).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][EQ_COL.QUEUE_ID - 1]) === queueId) {
        var status = String(data[i][EQ_COL.STATUS - 1]);
        if (status !== 'Pending') return { success: false, error: 'Cannot cancel — status is ' + status };
        sheet.getRange(i + 2, EQ_COL.STATUS).setValue('Cancelled');
        Logger.log('[EmailQueue] Cancelled queueId=' + queueId);
        return { success: true };
      }
    }
    return { success: false, error: 'Queue ID not found' };
  } catch(e) {
    Logger.log('[EmailQueue] cancelQueuedEmail ERROR: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── Public: Get queue rows for UI ─────────────────────────────
function getEmailQueue() {
  try {
    var sheet   = _getEmailQueueSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, rows: [] };

    var data = sheet.getRange(2, 1, lastRow - 1, EQ_COL.ERROR).getValues();
    var tz   = Session.getScriptTimeZone();

    function fmtTs(v) {
      if (!v) return '';
      if (v instanceof Date) return Utilities.formatDate(v, tz, 'dd/MM/yyyy HH:mm');
      return String(v);
    }

    var rows = data.map(function(r, idx) {
      return {
        rowIndex:      idx + 2,
        queueId:       String(r[EQ_COL.QUEUE_ID       - 1] || ''),
        scheduledDate: String(r[EQ_COL.SCHEDULED_DATE - 1] || ''),
        emailType:     String(r[EQ_COL.EMAIL_TYPE      - 1] || ''),
        jlid:          String(r[EQ_COL.JLID            - 1] || ''),
        recipient:     String(r[EQ_COL.RECIPIENT        - 1] || ''),
        learnerName:   String(r[EQ_COL.LEARNER_NAME    - 1] || ''),
        status:        String(r[EQ_COL.STATUS           - 1] || 'Pending'),
        createdAt:     fmtTs(r[EQ_COL.CREATED_AT       - 1]),
        createdBy:     String(r[EQ_COL.CREATED_BY       - 1] || ''),
        sentAt:        fmtTs(r[EQ_COL.SENT_AT           - 1]),
        error:         String(r[EQ_COL.ERROR            - 1] || '')
      };
    }).filter(function(r) { return r.queueId; });

    return { success: true, rows: rows };
  } catch(e) {
    Logger.log('[EmailQueue] getEmailQueue ERROR: ' + e.message);
    return { success: false, rows: [], error: e.message };
  }
}

// ── Daily trigger: process due emails ────────────────────────
function processEmailQueue() {
  Logger.log('[EmailQueue] processEmailQueue started');
  try {
    var sheet   = _getEmailQueueSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) { Logger.log('[EmailQueue] Queue empty.'); return; }

    var today = new Date(); today.setHours(0, 0, 0, 0);
    var data  = sheet.getRange(2, 1, lastRow - 1, EQ_COL.ERROR).getValues();
    var sent  = 0; var failed = 0;

    data.forEach(function(r, idx) {
      var sheetRow = idx + 2;
      var status   = String(r[EQ_COL.STATUS - 1] || '');
      if (status !== 'Pending') return; // skip sent/cancelled/failed

      var schedStr  = String(r[EQ_COL.SCHEDULED_DATE - 1] || '');
      var schedDate = _eqParseDate(schedStr);
      if (!schedDate) {
        Logger.log('[EmailQueue] Row ' + sheetRow + ': cannot parse date "' + schedStr + '"');
        return;
      }
      schedDate.setHours(0, 0, 0, 0);
      if (schedDate > today) return; // not yet due

      // Due — attempt send
      var emailType  = String(r[EQ_COL.EMAIL_TYPE      - 1] || '');
      var jlid       = String(r[EQ_COL.JLID            - 1] || '');
      var recipient  = String(r[EQ_COL.RECIPIENT        - 1] || '');
      var learner    = String(r[EQ_COL.LEARNER_NAME    - 1] || '');
      var fdJson     = String(r[EQ_COL.FORM_DATA_JSON  - 1] || '{}');

      Logger.log('[EmailQueue] Processing row ' + sheetRow + ' type=' + emailType + ' jlid=' + jlid);

      try {
        var fd = JSON.parse(fdJson);
        var result = _dispatchQueuedEmail(emailType, fd, recipient);

        if (result && result.success) {
          sheet.getRange(sheetRow, EQ_COL.STATUS).setValue('Sent');
          sheet.getRange(sheetRow, EQ_COL.SENT_AT).setValue(new Date());
          Logger.log('[EmailQueue] Sent row ' + sheetRow + ' ' + emailType + ' → ' + recipient);
          sent++;
        } else {
          var errMsg = (result && result.message) || 'Unknown error';
          sheet.getRange(sheetRow, EQ_COL.STATUS).setValue('Failed');
          sheet.getRange(sheetRow, EQ_COL.ERROR).setValue(errMsg);
          Logger.log('[EmailQueue] Failed row ' + sheetRow + ': ' + errMsg);
          failed++;
        }
      } catch(sendErr) {
        sheet.getRange(sheetRow, EQ_COL.STATUS).setValue('Failed');
        sheet.getRange(sheetRow, EQ_COL.ERROR).setValue(sendErr.message);
        Logger.log('[EmailQueue] Send error row ' + sheetRow + ': ' + sendErr.message);
        failed++;
      }
    });

    Logger.log('[EmailQueue] processEmailQueue done. sent=' + sent + ' failed=' + failed);
  } catch(e) {
    Logger.log('[EmailQueue] processEmailQueue ERROR: ' + e.message + '\n' + e.stack);
  }
}

// ── Dispatch to correct send function ────────────────────────
function _dispatchQueuedEmail(emailType, formData, recipientEmail) {
  switch(emailType) {
    case 'minecraft':
      return sendGenericEmail('Minecraft Install', formData, recipientEmail || formData.recipientEmail, [], formData.comments || '');
    case 'roblox':
      return sendGenericEmail('Roblox Install', formData, recipientEmail || formData.recipientEmail, [], formData.comments || '');
    case 'onboardingParent':
      return sendParentOnboardingWithInvoice(formData, []);
    case 'renewal':
      return sendRenewalEmail(formData, []);
    case 'invoice':
      return sendInvoiceEmail(formData, []);
    default:
      return { success: false, message: 'Unknown email type: ' + emailType };
  }
}

// ── Register daily trigger ────────────────────────────────────
function setupEmailQueueTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processEmailQueue') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('processEmailQueue')
    .timeBased().everyDays(1).atHour(8).create();
  Logger.log('[EmailQueue] Daily trigger registered for processEmailQueue at 8am.');
}
