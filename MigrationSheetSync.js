// ============================================================
// Migration Pipeline → Google Sheet Sync
// Runs daily at 7 AM via time-based GAS trigger.
// Fetches all migration tickets from HubSpot (Jan 1 2026 →
// today), excludes cancelled/rejected stages, and rewrites
// the tracking sheet in full on every run.
// ============================================================

var SYNC_CONFIG = {
  SHEET_ID:    '11VjxS9TYbSHw6_UMyavvEeSB2JsMCDSHz2wHTaHe1oo',
  PIPELINE_ID: '66161281',
  FROM_DATE:   '2026-01-01',          // pull tickets created on or after this date

  // Stage 10 variants — always excluded
  EXCLUDE_STAGES: ['133821818', '153457301'],

  STAGE_LABELS: {
    '128913747':  'Migration Triggered',
    '128913748':  'WIP',
    '128913750':  'WIP - TP Approval Pending',
    '128913752':  'WIP - CLS Approval Pending',
    '1030980247': 'WIP - Rejected by CLS',
    '133755411':  'WIP - Approved by CLS',
    '1065336836': 'Execution Pending',
    '128913749':  'WIP - PR Approval Pending',
    '128913753':  'Migration Completed'
  },

  // HubSpot stores "date entered stage" as hs_date_entered_{stageId}
  // We request all of them so we can pick the right one per ticket.
  DATE_ENTERED_PROPS: [
    'hs_date_entered_128913747',
    'hs_date_entered_128913748',
    'hs_date_entered_128913750',
    'hs_date_entered_128913752',
    'hs_date_entered_1030980247',
    'hs_date_entered_133755411',
    'hs_date_entered_1065336836',
    'hs_date_entered_128913749',
    'hs_date_entered_128913753'
  ],

  COLUMNS: [
    'Ticket ID',
    'Learner Name',
    'Learner UID',
    'Current Teacher',
    'New Teacher',
    'Reason of Migration',
    'Stage',
    'Date Entered Stage',
    'Ticket Created Date'
  ]
};

// ── Helpers ──────────────────────────────────────────────────

function _fmtDate(val) {
  if (!val) return '';
  try {
    var d = new Date(typeof val === 'string' && /^\d+$/.test(val) ? parseInt(val, 10) : val);
    if (isNaN(d.getTime())) return String(val);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } catch (e) { return String(val); }
}

function _getToken() {
  return PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY')
      || CONFIG.HUBSPOT_API_KEY;
}

// ── Main sync function ────────────────────────────────────────

function syncMigrationPipelineToSheet() {
  Logger.log('[MigrationSync] Starting sync — ' + new Date().toLocaleString());

  var token   = _getToken();
  var headers = { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' };
  var fromMs  = new Date(SYNC_CONFIG.FROM_DATE).getTime();

  var properties = [
    'subject',
    'learner_full_name',
    'learner_uid',
    'current_teacher__t_',
    'new_teacher',
    'reason_of_migration__t_',
    'hs_pipeline_stage',
    'createdate'
  ].concat(SYNC_CONFIG.DATE_ENTERED_PROPS);

  // ── 1. Paginate through all matching tickets ─────────────────
  var rows  = [];
  var after = undefined;
  var page  = 0;

  do {
    var body = {
      filterGroups: [{
        filters: [
          { propertyName: 'hs_pipeline', operator: 'EQ',  value: SYNC_CONFIG.PIPELINE_ID },
          { propertyName: 'createdate',  operator: 'GTE', value: fromMs }
        ]
      }],
      properties: properties,
      sorts:      [{ propertyName: 'createdate', direction: 'DESCENDING' }],
      limit:      100
    };
    if (after) body.after = after;

    var resp = UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/objects/tickets/search', {
      method:             'post',
      headers:            headers,
      payload:            JSON.stringify(body),
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      Logger.log('[MigrationSync] HubSpot error: ' + resp.getContentText());
      break;
    }

    var data = JSON.parse(resp.getContentText());
    if (!data.results || data.results.length === 0) break;

    data.results.forEach(function(ticket) {
      var props   = ticket.properties || {};
      var stageId = String(props.hs_pipeline_stage || '').trim();

      // Skip excluded stages (Stage 10)
      if (SYNC_CONFIG.EXCLUDE_STAGES.indexOf(stageId) !== -1) return;

      var stageLabel     = SYNC_CONFIG.STAGE_LABELS[stageId] || ('Stage ' + stageId);
      var dateEnteredRaw = props['hs_date_entered_' + stageId] || '';
      var dateEntered    = _fmtDate(dateEnteredRaw);
      var createdDate    = _fmtDate(props.createdate || ticket.createdAt);

      // Learner name: prefer dedicated property, fall back to ticket subject
      var learnerName = (props.learner_full_name || props.subject || '').trim();

      rows.push([
        ticket.id                               || '',  // Ticket ID
        learnerName,                                    // Learner Name
        props.learner_uid                       || '',  // Learner UID
        props.current_teacher__t_               || '',  // Current Teacher
        props.new_teacher                       || '',  // New Teacher
        props.reason_of_migration__t_           || '',  // Reason of Migration
        stageLabel,                                     // Stage
        dateEntered,                                    // Date Entered Stage
        createdDate                                     // Ticket Created Date
      ]);
    });

    after = (data.paging && data.paging.next) ? data.paging.next.after : undefined;
    page++;

  } while (after && page < 100); // safety cap at 10,000 tickets

  Logger.log('[MigrationSync] Fetched ' + rows.length + ' tickets across ' + page + ' pages.');

  // ── 2. Write to Google Sheet ─────────────────────────────────
  var ss    = SpreadsheetApp.openById(SYNC_CONFIG.SHEET_ID);
  var sheet = ss.getSheets()[0];

  sheet.clearContents();

  // Header row
  var headerRange = sheet.getRange(1, 1, 1, SYNC_CONFIG.COLUMNS.length);
  headerRange.setValues([SYNC_CONFIG.COLUMNS]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4f46e5');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontSize(10);
  sheet.setFrozenRows(1);

  // Data rows
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, SYNC_CONFIG.COLUMNS.length).setValues(rows);

    // Zebra stripe for readability
    for (var i = 0; i < rows.length; i++) {
      var rowRange = sheet.getRange(i + 2, 1, 1, SYNC_CONFIG.COLUMNS.length);
      rowRange.setBackground(i % 2 === 0 ? '#ffffff' : '#f8f7ff');
    }
  }

  // Auto-resize all columns
  for (var c = 1; c <= SYNC_CONFIG.COLUMNS.length; c++) {
    sheet.autoResizeColumn(c);
  }

  // Last synced timestamp in top-right (column 10)
  sheet.getRange(1, SYNC_CONFIG.COLUMNS.length + 1)
    .setValue('Last synced: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'))
    .setFontColor('#94a3b8')
    .setFontSize(9)
    .setFontWeight('normal');

  Logger.log('[MigrationSync] Sheet updated. ' + rows.length + ' rows written.');
  return { success: true, rowsWritten: rows.length };
}

// ── Trigger management ────────────────────────────────────────

/**
 * Run this ONCE from the GAS editor to set up the daily 7 AM trigger.
 * After that, syncMigrationPipelineToSheet() runs automatically every day.
 */
function setupMigrationSyncTrigger() {
  // Remove any existing trigger for this function to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncMigrationPipelineToSheet') {
      ScriptApp.deleteTrigger(t);
      Logger.log('[MigrationSync] Removed existing trigger.');
    }
  });

  // Create daily trigger at 7 AM in the script's timezone
  ScriptApp.newTrigger('syncMigrationPipelineToSheet')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();

  Logger.log('[MigrationSync] Trigger created: syncMigrationPipelineToSheet runs daily at 7 AM.');
}

/**
 * Check what triggers are currently active for this script.
 */
function listMigrationSyncTriggers() {
  var triggers = ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction() === 'syncMigrationPipelineToSheet';
  });
  if (triggers.length === 0) {
    Logger.log('[MigrationSync] No trigger found. Run setupMigrationSyncTrigger() to create one.');
  } else {
    triggers.forEach(function(t) {
      Logger.log('[MigrationSync] Trigger active: ' + t.getHandlerFunction() + ' | Type: ' + t.getTriggerSource());
    });
  }
}
