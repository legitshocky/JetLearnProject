// =============================================
// CREDENTIALS SERVICE
// Manages Scratch (and future platform) credentials:
//   - Generate next sequential username
//   - Log to credentials spreadsheet
//   - Auto-update Google Calendar event description
// =============================================

var CREDS_SPREADSHEET_ID = '1KsyxldnHpm7gEyTcmmQFkz-uaqTM_FMhNTxh7OXBCTk';
var SCRATCH_SHEET_TAB    = 'Scratch Credentials';
var SCRATCH_PREFIX       = 'SHJLK';
var SCRATCH_PASSWORD     = 'jetlearn';

// ─── Public ──────────────────────────────────────────────────────────────────

/**
 * Generate next Scratch username for a learner, log to sheet, update calendar.
 * Called from client after user has created the account on scratch.mit.edu.
 *
 * @param {string} jlid        - e.g. "JL55030989090C"
 * @param {string} learnerName - display name for logging
 * @returns {{ success, username, password, calendarUpdated, registerLink, nextUsername }}
 */
function generateScratchCredentials(jlid, learnerName) {
  try {
    if (!jlid) return { success: false, error: 'JLID required' };

    // 1. Get next username
    var username = _getNextScratchUsername();

    // 2. Log to credentials sheet
    var ss    = SpreadsheetApp.openById(CREDS_SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SCRATCH_SHEET_TAB);
    if (!sheet) return { success: false, error: 'Sheet "' + SCRATCH_SHEET_TAB + '" not found in credentials spreadsheet' };
    sheet.appendRow([username, SCRATCH_PASSWORD, learnerName || jlid, jlid, new Date()]);

    // 3. Update calendar event description
    var calUpdated = _updateCalendarWithCredentials(jlid, 'Scratch = ' + username + '\npass = ' + SCRATCH_PASSWORD);

    Logger.log('[Credentials] Generated ' + username + ' for ' + jlid + ' | calendar=' + calUpdated);

    return {
      success:         true,
      username:        username,
      password:        SCRATCH_PASSWORD,
      calendarUpdated: calUpdated,
      registerLink:    'https://scratch.mit.edu/join',
      platform:        'Scratch'
    };

  } catch(e) {
    Logger.log('[CredentialsService] generateScratchCredentials error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Peek at the next Scratch username without committing — shown in UI before generation.
 */
function peekNextScratchUsername() {
  try {
    return { success: true, username: _getNextScratchUsername() };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ─── Private ─────────────────────────────────────────────────────────────────

/**
 * Read Scratch Credentials sheet, find highest SHJLK number, return incremented username.
 */
function _getNextScratchUsername() {
  var ss      = SpreadsheetApp.openById(CREDS_SPREADSHEET_ID);
  var sheet   = ss.getSheetByName(SCRATCH_SHEET_TAB);
  var data    = sheet.getDataRange().getValues();
  var lastNum = 0;

  for (var i = 0; i < data.length; i++) {
    var u = String(data[i][0] || '').trim().toUpperCase();
    if (u.startsWith(SCRATCH_PREFIX)) {
      var n = parseInt(u.replace(SCRATCH_PREFIX, ''), 10);
      if (!isNaN(n) && n > lastNum) lastNum = n;
    }
  }

  var next = lastNum + 1;
  // Zero-pad to at least 2 digits: SHJLK01 … SHJLK09 SHJLK10 … SHJLK99
  return SCRATCH_PREFIX + (next < 10 ? '0' : '') + next;
}

/**
 * Search Google Calendar for events whose title contains the JLID.
 * Appends credText to each matching event's description.
 * Searches from today → 180 days ahead.
 *
 * @returns {boolean} true if at least one event was updated
 */
function _updateCalendarWithCredentials(jlid, credText) {
  try {
    var cal = CalendarApp.getCalendarById(CONFIG.CLASS_SCHEDULE_CALENDAR_ID);
    if (!cal) {
      Logger.log('[Credentials] Calendar not found: ' + CONFIG.CLASS_SCHEDULE_CALENDAR_ID);
      return false;
    }

    var now    = new Date();
    var future = new Date(now.getTime() + 180 * 24 * 60 * 60 * 1000);
    var events = cal.getEvents(now, future);

    // JLID in calendar title looks like: "Mazin (JL55030989090C) : ..."
    // Strip trailing C for safer matching — covers both "JL...C" and bare "JL..."
    var jlidBase = jlid.replace(/C$/i, '');
    var updated  = 0;

    for (var i = 0; i < events.length; i++) {
      var title = events[i].getTitle();
      if (title.indexOf(jlidBase) === -1 && title.indexOf(jlid) === -1) continue;

      var desc = events[i].getDescription() || '';

      // Remove stale Scratch credentials block if present
      desc = desc.replace(/Scratch = SHJLK\d+\s*\npass = \w+/g, '').trim();

      events[i].setDescription(desc + (desc ? '\n\n' : '') + credText);
      updated++;
    }

    Logger.log('[Credentials] Calendar events updated: ' + updated + ' for JLID ' + jlid);
    return updated > 0;

  } catch(e) {
    Logger.log('[CredentialsService] _updateCalendarWithCredentials error: ' + e.message);
    return false;
  }
}
// v1
