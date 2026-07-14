// =============================================
// CREDENTIALS SERVICE
// Manages Scratch (and future platform) credentials:
//   - Generate next sequential username
//   - Log to credentials spreadsheet
//   - Auto-update Google Calendar event description
// =============================================

var CREDS_SPREADSHEET_ID = '1KsyxldnHpm7gEyTcmmQFkz-uaqTM_FMhNTxh7OXBCTk';
var SCRATCH_SHEET_TAB    = 'Scratch Credentials';
var SCRATCH_PREFIX       = 'JLRCB';
var SCRATCH_PASSWORD     = 'jetlearn';

// ─── Public ──────────────────────────────────────────────────────────────────

/**
 * Assign the next available Scratch credential (col A = username, col B = password)
 * to a learner by writing their name (col C) and JLID (col D).
 * Picks the first row where col C is empty — no new usernames are generated.
 *
 * @param {string} jlid        - e.g. "JL55030989090C"
 * @param {string} learnerName - display name
 * @returns {{ success, username, password, calendarUpdated }}
 */
function generateScratchCredentials(jlid, learnerName) {
  try {
    if (!jlid) return { success: false, error: 'JLID required' };

    var ss    = SpreadsheetApp.openById(CREDS_SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SCRATCH_SHEET_TAB);
    if (!sheet) return { success: false, error: 'Sheet "' + SCRATCH_SHEET_TAB + '" not found' };

    var data = sheet.getDataRange().getValues();

    // Find first row where col A (username) exists and col C (learner name) is empty
    var targetRow = -1;
    var username  = '';
    var password  = '';
    for (var i = 0; i < data.length; i++) {
      var colA = String(data[i][0] || '').trim();
      var colC = String(data[i][2] || '').trim();
      if (colA && !colC) {
        targetRow = i + 1; // 1-indexed sheet row
        username  = colA;
        password  = String(data[i][1] || SCRATCH_PASSWORD).trim();
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, error: 'No available Scratch credentials left. Please add more to the sheet.' };
    }

    // Write learner name (col C) and JLID (col D) to claim the credential
    sheet.getRange(targetRow, 3).setValue(learnerName || jlid);
    sheet.getRange(targetRow, 4).setValue(jlid);

    // Update calendar event description
    var credText = 'Student Account on MIT Scratch coding platform:\nLogin URL: https://scratch.mit.edu/\nLogin Name: ' + username + '\nTemporary Password: ' + password;
    var calUpdated = _updateCalendarWithCredentials(jlid, credText);

    Logger.log('[Credentials] Assigned ' + username + ' to ' + jlid + ' | calendar=' + calUpdated);

    return {
      success:          true,
      username:         username,
      password:         password,
      calendarUpdated:  calUpdated > 0,
      eventsUpdated:    calUpdated,
      platform:         'Scratch'
    };

  } catch(e) {
    Logger.log('[CredentialsService] generateScratchCredentials error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Preview the next available username without assigning it.
 */
function peekNextScratchUsername() {
  try {
    var ss    = SpreadsheetApp.openById(CREDS_SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SCRATCH_SHEET_TAB);
    if (!sheet) return { success: false, error: 'Sheet not found' };
    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      var colA = String(data[i][0] || '').trim();
      var colC = String(data[i][2] || '').trim();
      if (colA && !colC) return { success: true, username: colA };
    }
    return { success: false, error: 'No available credentials left.' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/**
 * Fetch upcoming calendar sessions for a JLID — shown as preview before generating credentials.
 * Returns session count, CET times, attendees, and current description of first event.
 */
function getScratchCalendarPreview(jlid) {
  try {
    var calendarId = CONFIG.CLASS_SCHEDULE_CALENDAR_ID;
    if (!calendarId) return { success: false, error: 'Calendar not configured.' };
    if (!jlid)       return { success: false, error: 'JLID required.' };

    var now    = new Date();
    var future = new Date();
    future.setFullYear(future.getFullYear() + 3);

    var items = Calendar.Events.list(calendarId, {
      q:            jlid,
      timeMin:      now.toISOString(),
      timeMax:      future.toISOString(),
      singleEvents: true,
      orderBy:      'startTime',
      maxResults:   500
    }).items || [];

    if (items.length === 0) {
      return { success: true, count: 0, sessions: [], description: '', attendees: [] };
    }

    var sessions = items.slice(0, 5).map(function(ev) {
      var start   = ev.start && ev.start.dateTime ? new Date(ev.start.dateTime) : null;
      var end     = ev.end   && ev.end.dateTime   ? new Date(ev.end.dateTime)   : null;
      var dayCET      = start ? Utilities.formatDate(start, 'Europe/Paris', 'EEEE') : '—';
      var startCET    = start ? Utilities.formatDate(start, 'Europe/Paris', 'HH:mm') : '—';
      var endCET      = end   ? Utilities.formatDate(end,   'Europe/Paris', 'HH:mm') : '—';
      return { day: dayCET, time: startCET + ' – ' + endCET + ' CET' };
    });

    // Last session = subscription end date
    var lastEvent = items[items.length - 1];
    var lastStart = lastEvent.start && lastEvent.start.dateTime ? new Date(lastEvent.start.dateTime) : null;
    var lastSessionDate = lastStart ? Utilities.formatDate(lastStart, 'Europe/Paris', 'EEE, d MMM yyyy') : '—';

    var firstAttendees = (items[0].attendees || []).map(function(a) {
      return a.displayName || a.email || '';
    }).filter(Boolean);

    var currentDesc = (items[0].description || '').trim();

    return {
      success:         true,
      count:           items.length,
      sessions:        sessions,
      lastSessionDate: lastSessionDate,
      attendees:       firstAttendees,
      description:     currentDesc
    };

  } catch(e) {
    Logger.log('[CredentialsService] getScratchCalendarPreview error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Search Google Calendar for events whose title contains the JLID.
 * Appends credText to each matching event's description.
 *
 * @returns {number} count of events updated
 */
function _updateCalendarWithCredentials(jlid, credText) {
  try {
    var calendarId = CONFIG.CLASS_SCHEDULE_CALENDAR_ID;
    if (!calendarId) {
      Logger.log('[Credentials] No calendar ID configured.');
      return 0;
    }

    var now = new Date();
    var future = new Date();
    future.setFullYear(future.getFullYear() + 3); // search up to 3 years ahead

    // Use Calendar API (same as verifySubscriptionWithCalendar) for accurate search
    var items = Calendar.Events.list(calendarId, {
      q: jlid,
      timeMin: now.toISOString(),
      timeMax: future.toISOString(),
      singleEvents: true,
      orderBy: 'startTime',
      maxResults: 500
    }).items || [];

    if (items.length === 0) {
      Logger.log('[Credentials] No upcoming calendar events found for JLID: ' + jlid);
      return 0;
    }

    var updated = 0;
    for (var i = 0; i < items.length; i++) {
      var event = items[i];
      var desc = (event.description || '').trim();

      // Remove any stale Scratch credentials block
      desc = desc.replace(/Student Account on MIT Scratch coding platform:[\s\S]*?Temporary Password: \S+/g, '').trim();

      var newDesc = desc + (desc ? '\n\n' : '') + credText;

      Calendar.Events.patch({ description: newDesc }, calendarId, event.id);
      updated++;
    }

    Logger.log('[Credentials] Updated ' + updated + ' calendar events for JLID ' + jlid);
    return updated;

  } catch(e) {
    Logger.log('[CredentialsService] _updateCalendarWithCredentials error: ' + e.message);
    return 0;
  }
}
// v1
