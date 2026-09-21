// ─────────────────────────────────────────────────────────────────────────────
// MigrationAutomationService.js
// Polls HubSpot migration pipeline every 10 min (time trigger).
// Picks up tickets in "Execution Pending" stage and executes them fully.
// CLS-gated reasons wait for "Approved by CLS" stage before executing.
// ─────────────────────────────────────────────────────────────────────────────

var MIG_AUTO = (function() {

  // ── Pipeline constants ──────────────────────────────────────────────────────
  var PIPELINE_ID        = '66161281';

  var STAGE = {
    TRIGGERED:       '128913747',
    WIP:             '128913748',
    TP_PENDING:      '128913750',
    CLS_PENDING:     '128913752',
    CLS_REJECTED:    '1030980247',
    CLS_APPROVED:    '133755411',
    EXEC_PENDING:    '1065336836',
    PR_PENDING:      '128913749',
    COMPLETED:       '128913753',
    CANCELLED:       ['133821818', '153457301']
  };

  // ── Reasons that require CLS approval before execution ─────────────────────
  var CLS_REQUIRED_REASONS = [
    'escalation on teacher',
    'teacher performance issue',
    'course change after prm',
    'teacher change after prm',
    'escalation on teacher post migration',
    'misalignment of teacher due to ops intervention only'
  ];

  // ── Ticket properties to fetch ─────────────────────────────────────────────
  var TICKET_PROPERTIES = [
    'subject', 'learner_uid', 'hs_pipeline_stage',
    'new_teacher', 'current_teacher__t_',
    'reason_of_migration__t_',
    'current_course__t_', 'future_course_1', 'future_course_2', 'future_course_3',
    'regular_class_day__t_', 'regular_class_time__in_cet_',
    'pre_migration_last_class_conducted_date__t_',
    'mig_auto_attempted', 'mig_auto_error'
  ];

  // ── Helpers ─────────────────────────────────────────────────────────────────

  // Convert "17:00" (24h CET) → "5:00 PM" (12h AM/PM) for bookClassesWithNewTeacher
  function _cetTo12h(t) {
    if (!t) return '';
    var m = String(t).match(/^(\d{1,2}):(\d{2})/);
    if (!m) return t;
    var h = parseInt(m[1], 10), mn = m[2];
    var ap = h >= 12 ? 'PM' : 'AM';
    var h12 = h % 12 || 12;
    return h12 + ':' + mn + ' ' + ap;
  }

  // Normalize abbreviated or lowercased day names to full title-case
  // e.g. "tue" / "Tue" / "TUESDAY" → "Tuesday"
  var _DAY_FULL = {
    sun: 'Sunday', mon: 'Monday', tue: 'Tuesday', wed: 'Wednesday',
    thu: 'Thursday', fri: 'Friday', sat: 'Saturday'
  };
  function _normalizeDay(d) {
    if (!d) return d;
    var key = String(d).toLowerCase().slice(0, 3);
    return _DAY_FULL[key] || d;
  }

  function _token() {
    return PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  }

  function _needsCls(reason) {
    var r = (reason || '').toLowerCase().trim();
    return CLS_REQUIRED_REASONS.some(function(cr) { return r.indexOf(cr) !== -1; });
  }

  function _patchTicketStage(ticketId, stageId) {
    monitoredFetch('https://api.hubapi.com/crm/v3/objects/tickets/' + ticketId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + _token(), 'Content-Type': 'application/json' },
      payload: JSON.stringify({ properties: { hs_pipeline_stage: stageId } }),
      muteHttpExceptions: true
    });
  }

  function _addTicketNote(ticketId, body) {
    try {
      var noteId = _createHubSpotNote(body);
      if (!noteId) return;
      monitoredFetch('https://api.hubapi.com/crm/v4/objects/notes/' + noteId + '/associations/tickets/' + ticketId + '/202', {
        method: 'PUT',
        headers: { 'Authorization': 'Bearer ' + _token(), 'Content-Type': 'application/json' },
        payload: '{}',
        muteHttpExceptions: true
      });
    } catch(e) {
      Logger.log('[MIG_AUTO] addTicketNote error: ' + e.message);
    }
  }

  function _createHubSpotNote(body) {
    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/notes', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + _token(), 'Content-Type': 'application/json' },
      payload: JSON.stringify({ properties: { hs_note_body: body, hs_timestamp: String(new Date().getTime()) } }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() === 201) return JSON.parse(resp.getContentText()).id;
    return null;
  }

  function _patchTicketProps(ticketId, props) {
    monitoredFetch('https://api.hubapi.com/crm/v3/objects/tickets/' + ticketId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + _token(), 'Content-Type': 'application/json' },
      payload: JSON.stringify({ properties: props }),
      muteHttpExceptions: true
    });
  }

  // ── Fetch all tickets ready for automation ──────────────────────────────────
  // Polls: Execution Pending + CLS Approved (for CLS-gated reasons)
  function _fetchActionableTickets() {
    var url  = 'https://api.hubapi.com/crm/v3/objects/tickets/search';
    var body = {
      filterGroups: [
        {
          filters: [
            { propertyName: 'hs_pipeline',       operator: 'EQ',     value: PIPELINE_ID },
            { propertyName: 'hs_pipeline_stage', operator: 'IN',     values: [STAGE.EXEC_PENDING, STAGE.CLS_APPROVED] },
            { propertyName: 'mig_auto_attempted', operator: 'NOT_HAS_PROPERTY' }
          ]
        }
      ],
      properties: TICKET_PROPERTIES,
      limit: 50
    };

    var resp = monitoredFetch(url, {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + _token(), 'Content-Type': 'application/json' },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      Logger.log('[MIG_AUTO] Ticket fetch failed: ' + resp.getContentText().substring(0, 200));
      return [];
    }
    return JSON.parse(resp.getContentText()).results || [];
  }

  // ── Build execution data object from ticket + deal ──────────────────────────
  function _buildExecutionData(ticket) {
    var props = ticket.properties || {};
    var jlid  = props.learner_uid || '';
    if (!jlid) return null;

    // Fetch deal + hybrid data (reuse existing function)
    var hybridResult = fetchMigrationHybridData(jlid);
    if (!hybridResult.success) {
      Logger.log('[MIG_AUTO] fetchMigrationHybridData failed for ' + jlid + ': ' + hybridResult.message);
      return null;
    }

    var deal = hybridResult.data;

    var newTeacher = getTeacherLabel(props.new_teacher) || '';
    var oldTeacher = deal.currentTeacher || '';
    var reason     = props.reason_of_migration__t_ || deal.migrationReason || '';
    var course     = props.current_course__t_
                       ? (getCourseLabel(props.current_course__t_) || props.current_course__t_)
                       : (deal.course || '');

    // Confirmed future courses (non-NA)
    var confirmedFutureCourses = [];
    [props.future_course_1, props.future_course_2, props.future_course_3].forEach(function(raw) {
      if (!raw) return;
      var s = raw.toLowerCase().trim();
      if (s === 'na' || s === 'n/a' || s === 'not applicable' || s === '-') return;
      try { confirmedFutureCourses.push(getCourseLabel(raw) || raw); } catch(e) { confirmedFutureCourses.push(raw); }
    });

    // Fetch existing calendar event once — extract class link, description, title prefix, and start time
    var _evClassLink = deal.zoomLink || '';
    var _evDesc      = '';
    var _evPrefix    = 'Migration';
    var _evStart     = '';
    var _evTz        = '';
    try {
      var _evRes = getExistingEventDescription(jlid);
      _evStart = _evRes.eventStart     || '';
      _evTz    = _evRes.eventTimeZone  || '';
      if (_evRes && _evRes.success && _evRes.description) {
        var _evLinkMatch = _evRes.description.match(/https?:\/\/\S+/);
        if (_evLinkMatch) _evClassLink = _evLinkMatch[0].replace(/[)\]>]+$/, '');
        _evDesc = _evRes.description;
        var _prefixMatch = (_evRes.eventTitle || '').match(/^([A-Z][A-Z0-9\/\- ]+?)\s*:/);
        if (_prefixMatch) _evPrefix = _prefixMatch[1].trim();
        Logger.log('[MIG_AUTO] Existing event: title=' + _evRes.eventTitle + ' link=' + _evClassLink + ' prefix=' + _evPrefix);
      }
      if (_evStart) Logger.log('[MIG_AUTO] Event start=' + _evStart + ' tz=' + _evTz);
    } catch(_evErr) {
      Logger.log('[MIG_AUTO] getExistingEventDescription error: ' + _evErr.message);
    }

    return {
      jlid:                   jlid,
      learner:                deal.learnerName || '',
      newTeacher:             newTeacher,
      oldTeacher:             oldTeacher,
      course:                 course,
      reasonOfMigration:      reason,
      confirmedFutureCourses: confirmedFutureCourses,
      timezone:               deal.timezone || '',
      manualTimezone:         deal.suggestedIana || '',
      classSessions:          (function() {
        // 1. Ticket schedule (has both day + CET time)
        var ts = deal.ticketSchedule;
        Logger.log('[MIG_AUTO] ticketSchedule=' + JSON.stringify(ts) + ' dealSessions=' + JSON.stringify(deal.classSessions));
        if (ts && ts.day && ts.time) return [{ day: _normalizeDay(ts.day), time: _cetTo12h(ts.time) }];
        // 2. Deal sessions that have a time
        var ds = (deal.classSessions || []).filter(function(s) { return s.time && s.time.trim(); });
        if (ds.length) return ds.map(function(s) { return { day: _normalizeDay(s.day), time: s.time }; });
        // 3. Derive day + time from existing calendar event start datetime
        if (_evStart && _evTz) {
          try {
            var evDt  = new Date(_evStart);
            var evDay  = Utilities.formatDate(evDt, _evTz, 'EEEE');   // e.g. "Tuesday"
            var evTime = Utilities.formatDate(evDt, _evTz, 'h:mm a'); // e.g. "2:00 PM"
            if (evDay && evTime) {
              Logger.log('[MIG_AUTO] classSessions derived from event: ' + evDay + ' ' + evTime + ' (' + _evTz + ')');
              return [{ day: evDay, time: evTime }];
            }
          } catch(_evTzErr) {
            Logger.log('[MIG_AUTO] Event time parse error: ' + _evTzErr.message);
          }
        }
        return (deal.classSessions || []);
      })(),
      classBookingIana:       (function() {
        // If using CET ticket time, book in CET so Calendar API interprets it correctly
        var ts = deal.ticketSchedule;
        if (ts && ts.day && ts.time) return 'Europe/Paris';
        var ds = (deal.classSessions || []).filter(function(s) { return s.time && s.time.trim(); });
        if (ds.length) return deal.suggestedIana || deal.timezone || 'Europe/London';
        // If falling back to event time, use the event's own timezone
        if (_evStart && _evTz) return _evTz;
        return deal.suggestedIana || deal.timezone || 'Europe/London';
      })(),
      zoomLink:               _evClassLink,
      existingEventDesc:      _evDesc,
      migrationPrefix:        _evPrefix,
      watiPhoneTargets:       deal.parentPhone ? [deal.parentPhone] : [],
      parentName:             deal.parentName || '',
      parentEmail:            deal.parentEmail || '',
      sendEmailToTeacher:     true,
      sendWhatsappToParent:   true,
      sendEmailToParentAlso:  false,
      sendCertificate:        (reason.toLowerCase().indexOf('course change') !== -1),
      performedBy:            'MigrationAutomation',
      _ticketId:              ticket.id,
      _dealId:                deal.dealId || ''
    };
  }

  // ── Validate ticket is safe to auto-execute ─────────────────────────────────
  function _validate(data, ticket) {
    var props  = ticket.properties || {};
    var stage  = props.hs_pipeline_stage || '';
    var reason = (data.reasonOfMigration || '').toLowerCase();

    // CLS-gated reasons must be in CLS_APPROVED stage
    if (_needsCls(reason) && stage !== STAGE.CLS_APPROVED) {
      return 'CLS approval required but ticket not in Approved stage';
    }

    if (!data.newTeacher)        return 'new_teacher missing on ticket';
    if (!data.jlid)              return 'learner_uid missing on ticket';
    if (!data.reasonOfMigration) return 'reason_of_migration missing on ticket';

    // CCTC: block if pre-migration last class date is in future
    var isCourseChange = reason.indexOf('course change') !== -1;
    if (isCourseChange && props.pre_migration_last_class_conducted_date__t_) {
      var preMs   = new Date(props.pre_migration_last_class_conducted_date__t_).getTime();
      var todayMs = new Date().setHours(0, 0, 0, 0);
      if (!isNaN(preMs) && preMs > todayMs) {
        return 'CCTC: pre-migration last class date is in the future (' +
               new Date(preMs).toDateString() + ') — cannot execute yet';
      }
    }

    return null; // valid
  }

  // ── Execute one ticket ──────────────────────────────────────────────────────
  function _executeTicket(ticket) {
    var ticketId = ticket.id;
    Logger.log('[MIG_AUTO] Processing ticket ' + ticketId);

    // Mark as attempted immediately to prevent double-processing
    _patchTicketProps(ticketId, { mig_auto_attempted: 'true' });

    var data = _buildExecutionData(ticket);
    if (!data) {
      _addTicketNote(ticketId, '⚠️ Auto-execution failed: could not build execution data (deal/ticket mismatch).');
      _patchTicketProps(ticketId, { mig_auto_error: 'Could not build execution data' });
      return;
    }

    var validationError = _validate(data, ticket);
    if (validationError) {
      Logger.log('[MIG_AUTO] Validation failed for ' + ticketId + ': ' + validationError);
      _addTicketNote(ticketId, '⚠️ Auto-execution skipped: ' + validationError);
      _patchTicketProps(ticketId, { mig_auto_error: validationError, mig_auto_attempted: 'skipped' });
      return;
    }

    // Run the full migration
    try {
      var result = sendMigrationEmail(data);

      if (result && result.success !== false) {
        Logger.log('[MIG_AUTO] ✅ Ticket ' + ticketId + ' executed successfully');

        // ── Auto-book classes ───────────────────────────────────────────────
        var bookingNote = '';
        try {
          if (!data.zoomLink) {
            bookingNote = '⚠️ Classes not auto-booked — no class link found on event or contact. Please book manually.';
            Logger.log('[MIG_AUTO] Skipping class booking for ' + ticketId + ': no zoom link');
          } else if (!data.classSessions || !data.classSessions.length) {
            bookingNote = '⚠️ Classes not auto-booked — no class sessions found on deal. Please book manually.';
            Logger.log('[MIG_AUTO] Skipping class booking for ' + ticketId + ': no classSessions');
          } else {
            var remainingRes = getRemainingClassesForJlid(data.jlid);
            var numEvents = (remainingRes && remainingRes.classesLeft > 0) ? remainingRes.classesLeft : 12;
            var iana = data.classBookingIana || data.manualTimezone || data.timezone || 'Europe/London';
            var today = new Date();
            var pad = function(n) { return ('0' + n).slice(-2); };
            var startDate = today.getFullYear() + '-' + pad(today.getMonth() + 1) + '-' + pad(today.getDate());

            var bookRes = bookClassesWithNewTeacher(
              data.jlid,
              data.learner,
              data.newTeacher,
              data.classSessions,
              data.course,
              startDate,
              iana,
              numEvents,
              data.parentEmail ? [data.parentEmail] : [],
              'MigrationAutomation',
              data.zoomLink,
              '',
              data.existingEventDesc || '',
              data.migrationPrefix || 'Migration'
            );
            if (bookRes && bookRes.success !== false) {
              bookingNote = '📅 ' + (bookRes.booked || numEvents) + ' classes auto-booked with ' + data.newTeacher + '.';
              Logger.log('[MIG_AUTO] Classes booked for ' + ticketId + ': ' + (bookRes.booked || numEvents));
            } else {
              bookingNote = '⚠️ Class booking failed: ' + ((bookRes && bookRes.message) || 'Unknown error') + '. Please book manually.';
              Logger.log('[MIG_AUTO] Class booking failed for ' + ticketId + ': ' + ((bookRes && bookRes.message) || ''));
            }
          }
        } catch(be) {
          bookingNote = '⚠️ Class booking error: ' + be.message + '. Please book manually.';
          Logger.log('[MIG_AUTO] Class booking exception for ' + ticketId + ': ' + be.message);
        }

        // CCTC: update future_course_1 to current course (new course after migration)
        _updateCctcFutureCourse(data, ticket);

        _addTicketNote(ticketId, '✅ Migration auto-executed successfully by MigrationAutomation.\n' + bookingNote);
        _patchTicketStage(ticketId, STAGE.COMPLETED);
        _patchTicketProps(ticketId, { mig_auto_error: '' });
      } else {
        var errMsg = (result && result.message) ? result.message : 'Unknown error';
        Logger.log('[MIG_AUTO] ❌ Ticket ' + ticketId + ' execution failed: ' + errMsg);
        _addTicketNote(ticketId, '❌ Auto-execution failed: ' + errMsg + '\nWill retry on next cycle if error is resolved.');
        _patchTicketProps(ticketId, { mig_auto_error: errMsg, mig_auto_attempted: 'failed' });
      }
    } catch(e) {
      Logger.log('[MIG_AUTO] Exception for ticket ' + ticketId + ': ' + e.message);
      _addTicketNote(ticketId, '❌ Auto-execution exception: ' + e.message);
      _patchTicketProps(ticketId, { mig_auto_error: e.message, mig_auto_attempted: 'failed' });
    }
  }

  // ── Main entry point — called by GAS time trigger ───────────────────────────
  function run() {
    Logger.log('[MIG_AUTO] === Starting automation run === ' + new Date().toISOString());

    var tickets = _fetchActionableTickets();
    Logger.log('[MIG_AUTO] Found ' + tickets.length + ' actionable tickets');

    tickets.forEach(function(ticket) {
      try {
        _executeTicket(ticket);
      } catch(e) {
        Logger.log('[MIG_AUTO] Unhandled error for ticket ' + ticket.id + ': ' + e.message);
      }
    });

    Logger.log('[MIG_AUTO] === Run complete ===');
  }

  // ── CCTC: set future_course_1 = new course on ticket after execution ─────
  function _updateCctcFutureCourse(data, ticket) {
    try {
      var reason = (data.reasonOfMigration || '').toLowerCase();
      var isCctc = reason.indexOf('course change') !== -1;
      if (!isCctc) return;
      var props = ticket.properties || {};
      var fc1 = props.future_course_1 || ''; // new course they're moving TO
      if (!fc1) return;
      var fc2 = props.future_course_2 || '';
      var fc3 = props.future_course_3 || '';

      // Update DEAL: current_course = new course, shift future courses up
      var dealId = data._dealId || '';
      if (dealId) {
        monitoredFetch('https://api.hubapi.com/crm/v3/objects/deals/' + dealId, {
          method: 'PATCH',
          headers: { 'Authorization': 'Bearer ' + _token(), 'Content-Type': 'application/json' },
          payload: JSON.stringify({ properties: { current_course: fc1 } }),
          muteHttpExceptions: true
        });
        Logger.log('[MIG_AUTO] CCTC: deal ' + dealId + ' current_course → ' + fc1);
      }

      Logger.log('[MIG_AUTO] CCTC ticket courses left as-is (no automated update).');
    } catch(e) {
      Logger.log('[MIG_AUTO] _updateCctcFutureCourse error: ' + e.message);
    }
  }

  // ── Single-ticket manual trigger (from UI) ───────────────────────────────
  function runForJlid(jlid) {
    if (!jlid) return { success: false, message: 'JLID required.' };

    var token = _token();
    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/tickets/search', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({
        filterGroups: [{
          filters: [
            { propertyName: 'hs_pipeline',      operator: 'EQ', value: PIPELINE_ID },
            { propertyName: 'learner_uid',       operator: 'EQ', value: jlid },
            { propertyName: 'hs_pipeline_stage', operator: 'IN', values: [STAGE.EXEC_PENDING, STAGE.CLS_APPROVED] }
          ]
        }],
        properties: TICKET_PROPERTIES,
        sorts: [{ propertyName: 'createdate', direction: 'DESCENDING' }],
        limit: 1
      }),
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) return { success: false, message: 'HubSpot search failed.' };
    var results = JSON.parse(resp.getContentText()).results || [];
    if (!results.length) {
      return { success: false, message: 'No ticket in Execution Pending or CLS Approved stage for ' + jlid + '. Move the ticket to that stage first.' };
    }

    var ticket = results[0];
    var ticketId = ticket.id;
    var props = ticket.properties || {};

    // Allow re-run from UI even if previously attempted
    if (props.mig_auto_attempted) _patchTicketProps(ticketId, { mig_auto_attempted: '' });

    var data = _buildExecutionData(ticket);
    if (!data) return { success: false, message: 'Could not build execution data — deal/ticket mismatch for ' + jlid + '.' };

    var validationError = _validate(data, ticket);
    if (validationError) {
      _patchTicketProps(ticketId, { mig_auto_error: validationError, mig_auto_attempted: 'skipped' });
      _addTicketNote(ticketId, '⚠️ Auto-execution skipped: ' + validationError);
      return { success: false, message: 'Validation failed: ' + validationError };
    }

    _patchTicketProps(ticketId, { mig_auto_attempted: 'true' });

    var messages = [];

    try {
      var result = sendMigrationEmail(data);
      if (!result || result.success === false) {
        var errMsg = (result && result.message) ? result.message : 'Migration email failed';
        _patchTicketProps(ticketId, { mig_auto_error: errMsg, mig_auto_attempted: 'failed' });
        _addTicketNote(ticketId, '❌ Auto-execution failed: ' + errMsg);
        return { success: false, message: errMsg };
      }
      messages.push('Emails sent.');
    } catch(e) {
      _patchTicketProps(ticketId, { mig_auto_error: e.message, mig_auto_attempted: 'failed' });
      _addTicketNote(ticketId, '❌ Auto-execution exception: ' + e.message);
      return { success: false, message: e.message };
    }

    var bookingMsg = '';
    try {
      Logger.log('[MIG_AUTO] booking check: zoomLink=' + (data.zoomLink ? 'Y' : 'N') + ' sessions=' + JSON.stringify(data.classSessions) + ' iana=' + data.classBookingIana);
      if (!data.zoomLink) {
        bookingMsg = 'No class link found (event or contact) — book classes manually.';
      } else if (!data.classSessions || !data.classSessions.length) {
        bookingMsg = 'No class sessions on deal — book manually.';
      } else {
        var rem = getRemainingClassesForJlid(data.jlid);
        var n = (rem && rem.classesLeft > 0) ? rem.classesLeft : 12;
        var iana = data.classBookingIana || data.manualTimezone || data.timezone || 'Europe/London';
        var today = new Date();
        var p = function(v) { return ('0' + v).slice(-2); };
        var sd = today.getFullYear() + '-' + p(today.getMonth() + 1) + '-' + p(today.getDate());
        var bk = bookClassesWithNewTeacher(
          data.jlid, data.learner, data.newTeacher, data.classSessions,
          data.course, sd, iana, n,
          data.parentEmail ? [data.parentEmail] : [],
          'MigrationAutomation', data.zoomLink, '',
          data.existingEventDesc || '',
          data.migrationPrefix || 'Migration'
        );
        if (bk && bk.success !== false) {
          var bookedCount = bk.occurrences || n;
          bookingMsg = 'Classes Booked: ' + bookedCount + ' | Timezone: ' + iana + ' | Booked with ' + data.newTeacher + '.';
        } else {
          bookingMsg = 'Class booking failed: ' + ((bk && bk.message) || 'Unknown') + '. Book manually.';
        }
      }
    } catch(be) {
      bookingMsg = 'Class booking error: ' + be.message + '. Book manually.';
    }

    messages.push(bookingMsg);
    // CCTC: update future_course_1 to current course
    _updateCctcFutureCourse(data, ticket);

    var _slotStatus = 'Matched by Ops';
    var _intervenedBy = 'Ops Intervention';
    _addTicketNote(ticketId, '✅ Migration auto-executed (manual trigger).\n' + bookingMsg);
    _patchTicketStage(ticketId, STAGE.COMPLETED);
    _patchTicketProps(ticketId, {
      mig_auto_error: '',
      migration_slot_status__t_: _slotStatus,
      migration_intervened_by: _intervenedBy
    });

    // Build timeline for UI progress popup
    var timeline = (result && result.timeline) ? result.timeline.slice() : [];
    var bookingStatus = (bookingMsg.indexOf('error') !== -1 || bookingMsg.indexOf('failed') !== -1 || bookingMsg.indexOf('manually') !== -1) ? 'warning' : 'success';
    timeline.push({ key: 'class_booking',   label: bookingMsg || 'Classes booked.', status: bookingStatus, durationMs: 0 });
    timeline.push({ key: 'ticket_complete', label: 'HubSpot ticket moved to Completed.', status: 'success', durationMs: 0 });

    return { success: true, message: '✅ ' + messages.join(' '), timeline: timeline };
  }

  return { run: run, runForJlid: runForJlid };

})();

// ── GAS time trigger entry point ─────────────────────────────────────────────
function runMigrationAutomation() {
  MIG_AUTO.run();
}

// ── Manual single-ticket auto migration — called from migration page UI ───────
function runAutoMigrateForJlid(jlid) {
  return MIG_AUTO.runForJlid(jlid);
}

// ── Automation stats — called from client dashboard ──────────────────────────
function getMigrationAutomationStats() {
  var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');

  function _search(filters, props, limit) {
    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/tickets/search', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ filterGroups: [{ filters: filters }], properties: props, limit: limit || 100 }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return [];
    return JSON.parse(resp.getContentText()).results || [];
  }

  var BASE = [
    { propertyName: 'hs_pipeline', operator: 'EQ', value: '66161281' }
  ];

  var PROPS = ['subject', 'learner_uid', 'hs_pipeline_stage', 'mig_auto_attempted', 'mig_auto_error',
               'createdate', 'hs_lastmodifieddate', 'reason_of_migration__t_', 'new_teacher'];

  // Auto-executed: mig_auto_attempted = 'true' and stage = Completed
  var autoExec = _search(BASE.concat([
    { propertyName: 'mig_auto_attempted', operator: 'EQ', value: 'true' },
    { propertyName: 'hs_pipeline_stage',  operator: 'EQ', value: '128913753' }
  ]), PROPS, 100);

  // Failed: mig_auto_attempted = 'failed'
  var failed = _search(BASE.concat([
    { propertyName: 'mig_auto_attempted', operator: 'EQ', value: 'failed' }
  ]), PROPS, 100);

  // Skipped: mig_auto_attempted = 'skipped'
  var skipped = _search(BASE.concat([
    { propertyName: 'mig_auto_attempted', operator: 'EQ', value: 'skipped' }
  ]), PROPS, 100);

  // Pending (still in Execution Pending, no mig_auto_attempted yet)
  var pending = _search(BASE.concat([
    { propertyName: 'hs_pipeline_stage',   operator: 'EQ',              value: '1065336836' },
    { propertyName: 'mig_auto_attempted',  operator: 'NOT_HAS_PROPERTY' }
  ]), PROPS, 50);

  function _mapTicket(t) {
    var p = t.properties || {};
    return {
      id:       t.id,
      subject:  p.subject || p.learner_uid || t.id,
      jlid:     p.learner_uid || '',
      stage:    p.hs_pipeline_stage || '',
      error:    p.mig_auto_error || '',
      reason:   p.reason_of_migration__t_ || '',
      newTeacher: p.new_teacher || '',
      modified: p.hs_lastmodifieddate || p.createdate || ''
    };
  }

  return {
    success:        true,
    autoExecuted:   autoExec.length,
    failed:         failed.length,
    skipped:        skipped.length,
    pendingQueue:   pending.length,
    needManual:     failed.length + skipped.length,
    failedTickets:  failed.map(_mapTicket),
    skippedTickets: skipped.map(_mapTicket),
    pendingTickets: pending.map(_mapTicket),
    fetchedAt:      new Date().toISOString()
  };
}
