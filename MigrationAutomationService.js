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
      classSessions:          deal.classSessions || [],
      watiPhoneTargets:       deal.parentPhone ? [deal.parentPhone] : [],
      parentName:             deal.parentName || '',
      parentEmail:            deal.parentEmail || '',
      sendEmailToTeacher:     true,
      sendWhatsappToParent:   true,
      sendEmailToParentAlso:  false,
      addComplimentaryClasses: false,
      performedBy:            'MigrationAutomation',
      _ticketId:              ticket.id
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
        _addTicketNote(ticketId, '✅ Migration auto-executed successfully by MigrationAutomation.');
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

  return { run: run };

})();

// ── GAS time trigger entry point ─────────────────────────────────────────────
function runMigrationAutomation() {
  MIG_AUTO.run();
}
