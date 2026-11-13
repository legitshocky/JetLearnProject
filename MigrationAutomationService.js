// ─────────────────────────────────────────────────────────────────────────────
// MigrationAutomationService.js
// Polls HubSpot migration pipeline every 10 min (time trigger).
// Picks up tickets in "Execution Pending" stage and executes them fully.
// CLS-gated reasons wait for "Approved by CLS" stage before executing.
//
// Rules this file keeps:
//  - Nothing is sent to a parent/teacher until we know the classes CAN be booked
//    (class link + a readable schedule) — otherwise the ticket is left in place as
//    "blocked" with the reason on the ticket, and re-checked every ~20 minutes.
//  - A ticket is only moved to Completed when the classes were really booked and the
//    ticket's required fields were written. Every HubSpot write is checked.
//  - mig_auto_attempted: 'true' = ran, 'blocked' = waiting on something (auto re-checked),
//    'failed' = ran but something broke after messages went out (never auto-retried —
//    use Auto Migrate), 'skipped' = legacy value, treated like 'blocked'.
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
    'migration_slot_status__t_',
    'migration_execution_status',   // label "Migration Intervened By"
    'mig_auto_attempted', 'mig_auto_error'
  ];

  var TIME_BUDGET_MS = 240000;       // stop starting new tickets after 4 min (Apps Script caps a run at 6 min)
  var RECHECK_SECONDS = 1200;        // blocked tickets are re-evaluated at most every 20 min

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
    var key = String(d).trim().toLowerCase().slice(0, 3);
    return _DAY_FULL[key] || d;
  }

  // Convert time to 12h AM/PM — handles both "13:00" (24h) and "1:00 PM" / "1 PM" (already 12h)
  function _toAmPm(t) {
    if (!t) return '';
    if (/[AP]M/i.test(t)) return t.trim();
    return _cetTo12h(t);
  }

  function _token() {
    return PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  }

  function _needsCls(reason) {
    var r = (reason || '').toLowerCase().trim();
    return CLS_REQUIRED_REASONS.some(function(cr) { return r.indexOf(cr) !== -1; });
  }

  function _httpOk(code) { return code === 200 || code === 201 || code === 204; }

  function _isNa(v) {
    var s = String(v || '').trim().toLowerCase();
    return !s || s === 'na' || s === 'n/a' || s === 'not applicable' || s === '-';
  }

  // Fetch hs_pipeline_stage history to determine who intervened on this ticket
  function _getTicketIntervention(ticketId) {
    try {
      var resp = monitoredFetch(
        'https://api.hubapi.com/crm/v3/objects/tickets/' + ticketId + '?propertiesWithHistory=hs_pipeline_stage',
        { method: 'GET', headers: { 'Authorization': 'Bearer ' + _token() }, muteHttpExceptions: true }
      );
      if (resp.getResponseCode() !== 200) return null;
      var data = JSON.parse(resp.getContentText());
      var history = (data.propertiesWithHistory && data.propertiesWithHistory.hs_pipeline_stage) || [];
      var stages = history.map(function(h) { return h.value; });
      Logger.log('[MIG_AUTO] stage history: ' + JSON.stringify(stages));

      var hadTP  = stages.indexOf('128913750') !== -1;
      var hadCLS = stages.indexOf('133755411') !== -1 || stages.indexOf('128913752') !== -1;

      var intervenedBy;
      if (hadCLS && hadTP)    intervenedBy = 'CLS Intervention;TP Intervention';
      else if (hadCLS)        intervenedBy = 'CLS Intervention';
      else if (hadTP)         intervenedBy = 'TP Intervention';
      else                    intervenedBy = 'Ops Intervention';

      var slotStatus = hadTP ? 'Matched by TP' : 'Matched by Ops';

      return { intervenedBy: intervenedBy, slotStatus: slotStatus };
    } catch(e) {
      Logger.log('[MIG_AUTO] _getTicketIntervention error: ' + e.message);
      return null;
    }
  }

  // Every HubSpot write returns { ok, code, body } and is logged — nothing fails silently any more.
  function _patchTicket(ticketId, props, label) {
    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/tickets/' + ticketId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + _token(), 'Content-Type': 'application/json' },
      payload: JSON.stringify({ properties: props }),
      muteHttpExceptions: true
    });
    var code = resp ? resp.getResponseCode() : 0;
    var body = (resp && code && !_httpOk(code)) ? resp.getContentText().slice(0, 300) : '';
    Logger.log('[MIG_AUTO] ' + (label || 'patch') + ' ticket ' + ticketId + ' → HTTP ' + code + (body ? ' body=' + body : ''));
    return { ok: _httpOk(code), code: code, body: body };
  }

  function _patchTicketProps(ticketId, props) { return _patchTicket(ticketId, props, 'props'); }
  function _patchTicketStage(ticketId, stageId) { return _patchTicket(ticketId, { hs_pipeline_stage: stageId }, 'stage→' + stageId); }

  function _createHubSpotNote(body) {
    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/notes', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + _token(), 'Content-Type': 'application/json' },
      payload: JSON.stringify({ properties: { hs_note_body: body, hs_timestamp: String(new Date().getTime()) } }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() === 201) return JSON.parse(resp.getContentText()).id;
    Logger.log('[MIG_AUTO] note create failed HTTP ' + resp.getResponseCode() + ' ' + resp.getContentText().slice(0, 200));
    return null;
  }

  // Notes were created but never attached to the ticket: association type 202 is note→CONTACT, so HubSpot
  // refused it and nobody checked. The default association is the right one (same call the deal notes use).
  function _addTicketNote(ticketId, body) {
    try {
      var noteId = _createHubSpotNote(body);
      if (!noteId) return false;
      var resp = monitoredFetch('https://api.hubapi.com/crm/v4/objects/notes/' + noteId + '/associations/default/tickets/' + ticketId, {
        method: 'PUT',
        headers: { 'Authorization': 'Bearer ' + _token(), 'Content-Type': 'application/json' },
        muteHttpExceptions: true
      });
      var code = resp.getResponseCode();
      if (!_httpOk(code)) {
        Logger.log('[MIG_AUTO] note ' + noteId + ' → ticket ' + ticketId + ' association FAILED HTTP ' + code + ' ' + resp.getContentText().slice(0, 200));
        return false;
      }
      return true;
    } catch(e) {
      Logger.log('[MIG_AUTO] addTicketNote error: ' + e.message);
      return false;
    }
  }

  // ── Fetch all tickets ready for automation ──────────────────────────────────
  // Execution Pending + CLS Approved (for CLS-gated reasons), either never attempted or
  // currently 'blocked'/'skipped' (waiting on something that may since have been fixed).
  function _fetchActionableTickets() {
    var url  = 'https://api.hubapi.com/crm/v3/objects/tickets/search';
    var base = [
      { propertyName: 'hs_pipeline',       operator: 'EQ', value: PIPELINE_ID },
      { propertyName: 'hs_pipeline_stage', operator: 'IN', values: [STAGE.EXEC_PENDING, STAGE.CLS_APPROVED] }
    ];
    var body = {
      filterGroups: [
        { filters: base.concat([{ propertyName: 'mig_auto_attempted', operator: 'NOT_HAS_PROPERTY' }]) },
        { filters: base.concat([{ propertyName: 'mig_auto_attempted', operator: 'IN', values: ['blocked', 'skipped'] }]) }
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

  // The ticket stores the learner's regular slot as e.g. day "Tuesday;Saturday" + time "9 PM;5 PM",
  // or several days with ONE time ("2 PM"). Pair them up; refuse (with a reason) if it is ambiguous.
  function _sessionsFromTicketSchedule(dayStr, timeStr) {
    var split = function(s) { return String(s || '').split(/[;,]/).map(function(x) { return x.trim(); }).filter(Boolean); };
    var days = split(dayStr), times = split(timeStr);
    if (!days.length || !times.length) return { sessions: [], error: '' };

    var pairs;
    if (days.length === times.length)  pairs = days.map(function(d, i) { return [d, times[i]]; });
    else if (times.length === 1)       pairs = days.map(function(d) { return [d, times[0]]; });
    else if (days.length === 1)        pairs = times.map(function(t) { return [days[0], t]; });
    else return { sessions: [], error: 'Ticket schedule is ambiguous — ' + days.length + ' days ("' + dayStr + '") but ' + times.length + ' times ("' + timeStr + '"). Fix the ticket so each day has a time.' };

    var sessions = pairs.map(function(p) { return { day: _normalizeDay(p[0]), time: _toAmPm(p[1]) }; });
    var bad = sessions.filter(function(s) { return _DAY_INDEX[s.day] === undefined || !_parse12hTime(s.time); });
    if (bad.length) return { sessions: [], error: 'Ticket schedule could not be read (day "' + dayStr + '", time "' + timeStr + '"). Use e.g. Tuesday + 5 PM.' };
    return { sessions: sessions, error: '' };
  }

  // ── Time zones ──────────────────────────────────────────────────────────────
  // The ticket (and the deal's first time value) hold the class slot in CET. The parent is always told the time in THEIR
  // timezone, and the teacher email converts back from the parent's timezone. So every downstream step (messages, deal sync,
  // teacher email, booking) must get the slot in the parent's timezone, converted for a real date so daylight saving is right.
  function _tzOffsetMin(date, tz) {
    var z = Utilities.formatDate(date, tz, 'Z');            // e.g. "+0100"
    var sign = z.charAt(0) === '-' ? -1 : 1;
    return sign * (parseInt(z.substr(1, 2), 10) * 60 + parseInt(z.substr(3, 2), 10));
  }
  function _instantInTz(y, mo, d, h, mi, tz) {
    var guess = Date.UTC(y, mo - 1, d, h, mi);
    var off = _tzOffsetMin(new Date(guess), tz);
    var t = guess - off * 60000;
    var off2 = _tzOffsetMin(new Date(t), tz);
    if (off2 !== off) t = guess - off2 * 60000;
    return new Date(t);
  }
  // The next real occurrence of {day,time} in fromTz, as an instant.
  function _nextInstant(session, fromTz) {
    var tm = _parse12hTime(session.time);
    if (!tm) throw new Error('cannot read the time "' + session.time + '"');
    var now = new Date();
    for (var i = 0; i < 8; i++) {
      var probe = new Date(now.getTime() + i * 86400000);
      if (Utilities.formatDate(probe, fromTz, 'EEEE') === session.day) {
        var y = parseInt(Utilities.formatDate(probe, fromTz, 'yyyy'), 10), mo = parseInt(Utilities.formatDate(probe, fromTz, 'M'), 10), d = parseInt(Utilities.formatDate(probe, fromTz, 'd'), 10);
        return _instantInTz(y, mo, d, tm.h, tm.m, fromTz);
      }
    }
    throw new Error('unknown weekday "' + session.day + '"');
  }
  function _slotInTz(session, fromTz, toTz) {
    var at = _nextInstant(session, fromTz);
    return { day: Utilities.formatDate(at, toTz, 'EEEE'), time: Utilities.formatDate(at, toTz, 'h:mm a') };
  }
  function _fmtInstant(at, tz) { return Utilities.formatDate(at, tz, 'EEE d MMM yyyy, h:mm a') + ' (' + tz + ')'; }

  // Links we recognise as the class link. A random URL in the event notes (e.g. scratch.mit.edu) must never become it.
  function _isClassLink(u) {
    return /^https?:\/\/(live\.jetlearn\.com\/join\/|[\w.-]*zoom\.us\/|meet\.google\.com\/|teams\.microsoft\.com\/)/i.test(String(u || ''));
  }

  // Where the class sessions come from, in priority order: ticket → deal → the learner's existing calendar event.
  function _resolveSessions(deal, evStart, evTz) {
    var ts = deal.ticketSchedule;
    Logger.log('[MIG_AUTO] ticketSchedule=' + JSON.stringify(ts) + ' dealSessions=' + JSON.stringify(deal.classSessions));

    // 1. Ticket schedule (day + CET time) — may hold several days and/or times
    if (ts && ts.day && ts.time) {
      var built = _sessionsFromTicketSchedule(ts.day, ts.time);
      if (built.sessions.length) return { sessions: built.sessions, error: '', sourceTz: 'Europe/Paris' };
      return { sessions: [], error: built.error, sourceTz: 'Europe/Paris' };
    }
    // 2. Deal sessions that have a time (normalize to 12h AM/PM)
    var ds = (deal.classSessions || []).filter(function(s) { return s.time && s.time.trim(); });
    if (ds.length) return {
      sessions: ds.map(function(s) {
        // time may be "6 PM;3 PM UK Time" — take only the first part (CET)
        var t = String(s.time).split(';')[0].trim();
        return { day: _normalizeDay(s.day), time: _toAmPm(t) };
      }),
      error: '', sourceTz: 'Europe/Paris'
    };
    // 3. Derive day + time from the existing calendar event start datetime
    if (evStart && evTz) {
      try {
        var evDt   = new Date(evStart);
        var evDay  = Utilities.formatDate(evDt, evTz, 'EEEE');   // e.g. "Tuesday"
        var evTime = Utilities.formatDate(evDt, evTz, 'h:mm a'); // e.g. "2:00 PM"
        if (evDay && evTime) {
          Logger.log('[MIG_AUTO] classSessions derived from event: ' + evDay + ' ' + evTime + ' (' + evTz + ')');
          return { sessions: [{ day: evDay, time: evTime }], error: '', sourceTz: evTz };
        }
      } catch(_evTzErr) {
        Logger.log('[MIG_AUTO] Event time parse error: ' + _evTzErr.message);
      }
    }
    return { sessions: (deal.classSessions || []), error: '', sourceTz: 'Europe/Paris' };
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
      if (_isNa(raw)) return;
      // objects, as the manual form sends them (the upskilling note reads .courseLabel; it re-checks the teacher itself)
      var lbl = raw;
      try { lbl = getCourseLabel(raw) || raw; } catch(e) {}
      confirmedFutureCourses.push({ courseLabel: lbl, isUpskilled: false });
    });

    // Course-change tickets: the NEW classes are for future_course_1, not the course being completed
    var isCctc = String(reason).toLowerCase().indexOf('course change') !== -1;
    var bookingCourse = course;
    if (isCctc && !_isNa(props.future_course_1)) {
      try { bookingCourse = getCourseLabel(props.future_course_1) || props.future_course_1; } catch(e) { bookingCourse = props.future_course_1; }
    }

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
        var _evUrls = (_evRes.description.match(/https?:\/\/[^\s)\]>]+/g) || []).filter(_isClassLink);
        if (_evUrls.length) _evClassLink = _evUrls[0];
        _evDesc = _evRes.description;
        var _prefixMatch = (_evRes.eventTitle || '').match(/^([A-Z][A-Z0-9\/\- ]+?)\s*:/);
        if (_prefixMatch) _evPrefix = _prefixMatch[1].trim();
        Logger.log('[MIG_AUTO] Existing event: title=' + _evRes.eventTitle + ' link=' + _evClassLink + ' prefix=' + _evPrefix);
      }
      if (_evStart) Logger.log('[MIG_AUTO] Event start=' + _evStart + ' tz=' + _evTz);
    } catch(_evErr) {
      Logger.log('[MIG_AUTO] getExistingEventDescription error: ' + _evErr.message);
    }

    // Never use an unrecognised URL (notes often contain scratch.mit.edu etc.) — default to the learner's standard class link
    if (!_isClassLink(_evClassLink)) {
      _evClassLink = 'https://live.jetlearn.com/join/' + cleanJlidForZoom(jlid);
    }

    // Parent emails: every contact on the deal, like the manual form's list
    var parentEmails = [];
    try {
      var er = deal.dealId ? getEmailsForDeal(deal.dealId) : null;
      ((er && er.all) || []).forEach(function(e) { e = String(e || '').trim(); if (e && isValidEmail(e) && parentEmails.indexOf(e) === -1) parentEmails.push(e); });
    } catch (ee) { Logger.log('[MIG_AUTO] getEmailsForDeal failed: ' + ee.message); }
    if (!parentEmails.length && deal.parentEmail && isValidEmail(deal.parentEmail)) parentEmails.push(deal.parentEmail);

    var sched = _resolveSessions(deal, _evStart, _evTz);

    // Convert the slot to the PARENT's timezone for everything the parent/teacher sees; booking uses the same converted slot.
    var srcTz = sched.sourceTz || 'Europe/Paris';
    var parentIana = deal.suggestedIana || '';
    var ticketSessions = sched.sessions;
    var messageSessions = sched.sessions;
    var timeError = '';
    var converted = false;
    if (sched.sessions.length && !sched.error) {
      if (!parentIana) {
        timeError = 'The parent\'s timezone is unknown, so the class time cannot be converted for the parent. Set the timezone on the deal.';
      } else {
        try {
          messageSessions = sched.sessions.map(function(sx) { return _slotInTz(sx, srcTz, parentIana); });
          converted = true;
        } catch (cvErr) {
          timeError = 'Could not convert the class time to the parent\'s timezone: ' + cvErr.message;
        }
      }
    }

    return {
      jlid:                   jlid,
      learner:                deal.learnerName || '',
      newTeacher:             newTeacher,
      oldTeacher:             oldTeacher,
      course:                 course,
      bookingCourse:          bookingCourse,
      reasonOfMigration:      reason,
      confirmedFutureCourses: confirmedFutureCourses,
      timezone:               deal.timezone || '',
      manualTimezone:         deal.suggestedIana || '',
      classSessions:          messageSessions,        // in the PARENT's timezone (what messages, deal sync and the teacher email expect)
      ticketSessions:         ticketSessions,         // as stored (CET)
      sessionsSourceTz:       srcTz,
      parentIana:             parentIana,
      timeError:              timeError,
      scheduleError:          sched.error,
      classBookingIana:       (function() {
        if (converted) return parentIana;
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
      classLink:              _evClassLink,   // used by WATI params + parent email
      existingEventDesc:      _evDesc,
      migrationPrefix:        _evPrefix,
      watiPhoneTargets:       deal.parentPhone ? [deal.parentPhone] : [],
      parentName:             deal.parentName || '',
      parentEmail:            deal.parentEmail || '',
      sendEmailToTeacher:     true,
      sendWhatsappToParent:   true,
      sendEmailToParentAlso:  parentEmails.length > 0,   // the parent is emailed too (same content as the WhatsApp)
      parentEmailTargets:     parentEmails,
      // the fields the manual form sends, so the teacher email and notes are identical
      clsManager:             deal.clsManagerName || '',
      jetGuide:               deal.jetGuideName || '',
      startDate:              Utilities.formatDate(new Date(), 'Europe/Paris', 'yyyy-MM-dd'),
      migrationType:          isCctc ? 'New Assignment' : 'Mid-Course',
      completedCourseForCert: course,
      sendCertificate:        isCctc,
      performedBy:            'MigrationAutomation',
      _ticketId:              ticket.id,
      _dealId:                deal.dealId || ''
    };
  }

  // Course-change tickets wait for the learner's last class in the OLD course. Cheap (no API calls).
  function _cctcWaitReason(props) {
    var reason = String(props.reason_of_migration__t_ || '').toLowerCase();
    if (reason.indexOf('course change') === -1) return '';
    var raw = props.pre_migration_last_class_conducted_date__t_;
    if (!raw) return '';
    var preMs   = new Date(raw).getTime();
    var todayMs = new Date().setHours(0, 0, 0, 0);
    if (!isNaN(preMs) && preMs > todayMs) {
      return 'Course change: the learner\'s last class in the current course is on ' + new Date(preMs).toDateString() +
             ' — it will run automatically after that date.';
    }
    return '';
  }

  // ── Validate ticket is safe to auto-execute ─────────────────────────────────
  // Returns a reason string if NOT ready, or null.
  function _validate(data, ticket) {
    var props  = ticket.properties || {};
    var stage  = props.hs_pipeline_stage || '';
    var reason = (data.reasonOfMigration || '').toLowerCase();

    // CLS-gated reasons must be in CLS_APPROVED stage
    if (_needsCls(reason) && stage !== STAGE.CLS_APPROVED) {
      return 'CLS approval required — move the ticket to "Approved by CLS" first.';
    }

    if (!data.newTeacher)        return 'New teacher is missing on the ticket.';
    if (!data.jlid)              return 'Learner ID (learner_uid) is missing on the ticket.';
    if (!data.reasonOfMigration) return 'Reason of migration is missing on the ticket.';

    return null; // valid
  }

  // Nothing goes to the parent/teacher unless we know the classes can be booked.
  function _preflight(data) {
    if (data.scheduleError) return data.scheduleError;
    if (data.timeError) return data.timeError;
    if (!data.classSessions || !data.classSessions.length) {
      return 'No class schedule found — add the regular class day and time (CET) to the ticket (or deal).';
    }
    var bad = data.classSessions.filter(function(s) { return _DAY_INDEX[s.day] === undefined || !_parse12hTime(s.time); });
    if (bad.length) return 'Class schedule could not be read: ' + JSON.stringify(bad);
    if (!data.zoomLink) return 'No class link found — add it to the deal or the learner\'s existing calendar event.';
    return null;
  }

  // ── Pre-run check ("the brain") ─────────────────────────────────────────────
  // Runs before ANYTHING is sent. hard = cannot run. soft = needs a human OK (the automatic run waits, a person can accept).
  function _hash(s) { var h = 5381; s = String(s); for (var i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0; return String(h >>> 0); }

  function _buildPlan(data) {
    var plan = {
      learner: data.learner, jlid: data.jlid, reason: data.reasonOfMigration,
      oldTeacher: data.oldTeacher, newTeacher: data.newTeacher,
      currentCourse: data.course, bookingCourse: data.bookingCourse, futureCourses: (data.confirmedFutureCourses || []).map(function(c) { return c.courseLabel; }),
      ticketSlotCET: data.ticketSessions || [], parentTimezone: data.parentIana || '', slotToTellParent: data.classSessions || [],
      sendsWhatsappToParent: !!data.sendWhatsappToParent, sendsEmailToTeacher: !!data.sendEmailToTeacher, sendsCertificate: !!data.sendCertificate,
      parentEmails: data.parentEmailTargets || [], migrationType: data.migrationType, clsManager: data.clsManager, jetGuide: data.jetGuide,
      classLink: data.zoomLink || ''
    };
    try {
      if (data.ticketSessions && data.ticketSessions.length && data.parentIana && data.sessionsSourceTz) {
        var at = _nextInstant(data.ticketSessions[0], data.sessionsSourceTz);
        plan.firstClassForParent = _fmtInstant(at, data.parentIana);
        plan.firstClassCET = _fmtInstant(at, 'Europe/Paris');
      }
    } catch (e) {}
    return plan;
  }

  function _aiReview(ticketId, plan) {
    var key = 'MIG_AI_' + ticketId + '_' + _hash(JSON.stringify(plan));
    try { var hit = CacheService.getScriptCache().get(key); if (hit) return JSON.parse(hit); } catch (e) {}
    var res = { available: false, verdict: 'ok', issues: [], summary: '' };
    try {
      var prompt = 'You are the safety check for JetLearn, an online coding school. A script is about to migrate a learner to a new teacher: '
        + 'it will WhatsApp the parent, email the teacher and book weekly classes. Review the plan below BEFORE it runs.\n\n'
        + 'Facts you must apply:\n'
        + '- The ticket stores the class slot in CET (Europe/Paris). The parent must be told that same moment in THEIR timezone ("slotToTellParent"), '
        + 'so for a UK parent a 6 PM CET class is 5 PM, not 6 PM. Check slotToTellParent and firstClassForParent agree with ticketSlotCET and firstClassCET.\n'
        + '- The class link must be a live.jetlearn.com/join, Zoom, Google Meet or Teams link, never an unrelated website.\n'
        + '- Check that the teacher change, reason and courses are consistent (for a course change the booking course is the NEW course), '
        + 'that class times are plausible for a child (roughly 07:00-21:30 in the parent timezone), and that nothing is missing.\n\n'
        + 'Plan (JSON):\n' + JSON.stringify(plan) + '\n\n'
        + 'Answer with JSON only: {"verdict":"ok"|"hold","issues":["short plain sentence"],"summary":"one plain sentence of what will happen"}. '
        + 'Use "hold" ONLY when something is clearly wrong and would send wrong information to a parent or teacher. Minor style points are not issues.';
      var out = JSON.parse(_callGemini([{ role: 'user', parts: [{ text: prompt }] }], null, 1024));
      res = { available: true, verdict: out.verdict === 'hold' ? 'hold' : 'ok', issues: (out.issues || []).slice(0, 4), summary: String(out.summary || '') };
      try { CacheService.getScriptCache().put(key, JSON.stringify(res), 21600); } catch (e) {}
    } catch (e) {
      Logger.log('[MIG_AUTO] AI review unavailable: ' + e.message);
    }
    return res;
  }

  function _guard(data, ticket) {
    var hard = [], soft = [];
    var plan = _buildPlan(data);

    if (!data.parentIana) hard.push('The parent\'s timezone is not set on the deal, so the class time cannot be told to the parent correctly. Set the timezone on the deal.');
    if (data.timeError) hard.push(data.timeError);
    if (!_isClassLink(data.zoomLink)) hard.push('The class link does not look like a class link (' + (data.zoomLink || 'empty') + ').');

    // class hours as the parent will see them
    (data.classSessions || []).forEach(function(s) {
      var tm = _parse12hTime(s.time);
      if (tm && (tm.h < 7 || tm.h >= 22)) soft.push('The class on ' + s.day + ' is at ' + s.time + ' in the parent\'s timezone (' + data.parentIana + ') — an unusual hour.');
    });

    // teacher already busy at that time (their calendar), ignoring this learner's own reserved / booked events
    try {
      var d = new Date(), pad = function(n) { return ('0' + n).slice(-2); };
      var start = d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
      var chk = checkBookingConflicts(data.newTeacher, data.classSessions, start, data.classBookingIana);
      if (chk && chk.success && chk.conflicts && chk.conflicts.length) {
        var real = chk.conflicts.filter(function(c) { return String(c.eventTitle || '').indexOf(data.jlid) === -1; });
        if (real.length) soft.push(data.newTeacher + ' is already booked at that time: ' + real.slice(0, 3).map(function(c) { return c.day + ' ' + c.time + ' — ' + c.eventTitle; }).join('; ') + '.');
      }
    } catch (ce) { Logger.log('[MIG_AUTO] conflict check failed: ' + ce.message); }

    var ai = { available: false, verdict: 'ok', issues: [], summary: '' };
    if (!hard.length) {
      ai = _aiReview(ticket.id, plan);
      if (ai.verdict === 'hold') soft.push('AI check: ' + (ai.issues.length ? ai.issues.join(' ') : 'something in the plan looks wrong.'));
    }
    return { hard: hard, soft: soft, ai: ai, plan: plan };
  }

  // Leave the ticket where it is, with the reason on it. Notes/writes only happen when the reason CHANGES
  // so a blocked ticket re-checked every 20 minutes never spams HubSpot.
  function _markBlocked(ticket, reason) {
    var props = ticket.properties || {};
    Logger.log('[MIG_AUTO] BLOCKED ticket ' + ticket.id + ': ' + reason);
    if (props.mig_auto_attempted === 'blocked' && String(props.mig_auto_error || '') === reason) return;
    _patchTicketProps(ticket.id, { mig_auto_attempted: 'blocked', mig_auto_error: reason });
    _addTicketNote(ticket.id, '⏸ Auto-migration is waiting — ' + reason +
      '\nNothing has been sent to the parent or teacher. It re-checks every ~20 minutes and runs automatically once this is fixed (or use Auto Migrate in the app).');
  }

  // Something broke AFTER we started (messages may already be out) — never auto-retried, stays in Execution Pending.
  function _markFailed(ticketId, reason) {
    Logger.log('[MIG_AUTO] FAILED ticket ' + ticketId + ': ' + reason);
    _patchTicketProps(ticketId, { mig_auto_attempted: 'failed', mig_auto_error: reason });
    _addTicketNote(ticketId, '❌ Auto-migration needs manual attention — ' + reason +
      '\nThe ticket was left in Execution Pending. It will NOT retry by itself (messages may already have gone out); fix it and use Auto Migrate.');
  }

  function _recentlyChecked(ticketId) {
    try { return !!CacheService.getScriptCache().get('MIG_AUTO_CHK_' + ticketId); } catch(e) { return false; }
  }
  function _markChecked(ticketId) {
    try { CacheService.getScriptCache().put('MIG_AUTO_CHK_' + ticketId, '1', RECHECK_SECONDS); } catch(e) {}
  }

  // Books the weekly sessions. Returns { ok, count, note, error }.
  function _bookClasses(data) {
    try {
      var remainingRes = getRemainingClassesForJlid(data.jlid);
      var numEvents = (remainingRes && remainingRes.classesLeft > 0) ? remainingRes.classesLeft : 12;
      var iana = data.classBookingIana || data.manualTimezone || data.timezone || 'Europe/London';
      var today = new Date();
      var pad = function(n) { return ('0' + n).slice(-2); };
      var startDate = today.getFullYear() + '-' + pad(today.getMonth() + 1) + '-' + pad(today.getDate());

      // Reserved slots we marked earlier for this learner: the manual flow deletes them once the real classes are booked
      var reserved = [];
      try {
        var lg = getBookingLogForJlid(data.jlid);
        if (lg && lg.success && lg.rows) {
          reserved = lg.rows.filter(function(r) { return r.status === 'Active' && String(r.eventTitle || '').indexOf('Reserved : ') === 0; });
        }
      } catch (rle) { Logger.log('[MIG_AUTO] reserved-slot lookup failed: ' + rle.message); }

      var bk = bookClassesWithNewTeacher(
        data.jlid, data.learner, data.newTeacher, data.classSessions,
        data.bookingCourse || data.course, startDate, iana, numEvents,
        data.parentEmail ? [data.parentEmail] : [],
        'MigrationAutomation', data.zoomLink, '',
        data.existingEventDesc || '', data.migrationPrefix || 'Migration'
      );
      if (bk && bk.success !== false) {
        var count = bk.occurrences || numEvents;
        var sessionText = data.classSessions.map(function(s) { return s.day + ' ' + s.time; }).join(', ');
        var note = '📅 ' + count + ' classes booked with ' + data.newTeacher + ' (' + sessionText + ', ' + iana + ').';
        reserved.forEach(function(r) {
          var label = (r.teacherName || 'teacher') + (r.sessions ? ' (' + r.sessions + ')' : '');
          try {
            var cr = cancelBookedClasses(r.rowIndex);
            note += (cr && cr.success)
              ? '\n🗑 Removed the earlier reserved slot with ' + label + '.'
              : '\n⚠ Could not remove the reserved slot with ' + label + ' (' + ((cr && cr.message) || 'unknown') + ') - remove it in Manage Bookings.';
          } catch (ce) {
            note += '\n⚠ Could not remove the reserved slot with ' + label + ' (' + ce.message + ') - remove it in Manage Bookings.';
          }
        });
        return { ok: true, count: count, note: note };
      }
      return { ok: false, error: (bk && bk.message) || 'Unknown booking error' };
    } catch(e) {
      return { ok: false, error: e.message };
    }
  }

  // Required ticket fields → then Completed. Existing values (set by people) are never overwritten.
  // Returns { ok, error }.
  function _completeTicket(ticket) {
    var props = ticket.properties || {};
    var intervention = _getTicketIntervention(ticket.id);
    var patch = { mig_auto_error: '' };
    if (!props.migration_slot_status__t_)   patch.migration_slot_status__t_  = (intervention && intervention.slotStatus)   || 'Matched by Ops';
    if (!props.migration_execution_status)  patch.migration_execution_status = (intervention && intervention.intervenedBy) || 'Ops Intervention';

    var pr = _patchTicketProps(ticket.id, patch);
    if (!pr.ok) return { ok: false, error: 'could not update the ticket properties (HTTP ' + pr.code + (pr.body ? ': ' + pr.body : '') + ')' };

    var st = _patchTicketStage(ticket.id, STAGE.COMPLETED);
    if (!st.ok) return { ok: false, error: 'HubSpot would not move the ticket to Completed (HTTP ' + st.code + (st.body ? ': ' + st.body : '') + ') — a required ticket field is probably missing' };
    return { ok: true, slotStatus: patch.migration_slot_status__t_ || props.migration_slot_status__t_,
             intervenedBy: patch.migration_execution_status || props.migration_execution_status };
  }

  // ── Execute one ticket ──────────────────────────────────────────────────────
  function _executeTicket(ticket) {
    var ticketId = ticket.id;
    var props = ticket.properties || {};
    Logger.log('[MIG_AUTO] Processing ticket ' + ticketId);

    // 1. Course-change tickets wait for the old course to finish (cheap, no API calls)
    var wait = _cctcWaitReason(props);
    if (wait) { _markBlocked(ticket, wait); return; }

    // 2. Blocked tickets are only re-checked every ~20 minutes
    var prior = String(props.mig_auto_attempted || '');
    if ((prior === 'blocked' || prior === 'skipped') && _recentlyChecked(ticketId)) return;
    _markChecked(ticketId);

    // 3. Build + validate + preflight — NOTHING is sent to anyone until all of this passes
    var data = _buildExecutionData(ticket);
    if (!data) {
      _markBlocked(ticket, 'Could not load the learner\'s deal from HubSpot (deal not found for this learner ID, or HubSpot unavailable) — will retry.');
      return;
    }
    var problem = _validate(data, ticket) || _preflight(data);
    if (problem) { _markBlocked(ticket, problem); return; }

    // 3b. The pre-run check — hard problems and anything needing a human OK keep it waiting (nothing sent)
    var guard = _guard(data, ticket);
    if (guard.hard.length || guard.soft.length) { _markBlocked(ticket, 'Needs a check before it runs — ' + guard.hard.concat(guard.soft).join(' ')); return; }
    data._aiSummary = guard.ai && guard.ai.available ? guard.ai.summary : '';

    // 4. Go. Mark in-flight first so an overlapping run can never double-send.
    _patchTicketProps(ticketId, { mig_auto_attempted: 'true', mig_auto_error: '' });

    try {
      var result;
      try { result = sendMigrationEmail(data); } catch(se) { result = { success: false, message: se.message }; }
      if (!result || result.success === false) {
        _markFailed(ticketId, 'the migration emails/WhatsApp step failed: ' + ((result && result.message) || 'unknown error') + ' (no classes were booked).');
        return;
      }

      // 5. Classes. "Executed" means booked — if booking fails the ticket stays put.
      var booking = _bookClasses(data);
      if (!booking.ok) {
        _markFailed(ticketId, 'messages were sent but the classes were NOT booked: ' + booking.error + '. Book the classes manually, then move the ticket to Completed.');
        return;
      }

      // 6. Course-change: deal moves to the new course
      _updateCctcFutureCourse(data, ticket);

      // 7. Required ticket fields, then Completed — each write is checked
      var fin = _completeTicket(ticket);
      if (!fin.ok) {
        _markFailed(ticketId, 'messages were sent and classes booked, but ' + fin.error + '.');
        return;
      }

      Logger.log('[MIG_AUTO] ✅ Ticket ' + ticketId + ' executed successfully');
      _addTicketNote(ticketId,
        '✅ Migration auto-executed by MigrationAutomation.\n' + booking.note +
        '\nSlot status: ' + fin.slotStatus + ' | Intervened by: ' + fin.intervenedBy +
        (data._aiSummary ? '\n🤖 Pre-run check: OK — ' + data._aiSummary : ''));
    } catch(e) {
      _markFailed(ticketId, 'unexpected error: ' + e.message);
    }
  }

  // ── Main entry point — called by GAS time trigger ───────────────────────────
  function run() {
    var cache = CacheService.getScriptCache();
    // Only one run at a time (a long run must not be overtaken by the next 10-min trigger).
    // NOT LockService.getScriptLock() — the API monitor takes that lock on every HubSpot call.
    if (cache.get('MIG_AUTO_RUNNING')) { Logger.log('[MIG_AUTO] Previous run still going — skipping this one.'); return; }
    cache.put('MIG_AUTO_RUNNING', String(new Date().getTime()), 330);
    var started = new Date().getTime();
    Logger.log('[MIG_AUTO] === Starting automation run === ' + new Date().toISOString());

    try {
      var tickets = _fetchActionableTickets();
      Logger.log('[MIG_AUTO] Found ' + tickets.length + ' actionable tickets');

      for (var i = 0; i < tickets.length; i++) {
        if (new Date().getTime() - started > TIME_BUDGET_MS) {
          Logger.log('[MIG_AUTO] Time budget reached — ' + (tickets.length - i) + ' ticket(s) left for the next run.');
          break;
        }
        try {
          _executeTicket(tickets[i]);
        } catch(e) {
          Logger.log('[MIG_AUTO] Unhandled error for ticket ' + tickets[i].id + ': ' + e.message);
        }
      }
    } finally {
      try { cache.remove('MIG_AUTO_RUNNING'); } catch(e) {}
    }

    // Course-change forms filed less than 4 weeks before a kit course: tell the teacher (once per ticket)
    try { if (typeof scanLateMigrationForms === 'function') scanLateMigrationForms(); } catch (lf) { Logger.log('[MIG_AUTO] late-form scan failed: ' + lf.message); }
    Logger.log('[MIG_AUTO] === Run complete ==='); 
  }

  // ── CCTC: set deal current_course = new course after execution ───────────
  function _updateCctcFutureCourse(data, ticket) {
    try {
      var reason = (data.reasonOfMigration || '').toLowerCase();
      var isCctc = reason.indexOf('course change') !== -1;
      if (!isCctc) return;
      var props = ticket.properties || {};
      var fc1 = props.future_course_1 || ''; // new course they're moving TO
      if (!fc1) return;

      // Update DEAL: current_course = new course
      var dealId = data._dealId || '';
      if (dealId) {
        var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/deals/' + dealId, {
          method: 'PATCH',
          headers: { 'Authorization': 'Bearer ' + _token(), 'Content-Type': 'application/json' },
          payload: JSON.stringify({ properties: { current_course: fc1 } }),
          muteHttpExceptions: true
        });
        Logger.log('[MIG_AUTO] CCTC: deal ' + dealId + ' current_course → ' + fc1 + ' HTTP ' + (resp ? resp.getResponseCode() : '?'));
      }
    } catch(e) {
      Logger.log('[MIG_AUTO] _updateCctcFutureCourse error: ' + e.message);
    }
  }

  // ── Single-ticket manual trigger (from UI) ───────────────────────────────
  // Everything a run needs, WITHOUT sending anything. { error } or { ticket, data }.
  function _prepare(jlid, overrideNewTeacher) {
    if (!jlid) return { error: 'JLID required.' };
    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/tickets/search', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + _token(), 'Content-Type': 'application/json' },
      payload: JSON.stringify({
        filterGroups: [{ filters: [
          { propertyName: 'hs_pipeline',      operator: 'EQ', value: PIPELINE_ID },
          { propertyName: 'learner_uid',       operator: 'EQ', value: jlid },
          { propertyName: 'hs_pipeline_stage', operator: 'IN', values: [STAGE.EXEC_PENDING, STAGE.CLS_APPROVED] }
        ] }],
        properties: TICKET_PROPERTIES, sorts: [{ propertyName: 'createdate', direction: 'DESCENDING' }], limit: 1
      }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return { error: 'HubSpot search failed.' };
    var results = JSON.parse(resp.getContentText()).results || [];
    if (!results.length) return { error: 'No ticket in Execution Pending or CLS Approved stage for ' + jlid + '. Move the ticket to that stage first.' };
    var ticket = results[0];
    var wait = _cctcWaitReason(ticket.properties || {});
    if (wait) return { error: wait };
    var data = _buildExecutionData(ticket);
    if (!data) return { error: 'Could not load the learner\'s deal for ' + jlid + '.' };
    if (overrideNewTeacher && overrideNewTeacher.trim()) data.newTeacher = overrideNewTeacher.trim();
    var problem = _validate(data, ticket) || _preflight(data);
    if (problem) return { error: problem };
    return { ticket: ticket, data: data };
  }

  // "Check before I run" — what will happen, and anything that needs attention. Sends nothing, changes nothing.
  function verifyForJlid(jlid, overrideNewTeacher) {
    try {
      var p = _prepare(jlid, overrideNewTeacher);
      if (p.error) return { success: false, message: p.error };
      var g = _guard(p.data, p.ticket);
      return { success: true, learner: p.data.learner, plan: g.plan, hard: g.hard, soft: g.soft, aiSummary: g.ai && g.ai.available ? g.ai.summary : '', aiAvailable: !!(g.ai && g.ai.available) };
    } catch (e) {
      return { success: false, message: 'Could not check: ' + e.message };
    }
  }

  // Wrapper: records live progress (read by the page through getMigrationProgress) around the real run
  function runForJlid(jlid, overrideNewTeacher, acceptWarnings) {
    var key = 'migprog_' + String(jlid || '').trim();
    if (jlid) { __PROG_KEY = key; _progWrite(key, function(o) { o.steps = []; o.done = false; o.result = ''; }); _progBegin('Checking the ticket, class link and schedule'); }
    var res;
    try { res = _runForJlidCore(jlid, overrideNewTeacher, acceptWarnings); }
    catch (e) { _progEnd('Checking the ticket, class link and schedule', false, e.message); res = { success: false, message: e.message }; throw e; }
    finally { if (jlid) { _progWrite(key, function(o) { o.done = true; o.result = (res && res.message) || ''; }); __PROG_KEY = ''; } }
    return res;
  }

  function _runForJlidCore(jlid, overrideNewTeacher, acceptWarnings) {
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

    // Same gates as the automatic run — nothing is sent unless the classes can be booked
    var wait = _cctcWaitReason(props);
    if (wait) { _markBlocked(ticket, wait); return { success: false, message: 'Not run — ' + wait }; }

    var data = _buildExecutionData(ticket);
    if (!data) return { success: false, message: 'Could not build execution data — deal/ticket mismatch for ' + jlid + '.' };

    // UI-selected teacher overrides whatever was fetched from ticket/deal
    if (overrideNewTeacher && overrideNewTeacher.trim()) {
      Logger.log('[MIG_AUTO] UI teacher override: ' + data.newTeacher + ' → ' + overrideNewTeacher.trim());
      data.newTeacher = overrideNewTeacher.trim();
    }

    var problem = _validate(data, ticket) || _preflight(data);
    if (problem) {
      _markBlocked(ticket, problem);
      return { success: false, message: 'Not run — nothing was sent. ' + problem };
    }

    // The pre-run check. Hard problems always stop it; softer ones need the person's OK (acceptWarnings).
    var guard = _guard(data, ticket);
    if (guard.hard.length) {
      _markBlocked(ticket, 'Needs a check before it runs — ' + guard.hard.join(' '));
      return { success: false, message: 'Not run — nothing was sent. ' + guard.hard.join(' ') };
    }
    if (guard.soft.length && !acceptWarnings) {
      return { success: false, needsConfirm: true, warnings: guard.soft, message: 'Not run — needs your OK first. ' + guard.soft.join(' ') };
    }

    _progEnd('Checking the ticket, class link and schedule', true);
    // Allow re-run from UI even if previously attempted
    _patchTicketProps(ticketId, { mig_auto_attempted: 'true', mig_auto_error: '' });

    var result;
    try {
      result = sendMigrationEmail(data);
    } catch(e) {
      _markFailed(ticketId, 'the migration emails/WhatsApp step threw an error: ' + e.message + ' (no classes were booked).');
      return { success: false, message: e.message };
    }
    if (!result || result.success === false) {
      var errMsg = (result && result.message) ? result.message : 'Migration email failed';
      _markFailed(ticketId, 'the migration emails/WhatsApp step failed: ' + errMsg + ' (no classes were booked).');
      return { success: false, message: errMsg };
    }

    var timeline = (result && result.timeline) ? result.timeline.slice() : [];

    _progBegin('Booking the classes on the calendar');
    var booking = _bookClasses(data);
    _progEnd('Booking the classes on the calendar', !!booking.ok, booking.ok ? booking.note : booking.error);
    if (!booking.ok) {
      _markFailed(ticketId, 'messages were sent but the classes were NOT booked: ' + booking.error + '. Book the classes manually, then move the ticket to Completed.');
      timeline.push({ key: 'class_booking', label: 'Class booking failed: ' + booking.error, status: 'failed', durationMs: 0 });
      return { success: false, message: 'Messages were sent, but classes were NOT booked: ' + booking.error + '. Ticket left in Execution Pending.', timeline: timeline };
    }

    _updateCctcFutureCourse(data, ticket);

    _progBegin('Moving the HubSpot ticket to Completed');
    var fin = _completeTicket(ticket);
    _progEnd('Moving the HubSpot ticket to Completed', !!fin.ok, fin.ok ? '' : fin.error);
    timeline.push({ key: 'class_booking', label: booking.note, status: 'success', durationMs: 0 });
    if (!fin.ok) {
      _markFailed(ticketId, 'messages were sent and classes booked, but ' + fin.error + '.');
      timeline.push({ key: 'ticket_complete', label: 'Could not complete the ticket: ' + fin.error, status: 'failed', durationMs: 0 });
      return { success: false, message: 'Classes booked, but ' + fin.error + '. Ticket left in Execution Pending.', timeline: timeline };
    }

    _addTicketNote(ticketId, [
      '✅ Auto Migration Executed',
      'JLID: ' + data.jlid,
      'New Teacher: ' + data.newTeacher,
      'Course: ' + (data.bookingCourse || data.course || '—'),
      'Sessions: ' + JSON.stringify(data.classSessions || []),
      'Timezone: ' + (data.classBookingIana || '—'),
      booking.note,
      'Slot Status: ' + fin.slotStatus,
      'Intervened By: ' + fin.intervenedBy
    ].join('\n'));
    timeline.push({ key: 'ticket_complete', label: 'HubSpot ticket moved to Completed.', status: 'success', durationMs: 0 });

    return { success: true, message: '✅ Emails sent. ' + booking.note, timeline: timeline };
  }

  // ── "Ready to migrate" list for the Learner Migration screen ────────────────
  // Every ticket still waiting to be executed (Execution Pending / CLS Approved), with what — if anything —
  // is stopping it. Judged from the ticket alone (no deal/calendar calls) so the list loads fast; the class link
  // and the deal schedule are checked when a migration is actually run (nothing is sent if they are missing).
  function listReady() {
    var props = TICKET_PROPERTIES.concat(['createdate', 'hs_lastmodifieddate', 'hs_date_entered_1065336836', 'hs_date_entered_133755411']);
    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/tickets/search', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + _token(), 'Content-Type': 'application/json' },
      payload: JSON.stringify({
        filterGroups: [{ filters: [
          { propertyName: 'hs_pipeline',       operator: 'EQ', value: PIPELINE_ID },
          { propertyName: 'hs_pipeline_stage', operator: 'IN', values: [STAGE.EXEC_PENDING, STAGE.CLS_APPROVED] }
        ] }],
        properties: props,
        sorts: [{ propertyName: 'createdate', direction: 'ASCENDING' }],
        limit: 100
      }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      return { success: false, message: 'HubSpot search failed (HTTP ' + resp.getResponseCode() + ').' };
    }
    var body = JSON.parse(resp.getContentText());
    var now = new Date().getTime();

    var rows = (body.results || []).map(function(t) {
      var p = t.properties || {};
      var reason = p.reason_of_migration__t_ || '';
      var stage = p.hs_pipeline_stage || '';
      var attempted = String(p.mig_auto_attempted || '');
      var learner = String(p.subject || '').replace(/\s*-?\s*Migration Request Raised\s*$/i, '').trim() || p.learner_uid || '';
      var newTeacher = p.new_teacher ? (getTeacherLabel(p.new_teacher) || p.new_teacher) : '';
      var oldTeacher = p.current_teacher__t_ ? (getTeacherLabel(p.current_teacher__t_) || p.current_teacher__t_) : '';
      var course = p.current_course__t_ ? (getCourseLabel(p.current_course__t_) || p.current_course__t_) : '';
      var nextCourse = _isNa(p.future_course_1) ? '' : (getCourseLabel(p.future_course_1) || p.future_course_1);

      // Class schedule as written on the ticket (may hold several days/times)
      var scheduleText = '', schedErr = '';
      if (p.regular_class_day__t_ && p.regular_class_time__in_cet_) {
        var built = _sessionsFromTicketSchedule(p.regular_class_day__t_, p.regular_class_time__in_cet_);
        if (built.sessions.length) scheduleText = built.sessions.map(function(s) { return s.day.slice(0, 3) + ' ' + s.time; }).join(', ');
        else schedErr = built.error;
      }

      // Why not ready? (first matching reason wins)
      var state = 'ready', why = '', canRun = true, canBulk = true;
      var wait = _cctcWaitReason(p);
      if (attempted === 'failed') {
        state = 'failed'; why = p.mig_auto_error || 'A previous run failed.'; canBulk = false;
      } else if (attempted === 'true') {
        state = 'started'; why = 'A run started but did not finish — check the ticket before running again.'; canRun = false; canBulk = false;
      } else if (!newTeacher) {
        state = 'waiting'; why = 'New teacher is not set on the ticket.'; canRun = false; canBulk = false;
      } else if (wait) {
        state = 'waiting'; why = wait; canRun = false; canBulk = false;
      } else if (_needsCls(reason) && stage !== STAGE.CLS_APPROVED) {
        state = 'waiting'; why = 'Needs CLS approval first — move the ticket to "Approved by CLS".'; canRun = false; canBulk = false;
      } else if (schedErr) {
        state = 'waiting'; why = schedErr; canRun = false; canBulk = false;
      } else if (attempted === 'blocked' || attempted === 'skipped') {
        state = 'blocked'; why = p.mig_auto_error || 'Waiting on something.';
      }

      var enteredRaw = stage === STAGE.CLS_APPROVED ? p.hs_date_entered_133755411 : p.hs_date_entered_1065336836;
      var sinceMs = new Date(enteredRaw || p.createdate || now).getTime();
      var days = isNaN(sinceMs) ? 0 : Math.max(0, Math.floor((now - sinceMs) / 86400000));

      return {
        ticketId: t.id, jlid: p.learner_uid || '', learner: learner,
        oldTeacher: oldTeacher, newTeacher: newTeacher, reason: reason,
        course: course, nextCourse: nextCourse, schedule: scheduleText,
        stage: stage, stageLabel: stage === STAGE.CLS_APPROVED ? 'Approved by CLS' : 'Execution Pending',
        daysWaiting: days, state: state, why: why, canRun: canRun && !!p.learner_uid, canBulk: canBulk && !!p.learner_uid
      };
    });

    return { success: true, rows: rows, total: body.total || rows.length, truncated: (body.total || 0) > rows.length, fetchedAt: new Date().toISOString() };
  }

  return { run: run, runForJlid: runForJlid, verifyForJlid: verifyForJlid, listReady: listReady };

})();

// ── Learner Migration screen: tickets waiting to be migrated ─────────────────
function getReadyToMigrate() {
  try {
    return MIG_AUTO.listReady();
  } catch(e) {
    Logger.log('[MIG_AUTO] getReadyToMigrate error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── GAS time trigger entry point ─────────────────────────────────────────────
function runMigrationAutomation() {
  MIG_AUTO.run();
}

// ── Manual single-ticket auto migration — called from migration page UI ───────
function runAutoMigrateForJlid(jlid, overrideNewTeacher, acceptWarnings) {
  return MIG_AUTO.runForJlid(jlid, overrideNewTeacher, acceptWarnings);
}

// Pre-run check for the UI: what will happen + anything that needs attention. Sends nothing.
function verifyAutoMigrate(jlid, overrideNewTeacher) {
  return MIG_AUTO.verifyForJlid(jlid, overrideNewTeacher);
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
  // Failed/blocked/skipped only matter while the ticket is still waiting to be executed —
  // once someone finishes it by hand (stage = Completed) it is no longer an open problem.
  var OPEN_STAGE = { propertyName: 'hs_pipeline_stage', operator: 'IN', values: ['1065336836', '133755411'] };

  var PROPS = ['subject', 'learner_uid', 'hs_pipeline_stage', 'mig_auto_attempted', 'mig_auto_error',
               'createdate', 'hs_lastmodifieddate', 'reason_of_migration__t_', 'new_teacher'];

  // Auto-executed: mig_auto_attempted = 'true' and stage = Completed
  var autoExec = _search(BASE.concat([
    { propertyName: 'mig_auto_attempted', operator: 'EQ', value: 'true' },
    { propertyName: 'hs_pipeline_stage',  operator: 'EQ', value: '128913753' }
  ]), PROPS, 100);

  // Failed: ran, but something broke afterwards — needs a person
  var failed = _search(BASE.concat([
    OPEN_STAGE,
    { propertyName: 'mig_auto_attempted', operator: 'EQ', value: 'failed' }
  ]), PROPS, 100);

  // Blocked: waiting on something (class link, schedule, CLS approval, course end date…) — re-checked automatically
  var blocked = _search(BASE.concat([
    OPEN_STAGE,
    { propertyName: 'mig_auto_attempted', operator: 'EQ', value: 'blocked' }
  ]), PROPS, 100);

  // Skipped: legacy value from before 'blocked' existed
  var skipped = _search(BASE.concat([
    OPEN_STAGE,
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
    blocked:        blocked.length,
    skipped:        skipped.length,
    pendingQueue:   pending.length,
    needManual:     failed.length + blocked.length + skipped.length,
    failedTickets:  failed.map(_mapTicket),
    blockedTickets: blocked.map(_mapTicket),
    skippedTickets: skipped.map(_mapTicket),
    pendingTickets: pending.map(_mapTicket),
    fetchedAt:      new Date().toISOString()
  };
}
