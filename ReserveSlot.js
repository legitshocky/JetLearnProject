// ─────────────────────────────────────────────────────────────────────────────
// reserveCalendarSlot
// Books the first requested slot (date + "HH:00 AM/PM - HH:00 AM/PM" in CET) as
// a 12-week weekly recurring event on the class-schedule calendar and the
// teacher's own calendar.
// Called from Teacher Persona "📅 Reserve Slot" button.
// Returns { success, message, title, start, end, occurrences, eventSeriesId }
// ─────────────────────────────────────────────────────────────────────────────
function _courseTypeLabel(courseName, jlid) {
  // Check course name first — GCSE and other specific courses override JLID suffix
  var c = String(courseName || '').toLowerCase();
  if (c.indexOf('gcse') > -1) return 'GCSE';
  if (c.indexOf('financial') > -1 || c.indexOf('finlit') > -1) return 'Financial Literacy';
  if (c.indexOf('math') > -1) return 'Fun with Maths';

  // Fall back to JLID suffix
  var j = String(jlid || '').toUpperCase();
  var m = j.match(/(FL|[CM])\d*$/);
  var suf = m ? m[1] : '';
  if (suf === 'FL') return 'Financial Literacy';
  if (suf === 'M') return 'Fun with Maths';
  if (suf === 'C') return 'AI-Coding';

  if (c.indexOf('coding') > -1 || c.indexOf('code') > -1 || c.indexOf('ai') > -1) return 'AI-Coding';
  return courseName || 'Lesson';
}

// ─────────────────────────────────────────────────────────────────────────────
// Looks up a teacher's calendar ID, jetlearn email, and teacher ID by name.
// Reads "Migration Teacher" tab (col A=name, col B=calendarId) and "Teacher Data"
// (col B=name, col A=teacherId, col I=email). Returns {calendarId, email, teacherId}.
// ─────────────────────────────────────────────────────────────────────────────
function _lookupTeacherCalendarInfo(teacherName) {
  var info = { calendarId: '', email: '', teacherId: '' };
  var tNorm = normalizeTeacherName(teacherName);
  try {
    var migSS    = SpreadsheetApp.openById(CONFIG.AUDIT_SHEET_ID);
    var migSheet = migSS.getSheetByName('Migration Teacher');
    if (migSheet) {
      var migRows = migSheet.getDataRange().getValues();
      for (var i = 0; i < migRows.length; i++) {
        var n   = String(migRows[i][0] || '').trim();
        var cid = String(migRows[i][1] || '').trim();
        if (n && normalizeTeacherName(n) === tNorm && cid.indexOf('@') > -1) {
          info.calendarId = cid;
          break;
        }
      }
    }
  } catch(e) { Logger.log('[_lookupTeacherCalendarInfo] Migration Teacher lookup error: ' + e.message); }

  try {
    var tdRows = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA);
    for (var j = 1; j < tdRows.length; j++) {
      var n2 = String(tdRows[j][1] || '').trim();
      if (n2 && normalizeTeacherName(n2) === tNorm) {
        info.email      = String(tdRows[j][8] || '').trim().toLowerCase();
        info.teacherId  = String(tdRows[j][0] || '').trim();
        break;
      }
    }
  } catch(e) { Logger.log('[_lookupTeacherCalendarInfo] Teacher Data lookup error: ' + e.message); }

  if (!info.calendarId) info.calendarId = info.email;
  if (!info.email && info.calendarId && info.calendarId.indexOf('@') > -1) info.email = info.calendarId;
  return info;
}

// Parses a GMT-offset timezone label into a fixed offset in minutes (no DST).
// Handles both "(GMT+05:30) ..." and "(GMT +5:30) ..." formats.
function _gmtLabelToOffsetMinutes(gmtLabel) {
  var m = String(gmtLabel || '').match(/GMT\s*([+-])\s*(\d{1,2}):(\d{2})/);
  if (!m) return 60; // default CET
  var sign = (m[1] === '-') ? -1 : 1;
  var hh = parseInt(m[2], 10);
  var mm = parseInt(m[3], 10);
  return sign * (hh * 60 + mm);
}

var _DAY_INDEX = { Sunday: 0, Monday: 1, Tuesday: 2, Wednesday: 3, Thursday: 4, Friday: 5, Saturday: 6 };

// Normalizes "(GMT +5:30) ..." / "(GMT+05:30) ..." into "(GMT+05:30) ..." so labels
// from the migration form can be matched against TIMEZONE_IANA_MAP keys.
function _normalizeGmtLabel(label) {
  return String(label || '').replace(/GMT\s*([+-])\s*(\d{1,2}):(\d{2})/, function(_, sign, hh, mm) {
    return 'GMT' + sign + (hh.length < 2 ? '0' + hh : hh) + ':' + mm;
  }).trim();
}

// Resolves a GMT-offset label to an IANA timezone (for DST-aware offset lookup),
// matching loosely on the normalized "(GMT+HH:MM) ..." prefix + city list.
function _gmtLabelToIana(label) {
  var norm = _normalizeGmtLabel(label);
  for (var key in TIMEZONE_IANA_MAP) {
    if (_normalizeGmtLabel(key) === norm) return TIMEZONE_IANA_MAP[key];
  }
  return null;
}

// Returns the UTC offset (in minutes) for `gmtLabel` on `dateStr`, DST-aware if a
// matching IANA timezone is known; otherwise falls back to the fixed GMT offset
// parsed from the label.
function _resolveOffsetMinutes(dateStr, gmtLabel) {
  var iana = _gmtLabelToIana(gmtLabel);
  if (iana) {
    try {
      return Math.round(_tzOffsetHours(dateStr, iana) * 60);
    } catch(e) {}
  }
  return _gmtLabelToOffsetMinutes(gmtLabel);
}

// ─────────────────────────────────────────────────────────────────────────────
// checkBookingConflicts
// Pre-booking guard: for each session's FIRST occurrence, checks the teacher's
// personal calendar and the master class calendar for overlapping events.
// classSessions: [{ day: 'Monday', time: '5:00 PM' }, ...]; iana e.g. "Asia/Kolkata".
// Returns { success, conflicts: [{day, time, eventTitle, calendar}], unverifiable }
// ─────────────────────────────────────────────────────────────────────────────
function checkBookingConflicts(teacherName, classSessions, startDate, iana) {
  try {
    if (!teacherName || !classSessions || !classSessions.length || !startDate) {
      return { success: false, message: 'Teacher, sessions, and start date required.' };
    }
    iana = iana || 'Europe/London';
    var info = _lookupTeacherCalendarInfo(teacherName);
    var teacherEmailLower = (info.email || '').toLowerCase();
    var calIdLower = (info.calendarId || '').toLowerCase();

    var sdp = String(startDate).split('-');
    var startBase = new Date(Date.UTC(parseInt(sdp[0],10), parseInt(sdp[1],10)-1, parseInt(sdp[2],10)));
    var nameParts = String(teacherName).trim().toLowerCase().split(/\s+/).filter(function(p){ return p.length >= 3; });

    var conflicts = [];
    var unverifiable = !info.calendarId;

    classSessions.forEach(function(sess) {
      var dayIdx = _DAY_INDEX[sess.day];
      var t = _parse12hTime(sess.time);
      if (dayIdx === undefined || !t) return;

      var diff = (dayIdx - startBase.getUTCDay() + 7) % 7;
      var occDate = new Date(startBase.getTime() + diff * 86400000);
      var oy = occDate.getUTCFullYear(), om = occDate.getUTCMonth() + 1, od = occDate.getUTCDate();
      var pad = function(n) { return ('0' + n).slice(-2); };
      var dateStr = oy + '-' + pad(om) + '-' + pad(od);

      var offMins;
      try { offMins = Math.round(_tzOffsetHours(dateStr, iana) * 60); }
      catch(e) { offMins = 0; }
      var slotStartMs = Date.UTC(oy, om-1, od, t.h, t.m) - offMins * 60000;
      var slotEndMs   = slotStartMs + 3600000;
      var tMin = new Date(slotStartMs).toISOString();
      var tMax = new Date(slotEndMs).toISOString();

      // Teacher's personal calendar — any busy event is a conflict
      if (info.calendarId) {
        try {
          var persList = Calendar.Events.list(info.calendarId, {
            timeMin: tMin, timeMax: tMax, singleEvents: true, maxResults: 20
          });
          (persList.items || []).forEach(function(ev) {
            if (ev.status === 'cancelled' || ev.transparency === 'transparent') return;
            conflicts.push({ day: sess.day, time: sess.time, eventTitle: ev.summary || 'Busy', calendar: 'teacher' });
          });
        } catch(pe) {
          Logger.log('[checkBookingConflicts] personal cal error: ' + pe.message);
          unverifiable = true;
        }
      }

      // Master class calendar — conflict only if the event involves this teacher
      try {
        var masterList = Calendar.Events.list(CONFIG.CLASS_SCHEDULE_CALENDAR_ID, {
          timeMin: tMin, timeMax: tMax, singleEvents: true, maxResults: 50
        });
        (masterList.items || []).forEach(function(ev) {
          if (ev.status === 'cancelled') return;
          var guests = (ev.attendees || []).map(function(a){ return (a.email || '').toLowerCase(); });
          var guestMatch = (calIdLower && guests.indexOf(calIdLower) > -1)
                        || (teacherEmailLower && guests.indexOf(teacherEmailLower) > -1);
          var titleLow = (ev.summary || '').toLowerCase();
          var nameMatch = nameParts.length > 0 && nameParts.every(function(p){ return titleLow.indexOf(p) > -1; });
          if (guestMatch || nameMatch) {
            conflicts.push({ day: sess.day, time: sess.time, eventTitle: ev.summary || 'Class', calendar: 'master' });
          }
        });
      } catch(me) {
        Logger.log('[checkBookingConflicts] master cal error: ' + me.message);
      }
    });

    // De-duplicate (same event may appear on both calendars)
    var seen = {};
    conflicts = conflicts.filter(function(c) {
      var k = c.day + '|' + c.time + '|' + c.eventTitle;
      if (seen[k]) return false;
      seen[k] = true; return true;
    });

    return { success: true, conflicts: conflicts, unverifiable: unverifiable };
  } catch(e) {
    Logger.log('[checkBookingConflicts] Error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// Parses "HH:MM AM/PM" (12-hour) into {h, m} 24-hour.
function _parse12hTime(timeStr) {
  var m = String(timeStr || '').match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (!m) return null;
  var h = parseInt(m[1], 10), mn = parseInt(m[2], 10), ap = m[3].toUpperCase();
  if (ap === 'PM' && h !== 12) h += 12;
  if (ap === 'AM' && h === 12) h = 0;
  return { h: h, m: mn };
}

// ─────────────────────────────────────────────────────────────────────────────
// bookClassesWithNewTeacher
// Books each weekly class session (from the Migration form's "Class Schedule"
// rows) as a recurring weekly event with the new teacher, starting from the
// next occurrence of that weekday on/after startDate, in the given GMT-offset
// timezone (e.g. "(GMT+01:00) ..."). Invites hello@jet-learn.com (organizer),
// the new teacher, and any extra parent emails.
// classSessions: [{ day: 'Monday', time: '5:00 PM' }, ...]
// Returns { success, message, booked: [{day, time, start, end}], occurrences }
// ─────────────────────────────────────────────────────────────────────────────
var _JETGUIDE_EMAILS = {
  'Abhishek Nayak':  'abhishek.nayak@jet-learn.com',
  'Anamika Parmar':  'anamika.parmar@jet-learn.com',
  'Sana Rais':       'sana.rais@jet-learn.com',
  'Satyam Mehra':    'satyam.mehra@jet-learn.com'
};

function _logBooking(jlid, learnerName, teacherName, courseName, booked, numEvents, startDate, gmtTimezoneLabel, performedBy, classLink, title) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.AUDIT_SHEET_ID);
    var sheet = ss.getSheetByName('Class Booking Log');
    if (!sheet) {
      sheet = ss.insertSheet('Class Booking Log');
      sheet.appendRow(['Timestamp', 'JLID', 'Learner', 'Teacher', 'Course', 'Sessions', 'Weeks', 'Start Date', 'Timezone', 'Performed By', 'Class Link', 'Event Title']);
      sheet.getRange(1, 1, 1, 12).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    var sessionStr = booked.map(function(b) { return b.day + ' ' + b.time; }).join(', ');
    sheet.appendRow([
      new Date(),
      jlid || '',
      learnerName || '',
      teacherName || '',
      courseName || '',
      sessionStr,
      numEvents || '',
      startDate || '',
      gmtTimezoneLabel || '',
      performedBy || '',
      classLink || '',
      title || ''
    ]);
  } catch(e) {
    Logger.log('[_logBooking] Failed to write booking log: ' + e.message);
  }
}

function bookClassesWithNewTeacher(jlid, learnerName, teacherName, classSessions, courseName, startDate, gmtTimezoneLabel, numEvents, extraEmails, performedBy, classLink, jetGuideName, eventDescription) {
  try {
    if (!teacherName) return { success: false, message: 'New teacher required.' };
    if (!classSessions || !classSessions.length) return { success: false, message: 'No class sessions to book.' };
    if (!startDate) return { success: false, message: 'Start date required.' };

    var info = _lookupTeacherCalendarInfo(teacherName);
    var masterCal = CalendarApp.getCalendarById(CONFIG.CLASS_SCHEDULE_CALENDAR_ID);
    if (!masterCal) return { success: false, message: 'Class schedule calendar not accessible.' };

    var NUM_EVENTS = (numEvents > 0) ? numEvents : 12;
    var courseType = _courseTypeLabel(courseName, jlid);
    var title = (learnerName || 'Learner') + ' (' + (jlid || 'N/A') + ') : Jetlearn ' + courseType + ' Lesson'
      + (info.teacherId ? (' (' + info.teacherId + ')') : '');

    var eventOptions = {};
    if (classLink) {
      eventOptions.location = classLink;
      eventOptions.description = 'Join Zoom Meeting : ' + classLink;
    }

    var guests = [];
    if (info.email) guests.push(info.email);
    (extraEmails || []).forEach(function(e) { if (e && guests.indexOf(e) === -1) guests.push(e); });
    var jetGuideEmail = jetGuideName ? _JETGUIDE_EMAILS[jetGuideName] : null;
    if (jetGuideEmail && guests.indexOf(jetGuideEmail) === -1) guests.push(jetGuideEmail);

    // gmtTimezoneLabel is now an IANA id (e.g. "Asia/Kolkata") passed directly from the booking picker
    var iana = gmtTimezoneLabel || 'Europe/London';

    var sdp = String(startDate).split('-');
    var sYr = parseInt(sdp[0], 10), sMo = parseInt(sdp[1], 10), sDy = parseInt(sdp[2], 10);
    // Find next occurrence using UTC date arithmetic
    var startBase = new Date(Date.UTC(sYr, sMo - 1, sDy));

    var attendees = guests.map(function(g) { return { email: g }; });

    var booked = [];
    var _sessIndex = 0;
    classSessions.forEach(function(sess) {
      var dayIdx = _DAY_INDEX[sess.day];
      var t = _parse12hTime(sess.time);
      if (dayIdx === undefined || !t) return;

      var diff = (dayIdx - startBase.getUTCDay() + 7) % 7;
      var occDate = new Date(startBase.getTime() + diff * 24 * 60 * 60 * 1000);
      var oy = occDate.getUTCFullYear(), om = occDate.getUTCMonth() + 1, od = occDate.getUTCDate();

      // Local datetime strings in the parent's timezone (no offset conversion — Calendar API handles it)
      var pad = function(n) { return ('0' + n).slice(-2); };
      var dateStr  = oy + '-' + pad(om) + '-' + pad(od);
      var startLocal = dateStr + 'T' + pad(t.h) + ':' + pad(t.m) + ':00';
      var endH = t.h + 1, endD = od;
      if (endH >= 24) { endH -= 24; endD += 1; }
      var endDateStr = oy + '-' + pad(om) + '-' + pad(endD);
      var endLocal   = endDateStr + 'T' + pad(endH) + ':' + pad(t.m) + ':00';

      var eventBody = {
        summary: title,
        start:   { dateTime: startLocal, timeZone: iana },
        end:     { dateTime: endLocal,   timeZone: iana },
        recurrence: ['RRULE:FREQ=WEEKLY;COUNT=' + NUM_EVENTS],
        attendees: attendees,
        guestsCanSeeOtherGuests: false,
        guestsCanInviteOthers: false,
        responseRequested: true,
        reminders: {
          useDefault: false,
          overrides: [
            { method: 'popup', minutes: 10 },
            { method: 'email',  minutes: 300 }  // 5 hours before
          ]
        },
        sendUpdates: 'all'
      };
      var descParts = [];
      if (classLink) descParts.push('Join Zoom Meeting : ' + classLink);
      if (eventDescription) descParts.push(eventDescription);
      if (descParts.length) eventBody.description = descParts.join('\n\n');
      if (classLink) eventBody.location = classLink;

      var created = Calendar.Events.insert(eventBody, CONFIG.CLASS_SCHEDULE_CALENDAR_ID);
      if (created && created.id) {
        _hideGuestList(CONFIG.CLASS_SCHEDULE_CALENDAR_ID, created.id);
        // Patch first occurrence of first session only with "Migration : <title>"
        if (_sessIndex === 0) {
          try {
            var instances = Calendar.Events.instances(CONFIG.CLASS_SCHEDULE_CALENDAR_ID, created.id, { maxResults: 1 });
            if (instances && instances.items && instances.items.length) {
              var firstInst = instances.items[0];
              Calendar.Events.patch({ summary: 'Migration : ' + title }, CONFIG.CLASS_SCHEDULE_CALENDAR_ID, firstInst.id);
            }
          } catch(me) {
            Logger.log('[bookClassesWithNewTeacher] Migration tag patch failed: ' + me.message);
          }
        }
      }
      _sessIndex++;

      if (info.calendarId && info.calendarId !== CONFIG.CLASS_SCHEDULE_CALENDAR_ID) {
        try {
          var teacherBody = {
            summary: title,
            start:   { dateTime: startLocal, timeZone: iana },
            end:     { dateTime: endLocal,   timeZone: iana },
            recurrence: ['RRULE:FREQ=WEEKLY;COUNT=' + NUM_EVENTS],
            sendUpdates: 'none'
          };
          if (descParts.length) teacherBody.description = descParts.join('\n\n');
          if (classLink) teacherBody.location = classLink;
          Calendar.Events.insert(teacherBody, info.calendarId);
        } catch(te) {
          Logger.log('[bookClassesWithNewTeacher] Could not write to teacher calendar: ' + te.message);
        }
      }

      booked.push({ day: sess.day, time: sess.time, start: startLocal, end: endLocal, timezone: iana });
    });

    if (!booked.length) return { success: false, message: 'No valid sessions could be parsed.' };

    Logger.log('[bookClassesWithNewTeacher] Booked ' + booked.length + ' session(s) x' + NUM_EVENTS + ' weeks for ' + (learnerName || jlid) + ' with ' + teacherName);
    _logBooking(jlid, learnerName, teacherName, courseName, booked, NUM_EVENTS, startDate, gmtTimezoneLabel, performedBy, classLink, title);

    return { success: true, title: title, booked: booked, occurrences: NUM_EVENTS };
  } catch(e) {
    Logger.log('[bookClassesWithNewTeacher] Error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// Hides the guest list from teacher/parent invitees (privacy) by patching
// guestsCanSeeOtherGuests=false on the recurring event series via Advanced Calendar API.
// CalendarApp's series id is "<eventId>" or "<eventId>_<recurrenceId>"; strip suffix for the master event.
function _hideGuestList(calendarId, seriesId) {
  try {
    var eventId = String(seriesId || '').split('_')[0];
    if (!eventId) return;
    Calendar.Events.patch({ guestsCanSeeOtherGuests: false, guestsCanInviteOthers: false }, calendarId, eventId);
  } catch(e) {
    Logger.log('[_hideGuestList] ' + calendarId + ' ' + seriesId + ': ' + e.message);
  }
}

function reserveCalendarSlot(jlid, learnerName, teacherName, teacherCalendarId, teacherEmail, requestedSlots, courseName, performedBy, teacherId, timeZone, numEvents, extraEmails) {
  try {
    if (!teacherName) return { success: false, message: 'Teacher name required.' };
    if (!requestedSlots || !requestedSlots.length) return { success: false, message: 'No requested slot to reserve.' };

    var slot = requestedSlots[0];
    if (!slot || !slot.date || !slot.slot) return { success: false, message: 'Invalid slot data.' };

    var range = _slotStringToUtcMs(slot.date, slot.slot, timeZone || 'Europe/Berlin');
    if (!range) return { success: false, message: 'Could not parse slot time: ' + slot.slot };

    var startMs = range[0], endMs = range[1];
    var startDt = new Date(startMs), endDt = new Date(endMs);

    var courseType = _courseTypeLabel(courseName, jlid);
    var title = 'Reserved : ' + (learnerName || 'Learner') + ' (' + (jlid || 'N/A') + ') : Jetlearn ' + courseType + ' Lesson'
      + (teacherId ? (' (' + teacherId + ')') : '');

    var desc  = 'JLID: ' + (jlid || 'N/A')
      + '\nLearner: ' + (learnerName || 'N/A')
      + '\nTeacher: ' + teacherName
      + '\nCourse: ' + (courseName || 'N/A')
      + '\nReserved via Teacher Intelligence Center'
      + (performedBy ? ('\nBy: ' + performedBy) : '');

    var guests = [];
    if (teacherEmail) guests.push(teacherEmail);
    (extraEmails || []).forEach(function(e) { if (e && guests.indexOf(e) === -1) guests.push(e); });

    var masterCal = CalendarApp.getCalendarById(CONFIG.CLASS_SCHEDULE_CALENDAR_ID);
    if (!masterCal) return { success: false, message: 'Class schedule calendar not accessible.' };

    var NUM_EVENTS = (numEvents > 0) ? numEvents : 12;
    var recurrence = CalendarApp.newRecurrence().addWeeklyRule().times(NUM_EVENTS);

    var series = masterCal.createEventSeries(title, startDt, endDt, recurrence, {
      description: desc,
      guests: guests.join(','),
      sendInvites: true
    });

    // Best-effort: also place series on teacher's own calendar if accessible and distinct
    if (teacherCalendarId && teacherCalendarId !== CONFIG.CLASS_SCHEDULE_CALENDAR_ID) {
      try {
        var teacherCal = CalendarApp.getCalendarById(teacherCalendarId);
        if (teacherCal) {
          teacherCal.createEventSeries(title, startDt, endDt, CalendarApp.newRecurrence().addWeeklyRule().times(NUM_EVENTS), { description: desc });
        }
      } catch(te) {
        Logger.log('[reserveCalendarSlot] Could not write to teacher calendar: ' + te.message);
      }
    }

    Logger.log('[reserveCalendarSlot] Booked ' + title + ' x' + NUM_EVENTS + ' weekly @ ' + startDt.toISOString() + ' - ' + endDt.toISOString());

    return {
      success: true,
      title: title,
      start: startDt.toISOString(),
      end: endDt.toISOString(),
      occurrences: NUM_EVENTS,
      eventSeriesId: series.getId ? series.getId() : ''
    };
  } catch(e) {
    Logger.log('[reserveCalendarSlot] Error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// Returns the UTC offset (in hours, can be fractional e.g. India = 5.5) of `timeZone`
// on the given date, using the IANA tz database via Utilities.formatDate.
function _tzOffsetHours(dateStr, timeZone) {
  try {
    var ref = new Date(dateStr + 'T12:00:00Z'); // midday UTC avoids date-boundary edge cases
    var s = Utilities.formatDate(ref, timeZone, 'Z'); // e.g. "+0100", "-0500", "+0530"
    var sign = s.charAt(0) === '-' ? -1 : 1;
    var hh = parseInt(s.substring(1, 3), 10);
    var mm = parseInt(s.substring(3, 5), 10);
    return sign * (hh + mm / 60);
  } catch(e) {
    return 1; // fall back to CET
  }
}

// Fetches the description/notes from the most recent upcoming (or recent past) calendar event
// for a given JLID on the class schedule calendar. Returns { success, description, eventTitle }.
function getExistingEventDescription(jlid) {
  try {
    if (!jlid) return { success: false, message: 'JLID required.' };

    // Use Advanced Calendar API to search by title — much faster than getEvents over a range
    var jlidUpper = String(jlid).toUpperCase().trim();
    var now = new Date();
    var results = Calendar.Events.list(CONFIG.CLASS_SCHEDULE_CALENDAR_ID, {
      q: jlidUpper,
      timeMin: new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000).toISOString(),
      timeMax: new Date(now.getTime() + 7  * 24 * 60 * 60 * 1000).toISOString(),
      singleEvents: true,
      orderBy: 'startTime',
      maxResults: 3
    });

    var items = (results && results.items) || [];
    var match = null;
    for (var i = 0; i < items.length; i++) {
      var desc = (items[i].description || '').trim();
      if (desc) { match = { desc: desc, title: items[i].summary || '' }; break; }
    }

    if (!match) return { success: false, message: 'No event with notes found for ' + jlid + '.' };

    // Strip HTML tags, convert <br> to newlines
    var clean = match.desc
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/p>/gi, '\n')
      .replace(/<[^>]+>/g, '')
      .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ').replace(/&#39;/g, "'").replace(/&quot;/g, '"')
      .replace(/\n{3,}/g, '\n\n')
      .trim();

    return { success: true, description: clean, eventTitle: match.title };
  } catch(e) {
    Logger.log('[getExistingEventDescription] ' + e.message);
    return { success: false, message: e.message };
  }
}

// One-time utility: patch ALL events on the class calendar to hide guest lists.
// Run once from GAS editor: patchAllEventsHideGuests()
function patchAllEventsHideGuests() {
  var now = new Date();
  var from = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  var to   = new Date(now.getTime() + 365 * 24 * 60 * 60 * 1000);
  var pageToken;
  var patched = 0, errors = 0;
  do {
    var params = { timeMin: from.toISOString(), timeMax: to.toISOString(), maxResults: 250, singleEvents: false };
    if (pageToken) params.pageToken = pageToken;
    var res = Calendar.Events.list(CONFIG.CLASS_SCHEDULE_CALENDAR_ID, params);
    (res.items || []).forEach(function(ev) {
      try {
        Calendar.Events.patch({ guestsCanSeeOtherGuests: false, guestsCanInviteOthers: false }, CONFIG.CLASS_SCHEDULE_CALENDAR_ID, ev.id);
        patched++;
      } catch(e) { errors++; Logger.log('Patch error ' + ev.id + ': ' + e.message); }
    });
    pageToken = res.nextPageToken;
  } while (pageToken);
  Logger.log('Patched: ' + patched + ', Errors: ' + errors);
}

// Parses "YYYY-MM-DD" + "HH:00 AM/PM - HH:00 AM/PM" as local time in `timeZone` into [startMs, endMs] UTC.
function _slotStringToUtcMs(dateStr, slotStr, timeZone) {
  var di = slotStr.indexOf(' - ');
  if (di === -1) return null;
  var startStr = slotStr.substring(0, di).trim();
  var endStr   = slotStr.substring(di + 3).trim();
  function p12(ts) {
    var m = ts.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
    if (!m) return null;
    var h = parseInt(m[1]), mn = parseInt(m[2]), ap = m[3].toUpperCase();
    if (ap === 'PM' && h !== 12) h += 12;
    if (ap === 'AM' && h === 12) h = 0;
    return { h: h, m: mn };
  }
  var st = p12(startStr), et = p12(endStr);
  if (!st || !et) return null;
  var dp = dateStr.split('-');
  var yr = parseInt(dp[0]), mo = parseInt(dp[1]), dy = parseInt(dp[2]);
  var tzOffMin = Math.round(_tzOffsetHours(dateStr, timeZone || 'Europe/Berlin') * 60);
  var endDy = (et.h <= st.h && et.h < 12) ? dy + 1 : dy;
  return [
    Date.UTC(yr, mo - 1, dy,    st.h, st.m - tzOffMin, 0),
    Date.UTC(yr, mo - 1, endDy, et.h, et.m - tzOffMin, 0)
  ];
}
