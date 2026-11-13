// ═══════════════════════════════════════════════════════════════════════════
// TICKET TRIAGE — SUGGEST ONLY
// For a migration ticket still at Triggered / WIP / CLS Rejected: reads the ticket notes (Gemini), works out the requested
// slots, finds teachers who are ready for the course and free on those slots (real calendars), and writes a suggestion.
// It NEVER moves a ticket, assigns a teacher or messages anyone. A person can post the suggestion as a ticket note.
// No new HubSpot properties: only existing ticket fields and the "Notes - Additional Comments" text are read.
// ═══════════════════════════════════════════════════════════════════════════
var TRI_PIPELINE_ID = '66161281';
var TRI_STAGES = { '128913747': 'Triggered', '128913748': 'WIP', '1030980247': 'CLS Rejected' };
var TRI_CLS_REASONS = ['escalation on teacher', 'teacher performance issue', 'course change after prm', 'teacher change after prm',
  'escalation on teacher post migration', 'misalignment of teacher due to ops intervention only'];
// Timezone words people write in notes -> IANA zones (DST handled by the zone).
// India, Australia, US/Canada, UK, Europe, Africa, Gulf, Asia-Pacific.
var TRI_TZ = {
  // UK / Ireland
  uk: 'Europe/London', ukt: 'Europe/London', gmt: 'Europe/London', bst: 'Europe/London', london: 'Europe/London', britain: 'Europe/London', ireland: 'Europe/Dublin',
  // Central Europe (and "Euro")
  cet: 'Europe/Paris', cest: 'Europe/Paris', euro: 'Europe/Paris', europe: 'Europe/Paris', paris: 'Europe/Paris', berlin: 'Europe/Berlin', germany: 'Europe/Berlin',
  // Eastern Europe
  eet: 'Europe/Athens', eest: 'Europe/Athens', athens: 'Europe/Athens',
  // India
  ist: 'Asia/Kolkata', india: 'Asia/Kolkata', indian: 'Asia/Kolkata',
  // Australia / New Zealand
  aest: 'Australia/Sydney', aedt: 'Australia/Sydney', aet: 'Australia/Sydney', sydney: 'Australia/Sydney', melbourne: 'Australia/Melbourne', australia: 'Australia/Sydney',
  brisbane: 'Australia/Brisbane', acst: 'Australia/Adelaide', acdt: 'Australia/Adelaide', adelaide: 'Australia/Adelaide', awst: 'Australia/Perth', perth: 'Australia/Perth',
  nzst: 'Pacific/Auckland', nzdt: 'Pacific/Auckland', nz: 'Pacific/Auckland', auckland: 'Pacific/Auckland',
  // US / Canada
  et: 'America/New_York', est: 'America/New_York', edt: 'America/New_York', newyork: 'America/New_York', toronto: 'America/Toronto', canada: 'America/Toronto',
  ct: 'America/Chicago', cst: 'America/Chicago', cdt: 'America/Chicago', chicago: 'America/Chicago',
  mt: 'America/Denver', mst: 'America/Denver', mdt: 'America/Denver', denver: 'America/Denver',
  pt: 'America/Los_Angeles', pst: 'America/Los_Angeles', pdt: 'America/Los_Angeles', la: 'America/Los_Angeles', losangeles: 'America/Los_Angeles', vancouver: 'America/Vancouver',
  // Africa
  sast: 'Africa/Johannesburg', southafrica: 'Africa/Johannesburg', johannesburg: 'Africa/Johannesburg',
  wat: 'Africa/Lagos', lagos: 'Africa/Lagos', nigeria: 'Africa/Lagos',
  eat: 'Africa/Nairobi', nairobi: 'Africa/Nairobi', kenya: 'Africa/Nairobi',
  cat: 'Africa/Maputo', cairo: 'Africa/Cairo', egypt: 'Africa/Cairo', africa: 'Africa/Johannesburg',
  // Gulf / Asia
  gst: 'Asia/Dubai', uae: 'Asia/Dubai', dubai: 'Asia/Dubai', sgt: 'Asia/Singapore', singapore: 'Asia/Singapore', hkt: 'Asia/Hong_Kong', hongkong: 'Asia/Hong_Kong',
  pkt: 'Asia/Karachi', pakistan: 'Asia/Karachi', bdt: 'Asia/Dhaka', bangladesh: 'Asia/Dhaka'
};
var TRI_DAYS = { sunday: 0, monday: 1, tuesday: 2, wednesday: 3, thursday: 4, friday: 5, saturday: 6 };
var TRI_DAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
var TRI_BUDGET_MS = 270000;
var TRI_MAX_TEACHERS = 24;
var TRI_MAX_SLOTS = 24;

function _triToken() { return PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY'); }
function _triIsNa(v) { var s = String(v == null ? '' : v).trim().toLowerCase(); return !s || s === 'n/a' || s === 'na' || s === '-' || s === '--'; }

// ── The list: tickets that still need a teacher ──────────────────────────────
function getTriageTickets() {
  try {
    var props = ['subject', 'learner_uid', 'hs_pipeline_stage', 'reason_of_migration__t_', 'current_course__t_', 'future_course_1', 'current_teacher__t_',
      'notes___additional_comments', 'createdate', 'hs_lastmodifieddate'];
    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/tickets/search', {
      method: 'post', headers: { 'Authorization': 'Bearer ' + _triToken(), 'Content-Type': 'application/json' },
      payload: JSON.stringify({
        filterGroups: [{ filters: [
          { propertyName: 'hs_pipeline', operator: 'EQ', value: TRI_PIPELINE_ID },
          { propertyName: 'hs_pipeline_stage', operator: 'IN', values: Object.keys(TRI_STAGES) }
        ] }],
        properties: props, sorts: [{ propertyName: 'createdate', direction: 'ASCENDING' }], limit: 100
      }), muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return { success: false, message: 'HubSpot search failed (HTTP ' + resp.getResponseCode() + ').' };
    var now = new Date().getTime();
    var rows = (JSON.parse(resp.getContentText()).results || []).map(function(t) {
      var p = t.properties || {};
      var reason = p.reason_of_migration__t_ || '';
      return {
        ticketId: t.id, jlid: p.learner_uid || '',
        learner: String(p.subject || '').replace(/\s*-?\s*Migration Request Raised\s*$/i, '').trim() || p.learner_uid || '',
        stage: p.hs_pipeline_stage, stageLabel: TRI_STAGES[p.hs_pipeline_stage] || '',
        reason: reason, needsCls: TRI_CLS_REASONS.indexOf(String(reason).trim().toLowerCase()) > -1,
        course: p.current_course__t_ ? (getCourseLabel(p.current_course__t_) || p.current_course__t_) : '',
        nextCourse: _triIsNa(p.future_course_1) ? '' : (getCourseLabel(p.future_course_1) || p.future_course_1),
        oldTeacher: p.current_teacher__t_ ? (getTeacherLabel(p.current_teacher__t_) || p.current_teacher__t_) : '',
        notes: _triIsNa(p.notes___additional_comments) ? '' : String(p.notes___additional_comments),
        daysWaiting: Math.max(0, Math.floor((now - new Date(p.createdate).getTime()) / 86400000))
      };
    }).filter(function(r) { return r.jlid; });
    return { success: true, rows: rows };
  } catch (e) {
    Logger.log('[Triage] list error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Gemini: free-text notes -> structured request ────────────────────────────
function _triParseNotes(t) {
  var prompt = 'You read a note a sales/ops person wrote on a migration ticket at JetLearn, an online coding school. The note asks for a NEW TEACHER and may state when the class should be. '
    + 'Extract exactly what it asks. Do not invent anything that is not in the note.\n\n'
    + 'Ticket facts: reason="' + t.reason + '", current course="' + t.course + '", next course="' + t.nextCourse + '".\n'
    + 'Note:\n"""' + t.notes + '"""\n\n'
    + 'Return JSON only:\n'
    + '{"slots":[{"day":"Monday","time":"18:00","tz":"CET"}],'
    + ' "windows":[{"days":["Monday","Tuesday"],"from":"17:00","to":"20:00","tz":"UKT"}],'
    + ' "classesPerWeek":1,'
    + ' "courseHint":"course name or abbreviation the note says to switch to, else null",'
    + ' "focus":"e.g. coding rather than robotics, else null",'
    + ' "humanSteps":["each thing a person must do that is NOT finding a teacher, e.g. free trial class first, schedule a call with the new teacher, add X to next lessons"],'
    + ' "confirmWithParent":true or false,'
    + ' "summary":"one plain sentence of what the note asks"}\n'
    + 'Rules: "slots" = one entry per specific day+time offered (24-hour HH:MM); several offered days/times = several entries (alternatives, not all required). '
    + '"windows" = a range such as "any weekday after 5pm" (weekday = Monday-Friday; "after 5pm" = from 17:00 to 20:00; "3pm-6pm" = from 15:00 to 18:00). '
    + '"tz" = the timezone WORD written in the note exactly as written or its closest standard form (CET, UK, IST, India, AEST, Sydney, EST, ET, CST, MST, PST, PT, Toronto, SAST, WAT, EAT, Lagos, Dubai...) or null if none is written. '
    + '"classesPerWeek" = how many classes a week are wanted (null if not stated; "one class per week" = 1). '
    + '"confirmWithParent" = true when the note says to confirm / check the time or anything else with the parent first. '
    + 'Use [] for empty lists.';
  var out = JSON.parse(_callGemini([{ role: 'user', parts: [{ text: prompt }] }], null, 1500));
  out.slots = (out.slots || []).filter(function(s) { return s && s.day && s.time; });
  out.windows = (out.windows || []).filter(function(w) { return w && w.days && w.days.length && w.from; });
  out.humanSteps = (out.humanSteps || []).map(String).filter(Boolean);
  return out;
}


// ── Backup reader: when the AI is busy / down, read the plain cases ("Friday 5:30 PM IST", "Fri 5pm or Tue 3:30pm PT") with rules ──
function _triFallbackParse(notes) {
  var text = String(notes || '');
  var slots = [];
  var tzWords = Object.keys(TRI_TZ);
  var findTz = function(chunk) {
    var words = String(chunk).toLowerCase().replace(/[^a-z ]/g, ' ').split(/\s+/);
    for (var i = 0; i < words.length; i++) { if (TRI_TZ[words[i]]) return words[i]; }
    return null;
  };
  var noteTz = findTz(text);
  var dayRe = /\b(mon|tue|tues|wed|thu|thur|thurs|fri|sat|sun)(?:day|sday|nesday|rsday|urday)?\b/ig;
  var m;
  while ((m = dayRe.exec(text)) !== null) {
    var after = text.slice(m.index + m[0].length, m.index + m[0].length + 40);
    var tm = after.match(/(\d{1,2})(?::(\d{2}))?\s*(am|pm)\b/i) || after.match(/\b([01]?\d|2[0-3]):([0-5]\d)\b()/);
    if (!tm) continue;
    var h = parseInt(tm[1], 10), mi = tm[2] ? parseInt(tm[2], 10) : 0, mer = (tm[3] || '').toLowerCase();
    if (mer === 'pm' && h < 12) h += 12; if (mer === 'am' && h === 12) h = 0;
    var tail = after.slice(after.indexOf(tm[0]) + tm[0].length, after.indexOf(tm[0]) + tm[0].length + 14);
    var key = String(m[1]).toLowerCase().slice(0, 3);
    var day = { mon: 'Monday', tue: 'Tuesday', wed: 'Wednesday', thu: 'Thursday', fri: 'Friday', sat: 'Saturday', sun: 'Sunday' }[key];
    slots.push({ day: day, time: ('0' + h).slice(-2) + ':' + ('0' + mi).slice(-2), tz: findTz(tail) || noteTz });
  }
  var cpw = (text.match(/(\d+|one|two|three)\s*(?:class(?:es)?|sessions?)\s*(?:per|a|\/)\s*week/i) || [])[1];
  var cpwNum = cpw ? ({ one: 1, two: 2, three: 3 }[String(cpw).toLowerCase()] || parseInt(cpw, 10)) : null;
  return { slots: slots, windows: [], classesPerWeek: cpwNum || null, courseHint: null, focus: null, humanSteps: [],
    confirmWithParent: /confirm[^.]{0,40}parent|check[^.]{0,30}parent/i.test(text), summary: slots.length ? 'Read with the simple reader (AI busy): ' + slots.map(function(s) { return s.day + ' ' + s.time + (s.tz ? ' ' + s.tz.toUpperCase() : ''); }).join(' or ') : '' };
}

// ── Which course matters: the one the learner will be TAUGHT by the new teacher ─────────────────────────────
// A course change -> the next course. Anything else (slot change, teacher change, escalation...) -> the current course.
// The other course is still shown / used for ranking, so the teacher is judged on both.
function _triCourseFocus(ticket) {
  var isCourseChange = /course\s*change/i.test(String(ticket.reason || ''));
  var cur = ticket.course || '', next = ticket.nextCourse || '';
  if (isCourseChange && next) return { primary: next, other: cur, source: 'the ticket\'s next course (this is a course change)' };
  if (cur) return { primary: cur, other: next, source: 'the learner\'s current course' + (isCourseChange ? '' : ' (not a course change)') };
  return { primary: next, other: '', source: 'the ticket\'s next course (no current course on the ticket)' };
}

// Slots as the Teacher Intelligence search wants them: CET date + "hh:mm AM - hh:mm PM"
function _triEngineSlots(slots) {
  var to12 = function(h, m) { var ap = h >= 12 ? 'PM' : 'AM'; var h12 = h % 12 || 12; return ('0' + h12).slice(-2) + ':' + ('0' + m).slice(-2) + ' ' + ap; };
  return slots.map(function(s) {
    var f = Utilities.formatDate(new Date(s.startMs), 'Europe/Paris', 'yyyy-MM-dd HH:mm').split(' ');
    var hm = f[1].split(':'), h = parseInt(hm[0], 10), m = parseInt(hm[1], 10);
    return { date: f[0], slot: to12(h, m) + ' - ' + to12((h + 1) % 24, m) };
  });
}

function _triDescribeDiff(min) {
  var a = Math.abs(min), dir = min < 0 ? 'earlier' : 'later';
  return (a % 60 === 0 ? (a / 60) + ' h' : a + ' min') + ' ' + dir;
}

// Free slots the search lists ("Fri, 13 Nov · 01:00 PM - 02:00 PM", CET) that are within 2 hours of a requested slot
function _triNearby(alternateSlots, slots) {
  var out = [], seen = {};
  var re = /([A-Za-z]{3}), (\d{1,2}) ([A-Za-z]{3}) · (\d{2}):(\d{2}) ([AP]M) - (\d{2}):(\d{2}) ([AP]M)/g, m;
  while ((m = re.exec(String(alternateSlots || ''))) !== null) {
    var key = (m[2] + ' ' + m[3]).toLowerCase();
    var h = parseInt(m[4], 10) % 12 + (m[6] === 'PM' ? 12 : 0), mi = parseInt(m[5], 10);
    slots.forEach(function(s) {
      var f = Utilities.formatDate(new Date(s.startMs), 'Europe/Paris', 'd MMM yyyy-MM-dd').split(' ');
      if ((f[0] + ' ' + f[1]).toLowerCase() !== key) return;
      var iso = f[2].split('-');
      var altMs = Date.UTC(parseInt(iso[0], 10), parseInt(iso[1], 10) - 1, parseInt(iso[2], 10), h, mi) - _tzOffsetHours(f[2], 'Europe/Paris') * 3600000;
      var diff = Math.round((altMs - s.startMs) / 60000);
      if (diff === 0 || Math.abs(diff) > 120) return;
      var label = Utilities.formatDate(new Date(altMs), 'Europe/Paris', 'EEE HH:mm') + ' CET';
      if (seen[label]) return; seen[label] = true;
      out.push({ label: label, diffMin: diff, text: label + ' (' + _triDescribeDiff(diff) + ')' });
    });
  }
  out.sort(function(a, b) { return Math.abs(a.diffMin) - Math.abs(b.diffMin); });
  return out;
}

// The Teacher Intelligence engine (searchMatchingTeachers): live calendars, course readiness, audit, red flags, ranked
function _triFromEngine(result, slots, course, otherCourse, need, warnings) {
  if (typeof searchMatchingTeachers !== 'function') return false;
  var step = 'Asking the Teacher Intelligence search (calendars, course readiness, audits)';
  _progBegin(step);
  var t0 = new Date().getTime();
  var res;
  try {
    res = searchMatchingTeachers({ currentCourse: course, futureCourse1: otherCourse || '', learnerAge: '', techTraits: '', mathTraits: '', requestedSlots: _triEngineSlots(slots) });
  } catch (e) { _progEnd(step, false, e.message); warnings.push('The Teacher Intelligence search failed (' + e.message + '); used the built-in calendar check instead.'); return false; }
  if (!res || !res.success) { _progEnd(step, false, (res && res.message) || 'no answer'); warnings.push('The Teacher Intelligence search returned no answer; used the built-in calendar check instead.'); return false; }
  var list = res.results || [];
  list.forEach(function(t, idx) {
    var label = String(t.slotMatch || '');
    var pm = label.match(/(\d+)\s*\/\s*(\d+)/);
    var matched = t.slotFullMatch ? slots.length : (pm ? parseInt(pm[1], 10) : 0);
    var nearby = t.slotFullMatch ? [] : _triNearby(t.alternateSlots, slots);
    var sc = parseFloat(String(t.avgClassScore || '').split('/')[0]);
    result.candidates.push({
      teacher: t.teacherName, fits: matched >= need, matchedCount: matched,
      freeSlots: t.slotFullMatch ? slots.map(function(s) { return s.label; }) : [],
      freeNote: (!t.slotFullMatch && matched > 0) ? 'free for ' + matched + ' of ' + slots.length + ' requested slots' : '',
      nearby: nearby, busySlots: [], courseProgress: t.currentCourseProgress, otherCourseProgress: t.futureCourse1Progress || '',
      auditScore: isNaN(sc) ? null : sc, auditGrade: t.auditGrade || '', redFlags: t.redFlagCount || 0, rank: idx
    });
  });
  result.candidates.sort(function(a, b) { return (b.fits - a.fits) || (b.matchedCount - a.matchedCount) || (b.nearby.length > 0) - (a.nearby.length > 0) || (a.rank - b.rank); });
  result.readyCount = list.length;
  result.viaEngine = true;
  if (res.relaxedFilter) warnings.push('Nobody was 71%+ ready, so teachers at 51%+ are included.');
  _progEnd(step, true, list.length + ' teachers · ' + Math.round((new Date().getTime() - t0) / 1000) + 's');
  return true;
}

// A short plain note ("Friday 5:30 PM IST") needs no AI: read it with rules (faster, and works when the AI is busy)
function _triSimpleNote(notes) {
  var text = String(notes || '').trim();
  if (!text || text.length > 110) return null;
  if (/trial|call|confirm|parent|weekday|weekend|any\b|either|after|before|between|robot|coding|focus|instead|first|add\b|course|please|suitable/i.test(text)) return null;
  var fb = _triFallbackParse(text);
  return fb.slots.length ? fb : null;
}

function _triTzOf(word) {
  var k = String(word || '').toLowerCase().replace(/[^a-z]/g, '');
  return TRI_TZ[k] || '';
}

function _triPadTime(s) {
  var m = String(s || '').match(/(\d{1,2})(?::(\d{2}))?/);
  if (!m) return null;
  return { h: Math.min(23, parseInt(m[1], 10)), m: m[2] ? parseInt(m[2], 10) : 0 };
}

// Next occurrence (starting tomorrow, local to tz) of a weekday + time -> { startMs, endMs, label, day }
function _triOccurrence(dayName, hhmm, ianaTz) {
  var dIdx = TRI_DAYS[String(dayName || '').trim().toLowerCase()];
  var tm = _triPadTime(hhmm);
  if (dIdx === undefined || !tm) return null;
  var today = Utilities.formatDate(new Date(), ianaTz, 'yyyy-MM-dd').split('-');
  var base = Date.UTC(parseInt(today[0], 10), parseInt(today[1], 10) - 1, parseInt(today[2], 10));
  var d;
  for (var i = 1; i <= 7; i++) { d = new Date(base + i * 86400000); if (d.getUTCDay() === dIdx) break; }
  var ds = d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
  var startMs = Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), tm.h, tm.m) - _tzOffsetHours(ds, ianaTz) * 3600000;
  var cet = Utilities.formatDate(new Date(startMs), 'Europe/Paris', 'EEE HH:mm');
  return { startMs: startMs, endMs: startMs + 3600000, label: cet + ' CET', day: TRI_DAY_NAMES[dIdx], local: dayName + ' ' + hhmm };
}

function _triBuildSlots(req, warnings) {
  var out = [], seen = {};
  var add = function(day, time, tzWord) {
    var iana = _triTzOf(tzWord);
    if (!iana) { iana = 'Europe/Paris'; if (warnings.indexOf('No timezone written in the note — assumed CET.') === -1) warnings.push('No timezone written in the note — assumed CET.'); }
    var o = _triOccurrence(day, time, iana);
    if (o && !seen[o.startMs]) { seen[o.startMs] = true; out.push(o); }
  };
  (req.slots || []).forEach(function(s) { add(s.day, s.time, s.tz); });
  (req.windows || []).forEach(function(w) {
    var from = _triPadTime(w.from), to = _triPadTime(w.to || '');
    var endH = to ? to.h : Math.min(21, from.h + 3);
    for (var h = from.h; h <= Math.min(endH, 21); h++) {
      (w.days || []).forEach(function(d) { add(d, ('0' + h).slice(-2) + ':' + ('0' + (h === from.h ? from.m : 0)).slice(-2), w.tz); });
    }
  });
  out.sort(function(a, b) { return a.startMs - b.startMs; });
  if (out.length > TRI_MAX_SLOTS) { warnings.push('The note allows many slots; checked the first ' + TRI_MAX_SLOTS + '.'); out = out.slice(0, TRI_MAX_SLOTS); }
  return out;
}

// ── Which course to search teachers for ──────────────────────────────────────
function _triResolveCourse(hint, ticket, headerCourses, warnings) {
  var h = String(hint || '').trim().toLowerCase();
  if (h) {
    var hit = headerCourses.filter(function(c) { return c.toLowerCase() === h; })[0]
      || headerCourses.filter(function(c) { return c.toLowerCase().indexOf(h) > -1; })[0]
      || headerCourses.filter(function(c) { return c.split(/[\s\-&]+/).filter(Boolean).map(function(w) { return w.charAt(0); }).join('').toLowerCase() === h.replace(/[^a-z]/g, ''); })[0];
    if (hit) return { course: hit.replace(/(^|\s)([a-z])/g, function(m, sp, ch) { return sp + ch.toUpperCase(); }), source: 'the note ("' + hint + '")' };
    warnings.push('The note names "' + hint + '" but no course in the teacher sheet matches it; used the ticket course instead.');
  }
  var c = ticket.nextCourse || ticket.course;
  return { course: c, source: ticket.nextCourse ? 'the ticket\'s next course' : 'the ticket\'s current course' };
}

// ── Who is ready for that course ─────────────────────────────────────────────
function _triReadyTeachers(course, excludeTeacher) {
  var sheet = _getCachedSheetData(CONFIG.SHEETS.TEACHER_COURSES) || [];
  var hidden = {}; try { var hm = getHiddenTeachersMap(); Object.keys(hm).forEach(function(n) { hidden[normalizeTeacherName(n)] = true; }); } catch (e) {}
  var exNorm = excludeTeacher ? normalizeTeacherName(excludeTeacher) : '';
  var out = [], seen = {};
  for (var i = 0; i < sheet.length; i++) {
    var raw = String(sheet[i][0] || '').trim();
    if (!raw || raw.toLowerCase() === 'teacher') continue;
    var norm = normalizeTeacherName(raw);
    if (seen[norm] || hidden[norm] || (exNorm && norm === exNorm)) continue;
    var prog = _lpCourseProgress(norm, course);
    var st = _parseProgressStatus(prog);
    if (st.teacherReady) { seen[norm] = true; out.push({ teacher: raw, norm: norm, progress: st.displayLabel, progNum: st.progNum }); }
  }
  return out;
}

// ── Busy checks ──────────────────────────────────────────────────────────────
function _triOverlaps(ev, a, b) {
  if (!ev || ev.status === 'cancelled' || ev.transparency === 'transparent') return false;
  if (/availability\s*hour/i.test(ev.summary || '')) return false;
  var s = ev.start && ev.start.dateTime ? new Date(ev.start.dateTime).getTime() : null;
  var e = ev.end && ev.end.dateTime ? new Date(ev.end.dateTime).getTime() : null;
  if (s === null || e === null) return false;     // all-day markers do not block a class hour
  return s < b && e > a;
}

function _triMasterBusy(masterEventsBySlot, slot, info, name) {
  var events = masterEventsBySlot[slot.startMs] || [];
  var parts = String(name).trim().toLowerCase().split(/\s+/).filter(function(p) { return p.length >= 3; });
  var calLow = (info.calendarId || '').toLowerCase(), mailLow = (info.email || '').toLowerCase();
  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    if (!_triOverlaps(ev, slot.startMs, slot.endMs)) continue;
    var guests = (ev.attendees || []).map(function(a) { return (a.email || '').toLowerCase(); });
    var guestMatch = (calLow && guests.indexOf(calLow) > -1) || (mailLow && guests.indexOf(mailLow) > -1);
    var title = (ev.summary || '').toLowerCase();
    var nameMatch = parts.length > 0 && parts.every(function(p) { return title.indexOf(p) > -1; });
    if (guestMatch || nameMatch) return ev.summary || 'Class';
  }
  return '';
}

function _triCalEvents(calId, minMs, maxMs) {
  var items = [], token = null, pages = 0;
  do {
    var args = { timeMin: new Date(minMs).toISOString(), timeMax: new Date(maxMs).toISOString(), singleEvents: true, maxResults: 250 };
    if (token) args.pageToken = token;
    var r = Calendar.Events.list(calId, args);
    items = items.concat(r.items || []);
    token = r.nextPageToken; pages++;
  } while (token && pages < 4);
  return items;
}

// ── The run ──────────────────────────────────────────────────────────────────
function triageSuggest(jlid) {
  jlid = String(jlid || '').trim();
  var key = 'migprog_' + jlid;
  if (jlid) { __PROG_KEY = key; _progWrite(key, function(o) { o.steps = []; o.done = false; o.result = ''; }); }
  var res;
  try { res = _triSuggestCore(jlid); }
  catch (e) { res = { success: false, message: e.message }; Logger.log('[Triage] ' + e.message + '\n' + (e.stack || '')); }
  finally { if (jlid) { _progWrite(key, function(o) { o.done = true; o.result = (res && (res.message || (res.recommendation && res.recommendation.why))) || ''; }); __PROG_KEY = ''; } }
  return res;
}

function _triSuggestCore(jlid) {
  if (!jlid) return { success: false, message: 'JLID required.' };
  var started = new Date().getTime();
  var warnings = [];

  _progBegin('Reading the ticket and its notes');
  var list = getTriageTickets();
  if (!list.success) { _progEnd('Reading the ticket and its notes', false, list.message); return list; }
  var ticket = list.rows.filter(function(r) { return String(r.jlid).toUpperCase() === jlid.toUpperCase(); })[0];
  if (!ticket) { _progEnd('Reading the ticket and its notes', false, 'No Triggered / WIP / Rejected ticket'); return { success: false, message: 'No ticket at Triggered, WIP or CLS Rejected for ' + jlid + '.' }; }
  _progEnd('Reading the ticket and its notes', true, ticket.stageLabel + ' · ' + (ticket.notes ? ticket.notes.length + ' characters of notes' : 'no notes'));

  // 1. Understand the note
  var req = { slots: [], windows: [], humanSteps: [], confirmWithParent: false, classesPerWeek: null, courseHint: null, focus: null, summary: '' };
  var simple = ticket.notes ? _triSimpleNote(ticket.notes) : null;
  if (simple) {
    _progBegin('Understanding the request');
    req = simple; _progEnd('Understanding the request', true, simple.summary.replace('Read with the simple reader (AI busy): ', ''));
  } else if (ticket.notes) {
    _progBegin('Understanding the request (AI)');
    try { req = _triParseNotes(ticket); _progEnd('Understanding the request (AI)', true, req.summary || ''); }
    catch (e) {
      _progEnd('Understanding the request (AI)', false, String(e.message).slice(0, 140));
      var fb = _triFallbackParse(ticket.notes);
      if (fb.slots.length) { req = fb; warnings.push('The AI was busy, so the note was read with a simple reader. Please check the slots shown, and anything else the note asks: ' + ticket.notes.slice(0, 200)); }
      else warnings.push('The AI could not read the note (' + String(e.message).slice(0, 120) + ') and the simple reader found no day and time. Read it yourself: ' + ticket.notes.slice(0, 200));
    }
  } else {
    warnings.push('There are no notes on this ticket, so there is no requested slot to check.');
  }

  // 2. Slots in real time
  var slots = _triBuildSlots(req, warnings);
  _progBegin('Working out the slots in CET');
  _progEnd('Working out the slots in CET', slots.length > 0, slots.length ? slots.map(function(s) { return s.label; }).join(', ') : 'no slot found in the note');

  // 3. Course + ready teachers
  var headerCourses = [];
  try { var tc = _lpGetTeacherCourseMap(); var cs = {}; Object.keys(tc).forEach(function(t) { Object.keys(tc[t]).forEach(function(c) { cs[c] = true; }); }); headerCourses = Object.keys(cs); } catch (e) {}
  var focus = _triCourseFocus(ticket);
  var cr = { course: focus.primary, source: focus.source };
  var hinted = _triResolveCourse(req.courseHint, ticket, headerCourses, warnings);
  if (req.courseHint && hinted.source.indexOf('the note') === 0) cr = hinted;          // the note names the course: that wins
  var otherCourse = (cr.course === focus.primary) ? focus.other : focus.primary;
  _progBegin('Finding teachers ready for ' + (cr.course || 'the course'));
  var ready = cr.course ? _triReadyTeachers(cr.course, ticket.oldTeacher) : [];
  _progEnd('Finding teachers ready for ' + (cr.course || 'the course'), ready.length > 0, ready.length + ' ready (71%+ upskilled) · ' + cr.source);

  var result = { success: true, ticket: ticket, course: cr.course, otherCourse: otherCourse, courseSource: cr.source,
    request: { summary: req.summary || '', slots: slots.map(function(s) { return { label: s.label, local: s.local }; }), classesPerWeek: req.classesPerWeek, humanSteps: req.humanSteps, confirmWithParent: !!req.confirmWithParent, focus: req.focus, warnings: warnings },
    candidates: [], readyCount: ready.length };

  // 4. Who is free — the Teacher Intelligence search first (same engine as that page); the built-in calendar check is the backup
  var need = Math.max(1, parseInt(req.classesPerWeek, 10) || 1);
  var engineOk = false;
  if (slots.length) engineOk = _triFromEngine(result, slots, cr.course, otherCourse, need, warnings);
  if (!engineOk && slots.length && ready.length) {
    var audit = {}; try { audit = _buildAuditScoreMap(); } catch (e) {}
    ready.forEach(function(r) { var a = audit[r.norm]; r.audit = a && a.avgScore != null ? a.avgScore : null; r.redFlags = a ? a.redFlags : 0; });
    ready.sort(function(a, b) { return ((b.audit || 0) - (a.audit || 0)) || (b.progNum - a.progNum); });
    var pool = ready.slice(0, TRI_MAX_TEACHERS);
    if (ready.length > pool.length) warnings.push('Checked the ' + pool.length + ' best-scoring of ' + ready.length + ' ready teachers.');

    var minMs = slots[0].startMs - 3600000, maxMs = slots[slots.length - 1].endMs + 3600000;
    var masterBySlot = {};
    _progBegin('Reading the class calendar for ' + slots.length + ' slot' + (slots.length === 1 ? '' : 's'));
    try { slots.forEach(function(s) { masterBySlot[s.startMs] = _triCalEvents(CONFIG.CLASS_SCHEDULE_CALENDAR_ID, s.startMs, s.endMs); }); _progEnd('Reading the class calendar for ' + slots.length + ' slot' + (slots.length === 1 ? '' : 's'), true); }
    catch (e) { _progEnd('Reading the class calendar for ' + slots.length + ' slot' + (slots.length === 1 ? '' : 's'), false, e.message); warnings.push('Class calendar could not be read (' + e.message + '); only personal calendars were checked.'); }

    var label = 'Checking ' + pool.length + ' teacher calendars';
    _progBegin(label);
    var timedOut = false, checked = 0;
    pool.forEach(function(r) {
      if (new Date().getTime() - started > TRI_BUDGET_MS) { timedOut = true; return; }
      var info = _lookupTeacherCalendarInfo(r.teacher);
      var personal = null, unverifiable = false;
      if (info.calendarId) { try { personal = _triCalEvents(info.calendarId, minMs, maxMs); } catch (e) { unverifiable = true; } } else unverifiable = true;
      var free = [], busy = [];
      slots.forEach(function(s) {
        var why = '';
        if (personal) { for (var i = 0; i < personal.length; i++) { if (_triOverlaps(personal[i], s.startMs, s.endMs)) { why = (personal[i].summary || 'Busy') + ' (their calendar)'; break; } } }
        if (!why) { var m = _triMasterBusy(masterBySlot, s, info, r.teacher); if (m) why = m + ' (class calendar)'; }
        if (why) busy.push({ slot: s.label, why: why }); else free.push({ slot: s.label, day: s.day });
      });
      var days = {}; free.forEach(function(f) { days[f.day] = true; });
      var fits = free.length >= need && Object.keys(days).length >= need;
      checked++;
      result.candidates.push({ teacher: r.teacher, fits: fits, freeSlots: free.map(function(f) { return f.slot; }), busySlots: busy, unverifiable: unverifiable && !personal,
        courseProgress: r.progress, auditScore: r.audit, redFlags: r.redFlags || 0 });
    });
    _progEnd(label, true, checked + ' checked' + (timedOut ? ' (stopped at the time limit)' : ''));
    if (timedOut) warnings.push('Ran out of time; ' + (pool.length - checked) + ' teachers were not checked.');
    result.candidates.sort(function(a, b) { return (b.fits - a.fits) || (b.freeSlots.length - a.freeSlots.length) || ((b.auditScore || 0) - (a.auditScore || 0)); });
  }

  // 5. The suggestion
  _progBegin('Writing the suggestion');
  var fitting = result.candidates.filter(function(c) { return c.fits; });
  var rec = { action: '', stageLabel: '', why: '' };
  if (req.humanSteps.length || req.confirmWithParent) {
    rec.action = 'needs_person'; rec.stageLabel = 'Stay where it is — needs a person';
    rec.why = 'The note asks for things only a person can do: ' + req.humanSteps.concat(req.confirmWithParent ? ['confirm with the parent first'] : []).join('; ') + '.';
  } else if (!slots.length) {
    rec.action = 'needs_person'; rec.stageLabel = 'Stay where it is — needs a person'; rec.why = 'No requested slot could be read from the ticket.';
  } else if (fitting.length) {
    rec.action = ticket.needsCls ? 'cls' : 'exec';
    rec.stageLabel = ticket.needsCls ? 'CLS Approval Pending' : 'Execution Pending';
    rec.why = fitting.length + ' teacher' + (fitting.length === 1 ? ' is' : 's are') + ' ready for ' + cr.course + ' and free' + (need > 1 ? ' for ' + need + ' classes a week' : '') + '. '
      + (ticket.needsCls ? 'The reason "' + ticket.reason + '" needs CLS approval first.' : 'The reason does not need CLS approval.');
  } else if (result.candidates.some(function(c) { return c.nearby && c.nearby.length; })) {
    var nb = result.candidates.filter(function(c) { return c.nearby && c.nearby.length; });
    rec.action = 'nearby'; rec.stageLabel = 'Stay where it is — ask the parent about a nearby time';
    rec.why = 'Nobody is free at the exact time, but ' + nb.length + ' teacher' + (nb.length === 1 ? ' is' : 's are') + ' free within 2 hours of it (for example ' + nb[0].teacher + ' at ' + nb[0].nearby[0].text + '). A different time needs the parent\'s OK.';
  } else {
    rec.action = 'tp'; rec.stageLabel = 'TP Approval Pending';
    rec.why = ready.length ? 'None of the ' + (result.candidates.length || ready.length) + ' checked teachers ready for ' + cr.course + ' is free on the requested slots.' : 'No teacher is ready for ' + (cr.course || 'this course') + '.';
  }
  if (ticket.stage === '1030980247') rec.why += ' This ticket was rejected by CLS — read their rejection note before acting.';
  result.recommendation = rec;
  result.checkedCount = result.candidates.length;
  result.noteText = _triNoteText(result);
  _progEnd('Writing the suggestion', true, rec.stageLabel);
  return result;
}

function _triNoteText(r) {
  var L = ['🤖 Teacher suggestion (suggest-only — nothing was changed)'];
  L.push('Request: ' + (r.request.summary || '—'));
  if (r.request.slots.length) L.push('Slots checked (CET): ' + r.request.slots.map(function(s) { return s.label; }).join(', '));
  L.push('Course: ' + (r.course || '—') + ' (from ' + r.courseSource + ')' + (r.otherCourse ? '; also considered: ' + r.otherCourse : ''));
  var fit = r.candidates.filter(function(c) { return c.fits; }).slice(0, 5);
  if (fit.length) { L.push('Teachers ready and free:'); fit.forEach(function(c) { L.push('• ' + c.teacher + ' — ' + (c.freeSlots.length ? 'free ' + c.freeSlots.join(', ') : c.freeNote) + (c.auditScore != null ? ' (audit ' + c.auditScore + ')' : '')); }); }
  else L.push('Teachers ready and free: none');
  var near = r.candidates.filter(function(c) { return !c.fits && c.nearby && c.nearby.length; }).slice(0, 5);
  if (near.length) { L.push('Free at a nearby time (within 2 hours — needs the parent\'s OK):'); near.forEach(function(c) { L.push('• ' + c.teacher + ' — ' + c.nearby.slice(0, 3).map(function(n) { return n.text; }).join(', ')); }); }
  if (r.request.humanSteps.length) L.push('Needs a person: ' + r.request.humanSteps.join('; '));
  if (r.request.confirmWithParent) L.push('Note says: confirm with the parent first.');
  L.push('Suggested next step: ' + r.recommendation.stageLabel + '. ' + r.recommendation.why);
  return L.join('\n');
}

function triagePostNote(ticketId, text) {
  try {
    if (!ticketId || !text) return { success: false, message: 'Ticket and text required.' };
    addNoteToHubSpotTicket(ticketId, String(text).replace(/\n/g, '<br>'));
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}
