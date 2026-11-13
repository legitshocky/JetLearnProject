// ═══════════════════════════════════════════════════════════════════════════
// LATE FORM LOG
// Every course-change (CCTC) migration ticket is judged on how much notice the form gave, and written to the review sheet:
//   "Late Form Log"      one row per ticket: who filed it, notice given, outcome. Two columns are filled in by a PERSON:
//                        "Exception" (e.g. "CLS told teacher") and "Exception Note". They are never overwritten.
//   "Late Form Summary"  counts that update themselves: late forms, of which exceptions, teacher-side, filed by CLS/TP/ops.
// Rows are written by the hourly late-form check and, once, for tickets from before it existed (backfill).
// Never sends anything and never blocks the check: every failure here is swallowed and logged.
// ═══════════════════════════════════════════════════════════════════════════
var LFL_LOG_TAB = 'Late Form Log';
var LFL_SUMMARY_TAB = 'Late Form Summary';
var LFL_BACKFILL_FROM = '2026-08-01';
var LFL_HEADERS = ['Ticket ID', 'Month', 'Form Filed', 'JLID', 'Learner', 'Teacher', 'CLS', 'TP', 'Filed By', 'Submitter Email', 'Current Course', 'Next Course', 'Kit',
  'Last Class', 'Notice (days)', 'Days Needed', 'Classes / Week', 'Outcome', 'Teacher Emailed', 'Exception', 'Exception Note', 'Ticket Link', 'Updated'];
var LFL = { ID: 1, MONTH: 2, FILED: 3, JLID: 4, LEARNER: 5, TEACHER: 6, CLS: 7, TP: 8, BY: 9, SUBMITTER: 10, CUR: 11, NEXT: 12, KIT: 13, LAST: 14, NOTICE: 15, NEEDED: 16, PERWEEK: 17,
  OUTCOME: 18, EMAILED: 19, EXC: 20, EXCNOTE: 21, LINK: 22, UPDATED: 23 };
var LFL_O = { ON_TIME: 'On time', TEACHER_LATE: 'Late - teacher filed', STAFF: 'Late - filed by CLS/TP/ops (not a teacher issue)', NO_KIT: 'Late - next course needs no kit',
  UNKNOWN_FILER: 'Late - filer not known yet', NO_DATA: 'Not enough data', BAD_DATE: 'Ignored - date looks wrong' };

function _lflSheets() {
  var ss = SpreadsheetApp.openById(OBL_SHEET_ID);
  var log = ss.getSheetByName(LFL_LOG_TAB);
  if (!log) {
    log = ss.insertSheet(LFL_LOG_TAB);
    log.appendRow(LFL_HEADERS); log.setFrozenRows(1);
    try {
      var rule = SpreadsheetApp.newDataValidation().requireValueInList(['CLS told teacher', 'TP told teacher', 'Ops told teacher', 'Other'], true).setAllowInvalid(true).build();
      log.getRange(2, LFL.EXC, 3000, 1).setDataValidation(rule);
    } catch (e) {}
  }
  return log;
}

// Pure: what happened to this ticket. Returns null when it is not a course-change ticket.
function _lflEvaluate(t, doneSet) {
  var p = t.properties || {};
  if (String(p.reason_of_migration__t_ || '').toLowerCase().indexOf('course change') === -1) return null;
  var created = _lfDate(p.createdate), lastClass = _lfDate(p.pre_migration_last_class_conducted_date__t_);
  var teacherName = ''; try { teacherName = getTeacherLabel(p.current_teacher__t_) || ''; } catch (e) {}
  var tr = _lfTeacher(teacherName);
  var submitter = String(p.hs_all_associated_contact_emails || '').trim();
  var list = submitter.toLowerCase().split(/[;,\s]+/).filter(Boolean);
  var filedBy = 'Unknown';
  if (list.length) {
    var cls = String(tr.clsEmail || '').toLowerCase(), tp = String(tr.tpEmail || '').toLowerCase();
    if (list.some(function(e) { return cls && e === cls; })) filedBy = 'CLS';
    else if (list.some(function(e) { return tp && e === tp; })) filedBy = 'TP';
    else if (list.some(function(e) { return /@jet-learn\.com$/.test(e); })) filedBy = 'Ops';
    else filedBy = 'Teacher';
  }
  var rawNext = p.future_course_1;
  var nextCourse = (rawNext && String(rawNext).toLowerCase() !== 'n/a') ? (getCourseLabel(rawNext) || rawNext) : '';
  var curCourse = p.current_course__t_ ? (getCourseLabel(p.current_course__t_) || p.current_course__t_) : '';
  var kit = nextCourse ? _lfKitForCourse(rawNext, nextCourse) : null;
  var perWeek = _lfPerWeekFromDays(p.regular_class_day__t_), required = _lfRequiredDays(perWeek);
  var learner = String(p.subject || '').replace(/\s*-\s*Migration Request.*$/i, '').trim() || p.learner_uid || '';
  var out = { id: t.id, created: created, jlid: p.learner_uid || '', learner: learner, teacher: teacherName, cls: tr.clsName || '', tp: tr.tpName || '', filedBy: filedBy, submitter: submitter,
    cur: curCourse, next: nextCourse, kit: kit || '', lastClass: lastClass, notice: '', needed: required, perWeek: perWeek, outcome: '', emailed: '' };
  if (!created || !lastClass) { out.outcome = LFL_O.NO_DATA; return out; }
  var lead = Math.round((lastClass.getTime() - created.getTime()) / 86400000);
  out.notice = lead;
  if (lead < -60) { out.outcome = LFL_O.BAD_DATE; return out; }
  if (lead >= required) { out.outcome = LFL_O.ON_TIME; return out; }
  if (!kit) { out.outcome = LFL_O.NO_KIT; return out; }
  if (!submitter) { out.outcome = LFL_O.UNKNOWN_FILER; return out; }
  if (filedBy !== 'Teacher') { out.outcome = LFL_O.STAFF; return out; }
  out.outcome = LFL_O.TEACHER_LATE;
  var floor = _lfDate(LATE_FORM_STARTS_ON);
  if (floor && created.getTime() < floor.getTime()) out.emailed = 'No (before the alerts started)';
  else if (!tr.email || !isValidEmail(tr.email)) out.emailed = 'No (no teacher email address)';
  else out.emailed = (doneSet && doneSet[t.id]) ? 'Yes' : 'Not yet';
  return out;
}

// Write / refresh the rows for these tickets. Existing rows keep the person's Exception columns.
function logLateForms(tickets, doneSet) {
  try {
    var log = _lflSheets();
    var last = log.getLastRow();
    var existing = {};
    if (last > 1) log.getRange(2, LFL.ID, last - 1, 1).getValues().forEach(function(r, i) { if (r[0] !== '') existing[String(r[0])] = i + 2; });
    var tz = Session.getScriptTimeZone(), now = new Date(), added = 0, updated = 0, fresh = [];
    (tickets || []).forEach(function(t) {
      var e = _lflEvaluate(t, doneSet);
      if (!e) return;
      var data = [e.id, e.created ? Utilities.formatDate(e.created, tz, 'yyyy-MM') : '', e.created || '', e.jlid, e.learner, e.teacher, e.cls, e.tp, e.filedBy, e.submitter, e.cur, e.next, e.kit,
        e.lastClass || '', e.notice, e.needed, e.perWeek, e.outcome, e.emailed];
      var link = 'https://app.hubspot.com/contacts/7729491/record/0-5/' + e.id;
      var row = existing[String(e.id)];
      if (row) {
        log.getRange(row, 1, 1, data.length).setValues([data]);
        log.getRange(row, LFL.LINK, 1, 2).setValues([[link, now]]);
        updated++;
      } else {
        fresh.push(data.concat(['', '', link, now]));
        added++;
      }
    });
    if (fresh.length) log.getRange(log.getLastRow() + 1, 1, fresh.length, LFL_HEADERS.length).setValues(fresh);
    lateFormRefreshSummary();
    return { success: true, added: added, updated: updated };
  } catch (e) {
    Logger.log('[LateFormLog] failed: ' + e.message);
    return { success: false, message: e.message };
  }
}

// The summary tab: headline counts + by month + by teacher. Labels are written by the script, the numbers are live formulas, so
// marking an Exception in the log updates them straight away.
function lateFormRefreshSummary() {
  try {
    var ss = SpreadsheetApp.openById(OBL_SHEET_ID);
    var log = _lflSheets();
    var sum = ss.getSheetByName(LFL_SUMMARY_TAB) || ss.insertSheet(LFL_SUMMARY_TAB);
    sum.clear();
    var Q = "'" + LFL_LOG_TAB + "'!";
    var col = function(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
    var rng = function(c) { return Q + col(c) + '2:' + col(c) + '3000'; };
    var OUT = rng(LFL.OUTCOME), EXC = rng(LFL.EXC), MON = rng(LFL.MONTH), TCH = rng(LFL.TEACHER), BY = rng(LFL.BY);
    var teacherLate = 'COUNTIF(' + OUT + ',"' + LFL_O.TEACHER_LATE + '")';
    var teacherExc = 'COUNTIFS(' + OUT + ',"' + LFL_O.TEACHER_LATE + '",' + EXC + ',"<>")';

    sum.getRange('A1').setValue('Late migration forms (course change, next course needs a kit)').setFontWeight('bold').setFontSize(13);
    sum.getRange('A2').setValue('Mark an Exception on a row of "' + LFL_LOG_TAB + '" (for example CLS told the teacher to fill it) and these numbers update.').setFontColor('#6b7280');
    var rows = [
      ['Course-change tickets checked', '=COUNTA(' + rng(LFL.ID) + ')'],
      ['Filed with enough notice', '=COUNTIF(' + OUT + ',"' + LFL_O.ON_TIME + '")'],
      ['Late, filed by a teacher', '=' + teacherLate],
      ['   of which exceptions (told to fill it)', '=' + teacherExc],
      ['   teacher-side (no exception)', '=B5-B6'],
      ['Late, filed by CLS / TP / ops', '=COUNTIF(' + OUT + ',"Late - filed by*")'],
      ['Late, but the next course needs no kit', '=COUNTIF(' + OUT + ',"' + LFL_O.NO_KIT + '")'],
      ['Late, filer not known yet', '=COUNTIF(' + OUT + ',"' + LFL_O.UNKNOWN_FILER + '")']
    ];
    sum.getRange(4, 1, rows.length, 2).setValues(rows.map(function(r) { return [r[0], r[1]]; }));
    sum.getRange('A13').setFormula('="Late forms from teachers: "&B5&" — "&B6&" exception(s), "&B7&" teacher-side. Filed by CLS/TP/ops (left out): "&B8&"."').setFontWeight('bold');

    var data = log.getLastRow() > 1 ? log.getRange(2, 1, log.getLastRow() - 1, LFL_HEADERS.length).getValues() : [];
    var months = {}, teachers = {};
    data.forEach(function(r) { if (r[LFL.MONTH - 1]) months[r[LFL.MONTH - 1]] = true; if (r[LFL.TEACHER - 1] && r[LFL.OUTCOME - 1] === LFL_O.TEACHER_LATE) teachers[r[LFL.TEACHER - 1]] = true; });
    var mKeys = Object.keys(months).sort().reverse(), tKeys = Object.keys(teachers).sort();

    sum.getRange('A15:E15').setValues([['Month', 'Checked', 'Late (teacher)', 'Exceptions', 'Teacher-side']]).setFontWeight('bold');
    var mRows = mKeys.map(function(m, i) {
      var r = 16 + i;
      return [m, '=COUNTIF(' + MON + ',A' + r + ')', '=COUNTIFS(' + MON + ',A' + r + ',' + OUT + ',"' + LFL_O.TEACHER_LATE + '")',
        '=COUNTIFS(' + MON + ',A' + r + ',' + OUT + ',"' + LFL_O.TEACHER_LATE + '",' + EXC + ',"<>")', '=C' + r + '-D' + r];
    });
    if (mRows.length) sum.getRange(16, 1, mRows.length, 5).setValues(mRows);

    var tStart = 16 + mRows.length + 2;
    sum.getRange(tStart - 1, 1, 1, 4).setValues([['Teacher (late forms)', 'Late', 'Exceptions', 'Teacher-side']]).setFontWeight('bold');
    var tRows = tKeys.map(function(t, i) {
      var r = tStart + i;
      return [t, '=COUNTIFS(' + TCH + ',A' + r + ',' + OUT + ',"' + LFL_O.TEACHER_LATE + '")', '=COUNTIFS(' + TCH + ',A' + r + ',' + OUT + ',"' + LFL_O.TEACHER_LATE + '",' + EXC + ',"<>")', '=B' + r + '-C' + r];
    });
    if (tRows.length) sum.getRange(tStart, 1, tRows.length, 4).setValues(tRows);
    sum.setColumnWidth(1, 330);
    return { success: true };
  } catch (e) {
    Logger.log('[LateFormLog] summary failed: ' + e.message);
    return { success: false, message: e.message };
  }
}

// Tickets from before the log existed. Safe to run again: existing rows are refreshed, not duplicated.
function lateFormBackfill() {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    if (!token) return { success: false, message: 'No HubSpot key.' };
    var from = _lfDate(LFL_BACKFILL_FROM);
    var done = _lfNotifiedSet();
    var all = [], after = null, pages = 0;
    do {
      var body = { filterGroups: [{ filters: [
        { propertyName: 'hs_pipeline', operator: 'EQ', value: LATE_FORM_PIPELINE_ID },
        { propertyName: 'createdate', operator: 'GTE', value: String(from ? from.getTime() : 0) }
      ] }], properties: LATE_FORM_PROPS, sorts: [{ propertyName: 'createdate', direction: 'DESCENDING' }], limit: 100 };
      if (after) body.after = after;
      var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/tickets/search', { method: 'post', headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' }, payload: JSON.stringify(body), muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) { Logger.log('[LateFormLog] backfill search failed: HTTP ' + resp.getResponseCode()); break; }
      var j = JSON.parse(resp.getContentText());
      all = all.concat(j.results || []);
      after = (j.paging && j.paging.next && j.paging.next.after) || null;
      pages++;
    } while (after && pages < 12);
    var res = logLateForms(all, done);
    try { PropertiesService.getScriptProperties().setProperty('LATE_FORM_LOG_BACKFILLED', new Date().toISOString()); } catch (e) {}
    Logger.log('[LateFormLog] backfill: ' + all.length + ' tickets read, ' + JSON.stringify(res));
    return { success: true, read: all.length, result: res };
  } catch (e) {
    Logger.log('[LateFormLog] backfill failed: ' + e.message);
    return { success: false, message: e.message };
  }
}
