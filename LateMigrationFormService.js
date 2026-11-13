/**
 * LateMigrationFormService.js
 *
 * Kits take about 4 weeks end to end, and the migration form is what starts the kit order. When a course-change
 * migration (CCTC) is filed less than 4 weeks before the learner's current course ends, and the NEXT course needs a kit,
 * the teacher is emailed (TP + CLS manager copied) so the process is followed next time. One email per ticket.
 *
 * Runs from the migration automation trigger and the daily kit follow-up trigger; safe to run often.
 */

var LATE_FORM_MIN_DAYS = 28;          // 1 class a week: the form is needed 4 weeks before the course ends
var LATE_FORM_MIN_DAYS_MULTI = 42;    // 2 or more classes a week: 6 weeks (the course finishes sooner, the kit takes as long)
// Notice needed, from how many classes a week the learner has
function _lfRequiredDays(perWeek) { return Math.round(Number(perWeek) || 1) >= 2 ? LATE_FORM_MIN_DAYS_MULTI : LATE_FORM_MIN_DAYS; }
function _lfPerWeekFromDays(dayStr) {
  var n = String(dayStr || '').split(/[;,]/).map(function(x) { return x.trim(); }).filter(Boolean).length;
  return n > 0 ? n : 1;
}
var LATE_FORM_LOOKBACK_DAYS = 3;
var LATE_FORM_GUIDE_URL = 'https://jetlearn-kit-links.web.app/teacher-guide/';
var LATE_FORM_PIPELINE_ID = '66161281';
var LATE_FORM_FORM_URL = 'https://share.hsforms.com/1RZKGkUG6Qc2kLHBVlTOmPw4lo43';   // the migration form teachers fill in
var LATE_FORM_STARTS_ON = '2026-11-07';   // only tickets created from this day on are checked (no emails for older ones)
var LATE_FORM_PROPS = ['hs_all_associated_contact_emails', 'regular_class_day__t_', 'subject', 'learner_uid', 'reason_of_migration__t_', 'current_teacher__t_', 'new_teacher', 'current_course__t_', 'future_course_1',
                       'pre_migration_last_class_conducted_date__t_', 'createdate', 'hs_pipeline_stage'];

function _lfDate(v) {
  if (!v) return null;
  var d = new Date(isNaN(Number(v)) ? v : Number(v));
  if (isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}
function _lfFmt(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'EEE d MMM yyyy'); }

function _lfNotifiedSet() {
  var raw = '';
  try { raw = PropertiesService.getScriptProperties().getProperty('LATE_FORM_NOTIFIED') || ''; } catch (e) {}
  var set = {};
  raw.split(',').forEach(function(id) { if (id) set[id] = true; });
  return set;
}
function _lfMarkNotified(set, id) {
  set[id] = true;
  var ids = Object.keys(set);
  if (ids.length > 400) ids = ids.slice(ids.length - 400);
  try { PropertiesService.getScriptProperties().setProperty('LATE_FORM_NOTIFIED', ids.join(',')); } catch (e) {}
}

// Which kit a course needs. HubSpot's course labels differ from the sheet names ("CoSpaces AR and VR",
// "Junior XI (Microbit Jr)"), so the known map is tried first and then the kit's own name in the course title.
function _lfKitForCourse(rawName, label) {
  var k = null;
  try { k = _getKitForCourse(label) || _getKitForCourse(rawName); } catch (e) {}
  if (k) return k;
  var n = String(label || '') + ' ' + String(rawName || '');
  if (/microbit/i.test(n)) return 'Microbit';
  if (/makey/i.test(n)) return 'Makey Makey';
  if (/cospaces|oculus|virtual reality|\bvr\b/i.test(n)) return 'VR Headset';
  if (/arduino/i.test(n)) return 'Arduino';
  return null;
}

// Is this teacher already upskilled on the course? true / false / null (could not tell)
function _lfUpskilled(teacherName, courseLabel) {
  try {
    var load = getTeacherSpecificLoad(resolveTeacherName(teacherName));
    if (!load || !load.success) return null;
    var want = normalizeCourseNameForMatch(courseLabel);
    return (load.courses || []).some(function(c) { return normalizeCourseNameForMatch(c.course) === want; });
  } catch (e) { return null; }
}

// teacher name -> email + TP + CLS (teacher sheet columns: 1 name, 6 TP name, 7 CLS name, 8 email, 9 CLS email, 10 TP email)
function _lfTeacher(name) {
  var out = { name: name, email: '', tpName: '', tpEmail: '', clsName: '', clsEmail: '' };
  try {
    var sheet = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA);
    var low = String(name || '').trim().toLowerCase();
    var norm = '';
    try { norm = normalizeTeacherName(resolveTeacherName(name)); } catch (e) {}
    for (var r = 1; (low || norm) && r < sheet.length; r++) {
      var rowName = String(sheet[r][1] || '').trim();
      var same = (low && rowName.toLowerCase() === low);
      if (!same && norm) { try { same = normalizeTeacherName(rowName) === norm; } catch (e) {} }
      if (same) {
        out.email = String(sheet[r][8] || '').trim(); out.tpName = String(sheet[r][6] || '').trim(); out.clsName = String(sheet[r][7] || '').trim();
        out.clsEmail = String(sheet[r][9] || '').trim(); out.tpEmail = String(sheet[r][10] || '').trim();
        break;
      }
    }
  } catch (e) { Logger.log('[LateForm] teacher lookup failed: ' + e.message); }
  return out;
}

// Who filled the form? The migration form attaches the SUBMITTER as the ticket's contact. Teachers use their own address
// (name.jetlearn@gmail.com); CLS, TP and ops use @jet-learn.com. A form filed by staff is not the teacher's mistake, so no email.
function _lfFilledByStaff(emailsStr, tr) {
  var list = String(emailsStr || '').toLowerCase().split(/[;,\s]+/).filter(Boolean);
  var staff = [String(tr.clsEmail || '').toLowerCase(), String(tr.tpEmail || '').toLowerCase()].filter(Boolean);
  return list.some(function(e) { return /@jet-learn\.com$/.test(e) || staff.indexOf(e) > -1; });
}

// What the email says, given the facts. Pure function so it can be checked on its own.
function _lfCompose(f) {
  var days = f.leadDays;
  var spec = { headerLabel: 'JetLearn · Migration form', tag: 'Delayed entry', tone: days <= 14 ? 'bad' : 'warn', learnerName: f.learner, jlid: f.jlid, teacherName: String(f.teacher || '').split(/\s+/)[0],
               paras: [], rows: [], actionTitle: 'Next learner', actionText: '', ctaLabel: 'Fill the migration form', ctaHref: LATE_FORM_FORM_URL, cta2Label: 'Read the teacher guide', cta2Href: LATE_FORM_GUIDE_URL };
  var notice = days > 0 ? 'only ' + days + ' day' + (days === 1 ? '' : 's') + ' before' : 'after';
  spec.paras.push('This is a delayed entry. The migration form for ' + f.learner + ' was filled on ' + _lfFmt(f.created) + ', ' + notice + ' the last class of their current course (' + _lfFmt(f.lastClass) + '). '
    + 'The next course, ' + f.nextCourse + ', needs a ' + f.kit + ' kit, and we ask for the form at least ' + (f.required / 7) + ' weeks ahead' + (f.perWeek >= 2 ? ' for a learner with ' + f.perWeek + ' classes a week' : '') + ' because that is what starts the kit order.');
  spec.paras.push('We are ordering the kit as quickly as we can, but it may not arrive in time. Your CLS manager will speak to you if lessons that need the kit have to wait. Please be careful with the next learner.');
  spec.rows.push({ l: 'Form filled', v: _lfFmt(f.created) });
  spec.rows.push({ l: 'Last class of current course', v: _lfFmt(f.lastClass) });
  spec.rows.push({ l: 'Notice given', v: (days > 0 ? days + ' days' : 'none') + ' (we need ' + f.required + (f.perWeek >= 2 ? ', for ' + f.perWeek + ' classes a week' : '') + ')' });
  spec.rows.push({ l: 'Next course', v: f.nextCourse });
  spec.rows.push({ l: 'Kit', v: f.kit });
  if (f.upskilled === false) spec.rows.push({ l: 'Upskilled on next course', v: 'Not yet' });
  else if (f.upskilled === true) spec.rows.push({ l: 'Upskilled on next course', v: 'Yes' });
  spec.actionText = f.upskilled === false
    ? 'Waiting to be upskilled is a common reason the form goes in late. Please fill the migration form ' + (f.required / 7) + ' weeks before the course ends, choose "Not Onboarded" if you are not trained on the next course, and tell ' + (f.tpName || 'your TP manager') + ' so the training and the kit order happen at the same time.'
    : 'Please fill the migration form for every learner 4 weeks before the course ends (6 weeks for a learner with 2 or more classes a week), after checking PRM Prep in Athena. For kit courses it is the first step of the kit order.';
  spec.footer = 'Your CLS and TP managers are copied so everyone sees the same dates. The guide explains the process and the questions we hear most.';
  return spec;
}

function scanLateMigrationForms() {
  var result = { checked: 0, sent: 0 };
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    if (!token) return result;
    var floor = _lfDate(LATE_FORM_STARTS_ON);
    var since = Math.max(new Date().getTime() - LATE_FORM_LOOKBACK_DAYS * 86400000, floor ? floor.getTime() : 0);
    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/tickets/search', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({
        filterGroups: [{ filters: [
          { propertyName: 'hs_pipeline', operator: 'EQ', value: LATE_FORM_PIPELINE_ID },
          { propertyName: 'createdate', operator: 'GTE', value: String(since) }
        ] }],
        properties: LATE_FORM_PROPS, sorts: [{ propertyName: 'createdate', direction: 'DESCENDING' }], limit: 100
      }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) { Logger.log('[LateForm] search failed: HTTP ' + resp.getResponseCode()); return result; }
    var tickets = JSON.parse(resp.getContentText()).results || [];
    var done = _lfNotifiedSet();

    tickets.forEach(function(t) {
      var p = t.properties || {};
      if (done[t.id]) return;
      if (String(p.reason_of_migration__t_ || '').toLowerCase().indexOf('course change') === -1) return;   // CCTC only
      result.checked++;

      var lastClass = _lfDate(p.pre_migration_last_class_conducted_date__t_);
      var created = _lfDate(p.createdate);
      if (!lastClass || !created) return;
      if (created.getTime() < (floor ? floor.getTime() : 0)) return;                    // before the start date: never alert
      var leadDays = Math.round((lastClass.getTime() - created.getTime()) / 86400000);
      if (leadDays < -60) { _lfMarkNotified(done, t.id); return; }                      // a typo'd date (e.g. year 2005), not a real late form
      var perWeek = _lfPerWeekFromDays(p.regular_class_day__t_);
      var required = _lfRequiredDays(perWeek);
      if (leadDays >= required) { _lfMarkNotified(done, t.id); return; }   // on time: never look at it again

      var rawNext = p.future_course_1;
      var nextCourse = rawNext ? (getCourseLabel(rawNext) || rawNext) : '';
      var kit = nextCourse ? _lfKitForCourse(rawNext, nextCourse) : null;
      if (!kit) { _lfMarkNotified(done, t.id); return; }                              // next course needs no kit

      var teacherName = '';
      try { teacherName = getTeacherLabel(p.current_teacher__t_) || ''; } catch (e) {}
      var tr = _lfTeacher(teacherName);
      // Filed by CLS / TP / ops instead of the teacher: nothing to tell the teacher
      var submitter = String(p.hs_all_associated_contact_emails || '').trim();
      if (!submitter) return;                                                      // contact not attached yet: look again on the next run
      if (_lfFilledByStaff(submitter, tr)) { _lfMarkNotified(done, t.id); return; }

      var learner = String(p.subject || '').replace(/\s*-\s*Migration Request.*$/i, '').trim() || p.learner_uid || 'the learner';
      if (!tr.email || !isValidEmail(tr.email)) {
        try { addNoteToHubSpotTicket(t.id, '⚠ Late migration form (' + leadDays + ' days notice, kit course ' + nextCourse + '), but no teacher email was found' + (teacherName ? ' for ' + teacherName : '') + ' so nobody was emailed.'); } catch (e) {}
        _lfMarkNotified(done, t.id);
        return;
      }

      var spec = _lfCompose({ learner: learner, jlid: p.learner_uid || '', teacher: teacherName, created: created, lastClass: lastClass, leadDays: leadDays, required: required, perWeek: perWeek, nextCourse: nextCourse, kit: kit,
                              upskilled: _lfUpskilled(teacherName, nextCourse), tpName: tr.tpName });
      try {
        var tpl = HtmlService.createTemplateFromFile('KitTeacherTemplate');
        tpl.data = spec;
        var cc = [tr.clsEmail, tr.tpEmail].filter(function(e, i, a) { return e && isValidEmail(e) && a.indexOf(e) === i; }).join(',');
        sendTrackedEmail({ to: tr.email, cc: cc, subject: 'Delayed entry: migration form for ' + learner + ' (' + (p.learner_uid || '') + ')', htmlBody: tpl.evaluate().getContent(), jlid: p.learner_uid || '' });
        try { addNoteToHubSpotTicket(t.id, '📨 Teacher ' + teacherName + ' emailed (CLS and TP copied): migration form filed ' + leadDays + ' days before the course change; the next course needs a ' + kit + ' kit. Guide: ' + LATE_FORM_GUIDE_URL); } catch (e) {}
        _lfMarkNotified(done, t.id);
        result.sent++;
      } catch (se) {
        Logger.log('[LateForm] send failed for ticket ' + t.id + ': ' + se.message);   // not marked: retried on the next run
      }
    });
    // The log: one row per course-change ticket, whatever happened to it (never blocks or changes the alerts above)
    try {
      if (typeof logLateForms === 'function') logLateForms(tickets, _lfNotifiedSet());
      if (typeof lateFormBackfill === 'function' && !PropertiesService.getScriptProperties().getProperty('LATE_FORM_LOG_BACKFILLED')) lateFormBackfill();   // once: tickets from before the log
    } catch (le) { Logger.log('[LateForm] log failed: ' + le.message); }
  } catch (e) {
    Logger.log('[LateForm] scan failed: ' + e.message);
  }
  return result;
}


// ═══════════════════════════════════════════════════════════════════════════
// Reminder BEFORE the deadline: "the migration form is due"
//
// From the course planner (getLearnerProgressions): a learner whose current course ends within 4 weeks / 4 classes (or within 5,
// as an early heads-up), whose next course needs a kit, and who has no migration ticket yet. One email per teacher (a digest of
// their learners), CLS + TP copied. Each learner is reminded once per step ("soon", then "now") per course.
// ═══════════════════════════════════════════════════════════════════════════
var KIT_FORM_REMINDER_MAX_EMAILS = 25;

function _kfStore() {
  var raw = '';
  try { raw = PropertiesService.getScriptProperties().getProperty('KIT_FORM_REMINDERS') || ''; } catch (e) {}
  var o = {};
  try { o = raw ? JSON.parse(raw) : {}; } catch (e) { o = {}; }
  var cutoff = new Date().getTime() - 150 * 86400000;
  Object.keys(o).forEach(function(k) { if (new Date(o[k]).getTime() < cutoff) delete o[k]; });
  return o;
}
function _kfSave(o) { try { PropertiesService.getScriptProperties().setProperty('KIT_FORM_REMINDERS', JSON.stringify(o)); } catch (e) {} }

// jlids (from a list) that already have a migration ticket created in the last 90 days, ignoring cancelled ones
function _kfJlidsWithTickets(jlids) {
  var have = {};
  var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  if (!token || !jlids.length) return have;
  var since = String(new Date().getTime() - 90 * 86400000);
  for (var i = 0; i < jlids.length; i += 50) {
    var chunk = jlids.slice(i, i + 50);
    try {
      var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/tickets/search', {
        method: 'post',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify({
          filterGroups: [{ filters: [
            { propertyName: 'hs_pipeline', operator: 'EQ', value: LATE_FORM_PIPELINE_ID },
            { propertyName: 'learner_uid', operator: 'IN', values: chunk },
            { propertyName: 'createdate', operator: 'GTE', value: since }
          ] }],
          properties: ['learner_uid', 'hs_pipeline_stage'], limit: 100
        }),
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() !== 200) { Logger.log('[FormReminder] ticket search failed: HTTP ' + resp.getResponseCode()); chunk.forEach(function(j) { have[j] = true; }); continue; }   // when unsure, do not nag
      (JSON.parse(resp.getContentText()).results || []).forEach(function(t) {
        var p = t.properties || {};
        if (String(p.hs_pipeline_stage) !== '133821818') have[p.learner_uid] = true;     // 133821818 = Migration Cancelled
      });
    } catch (e) { Logger.log('[FormReminder] ticket search error: ' + e.message); chunk.forEach(function(j) { have[j] = true; }); }
  }
  return have;
}

// The email for one teacher: a digest of their learners. Pure function.
function _kfCompose(teacher, items, tpName) {
  var now = items.some(function(i) { return i.tier === 'now'; });
  var notReady = items.some(function(i) { return !i.upskilledOk; });
  var spec = { headerLabel: 'JetLearn · Migration form due', tag: now ? 'Fill the form now' : 'Form due next week', tone: now ? 'warn' : 'info', learnerName: items.length === 1 ? items[0].learner : items.length + ' learners', jlid: items.length === 1 ? items[0].jlid : '',
               teacherName: String(teacher || '').split(/\s+/)[0], paras: [], rows: [], actionTitle: 'Before you fill it in', actionText: '', ctaLabel: 'Fill the migration form', ctaHref: LATE_FORM_FORM_URL, cta2Label: 'Read the teacher guide', cta2Href: LATE_FORM_GUIDE_URL };
  spec.paras.push(items.length === 1
    ? items[0].learner + ' is close to finishing ' + items[0].currentCourse + ' and the next course, ' + items[0].nextCourse + ', needs a ' + items[0].kit + ' kit.'
    : items.length + ' of your learners are close to finishing their course and the next course needs a kit.');
  spec.paras.push('The migration form is what starts the kit order. It is needed 4 weeks before the course ends, or 6 weeks for a learner with 2 or more classes a week, so the kit is with the learner on day 1. '
    + (now ? 'That point has been reached, so please fill it in now.' : 'That point is about a week away, so you can fill it in already.'));
  items.forEach(function(i) {
    spec.rows.push({ l: i.learner + ' (' + i.jlid + ')', v: i.nextCourse + ' · ' + i.kit + ' kit · ' + (i.classesLeft === 1 ? '1 class' : i.classesLeft + ' classes') + ' left' + (i.when ? ', ends about ' + i.when : '') + (i.perWeek >= 2 ? ' · ' + i.perWeek + ' classes a week, form needed ' + i.needWeeks + ' weeks ahead' : '') + (i.upskilledOk ? '' : ' · not upskilled yet') });
  });
  spec.actionText = 'Check PRM Prep in Athena first (the last class date, the next 3 courses, your upskilling). Type the learner ID exactly as shown, with no brackets or spaces. '
    + (notReady ? 'For the course you are not trained on, choose "Not Onboarded" and send the form anyway; ' + (tpName ? 'tell ' + tpName + ' ' : 'tell your TP manager ') + 'so the training and the kit order happen at the same time.' : 'Pick your upskilling level for the next course.');
  spec.footer = 'Your CLS and TP managers are copied. You get this once for each step, and not at all once the form is in.';
  return spec;
}

function runKitFormReminders() {
  var result = { candidates: 0, emails: 0, skipped: 0 };
  try {
    var lp = getLearnerProgressions(false);
    if (!lp || !lp.success || !lp.learners || !lp.learners.length) return result;
    var today = new Date().getTime();
    var cands = [];
    lp.learners.forEach(function(l) {
      var nc = (l.nextCourses || [])[0];
      if (!nc || !nc.name || !l.jlid || !l.teacher) return;
      var kit = _lfKitForCourse(nc.name, nc.name);
      if (!kit) return;
      var weeks = l.projectedWeeks, left = (l.classesLeft === undefined || l.classesLeft === null) ? null : Number(l.classesLeft);
      var perWeek = Math.max(1, Math.round(Number(l.classFrequency) || 1));
      var needWeeks = _lfRequiredDays(perWeek) / 7;            // 4 weeks for 1 class a week, 6 for 2 or more
      var tier = null;
      if (weeks !== null && weeks !== undefined) {
        if (weeks <= needWeeks) tier = 'now'; else if (weeks <= needWeeks + 1) tier = 'soon';
      } else if (left !== null) {                              // no pace known: fall back to "4 classes"
        if (left <= 4) tier = 'now'; else if (left <= 5) tier = 'soon';
      }
      if (!tier) return;
      var last = _lfDate(l.lastClassDate);                                         // learners who have stopped attending are left alone
      if (!last || (today - last.getTime()) > 35 * 86400000) return;
      cands.push({ l: l, nc: nc, kit: kit, tier: tier, left: left, perWeek: perWeek, needWeeks: needWeeks });
    });
    result.candidates = cands.length;
    if (!cands.length) return result;

    var haveTicket = _kfJlidsWithTickets(cands.map(function(c) { return c.l.jlid; }));
    var store = _kfStore();
    var byTeacher = {};
    cands.forEach(function(c) {
      if (haveTicket[c.l.jlid]) { result.skipped++; return; }
      var key = c.l.jlid + '|' + c.l.currentCourse + '|' + c.tier;
      if (store[key]) return;
      // the 'soon' step is skipped if the 'now' step was already sent
      if (c.tier === 'soon' && store[c.l.jlid + '|' + c.l.currentCourse + '|now']) return;
      (byTeacher[c.l.teacher] = byTeacher[c.l.teacher] || []).push({ c: c, key: key });
    });

    Object.keys(byTeacher).forEach(function(tName) {
      if (result.emails >= KIT_FORM_REMINDER_MAX_EMAILS) return;
      var tr = _lfTeacher(tName);
      var group = byTeacher[tName];
      if (!tr.email || !isValidEmail(tr.email)) { Logger.log('[FormReminder] no email for teacher ' + tName); result.skipped += group.length; return; }
      var items = group.map(function(g) {
        var nc = g.c.nc;
        var projected = g.c.l.projectedDate ? _lfDate(g.c.l.projectedDate) : null;
        return { learner: g.c.l.learnerName || g.c.l.jlid, jlid: g.c.l.jlid, currentCourse: g.c.l.currentCourse, nextCourse: nc.name, kit: g.c.kit, classesLeft: g.c.left === null ? 0 : g.c.left,
                 when: projected ? _lfFmt(projected) : '', tier: g.c.tier, upskilledOk: !!nc.teacherReady, perWeek: g.c.perWeek, needWeeks: g.c.needWeeks };
      });
      try {
        var spec = _kfCompose(tName, items, tr.tpName);
        var tpl = HtmlService.createTemplateFromFile('KitTeacherTemplate');
        tpl.data = spec;
        var cc = [tr.clsEmail, tr.tpEmail].filter(function(e, i, a) { return e && isValidEmail(e) && a.indexOf(e) === i; }).join(',');
        sendTrackedEmail({ to: tr.email, cc: cc, subject: 'Migration form due: ' + (items.length === 1 ? items[0].learner + ' (' + items[0].jlid + ')' : items.length + ' learners'), htmlBody: tpl.evaluate().getContent(), jlid: items.length === 1 ? items[0].jlid : '' });
        var stamp = new Date().toISOString().substring(0, 10);
        group.forEach(function(g) {
          store[g.key] = stamp;
          if (g.c.l.dealId) { try { _addNoteToDeal(g.c.l.dealId, '[Kit] Reminder sent to ' + tName + ' (CLS and TP copied): migration form due, next course ' + g.c.nc.name + ' needs a ' + g.c.kit + ' kit.'); } catch (ne) {} }
        });
        _kfSave(store);
        result.emails++;
      } catch (se) {
        Logger.log('[FormReminder] send failed for ' + tName + ': ' + se.message);    // not marked: tried again next run
      }
    });
  } catch (e) {
    Logger.log('[FormReminder] run failed: ' + e.message);
  }
  return result;
}

// ═══════════════════════════════════════════════════════════════════════════
// Triggers. Run setupTeacherKitTriggers() ONCE from the Apps Script editor (select it, click Run).
//   - every hour : check new course-change forms and email the teacher if they were filed late
//   - every day  : form reminders (before the deadline) and kit follow-ups for teachers, around 8am script time
// Both are safe to run more often: nothing is ever sent twice for the same ticket / learner step.
// ═══════════════════════════════════════════════════════════════════════════
function lateFormScanTrigger() {
  try { scanLateMigrationForms(); } catch (e) { Logger.log('[TeacherKit] late-form trigger failed: ' + e.message); }
}

function kitTeacherDailyTrigger() {
  try { runKitFormReminders(); } catch (e) { Logger.log('[TeacherKit] reminders trigger failed: ' + e.message); }
  try { if (typeof kitTeacherFollowups === 'function') kitTeacherFollowups(); } catch (e) { Logger.log('[TeacherKit] follow-ups trigger failed: ' + e.message); }
}

function setupTeacherKitTriggers() {
  var handlers = ['lateFormScanTrigger', 'kitTeacherDailyTrigger'];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (handlers.indexOf(t.getHandlerFunction()) > -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('lateFormScanTrigger').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('kitTeacherDailyTrigger').timeBased().atHour(8).everyDays(1).create();
  Logger.log('Teacher kit triggers installed: late-form check every hour, reminders and follow-ups daily around 8am.');
  return listTeacherKitTriggers();
}

// Which of the automatic jobs are installed right now (also prints to the log)
function listTeacherKitTriggers() {
  var names = { runMigrationAutomation: 'Migration automation (every 10 minutes; also checks late forms)', sendKitFollowUps: 'Kit follow-ups (daily; also checks late forms and sends reminders)',
                lateFormScanTrigger: 'Late-form check (hourly)', kitTeacherDailyTrigger: 'Form reminders + teacher kit follow-ups (daily)' };
  var found = ScriptApp.getProjectTriggers().map(function(t) { return t.getHandlerFunction(); }).filter(function(h) { return names[h]; });
  var out = Object.keys(names).map(function(h) { return { job: names[h], installed: found.indexOf(h) > -1 }; });
  out.forEach(function(o) { Logger.log((o.installed ? 'ON   ' : 'OFF  ') + o.job); });
  return out;
}

// Run both checks right now (for testing from the editor)
function runTeacherKitChecksNow() {
  var res = { lateForms: scanLateMigrationForms(), reminders: runKitFormReminders() };
  Logger.log(JSON.stringify(res));
  return res;
}
