// ═══════════════════════════════════════════════════════════════════════════
// TP WORKSPACE — upskilling requirements that TP managers see and act on in the app (not just an email).
//
// Data: the existing "Upskill Tracker" sheet in JetLearn App Data. Columns A-K are the old layout (kept, so nothing that reads it breaks);
// L onwards are new. A row is a NEW-style request only when column AB ("Version") says "v2" — every older row is ignored, so
// old gaps never reach the dashboard. One row = one teacher + one course (learners waiting on it are collected in the JLID column).
//
// Status flow:  Requested -> Started -> Verifying -> Verified
//                         \-> Not Possible (reason)         Verifying -> Rejected (still not on the data after 48 h) -> Started/Verifying again
// Rules: 75%+ on the Teacher Courses sheet = upskilled. "Done" is cross-checked at once and again after 24-48 h (Moodle lags the sheet).
//        No start in 14 days -> HubSpot task for the TP manager. Reminder emails every 2 days while something needs them.
// Access: TP managers (role "TP Manager", listed in the "TP Managers" sheet) see only their teachers. Super Admin / Admin see everything.
// ═══════════════════════════════════════════════════════════════════════════

var TPW = {
  VERSION: 'v2',
  PASS_PCT: 75,
  FAST_PCT: 50,
  VERIFY_AFTER_HOURS: 24,
  REJECT_AFTER_HOURS: 48,
  NO_START_TASK_DAYS: 14,
  DIGEST_EVERY_DAYS: 2,
  SHEET_MANAGERS: 'TP Managers',
  SHEET_ACTIVITY: 'Upskill Activity',
  OPEN: ['Requested', 'Started', 'Verifying', 'Rejected']
};

// 1-based column numbers in the Upskill Tracker sheet
var TPC = {
  ID: 1, TEACHER: 2, MANAGER: 3, JLIDS: 4, LEARNERS: 5, COURSE: 6, HS_TASK: 7, STATUS: 8, CREATED: 9, COMPLETED: 10, NOTES: 11,
  TIER: 12, PCT_REQ: 13, SOURCE: 14, STARTED: 15, DONE_MARKED: 16, VERIFIED: 17, LAST_CHECK: 18, LAST_PCT: 19,
  REASON_CODE: 20, REASON_NOTE: 21, LAST_REMINDER: 22, REMINDERS: 23, REQUESTED_BY: 24, LEARNER_COUNT: 25, REJECTED_COUNT: 26,
  UPDATED: 27, VERSION: 28
};
var TPW_HEADERS = ['ID', 'Teacher Name', 'TP Manager', 'JLID', 'Learner', 'Gap Courses', 'HubSpot Task ID', 'Status', 'Created At', 'Completed At', 'Notes',
  'Tier', 'Progress At Request', 'Source', 'Started At', 'Done Marked At', 'Verified At', 'Last Check At', 'Last Check %',
  'Reason Code', 'Reason Note', 'Last Reminder At', 'Reminders Sent', 'Requested By', 'Learner Count', 'Rejected Count', 'Updated At', 'Version'];

// Why a teacher cannot be upskilled on a course. The TP picks one; Other needs a note.
var TPW_REASONS = [
  { code: 'FOUNDATION_POD',   label: 'Foundation pod teacher — Pro roadmap courses are not allowed' },
  { code: 'POD_LEVEL',        label: "Teacher's pod level does not cover this course (Advanced / Professional pod rules)" },
  { code: 'SYSTEM_UNSUPPORTED', label: "Teacher's laptop / system does not support this course's software" },
  { code: 'INTERNET',         label: "Teacher's internet cannot run this course's tools reliably" },
  { code: 'DECLINED',         label: 'Teacher declined / not interested' },
  { code: 'NOTICE',           label: 'Teacher is serving notice / leaving' },
  { code: 'ON_LEAVE',         label: 'Teacher is on extended leave' },
  { code: 'AT_CAPACITY',      label: 'Teacher is already in other upskilling — at capacity' },
  { code: 'QUALITY',          label: 'Quality / audit concerns — not ready for a new course' },
  { code: 'PREREQ',           label: 'Missing prerequisite course' },
  { code: 'PART_TIME',        label: 'Part-time availability — cannot take more courses' },
  { code: 'LICENCE',          label: "Course tools / licence not available in the teacher's region" },
  { code: 'NO_TRAINING',      label: 'Training for this course is not live on Moodle yet' },
  { code: 'SUBJECT_MISMATCH', label: 'Subject mismatch (for example a maths-only teacher on a coding course)' },
  { code: 'OTHER',            label: 'Other (note required)' }
];

// ───────────────────────────── small helpers ─────────────────────────────
function _tpwNorm(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }
function _tpwTz() { return Session.getScriptTimeZone() || 'Asia/Kolkata'; }
function _tpwToDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var s = String(v).trim();
  if (!s) return null;
  if (/^\d{4}-\d\d-\d\dT/.test(s)) { var di = new Date(s); return isNaN(di.getTime()) ? null : di; } // ISO strings made by _tpwIso
  try { var d = Utilities.parseDate(s, _tpwTz(), 'yyyy-MM-dd HH:mm:ss'); if (d && !isNaN(d.getTime())) return d; } catch (e) {}
  var d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}
function _tpwIso(v) { var d = _tpwToDate(v); return d ? d.toISOString() : ''; }
function _tpwHoursSince(v) { var d = _tpwToDate(v); return d ? (Date.now() - d.getTime()) / 3600000 : null; }
function _tpwDaysSince(v) { var h = _tpwHoursSince(v); return h === null ? null : Math.floor(h / 24); }
function _tpwAppUrl() { try { return ScriptApp.getService().getUrl(); } catch (e) { return ''; } }
function _tpwEsc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

function _tpwWithLock(fn) {
  var lock = LockService.getScriptLock();
  var got = false;
  try { got = lock.tryLock(20000); } catch (e) {}
  try { return fn(); } finally { if (got) { try { lock.releaseLock(); } catch (e2) {} } }
}

// ───────────────────────────── sheets ─────────────────────────────
function _tpwTracker() {
  var sh = getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.UPSKILL_TRACKER);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var head = sh.getRange(1, 1, 1, Math.max(lastCol, TPW_HEADERS.length)).getValues()[0];
  var changed = false;
  for (var i = 0; i < TPW_HEADERS.length; i++) {
    if (!String(head[i] || '').trim()) { sh.getRange(1, i + 1).setValue(TPW_HEADERS[i]); changed = true; }
  }
  if (changed) { try { sh.getRange(1, 12, 1, TPW_HEADERS.length - 11).setFontWeight('bold'); } catch (e) {} }
  return sh;
}

function _tpwActivitySheet() {
  var sh = getOrCreateAppDataSheet(TPW.SHEET_ACTIVITY);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'Request ID', 'Teacher', 'Course', 'By', 'Role', 'Action', 'From', 'To', 'Detail']);
    try { sh.getRange(1, 1, 1, 10).setFontWeight('bold'); sh.setFrozenRows(1); } catch (e) {}
  }
  return sh;
}

function _tpwLog(req, by, role, action, from, to, detail) {
  try {
    _tpwActivitySheet().appendRow([new Date(), req ? req.id : '', req ? req.teacher : '', req ? req.course : '', by || 'system', role || '', action, from || '', to || '', detail || '']);
  } catch (e) { Logger.log('[tpw] activity log failed: ' + e.message); }
}

function _tpwManagersSheet() {
  var sh = getOrCreateAppDataSheet(TPW.SHEET_MANAGERS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Name', 'Email', 'Active', 'Invited At', 'Added By', 'Added At', 'Last Login', 'HubSpot Owner ID', 'First Login']);
    try { sh.getRange(1, 1, 1, 9).setFontWeight('bold'); sh.setFrozenRows(1); } catch (e) {}
  } else if (!String(sh.getRange(1, 8).getValue() || '').trim()) {
    sh.getRange(1, 8, 1, 2).setValues([['HubSpot Owner ID', 'First Login']]);
  }
  return sh;
}

function _tpwReadManagers() {
  var sh = _tpwManagersSheet();
  var vals = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < vals.length; i++) {
    var name = String(vals[i][0] || '').trim(), email = String(vals[i][1] || '').trim().toLowerCase();
    if (!name && !email) continue;
    out.push({ row: i + 1, name: name, email: email, active: /^(yes|true|y|1)$/i.test(String(vals[i][2] || '').trim()),
      invitedAt: _tpwIso(vals[i][3]), addedBy: String(vals[i][4] || ''), addedAt: _tpwIso(vals[i][5]), lastLogin: _tpwIso(vals[i][6]),
      hubspotId: String(vals[i][7] || '').trim(), firstLogin: _tpwIso(vals[i][8]) });
  }
  return out;
}

// Used by the sign-in functions: is this email an ACTIVE TP manager? (returns the manager row or null)
function _tpwFindManagerByEmail(email) {
  try {
    var e = String(email || '').trim().toLowerCase();
    if (!e) return null;
    var list = _tpwReadManagers();
    for (var i = 0; i < list.length; i++) { if (list[i].email === e && list[i].active) return list[i]; }
  } catch (err) { Logger.log('[tpw] _tpwFindManagerByEmail: ' + err.message); }
  return null;
}

function _tpwTouchManagerLogin(email) {
  try {
    var e = String(email || '').trim().toLowerCase();
    var list = _tpwReadManagers();
    for (var i = 0; i < list.length; i++) {
      if (list[i].email === e) {
        var sh = _tpwManagersSheet();
        sh.getRange(list[i].row, 7).setValue(new Date());
        if (!list[i].firstLogin) sh.getRange(list[i].row, 9).setValue(new Date());
        return;
      }
    }
  } catch (err) {}
}

// ───────────────────────────── who is asking, and which teachers they may see ─────────────────────────────
// Effective manager of every teacher: the Persona sheet's Manager column wins (same rule the Teacher Intelligence list uses), else Teacher Data.
function _tpwTeacherDirectory() {
  var list = getTeacherData();
  var persona = {};
  try {
    var pd = _getCachedSheetData(CONFIG.SHEETS.PERSONA_DATA, CONFIG.PERSONA_SHEET_ID);
    if (pd && pd.length > 2) {
      var hm = {};
      pd[1].forEach(function(h, i) { if (h) hm[String(h).trim()] = i; });
      var nc = hm['Teacher Name'], mc = hm['Manager'];
      if (nc !== undefined && mc !== undefined) {
        for (var r = 2; r < pd.length; r++) {
          var n = _tpwNorm(pd[r][nc]);
          if (n) persona[n] = String(pd[r][mc] || '').trim();
        }
      }
    }
  } catch (e) { Logger.log('[tpw] persona manager map: ' + e.message); }
  return list.map(function(t) {
    var key = _tpwNorm(t.name);
    var resolved = '';
    try { resolved = _tpwNorm(resolveTeacherName(t.name)); } catch (e) {}
    var pm = persona[key] || persona[resolved] || '';
    return {
      name: t.name, status: t.status || 'Active', hidden: t.hiddenInSearch === 'Yes',
      manager: pm || t.manager || '',
      personaMgr: pm, dataMgr: String(t.manager || '').trim(),
      source: pm ? 'Persona sheet' : (t.manager ? 'Teacher Data' : ''),
      tpEmail: String(t.tpManagerEmail || '').trim().toLowerCase()
    };
  });
}

function _tpwAllManagerNames(dir) {
  var seen = {}, out = [];
  var add = function(n) { var k = _tpwNorm(n); if (k && !seen[k]) { seen[k] = true; out.push(String(n).trim()); } };
  _tpwReadManagers().forEach(function(m) { add(m.name); });
  (dir || _tpwTeacherDirectory()).forEach(function(t) { add(t.manager); });
  try { getTPManagers().forEach(add); } catch (e) {}
  return out.sort();
}

function _tpwScopeFor(mgr, dir) {
  var out = {};
  (dir || _tpwTeacherDirectory()).forEach(function(t) {
    if ((mgr.name && _tpwNorm(t.manager) === _tpwNorm(mgr.name)) || (mgr.email && t.tpEmail && t.tpEmail === mgr.email)) out[_tpwNorm(t.name)] = t;
  });
  return out;
}

// Resolves a session token into { ok, isAdmin, role, email, name, scope } — every TP function starts here.
function _tpwCtx(token, viewAs) {
  var s = _getSession(token);
  if (!s) return { ok: false, message: 'Your session expired. Please sign in again.' };
  var role = s.role;
  var perms = s.permissions || [];
  var isAdmin = (role === ROLES.SUPER_ADMIN || role === ROLES.ADMIN);
  var email = String(s.username || '').trim().toLowerCase();
  if (isAdmin && viewAs) {
    // An admin looking at exactly what one TP manager sees (and able to act as them). Actions are logged as "Admin (as Manager)".
    var vm = _tpwReadManagers().filter(function(m) { return _tpwNorm(m.name) === _tpwNorm(viewAs); })[0];
    if (!vm) vm = { name: String(viewAs).trim(), email: '' }; // a manager who is in Teacher Data but not added to the TP Managers sheet yet
    var vscope = _tpwScopeFor(vm);
    if (!Object.keys(vscope).length) return { ok: false, message: 'No teachers found for ' + viewAs + '.' };
    return { ok: true, isAdmin: false, viewingAs: true, realAdmin: true, role: ROLES.TP_MANAGER, email: vm.email, name: vm.name,
      actor: String(s.username || 'Admin') + ' (as ' + vm.name + ')', scope: vscope };
  }
  if (isAdmin) return { ok: true, isAdmin: true, role: role, email: email, name: String(s.username || ''), actor: String(s.username || ''), scope: null };
  if (role !== ROLES.TP_MANAGER) return { ok: false, message: 'The TP Workspace is only available to TP managers and admins.' };
  var mgr = _tpwFindManagerByEmail(email);
  if (!mgr) return { ok: false, message: 'Your TP manager access is not active. Please contact the admin.' };
  return { ok: true, isAdmin: false, role: role, email: email, name: mgr.name, actor: mgr.name, scope: _tpwScopeFor(mgr) };
}

function _tpwCanSee(ctx, teacherName) { return ctx.isAdmin || !!(ctx.scope && ctx.scope[_tpwNorm(teacherName)]); }

// ───────────────────────────── progress lookup (Teacher Courses sheet) ─────────────────────────────
// One read of the sheet, then lookups in memory. Returns f(teacher, course) -> { teacherFound, courseFound, text, pct }
function _tpwProgressLookup() {
  var data = _getCachedSheetData(CONFIG.SHEETS.TEACHER_COURSES);
  var headerIdx = 0;
  for (var i = 0; i < Math.min(data.length, 10); i++) { if (String(data[i][0]).trim().toLowerCase() === 'teacher') { headerIdx = i; break; } }
  var headers = data[headerIdx] || [];
  var rows = data.slice(headerIdx + 1);
  var colByName = {};
  for (var c = 4; c < headers.length; c++) {
    var h = String(headers[c] || '').trim();
    if (h) colByName[normalizeCourseNameForMatch(h)] = c;
  }
  var rowCache = {};
  return function(teacher, course) {
    var key = _tpwNorm(teacher);
    if (!(key in rowCache)) { try { rowCache[key] = _findTeacherCoursesRow(rows, teacher); } catch (e) { rowCache[key] = null; } }
    var row = rowCache[key];
    if (!row) return { teacherFound: false, courseFound: false, text: '', pct: 0 };
    var cn = normalizeCourseNameForMatch(course);
    var ci = colByName[cn];
    if (ci === undefined) { try { ci = colByName[normalizeCourseNameForMatch(canonicalCourseName(course))]; } catch (e2) {} }
    if (ci === undefined) return { teacherFound: true, courseFound: false, text: 'Not onboarded', pct: 0 };
    var txt = String(row[ci] || '').trim();
    if (!txt || /not\s*onboarded/i.test(txt)) return { teacherFound: true, courseFound: true, text: 'Not onboarded', pct: 0 };
    return { teacherFound: true, courseFound: true, text: txt, pct: _upskillPct(txt) };
  };
}

// Pod and course level, read from whichever of these sheets has a "pod" column. Blank when none does.
function _tpwPodMap() {
  var map = {};
  var tryOne = function(sheetName, sheetId, headerRow, nameHeader) {
    try {
      var d = _getCachedSheetData(sheetName, sheetId);
      if (!d || d.length <= headerRow + 1) return;
      var hdr = d[headerRow].map(function(h) { return String(h || '').trim(); });
      var nc = hdr.indexOf(nameHeader);
      var pc = -1;
      hdr.forEach(function(h, i) { if (pc < 0 && /\bpod\b/i.test(h)) pc = i; });
      if (nc < 0 || pc < 0) return;
      for (var r = headerRow + 1; r < d.length; r++) {
        var n = _tpwNorm(d[r][nc]), p = String(d[r][pc] || '').trim();
        if (n && p && !map[n]) map[n] = p;
      }
    } catch (e) {}
  };
  tryOne(CONFIG.SHEETS.PERSONA_DATA, CONFIG.PERSONA_SHEET_ID, 1, 'Teacher Name');
  tryOne('Teacher Persona Mapping', undefined, 0, 'Teacher Name');
  tryOne(CONFIG.SHEETS.TEACHER_DATA, undefined, 0, 'Name');
  return map;
}

function _tpwCourseTags() {
  var map = {};
  try {
    var d = _getCachedSheetData(CONFIG.SHEETS.COURSE_NAME);
    for (var r = 1; r < d.length; r++) {
      var n = normalizeCourseNameForMatch(d[r][0]);
      if (n) map[n] = String(d[r][2] || '').trim();
    }
  } catch (e) {}
  return map;
}

// ───────────────────────────── request rows ─────────────────────────────
function _tpwRowToReq(vals, rowNum) {
  var g = function(c) { return vals[c - 1]; };
  return {
    row: rowNum, id: String(g(TPC.ID) || ''), teacher: String(g(TPC.TEACHER) || '').trim(), manager: String(g(TPC.MANAGER) || '').trim(),
    jlids: String(g(TPC.JLIDS) || '').split(/[,\s]+/).filter(Boolean), learners: String(g(TPC.LEARNERS) || ''),
    course: String(g(TPC.COURSE) || '').trim(), hsTask: String(g(TPC.HS_TASK) || ''), status: String(g(TPC.STATUS) || ''),
    createdAt: _tpwIso(g(TPC.CREATED)), completedAt: _tpwIso(g(TPC.COMPLETED)), notes: String(g(TPC.NOTES) || ''),
    tier: String(g(TPC.TIER) || ''), pctAtRequest: g(TPC.PCT_REQ), source: String(g(TPC.SOURCE) || ''),
    startedAt: _tpwIso(g(TPC.STARTED)), doneMarkedAt: _tpwIso(g(TPC.DONE_MARKED)), verifiedAt: _tpwIso(g(TPC.VERIFIED)),
    lastCheckAt: _tpwIso(g(TPC.LAST_CHECK)), lastCheckPct: g(TPC.LAST_PCT), reasonCode: String(g(TPC.REASON_CODE) || ''), reasonNote: String(g(TPC.REASON_NOTE) || ''),
    lastReminderAt: _tpwIso(g(TPC.LAST_REMINDER)), reminders: Number(g(TPC.REMINDERS) || 0), requestedBy: String(g(TPC.REQUESTED_BY) || ''),
    learnerCount: Number(g(TPC.LEARNER_COUNT) || 0), rejectedCount: Number(g(TPC.REJECTED_COUNT) || 0), updatedAt: _tpwIso(g(TPC.UPDATED)),
    _version: String(g(TPC.VERSION) || '')
  };
}

function _tpwReadRequests() {
  var sh = _tpwTracker();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, TPW_HEADERS.length).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][TPC.VERSION - 1] || '') !== TPW.VERSION) continue; // older rows are not part of this
    out.push(_tpwRowToReq(vals[i], i + 2));
  }
  return out;
}

function _tpwSet(sh, row, changes) {
  Object.keys(changes).forEach(function(col) { sh.getRange(row, Number(col)).setValue(changes[col]); });
  sh.getRange(row, TPC.UPDATED).setValue(new Date());
}

function _tpwManagerForTeacher(teacher, dir) {
  var key = _tpwNorm(teacher), found = null;
  (dir || _tpwTeacherDirectory()).forEach(function(t) { if (!found && _tpwNorm(t.name) === key) found = t; });
  if (!found) { try { var rk = _tpwNorm(resolveTeacherName(teacher)); (dir || []).forEach(function(t) { if (!found && _tpwNorm(t.name) === rk) found = t; }); } catch (e) {} }
  return found;
}

/**
 * Adds an upskilling requirement. Same teacher + course already open -> the learner is added to that request instead of a duplicate.
 * opts: { teacher, course, tier, pct, jlid, learner, source, requestedBy, notes }
 */
function tpwCreateRequest(opts) {
  return _tpwWithLock(function() {
    var teacher = String(opts.teacher || '').trim(), course = String(opts.course || '').trim();
    if (!teacher || !course) return { success: false, message: 'Teacher and course are required.' };
    var sh = _tpwTracker();
    var existing = _tpwReadRequests().filter(function(r) {
      return TPW.OPEN.indexOf(r.status) > -1 && _tpwNorm(r.teacher) === _tpwNorm(teacher) && normalizeCourseNameForMatch(r.course) === normalizeCourseNameForMatch(course)
        && ((r.source === 'Test') === (opts.source === 'Test')); // test requests never merge into real ones
    })[0];
    var jlid = String(opts.jlid || '').trim();
    if (existing) {
      var ch = {};
      if (jlid && existing.jlids.indexOf(jlid) === -1) {
        ch[TPC.JLIDS] = existing.jlids.concat([jlid]).join(', ');
        ch[TPC.LEARNERS] = (existing.learners ? existing.learners + '; ' : '') + (opts.learner || jlid);
        ch[TPC.LEARNER_COUNT] = existing.learnerCount + 1;
      }
      if (opts.tier === 'upskill' && existing.tier === 'fast') ch[TPC.TIER] = 'upskill';
      if (Object.keys(ch).length) _tpwSet(sh, existing.row, ch);
      _tpwLog(existing, opts.requestedBy || 'system', '', 'Requirement repeated', '', existing.status, jlid ? 'Another learner waiting: ' + jlid : 'Raised again');
      return { success: true, id: existing.id, duplicate: true };
    }
    var dir = _tpwTeacherDirectory();
    var t = _tpwManagerForTeacher(teacher, dir);
    var manager = (t && t.manager) || String(opts.manager || '');
    var id = Utilities.getUuid();
    var now = new Date();
    var row = new Array(TPW_HEADERS.length);
    for (var i = 0; i < row.length; i++) row[i] = '';
    row[TPC.ID - 1] = id; row[TPC.TEACHER - 1] = teacher; row[TPC.MANAGER - 1] = manager;
    row[TPC.JLIDS - 1] = jlid; row[TPC.LEARNERS - 1] = String(opts.learner || ''); row[TPC.COURSE - 1] = course;
    row[TPC.STATUS - 1] = 'Requested'; row[TPC.CREATED - 1] = now; row[TPC.NOTES - 1] = String(opts.notes || '');
    row[TPC.TIER - 1] = opts.tier === 'fast' ? 'fast' : 'upskill'; row[TPC.PCT_REQ - 1] = (opts.pct === undefined || opts.pct === null) ? '' : opts.pct;
    row[TPC.SOURCE - 1] = String(opts.source || 'Manual'); row[TPC.REQUESTED_BY - 1] = String(opts.requestedBy || 'system');
    row[TPC.LEARNER_COUNT - 1] = jlid ? 1 : 0; row[TPC.REMINDERS - 1] = 0; row[TPC.REJECTED_COUNT - 1] = 0;
    row[TPC.UPDATED - 1] = now; row[TPC.VERSION - 1] = TPW.VERSION;
    sh.appendRow(row);
    var req = { id: id, teacher: teacher, course: course };
    _tpwLog(req, opts.requestedBy || 'system', '', 'Requested', '', 'Requested', (opts.source || 'Manual') + (opts.tier === 'fast' ? ' · fast track' : ' · upskilling required'));
    return { success: true, id: id, duplicate: false };
  });
}

// Called by the ticket note: every course that is not fully upskilled becomes a request on the TP manager's dashboard.
function tpwCreateRequestsFromGaps(info) {
  var made = 0;
  try {
    (info.gaps || []).forEach(function(g) {
      var r = tpwCreateRequest({ teacher: info.teacher, course: g.label, tier: g.tier, pct: g.pct, jlid: info.jlid, learner: info.learner,
        source: 'Ticket note', requestedBy: 'system', manager: info.manager });
      if (r && r.success) made++;
    });
  } catch (e) { Logger.log('[tpw] tpwCreateRequestsFromGaps: ' + e.message); }
  return made;
}

// ───────────────────────────── dashboard data ─────────────────────────────
function _tpwDecorate(reqs, ctxDir) {
  var lookup = _tpwProgressLookup();
  var pods = _tpwPodMap();
  var tags = _tpwCourseTags();
  var reasonByCode = {};
  TPW_REASONS.forEach(function(r) { reasonByCode[r.code] = r.label; });
  return reqs.map(function(r) {
    var live = lookup(r.teacher, r.course);
    var pod = pods[_tpwNorm(r.teacher)] || '';
    var tag = tags[normalizeCourseNameForMatch(r.course)] || '';
    var hint = '';
    if (/foundation/i.test(pod) && /pro/i.test(tag)) hint = 'Foundation pod teacher and a Pro course — this is usually not allowed.';
    var created = _tpwHoursSince(r.createdAt);
    return {
      id: r.id, teacher: r.teacher, manager: r.manager, course: r.course, jlids: r.jlids, learners: r.learners, learnerCount: r.learnerCount,
      status: r.status, tier: r.tier, pctAtRequest: r.pctAtRequest, source: r.source, isTest: r.source === 'Test', hsTask: r.hsTask,
      createdAt: r.createdAt, startedAt: r.startedAt, doneMarkedAt: r.doneMarkedAt, verifiedAt: r.verifiedAt, completedAt: r.completedAt,
      lastCheckAt: r.lastCheckAt, lastCheckPct: r.lastCheckPct, rejectedCount: r.rejectedCount, notes: r.notes,
      reasonCode: r.reasonCode, reasonLabel: reasonByCode[r.reasonCode] || '', reasonNote: r.reasonNote, requestedBy: r.requestedBy,
      ageDays: created === null ? 0 : Math.floor(created / 24),
      daysSinceStart: r.startedAt ? _tpwDaysSince(r.startedAt) : null,
      hoursSinceDone: r.doneMarkedAt ? Math.floor(_tpwHoursSince(r.doneMarkedAt)) : null,
      liveText: live.text, livePct: live.pct, liveTeacherFound: live.teacherFound,
      pod: pod, courseTag: tag, hint: hint
    };
  });
}

function tpwBootstrap(token, viewAs) {
  try {
    var ctx = _tpwCtx(token, viewAs);
    if (!ctx.ok) return { success: false, message: ctx.message };
    var out = { success: true, isAdmin: ctx.isAdmin, name: ctx.name, email: ctx.email, reasons: TPW_REASONS, passPct: TPW.PASS_PCT, fastPct: TPW.FAST_PCT,
      rejectAfterHours: TPW.REJECT_AFTER_HOURS, noStartDays: TPW.NO_START_TASK_DAYS };
    if (ctx.isAdmin) {
      var dir = _tpwTeacherDirectory();
      out.managers = _tpwAllManagerNames(dir);
      out.teachers = dir.filter(function(t) { return ['Active', 'EWS', 'Friendly'].indexOf(t.status) > -1; }).map(function(t) { return { name: t.name, manager: t.manager }; })
        .sort(function(a, b) { return a.name.localeCompare(b.name); });
      try { out.courses = getCourseNames(); } catch (e) { out.courses = []; }
    } else {
      out.teacherCount = Object.keys(ctx.scope).length;
    }
    if (ctx.viewingAs) { out.viewingAs = ctx.name; out.realAdmin = true; out.managers = _tpwAllManagerNames(); }
    if (ctx.isAdmin) out.realAdmin = true;
    return out;
  } catch (e) {
    Logger.log('[tpwBootstrap] ' + e.message);
    return { success: false, message: e.message };
  }
}

function tpwGetRequests(token, viewAs) {
  try {
    var ctx = _tpwCtx(token, viewAs);
    if (!ctx.ok) return { success: false, message: ctx.message };
    var reqs = _tpwReadRequests().filter(function(r) { return _tpwCanSee(ctx, r.teacher); });
    var list = _tpwDecorate(reqs);
    return { success: true, requests: list, serverTime: new Date().toISOString() };
  } catch (e) {
    Logger.log('[tpwGetRequests] ' + e.message);
    return { success: false, message: e.message };
  }
}

// Everything a TP needs to decide on one teacher: courses with percentages, audit grade, pod, load.
function tpwTeacherSnapshot(token, teacher, viewAs) {
  try {
    var ctx = _tpwCtx(token, viewAs);
    if (!ctx.ok) return { success: false, message: ctx.message };
    if (!_tpwCanSee(ctx, teacher)) return { success: false, message: 'That teacher is not on your list.' };
    var load = {};
    try { load = getTeacherSpecificLoad(teacher, true) || {}; } catch (e) {}
    var courses = ((load && load.success) ? (load.courses || []) : []).map(function(c) {
      var p = _upskillPct(c.proficiency);
      return { course: c.course, text: c.proficiency, pct: p, tier: _upskillTier(p) };
    }).sort(function(a, b) { return b.pct - a.pct; });
    var pods = _tpwPodMap();
    var info = { success: true, teacher: teacher, pod: pods[_tpwNorm(teacher)] || '', courses: courses, hasLoadData: !!(load && load.success) };
    ['tpManager', 'clsManager', 'status', 'activeLearners', 'lastActivity'].forEach(function(k) { if (load && load[k] !== undefined) info[k] = load[k]; });
    var dir = _tpwTeacherDirectory();
    var t = _tpwManagerForTeacher(teacher, dir);
    if (t) { info.status = info.status || t.status; info.hidden = t.hidden; }
    return info;
  } catch (e) {
    Logger.log('[tpwTeacherSnapshot] ' + e.message);
    return { success: false, message: e.message };
  }
}

function tpwGetHistory(token, id, viewAs) {
  try {
    var ctx = _tpwCtx(token, viewAs);
    if (!ctx.ok) return { success: false, message: ctx.message };
    var req = _tpwReadRequests().filter(function(r) { return r.id === id; })[0];
    if (!req || !_tpwCanSee(ctx, req.teacher)) return { success: false, message: 'Request not found.' };
    var sh = _tpwActivitySheet();
    var vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues() : [];
    var items = vals.filter(function(v) { return String(v[1]) === id; }).map(function(v) {
      return { at: _tpwIso(v[0]), by: String(v[4] || ''), action: String(v[6] || ''), from: String(v[7] || ''), to: String(v[8] || ''), detail: String(v[9] || '') };
    }).reverse();
    return { success: true, items: items };
  } catch (e) { return { success: false, message: e.message }; }
}

// ───────────────────────────── actions ─────────────────────────────
function _tpwCheckNow(req) {
  var f = _tpwProgressLookup();
  return f(req.teacher, req.course);
}

function tpwAction(token, id, action, payload, viewAs) {
  try {
    var ctx = _tpwCtx(token, viewAs);
    if (!ctx.ok) return { success: false, message: ctx.message };
    payload = payload || {};
    return _tpwWithLock(function() {
      var sh = _tpwTracker();
      var req = _tpwReadRequests().filter(function(r) { return r.id === id; })[0];
      if (!req) return { success: false, message: 'Request not found.' };
      if (!_tpwCanSee(ctx, req.teacher)) return { success: false, message: 'That teacher is not on your list.' };
      var by = ctx.actor || ctx.name || ctx.email, now = new Date();
      var from = req.status;
      var note = String(payload.note || '').trim();

      if (action === 'start') {
        if (['Requested', 'Rejected'].indexOf(from) === -1) return { success: false, message: 'This request is already ' + from + '.' };
        var ch = {}; ch[TPC.STATUS] = 'Started'; ch[TPC.STARTED] = req.startedAt ? _tpwToDate(req.startedAt) : now;
        if (note) ch[TPC.NOTES] = (req.notes ? req.notes + ' | ' : '') + by + ': ' + note;
        _tpwSet(sh, req.row, ch);
        _tpwLog(req, by, ctx.role, 'Upskilling started', from, 'Started', note);
        return { success: true, message: 'Marked as upskilling started.' };
      }

      if (action === 'done') {
        if (['Started', 'Requested', 'Rejected'].indexOf(from) === -1) return { success: false, message: 'This request is ' + from + '.' };
        var chk = _tpwCheckNow(req);
        var c2 = {};
        c2[TPC.LAST_CHECK] = now; c2[TPC.LAST_PCT] = chk.pct;
        if (!req.startedAt) c2[TPC.STARTED] = now;
        if (chk.teacherFound && chk.pct >= TPW.PASS_PCT) {
          c2[TPC.STATUS] = 'Verified'; c2[TPC.DONE_MARKED] = now; c2[TPC.VERIFIED] = now; c2[TPC.COMPLETED] = now;
          if (note) c2[TPC.NOTES] = (req.notes ? req.notes + ' | ' : '') + by + ': ' + note;
          _tpwSet(sh, req.row, c2);
          _tpwLog(req, by, ctx.role, 'Marked done — verified at once', from, 'Verified', 'Sheet shows ' + chk.text);
          return { success: true, message: 'Verified — the sheet already shows ' + chk.text + '.', verified: true };
        }
        c2[TPC.STATUS] = 'Verifying'; c2[TPC.DONE_MARKED] = now;
        if (note) c2[TPC.NOTES] = (req.notes ? req.notes + ' | ' : '') + by + ': ' + note;
        _tpwSet(sh, req.row, c2);
        _tpwLog(req, by, ctx.role, 'Marked done — waiting for data', from, 'Verifying', 'Sheet shows ' + (chk.teacherFound ? chk.text : 'no row for this teacher') + '; will recheck for up to ' + TPW.REJECT_AFTER_HOURS + ' h');
        return { success: true, message: 'Noted. The sheet shows ' + (chk.teacherFound ? chk.text : 'nothing yet') + '. Moodle can take 24–48 hours to reach the sheet — we recheck automatically and will reject it if it is still below ' + TPW.PASS_PCT + '% after ' + TPW.REJECT_AFTER_HOURS + ' hours.', verified: false };
      }

      if (action === 'notpossible') {
        if (TPW.OPEN.indexOf(from) === -1) return { success: false, message: 'This request is already ' + from + '.' };
        var code = String(payload.reasonCode || '');
        var reason = TPW_REASONS.filter(function(r) { return r.code === code; })[0];
        if (!reason) return { success: false, message: 'Pick a reason.' };
        if (code === 'OTHER' && !note) return { success: false, message: 'Please add a note for "Other".' };
        var c3 = {}; c3[TPC.STATUS] = 'Not Possible'; c3[TPC.REASON_CODE] = code; c3[TPC.REASON_NOTE] = note; c3[TPC.COMPLETED] = now;
        _tpwSet(sh, req.row, c3);
        _tpwLog(req, by, ctx.role, 'Marked not possible', from, 'Not Possible', reason.label + (note ? ' — ' + note : ''));
        return { success: true, message: 'Recorded as not possible. The admin team can see the reason.' };
      }

      if (action === 'reopen') {
        if (['Not Possible', 'Cancelled', 'Verified'].indexOf(from) === -1) return { success: false, message: 'This request is already open.' };
        if (!ctx.isAdmin && from !== 'Not Possible') return { success: false, message: 'Only an admin can reopen this.' };
        var c4 = {}; c4[TPC.STATUS] = 'Requested'; c4[TPC.COMPLETED] = ''; c4[TPC.VERIFIED] = ''; c4[TPC.REASON_CODE] = ''; c4[TPC.REASON_NOTE] = ''; c4[TPC.CREATED] = now;
        _tpwSet(sh, req.row, c4);
        _tpwLog(req, by, ctx.role, 'Reopened', from, 'Requested', note);
        return { success: true, message: 'Reopened.' };
      }

      if (action === 'note') {
        if (!note) return { success: false, message: 'Write a note first.' };
        var c5 = {}; c5[TPC.NOTES] = (req.notes ? req.notes + ' | ' : '') + by + ': ' + note;
        _tpwSet(sh, req.row, c5);
        _tpwLog(req, by, ctx.role, 'Note', from, from, note);
        return { success: true, message: 'Note saved.' };
      }

      // admin-only actions
      if (!ctx.isAdmin) return { success: false, message: 'Only an admin can do that.' };
      if (action === 'verify') {
        var c6 = {}; c6[TPC.STATUS] = 'Verified'; c6[TPC.VERIFIED] = now; c6[TPC.COMPLETED] = now;
        if (!req.doneMarkedAt) c6[TPC.DONE_MARKED] = now;
        _tpwSet(sh, req.row, c6);
        _tpwLog(req, by, ctx.role, 'Verified manually', from, 'Verified', note);
        return { success: true, message: 'Verified.' };
      }
      if (action === 'cancel') {
        var c7 = {}; c7[TPC.STATUS] = 'Cancelled'; c7[TPC.COMPLETED] = now;
        _tpwSet(sh, req.row, c7);
        _tpwLog(req, by, ctx.role, 'Cancelled', from, 'Cancelled', note);
        return { success: true, message: 'Cancelled.' };
      }
      return { success: false, message: 'Unknown action.' };
    });
  } catch (e) {
    Logger.log('[tpwAction] ' + e.message);
    return { success: false, message: e.message };
  }
}

// Admin: raise a requirement for one or many teachers on one course.
function tpwAdminCreate(token, payload) {
  try {
    var ctx = _tpwCtx(token);
    if (!ctx.ok || !ctx.isAdmin) return { success: false, message: 'Only an admin can create requirements.' };
    var teachers = (payload.teachers || []).map(function(t) { return String(t || '').trim(); }).filter(Boolean);
    var course = String(payload.course || '').trim();
    if (!teachers.length || !course) return { success: false, message: 'Pick at least one teacher and a course.' };
    var lookup = _tpwProgressLookup();
    var made = 0, dup = 0, skipped = [];
    teachers.forEach(function(t) {
      var p = lookup(t, course);
      if (!payload.test && p.teacherFound && p.pct >= TPW.PASS_PCT) { skipped.push(t + ' (already ' + p.text + ')'); return; }
      var r = tpwCreateRequest({ teacher: t, course: course, tier: p.pct >= TPW.FAST_PCT ? 'fast' : 'upskill', pct: p.pct, source: payload.test ? 'Test' : 'Admin requirement',
        requestedBy: ctx.name || ctx.email, notes: String(payload.note || '').trim() });
      if (r && r.success) { if (r.duplicate) dup++; else made++; }
    });
    return { success: true, created: made, duplicates: dup, skipped: skipped,
      message: made + ' requirement' + (made === 1 ? '' : 's') + ' created' + (dup ? ', ' + dup + ' already open' : '') + (skipped.length ? ', ' + skipped.length + ' skipped (already upskilled)' : '') + '.' };
  } catch (e) { return { success: false, message: e.message }; }
}

// ───────────────────────────── TP managers (admin) ─────────────────────────────
function tpwListManagers(token) {
  try {
    var ctx = _tpwCtx(token);
    if (!ctx.ok || !ctx.isAdmin) return { success: false, message: 'Admin only.' };
    var mgrs = _tpwReadManagers();
    var dir = _tpwTeacherDirectory();
    var reqs = _tpwReadRequests().filter(function(r) { return r.source !== 'Test'; });
    var usage = _usageIndex();
    var iso = function(d) { return d ? d.toISOString() : ''; };
    var hsFor = function(name) { try { return getHubSpotOwnerIdByName(name) || ''; } catch (e) { return ''; } };
    mgrs.forEach(function(m) {
      m.teacherCount = Object.keys(_tpwScopeFor(m, dir)).length;
      var mine = reqs.filter(function(r) { return _tpwNorm(r.manager) === _tpwNorm(m.name); });
      m.open = mine.filter(function(r) { return TPW.OPEN.indexOf(r.status) > -1; }).length;
      m.needsAction = mine.filter(function(r) { return r.status === 'Requested' || r.status === 'Rejected'; }).length;
      var u = _usageFor(usage, m.email, m.name);
      if (u) { m.firstLogin = m.firstLogin || iso(u.first); m.lastLogin = iso(u.last) || m.lastLogin; m.logins = u.total; } else m.logins = 0;
      m.hubspotDetected = hsFor(m.name);
    });
    // every manager name the app knows about that is not in the TP Managers sheet yet (Teacher Data, Persona sheet, the old hard-coded list)
    var listed = {};
    mgrs.forEach(function(m) { listed[_tpwNorm(m.name)] = true; });
    var unlisted = {};
    _tpwAllManagerNames(dir).forEach(function(n) {
      var k = _tpwNorm(n);
      if (!listed[k]) unlisted[k] = { name: n, email: '', teacherCount: 0, hubspotDetected: hsFor(n) };
    });
    dir.forEach(function(t) {
      var k = _tpwNorm(t.manager);
      if (k && unlisted[k]) { unlisted[k].teacherCount++; if (!unlisted[k].email && t.tpEmail) unlisted[k].email = t.tpEmail; }
    });
    var ul = Object.keys(unlisted).map(function(k) { return unlisted[k]; }).sort(function(a, b) { return b.teacherCount - a.teacherCount; });
    return { success: true, managers: mgrs, unlisted: ul, automation: tpwAutomationStatus() };
  } catch (e) { return { success: false, message: e.message }; }
}

function _tpwEnsureProfile(email, name, active) {
  try {
    var sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.USER_PROFILES);
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim().toLowerCase().replace(/\s/g, ''); });
    var ei = headers.indexOf('email'), ri = headers.indexOf('role'), ai = headers.indexOf('isactive');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][ei] || '').toLowerCase().trim() === email) {
        var cur = String(data[i][ri] || '');
        // never downgrade an admin to TP Manager
        if (cur !== ROLES.SUPER_ADMIN && cur !== ROLES.ADMIN) sheet.getRange(i + 1, ri + 1).setValue(ROLES.TP_MANAGER);
        if (ai > -1) sheet.getRange(i + 1, ai + 1).setValue(!!active);
        delete _sheetDataCache[CONFIG.MIGRATION_SHEET_ID + '_' + CONFIG.SHEETS.USER_PROFILES];
        return;
      }
    }
    if (active) {
      sheet.appendRow([name, '', ROLES.TP_MANAGER, email, true, new Date(), new Date(), '', '', '']);
      delete _sheetDataCache[CONFIG.MIGRATION_SHEET_ID + '_' + CONFIG.SHEETS.USER_PROFILES];
    }
  } catch (e) { Logger.log('[tpw] _tpwEnsureProfile: ' + e.message); }
}

// Add or update a TP manager. active=true also creates / updates their login profile (role TP Manager).
function tpwSaveManager(token, m) {
  try {
    var ctx = _tpwCtx(token);
    if (!ctx.ok || !ctx.isAdmin) return { success: false, message: 'Admin only.' };
    var name = String(m.name || '').trim(), email = String(m.email || '').trim().toLowerCase();
    if (!name) return { success: false, message: 'Name is required (it must match the Manager column in Teacher Data).' };
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { success: false, message: 'Enter a valid email address.' };
    var active = !!m.active;
    return _tpwWithLock(function() {
      var sh = _tpwManagersSheet();
      var list = _tpwReadManagers();
      var hit = list.filter(function(x) { return x.email === email || _tpwNorm(x.name) === _tpwNorm(name); })[0];
      var hsId = String(m.hubspotId || '').trim();
      if (hit) {
        sh.getRange(hit.row, 1, 1, 3).setValues([[name, email, active ? 'Yes' : 'No']]);
        if (m.hubspotId !== undefined) sh.getRange(hit.row, 8).setValue(hsId);
      } else {
        sh.appendRow([name, email, active ? 'Yes' : 'No', '', ctx.name || ctx.email, new Date(), '', hsId, '']);
      }
      _tpwEnsureProfile(email, name, active);
      return { success: true, message: name + (active ? ' can now sign in to the TP Workspace.' : ' is switched off.') };
    });
  } catch (e) { return { success: false, message: e.message }; }
}

function tpwInviteManager(token, email) {
  try {
    var ctx = _tpwCtx(token);
    if (!ctx.ok || !ctx.isAdmin) return { success: false, message: 'Admin only.' };
    email = String(email || '').trim().toLowerCase();
    var mgr = _tpwReadManagers().filter(function(x) { return x.email === email; })[0];
    if (!mgr) return { success: false, message: 'Add the manager first.' };
    if (!mgr.active) return { success: false, message: 'Switch the manager on before sending the invite.' };
    var dir = _tpwTeacherDirectory();
    var count = Object.keys(_tpwScopeFor(mgr, dir)).length;
    var html = _tpwWrapEmail('Your TP Workspace is ready',
      '<p style="margin:0 0 12px;">Hi ' + _tpwEsc(mgr.name.split(' ')[0]) + ',</p>'
      + '<p style="margin:0 0 12px;">From now on, upskilling requirements for your teachers will not only arrive as an email — you can see them in one place and act on them: <b>start upskilling, mark it done, or tell us why it is not possible</b>.</p>'
      + '<p style="margin:0 0 12px;">You will see <b>' + count + ' teachers</b> (the ones you manage), their Teacher Intelligence and Persona details, and every upskilling request raised for them.</p>'
      + '<p style="margin:0 0 12px;">Sign in with <b>' + _tpwEsc(mgr.email) + '</b> using “Sign in with Google”.</p>',
      'Open TP Workspace', _tpwAppUrl());
    sendTrackedEmail({ to: mgr.email, subject: 'Your JetLearn TP Workspace is ready', htmlBody: html, jlid: 'TP-INVITE' });
    _tpwManagersSheet().getRange(mgr.row, 4).setValue(new Date());
    return { success: true, message: 'Invite sent to ' + mgr.email + '.' };
  } catch (e) { return { success: false, message: e.message }; }
}

// ───────────────────────────── scoped versions of the existing Teacher Intelligence / Persona calls ─────────────────────────────
function tpwTeacherDetailsScoped(token, viewAs) {
  var ctx = _tpwCtx(token, viewAs);
  if (!ctx.ok) throw new Error(ctx.message);
  var all = getTeacherDetailsForTable();
  if (ctx.isAdmin) return all;
  return (all || []).filter(function(t) { return ctx.scope[_tpwNorm(t.name)]; });
}

function tpwSetVisibilityScoped(token, teacherName, isHidden, viewAs) {
  var ctx = _tpwCtx(token, viewAs);
  if (!ctx.ok) return { success: false, message: ctx.message };
  if (!_tpwCanSee(ctx, teacherName)) return { success: false, message: 'That teacher is not on your list.' };
  var r = setTeacherVisibility(teacherName, isHidden);
  try { _tpwLog({ id: '', teacher: teacherName, course: '' }, ctx.actor || ctx.name || ctx.email, ctx.role, isHidden ? 'Teacher hidden in search' : 'Teacher visible in search', '', '', ''); } catch (e) {}
  return r;
}

// ───────────────────────────── automation: verification ─────────────────────────────
// Runs every few hours. Rechecks every "Verifying" request against the Teacher Courses sheet.
function tpwVerifyJob() {
  var summary = { verified: 0, rejected: 0, waiting: 0 };
  try {
    var lookup = _tpwProgressLookup();
    var sh = _tpwTracker();
    var verifying = _tpwReadRequests().filter(function(r) { return r.status === 'Verifying'; });
    var rejectedNow = {};
    verifying.forEach(function(r) {
      var chk = lookup(r.teacher, r.course);
      var now = new Date();
      var hrs = _tpwHoursSince(r.doneMarkedAt);
      if (chk.teacherFound && chk.pct >= TPW.PASS_PCT) {
        var ch = {}; ch[TPC.STATUS] = 'Verified'; ch[TPC.VERIFIED] = now; ch[TPC.COMPLETED] = now; ch[TPC.LAST_CHECK] = now; ch[TPC.LAST_PCT] = chk.pct;
        _tpwSet(sh, r.row, ch);
        _tpwLog(r, 'system', '', 'Verified on data', 'Verifying', 'Verified', 'Sheet shows ' + chk.text);
        summary.verified++;
      } else if (hrs !== null && hrs >= TPW.REJECT_AFTER_HOURS) {
        var c2 = {}; c2[TPC.STATUS] = 'Rejected'; c2[TPC.LAST_CHECK] = now; c2[TPC.LAST_PCT] = chk.pct; c2[TPC.REJECTED_COUNT] = r.rejectedCount + 1;
        _tpwSet(sh, r.row, c2);
        _tpwLog(r, 'system', '', 'Rejected — not on the data', 'Verifying', 'Rejected', 'Marked done ' + Math.floor(hrs) + ' h ago, sheet shows ' + (chk.teacherFound ? chk.text : 'no row for this teacher'));
        r.status = 'Rejected'; r.lastCheckPct = chk.pct; r._text = chk.teacherFound ? chk.text : 'no row for this teacher';
        var key = _tpwNorm(r.manager);
        if (r.source !== 'Test') (rejectedNow[key] = rejectedNow[key] || []).push(r);
        summary.rejected++;
      } else {
        var c3 = {}; c3[TPC.LAST_CHECK] = now; c3[TPC.LAST_PCT] = chk.pct;
        _tpwSet(sh, r.row, c3);
        summary.waiting++;
      }
    });
    Object.keys(rejectedNow).forEach(function(k) { _tpwEmailRejections(rejectedNow[k]); });
  } catch (e) { Logger.log('[tpwVerifyJob] ' + e.message); }
  Logger.log('[tpwVerifyJob] ' + JSON.stringify(summary));
  return summary;
}

function _tpwManagerEmailFor(managerName) {
  var m = _tpwReadManagers().filter(function(x) { return _tpwNorm(x.name) === _tpwNorm(managerName) && x.email; })[0];
  if (m) return { name: m.name, email: m.email };
  var dir = _tpwTeacherDirectory();
  for (var i = 0; i < dir.length; i++) { if (_tpwNorm(dir[i].manager) === _tpwNorm(managerName) && dir[i].tpEmail) return { name: managerName, email: dir[i].tpEmail }; }
  return null;
}

function _tpwEmailRejections(reqs) {
  try {
    if (!reqs.length) return;
    var to = _tpwManagerEmailFor(reqs[0].manager);
    if (!to) { Logger.log('[tpw] no email for manager ' + reqs[0].manager); return; }
    var rows = reqs.map(function(r) {
      return '<tr><td style="padding:8px 10px;border-top:1px solid #eef0f4;">' + _tpwEsc(r.teacher) + '</td><td style="padding:8px 10px;border-top:1px solid #eef0f4;">' + _tpwEsc(r.course)
        + '</td><td style="padding:8px 10px;border-top:1px solid #eef0f4;color:#b91c1c;">' + _tpwEsc(r._text || '') + '</td></tr>';
    }).join('');
    var body = '<p style="margin:0 0 12px;">Hi ' + _tpwEsc(to.name.split(' ')[0]) + ',</p>'
      + '<p style="margin:0 0 12px;">You marked ' + (reqs.length === 1 ? 'this teacher' : 'these teachers') + ' as upskilled, but <b>' + TPW.REJECT_AFTER_HOURS + ' hours later the data still does not show it</b> (we need ' + TPW.PASS_PCT + '% or more). '
      + 'We gave Moodle time to catch up — it has not. Please check Moodle, and mark it done again once it shows.</p>'
      + '<table style="width:100%;border-collapse:collapse;font-size:14px;margin:0 0 14px;"><tr style="background:#f6f7fb;text-align:left;"><th style="padding:8px 10px;">Teacher</th><th style="padding:8px 10px;">Course</th><th style="padding:8px 10px;">Data shows</th></tr>' + rows + '</table>';
    sendTrackedEmail({ to: to.email, subject: '↩️ Upskilling not on the data yet — ' + reqs.length + ' sent back to you', htmlBody: _tpwWrapEmail('Sent back: not on the data yet', body, 'Open TP Workspace', _tpwAppUrl()), jlid: 'TP-REJECT' });
  } catch (e) { Logger.log('[tpw] _tpwEmailRejections: ' + e.message); }
}

// ───────────────────────────── automation: reminders + 14-day task ─────────────────────────────
function _tpwWrapEmail(title, innerHtml, ctaLabel, ctaUrl) {
  return '<html><body style="margin:0;padding:0;background:#f3f4f8;font-family:Segoe UI,Arial,sans-serif;color:#1f2937;">'
    + '<div style="max-width:600px;margin:0 auto;padding:24px 16px;">'
    + '<div style="background:#ffffff;border-radius:14px;overflow:hidden;border:1px solid #e5e7ef;">'
    + '<div style="background:#1e1b4b;color:#ffffff;padding:18px 24px;font-size:18px;font-weight:700;">' + _tpwEsc(title) + '</div>'
    + '<div style="padding:22px 24px;font-size:15px;line-height:1.55;">' + innerHtml
    + (ctaUrl ? '<p style="margin:18px 0 4px;"><a href="' + _tpwEsc(ctaUrl) + '" style="background:#6366f1;color:#ffffff;text-decoration:none;padding:11px 20px;border-radius:8px;font-weight:600;display:inline-block;">' + _tpwEsc(ctaLabel || 'Open') + '</a></p>' : '')
    + '</div></div><p style="text-align:center;color:#9ca3af;font-size:12px;margin:14px 0 0;">JetLearn · TP Workspace</p></div></body></html>';
}

var TPW_DIGEST_VARIANTS = [
  { subject: function(n) { return '🚀 ' + n + ' upskilling request' + (n === 1 ? ' is' : 's are') + ' waiting for you'; },
    title: 'Your teachers are ready to grow — they need a nudge from you',
    lead: function(n, d) { return 'Every request below is one click from moving: <b>Start upskilling</b>. The oldest has been waiting ' + d + ' day' + (d === 1 ? '' : 's') + '.'; } },
  { subject: function(n) { return '⏳ Clock check: ' + n + ' teacher' + (n === 1 ? '' : 's') + ' still waiting on your call'; },
    title: 'Tick-tock — learners are waiting on these courses',
    lead: function(n, d) { return 'Each course here is blocking a migration or a new learner. Oldest: <b>' + d + ' day' + (d === 1 ? '' : 's') + '</b>. A start date today beats a perfect plan next week.'; } },
  { subject: function(n) { return '📣 Utmost attention: upskilling needs your action'; },
    title: 'This one needs your full attention',
    lead: function(n, d) { return 'If upskilling has not started within ' + TPW.NO_START_TASK_DAYS + ' days, a HubSpot task is created in your name. Oldest here: <b>' + d + ' day' + (d === 1 ? '' : 's') + '</b>.'; } },
  { subject: function(n) { return '🧩 Small clicks, big unlock: ' + n + ' upskilling request' + (n === 1 ? '' : 's'); },
    title: 'Three buttons. That is all it takes.',
    lead: function(n, d) { return '<b>Start upskilling</b> · <b>Mark done</b> · or tell us <b>why it is not possible</b>. Anything is better than silence — silence is what blocks the learner.'; } },
  { subject: function(n) { return '🔥 ' + n + ' still open — please close the loop'; },
    title: 'Close the loop on these upskilling requests',
    lead: function(n, d) { return 'We can only plan teacher allocation properly when your dashboard is up to date. Oldest open request: <b>' + d + ' day' + (d === 1 ? '' : 's') + '</b>.'; } },
  { subject: function(n) { return '🏁 Almost there — ' + n + ' upskilling request' + (n === 1 ? '' : 's') + ' need a status'; },
    title: 'Give each request a status and we are done',
    lead: function(n, d) { return 'Started, done, or not possible — pick one for each line below. If a teacher has finished the training, mark it done and we will check the data for you.'; } }
];

function _tpwDigestFor(mgrName, items, count) {
  var v = TPW_DIGEST_VARIANTS[count % TPW_DIGEST_VARIANTS.length];
  var oldest = items.reduce(function(m, r) { return Math.max(m, r._age || 0); }, 0);
  var first = String(mgrName || '').split(' ')[0] || 'there';
  var rows = items.slice().sort(function(a, b) { return (b._age || 0) - (a._age || 0); }).slice(0, 25).map(function(r) {
    var tag = r.status === 'Rejected' ? '<span style="color:#b91c1c;font-weight:600;">Sent back</span>'
      : (r.status === 'Started' ? '<span style="color:#b45309;font-weight:600;">In progress</span>' : (r.tier === 'fast' ? '<span style="color:#b45309;font-weight:600;">Fast track</span>' : '<span style="color:#b91c1c;font-weight:600;">Upskilling</span>'));
    return '<tr><td style="padding:8px 10px;border-top:1px solid #eef0f4;">' + _tpwEsc(r.teacher) + '</td><td style="padding:8px 10px;border-top:1px solid #eef0f4;">' + _tpwEsc(r.course)
      + '</td><td style="padding:8px 10px;border-top:1px solid #eef0f4;">' + tag + '</td><td style="padding:8px 10px;border-top:1px solid #eef0f4;text-align:right;">' + (r._age || 0) + 'd</td></tr>';
  }).join('');
  var more = items.length > 25 ? '<p style="color:#6b7280;margin:0 0 12px;">…and ' + (items.length - 25) + ' more in your dashboard.</p>' : '';
  var body = '<p style="margin:0 0 12px;">Hi ' + _tpwEsc(first) + ',</p><p style="margin:0 0 14px;">' + v.lead(items.length, oldest) + '</p>'
    + '<table style="width:100%;border-collapse:collapse;font-size:14px;margin:0 0 12px;"><tr style="background:#f6f7fb;text-align:left;"><th style="padding:8px 10px;">Teacher</th><th style="padding:8px 10px;">Course</th><th style="padding:8px 10px;">Needs</th><th style="padding:8px 10px;text-align:right;">Waiting</th></tr>' + rows + '</table>' + more;
  return { subject: v.subject(items.length), html: _tpwWrapEmail(v.title, body, 'Open my dashboard', _tpwAppUrl()) };
}

function _tpwCreateHsTask(ownerId, subject, body) {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    if (!token) return '';
    var due = new Date(); due.setDate(due.getDate() + 2);
    var props = { hs_task_subject: subject, hs_task_body: body, hs_task_status: 'NOT_STARTED', hs_task_priority: 'HIGH', hs_task_type: 'TODO', hs_timestamp: due.getTime() };
    if (ownerId) props.hubspot_owner_id = String(ownerId);
    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/tasks', {
      method: 'post', headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ properties: props }), muteHttpExceptions: true
    });
    if (resp.getResponseCode() >= 300) { Logger.log('[tpw] HubSpot task HTTP ' + resp.getResponseCode() + ' ' + resp.getContentText().substring(0, 200)); return ''; }
    var j = JSON.parse(resp.getContentText());
    return j && j.id ? String(j.id) : '';
  } catch (e) { Logger.log('[tpw] _tpwCreateHsTask: ' + e.message); return ''; }
}

// Daily at 9 am (script time zone). Skips Sundays.
function tpwDailyJob() {
  var out = { tasks: 0, digests: 0 };
  try {
    if (new Date().getDay() === 0) return out;
    var sh = _tpwTracker();
    var props = PropertiesService.getScriptProperties();
    var reqs = _tpwReadRequests().filter(function(r) { return TPW.OPEN.indexOf(r.status) > -1; });

    // 1) no start in 14 days -> HubSpot task, once
    reqs.forEach(function(r) {
      if (r.status !== 'Requested' || r.hsTask || r.source === 'Test') return;
      var age = _tpwDaysSince(r.createdAt);
      if (age === null || age < TPW.NO_START_TASK_DAYS) return;
      var ownerId = '';
      try {
        var mrow = _tpwReadManagers().filter(function(x) { return _tpwNorm(x.name) === _tpwNorm(r.manager); })[0];
        ownerId = (mrow && mrow.hubspotId) || getHubSpotOwnerIdByName(r.manager) || '';
      } catch (e) {}
      var id = _tpwCreateHsTask(ownerId,
        'Upskilling not started (' + age + ' days): ' + r.teacher + ' → ' + r.course,
        r.teacher + ' needs upskilling on ' + r.course + (r.learnerCount ? ' (' + r.learnerCount + ' learner' + (r.learnerCount === 1 ? '' : 's') + ' waiting)' : '') + '.\n\nThis requirement has been open for ' + age + ' days with no start. '
        + 'Please open the TP Workspace and mark it Started, Done, or Not possible (with the reason).\n' + _tpwAppUrl());
      if (id) {
        var ch = {}; ch[TPC.HS_TASK] = id; _tpwSet(sh, r.row, ch);
        _tpwLog(r, 'system', '', 'HubSpot task created', 'Requested', 'Requested', age + ' days without a start');
        out.tasks++;
      } else {
        _tpwLog(r, 'system', '', 'HubSpot task failed', 'Requested', 'Requested', 'Could not create the task — will retry tomorrow');
      }
    });

    // 2) reminder email per manager: things that need them (not started, or sent back), at most once every 2 days
    var byMgr = {};
    reqs.forEach(function(r) {
      if ((r.status !== 'Requested' && r.status !== 'Rejected') || r.source === 'Test') return;
      var k = _tpwNorm(r.manager);
      if (!k) return;
      r._age = _tpwDaysSince(r.createdAt) || 0;
      (byMgr[k] = byMgr[k] || []).push(r);
    });
    Object.keys(byMgr).forEach(function(k) {
      var items = byMgr[k];
      var pk = 'TPW_DIGEST_' + k.replace(/[^a-z0-9]/g, '_');
      var lastRaw = props.getProperty(pk) || '';
      var last = lastRaw ? JSON.parse(lastRaw) : { at: 0, n: 0 };
      if (Date.now() - last.at < (TPW.DIGEST_EVERY_DAYS * 24 - 2) * 3600000) return;
      var to = _tpwManagerEmailFor(items[0].manager);
      if (!to) { Logger.log('[tpw] no email for ' + items[0].manager); return; }
      var d = _tpwDigestFor(to.name, items, last.n);
      sendTrackedEmail({ to: to.email, subject: d.subject, htmlBody: d.html, jlid: 'TP-REMINDER' });
      props.setProperty(pk, JSON.stringify({ at: Date.now(), n: last.n + 1 }));
      items.forEach(function(r) { var ch = {}; ch[TPC.LAST_REMINDER] = new Date(); ch[TPC.REMINDERS] = r.reminders + 1; _tpwSet(sh, r.row, ch); });
      out.digests++;
    });
  } catch (e) { Logger.log('[tpwDailyJob] ' + e.message); }
  Logger.log('[tpwDailyJob] ' + JSON.stringify(out));
  return out;
}

// Idempotent: (re)creates the two time triggers.
function setupTPWorkspaceTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var f = t.getHandlerFunction();
    if (f === 'tpwVerifyJob' || f === 'tpwDailyJob') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('tpwVerifyJob').timeBased().everyHours(4).create();
  ScriptApp.newTrigger('tpwDailyJob').timeBased().everyDays(1).atHour(9).create();
  Logger.log('[tpw] triggers installed: verify every 4 h, reminders daily at 9');
  return tpwAutomationStatus();
}

function tpwAutomationStatus() {
  var v = false, d = false;
  try {
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'tpwVerifyJob') v = true;
      if (t.getHandlerFunction() === 'tpwDailyJob') d = true;
    });
  } catch (e) {}
  return { verify: v, reminders: d, on: v && d };
}

function tpwEnableAutomation(token) {
  var ctx = _tpwCtx(token);
  if (!ctx.ok || !ctx.isAdmin) return { success: false, message: 'Admin only.' };
  try { return { success: true, automation: setupTPWorkspaceTriggers() }; } catch (e) { return { success: false, message: e.message }; }
}

// Admin: run the checks now instead of waiting for the schedule.
function tpwRunNow(token, which) {
  var ctx = _tpwCtx(token);
  if (!ctx.ok || !ctx.isAdmin) return { success: false, message: 'Admin only.' };
  try {
    if (which === 'verify') return { success: true, result: tpwVerifyJob() };
    return { success: true, result: tpwDailyJob() };
  } catch (e) { return { success: false, message: e.message }; }
}

// ───────────────────────────── admin: test data + sample emails ─────────────────────────────
// Removes every request created in test mode (and nothing else).
function tpwDeleteTests(token) {
  var ctx = _tpwCtx(token);
  if (!ctx.ok || !ctx.isAdmin) return { success: false, message: 'Admin only.' };
  try {
    return _tpwWithLock(function() {
      var sh = _tpwTracker();
      var reqs = _tpwReadRequests().filter(function(r) { return r.source === 'Test'; });
      reqs.sort(function(a, b) { return b.row - a.row; }).forEach(function(r) { sh.deleteRow(r.row); });
      return { success: true, message: reqs.length + ' test requirement' + (reqs.length === 1 ? '' : 's') + ' removed.' };
    });
  } catch (e) { return { success: false, message: e.message }; }
}

// Admin: email yourself what TP managers receive (a reminder + a "sent back" email), built from sample rows.
function tpwSendSamples(token) {
  var ctx = _tpwCtx(token);
  if (!ctx.ok || !ctx.isAdmin) return { success: false, message: 'Admin only.' };
  var to = /@/.test(ctx.email) ? ctx.email : '';
  if (!to) return { success: false, message: 'Your sign-in has no email address, so there is nowhere to send the samples.' };
  try {
    var props = PropertiesService.getScriptProperties();
    var n = Number(props.getProperty('TPW_SAMPLE_N') || 0);
    props.setProperty('TPW_SAMPLE_N', String(n + 1));
    var items = [
      { teacher: 'Sample Teacher A', course: 'Python with AI', status: 'Requested', tier: 'upskill', _age: 16 },
      { teacher: 'Sample Teacher B', course: 'Design with Roblox', status: 'Requested', tier: 'fast', _age: 5 },
      { teacher: 'Sample Teacher C', course: 'Explore GenAI', status: 'Rejected', tier: 'fast', _age: 9 }
    ];
    var d = _tpwDigestFor('Alex Sample', items, n);
    sendTrackedEmail({ to: to, subject: '[SAMPLE] ' + d.subject, htmlBody: d.html, jlid: 'TP-SAMPLE' });
    var rej = _tpwWrapEmail('Sent back: not on the data yet',
      '<p style="margin:0 0 12px;">Hi Alex,</p><p style="margin:0 0 12px;">You marked this teacher as upskilled, but <b>' + TPW.REJECT_AFTER_HOURS + ' hours later the data still does not show it</b> (we need ' + TPW.PASS_PCT + '% or more). We gave Moodle time to catch up — it has not. Please check Moodle, and mark it done again once it shows.</p>'
      + '<table style="width:100%;border-collapse:collapse;font-size:14px;"><tr style="background:#f6f7fb;text-align:left;"><th style="padding:8px 10px;">Teacher</th><th style="padding:8px 10px;">Course</th><th style="padding:8px 10px;">Data shows</th></tr><tr><td style="padding:8px 10px;border-top:1px solid #eef0f4;">Sample Teacher C</td><td style="padding:8px 10px;border-top:1px solid #eef0f4;">Explore GenAI</td><td style="padding:8px 10px;border-top:1px solid #eef0f4;color:#b91c1c;">45%</td></tr></table>',
      'Open TP Workspace', _tpwAppUrl());
    sendTrackedEmail({ to: to, subject: '[SAMPLE] ↩️ Upskilling not on the data yet — 1 sent back to you', htmlBody: rej, jlid: 'TP-SAMPLE' });
    return { success: true, message: 'Two sample emails sent to ' + to + '. Press again to see a different reminder message.' };
  } catch (e) { return { success: false, message: e.message }; }
}

// ───────────────────────────── read-only checks on where each teacher comes from (data is imported — nothing here edits it) ─────────────────────────────
var TPW_LIVE_STATUSES = ['Active', 'EWS', 'Friendly'];

// The exact teachers a manager sees, and which sheet put each one on their list.
function tpwManagerTeachers(token, managerName) {
  try {
    var ctx = _tpwCtx(token);
    if (!ctx.ok || !ctx.isAdmin) return { success: false, message: 'Admin only.' };
    var mgr = _tpwReadManagers().filter(function(m) { return _tpwNorm(m.name) === _tpwNorm(managerName); })[0] || { name: String(managerName || '').trim(), email: '' };
    var scope = _tpwScopeFor(mgr);
    var list = Object.keys(scope).map(function(k) {
      var t = scope[k];
      var viaName = _tpwNorm(t.manager) === _tpwNorm(mgr.name);
      return { name: t.name, status: t.status, hidden: t.hidden, source: viaName ? t.source : 'TP Manager Email column', manager: t.manager };
    }).sort(function(a, b) { return a.name.localeCompare(b.name); });
    var live = list.filter(function(t) { return TPW_LIVE_STATUSES.indexOf(t.status) > -1; }).length;
    return { success: true, manager: mgr.name, total: list.length, live: live, teachers: list };
  } catch (e) { return { success: false, message: e.message }; }
}

// Teachers nobody can see, teachers under a manager who is not in the system, and teachers whose two sheets disagree.
function tpwCoverage(token) {
  try {
    var ctx = _tpwCtx(token);
    if (!ctx.ok || !ctx.isAdmin) return { success: false, message: 'Admin only.' };
    var dir = _tpwTeacherDirectory().filter(function(t) { return TPW_LIVE_STATUSES.indexOf(t.status) > -1; });
    var inSystem = {};
    _tpwReadManagers().forEach(function(m) { inSystem[_tpwNorm(m.name)] = true; });
    var noMgr = [], notIn = {}, conflicts = [];
    dir.forEach(function(t) {
      var m = _tpwNorm(t.manager);
      if (!m && !t.tpEmail) noMgr.push({ name: t.name, status: t.status });
      else if (m && !inSystem[m]) { (notIn[m] = notIn[m] || { manager: t.manager, teachers: [] }).teachers.push(t.name); }
      if (t.personaMgr && t.dataMgr && _tpwNorm(t.personaMgr) !== _tpwNorm(t.dataMgr)) conflicts.push({ name: t.name, persona: t.personaMgr, data: t.dataMgr });
    });
    return { success: true, activeTeachers: dir.length, noManager: noMgr,
      notInSystem: Object.keys(notIn).map(function(k) { return notIn[k]; }).sort(function(a, b) { return b.teachers.length - a.teachers.length; }),
      conflicts: conflicts };
  } catch (e) { return { success: false, message: e.message }; }
}

// Just the teacher names, fast (no HubSpot, no audit-log scan) — what the Teacher Intelligence "Select Teacher" box needs.
// Admin: every teacher. TP manager (or admin viewing as one): only their own.
function tpwTeacherNames(token, viewAs) {
  var ctx = _tpwCtx(token, viewAs);
  if (!ctx.ok) throw new Error(ctx.message);
  var dir = _tpwTeacherDirectory();
  var names = dir.filter(function(t) { return ctx.isAdmin || (ctx.scope && ctx.scope[_tpwNorm(t.name)]); })
    .map(function(t) { return t.name; });
  var seen = {};
  return names.filter(function(n) { var k = _tpwNorm(n); if (!k || seen[k]) return false; seen[k] = true; return true; })
    .sort(function(a, b) { return a.localeCompare(b); });
}
