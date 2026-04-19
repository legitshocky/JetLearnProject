/**
 * LearnerProgressionService.js
 * ────────────────────────────────────────────────────────────────────
 * Reads Athena CPRS + PRMS sheets (pasted from daily CSV exports),
 * computes per-learner course completion timelines, checks teacher
 * upskill readiness for next courses, and drives the Course Planner
 * dashboard + Smart Migration trigger flow.
 *
 * Sheet requirements (in MIGRATION_SHEET_ID spreadsheet):
 *   "Athena CPRS"  — paste from Athena_cprs_*.csv  (full headers row 1)
 *   "Athena PRMS"  — paste from Athena_prms_*.csv  (full headers row 1)
 * ────────────────────────────────────────────────────────────────────
 */

// ── Constants ──────────────────────────────────────────────────────
var LP_ATHENA_SHEET_ID = '1EodMl-ls6hJe7ONOp4Yyt5r901J4d5PtjgF5iXeYDk8';

// ── Cache helpers (chunked — CacheService is 100KB per key) ────────
var LP_CACHE_KEY  = 'LP_PROG_V4';  // bumped — migrationNeeded now uses !teacherReady (< 71%)
var LP_CACHE_TTL  = 600;  // 10 minutes
var LP_CACHE_CHUNK = 90000; // 90KB per chunk (leave buffer)

function _lpCacheStore(data) {
  try {
    var sc   = CacheService.getScriptCache();
    var json = JSON.stringify(data);
    var n    = Math.ceil(json.length / LP_CACHE_CHUNK);
    sc.put(LP_CACHE_KEY + '_meta', JSON.stringify({ chunks: n, ts: Date.now(), len: json.length }), LP_CACHE_TTL);
    for (var i = 0; i < n; i++) {
      sc.put(LP_CACHE_KEY + '_' + i, json.substring(i * LP_CACHE_CHUNK, (i + 1) * LP_CACHE_CHUNK), LP_CACHE_TTL);
    }
    Logger.log('[LP] Cached ' + data.learners.length + ' learners in ' + n + ' chunk(s) (' + json.length + ' bytes)');
  } catch(e) { Logger.log('[LP] _lpCacheStore error: ' + e.message); }
}

function _lpCacheLoad() {
  try {
    var sc      = CacheService.getScriptCache();
    var metaRaw = sc.get(LP_CACHE_KEY + '_meta');
    if (!metaRaw) return null;
    var meta   = JSON.parse(metaRaw);
    var parts  = [];
    for (var i = 0; i < meta.chunks; i++) {
      var part = sc.get(LP_CACHE_KEY + '_' + i);
      if (!part) return null;  // chunk expired — treat as cache miss
      parts.push(part);
    }
    var result = JSON.parse(parts.join(''));
    result._fromCache = true;
    result._cachedAt  = meta.ts;
    return result;
  } catch(e) { Logger.log('[LP] _lpCacheLoad error: ' + e.message); return null; }
}

function clearCoursePlannerCache() {
  try {
    var sc      = CacheService.getScriptCache();
    var metaRaw = sc.get(LP_CACHE_KEY + '_meta');
    if (metaRaw) {
      var meta = JSON.parse(metaRaw);
      var keys = [LP_CACHE_KEY + '_meta'];
      for (var i = 0; i < (meta.chunks || 20); i++) keys.push(LP_CACHE_KEY + '_' + i);
      sc.removeAll(keys);
    }
    _lpTeacherCourseMapCache = null;
    Logger.log('[LP] Course Planner cache cleared');
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}
var LP_CPRS_SHEET  = 'CPR';   // tab name in Athena spreadsheet
var LP_PRMS_SHEET  = 'PRM';   // tab name in Athena spreadsheet
var LP_ALERT_WEEKS_CRITICAL = 4;   // ≤ 4 weeks → critical
var LP_ALERT_WEEKS_WARNING  = 8;   // ≤ 8 weeks → warning
var LP_FREQ_WINDOW_DAYS     = 28;  // look-back for frequency calc
var LP_FREQ_FALLBACK        = 1.0; // if too few sessions to measure
var LP_UPSKILL_THRESHOLD    = 71;  // % to consider teacher "ready"

// ── Utility: extract JLID from "JL52767973402C - Atharva Gokul" ───
function parseLearnerJLID(str) {
  if (!str) return '';
  var m = String(str).match(/JL[\w]+C/);
  return m ? m[0] : String(str).split(' - ')[0].trim();
}

// ── Utility: extract learner display name ─────────────────────────
function parseLearnerName(str) {
  if (!str) return '';
  var parts = String(str).split(' - ');
  return parts.length > 1 ? parts.slice(1).join(' - ').trim() : str.trim();
}

// ── Utility: extract teacher display name ────────────────────────
function parseTeacherName(str) {
  if (!str) return '';
  var parts = String(str).split(' - ');
  return parts.length > 1 ? parts.slice(1).join(' - ').trim() : str.trim();
}

// ── Utility: extract teacher TJL code ────────────────────────────
function parseTeacherCode(str) {
  if (!str) return '';
  var m = String(str).match(/TJL\d+/);
  return m ? m[0] : '';
}

// ── Utility: parse date value from sheet cell ────────────────────
function lpParseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  var s = String(v).trim();
  // handle YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return new Date(s);
  // handle DD-MM-YY or DD/MM/YYYY
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    var yr = parseInt(m[3]); if (yr < 100) yr += 2000;
    return new Date(yr, parseInt(m[2]) - 1, parseInt(m[1]));
  }
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// ── Build CPRS map: jlid → [{ date, happened, classNumber, course, teacher, cancelReason }]
function _buildCPRSMap() {
  var map = {};
  try {
    var data = _getCachedSheetData(LP_CPRS_SHEET, LP_ATHENA_SHEET_ID);
    if (!data || data.length < 2) {
      Logger.log('[LP] CPR sheet empty or missing (Athena spreadsheet)');
      return map;
    }
    var h = data[0].map(function(x){ return String(x).trim().toLowerCase(); });
    var iDate    = h.indexOf('class date');
    var iTeacher = h.indexOf('teacher');
    var iLearner = h.indexOf('learner');
    var iCourse  = h.indexOf('course');
    var iHappen  = h.indexOf('class happened?');
    if (iHappen === -1) iHappen = h.indexOf('class happened');
    var iClassNo = h.indexOf('class number');
    var iCancel  = h.indexOf('reason for cancellation');
    if (iCancel === -1) iCancel = h.indexOf('cancellation reason');

    Logger.log('[LP] CPRS headers found: date=' + iDate + ' learner=' + iLearner + ' happened=' + iHappen + ' classNo=' + iClassNo);

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var learnerRaw = String(row[iLearner] || '').trim();
      if (!learnerRaw) continue;
      var jlid = parseLearnerJLID(learnerRaw);
      if (!jlid) continue;
      var happenedRaw = iHappen > -1 ? String(row[iHappen] || '').trim().toLowerCase() : '';
      var happened = happenedRaw === 'yes' || happenedRaw === 'true' || happenedRaw === '1';
      var dateVal = iDate > -1 ? lpParseDate(row[iDate]) : null;
      var classNo = iClassNo > -1 ? (parseInt(row[iClassNo]) || 0) : 0;
      if (!map[jlid]) map[jlid] = [];
      map[jlid].push({
        date        : dateVal,
        happened    : happened,
        classNumber : classNo,
        course      : iCourse  > -1 ? String(row[iCourse]  || '').trim() : '',
        teacher     : iTeacher > -1 ? String(row[iTeacher] || '').trim() : '',
        cancelReason: iCancel  > -1 ? String(row[iCancel]  || '').trim() : ''
      });
    }
    Logger.log('[LP] CPRS map built: ' + Object.keys(map).length + ' learners');
  } catch(e) {
    Logger.log('[LP] _buildCPRSMap error: ' + e.message);
  }
  return map;
}

// ── Build PRMS map: jlid → { currentCourse, classesLeft, nextCourses, finalized, prmType, teacher, learnerName, prmDate }
function _buildPRMSMap() {
  var map = {};
  try {
    var data = _getCachedSheetData(LP_PRMS_SHEET, LP_ATHENA_SHEET_ID);
    if (!data || data.length < 2) {
      Logger.log('[LP] PRM sheet empty or missing (Athena spreadsheet)');
      return map;
    }
    var h = data[0].map(function(x){ return String(x).trim().toLowerCase(); });
    var iTeacher    = h.indexOf('teacher');
    var iLearner    = h.indexOf('learner');
    var iPrmDate    = h.indexOf('prm date');
    if (iPrmDate === -1) iPrmDate = h.indexOf('prm prep date');
    var iCurCourse  = h.indexOf('current course');
    var iClassesLeft= h.indexOf('# classes left');
    if (iClassesLeft === -1) iClassesLeft = h.indexOf('classes left');
    var iFinalized  = h.indexOf('next courses finalized?');
    if (iFinalized  === -1) iFinalized  = h.indexOf('next courses finalized');
    var iPrmType    = h.indexOf('prm type');
    var iClsNotes   = h.indexOf("cls team notes");

    // Find next course pairs: "next course #1" / "number of classes #1"
    var nextCoursePairs = []; // [{nameIdx, countIdx}]
    for (var ni = 1; ni <= 8; ni++) {
      var nIdx = h.indexOf('next course #' + ni);
      if (nIdx === -1) nIdx = h.indexOf('next course #0' + ni);
      var cIdx = h.indexOf('number of classes #' + ni);
      if (cIdx === -1) cIdx = h.indexOf('number of classes #0' + ni);
      if (nIdx > -1) nextCoursePairs.push({ nameIdx: nIdx, countIdx: cIdx });
    }

    Logger.log('[LP] PRMS headers: learner=' + iLearner + ' curCourse=' + iCurCourse + ' classesLeft=' + iClassesLeft + ' nextPairs=' + nextCoursePairs.length);

    // Collect all rows per JLID, then keep the LATEST (most recent PRM date)
    var rawMap = {}; // jlid → array of rows
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var learnerRaw = String(row[iLearner] || '').trim();
      if (!learnerRaw) continue;
      var jlid = parseLearnerJLID(learnerRaw);
      if (!jlid) continue;
      if (!rawMap[jlid]) rawMap[jlid] = [];
      rawMap[jlid].push({ row: row, prmDate: iPrmDate > -1 ? lpParseDate(row[iPrmDate]) : null });
    }

    Object.keys(rawMap).forEach(function(jlid) {
      // Sort by PRM date descending, take latest
      var entries = rawMap[jlid].sort(function(a, b) {
        var da = a.prmDate ? a.prmDate.getTime() : 0;
        var db = b.prmDate ? b.prmDate.getTime() : 0;
        return db - da;
      });
      var best = entries[0].row;
      var prmDate = entries[0].prmDate;

      var learnerRaw = String(best[iLearner] || '').trim();
      var nextCourses = [];
      nextCoursePairs.forEach(function(pair) {
        var name  = String(best[pair.nameIdx]  || '').trim();
        var count = pair.countIdx > -1 ? parseInt(best[pair.countIdx]) || 0 : 0;
        if (name) nextCourses.push({ name: name, count: count || 12 });
      });

      var finalizedRaw = iFinalized > -1 ? String(best[iFinalized] || '').trim().toLowerCase() : '';
      var finalized    = finalizedRaw === 'yes' || finalizedRaw === 'true';

      map[jlid] = {
        learnerName   : parseLearnerName(learnerRaw),
        teacher       : iTeacher   > -1 ? String(best[iTeacher]    || '').trim() : '',
        currentCourse : iCurCourse > -1 ? String(best[iCurCourse]  || '').trim() : '',
        classesLeft   : iClassesLeft > -1 ? (parseInt(best[iClassesLeft]) || 0) : 0,
        nextCourses   : nextCourses,
        finalized     : finalized,
        prmType       : iPrmType  > -1 ? String(best[iPrmType]  || '').trim() : '',
        clsNotes      : iClsNotes > -1 ? String(best[iClsNotes] || '').trim() : '',
        prmDate       : prmDate
      };
    });

    Logger.log('[LP] PRMS map built: ' + Object.keys(map).length + ' learners');
  } catch(e) {
    Logger.log('[LP] _buildPRMSMap error: ' + e.message);
  }
  return map;
}

// ── Convert day-of-week slots to upcoming calendar dates ─────────
// Input:  [{day: "Friday", time: "2:00 AM"}]
// Output: [{date: "2026-04-24", slot: "02:00 AM"}, {date: "2026-05-01", slot: "02:00 AM"}]
// searchMatchingTeachers requires date-based slots, not day-name slots.
function _convertDayTimeSlotsToDateSlots(sessions) {
  var DAY_MAP = { sunday:0, monday:1, tuesday:2, wednesday:3, thursday:4, friday:5, saturday:6 };
  var result = [];
  if (!sessions || sessions.length === 0) return result;
  var today = new Date();
  today.setHours(0,0,0,0);
  var tz = Session.getScriptTimeZone();

  sessions.forEach(function(s) {
    var dayLow = (s.day || '').toLowerCase().trim();
    var targetDay = DAY_MAP[dayLow];
    if (targetDay === undefined) return;

    // Normalize: "2:00 AM" → "02:00 AM"
    var timeStr = (s.time || '').trim();
    var tMatch  = timeStr.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
    var normTime = timeStr;
    if (tMatch) {
      var hr = parseInt(tMatch[1]);
      normTime = (hr < 10 ? '0' + hr : String(hr)) + ':' + tMatch[2] + ' ' + tMatch[3].toUpperCase();
    }
    if (!normTime) return;

    // Produce next 2 occurrences of this weekday
    var diff = (targetDay - today.getDay() + 7) % 7;
    for (var occ = 0; occ < 2; occ++) {
      var d = new Date(today.getTime() + (diff + occ * 7) * 86400000);
      result.push({ date: Utilities.formatDate(d, tz, 'yyyy-MM-dd'), slot: normTime });
    }
  });
  return result;
}

// ── Fetch learner class schedule from master Google Calendar ──────
// Searches hello@jet-learn.com calendar for events containing learner
// name or JLID in the next 30 days. Returns day + time in CET.
function getLearnerCalendarSchedule(jlid, learnerName) {
  try {
    var cal = CalendarApp.getCalendarById('hello@jet-learn.com');
    if (!cal) { Logger.log('[LP] Calendar hello@jet-learn.com not accessible'); return { found: false }; }

    var now   = new Date();
    var end   = new Date(now.getTime() + 14 * 86400000); // 2 weeks (was 6 — faster fetch)
    var events = cal.getEvents(now, end);

    var nameLow = (learnerName || '').toLowerCase().replace(/\s+/g,' ').trim();
    var jlidLow = (jlid        || '').toLowerCase();

    // Collect unique day+time slots from matching events
    var slotMap = {};  // "Friday|02:00 AM" → {day, timeCET, count}
    events.forEach(function(ev) {
      var title = (ev.getTitle() || '').toLowerCase();
      if (!nameLow && !jlidLow) return;
      if (nameLow && title.indexOf(nameLow) === -1 && title.indexOf(jlidLow) === -1) return;
      if (!nameLow && jlidLow && title.indexOf(jlidLow) === -1) return;

      var startTime = ev.getStartTime();
      var dayStr    = Utilities.formatDate(startTime, 'Europe/Paris', 'EEEE');   // "Friday"
      var timeCET   = Utilities.formatDate(startTime, 'Europe/Paris', 'h:mm a'); // "2:00 AM"
      var key       = dayStr + '|' + timeCET;
      if (!slotMap[key]) slotMap[key] = { day: dayStr, timeCET: timeCET, count: 0 };
      slotMap[key].count++;
    });

    var slots = Object.keys(slotMap).map(function(k){ return slotMap[k]; });
    if (slots.length === 0) return { found: false };

    // Sort by frequency (most common = most likely regular slot)
    slots.sort(function(a,b){ return b.count - a.count; });
    return { found: true, slots: slots, primary: slots[0] };
  } catch(e) {
    Logger.log('[LP] getLearnerCalendarSchedule error: ' + e.message);
    return { found: false, error: e.message };
  }
}

// ── Compute class frequency (sessions/week) from last N days ──────
function _computeFrequency(cprsRows) {
  if (!cprsRows || cprsRows.length === 0) return LP_FREQ_FALLBACK;
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - LP_FREQ_WINDOW_DAYS);
  var recentDone = cprsRows.filter(function(r) {
    return r.happened && r.date && r.date >= cutoff;
  });
  if (recentDone.length < 2) return LP_FREQ_FALLBACK;
  return Math.round((recentDone.length / (LP_FREQ_WINDOW_DAYS / 7)) * 10) / 10;
}

// ── Cached teacher→course map (mirrors TeacherService.js searchMatchingTeachers logic) ─
// Built once per execution. Key: normalizedTeacherName → { courseLower: progressString }
var _lpTeacherCourseMapCache = null;
function _lpGetTeacherCourseMap() {
  if (_lpTeacherCourseMapCache) return _lpTeacherCourseMapCache;
  var map = {};
  try {
    var tcSheet = _getCachedSheetData(CONFIG.SHEETS.TEACHER_COURSES);
    if (!tcSheet || tcSheet.length < 2) { _lpTeacherCourseMapCache = map; return map; }

    // Find header row (row where col-0 is "teacher")
    var tcHeaderIdx = 0;
    for (var hi = 0; hi < Math.min(tcSheet.length, 10); hi++) {
      if (String(tcSheet[hi][0]).trim().toLowerCase() === 'teacher') { tcHeaderIdx = hi; break; }
    }
    var tcHeaders   = tcSheet[tcHeaderIdx];
    var COURSE_START = 4; // matches TeacherService constant

    tcSheet.slice(tcHeaderIdx + 1).forEach(function(row) {
      var rawName = String(row[0] || '').trim();
      if (!rawName || rawName.toLowerCase() === 'teacher') return;
      var key = normalizeTeacherName(rawName);
      var courseMap = {};
      for (var ci = COURSE_START; ci < tcHeaders.length; ci++) {
        var cName = String(tcHeaders[ci] || '').trim();
        var prog  = String(row[ci]  || '').trim();
        if (!cName) continue;
        // Store raw value. null=empty cell. "Not Onboarded"=cell has that text.
        courseMap[cName.toLowerCase()] = prog || null;
      }
      map[key] = courseMap;
    });
    Logger.log('[LP] _lpGetTeacherCourseMap built: ' + Object.keys(map).length + ' teachers');
  } catch(e) {
    Logger.log('[LP] _lpGetTeacherCourseMap error: ' + e.message);
  }
  _lpTeacherCourseMapCache = map;
  return map;
}

// ── Fuzzy course lookup — exact → prefix → 15-char substr → word-overlap ≥60% ──
// Mirrors getCourseProgress() closure inside TeacherService.searchMatchingTeachers exactly.
function _lpCourseProgress(tNorm, courseName) {
  if (!tNorm || !courseName) return null;
  var tcMap   = _lpGetTeacherCourseMap();
  var courses = tcMap[tNorm];

  // Fuzzy fallback: Teacher Courses sheet may store "Shobhit" but CPRS has "Shobhit Gupta"
  // Try: first-word match, then last-word match
  if (!courses) {
    var parts     = tNorm.split(' ');
    var firstWord = parts[0];
    var lastWord  = parts[parts.length - 1];
    var mapKeys   = Object.keys(tcMap);
    for (var ki = 0; ki < mapKeys.length; ki++) {
      var mk      = mapKeys[ki];
      var mkParts = mk.split(' ');
      var mkFirst = mkParts[0];
      var mkLast  = mkParts[mkParts.length - 1];
      // Sheet has "Shobhit" and CPRS has "Shobhit Gupta" → first words match
      if (firstWord.length > 2 && mkFirst === firstWord) {
        courses = tcMap[mk];
        Logger.log('[LP] Teacher first-name match: "' + tNorm + '" → "' + mk + '"');
        break;
      }
      // Sheet has "Gupta Shobhit" (reversed) or last-name only
      if (lastWord.length > 2 && mkLast === lastWord && lastWord !== firstWord) {
        courses = tcMap[mk];
        Logger.log('[LP] Teacher last-name match: "' + tNorm + '" → "' + mk + '"');
        break;
      }
    }
  }
  if (!courses) return null;  // teacher truly NOT found in Teacher Courses sheet

  var cLower = courseName.toLowerCase().trim();
  // 1. exact
  if (courses[cLower] !== undefined) return courses[cLower] || 'Not Onboarded';

  var courseKeys = Object.keys(courses);
  for (var ci = 0; ci < courseKeys.length; ci++) {
    var ck = courseKeys[ci];
    // 2. prefix (either direction)
    if (ck.indexOf(cLower) === 0 || cLower.indexOf(ck) === 0) return courses[ck] || 'Not Onboarded';
    // 3. 15-char substring
    var prefix = cLower.substring(0, 15);
    if (prefix.length >= 10 && ck.indexOf(prefix) !== -1) return courses[ck] || 'Not Onboarded';
    // 4. word overlap ≥60% of meaningful words (>3 chars)
    var appWords   = cLower.split(/\s+/).filter(function(w){ return w.length > 3; });
    var sheetWords = ck.split(/\s+/).filter(function(w){ return w.length > 3; });
    if (appWords.length > 0 && sheetWords.length > 0) {
      var overlap = appWords.filter(function(w){ return sheetWords.indexOf(w) > -1; }).length;
      var minLen  = Math.min(appWords.length, sheetWords.length);
      if (overlap >= Math.ceil(minLen * 0.6)) return courses[ck] || 'Not Onboarded';
    }
  }
  // Teacher IS in the sheet but has no record for this course → Not Onboarded (not a name-match issue)
  return 'Not Onboarded';
}

// ── Parse a raw progress string into structured status ───────────────────
// prog values from sheet: null (not found), "Not Onboarded", "0%", "61-70%", "100%", etc.
function _parseProgressStatus(prog) {
  var progressKnown  = prog !== null;
  var isNotOnboarded = !prog || String(prog).toLowerCase().trim() === 'not onboarded';
  var progNum = (!isNotOnboarded && prog) ? (parseInt(String(prog)) || 0) : 0;
  return {
    prog          : prog,
    progNum       : progNum,
    progressKnown : progressKnown,
    teacherReady  : progressKnown && !isNotOnboarded && progNum >= LP_UPSKILL_THRESHOLD,
    inProgress    : progressKnown && !isNotOnboarded && progNum > 0 && progNum < LP_UPSKILL_THRESHOLD,
    notStarted    : progressKnown && (isNotOnboarded || progNum === 0),
    displayLabel  : !progressKnown ? 'Not Onboarded' : (prog || 'Not Onboarded')
  };
}

// ── Public wrapper: try both raw name and resolved/canonical name ──────────
function _getTeacherCourseProgressStandalone(teacherRawName, courseName) {
  if (!teacherRawName || !courseName) return null;
  try {
    var tNorm  = normalizeTeacherName(teacherRawName);
    var prog   = _lpCourseProgress(tNorm, courseName);
    // Also try resolveTeacherName if available (handles aliases)
    if (!prog) {
      try {
        var tCanon = normalizeTeacherName(resolveTeacherName(teacherRawName));
        if (tCanon && tCanon !== tNorm) prog = _lpCourseProgress(tCanon, courseName);
      } catch(re) { /* resolveTeacherName may not be available */ }
    }
    return prog; // null = not found; string = progress value
  } catch(e) {
    Logger.log('[LP] _getTeacherCourseProgressStandalone error: ' + e.message);
    return null;
  }
}

// ── Check HubSpot learner health (Critical flag) ──────────────────
// Batches lookups — returns map jlid → { critical, health, classSessions }
function _buildHealthMap(jlids) {
  var healthMap = {};
  if (!jlids || jlids.length === 0) return healthMap;
  // HubSpot search supports up to 100 filterGroups per request
  var batchSize = 100;
  var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  if (!token) { Logger.log('[LP] No HubSpot token — health data unavailable'); return healthMap; }

  for (var b = 0; b < jlids.length; b += batchSize) {
    var batch = jlids.slice(b, b + batchSize);
    try {
      var filters = batch.map(function(id) {
        return { filters: [{ propertyName: 'jetlearner_id', operator: 'EQ', value: id }] };
      });
      var body = {
        filterGroups          : filters,
        properties            : ['jetlearner_id', 'learner_health', 'class_timings', 'regular_class_day', 'frequency_of_classes', 'dealname', 'dealstage', 'closedate'],
        propertiesWithHistory : ['current_course__t_'],
        limit                 : batchSize
      };
      var resp = UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/objects/deals/search', {
        method: 'post',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify(body),
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) {
        var results = JSON.parse(resp.getContentText()).results || [];
        results.forEach(function(r) {
          var p    = r.properties;
          var id   = p.jetlearner_id || '';
          var hlth = String(p.learner_health || '').toLowerCase();
          // Parse class_timings string → [{day, time}]
          // Format examples: "Monday at 06:00 PM", "Monday;Wednesday at 05:00 PM"
          var rawTimings = p.class_timings || p.regular_class_day || '';
          var parsedSessions = [];
          if (rawTimings) {
            rawTimings.split(';').forEach(function(seg) {
              seg = seg.trim();
              var m = seg.match(/(\w+)\s+at\s+(\d{1,2}:\d{2}\s*(?:AM|PM))/i);
              if (m) { parsedSessions.push({ day: m[1], time: m[2].trim() }); }
              else if (seg) { parsedSessions.push({ day: seg, time: '' }); }
            });
          }
          // HubSpot closedate is milliseconds epoch (string)
          var closeDateMs = parseInt(p.closedate || '') || 0;
          // Parse current_course__t_ property history → ordered course list (oldest first)
          var rawCourseHistory = (r.propertiesWithHistory && r.propertiesWithHistory['current_course__t_']) || [];
          // HubSpot returns newest-first; sort ascending by timestamp
          var sortedHistory = rawCourseHistory.slice().sort(function(a, b) {
            return new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime();
          });
          // Dedupe: keep first occurrence of each distinct course name
          var courseHistory = [];
          sortedHistory.forEach(function(h) {
            var v = (h.value || '').trim();
            if (v && courseHistory.indexOf(v) === -1) courseHistory.push(v);
          });
          Logger.log('[LP] courseHistory for ' + id + ': ' + JSON.stringify(courseHistory));

          healthMap[id] = {
            critical       : hlth === 'critical' || hlth === '4' || hlth.indexOf('critical') > -1,
            health         : p.learner_health || '',
            classSessions  : parsedSessions,   // [{day, time}]
            classTimingsRaw: rawTimings,
            hsFrequency    : p.frequency_of_classes || '',
            dealId         : r.id || '',
            dealStage      : p.dealstage || '',
            dealName       : p.dealname || '',
            closeDate      : closeDateMs,  // epoch ms or 0
            courseHistory  : courseHistory  // ['Fundamentals of Python', 'Python Edublocks', 'Python Game Dev', ...]
          };
        });
      }
      Utilities.sleep(50); // minimal buffer — batch of 100 is well within HubSpot rate limits
    } catch(he) {
      Logger.log('[LP] Health batch error: ' + he.message);
    }
  }
  return healthMap;
}

// ─────────────────────────────────────────────────────────────────────
// MAIN: getLearnerProgressions()
// Called by the Course Planner dashboard page.
// ─────────────────────────────────────────────────────────────────────
function getLearnerProgressions(forceRefresh) {
  try {
    Logger.log('[LP] getLearnerProgressions start (forceRefresh=' + !!forceRefresh + ')');

    // ── Serve from cache if available ────────────────────────────
    if (!forceRefresh) {
      var cached = _lpCacheLoad();
      if (cached) {
        var ageMin = Math.round((Date.now() - (cached._cachedAt || 0)) / 60000);
        Logger.log('[LP] Returning cached data (' + ageMin + 'm old, ' + (cached.learners || []).length + ' learners)');
        return cached;
      }
    }

    _lpTeacherCourseMapCache = null; // reset so fresh sheet data is used on full recompute
    var cprsMap  = _buildCPRSMap();
    var prmsMap  = _buildPRMSMap();
    var jlids    = Object.keys(prmsMap);
    if (jlids.length === 0) {
      return { success: true, learners: [], stats: { total:0, criticalCount:0, warningCount:0, migrationsDue:0, teacherGapCount:0 }, message: 'No PRMS data found. Upload Athena PRMS sheet.' };
    }

    // Batch health lookup
    var healthMap = _buildHealthMap(jlids);

    var today = new Date();
    var learners = [];

    jlids.forEach(function(jlid) {
      var prms      = prmsMap[jlid];
      var cprsRows  = cprsMap[jlid] || [];
      var healthInfo = healthMap[jlid] || {};   // moved up — needed for courseHistory at line 663

      // ── Adjusted classesLeft ────────────────────────────────────────
      // PRMS "# Classes Left" is a snapshot from PRM date (can be months stale).
      // Subtract sessions that happened AFTER the PRM date to get a live count.
      var prmDate       = prms.prmDate;                          // Date or null
      var classesLeftRaw = prms.classesLeft;                     // as of PRM date
      var sessionsDoneAfterPRM = 0;
      if (prmDate) {
        sessionsDoneAfterPRM = cprsRows.filter(function(r) {
          return r.happened && r.date && r.date > prmDate;
        }).length;
      }
      var classesLeft = Math.max(0, classesLeftRaw - sessionsDoneAfterPRM);

      // PRM staleness (days since PRM date)
      var prmStaleDays = prmDate ? Math.floor((today.getTime() - prmDate.getTime()) / 86400000) : null;

      var classesDone   = cprsRows.filter(function(r){ return r.happened; }).length;
      var frequency     = _computeFrequency(cprsRows);
      var projectedWeeks = frequency > 0 ? Math.round((classesLeft / frequency) * 10) / 10 : null;
      var projectedDate  = projectedWeeks != null
        ? new Date(today.getTime() + projectedWeeks * 7 * 24 * 60 * 60 * 1000)
        : null;

      // ── Subject category (Coding vs Maths) ─────────────────────────
      var courseLow = (prms.currentCourse || '').toLowerCase();
      var isMaths = /\bmath/i.test(courseLow) || /maths year/i.test(courseLow)
                 || /year \d.*maths/i.test(courseLow);
      var subject = isMaths ? 'maths' : 'coding';

      // ── Total course length estimate ────────────────────────────────
      // classesDone (from CPRS) + classesLeft (adjusted) = estimated total
      var totalCourseClasses = classesDone + classesLeft;

      // Teacher info — prefer most recent CPRS session (accurate, vs PRMS which may be stale snapshot)
      var teacherRaw  = prms.teacher;
      var cprsTeacher = '';
      if (cprsRows.length > 0) {
        // Sort CPRS descending by date, pick first row that has a teacher
        var cprsDesc = cprsRows.slice().sort(function(a,b) {
          return (b.date ? b.date.getTime() : 0) - (a.date ? a.date.getTime() : 0);
        });
        for (var ri = 0; ri < cprsDesc.length; ri++) {
          if (cprsDesc[ri].teacher) {
            cprsTeacher = parseTeacherName(cprsDesc[ri].teacher);
            break;
          }
        }
      }
      var teacherName = cprsTeacher || parseTeacherName(teacherRaw);
      var teacherCode = parseTeacherCode(teacherRaw);

      // ── Course history: HubSpot property history (ground truth) ────────
      // HubSpot audits every change to current_course__t_ with a timestamp.
      // This is far more reliable than CPRS (which only covers the export window).
      // Fall back to CPRS-based logic if HubSpot history is unavailable.
      var hsHistory = (healthInfo && healthInfo.courseHistory) || [];
      var coursesWithTeacher;
      var courseNumberWithTeacher;
      var prmTypeLow = (prms.prmType || '').toLowerCase().trim();
      var isRenewal  = prmTypeLow.indexOf('renewal') > -1
                    || (prmTypeLow.indexOf('term') > -1 && prmTypeLow.indexOf('first') === -1);

      if (hsHistory.length > 0) {
        // ── Primary: HubSpot course history ─────────────────────────
        coursesWithTeacher      = hsHistory;         // full journey, oldest→newest
        courseNumberWithTeacher = hsHistory.length;  // exact count
      } else {
        // ── Fallback: derive from CPRS rows ─────────────────────────
        var teacherNormLocal = normalizeTeacherName(teacherName);
        coursesWithTeacher = [];
        var cprsChron = cprsRows.slice().sort(function(a, b) {
          var da = a.date ? a.date.getTime() : 0;
          var db = b.date ? b.date.getTime() : 0;
          return da - db;
        });
        cprsChron.forEach(function(r) {
          if (!r.happened) return;
          var rowTeacher = normalizeTeacherName(parseTeacherName(r.teacher));
          if (rowTeacher !== teacherNormLocal) return;
          var cName = (r.course || '').trim();
          if (!cName) return;
          if (coursesWithTeacher.length === 0 || coursesWithTeacher[coursesWithTeacher.length - 1] !== cName) {
            if (coursesWithTeacher.indexOf(cName) === -1) coursesWithTeacher.push(cName);
          }
        });
        var courseNumberFromCPRS = Math.max(1, coursesWithTeacher.length || 1);
        // PRMS prmType fallback: Renewal / Term (not First Term) → 2nd+ course
        courseNumberWithTeacher = courseNumberFromCPRS >= 2
          ? courseNumberFromCPRS
          : (isRenewal ? 2 : 1);
      }

      // Cancellation stats (last 30 days)
      var cutoff30 = new Date(); cutoff30.setDate(cutoff30.getDate() - 30);
      var recentCancels = cprsRows.filter(function(r){ return !r.happened && r.date && r.date >= cutoff30; });
      var cancelReasonSet = {};
      recentCancels.forEach(function(r){ if (r.cancelReason) cancelReasonSet[r.cancelReason] = true; });

      // Next courses with teacher upskill check
      var nextCourses = prms.nextCourses.map(function(nc) {
        var prog = _getTeacherCourseProgressStandalone(teacherName, nc.name);
        var ps   = _parseProgressStatus(prog);
        return {
          name           : nc.name,
          count          : nc.count,
          teacherProgress: ps.displayLabel,
          progressKnown  : ps.progressKnown,
          teacherReady   : ps.teacherReady,
          inProgress     : ps.inProgress,
          notStarted     : ps.notStarted
        };
      });

      var teacherReadyForNext = nextCourses.length > 0 && nextCourses[0].teacherReady;

      // Migration needed when teacher not found in sheet OR confirmed 0%/Not Onboarded on next course
      // AND learner is on 2nd+ course with same teacher.
      // In-progress upskilling (1–70%) → fast-pace upskill, NOT a migration case.
      var nc0 = nextCourses[0];
      // Migration needed when teacher is NOT ready (< 71%) for next course AND learner is on 2nd+ course.
      // Covers: not in sheet, Not Onboarded (0%), AND in-progress (1–70%) — all are migration risk.
      // Only excluded: teacher ≥ 71% (teacherReady = true) → same-teacher course transition.
      var migrationNeeded = !!nc0
                         && !nc0.teacherReady   // not in sheet OR < 71% upskilled
                         && courseNumberWithTeacher >= 2;
      var teacherGapCourses = nextCourses.filter(function(nc){ return nc.progressKnown && !nc.teacherReady; }).length;

      // Debug log — remove after confirming CLS button appears
      if (courseNumberWithTeacher >= 2 || (nc0 && !nc0.progressKnown)) {
        Logger.log('[LP] migrationNeeded check for ' + jlid + ' (' + prms.learnerName + '): ' +
          'courseNum=' + courseNumberWithTeacher +
          ' nc0=' + (nc0 ? (nc0.name + ' progressKnown=' + nc0.progressKnown + ' notStarted=' + nc0.notStarted) : 'none') +
          ' → migrationNeeded=' + migrationNeeded);
      }

      // Health + deal info  (healthInfo already defined at top of loop)
      var healthCritical = healthInfo.critical || false;
      var clsRequired    = healthCritical;
      var dealId         = healthInfo.dealId    || '';
      var dealStage      = healthInfo.dealStage || '';
      var dealName       = healthInfo.dealName  || '';
      // Renewal detection: closedate epoch ms + PRMS prmType
      var renewalApproaching = false;
      var daysToRenewal      = null;
      var isRenewalDeal      = prmTypeLow.indexOf('renewal') > -1;
      if (healthInfo.closeDate && healthInfo.closeDate > 0) {
        var closeDateObj = new Date(healthInfo.closeDate);
        daysToRenewal    = Math.floor((closeDateObj.getTime() - today.getTime()) / 86400000);
        renewalApproaching = daysToRenewal >= 0 && daysToRenewal <= 30;
      }

      // Alert level — 1st course learners never go above 'warning' on time alone,
      // and never trigger migration alerts (CCTC is 2nd course+ only).
      var alertLevel = 'ok';
      if (projectedWeeks != null) {
        if (courseNumberWithTeacher >= 2 && projectedWeeks <= LP_ALERT_WEEKS_CRITICAL) alertLevel = 'critical';
        else if (projectedWeeks <= LP_ALERT_WEEKS_WARNING) alertLevel = 'warning';
      }
      // Boost to warning if teacher not ready AND on 2nd+ course within warning window
      if (alertLevel === 'ok' && migrationNeeded && projectedWeeks != null && projectedWeeks <= LP_ALERT_WEEKS_WARNING) alertLevel = 'warning';

      learners.push({
        jlid                 : jlid,
        learnerName          : prms.learnerName,
        teacher              : teacherName,
        teacherCode          : teacherCode,
        teacherRaw           : teacherRaw,
        currentCourse        : prms.currentCourse,
        classesLeft          : classesLeft,
        classesLeftRaw       : classesLeftRaw,       // original PRMS number
        sessionsDoneAfterPRM : sessionsDoneAfterPRM, // how many subtracted
        classesDone          : classesDone,
        classFrequency       : frequency,
        projectedWeeks       : projectedWeeks,
        projectedDate        : projectedDate ? Utilities.formatDate(projectedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : null,
        prmStaleDays         : prmStaleDays,
        subject              : subject,              // 'coding' or 'maths'
        alertLevel           : alertLevel,
        migrationNeeded      : migrationNeeded,
        teacherReadyForNext  : teacherReadyForNext,
        teacherGapCourses    : teacherGapCourses,
        nextCourses          : nextCourses,
        nextCoursesFinalized : prms.finalized,
        prmType              : prms.prmType,
        clsNotes             : prms.clsNotes,
        prmDate              : prms.prmDate ? Utilities.formatDate(prms.prmDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : null,
        learnerHealthCritical: healthCritical,
        learnerHealth        : healthInfo.health || '',
        clsApprovalRequired  : clsRequired,
        classSessions        : healthInfo.classSessions || [],   // [{day,time}]
        classTimingsRaw      : healthInfo.classTimingsRaw || '',
        dealId               : dealId,
        dealStage            : dealStage,
        dealName             : dealName,
        isRenewalDeal        : isRenewalDeal,
        renewalApproaching   : renewalApproaching,
        daysToRenewal        : daysToRenewal,
        totalCourseClasses   : totalCourseClasses,
        courseNumberWithTeacher: courseNumberWithTeacher,
        coursesWithTeacher   : coursesWithTeacher,
        recentCancellations  : recentCancels.length,
        cancelReasons        : Object.keys(cancelReasonSet)
      });
    });

    // Sort: critical first, then warning, then ok; within same level sort by projectedWeeks ascending
    var levelOrder = { critical: 0, warning: 1, ok: 2 };
    learners.sort(function(a, b) {
      var la = levelOrder[a.alertLevel] || 2, lb = levelOrder[b.alertLevel] || 2;
      if (la !== lb) return la - lb;
      var wa = a.projectedWeeks != null ? a.projectedWeeks : 999;
      var wb = b.projectedWeeks != null ? b.projectedWeeks : 999;
      return wa - wb;
    });

    var roadmapGaps = learners.filter(function(l){ return l.migrationNeeded; }).length;
    var stats = {
      total           : learners.length,
      criticalCount   : learners.filter(function(l){ return l.alertLevel === 'critical'; }).length,
      warningCount    : learners.filter(function(l){ return l.alertLevel === 'warning'; }).length,
      migrationsDue   : learners.filter(function(l){ return l.migrationNeeded && l.alertLevel !== 'ok'; }).length,
      roadmapGaps     : roadmapGaps,   // teachers missing upskill for ≥1 next course
      teacherGapCount : roadmapGaps,   // alias kept for backward compat
      unfinalizedCount: learners.filter(function(l){ return !l.nextCoursesFinalized; }).length
    };

    Logger.log('[LP] getLearnerProgressions done: ' + learners.length + ' learners, ' + stats.criticalCount + ' critical');
    var result = { success: true, learners: learners, stats: stats };
    _lpCacheStore(result);  // cache for next 10 min
    return result;
  } catch(e) {
    Logger.log('[LP] getLearnerProgressions error: ' + e.message);
    return { success: false, message: e.message, learners: [], stats: {} };
  }
}

// ─────────────────────────────────────────────────────────────────────
// getUpskillGapsForRoadmap(jlid)
// Returns teacher upskill status for each next course in roadmap.
// ─────────────────────────────────────────────────────────────────────
/**
 * @param {string} jlid
 * @param {Array}  learnerSlots  [{day, time}] — learner's class schedule from HubSpot
 * @param {string} learnerName   — for calendar lookup
 */
function getUpskillGapsForRoadmap(jlid, learnerSlots, learnerName) {
  try {
    var prmsMap = _buildPRMSMap();
    var prms    = prmsMap[jlid];
    if (!prms) return { success: false, message: 'Learner not found in PRMS data.' };

    var teacherName  = parseTeacherName(prms.teacher);
    var rawSlots     = learnerSlots || [];
    // Convert day-of-week slots → upcoming calendar dates for searchMatchingTeachers
    var dateSlots    = _convertDayTimeSlotsToDateSlots(rawSlots);

    Logger.log('[LP] getUpskillGapsForRoadmap ' + jlid + ' rawSlots=' + JSON.stringify(rawSlots) + ' dateSlots=' + JSON.stringify(dateSlots));

    // Calendar schedule: use passed slots if available (already fetched in main getLearnerProgressions).
    // Only fall back to live calendar fetch when no slots were passed — avoids fetching ALL events
    // (can be 500+ rows for hello@jet-learn.com) every time a panel opens.
    var calSchedule = { found: false };
    if (rawSlots.length > 0) {
      // Slots already available from HubSpot deal data — no calendar call needed
      calSchedule = { found: true, slots: rawSlots.map(function(s){ return { day: s.day, timeCET: s.time, count: 1 }; }), primary: rawSlots[0] || null };
      Logger.log('[LP] calSchedule for ' + jlid + ': using passed slots (' + rawSlots.length + ')');
    } else {
      var lName = learnerName || prms.learnerName || '';
      if (lName) {
        calSchedule = getLearnerCalendarSchedule(jlid, lName);
        Logger.log('[LP] calSchedule for ' + jlid + ': calendar fetch, found=' + calSchedule.found);
      }
    }

    var gaps = prms.nextCourses.map(function(nc, idx) {
      var prog = _getTeacherCourseProgressStandalone(teacherName, nc.name);
      var ps   = _parseProgressStatus(prog);
      var progressKnown = ps.progressKnown;
      var teacherReady  = ps.teacherReady;
      var inProgress    = ps.inProgress;
      var notStarted    = ps.notStarted;

      var matchedTeachers  = [];   // upskilled + confirmed slot match
      var anyUpskilled     = [];   // upskilled, no slot check
      var altCount         = 0;

      // Teacher search skipped here — too slow for panel open (opens separate spreadsheet).
      // Full search runs when "Trigger Migration" / "CLS Review" is clicked via triggerSmartMigration.

      return {
        index                    : idx + 1,
        course                   : nc.name,
        count                    : nc.count,
        teacherProgress          : ps.displayLabel,
        progressKnown            : progressKnown,
        teacherReady             : teacherReady,
        inProgress               : inProgress,
        notStarted               : notStarted,
        requiresAction           : progressKnown && !teacherReady,
        alternativeTeachersCount : altCount,
        matchedTeachers          : matchedTeachers,
        anyUpskilled             : anyUpskilled
      };
    });

    return {
      success      : true,
      teacher      : teacherName,
      jlid         : jlid,
      gaps         : gaps,
      calSchedule  : calSchedule,                 // {found, slots:[{day,timeCET,count}], primary}
      rawSlots     : rawSlots,                     // original HubSpot slots for display
      dateSlots    : dateSlots                     // converted for debug
    };
  } catch(e) {
    Logger.log('[LP] getUpskillGapsForRoadmap error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────────────
// requestTPUpskilling(jlid, teacherName, courseName)
// Creates a HubSpot upskill task for the TP Manager and sends an email.
// ─────────────────────────────────────────────────────────────────────
function requestTPUpskilling(jlid, teacherName, courseName, fastPace) {
  try {
    Logger.log('[LP] requestTPUpskilling: ' + jlid + ' | ' + teacherName + ' | ' + courseName + ' | fastPace=' + fastPace);
    if (!teacherName || !courseName) return { success: false, message: 'Teacher and course required.' };

    // Get TP Manager email from teacher record
    var teacherData   = getTeacherData ? getTeacherData() : [];
    var teacherRecord = teacherData.find(function(t) {
      return String(t.name||'').trim().toLowerCase() === String(teacherName).trim().toLowerCase();
    });
    var tpManagerName  = (teacherRecord && teacherRecord.manager)          || '';
    var tpManagerEmail = (teacherRecord && teacherRecord.tpManagerEmail)   || '';
    var learnerName    = '';
    try {
      var prmsMap = _buildPRMSMap();
      if (prmsMap[jlid]) learnerName = prmsMap[jlid].learnerName || '';
    } catch(e) {}

    // Create HubSpot upskill task
    try {
      var dealId = '';
      try {
        var hsRes = fetchHubspotByJlid(jlid);
        if (hsRes.success) dealId = hsRes.data.dealId || '';
      } catch(he) {}
      var tpManagerHsId = tpManagerName ? (getHubSpotOwnerIdByName(tpManagerName) || '') : '';
      createUpskillTaskOnHubSpot(
        jlid, teacherName, learnerName || jlid,
        tpManagerHsId, [courseName], dealId
      );
    } catch(hte) {
      Logger.log('[LP] Upskill HubSpot task failed: ' + hte.message);
    }

    // Send email to TP Manager if we have their email
    if (tpManagerEmail) {
      try {
        var urgency = fastPace ? '[FAST-PACE] ' : '[UPSKILL NEEDED] ';
        var subject = urgency + teacherName + ' — ' + courseName;
        var body = 'Hi ' + (tpManagerName || 'TP Manager') + ',\n\n'
          + 'This is an automated request from the JetLearn Course Planner.\n\n'
          + 'Learner : ' + (learnerName || jlid) + '\n'
          + 'Teacher : ' + teacherName + '\n'
          + 'Course  : ' + courseName + '\n'
          + 'Priority: ' + (fastPace ? 'FAST-PACE — teacher has started but is below threshold. Please accelerate.' : 'URGENT START — teacher has not begun this course yet.') + '\n\n'
          + 'The learner\'s roadmap requires the above course next. '
          + (fastPace
              ? 'Please prioritise completing ' + teacherName + '\'s upskilling on "' + courseName + '" before the current course ends.'
              : 'Please arrange upskilling for ' + teacherName + ' on "' + courseName + '" immediately — current course is nearing completion.')
          + '\n\nRegards,\nJetLearn CLS System';
        GmailApp.sendEmail(tpManagerEmail, subject, body, { name: 'JetLearn CLS' });
      } catch(me) {
        Logger.log('[LP] TP upskill email failed: ' + me.message);
      }
    }

    try {
      var actionType = fastPace ? 'TP Fast-pace Upskill Request' : 'TP Upskill Request';
      logAction(actionType, jlid, learnerName, teacherName, '', courseName, 'Requested',
        (fastPace ? 'Fast-pace' : 'Upskill') + ' request sent to TP: ' + (tpManagerName || 'unknown'), 'Roadmap Gap', '');
    } catch(ae) {}

    return {
      success  : true,
      fastPace : !!fastPace,
      message  : (fastPace ? 'Fast-pace upskill request sent' : 'Upskill request sent')
                 + (tpManagerEmail ? ' to ' + tpManagerEmail : '') + ' for ' + teacherName + ' on ' + courseName
    };
  } catch(e) {
    Logger.log('[LP] requestTPUpskilling error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Fetch HubSpot enum property options (cached 6h) ──────────────────
function _getHubspotEnumOptions(propertyName) {
  var cacheKey = 'HS_PROP_' + propertyName;
  var sc = CacheService.getScriptCache();
  var cached = sc.get(cacheKey);
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    var resp  = UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/properties/tickets/' + propertyName, {
      headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return [];
    var data    = JSON.parse(resp.getContentText());
    var options = (data.options || []).map(function(o){ return { value: o.value, label: o.label }; });
    sc.put(cacheKey, JSON.stringify(options), 21600);
    return options;
  } catch(e) { return []; }
}

// ── Match a free-text course name to a HubSpot enum option value ──────
function _matchCourseToHubspotOption(courseName, options) {
  if (!courseName || !options || options.length === 0) return '';
  var norm = function(s){ return String(s).toLowerCase().replace(/[^a-z0-9]/g,''); };
  var cn   = norm(courseName);
  // 1. Exact value match
  for (var i = 0; i < options.length; i++) { if (norm(options[i].value) === cn) return options[i].value; }
  // 2. Exact label match
  for (var i = 0; i < options.length; i++) { if (norm(options[i].label) === cn) return options[i].value; }
  // 3. Contains match (longer wins)
  var best = ''; var bestLen = 0;
  options.forEach(function(o) {
    var ol = norm(o.label);
    if ((cn.indexOf(ol) > -1 || ol.indexOf(cn) > -1) && ol.length > bestLen) { best = o.value; bestLen = ol.length; }
  });
  return best;
}

// ── Pick migration_timeline enum value based on weeks to completion ───
function _pickMigrationTimeline(projectedWeeks, options) {
  if (!options || options.length === 0) return '';
  // Pick the option whose label contains the closest week count
  var wks = projectedWeeks != null ? Math.round(projectedWeeks) : 4;
  // Sort options by the number in their label; pick closest
  var scored = options.map(function(o) {
    var m = String(o.label).match(/(\d+)/);
    var n = m ? parseInt(m[1]) : 99;
    return { value: o.value, label: o.label, n: n, dist: Math.abs(n - wks) };
  }).sort(function(a, b){ return a.dist - b.dist; });
  return scored.length > 0 ? scored[0].value : '';
}

// ─────────────────────────────────────────────────────────────────────
// triggerSmartMigration(jlid)
// Finds best-matched teachers for learner's next course + creates
// a HubSpot migration ticket on pipeline 66161281.
// ─────────────────────────────────────────────────────────────────────
// preMatchedTeachers — optional, pass from handleCLSMigrationApproval to skip slow re-search.
// Teachers already found when Slack CLS Review message was sent; cached in ScriptCache.
function triggerSmartMigration(jlid, preMatchedTeachers) {
  try {
    Logger.log('[LP] triggerSmartMigration: ' + jlid);
    if (!jlid) return { success: false, message: 'JLID required.' };

    // 1. Get learner progression data (from cache — fast)
    var prog = getLearnerProgressions();
    if (!prog.success) return { success: false, message: 'Failed to load progressions: ' + prog.message };
    var learner = null;
    for (var i = 0; i < prog.learners.length; i++) {
      if (prog.learners[i].jlid === jlid) { learner = prog.learners[i]; break; }
    }
    if (!learner) return { success: false, message: 'Learner ' + jlid + ' not found in PRMS data.' };

    var nextCourse  = learner.nextCourses.length > 0 ? learner.nextCourses[0].name : learner.currentCourse;
    var nc1         = learner.nextCourses.length > 1 ? learner.nextCourses[1].name : '';
    var nc2         = learner.nextCourses.length > 2 ? learner.nextCourses[2].name : '';

    // 2. Use pre-matched teachers (from CLS Review send) or skip search entirely.
    // searchMatchingTeachers opens a separate spreadsheet and takes 3-4 min — too slow here.
    var matchedTeachers = preMatchedTeachers || [];
    if (matchedTeachers.length === 0) {
      // Try cache from when CLS Review was sent
      try {
        var cached = CacheService.getScriptCache().get('CLS_TEACHERS_' + jlid);
        if (cached) matchedTeachers = JSON.parse(cached);
      } catch(ce) {}
    }
    Logger.log('[LP] triggerSmartMigration: ' + matchedTeachers.length + ' pre-matched teachers for ' + jlid);

    // 3. Create HubSpot ticket
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    if (!token) return { success: false, message: 'HubSpot API key not configured.' };

    var subjectPrefix = learner.migrationNeeded ? '[CCTC] ' : '[Course Transition] ';
    var subject = subjectPrefix + learner.learnerName + ' — ' + learner.currentCourse + ' → ' + nextCourse;

    var tz      = Session.getScriptTimeZone();
    var today   = Utilities.formatDate(new Date(), tz, 'dd-MMM-yyyy');
    var projStr = learner.projectedDate ? Utilities.formatDate(new Date(learner.projectedDate), tz, 'dd-MMM-yyyy') : 'Unknown';
    var freq    = learner.classFrequency ? (Math.round(learner.classFrequency * 10) / 10) + ' class/wk' : '—';

    // ── Notes: full ticket details — CLS fills enum dropdowns manually ────
    var sep = '─────────────────────────────────────────';
    var notesLines = [
      'Auto-triggered by Course Planner  |  ' + today,
      sep,
      '',
      '📋 MIGRATION DETAILS',
      'Learner       : ' + learner.learnerName + ' (' + jlid + ')',
      'Current Course: ' + learner.currentCourse,
      'Classes Left  : ' + learner.classesLeft + '  |  Frequency: ' + freq,
      'Est. Completion: ' + projStr + ' (~' + (learner.projectedWeeks != null ? Math.round(learner.projectedWeeks) : '?') + ' weeks)',
      '',
      '📚 COURSE ROADMAP',
      'Current   → ' + learner.currentCourse,
      'Next (1)  → ' + nextCourse + (learner.nextCourses[0] ? '  [Teacher: ' + (learner.nextCourses[0].teacherProgress || 'Unknown') + ']' : ''),
      nc1 ? ('Next (2)  → ' + nc1) : '',
      nc2 ? ('Next (3)  → ' + nc2) : '',
      '',
      '👩‍🏫 CURRENT TEACHER',
      'Teacher: ' + (learner.teacher || '—') + (learner.teacherCode ? ' (' + learner.teacherCode + ')' : ''),
      'Upskill on Next Course: ' + (learner.nextCourses[0] ? learner.nextCourses[0].teacherProgress : '—'),
      'Migration Type: ' + (learner.migrationNeeded ? 'CCTC (teacher not ready < 71%)' : 'Course Transition (teacher ready)'),
      '',
      '❤️ LEARNER HEALTH',
      'Health: ' + (learner.learnerHealth || 'Not set') + (learner.learnerHealthCritical ? '  ⚠️ CRITICAL' : ''),
      'Deal Stage: ' + (learner.dealStage || '—'),
      '',
      sep,
      '🏆 PRE-MATCHED TEACHERS FOR: ' + nextCourse,
    ];
    if (matchedTeachers.length > 0) {
      matchedTeachers.forEach(function(t, idx) {
        var slots = t.alternateSlots || t.slotMatch || '—';
        notesLines.push('  ' + (idx+1) + '. ' + t.name
          + '  |  Grade: ' + (t.auditGrade || '—')
          + '  |  ' + nextCourse + ': ' + (t.courseReady || t.currentCourseProgress || 'N/A')
          + '  |  Available: ' + slots);
      });
    } else {
      notesLines.push('  (No pre-matched teachers — run Teacher Intelligence for this learner\'s slot)');
    }
    // Remove empty strings
    var filteredNotes = notesLines.filter(function(l){ return l !== null && l !== undefined; });

    // ── Fetch HubSpot enum options for auto-fill ──────────────────────
    var courseOpts    = _getHubspotEnumOptions('current_course__t_');
    var fc1Opts       = _getHubspotEnumOptions('future_course_1');
    var fc2Opts       = _getHubspotEnumOptions('future_course_2');
    var fc3Opts       = _getHubspotEnumOptions('future_course_3');
    var tlOpts        = _getHubspotEnumOptions('migration_timeline');
    var tlfOpts       = _getHubspotEnumOptions('timeline_of_form_filled');

    var curCourseVal  = _matchCourseToHubspotOption(learner.currentCourse, courseOpts);
    var fc1Val        = _matchCourseToHubspotOption(nextCourse, fc1Opts);
    var fc2Val        = nc1 ? _matchCourseToHubspotOption(nc1, fc2Opts) : '';
    var fc3Val        = nc2 ? _matchCourseToHubspotOption(nc2, fc3Opts) : '';
    var tlVal         = _pickMigrationTimeline(learner.projectedWeeks, tlOpts);
    var tlfVal        = _pickMigrationTimeline(learner.projectedWeeks, tlfOpts);

    Logger.log('[LP] Enum auto-fill: curCourse=' + curCourseVal + ' fc1=' + fc1Val + ' fc2=' + fc2Val + ' timeline=' + tlVal);

    // Only send text/ID fields on creation to avoid 400 from invalid enum values.
    // Enum fields (migration_timeline, timeline_of_form_filled, are_you_trained_on_the_next_course_)
    // and date picker fields (pre_migration_last_class_conducted_date__t_) are omitted here —
    // CLS will fill them manually, or update via a follow-up PATCH once enum values are confirmed.
    // Build ticket properties — enum fields auto-filled via HubSpot property API lookup.
    // Only included when a valid match was found; notes always have the full details as fallback.
    var ticketProps = {
      subject                : subject,
      hs_pipeline            : '66161281',
      hs_pipeline_stage      : '128913747',
      content                : filteredNotes.join('\n'),
      learner_uid            : jlid,
      reason_of_migration__t_: 'Course Change Teacher Change'
    };
    if (curCourseVal) ticketProps['current_course__t_']      = curCourseVal;
    if (fc1Val)       ticketProps['future_course_1']         = fc1Val;
    if (fc2Val)       ticketProps['future_course_2']         = fc2Val;
    if (fc3Val)       ticketProps['future_course_3']         = fc3Val;
    if (tlVal)        ticketProps['migration_timeline']      = tlVal;
    if (tlfVal)       ticketProps['timeline_of_form_filled'] = tlfVal;
    var ticketPayload = { properties: ticketProps };

    var ticketResp = UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/objects/tickets', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(ticketPayload),
      muteHttpExceptions: true
    });

    var ticketCode = ticketResp.getResponseCode();
    var ticketBody = ticketResp.getContentText();
    if (ticketCode !== 201 && ticketCode !== 200) {
      Logger.log('[LP] Ticket creation failed (' + ticketCode + '): ' + ticketBody);
      // Parse HubSpot error for specific field that failed
      var errMsg = 'HubSpot ticket creation failed (' + ticketCode + ').';
      try {
        var errData = JSON.parse(ticketBody);
        if (errData.message) errMsg += ' ' + errData.message;
        if (errData.errors && errData.errors.length > 0) {
          errMsg += ' Fields: ' + errData.errors.map(function(e){ return e.context ? JSON.stringify(e.context) : e.message; }).join('; ');
        }
      } catch(pe) {}
      return { success: false, message: errMsg };
    }

    var ticketData = JSON.parse(ticketResp.getContentText());
    var ticketId   = ticketData.id || '';
    var hsLink     = ticketId ? 'https://app.hubspot.com/contacts/7729491/record/0-5/' + ticketId : '';

    Logger.log('[LP] Ticket created: ' + ticketId + ' for ' + jlid);
    return {
      success         : true,
      ticketId        : ticketId,
      hubspotLink     : hsLink,
      subject         : subject,
      matchedTeachers : matchedTeachers,
      clsApprovalRequired: learner.clsApprovalRequired,
      learner         : learner
    };
  } catch(e) {
    Logger.log('[LP] triggerSmartMigration error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────────────
// getCoursePlannerSummary()
// Lightweight stats-only call for sidebar badge (no full learner list).
// ─────────────────────────────────────────────────────────────────────
function getCoursePlannerSummary() {
  try {
    var result = getLearnerProgressions();
    if (!result.success) return { success: true, criticalCount: 0, warningCount: 0 };
    return { success: true, stats: result.stats };
  } catch(e) {
    return { success: true, stats: { criticalCount: 0, warningCount: 0 } };
  }
}

// ─────────────────────────────────────────────────────────────────────
// HubSpot Migration Form Submission
// Form: https://share.hsforms.com/1RZKGkUG6Qc2kLHBVlTOmPw4lo43
// Portal: 7729491
// ─────────────────────────────────────────────────────────────────────

var LP_HS_PORTAL_ID           = '7729491';
var LP_MIGRATION_FORM_SHARE   = 'https://share.hsforms.com/1RZKGkUG6Qc2kLHBVlTOmPw4lo43';
var LP_FORM_GUID_CACHE_KEY    = 'LP_MIG_FORM_GUID';
var LP_FORM_FIELDS_CACHE_KEY  = 'LP_MIG_FORM_FIELDS';

// ── Fetch migration form GUID from HubSpot (cached 6 hrs) ────────────
function _findMigrationFormGuid() {
  var sc = CacheService.getScriptCache();
  var cached = sc.get(LP_FORM_GUID_CACHE_KEY);
  if (cached) return cached;

  // 1. Check Script Properties for manually set GUID (highest priority)
  var hardcoded = PropertiesService.getScriptProperties().getProperty('LP_MIGRATION_FORM_GUID');
  if (hardcoded) {
    Logger.log('[LP] Using hardcoded form GUID from Script Properties: ' + hardcoded);
    sc.put(LP_FORM_GUID_CACHE_KEY, hardcoded, 21600);
    return hardcoded;
  }

  // 2. Extract GUID from the HubSpot share URL (works without Forms API scope)
  // The share form page embeds the GUID in its HTML as "formId":"xxxxxxxx-..."
  try {
    Logger.log('[LP] Fetching share URL to extract form GUID: ' + LP_MIGRATION_FORM_SHARE);
    var shareResp = UrlFetchApp.fetch(LP_MIGRATION_FORM_SHARE, {
      muteHttpExceptions: true,
      followRedirects: true
    });
    if (shareResp.getResponseCode() === 200) {
      var html = shareResp.getContentText();
      // HubSpot embeds GUID in multiple ways — try each
      var guidPatterns = [
        /"formId"\s*:\s*"([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})"/i,
        /"guid"\s*:\s*"([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})"/i,
        /formId=([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i,
        /\/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})(?:\/|"|\?)/i
      ];
      for (var pi = 0; pi < guidPatterns.length; pi++) {
        var m = html.match(guidPatterns[pi]);
        if (m && m[1]) {
          Logger.log('[LP] Extracted GUID from share URL (pattern ' + pi + '): ' + m[1]);
          sc.put(LP_FORM_GUID_CACHE_KEY, m[1], 21600);
          return m[1];
        }
      }
      Logger.log('[LP] Share URL fetched (' + html.length + ' chars) but GUID pattern not found. HTML snippet: ' + html.substring(0, 500));
    } else {
      Logger.log('[LP] Share URL returned HTTP ' + shareResp.getResponseCode());
    }
  } catch(shareErr) {
    Logger.log('[LP] Share URL fetch error: ' + shareErr.message);
  }

  // 3. Try Forms API (requires forms scope — may fail if token lacks it)
  var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  try {
    var resp = UrlFetchApp.fetch('https://api.hubapi.com/forms/v2/forms?limit=500', {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() === 200) {
      var forms = JSON.parse(resp.getContentText());
      for (var i = 0; i < forms.length; i++) {
        var f = forms[i];
        if (f.name && /migration/i.test(f.name)) {
          Logger.log('[LP] Found migration form (v2 API): "' + f.name + '" GUID: ' + f.guid);
          sc.put(LP_FORM_GUID_CACHE_KEY, f.guid, 21600);
          return f.guid;
        }
      }
      Logger.log('[LP] v2 API: ' + forms.length + ' forms, none named "migration". Names: ' + forms.slice(0,20).map(function(x){return x.name;}).join(' | '));
    } else {
      Logger.log('[LP] Forms v2 API returned ' + resp.getResponseCode() + ' — token may lack forms scope');
    }
  } catch(apiErr) {
    Logger.log('[LP] Forms API error: ' + apiErr.message);
  }

  return null;
}

// ── Debug: run this manually in GAS editor to find the form GUID ────
function debugFindMigrationFormGuid() {
  // Clear cached result first so we always do fresh lookup
  var sc = CacheService.getScriptCache();
  sc.remove(LP_FORM_GUID_CACHE_KEY);
  var guid = _findMigrationFormGuid();
  Logger.log('[DEBUG] Form GUID result: ' + (guid || 'NOT FOUND'));
  Logger.log('[DEBUG] To set manually: GAS Project Settings → Script Properties → LP_MIGRATION_FORM_GUID = ' + (guid || 'YOUR-GUID-HERE'));
  return guid;
}

// ── Get form field definitions (to verify internal names) ────────────
function getMigrationFormFields() {
  var token  = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  var guid   = _findMigrationFormGuid();
  if (!guid) return { success: false, message: 'Migration form not found in HubSpot.' };

  try {
    var resp = UrlFetchApp.fetch('https://api.hubapi.com/forms/v2/forms/' + guid, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    var form = JSON.parse(resp.getContentText());
    var fields = [];
    (form.formFieldGroups || []).forEach(function(grp) {
      (grp.fields || []).forEach(function(f) {
        fields.push({
          name       : f.name,
          label      : f.label,
          fieldType  : f.fieldType,
          required   : f.required,
          options    : (f.options || []).map(function(o){ return o.value; })
        });
      });
    });
    return { success: true, formName: form.name, guid: guid, fields: fields };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ── Build migration form field payload from learner data ─────────────
function _buildMigrationFormFields(learner, userEmail) {
  var tz          = Session.getScriptTimeZone();
  var nextCourses = learner.nextCourses || [];
  var today       = new Date();

  // Pre-migration last class date = projected completion date
  var projDate    = learner.projectedDate ? new Date(learner.projectedDate) : null;
  // If projected date is null, estimate 4 weeks from today
  if (!projDate || isNaN(projDate.getTime())) projDate = new Date(today.getTime() + 28 * 86400000);
  var projDateStr = Utilities.formatDate(projDate, tz, 'MM/dd/yyyy'); // HubSpot date format

  var notesText   = [
    'Auto-triggered by JetLearn Course Planner',
    'Trigger date: ' + Utilities.formatDate(today, tz, 'dd-MMM-yyyy'),
    'Classes remaining: ' + (learner.classesLeft != null ? learner.classesLeft : '—'),
    'Class frequency: '   + (learner.classFrequency ? learner.classFrequency + '/wk' : '—'),
    'Estimated completion: ' + Utilities.formatDate(projDate, tz, 'dd-MMM-yyyy'),
    'Teacher ready for next course: ' + (learner.teacherReadyForNext ? 'Yes' : 'No'),
    learner.learnerHealthCritical ? 'CRITICAL LEARNER — CLS approval required' : ''
  ].filter(Boolean).join('\n');

  // Field name → value map
  // NOTE: Internal HubSpot field names discovered via getMigrationFormFields()
  // Adjust names below if submission returns field validation errors
  return [
    { name: 'email',                                        value: userEmail || '' },
    { name: 'learner_uid',                                  value: learner.jlid || '' },
    { name: 'reason_of_migration__t_',                      value: 'Course Change Teacher Change' },
    { name: 'migration_timeline',                           value: 'After Current Course Completion' },
    { name: 'timeline_of_form_filled',                      value: '4 weeks before Course Completion' },
    { name: 'pre_migration_last_class_conducted_date__t_',  value: projDateStr },
    { name: 'current_course__t_',                           value: learner.currentCourse || '' },
    { name: 'future_course_1',                              value: nextCourses[0] ? nextCourses[0].name : '' },
    { name: 'future_course_2',                              value: nextCourses[1] ? nextCourses[1].name : 'NA' },
    { name: 'future_course_3',                              value: nextCourses[2] ? nextCourses[2].name : 'NA' },
    // Always "Not Onboarded" when triggering migration — if teacher was upskilled, ticket auto-cancels per HS workflow
    { name: 'are_you_trained_on_the_next_course_',          value: 'Not Onboarded' },
    { name: 'notes___additional_comments',                  value: notesText || 'NA' }
  ].filter(function(f) { return f.value !== ''; });
}

// ── Build pre-fill URL (fallback if API submission fails) ─────────────
function _buildPrefillUrl(learner, userEmail) {
  var tz          = Session.getScriptTimeZone();
  var nextCourses = learner.nextCourses || [];
  var projDate    = learner.projectedDate ? new Date(learner.projectedDate) : new Date(new Date().getTime() + 28*86400000);
  var params = {
    email                                       : userEmail || '',
    learner_uid                                 : learner.jlid || '',
    reason_of_migration__t_                     : 'Course Change Teacher Change',
    migration_timeline                          : 'After Current Course Completion',
    timeline_of_form_filled                     : '4 weeks before the Course Completion',
    'pre_migration_last_class_conducted_date__t_': Utilities.formatDate(projDate, tz, 'MM/dd/yyyy'),
    current_course__t_                          : learner.currentCourse || '',
    future_course_1                             : nextCourses[0] ? nextCourses[0].name : '',
    future_course_2                             : nextCourses[1] ? nextCourses[1].name : 'NA',
    future_course_3                             : nextCourses[2] ? nextCourses[2].name : 'NA',
    // Always Not Onboarded — teacher upskilled = HS workflow auto-cancels ticket
    are_you_trained_on_the_next_course_         : 'Not Onboarded'
  };
  var qs = Object.keys(params).filter(function(k){ return params[k]; }).map(function(k) {
    return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
  }).join('&');
  return LP_MIGRATION_FORM_SHARE + (qs ? '?' + qs : '');
}

// ── Main: submit form to HubSpot + return preview data ───────────────
function submitMigrationHSForm(jlid, userEmail) {
  try {
    Logger.log('[LP] submitMigrationHSForm: ' + jlid + ' by ' + userEmail);

    // Get learner data
    var prog = getLearnerProgressions();
    if (!prog.success) return { success: false, message: 'Could not load progressions: ' + prog.message };

    var learner = null;
    for (var i = 0; i < prog.learners.length; i++) {
      if (prog.learners[i].jlid === jlid) { learner = prog.learners[i]; break; }
    }
    if (!learner) return { success: false, message: 'Learner ' + jlid + ' not found.' };

    var formFields    = _buildMigrationFormFields(learner, userEmail);
    var prefillUrl    = _buildPrefillUrl(learner, userEmail);
    var token         = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    var guid          = _findMigrationFormGuid();

    if (!guid) {
      Logger.log('[LP] Form GUID not found — returning prefill URL for manual submission');
      return {
        success       : false,
        formNotFound  : true,
        message       : 'Migration form not found in HubSpot (check form name contains "migration"). Use the pre-fill link to submit manually.',
        prefillUrl    : prefillUrl,
        fieldPreview  : formFields
      };
    }

    // Submit via HubSpot Forms v3 API
    var payload = {
      fields  : formFields,
      context : {
        pageUri  : 'https://jetlearn.com/course-planner',
        pageName : 'JetLearn Course Planner — Auto Migration Trigger'
      }
    };

    var resp = UrlFetchApp.fetch(
      'https://api.hsforms.com/submissions/v3/integration/submit/' + LP_HS_PORTAL_ID + '/' + guid,
      {
        method          : 'post',
        headers         : { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload         : JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );

    var code    = resp.getResponseCode();
    var body    = resp.getContentText();
    Logger.log('[LP] Form submission response (' + code + '): ' + body.substring(0, 300));

    if (code === 200 || code === 201) {
      return {
        success      : true,
        message      : 'Migration form submitted to HubSpot successfully.',
        fieldPreview : formFields,
        prefillUrl   : prefillUrl
      };
    }

    // Non-fatal: log field validation errors so we can fix names
    var errData = {};
    try { errData = JSON.parse(body); } catch(pe) {}
    var errMsg = errData.message || errData.error || body.substring(0, 400);

    // Still return field preview + prefill URL so user can submit manually
    return {
      success       : false,
      apiError      : true,
      message       : 'HubSpot form submission failed (' + code + '): ' + errMsg,
      prefillUrl    : prefillUrl,
      fieldPreview  : formFields
    };

  } catch(e) {
    Logger.log('[LP] submitMigrationHSForm error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── getMigrationFormPreview() — client calls this first to show form preview
// before user confirms submission (no side effects, just builds the payload)
function getMigrationFormPreview(jlid, userEmail) {
  try {
    var prog = getLearnerProgressions();
    if (!prog.success) return { success: false, message: prog.message };

    var learner = null;
    for (var i = 0; i < prog.learners.length; i++) {
      if (prog.learners[i].jlid === jlid) { learner = prog.learners[i]; break; }
    }
    if (!learner) return { success: false, message: 'Learner ' + jlid + ' not found.' };

    return {
      success      : true,
      fieldPreview : _buildMigrationFormFields(learner, userEmail),
      prefillUrl   : _buildPrefillUrl(learner, userEmail),
      learner      : learner
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════════════
// CLS APPROVAL WORKFLOW — Slack-based CCTC migration review
// ═══════════════════════════════════════════════════════════════════════
//
// Setup (one-time, in Script Properties):
//   SLACK_BOT_TOKEN   — Bot User OAuth Token from api.slack.com/apps
//                       Scopes needed: chat:write, chat:write.public
//   SLACK_CLS_CHANNEL — Channel ID or name (e.g. C0123ABCDEF or #cls-migrations)
//
// Slack App Interactivity:
//   Enable Interactivity → set Request URL to this GAS web app's deployed URL
//   (Deploy → Manage deployments → copy the /exec URL)
// ═══════════════════════════════════════════════════════════════════════

var CLS_Q = {
  ROW_ID      : 0,  // A
  JLID        : 1,  // B
  LEARNER     : 2,  // C
  TEACHER     : 3,  // D
  CUR_COURSE  : 4,  // E
  NEXT_COURSE : 5,  // F
  STATUS      : 6,  // G — Pending / Approved / Declined
  REASON      : 7,  // H
  SLACK_TS    : 8,  // I
  SLACK_CH    : 9,  // J
  TRIGGERED   : 10, // K
  RESOLVED    : 11, // L
  TRIGGERED_BY: 12  // M
};
var CLS_Q_SHEET = 'CLS Approval Queue';

// ── Sheet helpers ─────────────────────────────────────────────────────
function _getOrCreateCLSQueueSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.APP_DATA_SHEET_ID);
  var sh = ss.getSheetByName(CLS_Q_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CLS_Q_SHEET);
    sh.getRange(1, 1, 1, 13).setValues([[
      'Row ID','JLID','Learner Name','Teacher','Current Course','Next Course',
      'Status','Decline Reason','Slack TS','Slack Channel',
      'Triggered At','Resolved At','Triggered By'
    ]]).setFontWeight('bold');
    Logger.log('[CLS] Created CLS Approval Queue sheet');
  }
  return sh;
}

function _clsQueueGetEntry(jlid) {
  try {
    var sh   = _getOrCreateCLSQueueSheet();
    var data = sh.getDataRange().getValues();
    // Walk newest → oldest; return most recent row for this jlid
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][CLS_Q.JLID]) === String(jlid)) {
        return { rowIndex: i + 1, data: data[i] };
      }
    }
  } catch(e) { Logger.log('[CLS] _clsQueueGetEntry error: ' + e.message); }
  return null;
}

function _clsQueueUpsert(jlid, fields) {
  // fields: { learner, teacher, curCourse, nextCourse, status, reason, slackTs, slackCh, triggeredBy }
  try {
    var sh    = _getOrCreateCLSQueueSheet();
    var entry = _clsQueueGetEntry(jlid);
    var now   = new Date();
    if (entry) {
      var row = sh.getRange(entry.rowIndex, 1, 1, 13).getValues()[0];
      if (fields.status !== undefined) row[CLS_Q.STATUS]   = fields.status;
      if (fields.reason !== undefined) row[CLS_Q.REASON]   = fields.reason;
      if (fields.slackTs)              row[CLS_Q.SLACK_TS] = fields.slackTs;
      if (fields.slackCh)              row[CLS_Q.SLACK_CH] = fields.slackCh;
      if (fields.status && fields.status !== 'Pending') row[CLS_Q.RESOLVED] = now;
      sh.getRange(entry.rowIndex, 1, 1, 13).setValues([row]);
    } else {
      sh.appendRow([
        Utilities.getUuid().substring(0, 8).toUpperCase(),
        jlid,
        fields.learner     || '',
        fields.teacher     || '',
        fields.curCourse   || '',
        fields.nextCourse  || '',
        fields.status      || 'Pending',
        fields.reason      || '',
        fields.slackTs     || '',
        fields.slackCh     || '',
        now,
        '',
        fields.triggeredBy || ''
      ]);
    }
    Logger.log('[CLS] Queue upsert: ' + jlid + ' → ' + (fields.status || 'updated'));
    return true;
  } catch(e) {
    Logger.log('[CLS] _clsQueueUpsert error: ' + e.message);
    return false;
  }
}

// ── Build Slack Block Kit message ─────────────────────────────────────
function _buildCLSSlackBlocks(learner, matchedTeachers) {
  var alertEmoji  = { critical: '🔴', warning: '🟡', ok: '🟢' };
  var emoji       = alertEmoji[learner.alertLevel] || '⚪';
  var weeksText   = learner.projectedWeeks != null
    ? (Math.round(learner.projectedWeeks * 10) / 10) + ' wks' : '—';
  var dateText    = learner.projectedDate
    ? new Date(learner.projectedDate).toLocaleDateString('en-GB',{day:'numeric',month:'short',year:'numeric'})
    : '—';

  var courseOrdinal = (function(n) {
    var s = ['th','st','nd','rd'], v = n % 100;
    return n + (s[(v-20)%10] || s[v] || s[0]) + ' course together';
  })(learner.courseNumberWithTeacher || 1);

  var journeyStr = (learner.coursesWithTeacher || []).join(' → ') || (learner.currentCourse || '—');
  if (journeyStr.length > 120) journeyStr = journeyStr.substring(0, 117) + '\u2026';

  var nextLines = (learner.nextCourses || []).slice(0, 3).map(function(nc, i) {
    var icon = nc.teacherReady ? '\u2705' : '\u274c';
    var prog = nc.teacherReady ? 'Teacher ready'
      : 'Not upskilled' + (nc.teacherProgress && nc.teacherProgress !== 'Not Onboarded' ? ' (' + nc.teacherProgress + ')' : '');
    return (i+1) + '. ' + nc.name + ' \u2014 ' + icon + ' ' + prog;
  }).join('\n') || '\u2014';

  var teacherLines = (matchedTeachers || []).slice(0, 3).map(function(t, i) {
    var medals = ['\ud83e\udd47','\ud83e\udd48','\ud83e\udd49'];
    var nc0    = (learner.nextCourses && learner.nextCourses[0]) ? learner.nextCourses[0].name : '';
    return medals[i] + ' ' + (t.name||'\u2014') + ' | Grade: ' + (t.auditGrade||'\u2014')
      + (nc0 ? ' | ' + nc0 + ': ' + (t.courseReady||'\u2014') : '');
  }).join('\n') || '_No pre-matched teachers found_';

  var clsAlert   = learner.learnerHealthCritical
    ? '\n\u26a0\ufe0f *Critical learner* \u2014 CLS approval mandatory' : '';
  var cancelNote = learner.recentCancellations > 0
    ? '\n\u26a0\ufe0f ' + learner.recentCancellations + ' cancellations (30d): ' + (learner.cancelReasons || []).join(', ') : '';

  return [
    { type: 'header', text: { type: 'plain_text', text: emoji + ' CCTC Review: ' + learner.learnerName, emoji: true } },
    {
      type: 'section',
      fields: [
        { type: 'mrkdwn', text: '*Learner:*\n' + learner.learnerName + '\n_' + learner.jlid + '_' },
        { type: 'mrkdwn', text: '*Teacher:*\n' + (learner.teacher||'\u2014') + '\n_' + courseOrdinal + '_' }
      ]
    },
    {
      type: 'section',
      fields: [
        { type: 'mrkdwn', text: '*Current Course:*\n' + (learner.currentCourse||'\u2014') },
        { type: 'mrkdwn', text: '*Classes Remaining:*\n' + (learner.classesLeft != null ? learner.classesLeft : '\u2014') + ' \u00b7 ' + (learner.classFrequency||'\u2014') + '/wk' }
      ]
    },
    {
      type: 'section',
      fields: [
        { type: 'mrkdwn', text: '*Est. Completion:*\n' + dateText + ' (' + weeksText + ')' },
        { type: 'mrkdwn', text: '*Course Journey:*\n' + journeyStr }
      ]
    },
    { type: 'divider' },
    { type: 'section', text: { type: 'mrkdwn', text: '*\ud83d\udcda Next Courses \u2014 Teacher Readiness:*\n' + nextLines + clsAlert + cancelNote } },
    { type: 'section', text: { type: 'mrkdwn', text: '*\ud83d\udc68\u200d\ud83c\udfeb Pre-matched Teachers:*\n' + teacherLines } },
    { type: 'divider' },
    { type: 'section', text: { type: 'mrkdwn', text: '*Should we trigger the CCTC migration?*' } },
    {
      type: 'actions',
      block_id: 'cls_migration_action',
      elements: [
        {
          type: 'button',
          text: { type: 'plain_text', text: '\u2705  Yes, Trigger Migration', emoji: true },
          style: 'primary',
          action_id: 'cls_approve_migration',
          value: learner.jlid,
          confirm: {
            title  : { type: 'plain_text', text: 'Confirm Migration' },
            text   : { type: 'mrkdwn', text: 'Create a HubSpot CCTC ticket for *' + learner.learnerName + '*?' },
            confirm: { type: 'plain_text', text: 'Yes, Create Ticket' },
            deny   : { type: 'plain_text', text: 'Cancel' }
          }
        },
        {
          type: 'button',
          text: { type: 'plain_text', text: '\u274c  No, Decline', emoji: true },
          style: 'danger',
          action_id: 'cls_decline_migration',
          value: learner.jlid
        }
      ]
    }
  ];
}

// ── Update Slack message via response_url ─────────────────────────────
function _slackUpdateMessage(responseUrl, blocks, fallbackText) {
  if (!responseUrl) return;
  try {
    UrlFetchApp.fetch(responseUrl, {
      method            : 'post',
      headers           : { 'Content-Type': 'application/json' },
      payload           : JSON.stringify({ replace_original: true, text: fallbackText || '', blocks: blocks }),
      muteHttpExceptions: true
    });
  } catch(e) { Logger.log('[CLS] _slackUpdateMessage error: ' + e.message); }
}

// ── Open decline reason modal in Slack ────────────────────────────────
function openSlackDeclineModal(triggerId, jlid, learnerName) {
  var token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
  if (!token || !triggerId) return;
  var modal = {
    type            : 'modal',
    callback_id     : 'cls_decline_reason',
    private_metadata: jlid,
    title           : { type: 'plain_text', text: 'Decline Migration' },
    submit          : { type: 'plain_text', text: 'Submit Decline' },
    close           : { type: 'plain_text', text: 'Cancel' },
    blocks: [
      { type: 'section', text: { type: 'mrkdwn', text: 'Declining CCTC migration for *' + (learnerName || jlid) + '*.\nPlease provide a reason.' } },
      {
        type    : 'input',
        block_id: 'decline_reason_block',
        label   : { type: 'plain_text', text: 'Reason for declining' },
        element : {
          type       : 'plain_text_input',
          action_id  : 'decline_reason_input',
          multiline  : true,
          placeholder: { type: 'plain_text', text: 'e.g. Teacher fast-pacing upskill, learner prefers same teacher, renewal in 2 weeks\u2026' }
        }
      }
    ]
  };
  try {
    UrlFetchApp.fetch('https://slack.com/api/views.open', {
      method            : 'post',
      headers           : { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json; charset=utf-8' },
      payload           : JSON.stringify({ trigger_id: triggerId, view: modal }),
      muteHttpExceptions: true
    });
  } catch(e) { Logger.log('[CLS] openSlackDeclineModal error: ' + e.message); }
}

// ─────────────────────────────────────────────────────────────────────
// sendCLSMigrationRequest(jlid, triggeredBy)
// ─────────────────────────────────────────────────────────────────────
function sendCLSMigrationRequest(jlid, triggeredBy) {
  try {
    Logger.log('[CLS] sendCLSMigrationRequest: ' + jlid);
    if (!jlid) return { success: false, message: 'JLID required.' };

    var token   = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
    var channel = PropertiesService.getScriptProperties().getProperty('SLACK_CLS_CHANNEL') || '#cls-migrations';
    if (!token) return { success: false, message: 'SLACK_BOT_TOKEN not set in Script Properties.' };

    var prog = getLearnerProgressions();
    if (!prog.success) return { success: false, message: 'Could not load progressions: ' + prog.message };
    var learner = null;
    for (var i = 0; i < prog.learners.length; i++) {
      if (prog.learners[i].jlid === jlid) { learner = prog.learners[i]; break; }
    }
    if (!learner) return { success: false, message: 'Learner ' + jlid + ' not found.' };

    // Block duplicate pending requests
    var existing = _clsQueueGetEntry(jlid);
    if (existing && String(existing.data[CLS_Q.STATUS]) === 'Pending') {
      return {
        success       : false,
        alreadyPending: true,
        message       : 'CLS approval request already pending for ' + learner.learnerName + '. Check Slack.',
        slackTs       : String(existing.data[CLS_Q.SLACK_TS] || '')
      };
    }

    // Pre-match teachers — pass learner's class slots so we get real slot availability back
    var matchedTeachers = [];
    try {
      var nc0       = (learner.nextCourses && learner.nextCourses[0]) ? learner.nextCourses[0].name : learner.currentCourse;
      var lSlots    = (learner.classSessions || []).filter(function(s){ return s.day && s.time; });
      var smtRes = searchMatchingTeachers({
        currentCourse: nc0, learnerAge: '', techTraits: '', mathTraits: '',
        requestedSlots: lSlots, requestedSlot: lSlots.length > 0 ? (lSlots[0].day + ' ' + lSlots[0].time) : '',
        futureCourse1: (learner.nextCourses && learner.nextCourses[1]) ? learner.nextCourses[1].name : ''
      });
      if (smtRes.success && smtRes.results) {
        matchedTeachers = smtRes.results.slice(0, 5).map(function(t) {
          return {
            name          : t.teacherName,
            auditGrade    : t.auditGrade    || '\u2014',
            courseReady   : t.currentCourseProgress || '\u2014',
            slotMatch     : t.slotMatch     || '',
            alternateSlots: t.alternateSlots || ''
          };
        });
      }
    } catch(smtErr) { Logger.log('[CLS] Teacher search non-fatal: ' + smtErr.message); }

    // Cache matched teachers so triggerSmartMigration can reuse them without re-searching
    if (matchedTeachers.length > 0) {
      try {
        CacheService.getScriptCache().put('CLS_TEACHERS_' + jlid, JSON.stringify(matchedTeachers), 86400);
      } catch(ce) {}
    }

    var blocks     = _buildCLSSlackBlocks(learner, matchedTeachers);
    var nc0name    = (learner.nextCourses && learner.nextCourses[0]) ? learner.nextCourses[0].name : '\u2014';
    var msgPayload = {
      channel     : channel,
      text        : 'CCTC Migration Review: ' + learner.learnerName + ' (' + jlid + ')',
      blocks      : blocks,
      unfurl_links: false
    };

    var resp     = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
      method            : 'post',
      headers           : { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json; charset=utf-8' },
      payload           : JSON.stringify(msgPayload),
      muteHttpExceptions: true
    });
    var respData = JSON.parse(resp.getContentText());
    if (!respData.ok) {
      Logger.log('[CLS] Slack post failed: ' + (respData.error || JSON.stringify(respData)));
      return { success: false, message: 'Slack error: ' + (respData.error || 'unknown') };
    }

    var slackTs = respData.ts      || '';
    var slackCh = respData.channel || channel;

    _clsQueueUpsert(jlid, {
      learner    : learner.learnerName,
      teacher    : learner.teacher,
      curCourse  : learner.currentCourse,
      nextCourse : nc0name,
      status     : 'Pending',
      slackTs    : slackTs,
      slackCh    : slackCh,
      triggeredBy: triggeredBy || ''
    });

    try {
      logAction('CLS Migration Request Sent', jlid, learner.learnerName, learner.teacher, '', learner.currentCourse,
        'Pending CLS Approval', 'Slack notification sent to ' + channel, 'CCTC', triggeredBy || '');
    } catch(ae) {}

    Logger.log('[CLS] Slack message sent ts=' + slackTs + ' ch=' + slackCh);
    return {
      success        : true,
      slackTs        : slackTs,
      slackChannel   : slackCh,
      learnerName    : learner.learnerName,
      message        : 'CLS approval request sent to ' + channel,
      matchedTeachers: matchedTeachers
    };
  } catch(e) {
    Logger.log('[CLS] sendCLSMigrationRequest error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Ticket creation diagnostic — run in GAS editor to test HubSpot ────
function diagTicket() {
  var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  if (!token) { Logger.log('[DIAG] ❌ HUBSPOT_API_KEY not set'); return; }
  Logger.log('[DIAG] token starts: ' + token.substring(0, 12));

  // Step 1: minimal payload — just required fields
  var payload1 = {
    properties: {
      subject          : '[DIAG TEST] Delete me',
      hs_pipeline      : '66161281',
      hs_pipeline_stage: '128913747'
    }
  };
  var r1 = UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/objects/tickets', {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload1),
    muteHttpExceptions: true
  });
  Logger.log('[DIAG] Step 1 (minimal): ' + r1.getResponseCode() + ' — ' + r1.getContentText().substring(0, 300));
  if (r1.getResponseCode() !== 201) { Logger.log('[DIAG] ❌ Even minimal payload fails — check pipeline ID or API key scope'); return; }

  var ticketId = JSON.parse(r1.getContentText()).id;
  Logger.log('[DIAG] ✅ Ticket created: ' + ticketId + ' — deleting now');

  // Step 2: delete test ticket
  UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/objects/tickets/' + ticketId, {
    method: 'delete',
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  // Step 3: test custom fields one at a time
  var customFields = {
    learner_uid            : 'JL_TEST_123',
    reason_of_migration__t_: 'Course Change Teacher Change',
    current_course__t_     : 'Test Course',
    future_course_1        : 'Next Course',
    future_course_2        : '',
    future_course_3        : ''
  };
  var r2 = UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/objects/tickets', {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ properties: Object.assign({ subject: '[DIAG TEST 2] Delete me', hs_pipeline: '66161281', hs_pipeline_stage: '128913747' }, customFields) }),
    muteHttpExceptions: true
  });
  Logger.log('[DIAG] Step 2 (with custom fields): ' + r2.getResponseCode() + ' — ' + r2.getContentText().substring(0, 500));
  if (r2.getResponseCode() === 201) {
    var t2id = JSON.parse(r2.getContentText()).id;
    Logger.log('[DIAG] ✅ Custom fields OK — deleting ticket ' + t2id);
    UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/objects/tickets/' + t2id, { method: 'delete', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
  } else {
    Logger.log('[DIAG] ❌ Custom fields causing 400 — check error above for which field');
  }
}

// ── Slack connectivity diagnostic — run this directly in GAS editor to test ──
function diagSlack() {
  var props   = PropertiesService.getScriptProperties().getProperties();
  var token   = props['SLACK_BOT_TOKEN']   || '';
  var channel = props['SLACK_CLS_CHANNEL'] || '';
  Logger.log('[DIAG] SLACK_BOT_TOKEN set: ' + (token.length > 0) + ' (length=' + token.length + ', starts=' + token.substring(0,10) + ')');
  Logger.log('[DIAG] SLACK_CLS_CHANNEL: "' + channel + '"');
  if (!token) { Logger.log('[DIAG] ❌ No token — set SLACK_BOT_TOKEN in Script Properties'); return; }
  // Test: post a simple message
  var resp = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json; charset=utf-8' },
    payload: JSON.stringify({ channel: channel || '#general', text: '✅ JetLearn Slack test from GAS — if you see this, Slack is connected!' }),
    muteHttpExceptions: true
  });
  var data = JSON.parse(resp.getContentText());
  Logger.log('[DIAG] Slack response: ok=' + data.ok + ' error=' + (data.error||'none') + ' channel=' + (data.channel||''));
  if (!data.ok) {
    Logger.log('[DIAG] ❌ Fix: ' + _slackErrorHint(data.error));
  } else {
    Logger.log('[DIAG] ✅ Message delivered to channel ' + data.channel);
  }
}
function _slackErrorHint(err) {
  var hints = {
    'not_in_channel': 'Bot is not a member — /invite @YourBotName in the Slack channel',
    'channel_not_found': 'Channel ID/name wrong — use the channel ID (C0123ABCDEF), not the name',
    'invalid_auth': 'Token is wrong or expired — regenerate Bot User OAuth Token',
    'missing_scope': 'Bot missing scope — add chat:write and chat:write.public in api.slack.com/apps',
    'account_inactive': 'Slack workspace inactive or bot deactivated'
  };
  return hints[err] || ('Unknown error: ' + err + ' — check api.slack.com/methods/chat.postMessage');
}

// ── handleCLSMigrationApproval — called from doPost on ✅ click ───────
function handleCLSMigrationApproval(jlid, slackUserId, slackUserName, responseUrl) {
  try {
    Logger.log('[CLS] handleCLSMigrationApproval: ' + jlid + ' by ' + slackUserName);
    var approvedBy  = slackUserName || slackUserId || 'CLS Manager';
    var entry       = _clsQueueGetEntry(jlid);
    var learnerName = entry ? String(entry.data[CLS_Q.LEARNER]    || jlid) : jlid;
    var teacher     = entry ? String(entry.data[CLS_Q.TEACHER]    || '')   : '';
    var curCourse   = entry ? String(entry.data[CLS_Q.CUR_COURSE] || '')   : '';

    var ticketResult = triggerSmartMigration(jlid);

    _clsQueueUpsert(jlid, { status: 'Approved', reason: 'Approved by ' + approvedBy });

    try {
      logAction('CLS Migration Approved', jlid, learnerName, teacher, '', curCourse,
        ticketResult.success ? 'Ticket Created' : 'Ticket Failed',
        'Approved via Slack by ' + approvedBy + (ticketResult.ticketId ? '. Ticket #' + ticketResult.ticketId : ''),
        'CCTC', approvedBy);
    } catch(ae) {}

    _slackUpdateMessage(responseUrl, [{
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: '\u2705 *Migration Approved* by @' + approvedBy
          + (ticketResult.success
            ? '\n*HubSpot Ticket:* <https://app.hubspot.com/contacts/7729491/record/0-5/' + ticketResult.ticketId + '|#' + ticketResult.ticketId + '> \u2014 ' + (ticketResult.subject || '')
            : '\n\u26a0\ufe0f Ticket creation failed: ' + (ticketResult.message || 'unknown'))
      }
    }], '\u2705 Approved: ' + learnerName);

    return { success: true, ticketResult: ticketResult };
  } catch(e) {
    Logger.log('[CLS] handleCLSMigrationApproval error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── handleCLSMigrationDecline — called from doPost on modal submit ────
function handleCLSMigrationDecline(jlid, reason, slackUserId, slackUserName, responseUrl) {
  try {
    Logger.log('[CLS] handleCLSMigrationDecline: ' + jlid + ' reason: ' + reason);
    var declinedBy  = slackUserName || slackUserId || 'CLS Manager';
    var entry       = _clsQueueGetEntry(jlid);
    var learnerName = entry ? String(entry.data[CLS_Q.LEARNER]    || jlid) : jlid;
    var teacher     = entry ? String(entry.data[CLS_Q.TEACHER]    || '')   : '';
    var curCourse   = entry ? String(entry.data[CLS_Q.CUR_COURSE] || '')   : '';

    _clsQueueUpsert(jlid, { status: 'Declined', reason: reason || 'No reason provided' });

    try {
      logAction('CLS Migration Declined', jlid, learnerName, teacher, '', curCourse,
        'Declined', 'Declined via Slack by ' + declinedBy + '. Reason: ' + (reason || 'Not provided'),
        'CCTC', declinedBy);
    } catch(ae) {}

    _slackUpdateMessage(responseUrl, [{
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: '\u274c *Migration Declined* by @' + declinedBy
          + '\n*Reason:* ' + (reason || '_No reason provided_')
      }
    }], '\u274c Declined: ' + learnerName);

    return { success: true };
  } catch(e) {
    Logger.log('[CLS] handleCLSMigrationDecline error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── getCLSApprovalQueue — returns queue rows for Course Planner UI ─────
function getCLSApprovalQueue() {
  try {
    var sh   = _getOrCreateCLSQueueSheet();
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: true, queue: [], stats: { pending:0, approved:0, declined:0 } };
    var queue = [];
    for (var i = data.length - 1; i >= 1 && queue.length < 100; i--) {
      var r = data[i];
      if (!r[CLS_Q.JLID]) continue;
      queue.push({
        jlid       : String(r[CLS_Q.JLID]),
        learnerName: String(r[CLS_Q.LEARNER]     || ''),
        teacher    : String(r[CLS_Q.TEACHER]     || ''),
        curCourse  : String(r[CLS_Q.CUR_COURSE]  || ''),
        nextCourse : String(r[CLS_Q.NEXT_COURSE] || ''),
        status     : String(r[CLS_Q.STATUS]      || ''),
        reason     : String(r[CLS_Q.REASON]      || ''),
        slackTs    : String(r[CLS_Q.SLACK_TS]    || ''),
        triggeredAt: r[CLS_Q.TRIGGERED] ? r[CLS_Q.TRIGGERED].toString() : '',
        resolvedAt : r[CLS_Q.RESOLVED]  ? r[CLS_Q.RESOLVED].toString()  : ''
      });
    }
    var pending  = queue.filter(function(q){ return q.status === 'Pending';  }).length;
    var approved = queue.filter(function(q){ return q.status === 'Approved'; }).length;
    var declined = queue.filter(function(q){ return q.status === 'Declined'; }).length;
    return { success: true, queue: queue, stats: { pending: pending, approved: approved, declined: declined } };
  } catch(e) {
    Logger.log('[CLS] getCLSApprovalQueue error: ' + e.message);
    return { success: false, queue: [], message: e.message };
  }
}
