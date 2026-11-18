/**
 * KitTeacherNotifyService.js
 *
 * Keeps the learner's teacher (with the CLS manager and TP manager copied) in the loop about a kit JetLearn sends:
 * ordered -> date moved -> late / needs CLS -> delivered, as ONE email thread per kit. Kits JetLearn orders only
 * ("sent by us"); Parent Will Buy is not covered here.
 *
 * Events: ordered | ordered_again | delayed | late | reminder | check_cls | delivered | reorder | cancelled
 *
 * Thread, last-told date and stage live in extra columns on the Kits sheet (KIT_COL.TEACHER_*).
 * Thresholds can be changed with the Script Property KIT_TEACHER_NOTIFY_CFG (JSON), e.g. {"remindAfterDays":3}.
 */

var KIT_TEACHER_CFG_DEFAULT = { remindAfterDays: 2, escalateAfterDays: 7, neededWithinDays: 7, reminderEveryDays: 3, maxPerRun: 20 };

function _ktCfg() {
  var cfg = {};
  Object.keys(KIT_TEACHER_CFG_DEFAULT).forEach(function(k) { cfg[k] = KIT_TEACHER_CFG_DEFAULT[k]; });
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('KIT_TEACHER_NOTIFY_CFG');
    if (raw) { var o = JSON.parse(raw); Object.keys(o).forEach(function(k) { if (typeof o[k] === 'number') cfg[k] = o[k]; }); }
  } catch (e) {}
  return cfg;
}

// Make sure the Kits sheet has the extra columns (and headers) before anything reads or writes them.
var _ktColsEnsured = false;
function _ktEnsureCols(sheet) {
  if (_ktColsEnsured) return;
  try {
    var need = KIT_COL.TEACHER_REMINDED_AT;
    var have = sheet.getMaxColumns();
    if (have < need) sheet.insertColumnsAfter(have, need - have);
    var heads = [[KIT_COL.TEACHER_THREAD, 'Teacher Thread ID'], [KIT_COL.TEACHER_NOTIFIED_AT, 'Teacher Notified At'], [KIT_COL.TEACHER_TOLD_ETA, 'Teacher Told ETA'],
                 [KIT_COL.TEACHER_STAGE, 'Teacher Stage'], [KIT_COL.TEACHER_REMINDED_AT, 'Teacher Reminded At']];
    heads.forEach(function(h) {
      var c = sheet.getRange(1, h[0]);
      if (!String(c.getValue() || '').trim()) c.setValue(h[1]);
    });
    _ktColsEnsured = true;
  } catch (e) { Logger.log('[KitTeacher] ensure columns failed: ' + e.message); }
}

// ── small helpers ───────────────────────────────────────────────────────────
function _ktToDate(v) {
  if (!v) return null;
  var d = null;
  if (Object.prototype.toString.call(v) === '[object Date]') d = new Date(v.getTime());
  else {
    var s = String(v).trim();
    if (!s) return null;
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) { var p = s.substring(0, 10).split('-'); d = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10)); }
    else d = _parseDMY(s);
  }
  if (!d || isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}
function _ktFmt(d) { return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'EEE d MMM yyyy') : ''; }
function _ktToday() { var t = new Date(); t.setHours(0, 0, 0, 0); return t; }
function _ktDaysBetween(a, b) { return Math.round((b.getTime() - a.getTime()) / 86400000); }
function _ktSameDay(a, b) { return !!a && !!b && a.getTime() === b.getTime(); }
function _ktFirst(name) { return String(name || '').trim().split(/\s+/)[0] || ''; }

// Teacher + CLS + TP for a learner. The teacher is the learner's CURRENT teacher on the deal.
function _ktPeople(jlid) {
  var out = { teacherName: '', teacherEmail: '', clsName: '', clsEmail: '', tpName: '', tpEmail: '', dealId: '', learnerName: '' };
  try {
    var hs = fetchHubspotByJlid(jlid);
    if (hs && hs.success && hs.data) {
      out.teacherName = hs.data.currentTeacher || '';
      out.dealId = hs.data.dealId || '';
      out.learnerName = hs.data.learnerName || '';
      out.clsName = hs.data.clsManagerName || '';
    }
  } catch (e) { Logger.log('[KitTeacher] deal lookup failed: ' + e.message); }
  try {
    var sheet = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA);
    var low = String(out.teacherName).trim().toLowerCase();
    for (var r = 1; low && r < sheet.length; r++) {
      if (String(sheet[r][1] || '').trim().toLowerCase() === low) {
        out.teacherEmail = String(sheet[r][8] || '').trim();
        out.tpName = String(sheet[r][6] || '').trim();
        out.clsName = String(sheet[r][7] || '').trim() || out.clsName;
        out.clsEmail = String(sheet[r][9] || '').trim();
        out.tpEmail = String(sheet[r][10] || '').trim();
        break;
      }
    }
  } catch (e) { Logger.log('[KitTeacher] teacher sheet lookup failed: ' + e.message); }
  if (!out.clsEmail && out.clsName) { try { out.clsEmail = findClsEmailByManagerName(out.clsName) || ''; } catch (e) {} }
  return out;
}

// ── the "brain": facts -> severity ──────────────────────────────────────────
function _ktFacts(row, event, opts, cfg) {
  var today = _ktToday();
  var eta = opts.newEta ? _ktToDate(opts.newEta) : _ktToDate(row[KIT_COL.ETA - 1]);
  var prevEta = opts.prevEta ? _ktToDate(opts.prevEta) : _ktToDate(row[KIT_COL.TEACHER_TOLD_ETA - 1]);
  var needed = _ktToDate(row[KIT_COL.MODULE_START_DATE - 1]);
  var daysLate = (eta && today > eta) ? _ktDaysBetween(eta, today) : 0;
  var daysToNeed = needed ? _ktDaysBetween(today, needed) : null;
  var atRisk = !!(needed && eta && eta > needed);
  var neededSoon = (daysToNeed !== null && daysToNeed <= cfg.neededWithinDays && daysToNeed >= -30);
  var severity = 'ok';
  if (daysLate >= cfg.remindAfterDays || atRisk) severity = 'late';
  if (daysLate >= cfg.escalateAfterDays || (neededSoon && (daysLate > 0 || atRisk))) severity = 'blocked';
  return {
    learner: String(row[KIT_COL.LEARNER_NAME - 1] || '').trim(), jlid: String(row[KIT_COL.JLID - 1] || '').trim(),
    kit: String(row[KIT_COL.KIT - 1] || '').trim(), eta: eta, prevEta: prevEta, needed: needed, daysLate: daysLate, daysToNeed: daysToNeed,
    atRisk: atRisk, neededSoon: neededSoon, severity: severity,
    delivered: _ktToDate(row[KIT_COL.DELIVERY_DATE - 1]),
    reminders: (String(row[KIT_COL.FOLLOWUP_SENT - 1]).toUpperCase() === 'TRUE' ? 1 : 0) + (String(row[KIT_COL.FOLLOWUP2_SENT - 1]).toUpperCase() === 'TRUE' ? 1 : 0)
  };
}

function _ktHash(s) { var h = 5381; s = String(s); for (var i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0; return String(h >>> 0); }

// Free-text reasons from the team are tidied into one plain sentence by Gemini. It may not add dates or promises;
// anything odd (or Gemini unavailable) falls back to the person's own words.
function _ktPlainReason(raw, facts) {
  raw = String(raw || '').trim();
  if (!raw) return '';
  var key = 'KT_RSN_' + _ktHash(raw);
  try { var hit = CacheService.getScriptCache().get(key); if (hit) return hit; } catch (e) {}
  var out = raw;
  try {
    var prompt = 'Rewrite this delivery-delay reason as ONE short, neutral, plain sentence (max 18 words) for a teacher. '
      + 'Do not add dates, numbers, promises or blame. Keep the meaning. Reason: "' + raw.replace(/"/g, "'") + '". Reply with JSON only: {"reason":"..."}';
    var j = JSON.parse(_callGemini([{ role: 'user', parts: [{ text: prompt }] }], null, 256));
    var r = String((j && j.reason) || '').trim();
    if (r && r.length <= 160 && !/\d/.test(r)) out = r;
  } catch (e) { Logger.log('[KitTeacher] reason wording skipped: ' + e.message); }
  try { CacheService.getScriptCache().put(key, out, 21600); } catch (e) {}
  return out;
}

// Reason shown to the teacher for the standard cases
function _ktReasonFor(event, opts) {
  var raw = String(opts.reason || '').trim();
  if (raw && raw !== 'parent_silent' && raw !== 'late' && raw !== 'needed_soon') return raw;
  return '';
}

// ── compose ─────────────────────────────────────────────────────────────────
function _ktCompose(event, f, people, opts, cfg) {
  var cls = people.clsName || 'your CLS manager';
  var kitOf = f.learner + '\'s ' + f.kit + ' kit';
  var rows = [];
  var spec = { tag: '', tone: 'info', paras: [], rows: rows, actionTitle: 'What to do', actionText: '', ctaLabel: '', ctaHref: '', headerLabel: 'JetLearn · Kit update' };
  var reason = _ktPlainReason(_ktReasonFor(event, opts), f);
  var etaTxt = _ktFmt(f.eta);
  var neededRow = f.needed ? { l: 'Kit needed from', v: _ktFmt(f.needed) + (f.atRisk ? ' (the kit may arrive after this)' : '') } : null;
  var mailto = people.clsEmail ? 'mailto:' + people.clsEmail + '?subject=' + encodeURIComponent('Kit for ' + f.learner + ' (' + f.jlid + ')') : '';
  var escalateOn = f.eta ? new Date(f.eta.getTime() + cfg.escalateAfterDays * 86400000) : null;

  if (event === 'ordered' || event === 'ordered_again') {
    spec.tag = event === 'ordered' ? 'Kit ordered' : 'New order placed'; spec.tone = event === 'ordered' ? 'good' : 'info';
    spec.paras.push(event === 'ordered'
      ? 'A ' + f.kit + ' kit has been ordered for ' + f.learner + '. Delivery is handled by the courier, not by JetLearn, so the date can move.'
      : 'A new order has been placed for ' + kitOf + ' after the earlier one was cancelled. Delivery is handled by the courier, so the date can move.');
    rows.push({ l: 'Kit', v: f.kit }); if (etaTxt) rows.push({ l: 'Expected by', v: etaTxt }); if (neededRow) rows.push(neededRow);
    spec.actionText = 'Keep teaching as planned. Please hold any kit-based activity until we confirm delivery in this thread.';
  } else if (event === 'delayed') {
    spec.tag = 'Running late'; spec.tone = f.severity === 'blocked' ? 'bad' : 'warn';
    spec.paras.push(kitOf + ' is running late. This is outside our control and we are following it up.');
    rows.push({ l: 'Expected by', v: etaTxt, old: f.prevEta && !_ktSameDay(f.prevEta, f.eta) ? _ktFmt(f.prevEta) : '' });
    if (reason) rows.push({ l: 'Reason', v: reason }); if (neededRow) rows.push(neededRow);
    spec.actionText = 'No change to your lessons for now.' + (escalateOn ? ' If the kit is still not here by ' + _ktFmt(escalateOn) + ', please connect with ' + cls + '.' : ' If it runs much later, please connect with ' + cls + '.');
    if (f.severity === 'blocked') { spec.actionText = 'This kit is needed soon. Please connect with ' + cls + ' so lessons that need it can be planned.'; spec.ctaLabel = 'Email ' + cls; spec.ctaHref = mailto; }
  } else if (event === 'late' || event === 'reminder') {
    spec.tag = 'Still not delivered'; spec.tone = 'warn';
    spec.paras.push(kitOf + ' was due on ' + etaTxt + ' and has not been confirmed as delivered' + (f.daysLate ? ' (' + f.daysLate + ' day' + (f.daysLate === 1 ? '' : 's') + ' late)' : '') + '. We are checking with logistics.');
    rows.push({ l: 'Was due', v: etaTxt }); if (neededRow) rows.push(neededRow);
    spec.actionText = 'No change to your lessons for now. If you need the kit for a lesson, please connect with ' + cls + '.';
  } else if (event === 'check_cls') {
    spec.headerLabel = 'JetLearn · Action needed'; spec.tag = 'Check with CLS'; spec.tone = 'bad';
    var why = String(opts.reason || '');
    if (why === 'parent_silent') {
      spec.paras.push('We have not been able to arrange ' + kitOf + '. The parent has not replied to ' + (f.reminders || 2) + ' reminders. Lessons that need the kit cannot go ahead without it.');
      rows.push({ l: 'Kit', v: f.kit }); rows.push({ l: 'Parent reminders', v: (f.reminders || 2) + ' sent, no reply' });
    } else if (why === 'late' || why === 'needed_soon') {
      spec.paras.push(kitOf + (f.daysLate ? ' is ' + f.daysLate + ' day' + (f.daysLate === 1 ? '' : 's') + ' late' : ' is needed soon and is not here yet') + '. Lessons that need the kit may be affected.');
      rows.push({ l: 'Was due', v: etaTxt });
    } else {
      spec.paras.push('There is a problem with ' + kitOf + (reason ? ': ' + reason : '.') + ' Lessons that need the kit cannot go ahead until it is sorted.');
      rows.push({ l: 'Kit', v: f.kit });
    }
    if (neededRow) rows.push(neededRow);
    spec.actionText = 'Please check with ' + cls + '. The options are to wait for the kit, move to a course that needs no kit, or pause. Please keep the current module going until CLS decides.';
    spec.ctaLabel = 'Email ' + cls; spec.ctaHref = mailto;
  } else if (event === 'delivered') {
    spec.tag = 'Delivered'; spec.tone = 'good';
    spec.paras.push(kitOf + ' arrived' + (f.delivered ? ' on ' + _ktFmt(f.delivered) : '') + '. You can use it in lessons from now on. If anything is missing or not working, tell ' + cls + ' and we will sort it out.');
  } else if (event === 'reorder') {
    spec.tag = 'Reordering'; spec.tone = 'warn';
    spec.paras.push('The earlier order for ' + kitOf + ' was cancelled' + (reason ? ' (' + reason + ')' : '') + '. A new kit is being ordered.');
    spec.actionText = 'Keep teaching the current module. We will confirm the new date in this thread.';
  } else if (event === 'cancelled') {
    spec.headerLabel = 'JetLearn · Action needed'; spec.tag = 'Order cancelled'; spec.tone = 'bad';
    spec.paras.push('The kit order for ' + f.learner + ' was cancelled' + (reason ? ' (' + reason + ')' : '') + '. No kit is coming right now.');
    spec.actionText = 'Please check with ' + cls + ' about how to continue (a course that needs no kit, or a pause).';
    spec.ctaLabel = 'Email ' + cls; spec.ctaHref = mailto;
  }
  spec.footer = event === 'delivered'
    ? 'This closes the kit thread.'
    : 'We will reply in this thread if anything changes. Questions about the kit: ' + cls + (people.tpName ? ' or ' + people.tpName + ' (TP manager)' : '') + '.';
  spec.learnerName = f.learner; spec.jlid = f.jlid; spec.teacherName = _ktFirst(people.teacherName);
  return spec;
}

function _ktSubject(f) { return 'Kit for ' + f.learner + ' (' + f.jlid + '), ' + f.kit; }

function _ktPlain(spec) {
  return [spec.tag, '', 'Hi ' + (spec.teacherName || 'there') + ','].concat(spec.paras).concat(spec.rows.map(function(r) { return r.l + ': ' + (r.old ? r.old + ' -> ' : '') + r.v; }))
    .concat(spec.actionText ? [spec.actionTitle + ': ' + spec.actionText] : []).concat(['', spec.footer]).join('\n');
}

// ── send: first email starts the thread; later ones reply in it ─────────────
function _ktFindThread(subject, toEmail) {
  var q = 'in:sent to:' + toEmail + ' subject:"' + subject.replace(/["()]/g, ' ').replace(/\s+/g, ' ').trim() + '" newer_than:2d';
  for (var attempt = 0; attempt < 2; attempt++) {
    try { var ts = GmailApp.search(q, 0, 1); if (ts && ts.length) return ts[0].getId(); } catch (e) { Logger.log('[KitTeacher] thread search failed: ' + e.message); }
    if (attempt === 0) Utilities.sleep(2500);
  }
  return '';
}

function _ktSend(threadId, f, people, spec) {
  var subject = _ktSubject(f);
  var tpl = HtmlService.createTemplateFromFile('KitTeacherTemplate');
  tpl.data = spec;
  var html = tpl.evaluate().getContent();
  var plain = _ktPlain(spec);
  var cc = [people.clsEmail, people.tpEmail].filter(function(e, i, a) { return e && isValidEmail(e) && a.indexOf(e) === i; }).join(',');

  if (threadId) {
    try {
      var th = GmailApp.getThreadById(threadId);
      if (th) {
        try { th.replyAll(plain, { htmlBody: html, name: CONFIG.EMAIL.FROM_NAME, from: CONFIG.EMAIL.FROM }); }
        catch (fe) { th.replyAll(plain, { htmlBody: html, name: CONFIG.EMAIL.FROM_NAME }); }
        return { threadId: threadId, replied: true };
      }
    } catch (te) { Logger.log('[KitTeacher] reply in thread failed, sending a new message: ' + te.message); }
    subject = 'Re: ' + subject;
  }
  sendTrackedEmail({ to: people.teacherEmail, cc: cc, subject: subject, htmlBody: html, jlid: f.jlid });
  var found = _ktFindThread(subject, people.teacherEmail);
  return { threadId: found || threadId || '', replied: false };
}

/**
 * Tell the teacher (CLS + TP copied) about a kit event. Safe to call repeatedly: nothing is sent twice for the same state.
 * opts: { reason, newEta, prevEta }  (dates as YYYY-MM-DD or DD/MM/YYYY)
 */
function _ktIsSalesPitch(reason) { return /sales\s*pitch/i.test(String(reason || '')); }

function kitTeacherNotify(rowIndex, event, opts) {
  opts = opts || {};
  try {
    var sheet = _getKitSheet();
    _ktEnsureCols(sheet);
    var cfg = _ktCfg();
    var row = sheet.getRange(rowIndex, 1, 1, KIT_COL.TEACHER_REMINDED_AT).getValues()[0];
    var jlid = String(row[KIT_COL.JLID - 1] || '').trim();
    if (!jlid) return { success: false, skipped: true, message: 'No learner ID on the row.' };
    // A kit sent as a sales pitch is not for a learner the teacher is teaching, so the teacher is never told about it.
    if (_ktIsSalesPitch(row[KIT_COL.REASON - 1])) return { success: true, skipped: true, salesPitch: true, message: 'Teacher not informed: this kit is a sales pitch.' };

    var thread = String(row[KIT_COL.TEACHER_THREAD - 1] || '').trim();
    var stage = String(row[KIT_COL.TEACHER_STAGE - 1] || '').trim();
    var toldEta = _ktToDate(row[KIT_COL.TEACHER_TOLD_ETA - 1]);
    var remindedAt = _ktToDate(row[KIT_COL.TEACHER_REMINDED_AT - 1]);
    var today = _ktToday();
    var deliveredAlready = String(row[KIT_COL.DELIVERY_DATE - 1] || '').trim() !== '';

    if (deliveredAlready && event !== 'delivered') return { success: true, skipped: true, message: 'Kit already delivered.' };
    if (event === 'ordered' && thread) {
      if (stage === 'reorder' || stage === 'cancelled') event = 'ordered_again'; else return { success: true, skipped: true, message: 'Teacher already told about this order.' };
    }
    if (event === 'delivered' && stage === 'delivered') return { success: true, skipped: true, message: 'Teacher already told it was delivered.' };
    if ((event === 'reminder' || event === 'late') && !thread) return { success: true, skipped: true, message: 'No thread yet.' };
    if ((event === 'reminder' || event === 'late') && remindedAt && _ktDaysBetween(remindedAt, today) < cfg.reminderEveryDays) return { success: true, skipped: true, message: 'Reminded recently.' };
    if (event === 'check_cls' && stage === 'check_cls' && remindedAt && _ktDaysBetween(remindedAt, today) < 7) return { success: true, skipped: true, message: 'Already asked to check with CLS this week.' };

    var f = _ktFacts(row, event, opts, cfg);
    if (event === 'delayed') {
      if (!f.eta) return { success: true, skipped: true, message: 'No expected date to tell the teacher.' };
      if (toldEta && _ktSameDay(toldEta, f.eta)) return { success: true, skipped: true, message: 'Teacher already has this date.' };
    }
    if (!f.kit) return { success: false, skipped: true, message: 'No kit name on the row.' };

    var people = _ktPeople(jlid);
    if (!people.teacherEmail || !isValidEmail(people.teacherEmail)) {
      return { success: false, message: 'No teacher email found' + (people.teacherName ? ' for ' + people.teacherName : ' (no current teacher on the deal)') + '.' };
    }
    f.learner = f.learner || people.learnerName;

    var spec = _ktCompose(event, f, people, opts, cfg);
    var sent = _ktSend(thread, f, people, spec);

    var nowStage = event === 'ordered' || event === 'ordered_again' ? 'ordered' : (event === 'late' ? 'reminder' : event);
    sheet.getRange(rowIndex, KIT_COL.TEACHER_THREAD).setValue(sent.threadId || thread || '');
    sheet.getRange(rowIndex, KIT_COL.TEACHER_NOTIFIED_AT).setValue(new Date());
    if (f.eta && (event === 'ordered' || event === 'ordered_again' || event === 'delayed')) sheet.getRange(rowIndex, KIT_COL.TEACHER_TOLD_ETA).setValue(f.eta);
    if (event === 'reorder' || event === 'cancelled') sheet.getRange(rowIndex, KIT_COL.TEACHER_TOLD_ETA).clearContent();
    sheet.getRange(rowIndex, KIT_COL.TEACHER_STAGE).setValue(nowStage);
    if (event === 'reminder' || event === 'late' || event === 'check_cls') sheet.getRange(rowIndex, KIT_COL.TEACHER_REMINDED_AT).setValue(new Date());

    if (people.dealId) {
      try { _addNoteToDeal(people.dealId, '[Kit] Teacher informed (' + spec.tag + ') — ' + (people.teacherName || 'teacher') + (people.clsName ? ', CLS ' + people.clsName + ' copied' : '') + '. ' + spec.paras[0]); } catch (ne) {}
    }
    return { success: true, event: event, teacher: people.teacherName, replied: sent.replied, message: 'Teacher ' + (people.teacherName || '') + ' informed (' + spec.tag + ')' + (sent.replied ? ' in the same thread' : '') + '.' };
  } catch (e) {
    Logger.log('[KitTeacher] notify failed (' + event + ' row ' + rowIndex + '): ' + e.message);
    return { success: false, message: 'Teacher email failed: ' + e.message };
  }
}

/**
 * Daily: look at kits whose teacher thread exists and nudge when they run late. Called from the kit follow-up trigger.
 * Only kits already in a teacher thread are touched, so older orders never trigger a surprise email.
 */
function kitTeacherFollowups() {
  try {
    var sheet = _getKitSheet();
    _ktEnsureCols(sheet);
    var cfg = _ktCfg();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { sent: 0 };
    var raw = sheet.getRange(2, 1, lastRow - 1, KIT_COL.TEACHER_REMINDED_AT).getValues();
    var today = _ktToday(), sent = 0;
    for (var i = 0; i < raw.length && sent < cfg.maxPerRun; i++) {
      var r = raw[i];
      if (!String(r[KIT_COL.TEACHER_THREAD - 1] || '').trim()) continue;
      if (_ktIsSalesPitch(r[KIT_COL.REASON - 1])) continue;   // sales-pitch kits: teacher is never nudged
      if (String(r[KIT_COL.DELIVERY_DATE - 1] || '').trim()) continue;
      if (String(r[KIT_COL.REFUNDED - 1] || '').trim().toUpperCase() === 'TRUE') continue;
      var stg = String(r[KIT_COL.TEACHER_STAGE - 1] || '').trim();
      if (stg === 'cancelled' || stg === 'reorder') continue;
      var f = _ktFacts(r, 'reminder', {}, cfg);
      if (!f.eta) continue;
      var event = null, opts = {};
      if (f.severity === 'blocked') { event = 'check_cls'; opts.reason = f.neededSoon && f.daysLate <= 0 ? 'needed_soon' : 'late'; }
      else if (f.daysLate >= cfg.remindAfterDays) event = 'reminder';
      if (!event) continue;
      var res = kitTeacherNotify(i + 2, event, opts);
      if (res && res.success && !res.skipped) sent++;
    }
    Logger.log('[KitTeacher] follow-ups sent: ' + sent);
    return { sent: sent };
  } catch (e) {
    Logger.log('[KitTeacher] follow-ups failed: ' + e.message);
    return { sent: 0, error: e.message };
  }
}

// What each Message Parent preset means for the teacher (used by the Message Parent window). '' = nothing for the teacher.
function _ktEventForPreset(presetId) {
  var map = { delayed: 'delayed', held_customs: 'delayed', delivery_failed: 'delayed', not_received_followup: 'late', late_checking: 'late', address_needed: 'check_cls', cannot_ship: 'check_cls' };
  return map[presetId] || '';
}

// Message Parent window -> "also tell the teacher" even when the parent message itself is skipped
function kitTeacherUpdateFromUi(p) {
  try {
    p = p || {};
    var ev = _ktEventForPreset(String(p.presetId || ''));
    if (!ev) return { success: false, message: 'Nothing to tell the teacher for this message.' };
    return kitTeacherNotify(parseInt(p.rowIndex, 10), ev, { reason: String(p.teacherReason || '').trim(), newEta: p.date || '', prevEta: p.prevDate || '' });
  } catch (e) { return { success: false, message: e.message }; }
}
