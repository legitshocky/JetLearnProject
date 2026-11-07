/**
 * KitLogisticsMessages.js
 *
 * Parent notifications about a kit's logistics (delivered, delayed, cancelled, refunded, reordered ...).
 * One preset = one WhatsApp message (WATI template `outbound_msg`) + one branded email, sent together.
 *
 * WATI template `outbound_msg` variables:  parent, text, text1
 *   text  = the main sentence (what happened)
 *   text1 = the follow-up sentence (what happens next / what we need)
 *
 * Placeholders usable in preset copy and in edited text:  {learner}  {kit}  {date}  {parent}
 */

var KIT_LOGISTICS_WATI_TEMPLATE = 'outbound_msg';

// tone drives the email badge colour: good | info | warn | bad
var KIT_LOGISTICS_PRESETS = [
  { id: 'delivered', group: 'Delivery', tone: 'good', label: 'Delivered (confirmed by logistics)',
    needsDate: true, dateLabel: 'Delivered on', tag: 'Delivered', rail: 'Delivery\nconfirmed',
    text:  "we checked with our logistics team and it has been noted that {learner}'s {kit} kit was delivered on {date}.",
    text1: "If you have not received it, just message us on our customer chat number and we will check with logistics straight away." },

  { id: 'dispatched', group: 'Delivery', tone: 'info', label: 'Dispatched / on its way',
    needsDate: true, dateLabel: 'Expected by', tag: 'On its way', rail: 'Your kit is\non its way',
    text:  "good news, {learner}'s {kit} kit has been dispatched and is on its way to you.",
    text1: "It is expected to reach you by {date}. We will let you know if anything changes." },

  { id: 'out_for_delivery', group: 'Delivery', tone: 'info', label: 'Out for delivery today',
    needsDate: false, tag: 'Out for delivery', rail: 'Out for\ndelivery',
    text:  "{learner}'s {kit} kit is out for delivery and should reach you today.",
    text1: "Please keep your phone handy for the courier, and tell us once it arrives." },

  { id: 'delayed', group: 'Delay', tone: 'warn', label: 'Delayed - new expected date',
    needsDate: true, dateLabel: 'New expected date', needsPrevDate: true, prevDateLabel: 'Originally expected', tag: 'Delayed', rail: 'Delivery\ndelayed',
    text:  "we checked with our logistics team and {learner}'s {kit} kit is running late with the courier.",
    text1: "It was originally expected on {prevdate} and is now expected on {date}. Sorry for the wait, we are following this up closely." },

  { id: 'late_checking', group: 'Delay', tone: 'warn', label: 'Due today but running late - checking',
    needsDate: false, needsPrevDate: true, prevDateLabel: 'Was due on', tag: 'Running late', rail: 'Running late,\nwe are checking',
    text:  "{learner}'s {kit} kit was due to reach you on {prevdate} but has not arrived yet. We have asked our logistics team to check where it is.",
    text1: "We will update you as soon as we hear back, usually within 1 to 2 working days. Sorry for the wait." },

  { id: 'held_customs', group: 'Delay', tone: 'warn', label: 'Held at customs / clearance',
    needsDate: false, tag: 'Delayed', rail: 'Held for\nclearance',
    text:  "{learner}'s {kit} kit is currently held for customs clearance, which is causing a delay.",
    text1: "No action is needed from you. We will update you as soon as it is released." },

  { id: 'not_received_followup', group: 'Delay', tone: 'info', label: 'You said not received - we are checking',
    needsDate: false, tag: 'We are on it', rail: 'Checking with\nlogistics',
    text:  "thank you for letting us know that {learner}'s {kit} kit has not reached you yet. We have raised this with our logistics team.",
    text1: "We will get back to you within 2 working days with a clear update." },

  { id: 'delivery_failed', group: 'Problem', tone: 'warn', label: 'Delivery attempt failed / returned',
    needsDate: false, tag: 'Action needed', rail: 'Delivery\nattempt failed',
    text:  "the courier could not deliver {learner}'s {kit} kit and it is on its way back to us.",
    text1: "Please message us on our customer chat number to confirm your address and a good time to receive it, and we will send it again." },

  { id: 'lost', group: 'Problem', tone: 'bad', label: 'Lost in transit - replacement coming',
    needsDate: false, tag: 'Replacement', rail: 'Kit lost,\nreplacement on it',
    text:  "our logistics team has confirmed that {learner}'s {kit} kit was lost in transit. We are very sorry about this.",
    text1: "We are arranging a replacement at no cost to you and will share the tracking details shortly." },

  { id: 'cancelled_logistics_reorder', group: 'Cancelled', tone: 'warn', label: 'Order cancelled (logistics issue) - reordering new kit',
    needsDate: false, tag: 'Reordering', rail: 'Order cancelled,\nreordering',
    text:  "the order for {learner}'s {kit} kit was cancelled due to a logistics issue with the carrier. We are very sorry about this.",
    text1: "We are placing a fresh order for a new kit and will share the new expected delivery date and tracking details as soon as they are confirmed." },

  { id: 'cancelled_logistics_refund', group: 'Cancelled', tone: 'bad', label: 'Order cancelled (logistics issue) - refunded',
    needsDate: false, tag: 'Refunded', rail: 'Order cancelled,\nrefunded',
    text:  "the order for {learner}'s {kit} kit was cancelled due to a logistics issue with the carrier and has been refunded. We are very sorry about this.",
    text1: "Your learning manager will get in touch with the best way forward." },

  { id: 'cancelled_delay_reorder', group: 'Cancelled', tone: 'warn', label: 'Cancelled due to delay - reordered',
    needsDate: true, dateLabel: 'New order expected by', tag: 'Reordered', rail: 'Order cancelled,\nreordered',
    text:  "because of a long delay in delivery we have cancelled the earlier order for {learner}'s {kit} kit and placed a fresh one.",
    text1: "The new order is expected to reach you by {date}. Sorry for the inconvenience." },

  { id: 'cancelled_delay_refund', group: 'Cancelled', tone: 'bad', label: 'Cancelled due to delay - refunded',
    needsDate: false, tag: 'Refunded', rail: 'Order cancelled,\nrefunded',
    text:  "because of a long delay in delivery we have cancelled the order for {learner}'s {kit} kit and the amount has been refunded.",
    text1: "Your learning manager will get in touch with the best way forward. Sorry for the inconvenience." },

  { id: 'refunded', group: 'Cancelled', tone: 'info', label: 'Refund processed',
    needsDate: true, dateLabel: 'Refunded on', tag: 'Refunded', rail: 'Refund\nprocessed',
    text:  "we checked with our logistics team and the order for {learner}'s {kit} kit has been refunded on {date}.",
    text1: "Nothing further is needed from you. Reply here if you have any questions." },

  { id: 'reordered', group: 'Cancelled', tone: 'info', label: 'New order placed (reorder)',
    needsDate: true, dateLabel: 'Expected by', tag: 'Reordered', rail: 'New order\nplaced',
    text:  "we have placed a fresh order for {learner}'s {kit} kit.",
    text1: "It is expected to reach you by {date}. We will share the tracking details as soon as it ships." },

  { id: 'address_needed', group: 'Address', tone: 'warn', label: 'We need your delivery address',
    needsDate: false, tag: 'Action needed', rail: 'Address\nneeded',
    text:  "we need your delivery address to dispatch {learner}'s {kit} kit.",
    text1: "Please message us on our customer chat number with your full address including pin or zip code and a phone number for the courier." },

  { id: 'cannot_ship', group: 'Address', tone: 'bad', label: 'We cannot ship to your location',
    needsDate: false, tag: 'Update', rail: 'Delivery\nupdate',
    text:  "our logistics partners are unable to deliver {learner}'s {kit} kit to your location right now.",
    text1: "Your learning manager will contact you shortly with an alternative so that {learner} does not miss out." },

  { id: 'custom', group: 'Custom', tone: 'info', label: 'Write my own',
    needsDate: false, tag: 'Update', rail: 'Kit\nupdate',
    text:  "here is an update on {learner}'s {kit} kit from our logistics team.",
    text1: "Reply here if you have any questions." }
];

function getKitLogisticsPresets() {
  var presets = KIT_LOGISTICS_PRESETS.map(function(p) { var c = {}; Object.keys(p).forEach(function(k) { c[k] = p[k]; }); c.teacherEvent = (typeof _ktEventForPreset === 'function') ? _ktEventForPreset(p.id) : ''; return c; });
  return { success: true, template: KIT_LOGISTICS_WATI_TEMPLATE, presets: presets };
}

function _klPreset(id) {
  for (var i = 0; i < KIT_LOGISTICS_PRESETS.length; i++) if (KIT_LOGISTICS_PRESETS[i].id === id) return KIT_LOGISTICS_PRESETS[i];
  return null;
}

// "28/10/2026" or "2026-10-28" -> "28 Oct 2026". Anything else is returned as typed.
function _klFormatDate(raw) {
  var s = String(raw || '').trim();
  if (!s) return '';
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/), d, mo, y;
  if (m) { d = +m[1]; mo = +m[2]; y = +m[3]; }
  else if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) { y = +m[1]; mo = +m[2]; d = +m[3]; }
  else return s;
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return s;
  var names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return d + ' ' + names[mo - 1] + ' ' + y;
}

function _klFill(str, vars) {
  return String(str || '').replace(/\{(learner|kit|date|prevdate|parent)\}/g, function(_, k) { return vars[k] || ''; });
}

// WhatsApp template params can't hold newlines/tabs or runs of 4+ spaces (Meta rule).
function _klWatiSafe(s) {
  return String(s || '').replace(/[\r\n\t]+/g, ' ').replace(/ {2,}/g, ' ').trim().substring(0, 900);
}

function _klFirstName(name) {
  var n = String(name || '').trim().split(/\s+/)[0];
  return n || 'there';
}

// Resolve copy for a payload -> { text, text1, subject, vars, preset }
function _klResolve(p) {
  var preset = _klPreset(String(p.presetId || 'custom')) || _klPreset('custom');
  var vars = {
    learner: String(p.learnerName || 'your child').trim() || 'your child',
    kit:     String(p.kitName || 'learning').trim() || 'learning',
    date:    _klFormatDate(p.date),
    prevdate: _klFormatDate(p.prevDate),
    parent:  _klFirstName(p.parentName)
  };
  if (preset.needsDate && !vars.date) vars.date = 'the date shared earlier';
  if (preset.needsPrevDate && !vars.prevdate) vars.prevdate = 'the earlier date';
  var text  = (p.text  !== undefined && String(p.text).trim()  !== '') ? String(p.text)  : _klFill(preset.text,  vars);
  var text1 = (p.text1 !== undefined && String(p.text1).trim() !== '') ? String(p.text1) : _klFill(preset.text1, vars);
  // user-edited copy may still contain placeholders
  text = _klFill(text, vars); text1 = _klFill(text1, vars);
  // template reads "Greetings from JetLearn! {{text}} {{text1}}" - each part must start a sentence
  var cap = function(s) { s = String(s).trim(); return s.charAt(0).toUpperCase() + s.slice(1); };
  text = cap(text); text1 = cap(text1);
  return { preset: preset, vars: vars, text: text.trim(), text1: text1.trim() };
}

function _klEmailData(res, p) {
  return {
    parentName:  res.vars.parent === 'there' ? '' : (p.parentName || res.vars.parent),
    learnerName: res.vars.learner,
    kitName:     res.vars.kit,
    tone:        res.preset.tone,
    tag:         res.preset.tag,
    railA:       String(res.preset.rail).split('\n')[0],
    railB:       String(res.preset.rail).split('\n')[1] || '',
    text:        res.text,
    text1:       res.text1,
    dateLabel:   res.preset.needsDate && res.vars.date ? res.preset.dateLabel : '',
    dateValue:   res.preset.needsDate ? res.vars.date : '',
    prevLabel:   res.preset.needsPrevDate && p.prevDate ? res.preset.prevDateLabel : '',
    prevValue:   res.preset.needsPrevDate && p.prevDate ? res.vars.prevdate : ''
  };
}

function _klEmailSubject(res) {
  var l = res.vars.learner, t = res.preset.tag;
  var map = {
    delivered: "Update: " + l + "'s kit was delivered",
    dispatched: l + "'s kit is on its way",
    out_for_delivery: l + "'s kit is out for delivery",
    late_checking: "Update: we are checking on " + l + "'s kit",
    delayed: "Update: " + l + "'s kit delivery is delayed",
    held_customs: "Update: " + l + "'s kit is held for clearance",
    not_received_followup: "We are checking on " + l + "'s kit",
    delivery_failed: "Action needed: " + l + "'s kit delivery",
    lost: "Update: replacement for " + l + "'s kit",
    cancelled_logistics_reorder: "Update: new order being placed for " + l + "'s kit",
    cancelled_logistics_refund: "Update: " + l + "'s kit order was cancelled",
    cancelled_delay_reorder: "Update: new order placed for " + l + "'s kit",
    cancelled_delay_refund: "Update: " + l + "'s kit order cancelled and refunded",
    refunded: "Refund processed for " + l + "'s kit",
    reordered: "New order placed for " + l + "'s kit",
    address_needed: "Action needed: delivery address for " + l + "'s kit",
    cannot_ship: "Update on " + l + "'s kit delivery"
  };
  return map[res.preset.id] || ("Update on " + l + "'s kit (" + t + ")");
}

function _klEmailHtml(res, p) {
  var tpl = HtmlService.createTemplateFromFile('KitLogisticsTemplate');
  tpl.data = _klEmailData(res, p);
  return tpl.evaluate().getContent();
}

/** Ops chose NOT to inform the parent after a date change: leave a trace on the deal. */
function logKitParentNotInformed(p) {
  try {
    var jlid = String((p && p.jlid) || '').trim().toUpperCase();
    if (!jlid) return { success: false, message: 'JLID missing' };
    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success || !hs.data.dealId) return { success: false, message: 'Deal not found' };
    _addNoteToDeal(hs.data.dealId, '[Kit Logistics] ' + (p.what || ('Delivery date changed' + (p.date ? ' to ' + _klFormatDate(p.date) : ''))) +
      ' - parent was NOT informed (skipped by ' + (String(p.who || '').trim() || 'ops') + ').');
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

/** Preview (no sending): returns the WhatsApp text and the rendered email HTML. */
function previewKitLogisticsMessage(p) {
  try {
    var res = _klResolve(p || {});
    return {
      success: true,
      text: res.text, text1: res.text1,
      subject: _klEmailSubject(res),
      html: _klEmailHtml(res, p || {})
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Send the message to the parent(s) over WhatsApp and email, and log it on the deal.
 * payload: { jlid, rowIndex, presetId, date, text, text1, sendWati, sendEmail, who }
 */
function sendKitLogisticsMessage(p) {
  try {
    p = p || {};
    var jlid = String(p.jlid || '').trim().toUpperCase();
    if (!jlid) return { success: false, message: 'JLID is missing - add it to the row first.' };
    var sendWati = p.sendWati !== false, sendEmail = p.sendEmail !== false;
    if (!sendWati && !sendEmail) return { success: false, message: 'Pick at least one channel (WhatsApp or Email).' };

    // double-click guard
    var cache = CacheService.getScriptCache();
    var dupKey = 'KIT_LOGI_' + jlid + '_' + String(p.presetId || '') + '_' + String(p.date || '');
    if (cache.get(dupKey)) return { success: false, message: 'This exact message was just sent. Wait a minute before sending it again.' };

    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success) return { success: false, message: 'HubSpot lookup failed for ' + jlid };
    var d = hs.data;

    var kitName = String(p.kitName || '').trim();
    if (!kitName && p.rowIndex) {
      try { kitName = String(_getKitSheet().getRange(parseInt(p.rowIndex, 10), KIT_COL.KIT).getValue() || '').trim(); } catch (ke) {}
    }
    p.kitName = kitName;
    p.learnerName = p.learnerName || d.learnerName;
    p.parentName = d.parentName || '';
    var res = _klResolve(p);

    // every contact on the deal, primary first
    var phones = [], emails = [];
    var seenP = {}, seenE = {};
    var addPhone = function(x) { x = _normalisePhone(String(x || '').replace(/\s*\(.*\)$/, '')); if (x && !seenP[x]) { seenP[x] = true; phones.push(x); } };
    addPhone(d.parentContact);
    if (d.dealId) {
      try { var pr = getPhoneNumbersForDeal(d.dealId); if (pr && pr.all) pr.all.forEach(addPhone); } catch (pe) { Logger.log('[KitLogistics] phones: ' + pe.message); }
      try {
        var er = getEmailsForDeal(d.dealId);
        if (er && er.all) er.all.forEach(function(e) { e = String(e || '').trim(); if (e && isValidEmail(e) && !seenE[e.toLowerCase()]) { seenE[e.toLowerCase()] = true; emails.push(e); } });
      } catch (ee) { Logger.log('[KitLogistics] emails: ' + ee.message); }
    }

    var out = { success: false, wati: { attempted: sendWati, sent: 0, failed: 0, errors: [] }, email: { attempted: sendEmail, sent: false, message: '' } };

    if (sendWati) {
      if (!phones.length) { out.wati.errors.push('No phone number on the deal.'); }
      var params = [
        { name: 'parent', value: res.vars.parent },
        { name: 'text',   value: _klWatiSafe(res.text) },
        { name: 'text1',  value: _klWatiSafe(res.text1) }
      ];
      phones.forEach(function(ph) {
        try { sendWatiMessage(ph, KIT_LOGISTICS_WATI_TEMPLATE, params); out.wati.sent++; }
        catch (we) { out.wati.failed++; out.wati.errors.push(ph.slice(-4) + ': ' + we.message); }
      });
    }

    if (sendEmail) {
      if (!emails.length) { out.email.message = 'No valid parent email on the deal.'; }
      else {
        var er2 = sendTrackedEmail({ to: emails.join(','), subject: _klEmailSubject(res), htmlBody: _klEmailHtml(res, p), jlid: jlid });
        out.email.sent = !!er2; out.email.message = 'Sent to ' + emails.length + ' address' + (emails.length > 1 ? 'es' : '');
      }
    }

    out.success = out.wati.sent > 0 || out.email.sent;
    if (out.success) cache.put(dupKey, '1', 60);

    // Also tell the teacher (same email thread, CLS + TP copied) when the person ticked it
    out.teacher = null;
    var tEvent = (typeof _ktEventForPreset === 'function') ? _ktEventForPreset(String(p.presetId || '')) : '';
    if (p.alsoTeacher && tEvent && p.rowIndex && typeof kitTeacherNotify === 'function') {
      try {
        out.teacher = kitTeacherNotify(parseInt(p.rowIndex, 10), tEvent, { reason: String(p.teacherReason || '').trim(), newEta: p.date || '', prevEta: p.prevDate || '' });
      } catch (te) { out.teacher = { success: false, message: te.message }; }
    }

    // log on the deal
    out.note = 'skipped';
    if (out.success && d.dealId) {
      try {
        var who = String(p.who || '').trim();
        _addNoteToDeal(d.dealId, '[Kit Logistics] Parent informed (' + res.preset.label + ') for ' + (kitName || 'kit') +
          (res.vars.date && res.preset.needsDate ? ' - ' + res.preset.dateLabel + ': ' + res.vars.date : '') +
          '. WhatsApp: ' + (out.wati.sent ? out.wati.sent + ' sent' : 'not sent') + '; Email: ' + (out.email.sent ? 'sent' : 'not sent') +
          (who ? '. By ' + who : '') + '.\nMessage: ' + res.text + ' ' + res.text1);
        out.note = 'added';
      } catch (ne) { out.note = 'failed'; Logger.log('[KitLogistics] note: ' + ne.message); }
    }

    var bits = [];
    if (sendWati)  bits.push('WhatsApp ' + (out.wati.sent ? out.wati.sent + ' sent' : 'FAILED' + (out.wati.errors.length ? ' (' + out.wati.errors[0] + ')' : '')));
    if (sendEmail) bits.push('Email ' + (out.email.sent ? 'sent' : 'FAILED (' + out.email.message + ')'));
    if (out.teacher) bits.push('Teacher ' + (out.teacher.success ? (out.teacher.skipped ? 'not emailed (' + (out.teacher.message || 'nothing new') + ')' : 'informed') : 'FAILED (' + (out.teacher.message || '') + ')'));
    out.message = bits.join(' | ');
    return out;
  } catch (e) {
    Logger.log('[KitLogistics] ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}
