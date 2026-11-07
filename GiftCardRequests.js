// ═══════════════════════════════════════════════════════════════════════════
// REQUESTS — everyone we have asked for their Amazon website + account email, and what came back.
// One row per learner while the request is open (Waiting -> Details received -> Issued, or Closed).
// ═══════════════════════════════════════════════════════════════════════════
var GC_REQ_SHEET_NAME = 'Gift Card Requests';
var GC_REQ_HEADERS = ['Requested At', 'JLID', 'Learner', 'Parent', 'Parent Emails', 'Parent Phone', 'Occasion', 'Offer', 'Channels', 'Requested By', 'Status',
                      'Website', 'Amazon Email', 'Reply Source', 'Reply At', 'Last Checked', 'Reminders', 'Issued At'];
var GC_REQ = { AT: 1, JLID: 2, LEARNER: 3, PARENT: 4, EMAILS: 5, PHONE: 6, OCCASION: 7, OFFER: 8, CHANNELS: 9, BY: 10, STATUS: 11,
               WEBSITE: 12, AMAZON_EMAIL: 13, REPLY_SOURCE: 14, REPLY_AT: 15, CHECKED_AT: 16, REMINDERS: 17, ISSUED_AT: 18 };

function _gcReqSheet() {
  var ss = SpreadsheetApp.openById(GC_SHEET_ID);
  var sh = ss.getSheetByName(GC_REQ_SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(GC_REQ_SHEET_NAME); sh.appendRow(GC_REQ_HEADERS); sh.setFrozenRows(1); }
  return sh;
}

function _gcReqIsOpen(status) { return status === 'Waiting' || status === 'Details received'; }

// Called whenever an ask goes out (WhatsApp or email). A second ask to the same learner is a reminder on the same row.
function gcRecordRequest(jlid, channel, occasionId, offerName, who, hsData) {
  try {
    jlid = String(jlid || '').trim().toUpperCase();
    var sh = _gcReqSheet(), last = sh.getLastRow();
    var vals = last > 1 ? sh.getRange(2, 1, last - 1, GC_REQ_HEADERS.length).getValues() : [];
    var occ = _gcOccasion(occasionId, offerName);
    for (var i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][GC_REQ.JLID - 1]).toUpperCase() === jlid && _gcReqIsOpen(String(vals[i][GC_REQ.STATUS - 1]))) {
        var row = i + 2, chans = String(vals[i][GC_REQ.CHANNELS - 1] || '').split(/\s*\+\s*/).filter(Boolean);
        if (chans.indexOf(channel) === -1) chans.push(channel);
        sh.getRange(row, GC_REQ.CHANNELS).setValue(chans.join(' + '));
        sh.getRange(row, GC_REQ.REMINDERS).setValue((parseInt(vals[i][GC_REQ.REMINDERS - 1], 10) || 0) + 1);
        sh.getRange(row, GC_REQ.CHECKED_AT).clearContent();
        sh.getRange(row, GC_REQ.OCCASION).setValue(occ.id); sh.getRange(row, GC_REQ.OFFER).setValue(occ.offer || '');
        return { success: true, rowIndex: row, reminder: true };
      }
    }
    var d = hsData || {};
    sh.appendRow([new Date(), jlid, d.learnerName || '', d.parentName || '', d.parentEmail || '', d.parentContact || '', occ.id, occ.offer || '', channel, who || '', 'Waiting',
                  '', '', '', '', '', 0, '']);
    return { success: true, rowIndex: sh.getLastRow(), reminder: false };
  } catch (e) {
    Logger.log('[GiftCard] gcRecordRequest failed: ' + e.message);
    return { success: false, message: e.message };
  }
}

function _gcReqRowToObj(v, rowIndex) {
  var occ = _gcOccasion(v[GC_REQ.OCCASION - 1], v[GC_REQ.OFFER - 1]);
  var at = (Object.prototype.toString.call(v[GC_REQ.AT - 1]) === '[object Date]') ? v[GC_REQ.AT - 1] : null;
  return {
    rowIndex: rowIndex, requestedAt: at ? _gcFormatDateTime(at) : String(v[GC_REQ.AT - 1] || ''),
    waitingDays: at ? Math.floor((new Date().getTime() - at.getTime()) / 86400000) : 0,
    jlid: String(v[GC_REQ.JLID - 1] || ''), learner: String(v[GC_REQ.LEARNER - 1] || ''), parent: String(v[GC_REQ.PARENT - 1] || ''),
    emails: String(v[GC_REQ.EMAILS - 1] || ''), phone: String(v[GC_REQ.PHONE - 1] || ''),
    occasionId: occ.id, occasion: occ.label, offer: occ.offer || '', channels: String(v[GC_REQ.CHANNELS - 1] || ''), by: String(v[GC_REQ.BY - 1] || ''),
    status: String(v[GC_REQ.STATUS - 1] || ''), website: String(v[GC_REQ.WEBSITE - 1] || ''), amazonEmail: String(v[GC_REQ.AMAZON_EMAIL - 1] || ''),
    replySource: String(v[GC_REQ.REPLY_SOURCE - 1] || ''), replyAt: String(v[GC_REQ.REPLY_AT - 1] || ''), reminders: parseInt(v[GC_REQ.REMINDERS - 1], 10) || 0
  };
}

// Open requests (plus anything issued in the last 7 days). Waiting ones are re-checked for replies (email + WhatsApp), at most
// every 10 minutes each and 12 per call, so opening the tab keeps the list current without being slow.
function gcGetRequests(recheck) {
  try {
    var sh = _gcReqSheet(), last = sh.getLastRow();
    if (last < 2) return { success: true, rows: [] };
    var vals = sh.getRange(2, 1, last - 1, GC_REQ_HEADERS.length).getValues();
    var now = new Date().getTime(), checked = 0, out = [];
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i], row = i + 2, status = String(v[GC_REQ.STATUS - 1] || '');
      if (!String(v[GC_REQ.JLID - 1] || '').trim()) continue;
      if (status === 'Waiting' && checked < 12) {
        var lc = v[GC_REQ.CHECKED_AT - 1];
        var lcMs = (Object.prototype.toString.call(lc) === '[object Date]') ? lc.getTime() : 0;
        if (recheck || !lcMs || (now - lcMs) > 600000) {
          checked++;
          var det = gcGetAmazonDetails(String(v[GC_REQ.JLID - 1]), String(v[GC_REQ.EMAILS - 1]), String(v[GC_REQ.PHONE - 1]));
          sh.getRange(row, GC_REQ.CHECKED_AT).setValue(new Date());
          if (det && det.success && det.found) {
            if (det.website) { v[GC_REQ.WEBSITE - 1] = det.website; sh.getRange(row, GC_REQ.WEBSITE).setValue(det.website); }
            if (det.amazonEmail) { v[GC_REQ.AMAZON_EMAIL - 1] = det.amazonEmail; sh.getRange(row, GC_REQ.AMAZON_EMAIL).setValue(det.amazonEmail); }
            v[GC_REQ.REPLY_SOURCE - 1] = det.source || ''; sh.getRange(row, GC_REQ.REPLY_SOURCE).setValue(det.source || '');
            v[GC_REQ.REPLY_AT - 1] = det.repliedAt || ''; sh.getRange(row, GC_REQ.REPLY_AT).setValue(det.repliedAt || '');
            if (det.website && det.amazonEmail) { status = 'Details received'; v[GC_REQ.STATUS - 1] = status; sh.getRange(row, GC_REQ.STATUS).setValue(status); }
          }
        }
      }
      if (status === 'Closed') continue;
      if (status === 'Issued') {
        var ia = v[GC_REQ.ISSUED_AT - 1];
        if (!(Object.prototype.toString.call(ia) === '[object Date]') || (now - ia.getTime()) > 7 * 86400000) continue;
      }
      out.push(_gcReqRowToObj(v, row));
    }
    var order = { 'Details received': 0, 'Waiting': 1, 'Issued': 2 };
    out.sort(function(a, b) { return (order[a.status] - order[b.status]) || (b.waitingDays - a.waitingDays); });
    return { success: true, rows: out, checked: checked };
  } catch (e) {
    Logger.log('[GiftCard] gcGetRequests ERROR: ' + e.message);
    return { success: false, message: e.message, rows: [] };
  }
}

function gcCloseRequest(rowIndex) {
  try {
    var sh = _gcReqSheet();
    sh.getRange(parseInt(rowIndex, 10), GC_REQ.STATUS).setValue('Closed');
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

// A gift card was recorded for this learner: the open request is done
function _gcMarkRequestIssued(jlid) {
  try {
    jlid = String(jlid || '').trim().toUpperCase();
    if (!jlid) return;
    var sh = _gcReqSheet(), last = sh.getLastRow();
    if (last < 2) return;
    var vals = sh.getRange(2, 1, last - 1, GC_REQ_HEADERS.length).getValues();
    vals.forEach(function(v, i) {
      if (String(v[GC_REQ.JLID - 1]).toUpperCase() === jlid && _gcReqIsOpen(String(v[GC_REQ.STATUS - 1]))) {
        sh.getRange(i + 2, GC_REQ.STATUS).setValue('Issued');
        sh.getRange(i + 2, GC_REQ.ISSUED_AT).setValue(new Date());
      }
    });
  } catch (e) { Logger.log('[GiftCard] mark issued failed: ' + e.message); }
}
