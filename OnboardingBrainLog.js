// ═══════════════════════════════════════════════════════════════════════════
// ONBOARDING REVIEW LOG
// One row per onboarding send: what the deal-driven "brain" proposed, what was actually sent, what changed, and the deal's raw numbers.
// A second tab has one row per changed field, ready for a pivot. Meant for a review in December: where does the brain get it wrong?
// Never blocks an onboarding: every failure here is swallowed and logged.
// ═══════════════════════════════════════════════════════════════════════════
var OBL_SHEET_ID = '1iNrejNX3HA01UqYEch94HuKLQCffSPofB8KbD4D9sI4';
var OBL_MAIN_GID = 330124279;
var OBL_CHANGES_TAB = 'Onboarding Brain Changes';

var OBL_MAIN_HEADERS = ['Timestamp', 'Sent By', 'JLID', 'Learner', 'Deal ID', 'Result', 'Brain Ran', 'Fields Changed', 'Changed Fields (proposed -> final)', 'Brain Warnings',
  'Deal: Plan', 'Deal: Tenure (months)', 'Deal: Amount', 'Deal: Currency', 'Deal: Payment Type', 'Deal: Installment Type', 'Deal: Installment Terms', 'Deal: Installment Months',
  'Deal: Classes / Week', 'Deal: Structured Classes', 'Deal: Free Classes', 'Deal: Journey Total', 'Deal: Start', 'Deal: End',
  'Final: Plan', 'Final: Tenure', 'Final: Free Classes', 'Final: Sessions / Week', 'Final: Start', 'Final: End', 'Final: Discount', 'Final: Payment Type', 'Final: Installments', 'Final: Frequency', 'Final: First Payment', 'Final: Total Payable',
  'Proposed (JSON)', 'Version'];
var OBL_CHANGES_HEADERS = ['Timestamp', 'Sent By', 'JLID', 'Learner', 'Field', 'Brain Proposed', 'Final Sent', 'Deal: Plan', 'Deal: Amount', 'Deal: Structured', 'Deal: Free', 'Deal: Classes / Week', 'Deal: Payment Type', 'Brain Warnings'];

function _oblMainSheet() {
  var ss = SpreadsheetApp.openById(OBL_SHEET_ID);
  var sh = null;
  ss.getSheets().forEach(function(s) { if (s.getSheetId() === OBL_MAIN_GID) sh = s; });
  if (!sh) sh = ss.insertSheet('Onboarding Review Log');
  if (sh.getLastRow() === 0) { sh.appendRow(OBL_MAIN_HEADERS); sh.setFrozenRows(1); }
  return sh;
}

function _oblChangesSheet() {
  var ss = SpreadsheetApp.openById(OBL_SHEET_ID);
  var sh = ss.getSheetByName(OBL_CHANGES_TAB);
  if (!sh) { sh = ss.insertSheet(OBL_CHANGES_TAB); sh.appendRow(OBL_CHANGES_HEADERS); sh.setFrozenRows(1); }
  return sh;
}

function _oblStr(v) { return (v === null || v === undefined) ? '' : String(v); }

// Compare the brain's proposal with what was sent. Numbers compare as numbers so "12" and "12.0" are the same.
function _oblSame(a, b) {
  var x = _oblStr(a).trim(), y = _oblStr(b).trim();
  if (x === y) return true;
  var isNum = function(t) { return /^[^0-9\-]{0,3}-?\d+(\.\d+)?$/.test(t); };   // 12, 12.00, €999.00 (not dates)
  if (!isNum(x) || !isNum(y)) return false;
  return Math.abs(parseFloat(x.replace(/[^0-9.\-]/g, '')) - parseFloat(y.replace(/[^0-9.\-]/g, ''))) < 0.005;
}

// review = { ran, proposed: {...}, final: {...}, warnings: [...], deal: {...} } built by the page; result = 'Success' | 'Failed'
function logOnboardingBrain(formData, result) {
  try {
    var r = (formData && formData.brainReview) || {};
    var ran = !!r.ran, p = r.proposed || {}, f = r.final || {}, d = r.deal || {};
    var keys = ['planName', 'tenure', 'freeClasses', 'sessionsPerWeek', 'startDate', 'endDate', 'discount', 'paymentType', 'installments', 'frequency', 'firstPaymentDate'];
    var changed = [];
    if (ran) keys.forEach(function(k) { if (!_oblSame(p[k], f[k])) changed.push({ field: k, proposed: _oblStr(p[k]), final: _oblStr(f[k]) }); });
    var warnings = (r.warnings || []).join(' | ');
    var now = new Date(), who = formData.performedBy || '';
    var main = _oblMainSheet();
    main.appendRow([now, who, formData.jlid || '', formData.learnerName || '', formData.dealId || '', result || '', ran ? 'Yes' : 'No', ran ? changed.length : '',
      changed.map(function(c) { return c.field + ': ' + c.proposed + ' -> ' + c.final; }).join('; '), warnings,
      d.plan || '', d.tenure || '', d.amount || '', d.currency || '', d.paymentType || '', d.installmentType || '', d.installmentTerms || '', d.installmentMonths || '',
      d.classesPerWeek || '', d.structured || '', d.free || '', d.journey || '', d.start || '', d.end || '',
      f.planName || '', f.tenure || '', f.freeClasses || '', f.sessionsPerWeek || '', f.startDate || '', f.endDate || '', f.discount || '', f.paymentType || '', f.installments || '', f.frequency || '', f.firstPaymentDate || '', f.totalPayable || '',
      ran ? JSON.stringify(p) : '', (typeof APP_VERSION !== 'undefined') ? APP_VERSION : '']);
    if (changed.length) {
      var cs = _oblChangesSheet();
      changed.forEach(function(c) {
        cs.appendRow([now, who, formData.jlid || '', formData.learnerName || '', c.field, c.proposed, c.final, d.plan || '', d.amount || '', d.structured || '', d.free || '', d.classesPerWeek || '', d.paymentType || '', warnings]);
      });
    }
    return { success: true };
  } catch (e) {
    Logger.log('[OnboardingBrainLog] failed: ' + e.message);
    return { success: false, message: e.message };
  }
}
