// ============================================================
// LEARNER OPS SERVICE
// Powers the "Learner Ops" page (formerly Teacher Onboarding Tracker) —
// three feeds giving learner-lifecycle visibility: New Learner Onboarding
// (last N days), Migration Activity (last N days), and Pause & Retention
// (live pipeline stages 15.1–15.5 / 13 in the AI-Coding Pipeline).
// ============================================================

var LOPS_PIPELINE_AI    = 'default';     // AI-Coding Pipeline
var LOPS_PIPELINE_MATHS = '117776157';   // Maths Pipeline

// Deal stage IDs — AI-Coding Pipeline (confirmed via HubSpot property definition)
var LOPS_STAGE_LABELS = {
  '2253783':    '1.1 Deal Created',
  'appointmentscheduled': '3.1 Free First Lesson Scheduled',
  'qualifiedtobuy':       '4. Trial Completed - Decision Pending',
  'contractsent':         '6.1 Trial Agreed to Sign-up',
  'closedwon':            '7. Payment Received',
  '1811602':    '8.1 Schedule Finalisation',
  'decisionmakerboughtin': '8.2 Onboarding Pending',
  '88556899':   '9. Learner Onboarded',
  '51004259':   '10. Upcoming PRMs (Next week)',
  '51023946':   '10.1 Agreed to Renew',
  '51031714':   '11.1 Installment Payment',
  '51004260':   '11.2 Payment Received',
  '117927192':  '12. No Show Absenteeism Check',
  '176052268':  '13. Retained & Save from pause',
  '180246061':  '14. Paid B&R Learners',
  '50954839':   '15.1 Urge on Pause',
  '57285418':   '15.2 Urge on Pause (WIP)',
  '59107960':   '15.3 Pause Save (WIP)',
  '57285419':   '15.4 Pause Follow-up',
  '51043392':   '15.5 Closed Lost'
};

var LOPS_PAUSE_STAGE_IDS = ['50954839', '57285418', '59107960', '57285419', '176052268'];

function _lopsStageLabel(id) { return LOPS_STAGE_LABELS[id] || id || ''; }

// ── Tab 1: New Learner Onboarding — deals whose payment-trigger date falls
// within the last `days` days, regardless of current stage (shows progression
// past Payment Received, and flags anyone stuck at the early stages) ────────
function getLearnerOpsOnboardingFeed(days) {
  try {
    days = days > 0 ? days : 30;
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    var since = new Date(Date.now() - days * 86400000);

    var body = {
      filterGroups: [{
        filters: [
          { propertyName: 'stage____payment_trigger_date', operator: 'GTE', value: String(since.getTime()) },
          { propertyName: 'pipeline', operator: 'IN', values: [LOPS_PIPELINE_AI, LOPS_PIPELINE_MATHS] }
        ]
      }],
      properties: [
        'dealname', 'jetlearner_id', 'dealstage', 'pipeline',
        'current_course__t_', 'current_teacher__t_', 'current_course', 'current_teacher',
        'timelines_chat_link', 'learner_practice_document_link',
        'stage____payment_trigger_date'
      ],
      sorts: [{ propertyName: 'stage____payment_trigger_date', direction: 'DESCENDING' }],
      limit: 100
    };

    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/deals/search', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return { success: false, message: 'HubSpot error ' + resp.getResponseCode() };

    var results = JSON.parse(resp.getContentText()).results || [];
    var rows = results.map(function(r) {
      var p = r.properties || {};
      var rawTeacher = p.current_teacher__t_ || p.current_teacher || '';
      var rawCourse  = p.current_course__t_  || p.current_course  || '';
      var payDate = p.stage____payment_trigger_date ? new Date(parseInt(p.stage____payment_trigger_date, 10) || p.stage____payment_trigger_date) : null;
      return {
        dealId:      p.hs_object_id || r.id,
        jlid:        p.jetlearner_id || '',
        learnerName: p.dealname || '',
        teacher:     rawTeacher ? (getTeacherLabel(rawTeacher) || rawTeacher) : '',
        course:      rawCourse  ? (getCourseLabel(rawCourse)   || rawCourse)  : '',
        stage:       p.dealstage || '',
        stageLabel:  _lopsStageLabel(p.dealstage),
        watiSent:    !!p.timelines_chat_link,
        docLinked:   !!p.learner_practice_document_link,
        paymentDate: payDate && !isNaN(payDate.getTime()) ? Utilities.formatDate(payDate, Session.getScriptTimeZone(), 'dd/MM/yyyy') : ''
      };
    });

    return { success: true, days: days, rows: rows };
  } catch(e) {
    Logger.log('[LOPS] getLearnerOpsOnboardingFeed ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Tab 3: Pause & Retention — live counts + list for stages 15.1–15.5/13 ──
function getLearnerOpsPauseRetention() {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    var dateProps = LOPS_PAUSE_STAGE_IDS.map(function(id) { return 'hs_date_entered_' + id; });

    var body = {
      filterGroups: [{
        filters: [{ propertyName: 'dealstage', operator: 'IN', values: LOPS_PAUSE_STAGE_IDS }]
      }],
      properties: ['dealname', 'jetlearner_id', 'dealstage', 'amount', 'deal_currency_code',
        'current_teacher__t_', 'current_teacher', 'hubspot_owner_id'
      ].concat(dateProps),
      sorts: [{ propertyName: 'hs_lastmodifieddate', direction: 'DESCENDING' }],
      limit: 100
    };

    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/deals/search', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return { success: false, message: 'HubSpot error ' + resp.getResponseCode() };

    var results = JSON.parse(resp.getContentText()).results || [];
    var rows = results.map(function(r) {
      var p = r.properties || {};
      var stageId = p.dealstage || '';
      var enteredRaw = p['hs_date_entered_' + stageId];
      var enteredMs = enteredRaw ? new Date(enteredRaw).getTime() : null;
      var daysInStage = enteredMs ? Math.floor((Date.now() - enteredMs) / 86400000) : null;
      var rawTeacher = p.current_teacher__t_ || p.current_teacher || '';
      return {
        dealId:      p.hs_object_id || r.id,
        jlid:        p.jetlearner_id || '',
        learnerName: p.dealname || '',
        teacher:     rawTeacher ? (getTeacherLabel(rawTeacher) || rawTeacher) : '',
        stage:       stageId,
        stageLabel:  _lopsStageLabel(stageId),
        amount:      p.amount || '',
        currency:    p.deal_currency_code || '',
        daysInStage: daysInStage
      };
    });

    var stats = {};
    LOPS_PAUSE_STAGE_IDS.forEach(function(id) { stats[id] = 0; });
    rows.forEach(function(r) { if (stats.hasOwnProperty(r.stage)) stats[r.stage]++; });

    return { success: true, rows: rows, stats: stats, stageLabels: LOPS_STAGE_LABELS };
  } catch(e) {
    Logger.log('[LOPS] getLearnerOpsPauseRetention ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}
