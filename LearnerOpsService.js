// ============================================================
// LEARNER OPS SERVICE
// Powers the "Learner Ops" page (formerly Teacher Onboarding Tracker) —
// three feeds giving learner-lifecycle visibility across BOTH the AI-Coding
// and Maths deal pipelines: New Learner Onboarding (last N days), Migration
// Activity (last N days), and Pause & Retention (live pause pipeline stages).
// ============================================================

var LOPS_PIPELINE_AI    = 'default';     // AI-Coding Pipeline
var LOPS_PIPELINE_MATHS = '117776157';   // Maths Pipeline
var LOPS_PIPELINE_SUBJECT = {};
LOPS_PIPELINE_SUBJECT[LOPS_PIPELINE_AI]    = 'Coding';
LOPS_PIPELINE_SUBJECT[LOPS_PIPELINE_MATHS] = 'Maths';

var LOPS_STAGE_CACHE_KEY = 'LOPS_PIPELINE_STAGES_V1';
var LOPS_STAGE_CACHE_TTL = 1800; // 30 min

// Fetches ALL deal pipelines + their stages directly from HubSpot's canonical
// pipelines endpoint (not a hand-maintained ID list) — so both AI-Coding and
// Maths pipeline stages are always accurate, even if HubSpot renumbers them.
// Returns { stageLabels: {stageId: label}, stageSubject: {stageId: 'Coding'|'Maths'|...} }
function _lopsGetPipelineStageMaps() {
  try {
    var cached = CacheService.getScriptCache().get(LOPS_STAGE_CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch(e) {}

  var stageLabels = {};
  var stageSubject = {};
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/pipelines/deals', {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() === 200) {
      var data = JSON.parse(resp.getContentText());
      (data.results || []).forEach(function(pipeline) {
        var subject = LOPS_PIPELINE_SUBJECT[pipeline.id] || pipeline.label || '';
        (pipeline.stages || []).forEach(function(stage) {
          stageLabels[stage.id] = stage.label;
          stageSubject[stage.id] = subject;
        });
      });
    } else {
      Logger.log('[LOPS] pipelines fetch HTTP ' + resp.getResponseCode());
    }
  } catch(e) {
    Logger.log('[LOPS] _lopsGetPipelineStageMaps ERROR: ' + e.message);
  }

  var result = { stageLabels: stageLabels, stageSubject: stageSubject };
  try { CacheService.getScriptCache().put(LOPS_STAGE_CACHE_KEY, JSON.stringify(result), LOPS_STAGE_CACHE_TTL); } catch(e) {}
  return result;
}

// Every stage across both pipelines whose label matches the pause/retention
// flow: 15.1 Urge on Pause → 15.2 (WIP) → 15.3 Pause Save (WIP) → 13. Retained
// & Save, or 15.4 Pause Follow-up if not saved. Excludes 15.5 Closed Lost —
// that's terminal churn, a different concern from active retention work.
function _lopsGetPauseStageIds(stageLabels) {
  var ids = [];
  for (var id in stageLabels) {
    if (/^(13\.|15\.[1-4]\b)/.test(stageLabels[id])) ids.push(id);
  }
  return ids;
}

function clearLearnerOpsPipelineCache() {
  try { CacheService.getScriptCache().remove(LOPS_STAGE_CACHE_KEY); } catch(e) {}
  return 'Cleared — next load will re-fetch pipeline stages from HubSpot.';
}

// ── Tab 1: New Learner Onboarding — deals whose payment-trigger date falls
// within the last `days` days, regardless of current stage, across BOTH
// pipelines (shows progression past Payment Received, flags stuck ones) ────
function getLearnerOpsOnboardingFeed(days) {
  try {
    days = days > 0 ? days : 30;
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    var since = new Date(Date.now() - days * 86400000);
    var maps = _lopsGetPipelineStageMaps();

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
        subject:     LOPS_PIPELINE_SUBJECT[p.pipeline] || p.pipeline || '',
        stage:       p.dealstage || '',
        stageLabel:  maps.stageLabels[p.dealstage] || p.dealstage || '',
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

// ── Tab 3: Pause & Retention — live counts + list across BOTH pipelines ────
function getLearnerOpsPauseRetention() {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    var maps = _lopsGetPipelineStageMaps();
    var pauseStageIds = _lopsGetPauseStageIds(maps.stageLabels);
    if (!pauseStageIds.length) return { success: false, message: 'Could not resolve pause-stage IDs from HubSpot pipelines.' };

    var dateProps = pauseStageIds.map(function(id) { return 'hs_date_entered_' + id; });

    var body = {
      filterGroups: [{
        filters: [{ propertyName: 'dealstage', operator: 'IN', values: pauseStageIds }]
      }],
      properties: ['dealname', 'jetlearner_id', 'dealstage', 'pipeline', 'amount', 'deal_currency_code',
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
        subject:     maps.stageSubject[stageId] || LOPS_PIPELINE_SUBJECT[p.pipeline] || '',
        stage:       stageId,
        stageLabel:  maps.stageLabels[stageId] || stageId,
        amount:      p.amount || '',
        currency:    p.deal_currency_code || '',
        daysInStage: daysInStage
      };
    });

    var stats = {};
    pauseStageIds.forEach(function(id) { stats[id] = 0; });
    rows.forEach(function(r) { if (stats.hasOwnProperty(r.stage)) stats[r.stage]++; });

    return { success: true, rows: rows, stats: stats, stageLabels: maps.stageLabels, pauseStageIds: pauseStageIds };
  } catch(e) {
    Logger.log('[LOPS] getLearnerOpsPauseRetention ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}
