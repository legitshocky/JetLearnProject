// =============================================
// AI PRODUCT MANAGER SERVICE
// Powered by Google Gemini (Free Tier)
// JetLearn Operations System
// =============================================

const AI_SUGGESTIONS_SHEET = 'AI_Suggestions';
const AI_CHAT_HISTORY_SHEET = 'AI_Chat_History';
const GEMINI_API_KEY_PROP   = 'GEMINI_API_KEY';
const GEMINI_BASE_URL       = 'https://generativelanguage.googleapis.com/v1beta/models/';

// ── Gemini API call helper ────────────────────────────────────────────────────
function _callGemini(contents, model, maxTokens) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(GEMINI_API_KEY_PROP);
  if (!apiKey) throw new Error('GEMINI_API_KEY not set in Script Properties.');

  const selectedModel = model || PropertiesService.getScriptProperties().getProperty('AI_SELECTED_MODEL') || 'gemini-2.5-flash';
  const url = GEMINI_BASE_URL + selectedModel + ':generateContent?key=' + apiKey;

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { 'Content-Type': 'application/json' },
    payload: JSON.stringify({
      contents: contents,
      generationConfig: {
        temperature:     0.4,
        maxOutputTokens: maxTokens || 8192
      }
    }),
    muteHttpExceptions: true
  });

  const raw  = response.getContentText();
  if (response.getResponseCode() !== 200) {
    throw new Error('Gemini API Error (' + response.getResponseCode() + '): ' + raw.substring(0, 200));
  }

  let json;
  try {
    json = JSON.parse(raw);
  } catch (parseErr) {
    throw new Error('Gemini returned non-JSON response: ' + raw.substring(0, 200));
  }

  if (json.error) throw new Error('Gemini API Error: ' + json.error.message);

  // Check finish reason — STOP is good, MAX_TOKENS means truncated
  const candidate    = json.candidates[0];
  const finishReason = candidate.finishReason || '';
  if (finishReason === 'MAX_TOKENS') {
    throw new Error('Response was cut off (too long). Try a simpler request or break it into smaller parts.');
  }

  const text = candidate.content.parts[0].text;
  return _sanitiseJsonResponse(text);
}

// ── JSON response sanitiser ───────────────────────────────────────────────────
function _sanitiseJsonResponse(text) {
  // Strip markdown code fences
  let cleaned = text.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();

  // Extract the outermost JSON object if there's extra text around it
  const firstBrace = cleaned.indexOf('{');
  const lastBrace  = cleaned.lastIndexOf('}');
  if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
    cleaned = cleaned.substring(firstBrace, lastBrace + 1);
  }

  // Attempt direct parse first
  try {
    JSON.parse(cleaned);
    return cleaned;
  } catch (e) {
    Logger.log('[AIService] Initial JSON parse failed: ' + e.message + ' — attempting repair');
  }

  // Repair: replace literal newlines inside string values with \n
  // This handles code blocks inside JSON that Gemini forgets to escape
  let repaired = '';
  let inString = false;
  let escape   = false;
  for (let i = 0; i < cleaned.length; i++) {
    const ch = cleaned[i];
    if (escape) {
      repaired += ch;
      escape = false;
      continue;
    }
    if (ch === '\\') {
      escape = true;
      repaired += ch;
      continue;
    }
    if (ch === '"') {
      inString = !inString;
      repaired += ch;
      continue;
    }
    if (inString && ch === '\n') {
      repaired += '\\n';
      continue;
    }
    if (inString && ch === '\r') {
      repaired += '\\r';
      continue;
    }
    if (inString && ch === '\t') {
      repaired += '\\t';
      continue;
    }
    repaired += ch;
  }

  // Try parse after repair
  try {
    JSON.parse(repaired);
    return repaired;
  } catch (e2) {
    Logger.log('[AIService] JSON repair also failed: ' + e2.message);
    throw new Error('AI response could not be parsed as JSON: ' + e2.message + '. Try approving again.');
  }
}

// ── Script property helpers (for model selector + chat history) ───────────────
function getScriptProperty(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function setScriptProperty(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
  return true;
}

// ── Sheet bootstrap ───────────────────────────────────────────────────────────
function _getOrCreateAISuggestionsSheet() {
  const ss    = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID);
  let   sheet = ss.getSheetByName(AI_SUGGESTIONS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(AI_SUGGESTIONS_SHEET);
    sheet.appendRow([
      'id', 'date', 'area', 'priority', 'problem',
      'suggestion', 'effort', 'status',
      'code_written', 'files_to_change',
      'outcome_notes', 'approved_date', 'rejected_reason'
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 13)
      .setBackground('#4a3c8a').setFontColor('#ffffff').setFontWeight('bold');
    Logger.log('[AIService] Created AI_Suggestions sheet.');
  }
  return sheet;
}

function _getOrCreateChatHistorySheet() {
  const ss    = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID);
  let   sheet = ss.getSheetByName(AI_CHAT_HISTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(AI_CHAT_HISTORY_SHEET);
    sheet.appendRow(['saved_at', 'history_json']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 2)
      .setBackground('#4a3c8a').setFontColor('#ffffff').setFontWeight('bold');
  }
  return sheet;
}

// ── Chat history persistence ──────────────────────────────────────────────────
function saveAIChatHistory(history) {
  try {
    const sheet = _getOrCreateChatHistorySheet();
    const data  = sheet.getDataRange().getValues();
    const json  = JSON.stringify(history);

    if (data.length > 1) {
      // Update existing row
      sheet.getRange(2, 1).setValue(new Date().toISOString());
      sheet.getRange(2, 2).setValue(json);
    } else {
      sheet.appendRow([new Date().toISOString(), json]);
    }
    return { success: true };
  } catch (e) {
    Logger.log('[AIService] saveAIChatHistory error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function getAIChatHistory() {
  try {
    const sheet = _getOrCreateChatHistorySheet();
    const data  = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, history: [] };
    const raw = data[1][1];
    return { success: true, history: raw ? JSON.parse(raw) : [] };
  } catch (e) {
    Logger.log('[AIService] getAIChatHistory error: ' + e.message);
    return { success: true, history: [] };
  }
}

// ── Data snapshot (same as before) ───────────────────────────────────────────
function _collectSystemSnapshot() {
  const snapshot = {};
  try {
    const now      = new Date();
    const cutoff90 = new Date(); cutoff90.setDate(now.getDate() - 90);
    const cutoff30 = new Date(); cutoff30.setDate(now.getDate() - 30);

    const auditData           = _getCachedSheetData(CONFIG.SHEETS.AUDIT_LOG);
    const recentMigrations    = [];
    const onboardings30       = [];
    const migrationsByJlid    = {};
    const migrationsByTeacher = {};

    if (auditData && auditData.length > 1) {
      auditData.slice(1).forEach(row => {
        const d      = parseSheetDate(row[0]);
        const action = String(row[1] || '');
        const jlid   = String(row[2] || '');
        const oldT   = String(row[4] || '');
        if (d && d >= cutoff90 && action.includes('Migration')) {
          recentMigrations.push({ date: row[0], jlid, learner: row[3] || '', oldTeacher: oldT, newTeacher: row[5] || '', reason: row[10] || 'Unspecified' });
          migrationsByJlid[jlid]   = (migrationsByJlid[jlid] || 0) + 1;
          if (oldT) migrationsByTeacher[oldT] = (migrationsByTeacher[oldT] || 0) + 1;
        }
        if (d && d >= cutoff30 && action.includes('Onboard')) {
          onboardings30.push({ date: row[0], jlid, learner: row[3] || '' });
        }
      });
    }

    const frequentMovers = Object.entries(migrationsByJlid)
      .filter(([, c]) => c >= 2).map(([jlid, count]) => ({ jlid, count })).sort((a, b) => b.count - a.count);
    const teachersLosingLearners = Object.entries(migrationsByTeacher)
      .map(([teacher, count]) => ({ teacher, count })).sort((a, b) => b.count - a.count).slice(0, 10);

    snapshot.migrations  = { total90days: recentMigrations.length, frequentMovers, frequentMoverCount: frequentMovers.length, teachersLosingMost: teachersLosingLearners, byReason: _groupBy(recentMigrations, 'reason'), recentSample: recentMigrations.slice(0, 15) };
    snapshot.onboardings = { last30days: onboardings30.length, list: onboardings30.slice(0, 10) };

    const teacherData   = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA);
    const teacherLoad   = [];
    if (teacherData && teacherData.length > 1) {
      teacherData.slice(1).forEach(row => {
        if (row[0]) teacherLoad.push({ name: String(row[0]).trim(), status: String(row[1] || '').trim(), load: Number(row[2]) || 0 });
      });
    }
    const activeTeachers = teacherLoad.filter(t => t.status === 'Active');
    const avgLoad        = activeTeachers.length ? Math.round(activeTeachers.reduce((s, t) => s + t.load, 0) / activeTeachers.length) : 0;
    snapshot.teacherLoad = { totalTeachers: teacherLoad.length, activeTeachers: activeTeachers.length, averageLoad: avgLoad, overloaded: activeTeachers.filter(t => t.load > 20), underloaded: activeTeachers.filter(t => t.load < 5), overloadedCount: activeTeachers.filter(t => t.load > 20).length, underloadedCount: activeTeachers.filter(t => t.load < 5).length };

    const sugSheet        = _getOrCreateAISuggestionsSheet();
    const sugData         = sugSheet.getDataRange().getValues();
    const pastSuggestions = [];
    if (sugData.length > 1) {
      sugData.slice(1).forEach(row => {
        pastSuggestions.push({ id: row[0], date: row[1], area: row[2], suggestion: row[5], status: row[7], outcome: row[10] || '' });
      });
    }
    snapshot.pastSuggestions = pastSuggestions.slice(-20);
    snapshot.projectFiles = [
      'AuditService.js','CacheService.js','Code.js','EmailService.js',
      'HubSpotService.js','InvoiceService.js','ReportService.js',
      'TeacherService.js','UserService.js','Utils.js','ValidationService.js',
      'WatiService.js','AIService.js',
      'Index.html','JavaScript.html','Styles.html','Report.html','AI.html'
    ];
  } catch (e) {
    Logger.log('[AIService] Snapshot error: ' + e.message);
  }
  return snapshot;
}

function _groupBy(arr, key) {
  return arr.reduce((acc, item) => { const k = item[key] || 'Unknown'; acc[k] = (acc[k] || 0) + 1; return acc; }, {});
}

// ── Lean snapshot — summaries only, no raw rows (avoids token overflow) ────────
function _collectLeanSnapshot() {
  const full = _collectSystemSnapshot();

  // Convert byReason object to a short string like "Attrition:97, Course Change:116"
  const byReasonStr = Object.entries(full.migrations.byReason || {})
    .sort((a,b) => b[1] - a[1])
    .slice(0, 6)
    .map(([k,v]) => k + ':' + v)
    .join(', ');

  // Top 3 teachers losing learners as a string
  const teachersLosingStr = (full.migrations.teachersLosingMost || [])
    .slice(0, 3)
    .map(t => t.teacher + '(' + t.count + ')')
    .join(', ');

  // Top 3 frequent movers as a string
  const frequentMoversStr = (full.migrations.frequentMovers || [])
    .slice(0, 3)
    .map(t => t.jlid + 'x' + t.count)
    .join(', ');

  // Top 3 overloaded teachers
  const overloadedStr = (full.teacherLoad.overloaded || [])
    .slice(0, 3)
    .map(t => t.name + ':' + t.load)
    .join(', ');

  // Past suggestions as short strings
  const pastStr = (full.pastSuggestions || [])
    .slice(-8)
    .map(s => '[' + s.status + '] ' + s.suggestion)
    .join(' | ');

  return {
    migrations_total_90d:       full.migrations.total90days,
    migrations_frequent_movers: full.migrations.frequentMoverCount,
    migrations_by_reason:       byReasonStr,
    migrations_top_movers:      frequentMoversStr,
    teachers_losing_most:       teachersLosingStr,
    onboardings_last_30d:       full.onboardings.last30days,
    teachers_total:             full.teacherLoad.totalTeachers,
    teachers_active:            full.teacherLoad.activeTeachers,
    teachers_avg_load:          full.teacherLoad.averageLoad,
    teachers_overloaded_count:  full.teacherLoad.overloadedCount,
    teachers_underloaded_count: full.teacherLoad.underloadedCount,
    teachers_overloaded_top3:   overloadedStr,
    past_suggestions:           pastStr,
    project_files:              'AuditService.js,CacheService.js,Code.js,EmailService.js,HubSpotService.js,InvoiceService.js,ReportService.js,TeacherService.js,UserService.js,Utils.js,ValidationService.js,WatiService.js,AIService.js,Index.html,JavaScript.html,Styles.html,Report.html,AI.html'
  };
}

// ── Run Analysis (now accepts model param) ────────────────────────────────────
function runAIPMAnalysis(model) {
  Logger.log('[AIService] runAIPMAnalysis started with model: ' + model);
  try {
    const snapshot = _collectLeanSnapshot();

    const prompt = 'You are an AI Product Manager for JetLearn (kids coding platform).' +
      ' Analyse this data and return EXACTLY 3 feature suggestions as JSON.' +
      '\nDATA: ' + JSON.stringify(snapshot) +
      '\nRULES: Do not suggest anything in past_suggestions that is Approved/Deployed/Pending. Be brief.' +
      '\nReturn ONLY valid JSON, no markdown, no explanation, starting with { :' +
      '\n{"suggestions":[' +
      '{"id":"id1","area":"Attrition|Teacher Load|Onboarding|Migrations|Milestones",' +
      '"priority":"Critical|High|Medium|Low","problem":"one sentence with numbers",' +
      '"suggestion":"short title","description":"one sentence","effort":"Low|Medium|High",' +
      '"files_to_change":["File.js"],"data_evidence":"key number from data"}' +
      ']}';

    const responseText = _callGemini([{ parts: [{ text: prompt }] }], model, 2048);
    const parsed       = JSON.parse(responseText);
    if (!parsed.suggestions || !Array.isArray(parsed.suggestions)) {
      return { success: false, message: 'AI returned unexpected format. Please try again.' };
    }

    const sheet     = _getOrCreateAISuggestionsSheet();
    const timestamp = new Date().toISOString();
    parsed.suggestions.forEach(s => {
      if (!s.id) s.id = 'sug_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
      sheet.appendRow([
        s.id, timestamp, s.area || '', s.priority || 'Medium',
        s.problem || '', s.suggestion || '', s.effort || 'Medium',
        'Pending', '', (s.files_to_change || []).join(', '), '', '', ''
      ]);
    });

    Logger.log('[AIService] ' + parsed.suggestions.length + ' suggestions saved.');
    return { success: true, suggestions: parsed.suggestions };
  } catch (e) {
    Logger.log('[AIService] runAIPMAnalysis error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Load suggestions ──────────────────────────────────────────────────────────
function getAISuggestions() {
  try {
    const sheet = _getOrCreateAISuggestionsSheet();
    const data  = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, suggestions: [] };

    const suggestions = data.slice(1).map(row => ({
      id: String(row[0] || ''), date: row[1] || '', area: String(row[2] || ''),
      priority: String(row[3] || 'Medium'), problem: String(row[4] || ''),
      suggestion: String(row[5] || ''), effort: String(row[6] || 'Medium'),
      status: String(row[7] || 'Pending'), codeWritten: String(row[8] || ''),
      filesToChange: String(row[9] || ''), outcomeNotes: String(row[10] || ''),
      approvedDate: String(row[11] || ''), rejectedReason: String(row[12] || '')
    }));

    suggestions.sort((a, b) => {
      if (a.status === 'Pending' && b.status !== 'Pending') return -1;
      if (b.status === 'Pending' && a.status !== 'Pending') return  1;
      return new Date(b.date) - new Date(a.date);
    });

    return { success: true, suggestions };
  } catch (e) {
    Logger.log('[AIService] getAISuggestions error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Approve → generates code (FIXED: saves full codeResult JSON) ──────────────
function approveSuggestion(suggestionId, model) {
  Logger.log('[AIService] approveSuggestion: ' + suggestionId);
  try {
    const sheet = _getOrCreateAISuggestionsSheet();
    const data  = sheet.getDataRange().getValues();
    let targetRow = -1, suggestion = null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(suggestionId)) {
        targetRow  = i + 1;
        suggestion = { id: data[i][0], area: data[i][2], priority: data[i][3], problem: data[i][4], suggestion: data[i][5], effort: data[i][6], filesToChange: data[i][9] };
        break;
      }
    }

    if (!suggestion) return { success: false, message: 'Suggestion not found: ' + suggestionId };

    const codePrompt = 'You are an expert Google Apps Script developer for JetLearn.' +
      '\n\nApproved feature to implement:' +
      '\nFEATURE: ' + suggestion.suggestion +
      '\nAREA: ' + suggestion.area +
      '\nPROBLEM SOLVED: ' + suggestion.problem +
      '\nFILES TO CHANGE: ' + suggestion.filesToChange +
      '\n\nTECHNICAL CONTEXT:' +
      '\n- Backend: Google Apps Script .js files' +
      '\n- Frontend: HTML + vanilla JS (UI logic in JavaScript.html)' +
      '\n- Data: Google Sheets via _getCachedSheetData(sheetName)' +
      '\n- Bridge: google.script.run.withSuccessHandler().functionName()' +
      '\n- UI helpers: showToast(type, title, msg), showLoading(msg), hideLoading()' +
      '\n- CONFIG object has SHEETS, EMAIL, MIGRATION_SHEET_ID' +
      '\n\nCRITICAL JSON RULES:' +
      '\n- Return ONLY a valid JSON object, no markdown, no code fences' +
      '\n- All code goes inside the "code" string field' +
      '\n- Escape ALL special characters in code strings: use \\n for newlines, \\\\ for backslashes, \\" for quotes' +
      '\n- Keep each file under 80 lines to avoid truncation' +
      '\n- If a file needs more code, split into multiple file entries' +
      '\n\nReturn ONLY this JSON structure:' +
      '\n{"files":[{"filename":"ExactFileName.js","action":"add_function OR modify_function","description":"What this does","code":"// escaped code here"}],"instructions":"1. Step one 2. Step two","testing":"How to verify"}';

    const responseText = _callGemini([{ parts: [{ text: codePrompt }] }], model, 4096);
    const codeResult   = JSON.parse(responseText);

    if (!codeResult.files || !Array.isArray(codeResult.files)) {
      return { success: false, message: 'AI returned unexpected code format. Try approving again.' };
    }

    sheet.getRange(targetRow, 8).setValue('Approved');
    sheet.getRange(targetRow, 9).setValue(JSON.stringify(codeResult)); // ✅ FIXED: full JSON
    sheet.getRange(targetRow, 12).setValue(new Date().toISOString());

    Logger.log('[AIService] Code generated for: ' + suggestionId);
    return { success: true, suggestion, codeResult };
  } catch (e) {
    Logger.log('[AIService] approveSuggestion error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Get approved code (FIXED: returns full codeResult) ───────────────────────
function getApprovedCode(suggestionId) {
  try {
    const sheet = _getOrCreateAISuggestionsSheet();
    const data  = sheet.getDataRange().getValues();

    Logger.log('[AIService] getApprovedCode looking for id: ' + suggestionId + ' in ' + (data.length - 1) + ' rows');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(suggestionId)) {
        const raw = String(data[i][8] || '').trim(); // col 9 = index 8 = code_written

        Logger.log('[AIService] Found row, raw code_written length: ' + raw.length);

        if (!raw) {
          return { success: false, message: 'No code stored yet. Please click Approve to generate the code first.' };
        }

        let parsed;
        try {
          parsed = JSON.parse(raw);
        } catch (parseErr) {
          Logger.log('[AIService] JSON parse error: ' + parseErr.message);
          return { success: false, message: 'Stored code is corrupted. Please approve the suggestion again to regenerate it.' };
        }

        // Handle OLD format: was just an array of filenames e.g. ["AIService.js"]
        // Handle NEW format: full codeResult object { files: [...], instructions: ..., testing: ... }
        if (Array.isArray(parsed)) {
          return { success: false, message: 'This suggestion was approved before the fix. Please approve it again to regenerate the code.' };
        }

        if (!parsed.files || !Array.isArray(parsed.files) || parsed.files.length === 0) {
          return { success: false, message: 'Code data is incomplete. Please approve the suggestion again.' };
        }

        return { success: true, codeResult: parsed };
      }
    }

    return { success: false, message: 'Suggestion not found in sheet.' };
  } catch (e) {
    Logger.log('[AIService] getApprovedCode error: ' + e.message);
    return { success: false, message: e.message };
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// CHAT ENGINE — Proper multi-turn architecture matching AI Studio
// ══════════════════════════════════════════════════════════════════════════════

// The system instruction is sent ONCE as a separate field (not as a turn in history).
// This matches exactly how AI Studio / Gemini API is designed to work:
//   systemInstruction → sets the persona + context permanently
//   contents          → only the actual conversation turns (user/model alternating)
// This means:
//   - System context never consumes conversation turn tokens
//   - History grows naturally without re-sending system prompt every time
//   - Model maintains persona across all turns correctly

// ══════════════════════════════════════════════════════════════════════════════
// sendAIChat — full multimodal, settings-aware, token-reporting chat engine
// params:
//   history  — [{role, parts:[{text}]}]  conversation turns
//   model    — model string e.g. 'gemini-2.5-flash'
//   settings — { temperature, topP, maxTokens, sysOverride }
//   parts    — current turn parts [{type:'text'|'image'|'pdf', ...}]
//              (only last user turn; history already contains previous turns as text)
// ══════════════════════════════════════════════════════════════════════════════
function sendAIChat(history, model, settings, parts) {
  settings = settings || {};
  parts    = parts    || [];
  Logger.log('[AIService] sendAIChat — turns: ' + (history||[]).length + ', model: ' + model + ', attachments: ' + parts.filter(function(p){return p.type!=='text';}).length);

  try {
    const snapshot = _collectLeanSnapshot();

    // ── System instruction: use override if set, else default ────────────────
    const defaultSystem =
      'You are an expert AI Product Manager and senior Google Apps Script developer for JetLearn — ' +
      'an online coding education platform for children aged 6-18. ' +
      'You have deep knowledge of the JetLearn operations system and access to live data.' +

      '\n\n## YOUR CAPABILITIES' +
      '\n- Analyse operational data and identify problems with specific numbers' +
      '\n- Suggest prioritised product features with clear business justification' +
      '\n- Write complete, production-ready Google Apps Script and HTML/JS code' +
      '\n- Explain technical concepts clearly to a non-technical product owner' +
      '\n- Remember the full conversation context and build on previous messages' +
      '\n- Analyse images, screenshots, wireframes, and PDFs when attached' +

      '\n\n## LIVE SYSTEM DATA (updated each session)' +
      '\n' + JSON.stringify(snapshot) +

      '\n\n## CODEBASE CONTEXT' +
      '\nBackend: AuditService.js, CacheService.js, Code.js, EmailService.js, HubSpotService.js, ' +
      'InvoiceService.js, ReportService.js, TeacherService.js, UserService.js, Utils.js, ' +
      'ValidationService.js, WatiService.js, AIService.js' +
      '\nFrontend: Index.html, JavaScript.html, Styles.html, Report.html, AI.html' +
      '\nData: _getCachedSheetData(sheetName), _getSpreadsheet(id)' +
      '\nBridge: google.script.run.withSuccessHandler(fn).backendFn(args)' +
      '\nUI: showToast(type,title,msg), showLoading(msg), hideLoading()' +

      '\n\n## RESPONSE FORMATTING' +
      '\n- Use **bold** for key terms and numbers' +
      '\n- Use ## and ### headers to organise long answers' +
      '\n- Use bullet and numbered lists freely' +
      '\n- ALL code in fenced code blocks with language tag: ```javascript or ```html' +
      '\n- Tables in markdown table syntax' +
      '\n- Always cite specific numbers from live data' +
      '\n- Write code completely — never say "add your logic here"' +
      '\n- When an image is attached: describe what you see and give specific advice based on it' +
      '\n- When a PDF is attached: read and reference specific sections' +
      '\n- When a code file is attached: reference specific line numbers and function names' +

      '\n\n## SUGGESTION CARDS' +
      '\nEnd responses with ```suggestion_card JSON when asked to create a feature suggestion.' +

      '\n\n## PERSONALITY' +
      '\nBe direct, confident, data-driven. Give strong recommendations. Push back if data contradicts the user\'s idea.';

    const systemText = (settings.sysOverride && settings.sysOverride.trim())
      ? settings.sysOverride.trim()
      : defaultSystem;

    const systemInstruction = { parts: [{ text: systemText }] };

    // ── Build previous history turns ─────────────────────────────────────────
    // All turns except the last (which is the current turn we're building from parts[])
    const prevHistory = (history || []).slice(0, -1);
    const contents    = _sanitiseChatHistory(prevHistory);

    // ── Build current user turn from parts array ─────────────────────────────
    // parts: [{type:'text', text}, {type:'image', mimeType, data}, {type:'pdf', mimeType, data}]
    const currentParts = [];
    parts.forEach(function(p) {
      if (p.type === 'text') {
        currentParts.push({ text: p.text });
      } else if (p.type === 'image') {
        currentParts.push({ inlineData: { mimeType: p.mimeType, data: p.data } });
      } else if (p.type === 'pdf') {
        // Gemini supports PDF as inlineData with application/pdf
        currentParts.push({ inlineData: { mimeType: 'application/pdf', data: p.data } });
      }
    });

    // Fallback: if parts was empty (old call style), use last history turn
    if (currentParts.length === 0 && history && history.length > 0) {
      const lastTurn = history[history.length - 1];
      const lastText = (lastTurn.parts && lastTurn.parts[0] && lastTurn.parts[0].text) || '';
      currentParts.push({ text: lastText });
    }

    contents.push({ role: 'user', parts: currentParts });

    // ── Generation config from settings ──────────────────────────────────────
    const genConfig = {
      temperature:     typeof settings.temperature === 'number' ? settings.temperature : 0.7,
      topP:            typeof settings.topP        === 'number' ? settings.topP        : 0.95,
      topK:            40,
      maxOutputTokens: typeof settings.maxTokens   === 'number' ? settings.maxTokens   : 8192,
      candidateCount:  1
    };

    // ── Call Gemini ───────────────────────────────────────────────────────────
    const apiKey = PropertiesService.getScriptProperties().getProperty(GEMINI_API_KEY_PROP);
    if (!apiKey) throw new Error('GEMINI_API_KEY not set in Script Properties.');

    const selectedModel = model
      || PropertiesService.getScriptProperties().getProperty('AI_SELECTED_MODEL')
      || 'gemini-2.5-flash';

    const url = GEMINI_BASE_URL + selectedModel + ':generateContent?key=' + apiKey;

    const payload = {
      systemInstruction: systemInstruction,
      contents:          contents,
      generationConfig:  genConfig,
      safetySettings: [
        { category: 'HARM_CATEGORY_HARASSMENT',        threshold: 'BLOCK_ONLY_HIGH' },
        { category: 'HARM_CATEGORY_HATE_SPEECH',       threshold: 'BLOCK_ONLY_HIGH' },
        { category: 'HARM_CATEGORY_SEXUALLY_EXPLICIT', threshold: 'BLOCK_ONLY_HIGH' },
        { category: 'HARM_CATEGORY_DANGEROUS_CONTENT', threshold: 'BLOCK_ONLY_HIGH' }
      ]
    };

    Logger.log('[AIService] → Gemini: model=' + selectedModel +
      ', turns=' + contents.length +
      ', temp=' + genConfig.temperature +
      ', maxTok=' + genConfig.maxOutputTokens);

    const response = UrlFetchApp.fetch(url, {
      method:             'post',
      headers:            { 'Content-Type': 'application/json' },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const raw  = response.getContentText();
    if (response.getResponseCode() !== 200) {
      throw new Error('Gemini API Error (' + response.getResponseCode() + '): ' + raw.substring(0, 200));
    }

    let json;
    try {
      json = JSON.parse(raw);
    } catch (parseErr) {
      throw new Error('Gemini returned non-JSON response: ' + raw.substring(0, 200));
    }

    if (json.error) throw new Error('Gemini API Error: ' + json.error.message);

    if (!json.candidates || json.candidates.length === 0) {
      const blocked = json.promptFeedback && json.promptFeedback.blockReason;
      throw new Error(blocked ? 'Blocked: ' + blocked : 'No response from Gemini.');
    }

    const candidate    = json.candidates[0];
    const finishReason = candidate.finishReason || 'STOP';

    const replyText = (candidate.content.parts || []).map(function(p){ return p.text || ''; }).join('');
    if (!replyText.trim()) throw new Error('Gemini returned an empty response.');

    // ── Extract suggestion card ───────────────────────────────────────────────
    var suggestion = null;
    var cardMatch  = replyText.match(/```suggestion_card[\s\S]*?({[\s\S]*?})[\s\S]*?```/);
    if (cardMatch) {
      try {
        suggestion    = JSON.parse(cardMatch[1].trim());
        suggestion.id = 'chat_sug_' + Date.now();
      } catch(e) { Logger.log('[AIService] suggestion_card parse failed: ' + e.message); }
    }

    // ── Token usage ───────────────────────────────────────────────────────────
    var tokenCount = 0;
    if (json.usageMetadata) {
      tokenCount = json.usageMetadata.totalTokenCount || 0;
      Logger.log('[AIService] Tokens — prompt: ' + json.usageMetadata.promptTokenCount +
        ', output: ' + json.usageMetadata.candidatesTokenCount +
        ', total: ' + tokenCount);
    }

    return {
      success:      true,
      reply:        replyText,
      suggestion:   suggestion,
      finishReason: finishReason,
      truncated:    finishReason === 'MAX_TOKENS',
      tokenCount:   tokenCount
    };

  } catch(e) {
    Logger.log('[AIService] sendAIChat error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function _sanitiseChatHistory(history) {
  if (!history || history.length === 0) {
    return [{ role: 'user', parts: [{ text: 'Hello' }] }];
  }

  const ATTACHMENT_PATTERN = /\*\*\[Attached file: ([^\]]+)\]\*\*\n```[\w]*\n[\s\S]*?\n```/g;
  const MAX_TURN_CHARS     = 8000;   // ~2000 tokens per old turn — keeps context tight
  const lastTurnIdx        = history.length - 1;

  const sanitised = [];
  let   lastRole  = null;

  history.forEach((turn, idx) => {
    const role = turn.role === 'model' ? 'model' : 'user';
    let   text = (turn.parts && turn.parts[0] && turn.parts[0].text) || '';
    if (!text.trim()) return;

    // For all turns EXCEPT the last one (current message being sent):
    // replace full file contents with a compact placeholder.
    // This stops 200KB files being re-transmitted every single request.
    if (idx < lastTurnIdx && role === 'user') {
      text = text.replace(ATTACHMENT_PATTERN, function(_, filename) {
        return '[Previously attached file: ' + filename + ' — contents already seen]';
      });
    }

    // Hard cap on any single old turn to prevent history bloat
    if (idx < lastTurnIdx && text.length > MAX_TURN_CHARS) {
      text = text.substring(0, MAX_TURN_CHARS) + '\n...[truncated for context window]';
    }

    if (role === lastRole) {
      sanitised[sanitised.length - 1].parts[0].text += '\n' + text;
    } else {
      sanitised.push({ role: role, parts: [{ text: text }] });
      lastRole = role;
    }
  });

  if (sanitised.length > 0 && sanitised[0].role !== 'user') {
    sanitised.unshift({ role: 'user', parts: [{ text: 'Continue.' }] });
  }

  // Log token estimate
  const totalChars = sanitised.reduce((sum, t) => sum + t.parts[0].text.length, 0);
  Logger.log('[AIService] History after sanitise — turns: ' + sanitised.length +
    ', chars: ' + totalChars + ' (~' + Math.round(totalChars/4) + ' tokens)');

  return sanitised;
}



// ── Save chat-generated suggestion ───────────────────────────────────────────
function saveChatSuggestion(s) {
  try {
    const sheet     = _getOrCreateAISuggestionsSheet();
    const timestamp = new Date().toISOString();
    if (!s.id) s.id = 'chat_sug_' + Date.now();
    sheet.appendRow([
      s.id, timestamp, s.area || 'General', s.priority || 'Medium',
      s.problem || '', s.suggestion || '', s.effort || 'Medium',
      'Pending', '', (s.files_to_change || []).join(', '), '', '', ''
    ]);
    return { success: true, id: s.id };
  } catch (e) {
    Logger.log('[AIService] saveChatSuggestion error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Reject ────────────────────────────────────────────────────────────────────
function rejectSuggestion(suggestionId, reason) {
  try {
    const sheet = _getOrCreateAISuggestionsSheet();
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(suggestionId)) {
        sheet.getRange(i + 1, 8).setValue('Rejected');
        sheet.getRange(i + 1, 13).setValue(reason || 'No reason provided');
        return { success: true };
      }
    }
    return { success: false, message: 'Suggestion not found.' };
  } catch (e) {
    Logger.log('[AIService] rejectSuggestion error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Mark Deployed ─────────────────────────────────────────────────────────────
function markSuggestionDeployed(suggestionId, outcomeNotes) {
  try {
    const sheet = _getOrCreateAISuggestionsSheet();
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(suggestionId)) {
        sheet.getRange(i + 1, 8).setValue('Deployed');
        sheet.getRange(i + 1, 11).setValue(outcomeNotes || '');
        return { success: true };
      }
    }
    return { success: false, message: 'Suggestion not found.' };
  } catch (e) {
    Logger.log('[AIService] markSuggestionDeployed error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Snooze ────────────────────────────────────────────────────────────────────
function snoozeSuggestion(suggestionId) {
  try {
    const sheet = _getOrCreateAISuggestionsSheet();
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(suggestionId)) {
        sheet.getRange(i + 1, 8).setValue('Snoozed');
        return { success: true };
      }
    }
    return { success: false, message: 'Suggestion not found.' };
  } catch (e) {
    Logger.log('[AIService] snoozeSuggestion error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Unsnooze (FIXED) ──────────────────────────────────────────────────────────
function unsnoozeSuggestion(suggestionId) {
  try {
    const sheet = _getOrCreateAISuggestionsSheet();
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(suggestionId)) {
        sheet.getRange(i + 1, 8).setValue('Pending');
        return { success: true };
      }
    }
    return { success: false, message: 'Suggestion not found.' };
  } catch (e) {
    Logger.log('[AIService] unsnoozeSuggestion error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function rankReplacementTeachersWithAI(targetTeacherName, candidates, targetContext) {
  try {
    Logger.log('[rankReplacementTeachersWithAI] Starting AI ranking for: ' + targetTeacherName);

    var apiKey = PropertiesService.getScriptProperties().getProperty(GEMINI_API_KEY_PROP);
    if (!apiKey) {
      Logger.log('[rankReplacementTeachersWithAI] No Gemini API key — skipping AI ranking.');
      return { success: false, message: 'No Gemini API key set.' };
    }

    // Compact candidate summaries — only what AI needs, nothing extra
    var candidateSummaries = candidates.map(function(c, idx) {
      return {
        rank:           idx + 1,
        name:           c.name,
        score:          c.matchScore,
        upskillCount:   c.upskillCount    || c.activeLearners || 0,
        upskillDiff:    c.upskillDiff     || 0,
        ageGroupMatch:  c.ageGroupMatch   || 'None',
        escalations:    c.escalations     || 0,
        escalationRisk: c.escalationRisk  || 'Clean'
      };
    });

    var targetUpskillCount = targetContext.upskillCount  || 0;
    var targetAgeGroups    = (targetContext.ageGroups    || []).join(', ') || 'N/A';
    var targetEscalations  = targetContext.escalations   || 0;
    var prompt =
      'You are a Teacher Ops Manager at JetLearn (kids coding platform).' +
      ' Pick the TOP 5 replacements for "' + targetTeacherName + '".' +
      '\n\nTARGET: ' + targetUpskillCount + ' courses, age groups: ' + targetAgeGroups + ', escalations: ' + targetEscalations +
      '\n\nCANDIDATES:\n' + JSON.stringify(candidateSummaries) +
      '\n\nPRIORITY ORDER: 1=stability(no escalations) 2=age group match 3=upskill count closeness 4=score' +
      '\n\nReturn ONLY this JSON (keep all string values under 20 words, no markdown):' +
      '\n{"summary":"<one short sentence>",' +
      '"recommendations":[' +
      '{"rank":1,"name":"<n>","matchScore":<num>,"confidence":"High|Medium|Low",' +
      '"reason":"<max 15 words citing one number>",' +
      '"warning":"<empty string or max 10 words>",' +
      '"caveat":"<max 12 words>"}' +
      ']}';

    var selectedModel = PropertiesService.getScriptProperties().getProperty('AI_SELECTED_MODEL') || 'gemini-2.5-flash';
    var url = GEMINI_BASE_URL + selectedModel + ':generateContent?key=' + apiKey;

    var payload = {
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: {
        temperature:     0.1,
        maxOutputTokens: 800
      }
    };

    var response = UrlFetchApp.fetch(url, {
      method:             'post',
      headers:            { 'Content-Type': 'application/json' },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var raw  = response.getContentText();
    if (response.getResponseCode() !== 200) {
      Logger.log('[rankReplacementTeachersWithAI] Gemini API error (' + response.getResponseCode() + '): ' + raw.substring(0, 200));
      return { success: false, message: 'Gemini API Error (' + response.getResponseCode() + ')' };
    }

    var json;
    try {
      json = JSON.parse(raw);
    } catch (parseErr) {
      Logger.log('[rankReplacementTeachersWithAI] Non-JSON response: ' + raw.substring(0, 200));
      return { success: false, message: 'Gemini returned non-JSON response.' };
    }

    if (json.error) {
      Logger.log('[rankReplacementTeachersWithAI] Gemini error: ' + json.error.message);
      return { success: false, message: json.error.message };
    }

    if (!json.candidates || !json.candidates[0]) {
      Logger.log('[rankReplacementTeachersWithAI] No candidates in response.');
      return { success: false, message: 'No response from Gemini.' };
    }

    var replyText = json.candidates[0].content.parts[0].text;
    Logger.log('[rankReplacementTeachersWithAI] Response length: ' + replyText.length + ' chars');
    Logger.log('[rankReplacementTeachersWithAI] Raw (first 400): ' + replyText.substring(0, 400));

    // ── Parse with multiple fallback strategies ───────────────────────────
    var parsed = null;

    // Strategy 1: standard sanitiser (strips markdown fences)
    try {
      var cleanedJson = _sanitiseJsonResponse(replyText);
      parsed = JSON.parse(cleanedJson);
      Logger.log('[rankReplacementTeachersWithAI] Strategy 1 (sanitiser) succeeded.');
    } catch (e1) {
      Logger.log('[rankReplacementTeachersWithAI] Strategy 1 failed: ' + e1.message);
    }

    // Strategy 2: find first { and last } in raw text and try parsing that slice
    if (!parsed) {
      try {
        var first = replyText.indexOf('{');
        var last  = replyText.lastIndexOf('}');
        if (first !== -1 && last > first) {
          parsed = JSON.parse(replyText.substring(first, last + 1));
          Logger.log('[rankReplacementTeachersWithAI] Strategy 2 (slice) succeeded.');
        }
      } catch (e2) {
        Logger.log('[rankReplacementTeachersWithAI] Strategy 2 failed: ' + e2.message);
      }
    }

    // Strategy 3: if response was truncated mid-array, try to recover
    // partial recommendations by extracting complete {...} objects from the array
    if (!parsed) {
      try {
        var summaryMatch = replyText.match(/"summary"\s*:\s*"([^"]+)"/);
        var summary = summaryMatch ? summaryMatch[1] : 'AI analysis partially available.';

        var recMatches = replyText.match(/\{[^{}]*"rank"\s*:\s*\d+[^{}]*\}/g);
        var partialRecs = [];
        if (recMatches) {
          recMatches.forEach(function(m) {
            try {
              var obj = JSON.parse(m);
              if (obj.name && obj.rank) partialRecs.push(obj);
            } catch (ignored) {}
          });
        }

        if (partialRecs.length > 0) {
          parsed = { summary: summary, recommendations: partialRecs };
          Logger.log('[rankReplacementTeachersWithAI] Strategy 3 (partial recovery) got ' + partialRecs.length + ' recs.');
        }
      } catch (e3) {
        Logger.log('[rankReplacementTeachersWithAI] Strategy 3 failed: ' + e3.message);
      }
    }

    if (!parsed || !parsed.recommendations || !Array.isArray(parsed.recommendations) || parsed.recommendations.length === 0) {
      Logger.log('[rankReplacementTeachersWithAI] All parse strategies failed — falling back.');
      return { success: false, message: 'Could not parse AI response after 3 attempts.' };
    }

    Logger.log('[rankReplacementTeachersWithAI] Parsed ' + parsed.recommendations.length + ' AI recommendations.');

    // Merge AI reasoning back onto original candidate objects
    var enriched = parsed.recommendations.map(function(aiRec) {
      var original = null;
      for (var i = 0; i < candidates.length; i++) {
        if (candidates[i].name.toLowerCase() === String(aiRec.name || '').toLowerCase()) {
          original = candidates[i];
          break;
        }
      }
      original = original || {};

      return {
        name:            aiRec.name        || original.name || 'Unknown',
        matchScore:      aiRec.matchScore  || original.matchScore      || 0,
        traitScore:      original.traitScore   || 0,
        ageScore:        original.ageScore     || 0,
        courseScore:     original.courseScore  || 0,
        escalationScore: original.escalationScore || 0,
        loadScore:       0,
        courseOverlap:   original.courseOverlap || 'N/A',
        overlapCount:    original.overlapCount  || 0,
        activeLearners:  original.activeLearners || 0,
        upskillCount:    original.upskillCount  || 0,
        upskillDiff:     original.upskillDiff   || 0,
        ageGroupMatch:   original.ageGroupMatch || 'N/A',
        escalations:     original.escalations   || 0,
        escalationRisk:  original.escalationRisk  || 'No Escalations',
        escalationColor: original.escalationColor || '#15803d',
        stability:       original.stability || { total: 0, risk: 'Stable' },
        aiRank:       aiRec.rank       || (parsed.recommendations.indexOf(aiRec) + 1),
        aiConfidence: aiRec.confidence || 'Medium',
        aiReason:     aiRec.reason     || '',
        aiWarning:    aiRec.warning    || '',
        aiCaveat:     aiRec.caveat     || '',
        aiEnriched:   true
      };
    }).slice(0, 5); // ✅ FIXED: guarantee max 5 results

    return {
      success:   true,
      data:      enriched,
      aiSummary: parsed.summary || ''
    };

  } catch (e) {
    Logger.log('[rankReplacementTeachersWithAI] Unexpected error: ' + e.message);
    return { success: false, message: e.message };
  }
}