// ============================================================
// PRACTICE DOC SERVICE
// Creates practice docs from template, shares with teacher + parent.
// Also posts onboarding verification note to HubSpot deal.
// ============================================================

var PD_TEMPLATE_ID    = '1bS_ogNmOQmWyPuBxwqz18r_ZngLUPF4BxQBiM4r8GwQ';
var PD_FOLDER_ID      = '1tL2edxPIZYTtrVgalyciJt9IMJmG44ja';  // Practice Docs shared drive folder
var PD_SUPPORT_EMAIL  = 'support@jet-learn.com';

// ── Subject label from JLID suffix ───────────────────────────
function _pdSubject(jlid) {
  var m = String(jlid || '').match(/JL\d+([A-Z0-9]+)$/i);
  var suffix = m ? m[1].toUpperCase() : '';
  var map = { 'C': 'AI- Coding', 'C2': 'AI- Coding', 'M': 'Maths', 'FL': 'FinLit' };
  return map[suffix] || 'AI- Coding';
}

// ── Update teacher on an existing practice doc ───────────────
// Removes any previous teacher editors (by cross-referencing Teacher Data sheet),
// adds the new teacher. Keeps support@ and parent commenter untouched.
function _updateExistingPracticeDocTeacher(docUrl, newTeacherName) {
  try {
    var match = String(docUrl || '').match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!match) {
      Logger.log('[PracticeDoc] _updateExistingPracticeDocTeacher: cannot extract file ID from: ' + docUrl);
      return { success: false, error: 'Cannot extract file ID from URL' };
    }
    var file = DriveApp.getFileById(match[1]);

    // Build set of all known teacher emails from Teacher Data sheet
    var allTeacherEmails = {};
    try {
      var tdData = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA);
      for (var i = 1; i < tdData.length; i++) {
        var em = String(tdData[i][8] || '').trim().toLowerCase();
        if (em) allTeacherEmails[em] = true;
      }
    } catch(te) { Logger.log('[PracticeDoc] Teacher email lookup error: ' + te.message); }

    var newTeacherEmail = _pdTeacherEmail(newTeacherName);
    var newEmailLow     = (newTeacherEmail || '').toLowerCase();

    // Remove any current editor who is a known teacher but NOT the new teacher
    var editors = file.getEditors();
    editors.forEach(function(ed) {
      var em = ed.getEmail().toLowerCase();
      if (em === PD_SUPPORT_EMAIL.toLowerCase()) return; // always keep support
      if (em === newEmailLow) return;                    // keep new teacher
      if (allTeacherEmails[em]) {
        file.removeEditor(ed.getEmail());
        Logger.log('[PracticeDoc] Removed old teacher: ' + em);
      }
    });

    // Add new teacher as editor
    if (newTeacherEmail) {
      file.addEditor(newTeacherEmail);
      Logger.log('[PracticeDoc] Added new teacher: ' + newTeacherEmail + ' to existing doc');
    } else {
      Logger.log('[PracticeDoc] New teacher "' + newTeacherName + '" not found in sheet — skipped editor update');
    }

    return { success: true, url: docUrl, teacherEmail: newTeacherEmail };
  } catch(e) {
    Logger.log('[PracticeDoc] _updateExistingPracticeDocTeacher ERROR: ' + e.message);
    return { success: false, error: e.message, url: docUrl };
  }
}

// ── Fetch existing practice doc URL from HubSpot deal ────────
function _pdFetchExistingDocUrl(dealId) {
  try {
    if (!dealId) return '';
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    var resp  = monitoredFetch(
      'https://api.hubapi.com/crm/v3/objects/deals/' + dealId + '?properties=learner_practice_document_link',
      { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (resp.getResponseCode() === 200) {
      return JSON.parse(resp.getContentText()).properties.learner_practice_document_link || '';
    }
  } catch(e) { Logger.log('[PracticeDoc] _pdFetchExistingDocUrl error: ' + e.message); }
  return '';
}

// ── Teacher email by name from Teacher Data sheet ─────────────
function _pdTeacherEmail(teacherName) {
  try {
    var data = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA);
    var nameLow = String(teacherName || '').toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1] || '').toLowerCase().trim() === nameLow) {
        return String(data[i][8] || '').trim(); // col 9 = email
      }
    }
    return '';
  } catch(e) {
    Logger.log('[PracticeDoc] _pdTeacherEmail error: ' + e.message);
    return '';
  }
}

// ── Teacher ID + Name lookup (returns "TJL1280 - Johncy Paul") ─
function _pdTeacherIdName(teacherName) {
  try {
    var data    = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA);
    var nameLow = String(teacherName || '').toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1] || '').toLowerCase().trim() === nameLow) {
        var tid  = String(data[i][0] || '').trim(); // col A = TJL code
        var name = String(data[i][1] || '').trim();
        return tid ? tid + ' - ' + name : name;
      }
    }
    return teacherName || '';
  } catch(e) {
    Logger.log('[PracticeDoc] _pdTeacherIdName error: ' + e.message);
    return teacherName || '';
  }
}

// ── Public: create + share practice doc ──────────────────────
function createPracticeDoc(jlid, learnerName, teacherName, parentEmail) {
  try {
    var subject   = _pdSubject(jlid);
    // FinLit uses " : " separator; Coding/Maths use single space
    var separator = (subject === 'FinLit') ? ' : ' : ' ';
    var docName   = 'JetLearn ' + subject + ' Practice Doc' + separator + learnerName + ' (' + jlid + ')';

    var template = DriveApp.getFileById(PD_TEMPLATE_ID);

    // Destination folder: script property overrides hardcoded default
    var folderId = PropertiesService.getScriptProperties().getProperty('PRACTICE_DOC_FOLDER_ID') || PD_FOLDER_ID;
    var folder = DriveApp.getFolderById(folderId);

    var newFile = template.makeCopy(docName, folder);
    var url     = newFile.getUrl();

    // support@ = editor (creates on behalf of support)
    newFile.addEditor(PD_SUPPORT_EMAIL);

    // teacher = editor
    var teacherEmail = _pdTeacherEmail(teacherName);
    if (teacherEmail) {
      newFile.addEditor(teacherEmail);
      Logger.log('[PracticeDoc] Teacher editor: ' + teacherEmail);
    }

    // parent = commenter
    if (parentEmail) {
      newFile.addCommenter(parentEmail);
      Logger.log('[PracticeDoc] Parent commenter: ' + parentEmail);
    }

    Logger.log('[PracticeDoc] Created: ' + docName + ' → ' + url);
    return { success: true, url: url, name: docName, teacherEmail: teacherEmail };

  } catch(e) {
    Logger.log('[PracticeDoc] createPracticeDoc ERROR: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── Public: post onboarding verification note to HubSpot deal ─
// noteData = {
//   totalAmount, currency, teacherName, course,
//   classTimings, timezone, committedClasses, practiceDocUrl
// }
function postOnboardingNote(dealId, noteData) {
  try {
    if (!dealId || !noteData) return { success: false, error: 'Missing dealId or noteData' };

    // Format today as DD-MM-YYYY
    var now   = new Date();
    var dd    = String(now.getDate()).padStart(2, '0');
    var mm    = String(now.getMonth() + 1).padStart(2, '0');
    var yyyy  = now.getFullYear();
    var today = dd + '-' + mm + '-' + yyyy;

    // Payment: strip symbols, combine amount + currency code e.g. "400GBP"
    var payment = String(noteData.totalAmount || '').replace(/[£€$\s]/g, '') + (noteData.currency || '');

    // Teacher: "TJL1280 - Johncy Paul"
    var teacherLabel = _pdTeacherIdName(noteData.teacherName);

    var lines = [
      'Payment : '           + payment,
      'Date : '              + today,
      'Course : '            + (noteData.course || ''),
      'Teacher upskilled : ' + teacherLabel,
      'TZ : '                + (noteData.timezone || ''),
      'CO : '                + (noteData.committedClasses || ''),
      'Time : '              + (noteData.classTimings || 'TBD'),
    ];

    if (noteData.practiceDocUrl) {
      lines.push('Practice Doc : ' + noteData.practiceDocUrl);
    }

    addNoteToHubSpotDeal(dealId, lines.join('\n'));
    Logger.log('[PracticeDoc] Onboarding note posted to deal ' + dealId);
    return { success: true };

  } catch(e) {
    Logger.log('[PracticeDoc] postOnboardingNote ERROR: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── TEST: run from GAS editor to test practice doc creation ──
// Select testCreatePracticeDoc in the function dropdown → Run
function testCreatePracticeDoc() {
  var JLID = 'JL39611449152C2';
  Logger.log('=== TEST createPracticeDoc for ' + JLID + ' ===');

  // 1. Fetch deal from HubSpot
  var res = fetchHubspotByJlid(JLID);
  if (!res || !res.success || !res.data) {
    Logger.log('FAIL: could not fetch deal — ' + JSON.stringify(res));
    return;
  }
  var d = res.data;
  Logger.log('Deal found: dealId=' + d.dealId + ' name=' + d.learnerName + ' teacher=' + d.currentTeacher + ' email=' + d.parentEmail);

  // 2. Strip "TJL1280 - " prefix so _pdTeacherEmail can match by name
  var teacherNameOnly = (d.currentTeacher || '').replace(/^TJL\d+\s*-\s*/i, '').trim();
  Logger.log('teacherNameOnly: "' + teacherNameOnly + '"');

  // 2b. Check template + folder access before creating
  var templateId = PropertiesService.getScriptProperties().getProperty('PRACTICE_DOC_TEMPLATE_ID') || PD_TEMPLATE_ID;
  var folderId   = PropertiesService.getScriptProperties().getProperty('PRACTICE_DOC_FOLDER_ID')   || PD_FOLDER_ID;
  Logger.log('Template ID: ' + templateId);
  Logger.log('Folder ID  : ' + folderId);
  try {
    var tpl = DriveApp.getFileById(templateId);
    Logger.log('Template OK: ' + tpl.getName());
  } catch(te) {
    Logger.log('TEMPLATE ACCESS FAILED: ' + te.message);
    Logger.log('Fix: share the template doc with the account running this script, giving Editor access.');
  }
  try {
    var fldr = DriveApp.getFolderById(folderId);
    Logger.log('Folder OK  : ' + fldr.getName());
  } catch(fe) {
    Logger.log('FOLDER ACCESS FAILED: ' + fe.message);
    Logger.log('Fix: share the destination folder with the account running this script, giving Editor access.');
  }

  // 3. Create the practice doc
  var docRes = createPracticeDoc(JLID, d.learnerName, teacherNameOnly, d.parentEmail);
  Logger.log('createPracticeDoc result: ' + JSON.stringify(docRes));

  if (!docRes.success) {
    Logger.log('FAIL: doc creation failed — ' + docRes.error);
    return;
  }

  Logger.log('Doc name : ' + docRes.name);
  Logger.log('Doc URL  : ' + docRes.url);
  Logger.log('Teacher email shared: ' + (docRes.teacherEmail || '(none — teacher not found in sheet)'));

  // 4. Patch HubSpot deal property
  if (d.dealId) {
    var patchOk = _obcPatchDeal(d.dealId, { learner_practice_document_link: docRes.url });
    Logger.log('HubSpot PATCH learner_practice_document_link: ' + (patchOk ? 'SUCCESS' : 'FAILED'));
  } else {
    Logger.log('WARNING: no dealId — HubSpot not patched');
  }

  Logger.log('=== TEST COMPLETE ===');
}

// ── Public: one-shot — create doc + post note together ────────
// Called from client after successful onboarding email send.
function createPracticeDocAndPostNote(jlid, learnerName, teacherName, parentEmail, dealId, noteData) {
  // Strip TJL prefix so _pdTeacherEmail can match by name
  var teacherNameOnly = String(teacherName || '').replace(/^TJL\d+\s*-\s*/i, '').trim() || teacherName;

  // ── Check if a practice doc already exists for this deal ──
  var existingUrl = _pdFetchExistingDocUrl(dealId);

  var docResult;
  if (existingUrl) {
    // Reuse existing doc — update teacher permissions instead of creating a new one
    Logger.log('[PracticeDoc] Existing doc found for deal ' + dealId + ', updating teacher to: ' + teacherNameOnly);
    var updateRes = _updateExistingPracticeDocTeacher(existingUrl, teacherNameOnly);
    docResult = {
      success:     true,
      url:         existingUrl,
      name:        '(existing doc — teacher updated)',
      teacherEmail: updateRes.teacherEmail || '',
      reused:      true
    };
  } else {
    // No existing doc — create new one
    docResult = createPracticeDoc(jlid, learnerName, teacherNameOnly, parentEmail);
  }

  if (docResult.success && dealId) {
    var patchOk = _obcPatchDeal(dealId, { learner_practice_document_link: docResult.url });
    Logger.log('[PracticeDoc] HubSpot patch dealId=' + dealId + ' ok=' + patchOk + ' url=' + docResult.url + ' reused=' + (!!docResult.reused));
    if (!patchOk) docResult.hsWarning = 'HubSpot link PATCH failed';
    if (noteData) {
      noteData.practiceDocUrl = docResult.url;
      postOnboardingNote(dealId, noteData);
    }
  }
  return docResult;
}
