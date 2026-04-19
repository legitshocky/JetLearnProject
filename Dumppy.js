/**
 * DEBUG: Tests ONLY the 3 fixed templates.
 */
function debugTestSpecificTemplates() {
  // --- 1. ENTER YOUR PHONE NUMBER HERE ---
  const MY_PHONE = "918369118156"; // <--- REPLACE THIS (No + symbol)
  // ---------------------------------------

  const TARGET_TEMPLATES = [
    "migration_teacher_affinity"
  ];

  Logger.log("🚀 Starting Specific Template Test...");
  
  // 2. Mock Data (Ensuring all variables like {{Course}}, {{Date}}, {{Weekday}} are present)
  const mockMigration = {
    learner: "Robin",
    newTeacher: "Batman", 
    oldTeacher: "Alfred",
    course: "Crime Fighting 101",
    classLink: "https://meet.google.com/bat-cave",
    startDate: "2026-02-01", // For {{Date}}
    classSessions: [{ day: "Monday", time: "08:00 PM" }], // For {{Weekday}}
    manualTimezone: "Asia/Kolkata",
    calculatedLocalTime: "10:00 AM IST" // For {{Time}}
  };

  const mockHubSpot = {
    parentName: "Sourav Pal",
    parentContact: MY_PHONE,
    timezone: "Asia/Kolkata"
  };

  // 3. Iterate and Send
  TARGET_TEMPLATES.forEach(templateId => {
      Logger.log(`\n---------------------------------`);
      Logger.log(`Testing: ${templateId}`);

      try {
        // Generate Params using the UPDATED logic
        const params = getWatiParameters(templateId, mockMigration, mockHubSpot);
        
        Logger.log("Params Generated: " + JSON.stringify(params));

        // Send
        const res = sendWatiMessage(MY_PHONE, templateId, params);
        
        if (res.success) {
          Logger.log(`✅ SUCCESS`);
        } else {
          Logger.log(`❌ FAILED`);
        }

      } catch (e) {
        Logger.log(`🚨 CRASH: ${e.message}`);
      }

      Utilities.sleep(1000);
  });

  Logger.log("\n✅ Done.");
}


// ─────────────────────────────────────────────────────────────────────────────
// TEST: Fallback email content for a real JLID + template
// Run this in Apps Script editor → Execution Log
// ─────────────────────────────────────────────────────────────────────────────
function testFallbackEmailContent() {
  var JLID        = 'JL39611449152C2';
  var TEMPLATE    = 'course_change_teacher_change_migration';

  Logger.log('=== FALLBACK EMAIL CONTENT TEST ===');
  Logger.log('JLID: ' + JLID);
  Logger.log('Template: ' + TEMPLATE);
  Logger.log('-----------------------------------');

  // 1. Fetch HubSpot data
  Logger.log('\n[1] Fetching HubSpot data...');
  var hsResult = fetchHubspotByJlid(JLID);
  if (!hsResult.success) {
    Logger.log('❌ HubSpot fetch failed: ' + hsResult.message);
    return;
  }
  var hsData = hsResult.data;
  Logger.log('✅ HubSpot OK');
  Logger.log('   Parent name:  ' + hsData.parentName);
  Logger.log('   Parent email: ' + hsData.parentEmail);
  Logger.log('   Learner:      ' + hsData.learnerName);
  Logger.log('   Course:       ' + hsData.course);

  // 2. Mock migration context (same shape as the form submission)
  var mockMigration = {
    jlid:              JLID,
    learner:           hsData.learnerName || 'Test Learner',
    newTeacher:        'Batman',        // replace with real teacher to test
    oldTeacher:        hsData.currentTeacher || 'Previous Teacher',
    course:            hsData.course || 'App It Up',
    classLink:         hsData.zoomLink || 'https://live.jetlearn.com/login',
    startDate:         '2026-04-25',
    watiTemplateName:  TEMPLATE,
    reasonOfMigration: 'Course Change Teacher Change',
    classSessions:     hsData.classSessions && hsData.classSessions.length > 0
                         ? hsData.classSessions
                         : [{ day: 'Monday', time: '06:00 PM' }],
    manualTimezone:    hsData.timezone || 'Europe/London'
  };

  // 3. Build WATI params — same as what would go to WhatsApp
  Logger.log('\n[2] Building WATI parameters...');
  var paramsList = getWatiParameters(TEMPLATE, mockMigration, hsData);
  Logger.log('✅ Params built:');
  paramsList.forEach(function(p) {
    Logger.log('   {{' + p.name + '}} = "' + p.value + '"');
  });

  // 4. Test direct ?templateName= lookup (single API call)
  Logger.log('\n[3] Testing direct template lookup...');
  try {
    var props = PropertiesService.getScriptProperties();
    var base  = (props.getProperty('WATI_API_ENDPOINT') || '').trim();
    var token = (props.getProperty('WATI_ACCESS_TOKEN') || '').trim();
    if (!token.startsWith('Bearer ')) token = 'Bearer ' + token;
    if (base.endsWith('/')) base = base.slice(0, -1);

    var directUrl = base + '/api/v1/getMessageTemplates?pageSize=1&templateName=' + encodeURIComponent(TEMPLATE);
    Logger.log('   Direct URL: ' + directUrl);
    var directRes = UrlFetchApp.fetch(directUrl, {
      method: 'get',
      headers: { Authorization: token },
      muteHttpExceptions: true
    });
    Logger.log('   HTTP: ' + directRes.getResponseCode());
    var directData = JSON.parse(directRes.getContentText());
    var directTemplates = directData.messageTemplates || directData.templates || directData.items || directData.data || [];
    Logger.log('   Results count: ' + directTemplates.length);
    if (directTemplates.length > 0) {
      var dt = directTemplates[0];
      var returnedName = dt.elementName || dt.name || '?';
      Logger.log('   elementName returned: ' + returnedName);
      if (returnedName.toLowerCase() === TEMPLATE.toLowerCase()) {
        Logger.log('   ✅ Name match — direct lookup WORKS, body: ' + (dt.body || '(empty)').substring(0, 100));
      } else {
        Logger.log('   ⚠ Name MISMATCH (got "' + returnedName + '") — WATI filter is not exact, falling back to pagination');
      }
    } else {
      Logger.log('   ⚠ Direct lookup returned 0 results — will fall back to pagination');
    }
  } catch(e) {
    Logger.log('   ❌ Exception: ' + e.message);
  }

  var templateBody = _fetchWatiTemplateBody(TEMPLATE);
  if (!templateBody) {
    Logger.log('\n❌ _fetchWatiTemplateBody returned null (see debug above)');
  } else {
    Logger.log('✅ Template body fetched:');
    Logger.log('--- RAW ---');
    Logger.log(templateBody);
    var filled = _substituteWatiParams(templateBody, paramsList);
    Logger.log('\n--- FILLED (what email will say) ---');
    Logger.log(filled);
  }

  // 6. Preview the full HTML that would be emailed
  Logger.log('\n[4] Building email HTML...');
  var html = getMigrationParentFallbackEmailHTML(
    templateBody
      ? { messageText: _substituteWatiParams(templateBody, paramsList), learnerName: mockMigration.learner, classLink: mockMigration.classLink }
      : { messageText: null, learnerName: mockMigration.learner,
          newTeacher: mockMigration.newTeacher, oldTeacher: mockMigration.oldTeacher,
          course: mockMigration.course, classLink: mockMigration.classLink }
  );
  Logger.log('✅ HTML length: ' + html.length + ' chars');
  Logger.log('   Subject would be: JetLearn - ' + mockMigration.learner + ' - Class Update');
  Logger.log('   Send to: ' + hsData.parentEmail);

  Logger.log('\n=== TEST COMPLETE ===');
  Logger.log('If filled text looks correct above, the real email will send identical content.');
}