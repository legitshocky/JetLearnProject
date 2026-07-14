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
    var directRes = monitoredFetch(directUrl, {
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

function testFetchDesc() {
  var res = getExistingEventDescription('JL31741746342C'); // replace with real JLID
  Logger.log(JSON.stringify(res));
}

function runTestChatLinkDummy() {
  var phone = '8369118156'; // <-- change this number
  var res = fetchWatiDirectLink(phone);
  Logger.log(JSON.stringify(res));
}


// ═══════════════════════════════════════════════════════════════════════════════
// CET TIME CONVERSION TEST — run: node Dummy.js (or paste in browser console)
// ═══════════════════════════════════════════════════════════════════════════════

function _testConvertCetToTimezone(cetDay, cetTime, timezoneStr, cetOffsetHours) {
    if (!cetDay || !cetTime) return null;
    cetTime = String(cetTime).replace(/_/g, ' ').replace(/\bam\b/i, 'AM').replace(/\bpm\b/i, 'PM').trim();
    var h, m;
    var ampmMatch  = String(cetTime).match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
    var noMinMatch = String(cetTime).match(/^(\d{1,2})\s*(AM|PM)$/i);
    var h24Match   = String(cetTime).match(/^(\d{1,2}):(\d{2})/);
    if (ampmMatch) {
        h = parseInt(ampmMatch[1]); m = parseInt(ampmMatch[2]);
        var mer = ampmMatch[3].toUpperCase();
        if (mer === 'PM' && h !== 12) h += 12;
        if (mer === 'AM' && h === 12) h = 0;
    } else if (noMinMatch) {
        h = parseInt(noMinMatch[1]); m = 0;
        var mer2 = noMinMatch[2].toUpperCase();
        if (mer2 === 'PM' && h !== 12) h += 12;
        if (mer2 === 'AM' && h === 12) h = 0;
    } else if (h24Match) {
        h = parseInt(h24Match[1]); m = parseInt(h24Match[2]);
    } else { return null; }
    var cetOffset = (cetOffsetHours || 1) * 60;
    var utcMins = h * 60 + m - cetOffset;
    var offMatch = String(timezoneStr || '').match(/GMT\s*([+-])\s*(\d{1,2}):(\d{2})/i);
    var offsetMins = 0;
    if (offMatch) offsetMins = (offMatch[1] === '+' ? 1 : -1) * (parseInt(offMatch[2]) * 60 + parseInt(offMatch[3]));
    var localMins = utcMins + offsetMins;
    var DAYS = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    var dayIdx = DAYS.findIndex(function(d){ return d.toLowerCase() === String(cetDay).trim().toLowerCase(); });
    if (dayIdx === -1) return null;
    var dayOffset = 0;
    if (localMins < 0) { localMins += 1440; dayOffset = -1; }
    else if (localMins >= 1440) { localMins -= 1440; dayOffset = 1; }
    var localDay = DAYS[(dayIdx + dayOffset + 7) % 7];
    var lh = Math.floor(localMins / 60), lm = localMins % 60;
    var lmer = lh >= 12 ? 'PM' : 'AM';
    var dh = lh % 12 || 12;
    return { day: localDay, time: String(dh).padStart(2,'0') + ':' + String(lm).padStart(2,'0') + ' ' + lmer };
}

function runCetConversionTests() {
    var tests = [
        // [label, cetDay, rawHubSpotTime, timezone, cetOffset, expectedDay, expectedTime]
        ['IST CET  "5 pm"',    'tuesday',  '5 pm',   '(GMT +5:30) Bombay, Calcutta, Madras, New Delhi', 1, 'Tuesday',   '09:30 PM'],
        ['IST CEST "5 pm"',    'tuesday',  '5 pm',   '(GMT +5:30) Bombay, Calcutta, Madras, New Delhi', 2, 'Tuesday',   '08:30 PM'],
        ['IST CET  "9 am"',    'thursday', '9 am',   '(GMT +5:30) Bombay, Calcutta, Madras, New Delhi', 1, 'Thursday',  '01:30 PM'],
        ['IST CET  "12:30 AM"','monday',   '12:30 AM','(GMT +5:30) Bombay, Calcutta, Madras, New Delhi',1, 'Monday',    '05:00 AM'],
        ['UTC CET  "5 pm"',    'tuesday',  '5 pm',   '(GMT +0:00) UTC',                                 1, 'Tuesday',   '04:00 PM'],
        ['GST CET  "10 am"',   'wednesday','10 am',  '(GMT +4:00) Abu Dhabi, Muscat',                   1, 'Wednesday', '01:00 PM'],
        ['EST CET  "5 pm"',    'friday',   '5 pm',   '(GMT -5:00) Eastern Time (US & Canada)',           1, 'Friday',    '11:00 AM'],
        ['Day back',           'monday',   '1 am',   '(GMT -3:00) Brasilia',                             1, 'Sunday',    '09:00 PM'],
        ['Day forward',        'friday',   '11 pm',  '(GMT +2:00) Cairo',                                1, 'Saturday',  '12:00 AM'],
        ['Null: no time',      'tuesday',  '',       '(GMT +5:30) Bombay, Calcutta, Madras, New Delhi', 1, null, null],
        ['Null: bad day',      'funday',   '5 pm',   '(GMT +5:30) Bombay, Calcutta, Madras, New Delhi', 1, null, null],
    ];

    var pass = 0, fail = 0;
    tests.forEach(function(t) {
        var label = t[0], cetDay = t[1], cetTime = t[2], tz = t[3], off = t[4], expDay = t[5], expTime = t[6];
        var result = _testConvertCetToTimezone(cetDay, cetTime, tz, off);
        var ok = expDay === null
            ? result === null
            : result && result.day === expDay && result.time === expTime;
        if (ok) {
            Logger.log('✅ PASS  ' + label + ' → ' + (result ? result.day + ' ' + result.time : 'null'));
            pass++;
        } else {
            Logger.log('❌ FAIL  ' + label);
            Logger.log('        Expected: ' + expDay + ' ' + expTime);
            Logger.log('        Got:      ' + (result ? result.day + ' ' + result.time : 'null'));
            fail++;
        }
    });
    Logger.log('\n' + pass + ' passed  /  ' + fail + ' failed');
}


// ── TEST WITH REAL JLID ───────────────────────────────────────────────────────
// Set your JLID below and run testCetForJlid() from Apps Script editor

function testCetForJlid() {
  var JLID = 'JL-XXXX'; // ← CHANGE THIS

  Logger.log('=== CET Time Fetch Test for: ' + JLID + ' ===');

  var result = fetchHubspotByJlid(JLID);

  if (!result.success) {
    Logger.log('❌ Fetch failed: ' + result.message);
    return;
  }

  var data = result.data;

  Logger.log('Learner:   ' + data.learnerName);
  Logger.log('Timezone:  ' + data.timezone);
  Logger.log('classSessions:    ' + JSON.stringify(data.classSessions));
  Logger.log('cetClassSessions: ' + JSON.stringify(data.cetClassSessions));

  if (!data.cetClassSessions || data.cetClassSessions.length === 0) {
    Logger.log('❌ cetClassSessions is empty — HubSpot has no class day/time for this deal');
    return;
  }

  var s = data.cetClassSessions[0];
  Logger.log('\nRaw CET → day: "' + s.day + '"  time: "' + s.time + '"');

  if (!s.time) {
    Logger.log('❌ No CET time found — regular_class_time_in_cet is blank on this deal');
    return;
  }

  // Simulate conversion (CET = UTC+1)
  var converted = _testConvertCetToTimezone(s.day, s.time, data.timezone, 1);
  Logger.log('Converted (CET):  ' + JSON.stringify(converted));

  // Simulate conversion (CEST = UTC+2)
  var convertedSummer = _testConvertCetToTimezone(s.day, s.time, data.timezone, 2);
  Logger.log('Converted (CEST): ' + JSON.stringify(convertedSummer));

  if (converted) {
    Logger.log('\n✅ Success — form would show: ' + converted.day + '  ' + converted.time + ' (winter) / ' + convertedSummer.time + ' (summer)');
  } else {
    Logger.log('\n❌ Conversion returned null — check timezone format: "' + data.timezone + '"');
  }
}
