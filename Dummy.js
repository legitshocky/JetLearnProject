
// ── TEST: Full class booking for JL39611449152C2 ───────────────────────────────

var TEST_JLID        = 'JL39611449152C2';
var TEST_NEW_TEACHER = 'Sangeeta Sarkar';  // TJL1043 — sangeeta.jetlearn@gmail.com
var TEST_SESSION     = [{ day: 'Monday', time: '5:30 PM' }];
var TEST_IANA        = 'Asia/Kolkata';  // IST

/**
 * Run this first to find the exact teacher name stored in the sheet.
 */
function testTeacherLookup() {
  var rows = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA);
  Logger.log('Total teacher rows: ' + rows.length);
  for (var i = 1; i < rows.length; i++) {
    var name = String(rows[i][1] || '').trim();
    if (name.toLowerCase().indexOf('sangeeta') !== -1) {
      Logger.log('Row ' + i + ': name="' + name + '" teacherId="' + rows[i][0] + '" email="' + rows[i][8] + '"');
    }
  }
}

/**
 * Full booking test — mirrors what the migration page does:
 *  - Fetches learner name + deal data from HubSpot
 *  - Fetches learner/parent emails from deal contacts
 *  - Reads existing calendar event to get class link + description
 *  - Detects event prefix (GMEET, DNRC, etc.) and applies it
 *  - Books with teacher as guest, parent email as guest
 *
 * DRY_RUN = true → logs only, no events created.
 */
function testBookClassWithNewTeacher() {
  var DRY_RUN = false;

  Logger.log('=== testBookClassWithNewTeacher ===');

  // 1. Deal data (learner name, course, dealId)
  var dealRes = fetchMigrationHybridData(TEST_JLID);
  if (!dealRes.success) { Logger.log('Deal fetch failed: ' + dealRes.message); return; }
  var deal        = dealRes.data;
  var learnerName = deal.learnerName || '';
  var courseName  = deal.course || '';
  var dealId      = deal.dealId || '';
  Logger.log('Learner name : ' + learnerName);
  Logger.log('Course       : ' + courseName);
  Logger.log('Deal ID      : ' + dealId);

  // 2. Learner/parent emails from deal contacts
  var extraEmails = [];
  if (dealId) {
    var emailRes = getEmailsForDeal(dealId);
    if (emailRes && emailRes.all && emailRes.all.length) {
      extraEmails = emailRes.all;
    } else if (deal.parentEmail) {
      extraEmails = [deal.parentEmail];
    }
  } else if (deal.parentEmail) {
    extraEmails = [deal.parentEmail];
  }
  Logger.log('Extra emails : ' + JSON.stringify(extraEmails));

  // 3. Existing calendar event — class link, description, and prefix
  var evRes = getExistingEventDescription(TEST_JLID);
  Logger.log('Event title  : ' + (evRes.eventTitle || '—'));

  var classLink    = '';
  var existingDesc = '';
  var eventPrefix  = 'Migration';  // default

  if (evRes.success && evRes.description) {
    var m = evRes.description.match(/https?:\/\/\S+/);
    classLink    = m ? m[0].replace(/[)\]>]+$/, '') : '';
    existingDesc = evRes.description;
  }

  // Extract prefix from existing event title (e.g. "GMEET : Ahaan..." → "GMEET")
  if (evRes.eventTitle) {
    var prefixMatch = evRes.eventTitle.match(/^([A-Z][A-Z0-9\/\- ]+?)\s*:/);
    if (prefixMatch) eventPrefix = prefixMatch[1].trim();
  }

  Logger.log('Class link   : ' + (classLink || '(none)'));
  Logger.log('Event prefix : ' + eventPrefix);

  // 4. Remaining classes
  var remRes    = getRemainingClassesForJlid(TEST_JLID);
  var numEvents = (remRes && remRes.classesLeft > 0) ? remRes.classesLeft : 12;
  Logger.log('Remaining    : ' + numEvents);

  var today     = new Date();
  var pad       = function(n) { return ('0' + n).slice(-2); };
  var startDate = today.getFullYear() + '-' + pad(today.getMonth() + 1) + '-' + pad(today.getDate());

  Logger.log('Start date   : ' + startDate);
  Logger.log('Teacher      : ' + TEST_NEW_TEACHER);
  Logger.log('DRY_RUN      : ' + DRY_RUN);

  if (DRY_RUN) {
    Logger.log('DRY RUN — no events created.');
    return;
  }

  var res = bookClassesWithNewTeacher(
    TEST_JLID,
    learnerName,
    TEST_NEW_TEACHER,
    TEST_SESSION,
    courseName,
    startDate,
    TEST_IANA,
    numEvents,
    extraEmails,
    'TestDummy',
    classLink,
    deal.jetGuideName || '',
    existingDesc,
    eventPrefix      // GMEET / DNRC / Migration etc. — used as first-occurrence title prefix
  );

  Logger.log('Result: ' + JSON.stringify(res));
}
