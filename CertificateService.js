/**
 * CertificateService.js
 * ─────────────────────────────────────────────────────────────────────
 * Generates course completion certificates from a Google Slides template
 * and emails them as PDF attachments to the parent on CCTC migration.
 *
 * Template file: JetLearn - Certificate Tool
 * File ID: 1QWy_mlcsF6K56I357rwyLDEUVzfcYTLH0TyP_gs83VI
 *   Slide 1 (index 0) → Foundation courses
 *   Slide 2 (index 1) → Maths courses
 *   Slide 3 (index 2) → Pro / Advanced courses
 *
 * Placeholders on each slide:
 *   {{learnerName}}  — learner's first name
 *   {{courseName}}   — completed course name
 *   {{year}}         — current year (e.g. 2026)
 * ─────────────────────────────────────────────────────────────────────
 */

var CERT_TEMPLATE_ID  = '1QWy_mlcsF6K56I357rwyLDEUVzfcYTLH0TyP_gs83VI';
var CERT_SAVE_FOLDER  = '1Eaub-wn5J7yMhYrHeQCiHKOXJEKR6vFX';

// ── Course → slide index mapping ─────────────────────────────────────

var CERT_FOUNDATION_COURSES = [
  'introduction to coding (code.org)',
  'animation with scratch jr',
  'introduction to coding ii (code.org)',
  'science adventures with sprite lab',
  'robotics with microbit (jr)',
  'building blocks of ai with google',
  'tynker ai animation lab',
  'learn with minecraft'
];

// Any course containing "maths" or "math" → slide 2
// Everything else not in foundation → slide 3 (Pro/Advanced)

function _getCertSlideIndex(courseName) {
  if (!courseName) return 2; // default to Pro
  var low = courseName.toLowerCase().trim();

  // Maths check
  if (/\bmath/i.test(low)) return 1;

  // Foundation check
  for (var i = 0; i < CERT_FOUNDATION_COURSES.length; i++) {
    if (low.indexOf(CERT_FOUNDATION_COURSES[i]) > -1 ||
        CERT_FOUNDATION_COURSES[i].indexOf(low) > -1) return 0;
  }

  // Default: Pro/Advanced
  return 2;
}

// ── Certificate Log ──────────────────────────────────────────────────
/**
 * Writes one row to the "Certificate Log" tab in JetLearn App Data.
 * Creates the sheet with headers if it doesn't exist yet.
 */
function _logCertificate(jlid, learnerName, courseName, year, parentEmail, sentBy, driveUrl, status, notes) {
  try {
    var HEADERS = [
      'Timestamp', 'JLID', 'Learner Name', 'Course', 'Year',
      'Parent Email', 'Sent By', 'Status', 'Drive URL', 'Notes'
    ];
    var sheet = getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.CERTIFICATE_LOG);

    // Write headers if sheet is brand new (only 1 row max or empty)
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      // Format header row
      var hRange = sheet.getRange(1, 1, 1, HEADERS.length);
      hRange.setFontWeight('bold').setBackground('#6366f1').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 160);  // Timestamp
      sheet.setColumnWidth(2, 140);  // JLID
      sheet.setColumnWidth(3, 160);  // Learner Name
      sheet.setColumnWidth(4, 200);  // Course
      sheet.setColumnWidth(5, 60);   // Year
      sheet.setColumnWidth(6, 220);  // Parent Email
      sheet.setColumnWidth(7, 180);  // Sent By
      sheet.setColumnWidth(8, 80);   // Status
      sheet.setColumnWidth(9, 280);  // Drive URL
      sheet.setColumnWidth(10, 200); // Notes
    }

    sheet.appendRow([
      new Date(),
      jlid        || '',
      learnerName || '',
      courseName  || '',
      year        || '',
      parentEmail || '',
      sentBy      || '',
      status      || 'Sent',
      driveUrl    || '',
      notes       || ''
    ]);
    SpreadsheetApp.flush();
    try { _clearAppDataCache(CONFIG.APP_DATA_SHEETS.CERTIFICATE_LOG); } catch(e) {}
  } catch(e) {
    Logger.log('[Cert] _logCertificate error: ' + e.message);
  }
}

// ── Get Certificate Log (for live log on dashboard) ──────────────────
/**
 * Returns last N rows from Certificate Log sheet.
 * Groups by date so UI can show "Today", "Yesterday", etc.
 */
function getCertificateLog(limit) {
  try {
    limit = limit || 50;
    var sheet = getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.CERTIFICATE_LOG);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, rows: [], todayCount: 0 };

    var startRow = Math.max(2, lastRow - limit + 1);
    var numRows  = lastRow - startRow + 1;
    var data     = sheet.getRange(startRow, 1, numRows, 10).getValues();

    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var todayCount = 0;

    // Reverse so newest first
    var rows = data.reverse().map(function(r) {
      var ts = r[0] ? new Date(r[0]) : null;
      var isToday = ts && ts >= today;
      if (isToday) todayCount++;
      return {
        timestamp:   ts ? Utilities.formatDate(ts, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm') : '',
        jlid:        String(r[1] || ''),
        learnerName: String(r[2] || ''),
        course:      String(r[3] || ''),
        year:        String(r[4] || ''),
        parentEmail: String(r[5] || ''),
        sentBy:      String(r[6] || ''),
        status:      String(r[7] || 'Sent'),
        driveUrl:    String(r[8] || ''),
        notes:       String(r[9] || ''),
        isToday:     isToday
      };
    });

    return { success: true, rows: rows, todayCount: todayCount };
  } catch(e) {
    Logger.log('[Cert] getCertificateLog error: ' + e.message);
    return { success: false, message: e.message, rows: [], todayCount: 0 };
  }
}

// ── Verify placeholders exist on all 3 slides (run once to check) ────
function verifyCertificateTemplate() {
  try {
    var pres   = SlidesApp.openById(CERT_TEMPLATE_ID);
    var slides = pres.getSlides();
    Logger.log('Total slides: ' + slides.length);

    var REQUIRED = ['{{learnerName}}', '{{courseName}}', '{{year}}'];
    var slideLabels = ['Foundation', 'Maths', 'Pro/Advanced'];

    slides.forEach(function(slide, idx) {
      Logger.log('\n── Slide ' + (idx + 1) + ' (' + (slideLabels[idx] || 'Extra') + ') ──');
      var shapes = slide.getShapes();
      var allText = '';
      shapes.forEach(function(shape) {
        try {
          var t = shape.getText().asString();
          if (t.trim()) {
            Logger.log('  Shape: "' + t.trim().substring(0, 80) + '"');
            allText += t;
          }
        } catch(e) {}
      });
      REQUIRED.forEach(function(ph) {
        var found = allText.indexOf(ph) > -1;
        Logger.log('  ' + ph + ' → ' + (found ? '✅ FOUND' : '❌ MISSING'));
      });
    });

    Logger.log('\n✅ Verification complete. Fix any MISSING placeholders before generating certificates.');
  } catch(e) {
    Logger.log('❌ Error: ' + e.message);
    Logger.log('Check that the script has access to the Slides file (share with the script owner email).');
  }
}

// ── Core: generate certificate PDF blob ──────────────────────────────
/**
 * @param {string} learnerName   - Full learner name
 * @param {string} courseName    - Course just completed
 * @param {number|string} [yearOverride] - Optional year (default: current year)
 * @returns {Blob} PDF blob, or null on error
 */
function generateCertificatePDF(learnerName, courseName, yearOverride) {
  var tempPresId = null;
  try {
    var year       = yearOverride ? String(yearOverride) : String(new Date().getFullYear());
    var slideIndex = _getCertSlideIndex(courseName);

    Logger.log('[Cert] Generating certificate: ' + learnerName + ' | ' + courseName + ' | slide ' + (slideIndex + 1));

    // 1. Open source template — get the target slide
    var srcPres  = SlidesApp.openById(CERT_TEMPLATE_ID);
    var srcSlide = srcPres.getSlides()[slideIndex];
    if (!srcSlide) throw new Error('Slide index ' + slideIndex + ' not found in template');

    // 2. Copy the ENTIRE template (preserves page dimensions / custom size)
    //    then remove every slide except the one we need
    //    Drive API can be flaky — retry up to 3 times
    var tempFile = null;
    var copyErr  = null;
    for (var attempt = 0; attempt < 3; attempt++) {
      try {
        if (attempt > 0) Utilities.sleep(2000);
        tempFile = DriveApp.getFileById(CERT_TEMPLATE_ID).makeCopy('_cert_temp_' + Date.now());
        copyErr  = null;
        break;
      } catch(ce) {
        copyErr = ce;
        Logger.log('[Cert] makeCopy attempt ' + (attempt + 1) + ' failed: ' + ce.message);
      }
    }
    if (!tempFile) throw new Error('Drive copy failed after 3 attempts: ' + (copyErr ? copyErr.message : 'unknown'));
    tempPresId   = tempFile.getId();
    var tempPres = SlidesApp.openById(tempPresId);

    // Keep only the target slide — delete all others
    var allSlides = tempPres.getSlides();
    for (var s = allSlides.length - 1; s >= 0; s--) {
      if (s !== slideIndex) allSlides[s].remove();
    }
    var pastedSlide = tempPres.getSlides()[0];

    // 3. Replace placeholders — preserves all existing font/size/color
    pastedSlide.replaceAllText('{{learnerName}}', learnerName);
    pastedSlide.replaceAllText('{{courseName}}',  courseName);
    pastedSlide.replaceAllText('{{year}}',        year);

    // 4. Auto-scale font size to prevent wrapping on long names
    var shapes = pastedSlide.getShapes();
    shapes.forEach(function(shape) {
      try {
        var txt = shape.getText().asString().trim();

        if (txt === learnerName) {
          // Learner name: scale down for long names
          var nameSize = learnerName.length <= 14 ? 36
                       : learnerName.length <= 20 ? 32
                       : learnerName.length <= 26 ? 28
                       : 24;
          shape.getText().getTextStyle().setFontSize(nameSize);
        }

        if (txt === courseName) {
          // Course name: scale down for very long course names
          var courseSize = courseName.length <= 18 ? 28
                         : courseName.length <= 26 ? 24
                         : courseName.length <= 36 ? 20
                         : 16;
          shape.getText().getTextStyle().setFontSize(courseSize);
        }
      } catch(se) {}
    });

    tempPres.saveAndClose();

    // 4. Export as PDF via DriveApp.getAs() — avoids UrlFetch bandwidth quota
    var fileName = learnerName.replace(/\s+/g, '_') + '_' + courseName.replace(/\s+/g, '_') + '_Certificate.pdf';
    var pdfBlob  = DriveApp.getFileById(tempPresId).getAs('application/pdf').setName(fileName);

    // Save to Drive folder
    var savedFile = null;
    var driveUrl  = null;
    try {
      var folder   = DriveApp.getFolderById(CERT_SAVE_FOLDER);
      savedFile    = folder.createFile(pdfBlob);
      driveUrl     = savedFile.getUrl();
      Logger.log('[Cert] Saved to Drive: ' + driveUrl);
    } catch(fe) {
      Logger.log('[Cert] Drive save warning (non-fatal): ' + fe.message);
    }

    Logger.log('[Cert] PDF generated: ' + fileName + ' (' + pdfBlob.getBytes().length + ' bytes)');

    // Return blob + metadata
    pdfBlob._driveUrl  = driveUrl;
    pdfBlob._savedFile = savedFile;
    return pdfBlob;

  } catch(e) {
    Logger.log('[Cert] generateCertificatePDF ERROR for "' + courseName + '": ' + e.message);
    // Re-throw so caller can surface actual error message (not just null)
    throw new Error(e.message);
  } finally {
    // Always clean up temp presentation
    if (tempPresId) {
      try { DriveApp.getFileById(tempPresId).setTrashed(true); } catch(ce) {}
    }
  }
}

// ── Send certificate email to parent ─────────────────────────────────
/**
 * Called from migration flow when CCTC is confirmed.
 * @param {string} jlid
 * @param {string} learnerName
 * @param {string} courseName       - Course being completed (triggers cert)
 * @param {string} parentEmail
 * @param {string} parentName
 * @param {string} performedBy      - CLS user email for audit log
 * @returns {{ success: boolean, message: string }}
 */
function sendCourseCertificateEmail(jlid, learnerName, courseName, parentEmail, parentName, performedBy) {
  try {
    if (!parentEmail) return { success: false, message: 'No parent email.' };

    Logger.log('[Cert] sendCourseCertificateEmail: ' + jlid + ' | ' + courseName + ' → ' + parentEmail);

    // 1. Generate PDF
    var pdfBlob = generateCertificatePDF(learnerName, courseName);
    if (!pdfBlob) return { success: false, message: 'Certificate PDF generation failed.' };

    // 2. Build email
    var year    = new Date().getFullYear();
    var subject = 'JetLearn - ' + learnerName + ' has completed ' + courseName + ' - Certificate Enclosed';

    // ── Course row ────────────────────────────────────────────────────────────
    var courseRowHtml = '<tr>'
      + '<td style="padding-top:13px;padding-bottom:13px;padding-left:20px;padding-right:10px;">'
      +   '<table role="presentation" cellpadding="0" cellspacing="0" border="0">'
      +     '<tr>'
      +       '<td width="24" valign="middle" style="padding-right:10px;">'
      +         '<span style="display:block;width:22px;height:22px;background-color:#ecfdf5;border:1.5px solid #6ee7b7;'
      +               'border-radius:6px;text-align:center;font-size:12px;line-height:20px;color:#059669;font-family:Arial;">&#10003;</span>'
      +       '</td>'
      +       '<td valign="middle" style="color:#1a1560;font-size:13.5px;font-weight:700;font-family:Arial,Helvetica,sans-serif;">'
      +         courseName
      +       '</td>'
      +     '</tr>'
      +   '</table>'
      + '</td>'
      + '<td style="padding-top:13px;padding-bottom:13px;padding-left:10px;padding-right:20px;'
      +       'color:#f97316;font-size:13px;font-weight:800;font-family:Arial,Helvetica,sans-serif;'
      +       'white-space:nowrap;width:60px;">'
      +   year
      + '</td>'
      + '</tr>';

    // ── HTML body (same template as sendBulkCertificates, single-course variant) ──
    var htmlBody =
      '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#eeeef6;padding-top:24px;padding-bottom:24px;">'
    + '<tr><td align="center" valign="top" style="padding-left:12px;padding-right:12px;">'

    // Card
    + '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="max-width:580px;background-color:#ffffff;border-radius:16px;border:1px solid #ddd9f5;">'

    // Header
    + '<tr><td align="center" valign="top" style="background-color:#1a1560;border-radius:16px 16px 0 0;padding-top:36px;padding-bottom:32px;padding-left:40px;padding-right:40px;">'
    +   '<img src="https://cdn.jsdelivr.net/gh/legitshocky/Jet-learn-Images@main/Logo.png" alt="JetLearn" width="140" style="display:block;margin-left:auto;margin-right:auto;margin-bottom:20px;border:0;outline:none;text-decoration:none;max-width:140px;height:auto;">'
    +   '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin-left:auto;margin-right:auto;margin-bottom:16px;">'
    +     '<tr><td style="background-color:#f97316;border-radius:100px;padding-top:5px;padding-bottom:5px;padding-left:16px;padding-right:16px;">'
    +       '<span style="color:#ffffff;font-size:11px;font-weight:700;font-family:Arial,Helvetica,sans-serif;letter-spacing:0.8px;text-transform:uppercase;white-space:nowrap;">&#127881; Certificate of Achievement</span>'
    +     '</td></tr>'
    +   '</table>'
    +   '<h1 style="color:#ffffff;font-size:24px;font-weight:800;line-height:1.3;margin-top:0;margin-bottom:8px;font-family:Arial,Helvetica,sans-serif;text-align:center;">' + learnerName + ' Did It!</h1>'
    +   '<p style="color:#c7d2fe;font-size:13px;font-weight:400;line-height:1.5;margin-top:0;margin-bottom:0;font-family:Arial,Helvetica,sans-serif;text-align:center;">World\'s Top Rated AI, Coding &amp; STEM Academy for Kids</p>'
    + '</td></tr>'

    // Orange bar
    + '<tr><td style="height:4px;background-color:#f97316;font-size:0;line-height:0;">&nbsp;</td></tr>'

    // Hero banner
    + '<tr><td style="background-color:#fff7ed;border-bottom:1px solid #fed7aa;padding-top:20px;padding-bottom:20px;padding-left:32px;padding-right:32px;">'
    +   '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr>'
    +     '<td width="52" valign="middle" style="font-size:40px;line-height:1;padding-right:14px;">&#127942;</td>'
    +     '<td valign="middle">'
    +       '<p style="color:#f97316;font-size:14px;margin-top:0;margin-bottom:4px;font-family:Arial,Helvetica,sans-serif;">&#9733;&#9733;&#9733;&#9733;&#9733;</p>'
    +       '<p style="color:#92400e;font-size:13px;font-weight:600;line-height:1.6;margin-top:0;margin-bottom:0;font-family:Arial,Helvetica,sans-serif;"><span style="color:#7c2d12;font-weight:700;">' + learnerName + '</span> has earned a certificate!</p>'
    +     '</td>'
    +   '</tr></table>'
    + '</td></tr>'

    // Body
    + '<tr><td style="background-color:#ffffff;padding-top:36px;padding-bottom:36px;padding-left:40px;padding-right:40px;">'
    +   '<p style="color:#111827;font-size:15px;font-weight:700;line-height:1.5;margin-top:0;margin-bottom:14px;font-family:Arial,Helvetica,sans-serif;">Dear ' + (parentName || 'Parent') + ',</p>'
    +   '<p style="color:#4b5563;font-size:14px;font-weight:400;line-height:1.85;margin-top:0;margin-bottom:24px;font-family:Arial,Helvetica,sans-serif;">We are absolutely thrilled to share that <span style="color:#1a1560;font-weight:700;">' + learnerName + '</span> has successfully completed the <span style="color:#1a1560;font-weight:700;">' + courseName + '</span> course at JetLearn. The certificate is attached &#8212; a milestone worth celebrating!</p>'

    // Cert table card
    +   '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-radius:14px;border:1.5px solid #ddd9f5;margin-bottom:24px;">'
    +     '<tr><td colspan="2" style="background-color:#1a1560;border-radius:14px 14px 0 0;padding-top:13px;padding-bottom:13px;padding-left:20px;padding-right:20px;">'
    +       '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr>'
    +         '<td style="color:#c7d2fe;font-size:11px;font-weight:700;font-family:Arial,Helvetica,sans-serif;letter-spacing:0.7px;text-transform:uppercase;">Certificate Earned</td>'
    +         '<td align="right"><span style="background-color:#f97316;color:#ffffff;font-size:11px;font-weight:700;font-family:Arial,Helvetica,sans-serif;padding-top:3px;padding-bottom:3px;padding-left:12px;padding-right:12px;border-radius:100px;">1 Course</span></td>'
    +       '</tr></table>'
    +     '</td></tr>'
    +     '<tr style="background-color:#f8f7ff;">'
    +       '<td style="padding-top:9px;padding-bottom:9px;padding-left:20px;padding-right:10px;font-size:11px;font-weight:700;color:#9ca3af;font-family:Arial,Helvetica,sans-serif;letter-spacing:0.7px;text-transform:uppercase;border-bottom:1px solid #f0eeff;">COURSE</td>'
    +       '<td style="padding-top:9px;padding-bottom:9px;padding-left:10px;padding-right:20px;font-size:11px;font-weight:700;color:#9ca3af;font-family:Arial,Helvetica,sans-serif;letter-spacing:0.7px;text-transform:uppercase;border-bottom:1px solid #f0eeff;white-space:nowrap;width:60px;">YEAR</td>'
    +     '</tr>'
    +     courseRowHtml
    +   '</table>'

    // Callout
    +   '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#fff7ed;border:1.5px solid #fed7aa;border-radius:14px;margin-bottom:24px;">'
    +     '<tr><td style="padding-top:18px;padding-bottom:18px;padding-left:20px;padding-right:20px;">'
    +       '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr>'
    +         '<td width="42" valign="top" style="padding-right:14px;">'
    +           '<span style="display:block;width:38px;height:38px;background-color:#f97316;border-radius:10px;text-align:center;font-size:20px;line-height:38px;">&#11088;</span>'
    +         '</td>'
    +         '<td valign="middle" style="color:#92400e;font-size:13.5px;font-weight:500;line-height:1.75;font-family:Arial,Helvetica,sans-serif;">'
    +           'We encourage you to save and share <span style="color:#7c2d12;font-weight:700;">' + learnerName + '\'s certificate</span> &#8212; it is proof of curiosity, persistence, and a love for learning that will carry them far.'
    +         '</td>'
    +       '</tr></table>'
    +     '</td></tr>'
    +   '</table>'

    // Closing
    +   '<p style="color:#4b5563;font-size:14px;font-weight:400;line-height:1.85;margin-top:0;margin-bottom:28px;font-family:Arial,Helvetica,sans-serif;">Thank you for trusting JetLearn with <span style="color:#1a1560;font-weight:700;">' + learnerName + '\'s</span> education. We are excited to see what they build next!</p>'

    // Divider
    +   '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:20px;"><tr><td style="height:1px;background-color:#f0eeff;font-size:0;line-height:0;">&nbsp;</td></tr></table>'

    // Sign-off
    +   '<p style="color:#6b7280;font-size:14px;font-weight:400;margin-top:0;margin-bottom:4px;font-family:Arial,Helvetica,sans-serif;">Warm regards,</p>'
    +   '<p style="color:#1a1560;font-size:18px;font-weight:800;margin-top:0;margin-bottom:4px;font-family:Arial,Helvetica,sans-serif;">Team JetLearn &#9889;</p>'
    +   '<p style="color:#9ca3af;font-size:12px;font-weight:400;margin-top:0;margin-bottom:0;font-family:Arial,Helvetica,sans-serif;">World\'s Top Rated AI, Coding &amp; STEM Academy for Kids</p>'
    + '</td></tr>'

    // Footer
    + '<tr><td style="background-color:#f5f3ff;border-top:1px solid #ddd9f5;border-radius:0 0 16px 16px;padding-top:16px;padding-bottom:16px;padding-left:40px;padding-right:40px;text-align:center;">'
    +   '<p style="color:#9ca3af;font-size:12px;font-weight:400;margin-top:0;margin-bottom:0;font-family:Arial,Helvetica,sans-serif;line-height:1.6;">'
    +     'JetLearn &nbsp;&middot;&nbsp; hello@jet-learn.com &nbsp;&middot;&nbsp; '
    +     '<a href="https://jet-learn.com" style="color:#6366f1;text-decoration:none;font-weight:600;">jet-learn.com</a>'
    +   '</p>'
    + '</td></tr>'

    + '</table>'
    + '</td></tr></table>';

    // 3. Send
    GmailApp.sendEmail(parentEmail, subject, '', {
      htmlBody   : htmlBody,
      attachments: [pdfBlob],
      name       : 'JetLearn',
      replyTo    : 'hello@jet-learn.com'
    });

    // 4. Audit log
    try {
      logAction(
        'Certificate Sent',
        jlid,
        learnerName + ' completed ' + courseName + '. Certificate emailed to ' + parentEmail,
        '',
        performedBy || ''
      );
    } catch(ae) {}

    // Log to Certificate Log sheet
    _logCertificate(
      jlid, learnerName, courseName, new Date().getFullYear(),
      parentEmail, performedBy,
      pdfBlob._driveUrl || '',
      'Sent', 'Single certificate — CCTC migration'
    );

    // Also log to Audit Log
    try {
      logAction('Certificate Sent', jlid,
        learnerName, '', '', courseName, 'Success',
        'Certificate emailed to ' + parentEmail, 'Course Completion', performedBy);
    } catch(ae) {}

    Logger.log('[Cert] Certificate sent to ' + parentEmail);

    // 5. HubSpot: note on deal + tick checkbox — via background trigger (avoids timeout)
    try {
      if (jlid) {
        var sc2 = CacheService.getScriptCache();
        var singlePayload = { jlid: jlid, learnerName: learnerName, parentEmail: parentEmail,
          performedBy: performedBy || '',
          courses: [{ name: courseName, year: new Date().getFullYear() }], failed: [] };
        sc2.put('CERT_HS_PENDING_' + jlid, JSON.stringify(singlePayload), 600);
        var qRaw2 = sc2.get('CERT_HS_QUEUE');
        var q2 = []; try { if (qRaw2) q2 = JSON.parse(qRaw2); } catch(e) {}
        if (q2.indexOf(jlid) === -1) q2.push(jlid);
        sc2.put('CERT_HS_QUEUE', JSON.stringify(q2), 600);
        ScriptApp.getProjectTriggers().forEach(function(t) {
          if (t.getHandlerFunction() === '_processPendingCertHS') ScriptApp.deleteTrigger(t);
        });
        ScriptApp.newTrigger('_processPendingCertHS').timeBased().after(5000).create();
      }
    } catch(hse) {
      Logger.log('[Cert] Single cert HS trigger (non-fatal): ' + hse.message);
    }

    return {
      success  : true,
      message  : 'Certificate sent to ' + parentEmail,
      driveUrl : pdfBlob._driveUrl  || null,
      driveFile: pdfBlob._savedFile || null
    };

  } catch(e) {
    Logger.log('[Cert] sendCourseCertificateEmail error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── HubSpot: create a note engagement on a deal ──────────────────────
// Uses legacy engagements v1 API — simpler, reliable, no association type guessing.
function _certCreateDealNote(dealId, noteBody, token) {
  try {
    // CRM v3 Notes API — associates directly with deal only (not tickets)
    var payload = {
      properties: {
        hs_note_body : noteBody,
        hs_timestamp : String(new Date().getTime())
      },
      associations: [{
        to   : { id: Number(dealId) },
        types: [{ associationCategory: 'HUBSPOT_DEFINED', associationTypeId: 214 }]
      }]
    };
    var resp = UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/objects/notes', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    Logger.log('[Cert] Deal note (v3) → ' + resp.getResponseCode() + ' ' + resp.getContentText().substring(0, 200));
  } catch(e) {
    Logger.log('[Cert] _certCreateDealNote error: ' + e.message);
  }
}

// ── HubSpot: tick course_completion_certificates_sent_to_the_parent ──
function _certTickCourseSent(dealId, courseName, token) {
  try {
    var PROP = 'course_completion_certificates_sent_to_the_parent';

    // 1. Fetch current ticked values
    var getResp = UrlFetchApp.fetch(
      'https://api.hubapi.com/crm/v3/objects/deals/' + dealId + '?properties=' + PROP,
      { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
    );
    var currentVal = '';
    if (getResp.getResponseCode() === 200) {
      try { currentVal = JSON.parse(getResp.getContentText()).properties[PROP] || ''; } catch(pe) {}
    }

    // 2. Find matching option value from HubSpot property definition
    var optionVal = courseName; // fallback: raw name
    try {
      var propResp = UrlFetchApp.fetch(
        'https://api.hubapi.com/crm/v3/properties/deals/' + PROP,
        { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
      );
      if (propResp.getResponseCode() === 200) {
        var options = JSON.parse(propResp.getContentText()).options || [];
        var norm = function(s){ return String(s).toLowerCase().replace(/[^a-z0-9]/g, ''); };
        var cn = norm(courseName);
        // Exact value match → label match → substring match
        var matched = '';
        for (var i = 0; i < options.length; i++) {
          if (norm(options[i].value) === cn) { matched = options[i].value; break; }
        }
        if (!matched) {
          for (var i = 0; i < options.length; i++) {
            if (norm(options[i].label) === cn) { matched = options[i].value; break; }
          }
        }
        if (!matched) {
          var bestLen = 0;
          options.forEach(function(o) {
            var ol = norm(o.label);
            if ((cn.indexOf(ol) > -1 || ol.indexOf(cn) > -1) && ol.length > bestLen) {
              matched = o.value; bestLen = ol.length;
            }
          });
        }
        if (matched) optionVal = matched;
      }
    } catch(oe) { Logger.log('[Cert] option fetch warn: ' + oe.message); }

    // 3. Append to existing semicolon-separated list (no duplicates)
    var existing = currentVal ? currentVal.split(';').map(function(s){ return s.trim(); }).filter(Boolean) : [];
    if (existing.indexOf(optionVal) === -1) existing.push(optionVal);
    var newVal = existing.join(';');

    // 4. PATCH deal
    var patchResp = UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/objects/deals/' + dealId, {
      method: 'patch',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ properties: { [PROP]: newVal } }),
      muteHttpExceptions: true
    });
    Logger.log('[Cert] Tick cert checkbox → ' + patchResp.getResponseCode() + ' val=' + newVal);
  } catch(e) {
    Logger.log('[Cert] _certTickCourseSent error: ' + e.message);
  }
}

// ── Get completed course history for a learner (from CPRS) ───────────
/**
 * Reads Athena CPRS sheet and returns distinct courses the learner has
 * sessions in, sorted by first session date. Each entry includes the
 * inferred completion year so CLS can override before bulk-sending.
 *
 * @param {string} jlid
 * @returns {{ success:boolean, courses:[{course,firstDate,lastDate,year,sessionsDone}], learnerName:string }}
 */
function getLearnerCourseHistory(query) {
  try {
    if (!query) return { success: false, message: 'JLID or name required.' };
    var cprsMap = _buildCPRSMap();

    // Find by JLID or by name (partial match)
    var jlid = query.trim();
    var rows = cprsMap[jlid] || [];

    if (!rows.length) {
      // Try name search — find JLID whose learner name contains query
      var qLow = query.toLowerCase();
      var prmsMap = _buildPRMSMap();
      Object.keys(prmsMap).forEach(function(k) {
        if (!rows.length) {
          var nm = (prmsMap[k].learnerName || '').toLowerCase();
          if (nm.indexOf(qLow) > -1) { jlid = k; rows = cprsMap[k] || []; }
        }
      });
    }
    if (!rows.length) return { success: false, message: 'No CPRS data found for ' + jlid };

    // Group by course name
    var courseMap = {};
    rows.forEach(function(r) {
      if (!r.course) return;
      var key = r.course.trim();
      if (!courseMap[key]) courseMap[key] = { sessions: [], happened: 0 };
      if (r.date) courseMap[key].sessions.push(r.date);
      if (r.happened) courseMap[key].happened++;
    });

    // Build sorted list
    var list = Object.keys(courseMap).map(function(courseName) {
      var dates = courseMap[courseName].sessions
        .filter(function(d){ return d instanceof Date && !isNaN(d); })
        .sort(function(a,b){ return a - b; });
      var firstDate = dates.length ? dates[0]           : null;
      var lastDate  = dates.length ? dates[dates.length-1] : null;
      var year      = lastDate ? lastDate.getFullYear() : new Date().getFullYear();
      return {
        course      : courseName,
        firstDate   : firstDate ? firstDate.toISOString().split('T')[0] : '',
        lastDate    : lastDate  ? lastDate.toISOString().split('T')[0]  : '',
        year        : year,
        sessionsDone: courseMap[courseName].happened
      };
    }).filter(function(c){ return c.sessionsDone > 0; })
      .sort(function(a,b){ return (a.firstDate || '') < (b.firstDate || '') ? -1 : 1; });

    // Learner name from first row
    var learnerName = '';
    try {
      var prmsMap = _buildPRMSMap();
      if (prmsMap[jlid]) learnerName = prmsMap[jlid].learnerName || '';
    } catch(e) {}

    return { success: true, courses: list, learnerName: learnerName, jlid: jlid };
  } catch(e) {
    Logger.log('[Cert] getLearnerCourseHistory error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Pool-based bulk PDF generation ───────────────────────────────────
/**
 * Generates N certificate PDFs using 3 shared pool copies (one per slide type).
 * Only 3 makeCopy calls regardless of how many certs — eliminates Drive rate limits.
 *
 * @param {string} learnerName
 * @param {Array<{name:string, year:number}>} courses
 * @returns {Array<{name, year, blob, driveUrl, error}>}
 */
function _generateBulkCertPDFs(learnerName, courses) {
  var SLIDE_NAMES = ['Foundation', 'Maths', 'Pro'];
  var poolIds     = {};   // slideIndex → fileId
  var origSizes   = {};   // 'si:placeholder' → fontSize (for reset)

  try {
    // ── Step 1: Read original font sizes from template (once) ──────────
    try {
      var tmpl = SlidesApp.openById(CERT_TEMPLATE_ID);
      [0, 1, 2].forEach(function(si) {
        var slide = tmpl.getSlides()[si];
        if (!slide) return;
        slide.getShapes().forEach(function(shape) {
          try {
            var txt = shape.getText().asString().trim();
            if (txt === '{{learnerName}}' || txt === '{{courseName}}' || txt === '{{year}}') {
              var sz = shape.getText().getTextStyle().getFontSize();
              if (sz) origSizes[si + ':' + txt] = sz;
            }
          } catch(e) {}
        });
      });
    } catch(re) {
      Logger.log('[Cert] origSizes read warn: ' + re.message);
    }

    // ── Step 2: Make 3 slim pool copies (one per slide type), with retry ─
    [0, 1, 2].forEach(function(si) {
      for (var a = 0; a < 4; a++) {
        try {
          if (a > 0) Utilities.sleep(3000);
          var copy  = DriveApp.getFileById(CERT_TEMPLATE_ID).makeCopy('_cert_pool_' + SLIDE_NAMES[si]);
          var pres  = SlidesApp.openById(copy.getId());
          var slides = pres.getSlides();
          for (var s = slides.length - 1; s >= 0; s--) {
            if (s !== si) slides[s].remove();
          }
          pres.saveAndClose();
          poolIds[si] = copy.getId();
          Logger.log('[Cert] Pool ready: ' + SLIDE_NAMES[si] + ' (' + copy.getId() + ')');
          break;
        } catch(ce) {
          Logger.log('[Cert] Pool ' + SLIDE_NAMES[si] + ' attempt ' + (a + 1) + ': ' + ce.message);
          if (a === 3) Logger.log('[Cert] Pool ' + SLIDE_NAMES[si] + ' UNAVAILABLE after 4 attempts');
        }
      }
    });

    // ── Step 3: Generate each cert using its pool copy ──────────────────
    var results = [];

    courses.forEach(function(c, idx) {
      if (idx > 0) Utilities.sleep(1500);
      var result = { name: c.name, year: c.year, blob: null, driveUrl: null, error: null };

      try {
        var si     = _getCertSlideIndex(c.name);
        var poolId = poolIds[si];
        if (!poolId) throw new Error('Pool unavailable for slide type ' + SLIDE_NAMES[si]);

        var year     = c.year ? String(c.year) : String(new Date().getFullYear());
        var fileName = learnerName.replace(/\s+/g,'_') + '_' + c.name.replace(/\s+/g,'_') + '_Certificate.pdf';
        Logger.log('[Cert] → ' + c.name + ' (' + year + ') pool=' + SLIDE_NAMES[si]);

        // Open pool, fill
        var pres  = SlidesApp.openById(poolId);
        var slide = pres.getSlides()[0];
        slide.replaceAllText('{{learnerName}}', learnerName);
        slide.replaceAllText('{{courseName}}',  c.name);
        slide.replaceAllText('{{year}}',        year);

        // Font scaling
        slide.getShapes().forEach(function(shape) {
          try {
            var txt = shape.getText().asString().trim();
            if (txt === learnerName) {
              var ns = learnerName.length <= 14 ? 36 : learnerName.length <= 20 ? 32 :
                       learnerName.length <= 26 ? 28 : 24;
              shape.getText().getTextStyle().setFontSize(ns);
            }
            if (txt === c.name) {
              var cs = c.name.length <= 18 ? 28 : c.name.length <= 26 ? 24 :
                       c.name.length <= 36 ? 20 : 16;
              shape.getText().getTextStyle().setFontSize(cs);
            }
          } catch(se) {}
        });
        pres.saveAndClose();

        // Export PDF — retry up to 3 times
        var blob = null;
        for (var pa = 0; pa < 3; pa++) {
          try {
            if (pa > 0) Utilities.sleep(3000);
            blob = DriveApp.getFileById(poolId).getAs('application/pdf').setName(fileName);
            break;
          } catch(pe) {
            Logger.log('[Cert] getAs attempt ' + (pa+1) + ' "' + c.name + '": ' + pe.message);
            if (pa === 2) throw pe;
          }
        }

        // Save to Drive folder
        try {
          var saved = DriveApp.getFolderById(CERT_SAVE_FOLDER).createFile(blob);
          result.driveUrl  = saved.getUrl();
          blob._driveUrl   = result.driveUrl;
        } catch(fe) {
          Logger.log('[Cert] Drive folder save (non-fatal): ' + fe.message);
        }

        result.blob = blob;
        Logger.log('[Cert] ✅ ' + c.name);

        // Reset pool slide for next reuse
        var presR  = SlidesApp.openById(poolId);
        var slideR = presR.getSlides()[0];
        slideR.replaceAllText(learnerName, '{{learnerName}}');
        slideR.replaceAllText(c.name,      '{{courseName}}');
        slideR.replaceAllText(year,        '{{year}}');
        // Restore original font sizes
        slideR.getShapes().forEach(function(shape) {
          try {
            var txt = shape.getText().asString().trim();
            var origSz = origSizes[si + ':' + txt];
            if (origSz) shape.getText().getTextStyle().setFontSize(origSz);
          } catch(se) {}
        });
        presR.saveAndClose();

      } catch(e) {
        result.error = e.message;
        Logger.log('[Cert] ❌ "' + c.name + '": ' + e.message);
      }

      results.push(result);
    });

    return results;

  } finally {
    // Always trash pool copies
    Object.keys(poolIds).forEach(function(si) {
      try { DriveApp.getFileById(poolIds[si]).setTrashed(true); } catch(e) {}
    });
    Logger.log('[Cert] Pool copies cleaned up');
  }
}

// ── Send multiple certificates in one email ───────────────────────────
/**
 * @param {object} data
 *   jlid, learnerName, parentEmail, parentName, performedBy
 *   courses: [{ name:string, year:number }]
 * @returns {{ success:boolean, sent:number, failed:string[], message:string }}
 */
function sendBulkCertificates(data) {
  try {
    var jlid        = data.jlid        || '';
    var learnerName = data.learnerName  || '';
    var parentEmail = data.parentEmail  || '';
    var parentName  = data.parentName   || 'Parent';
    var performedBy = data.performedBy  || '';
    var courses     = data.courses      || [];

    if (!parentEmail) return { success: false, message: 'No parent email.' };
    if (!courses.length) return { success: false, message: 'No courses selected.' };

    // ── Generate all PDFs via pool approach (3 makeCopy calls max) ─────
    var results  = _generateBulkCertPDFs(learnerName, courses);
    var blobs    = [];
    var failed   = [];
    var sentList = [];

    results.forEach(function(r) {
      if (r.blob) {
        blobs.push(r.blob);
        sentList.push(r.name + ' (' + r.year + ')');
      } else {
        failed.push(r.name + (r.error ? ' [' + r.error + ']' : ''));
      }
    });

    if (!blobs.length) return { success: false, message: 'All PDF generations failed.', failed: failed };

    var isBulk  = courses.length > 1;
    var subject = isBulk
      ? 'JetLearn - ' + learnerName + ' - Course Completion Certificates (' + courses.length + ' courses)'
      : 'JetLearn - ' + learnerName + ' has completed ' + courses[0].name + ' - Certificate Enclosed';

    // ── Course rows (table-safe, fully inline) ────────────────────────────────
    var courseRowsHtml = sentList.map(function(n, i) {
      var courseName   = n.split(' (')[0];
      var year         = n.match(/\((\d+)\)/)[1];
      var isLast       = i === sentList.length - 1;
      var rowBorder    = isLast ? '' : 'border-bottom:1px solid #f8f7ff;';
      return '<tr>'
        + '<td style="padding-top:13px;padding-bottom:13px;padding-left:20px;padding-right:10px;' + rowBorder + '">'
        +   '<table role="presentation" cellpadding="0" cellspacing="0" border="0">'
        +     '<tr>'
        +       '<td width="24" valign="middle" style="padding-right:10px;">'
        +         '<span style="display:block;width:22px;height:22px;background-color:#ecfdf5;border:1.5px solid #6ee7b7;'
        +               'border-radius:6px;text-align:center;font-size:12px;line-height:20px;color:#059669;font-family:Arial;">&#10003;</span>'
        +       '</td>'
        +       '<td valign="middle" style="color:#1a1560;font-size:13.5px;font-weight:700;font-family:Arial,Helvetica,sans-serif;">'
        +         courseName
        +       '</td>'
        +     '</tr>'
        +   '</table>'
        + '</td>'
        + '<td style="padding-top:13px;padding-bottom:13px;padding-left:10px;padding-right:20px;'
        +       'color:#f97316;font-size:13px;font-weight:800;font-family:Arial,Helvetica,sans-serif;'
        +       'white-space:nowrap;width:60px;' + rowBorder + '">'
        +   year
        + '</td>'
        + '</tr>';
    }).join('');

    // ── HTML body ─────────────────────────────────────────────────────────────
    var certCountLabel = isBulk ? courses.length + ' Courses' : '1 Course';
    var heroText = isBulk
      ? 'A big achievement deserves a big celebration.<br><span style="color:#7c2d12;font-weight:700;">' + learnerName + '</span> has earned ' + courses.length + ' certificates!'
      : '<span style="color:#7c2d12;font-weight:700;">' + learnerName + '</span> has earned a certificate!';
    var introText = isBulk
      ? 'We are absolutely thrilled to share that <span style="color:#1a1560;font-weight:700;">' + learnerName + '</span> has successfully completed the following courses at JetLearn. Each certificate is attached as a separate PDF &#8212; a milestone worth celebrating!'
      : 'We are absolutely thrilled to share that <span style="color:#1a1560;font-weight:700;">' + learnerName + '</span> has successfully completed the <span style="color:#1a1560;font-weight:700;">' + courses[0].name + '</span> course at JetLearn. The certificate is attached &#8212; a milestone worth celebrating!';

    var htmlBody =
      '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#eeeef6;padding-top:24px;padding-bottom:24px;">'
    + '<tr><td align="center" valign="top" style="padding-left:12px;padding-right:12px;">'

    // Card
    + '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="max-width:580px;background-color:#ffffff;border-radius:16px;border:1px solid #ddd9f5;">'

    // Header
    + '<tr><td align="center" valign="top" style="background-color:#1a1560;border-radius:16px 16px 0 0;padding-top:36px;padding-bottom:32px;padding-left:40px;padding-right:40px;">'
    +   '<img src="https://cdn.jsdelivr.net/gh/legitshocky/Jet-learn-Images@main/Logo.png" alt="JetLearn" width="140" style="display:block;margin-left:auto;margin-right:auto;margin-bottom:20px;border:0;outline:none;text-decoration:none;max-width:140px;height:auto;">'
    +   '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin-left:auto;margin-right:auto;margin-bottom:16px;">'
    +     '<tr><td style="background-color:#f97316;border-radius:100px;padding-top:5px;padding-bottom:5px;padding-left:16px;padding-right:16px;">'
    +       '<span style="color:#ffffff;font-size:11px;font-weight:700;font-family:Arial,Helvetica,sans-serif;letter-spacing:0.8px;text-transform:uppercase;white-space:nowrap;">&#127881; Certificate of Achievement</span>'
    +     '</td></tr>'
    +   '</table>'
    +   '<h1 style="color:#ffffff;font-size:24px;font-weight:800;line-height:1.3;margin-top:0;margin-bottom:8px;font-family:Arial,Helvetica,sans-serif;text-align:center;">'
    +     (isBulk ? 'Celebrating ' + learnerName + '\'s Success!' : learnerName + ' Did It!')
    +   '</h1>'
    +   '<p style="color:#c7d2fe;font-size:13px;font-weight:400;line-height:1.5;margin-top:0;margin-bottom:0;font-family:Arial,Helvetica,sans-serif;text-align:center;">World\'s Top Rated AI, Coding &amp; STEM Academy for Kids</p>'
    + '</td></tr>'

    // Orange bar
    + '<tr><td style="height:4px;background-color:#f97316;font-size:0;line-height:0;">&nbsp;</td></tr>'

    // Hero banner
    + '<tr><td style="background-color:#fff7ed;border-bottom:1px solid #fed7aa;padding-top:20px;padding-bottom:20px;padding-left:32px;padding-right:32px;">'
    +   '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr>'
    +     '<td width="52" valign="middle" style="font-size:40px;line-height:1;padding-right:14px;">&#127942;</td>'
    +     '<td valign="middle">'
    +       '<p style="color:#f97316;font-size:14px;margin-top:0;margin-bottom:4px;font-family:Arial,Helvetica,sans-serif;">&#9733;&#9733;&#9733;&#9733;&#9733;</p>'
    +       '<p style="color:#92400e;font-size:13px;font-weight:600;line-height:1.6;margin-top:0;margin-bottom:0;font-family:Arial,Helvetica,sans-serif;">' + heroText + '</p>'
    +     '</td>'
    +   '</tr></table>'
    + '</td></tr>'

    // Body
    + '<tr><td style="background-color:#ffffff;padding-top:36px;padding-bottom:36px;padding-left:40px;padding-right:40px;">'
    +   '<p style="color:#111827;font-size:15px;font-weight:700;line-height:1.5;margin-top:0;margin-bottom:14px;font-family:Arial,Helvetica,sans-serif;">Dear ' + (parentName || 'Parent') + ',</p>'
    +   '<p style="color:#4b5563;font-size:14px;font-weight:400;line-height:1.85;margin-top:0;margin-bottom:24px;font-family:Arial,Helvetica,sans-serif;">' + introText + '</p>'

    // Cert table card
    +   '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-radius:14px;border:1.5px solid #ddd9f5;margin-bottom:24px;">'
    +     '<tr><td colspan="2" style="background-color:#1a1560;border-radius:14px 14px 0 0;padding-top:13px;padding-bottom:13px;padding-left:20px;padding-right:20px;">'
    +       '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr>'
    +         '<td style="color:#c7d2fe;font-size:11px;font-weight:700;font-family:Arial,Helvetica,sans-serif;letter-spacing:0.7px;text-transform:uppercase;">Certificates Earned</td>'
    +         '<td align="right"><span style="background-color:#f97316;color:#ffffff;font-size:11px;font-weight:700;font-family:Arial,Helvetica,sans-serif;padding-top:3px;padding-bottom:3px;padding-left:12px;padding-right:12px;border-radius:100px;">' + certCountLabel + '</span></td>'
    +       '</tr></table>'
    +     '</td></tr>'
    +     '<tr style="background-color:#f8f7ff;">'
    +       '<td style="padding-top:9px;padding-bottom:9px;padding-left:20px;padding-right:10px;font-size:11px;font-weight:700;color:#9ca3af;font-family:Arial,Helvetica,sans-serif;letter-spacing:0.7px;text-transform:uppercase;border-bottom:1px solid #f0eeff;">COURSE</td>'
    +       '<td style="padding-top:9px;padding-bottom:9px;padding-left:10px;padding-right:20px;font-size:11px;font-weight:700;color:#9ca3af;font-family:Arial,Helvetica,sans-serif;letter-spacing:0.7px;text-transform:uppercase;border-bottom:1px solid #f0eeff;white-space:nowrap;width:60px;">YEAR</td>'
    +     '</tr>'
    +     courseRowsHtml
    +   '</table>'

    // Callout
    +   '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#fff7ed;border:1.5px solid #fed7aa;border-radius:14px;margin-bottom:24px;">'
    +     '<tr><td style="padding-top:18px;padding-bottom:18px;padding-left:20px;padding-right:20px;">'
    +       '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr>'
    +         '<td width="42" valign="top" style="padding-right:14px;">'
    +           '<span style="display:block;width:38px;height:38px;background-color:#f97316;border-radius:10px;text-align:center;font-size:20px;line-height:38px;">&#11088;</span>'
    +         '</td>'
    +         '<td valign="middle" style="color:#92400e;font-size:13.5px;font-weight:500;line-height:1.75;font-family:Arial,Helvetica,sans-serif;">'
    +           'We encourage you to save and share <span style="color:#7c2d12;font-weight:700;">' + learnerName + '\'s certificates</span> &#8212; they are proof of curiosity, persistence, and a love for learning that will carry them far.'
    +         '</td>'
    +       '</tr></table>'
    +     '</td></tr>'
    +   '</table>'

    // Closing
    +   '<p style="color:#4b5563;font-size:14px;font-weight:400;line-height:1.85;margin-top:0;margin-bottom:28px;font-family:Arial,Helvetica,sans-serif;">Thank you for trusting JetLearn with <span style="color:#1a1560;font-weight:700;">' + learnerName + '\'s</span> education. We are excited to see what they build next!</p>'

    // Divider
    +   '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:20px;"><tr><td style="height:1px;background-color:#f0eeff;font-size:0;line-height:0;">&nbsp;</td></tr></table>'

    // Sign-off
    +   '<p style="color:#6b7280;font-size:14px;font-weight:400;margin-top:0;margin-bottom:4px;font-family:Arial,Helvetica,sans-serif;">Warm regards,</p>'
    +   '<p style="color:#1a1560;font-size:18px;font-weight:800;margin-top:0;margin-bottom:4px;font-family:Arial,Helvetica,sans-serif;">Team JetLearn &#9889;</p>'
    +   '<p style="color:#9ca3af;font-size:12px;font-weight:400;margin-top:0;margin-bottom:0;font-family:Arial,Helvetica,sans-serif;">World\'s Top Rated AI, Coding &amp; STEM Academy for Kids</p>'
    + '</td></tr>'

    // Footer
    + '<tr><td style="background-color:#f5f3ff;border-top:1px solid #ddd9f5;border-radius:0 0 16px 16px;padding-top:16px;padding-bottom:16px;padding-left:40px;padding-right:40px;text-align:center;">'
    +   '<p style="color:#9ca3af;font-size:12px;font-weight:400;margin-top:0;margin-bottom:0;font-family:Arial,Helvetica,sans-serif;line-height:1.6;">'
    +     'JetLearn &nbsp;&middot;&nbsp; hello@jet-learn.com &nbsp;&middot;&nbsp; '
    +     '<a href="https://jet-learn.com" style="color:#6366f1;text-decoration:none;font-weight:600;">jet-learn.com</a>'
    +   '</p>'
    + '</td></tr>'

    + '</table>'
    + '</td></tr></table>';

    // ── Send ──────────────────────────────────────────────────────────────────
    GmailApp.sendEmail(parentEmail, subject, '', {
      htmlBody   : htmlBody,
      attachments: blobs,
      name       : 'JetLearn',
      replyTo    : 'hello@jet-learn.com'
    });

    // ── Log each cert (use results array for accurate status + driveUrl) ─────
    results.forEach(function(r) {
      _logCertificate(
        jlid, learnerName, r.name, r.year,
        parentEmail, performedBy,
        r.driveUrl || '',
        r.blob ? 'Sent' : 'Failed',
        r.blob ? ('Bulk send — ' + blobs.length + ' certificate(s)') : ('Failed: ' + (r.error || 'unknown'))
      );
    });

    // ── Audit log ─────────────────────────────────────────────────────────────
    try {
      logAction('Bulk Certificates Sent', jlid, learnerName, '', '',
        sentList.join(', '), 'Success',
        blobs.length + ' cert(s) emailed to ' + parentEmail,
        'Bulk Certificate Request', performedBy);
    } catch(ae) {}

    Logger.log('[Cert] Bulk: ' + blobs.length + ' certs sent to ' + parentEmail);

    // ── HubSpot: note + checkbox tick — fire via 1-min trigger to avoid timeout ──
    try {
      if (jlid) {
        var sc3 = CacheService.getScriptCache();
        // Strip error detail from failed names before storing in HS payload
        var failedNames = failed.map(function(f){ return f.replace(/\s*\[.*\]$/, ''); });
        var certPayload = { jlid: jlid, learnerName: learnerName, parentEmail: parentEmail,
                            performedBy: performedBy || '', courses: courses, failed: failedNames };
        sc3.put('CERT_HS_PENDING_' + jlid, JSON.stringify(certPayload), 600);
        // Queue: append jlid so trigger knows which keys to process
        var qRaw = sc3.get('CERT_HS_QUEUE');
        var q = []; try { if (qRaw) q = JSON.parse(qRaw); } catch(e) {}
        if (q.indexOf(jlid) === -1) q.push(jlid);
        sc3.put('CERT_HS_QUEUE', JSON.stringify(q), 600);
        ScriptApp.getProjectTriggers().forEach(function(t) {
          if (t.getHandlerFunction() === '_processPendingCertHS') ScriptApp.deleteTrigger(t);
        });
        ScriptApp.newTrigger('_processPendingCertHS').timeBased().after(5000).create();
      }
    } catch(hse2) {
      Logger.log('[Cert] Bulk HS trigger setup (non-fatal): ' + hse2.message);
    }

    return {
      success : true,
      sent    : blobs.length,
      failed  : failed,
      message : blobs.length + ' certificate(s) sent to ' + parentEmail
        + (failed.length ? '. Failed: ' + failed.join(', ') : '')
    };

  } catch(e) {
    Logger.log('[Cert] sendBulkCertificates error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Resend failed certificates from log ──────────────────────────────
/**
 * Called from the Certificate Log "Resend Selected" button.
 * @param {Array<{jlid, learnerName, course, year, parentEmail, sentBy}>} items
 * @returns {{ success:boolean, message:string }}
 */
function resendFailedCertificates(items) {
  try {
    if (!items || !items.length) return { success: false, message: 'No items to resend.' };

    // Group by learner+email (multiple learners could be selected)
    var groups = {};
    items.forEach(function(item) {
      var key = (item.jlid || item.learnerName) + '||' + item.parentEmail;
      if (!groups[key]) {
        groups[key] = {
          jlid       : item.jlid        || '',
          learnerName: item.learnerName  || '',
          parentEmail: item.parentEmail  || '',
          parentName : 'Parent',
          performedBy: item.sentBy       || '',
          courses    : []
        };
      }
      groups[key].courses.push({ name: item.course, year: item.year || new Date().getFullYear() });
    });

    var totalSent   = 0;
    var totalFailed = [];
    var messages    = [];

    Object.keys(groups).forEach(function(key) {
      var g = groups[key];
      Logger.log('[Cert] Resending ' + g.courses.length + ' cert(s) for ' + g.learnerName);
      var result = sendBulkCertificates(g);
      if (result.success) {
        totalSent += result.sent || 0;
        if (result.failed && result.failed.length) {
          result.failed.forEach(function(f){ totalFailed.push(f); });
        }
        messages.push(g.learnerName + ': ' + (result.sent || 0) + ' sent');
      } else {
        totalFailed.push(g.learnerName + ': ' + (result.message || 'failed'));
        messages.push(g.learnerName + ': failed — ' + result.message);
      }
    });

    var msg = totalSent + ' certificate(s) resent.'
      + (totalFailed.length ? ' Failed: ' + totalFailed.join(', ') : '');

    return {
      success: totalSent > 0 || totalFailed.length === 0,
      sent   : totalSent,
      failed : totalFailed,
      message: msg
    };

  } catch(e) {
    Logger.log('[Cert] resendFailedCertificates error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Quick test (run from editor) ──────────────────────────────────────
function testCertificateGeneration() {
  var tests = [
    { learner: 'Ahaan Padia',  course: 'Animation with Scratch Jr',   expected: 'Slide 1 (Foundation)' },
    { learner: 'Riya Sharma',  course: 'Maths Year 4',                expected: 'Slide 2 (Maths)'      },
    { learner: 'Arjun Mehta', course: 'Python Game Developer',        expected: 'Slide 3 (Pro)'        }
  ];

  tests.forEach(function(t) {
    Logger.log('\n── Test: ' + t.learner + ' | ' + t.course + ' (expect: ' + t.expected + ')');
    var idx = _getCertSlideIndex(t.course);
    Logger.log('   Slide index: ' + idx + ' → ' + ['Foundation','Maths','Pro/Advanced'][idx]);

    var pdf = generateCertificatePDF(t.learner, t.course);
    if (pdf) {
      Logger.log('   ✅ PDF generated: ' + pdf.getName() + ' (' + Math.round(pdf.getBytes().length / 1024) + ' KB)');
      Logger.log('   📁 Drive URL: ' + (pdf._driveUrl || '(not saved)'));
    } else {
      Logger.log('   ❌ PDF generation failed');
    }
  });
}
