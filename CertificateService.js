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
    var tempFile = DriveApp.getFileById(CERT_TEMPLATE_ID)
                           .makeCopy('_cert_temp_' + Date.now());
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

    // 4. Export as PDF via Drive API
    var token    = ScriptApp.getOAuthToken();
    var exportUrl = 'https://docs.google.com/presentation/d/' + tempPresId + '/export/pdf';
    var response  = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error('PDF export failed: HTTP ' + response.getResponseCode());
    }

    var fileName = learnerName.replace(/\s+/g, '_') + '_' + courseName.replace(/\s+/g, '_') + '_Certificate.pdf';
    var pdfBlob  = response.getBlob().setName(fileName);

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
    Logger.log('[Cert] generateCertificatePDF error: ' + e.message);
    return null;
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

    var blobs    = [];
    var failed   = [];
    var sentList = [];

    courses.forEach(function(c) {
      try {
        var blob = generateCertificatePDF(learnerName, c.name, c.year);
        if (blob) {
          blobs.push(blob);
          sentList.push(c.name + ' (' + c.year + ')');
        } else {
          failed.push(c.name);
        }
      } catch(e) {
        failed.push(c.name + ': ' + e.message);
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

    // ── Log each cert ─────────────────────────────────────────────────────────
    courses.forEach(function(c) {
      var matchBlob = blobs.filter(function(b){ return b.getName().indexOf(c.name.replace(/\s+/g,'_')) > -1; });
      var driveUrl  = matchBlob.length ? (matchBlob[0]._driveUrl || '') : '';
      _logCertificate(
        jlid, learnerName, c.name, c.year,
        parentEmail, performedBy, driveUrl,
        failed.indexOf(c.name) > -1 ? 'Failed' : 'Sent',
        'Bulk send — ' + blobs.length + ' certificate(s)'
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
