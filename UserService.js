// =============================================
// FILE: UserService.js
// =============================================

// =============================================
// AUTHENTICATION & USER MANAGEMENT
// =============================================

function getUserProfiles() {
  try {
    const data = _getCachedSheetData(CONFIG.SHEETS.USER_PROFILES);
    if (!data || data.length <= 1) return [];
    
    const headers = data[0].map(h => String(h).trim().toLowerCase().replace(/\s/g, ''));
    
    return data.slice(1).map(row => ({
      username: row[headers.indexOf('username')],
      password: row[headers.indexOf('password')],
      role: row[headers.indexOf('role')],
      email: row[headers.indexOf('email')],
      isActive: row[headers.indexOf('isactive')] === true || 
                String(row[headers.indexOf('isactive')]).toLowerCase() === 'true',
      lastLogin: row[headers.indexOf('lastlogin')],
      mustChangePassword: row[headers.indexOf('mustchangepassword')] === true || 
                          String(row[headers.indexOf('mustchangepassword')]).toLowerCase() === 'true'
    }));
  } catch (e) {
    Logger.log('Error in getUserProfiles: ' + e.message);
    return [];
  }
}

function authenticateUser(username, password) {
  Logger.log('authenticateUser called for username: ' + username);

  if (!username || !password) {
    return { success: false, role: ROLES.GUEST, message: 'Missing credentials' };
  }

  try {
    const userProfiles = getUserProfiles();
    const user = userProfiles.find(u => u.username === username);

    if (!user) {
      return { success: false, role: ROLES.GUEST, message: 'Invalid credentials' };
    }

    if (user.password !== password) {
      logUserActivity(username, 'Failed Login', 'Invalid credentials');
      return { success: false, role: ROLES.GUEST, message: 'Invalid credentials' };
    }

    if (!user.isActive) {
      return { success: false, role: ROLES.GUEST, message: 'Account inactive' };
    }

    // ✅ Build response FIRST, then do slow writes after
    const response = {
      success: true,
      role: user.role,
      username: username,
      permissions: PERMISSIONS[user.role] || [],
      mustChangePassword: user.mustChangePassword === true ||
                          String(user.mustChangePassword).toLowerCase() === 'true'
    };

    // ✅ Do sheet writes AFTER preparing response (still sync but ordered for clarity)
    try { updateUserLastLogin(username); } catch(e) { Logger.log('LastLogin update failed: ' + e.message); }
    try { logUserActivity(username, 'Successful Login', 'User logged in'); } catch(e) {}

    return response;

  } catch (error) {
    Logger.log('Error in authenticateUser: ' + error.message);
    return { success: false, role: ROLES.GUEST, message: 'Authentication error' };
  }
}


function verifyUserSession(username) {
  Logger.log('verifyUserSession called for: ' + username);
  try {
      const userProfiles = getUserProfiles();
      // Empty list = sheet unreadable (permission/quota error) — treat as transient
      if (!userProfiles || userProfiles.length === 0) {
        Logger.log('verifyUserSession: user list empty (transient sheet error)');
        return { success: false, transient: true, message: 'Session check temporarily unavailable.' };
      }
      const user = userProfiles.find(u => u.username === username);

      if (!user || !user.isActive) {
        Logger.log('Session verification failed for user: ' + username);
        return { success: false, message: 'Invalid or inactive session.' };
      }

      Logger.log('Session verification successful for user: ' + username);
      return {
        success: true,
        role: user.role,
        username: user.username,
        permissions: PERMISSIONS[user.role] || []
      };
  } catch (error) {
    Logger.log('Error in verifyUserSession: ' + error.message);
    return { success: false, transient: true, message: 'Session verification error.' };
  }
}


function createDefaultUsers() {
  try {
    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.USER_PROFILES);

    if (sheet.getLastRow() === 0 || sheet.getRange('A1').isBlank()) {
      sheet.getRange(1, 1, 1, 9).setValues([
        ['Username', 'Password', 'Role', 'Email', 'IsActive', 'LastLogin', 'CreatedDate', 'ResetToken', 'TokenExpiry']
      ]);
    }

    const defaultUsers = [
      ['Admin', 'JetLearn2025$', ROLES.ADMIN, 'admin@jet-learn.com', true, '', new Date(), '', ''],
      ['Ops_team', 'Opsteam@2025$', ROLES.USER, 'ops@jet-learn.com', true, '', new Date(), '', '']
    ];

    sheet.getRange(sheet.getLastRow() + 1, 1, defaultUsers.length, 9).setValues(defaultUsers);
    Logger.log('Default users created');

    delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.USER_PROFILES}`];

  } catch (error) {
    Logger.log('Error creating default users: ' + error.message);
  }
}

function updateUserLastLogin(username) {
  try {
    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.USER_PROFILES);
    const data = _getCachedSheetData(CONFIG.SHEETS.USER_PROFILES);
    const headers = data[0].map(h => h.toLowerCase().replace(/\s/g, ''));
    const usernameColIndex = headers.indexOf('username');
    const lastLoginColIndex = headers.indexOf('lastlogin');

    if (usernameColIndex === -1 || lastLoginColIndex === -1) {
      Logger.log('User Profiles sheet missing Username or LastLogin column.');
      return;
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][usernameColIndex] === username) {
        sheet.getRange(i + 1, lastLoginColIndex + 1).setValue(new Date());
        delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.USER_PROFILES}`];
        break;
      }
    }
  } catch (error) {
    Logger.log('Error updating last login: ' + error.message);
  }
}

function requestPasswordReset(email) {
  Logger.log('requestPasswordReset called for email: ' + email);
  try {
    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.USER_PROFILES);
    const data = _getCachedSheetData(CONFIG.SHEETS.USER_PROFILES);

    if (data.length < 1) {
        Logger.log("User Profiles sheet is empty or does not have headers. Cannot process password reset.");
        return { success: false, message: "Server configuration error: User profiles not set up." };
    }

    const headers = data[0].map(h => h.toLowerCase().replace(/\s/g, ''));
    const emailCol = headers.indexOf('email');
    const usernameCol = headers.indexOf('username');
    const tokenCol = headers.indexOf('resettoken');
    const expiryCol = headers.indexOf('tokenexpiry');

    if (emailCol === -1 || tokenCol === -1 || expiryCol === -1 || usernameCol === -1) {
      Logger.log("User Profiles sheet missing one or more required columns for password reset (Email, ResetToken, TokenExpiry, Username).");
      return { success: false, message: "Server configuration error: Required columns for password reset not found. Please check 'User Profiles' sheet headers." };
    }

    let userRowDataIndex = -1;
    for (let i = 1; i < data.length; i++) {
        if (data[i][emailCol] && String(data[i][emailCol]).trim().toLowerCase() === email.trim().toLowerCase()) {
            userRowDataIndex = i;
            break;
        }
    }

    if (userRowDataIndex === -1) {
      Logger.log('Email address not found in user profiles: ' + email);
      return { success: false, message: 'Email address not found.' };
    }

    const userSheetRowIndex = userRowDataIndex + 1;

    const token = Utilities.getUuid();
    const expiry = new Date(new Date().getTime() + 60 * 60 * 1000);

    sheet.getRange(userSheetRowIndex, tokenCol + 1).setValue(token);
    sheet.getRange(userSheetRowIndex, expiryCol + 1).setValue(expiry);
    Logger.log(`Generated token for ${email}: ${token}, expires: ${expiry.toLocaleString()}`);

    const webAppUrl = ScriptApp.getService().getUrl();
    const resetUrl = `${webAppUrl}?resetToken=${token}`;
    Logger.log('Generated reset URL: ' + resetUrl);

    const username = data[userRowDataIndex][usernameCol];

    const emailBody = `
      <p>Hello ${username},</p>
      <p>A password reset was requested for your JetLearn System account. Please click the link below to reset your password. This link is valid for 1 hour.</p>
      <p><a href="${resetUrl}" style="display: inline-block; padding: 10px 15px; background-color: #4a3c8a; color: white; text-decoration: none; border-radius: 5px;">Reset Password</a></p>
      <p>If you did not request this, please ignore this email.</p>
      <p>Thanks,<br>The JetLearn Team</p>
    `;

    MailApp.sendEmail({
      to: email,
      subject: 'JetLearn System - Password Reset Request',
      htmlBody: emailBody,
      name: CONFIG.EMAIL.FROM_NAME,
      from: CONFIG.EMAIL.FROM
    });
    Logger.log('Password reset email sent to: ' + email);

    delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.USER_PROFILES}`];

    return { success: true, message: 'A password reset link has been sent to your email.' };
  } catch (error) {
    Logger.log('Error in requestPasswordReset: ' + error.message);
    if (error.stack) {
        Logger.log('Stack trace: ' + error.stack);
    }
    return { success: false, message: 'An error occurred. Please try again later.' };
  }
}

function resetPassword(token, newPassword) {
  Logger.log('resetPassword called with token: ' + token);
  try {
    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.USER_PROFILES);
    const data = _getCachedSheetData(CONFIG.SHEETS.USER_PROFILES);

    if (data.length < 1) {
        Logger.log("User Profiles sheet is empty or does not have headers. Cannot process password reset.");
        return { success: false, message: "Server configuration error: User profiles not set up." };
    }

    const headers = data[0].map(h => h.toLowerCase().replace(/\s/g, ''));
    const tokenCol = headers.indexOf('resettoken');
    const expiryCol = headers.indexOf('tokenexpiry');
    const passwordCol = headers.indexOf('password');

    if (tokenCol === -1 || expiryCol === -1 || passwordCol === -1) {
      Logger.log("User Profiles sheet missing one or more required columns for password reset (ResetToken, TokenExpiry, Password).");
      return { success: false, message: "Server configuration error: Required columns for password reset not found. Please check 'User Profiles' sheet headers." };
    }

    let userRowDataIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][tokenCol] && String(data[i][tokenCol]).trim() === String(token).trim()) {
        userRowDataIndex = i;
        break;
      }
    }

    if (userRowDataIndex === -1) {
      Logger.log('Invalid or non-existent reset token provided: ' + token);
      return { success: false, message: 'Invalid or expired reset token.' };
    }

    const userSheetRowIndex = userRowDataIndex + 1;

    const expiryDate = new Date(data[userRowDataIndex][expiryCol]);
    if (isNaN(expiryDate.getTime())) {
        Logger.log(`Invalid expiry date for token ${token}: ${data[userRowDataIndex][expiryCol]}`);
        return { success: false, message: 'Invalid token expiry date. Please request a new link.' };
    }

    if (expiryDate < new Date()) {
      Logger.log(`Token ${token} has expired.`);
      sheet.getRange(userSheetRowIndex, tokenCol + 1).setValue('');
      sheet.getRange(userSheetRowIndex, expiryCol + 1).setValue('');
      delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.USER_PROFILES}`];
      return { success: false, message: 'Password reset token has expired.' };
    }

    if (!newPassword || newPassword.length < 6) {
        return { success: false, message: 'New password must be at least 6 characters long.' };
    }

    sheet.getRange(userSheetRowIndex, passwordCol + 1).setValue(newPassword);
    Logger.log(`Password updated for user at row ${userSheetRowIndex}.`);

    sheet.getRange(userSheetRowIndex, tokenCol + 1).setValue('');
    sheet.getRange(userSheetRowIndex, expiryCol + 1).setValue('');
    Logger.log(`Token ${token} invalidated after use.`);

    delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.USER_PROFILES}`];

    return { success: true, message: 'Your password has been reset successfully.' };
  } catch (error) {
    Logger.log('Error in resetPassword: ' + error.message);
    if (error.stack) {
        Logger.log('Stack trace: ' + error.stack);
    }
    return { success: false, message: 'An error occurred while resetting the password.' };
  }
}

function getActiveUsers() {
  Logger.log('getActiveUsers called');

  try {
    const users = getUserProfiles();
    return users.map(user => ({
      username: user.username,
      role: user.role,
      email: user.email,
      lastLogin: user.lastLogin ? new Date(user.lastLogin).toLocaleString() : 'Never',
      createdDate: user.createdDate ? new Date(user.createdDate).toLocaleDateString() : 'N/A',
      isActive: user.isActive
    }));
  } catch (error) {
    Logger.log('Error getting active users: ' + error.message);
    return [];
  }
}


function updateUser(userData, currentUser) {
  Logger.log(`updateUser called for '${userData.username}' by user '${currentUser.username}'`);

  try {
    if (!currentUser || !hasPermission(currentUser.role, 'manage_users')) {
      logUserActivity(currentUser.username, 'Update User Failed', `Permission denied to update ${userData.username}.`);
      return { success: false, message: 'Permission denied. Only Admins can manage users.' };
    }

    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.USER_PROFILES);
    const data = _getCachedSheetData(CONFIG.SHEETS.USER_PROFILES);

    let rowIndex = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === userData.username) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: `User '${userData.username}' not found.` };
    }

    const headers = data[0].map(h => h.toLowerCase().replace(/\s/g, ''));
    const changes = [];

    if (userData.password) {
      const colIndex = headers.indexOf('password');
      if (colIndex > -1) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(userData.password);
        changes.push('password');
      }
    }

    if (userData.role) {
      if (!Object.values(ROLES).includes(userData.role)) {
        return { success: false, message: `Invalid role: ${userData.role}` };
      }
      const colIndex = headers.indexOf('role');
      if (colIndex > -1) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(userData.role);
        changes.push(`role to '${userData.role}'`);
      }
    }

    if (userData.email) {
       if (!isValidEmail(userData.email)) {
        return { success: false, message: `Invalid email address: ${userData.email}` };
      }
      const colIndex = headers.indexOf('email');
      if (colIndex > -1) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(userData.email);
        changes.push(`email to '${userData.email}'`);
      }
    }

    if (userData.hasOwnProperty('isActive')) {
      const colIndex = headers.indexOf('isactive');
      if (colIndex > -1) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(userData.isActive);
        changes.push(`status to '${userData.isActive ? 'Active' : 'Inactive'}'`);
      }
    }

    if (changes.length > 0) {
      const logDetails = `Admin '${currentUser.username}' updated user '${userData.username}': changed ${changes.join(', ')}.`;
      logAction('User Updated', '', '', '', '', '', 'Success', logDetails);
      Logger.log(logDetails);
      delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.USER_PROFILES}`];
      return { success: true, message: 'User updated successfully.' };
    } else {
      return { success: true, message: 'No changes were applied.' };
    }

  } catch (error) {
    Logger.log('Error updating user: ' + error.message);
    return { success: false, message: 'An unexpected error occurred while updating the user.' };
  }
}

// ── changeOwnPassword ─────────────────────────────────────────────────
// Called on first-login forced password change.
// No admin permission required — user can only change their OWN password.
// Clears mustChangePassword flag after save.
function changeOwnPassword(username, newPassword) {
  try {
    if (!username || !newPassword) return { success: false, message: 'Username and password required.' };
    if (newPassword.length < 6) return { success: false, message: 'Password must be at least 6 characters.' };

    var sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.USER_PROFILES);
    var data  = _getCachedSheetData(CONFIG.SHEETS.USER_PROFILES);
    if (!data || data.length < 2) return { success: false, message: 'User profiles sheet not found.' };

    var headers = data[0].map(function(h){ return String(h).toLowerCase().replace(/\s/g,''); });
    var passIdx = headers.indexOf('password');
    var mcpIdx  = headers.indexOf('mustchangepassword');

    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(username).trim()) { rowIndex = i + 1; break; }
    }
    if (rowIndex === -1) return { success: false, message: 'User not found.' };

    if (passIdx > -1) sheet.getRange(rowIndex, passIdx + 1).setValue(newPassword);
    if (mcpIdx  > -1) sheet.getRange(rowIndex, mcpIdx  + 1).setValue(false);

    // Bust cache so next login reads fresh data
    var cacheKey = CONFIG.MIGRATION_SHEET_ID + '_' + CONFIG.SHEETS.USER_PROFILES;
    if (typeof _sheetDataCache !== 'undefined') delete _sheetDataCache[cacheKey];

    Logger.log('[Users] changeOwnPassword: ' + username + ' — password updated, mustChangePassword cleared');
    return { success: true, message: 'Password updated successfully.' };
  } catch(e) {
    Logger.log('[Users] changeOwnPassword error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function updateSystemConfig(newConfig) {
  Logger.log('updateSystemConfig called');

  try {
    if (newConfig.emailFrom && !isValidEmail(newConfig.emailFrom)) {
      return { success: false, message: 'Invalid email address' };
    }

    if (newConfig.paginationLimit && (isNaN(newConfig.paginationLimit) || newConfig.paginationLimit < 10 || newConfig.paginationLimit > 100)) {
      return { success: false, message: 'Pagination limit must be between 10 and 100' };
    }

    Logger.log('System config (pseudo) updated: ' + JSON.stringify(newConfig));
    return { success: true, message: 'System configuration updated successfully' };
  } catch (error) {
    Logger.log('Error updating system config: ' + error.message);
    return { success: false, message: 'Error updating system config: ' + error.message };
  }
}

function getSystemSettings() {
  return {
    emailFrom: CONFIG.EMAIL.FROM,
    emailFromName: CONFIG.EMAIL.FROM_NAME,
    paginationLimit: CONFIG.PAGINATION_LIMIT,
    auditRetentionDays: 90
  };
}

function hasPermission(userRole, permission) {
  const userPermissions = PERMISSIONS[userRole] || [];
  return userPermissions.includes(permission);
}

function addNewUser(userData, actingUser) {
  try {
    if (!actingUser || !hasPermission(actingUser.role, 'create_users')) {
      return { success: false, message: 'Permission denied. Only Super Admins can create users.' };
    }

    const validation = validateInput(userData, ['username', 'email', 'role']);
    if (!validation.isValid) return { success: false, message: validation.message };
    if (!isValidEmail(userData.email)) return { success: false, message: 'Invalid email address.' };

    const sheet = getOrCreateSheet(CONFIG.SHEETS.USER_PROFILES);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];

    const emailColIndex = headers.indexOf('Email');
    const usernameColIndex = headers.indexOf('Username');
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][emailColIndex]).toLowerCase() === userData.email.toLowerCase()) {
        return { success: false, message: 'A user with this email already exists.' };
      }
      if (String(allData[i][usernameColIndex]).toLowerCase() === userData.username.toLowerCase()) {
        return { success: false, message: 'Username is already taken.' };
      }
    }

    const tempPassword = generateTempPassword();

    sheet.appendRow([
      userData.username,  // Username
      tempPassword,       // Password
      userData.role,      // Role
      userData.email,     // Email
      true,               // IsActive
      '',                 // LastLogin
      new Date(),         // CreatedDate
      '',                 // ResetToken
      '',                 // TokenExpiry
      true                // MustChangePassword
    ]);

    SpreadsheetApp.flush();

    const emailResult = sendWelcomeEmail(userData.email, userData.username, tempPassword);

    logAuditAction('User Created', `Super Admin '${actingUser.username}' created new ${userData.role} account for '${userData.username}' (${userData.email})`);
    Logger.log(`User created: ${userData.username} by ${actingUser.username}`);

    return {
      success: true,
      message: `User '${userData.username}' created. Welcome email ${emailResult.success ? 'sent ✓' : 'FAILED — check logs'}.`
    };

  } catch (e) {
    logError('addNewUser', e);
    return { success: false, message: 'Error adding user: ' + e.message };
  }
}

// =============================================
// SEND WELCOME EMAIL (Super Admin only)
// =============================================
function sendWelcomeEmail(toEmail, username, tempPassword) {
  try {
    const platformUrl = ScriptApp.getService().getUrl();
    const platformLabel = 'JetLearn Operation System';
    const subject = `Your account is ready`;

    const htmlBody = `
<!DOCTYPE html>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <meta name="x-apple-disable-message-reformatting" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <title>Account Created — JetLearn</title>
  <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet" />

  <style>
    * { box-sizing: border-box; }
    body, table, td, a { -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; }
    table, td { mso-table-lspace: 0pt; mso-table-rspace: 0pt; }
    img { border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; -ms-interpolation-mode: bicubic; }
    body { margin: 0 !important; padding: 0 !important; width: 100% !important; }

    body { background-color: #f0ede8; font-family: 'DM Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }

    .email-bg       { background-color: #f0ede8; }
    .card-outer     { background-color: #ffffff; border-radius: 16px; border: 1px solid #e4e1dc; overflow: hidden; }
    .rail-bg        { background-color: #f5f3ff; border-right: 1px solid #e4e1f8; }
    .rail-brand     { color: #5a5888; }
    .rail-heading   { color: #2d2a6e; }
    .rail-lbl       { color: #5a5888; }
    .rail-val       { color: #333355; }
    .rail-divider   { background-color: #e0ddf8; }
    .rail-footer    { color: #7068a8; }
    .rail-icon-bg   { background-color: #ede9ff; }

    .main-bg        { background-color: #ffffff; }
    .action-tag     { color: #1e40af; background-color: #eff6ff; border: 1px solid #93c5fd; }
    .header-text    { color: #666666; }
    .header-strong  { color: #333333; }
    .section-lbl    { color: #767370; }
    .detail-lbl     { color: #767676; }
    .detail-val     { color: #111111; }
    .detail-border  { border-bottom: 1px solid #f8f6f3; }
    .notice-bg      { background-color: #fffbeb; border: 1px solid #fcd34d; }
    .notice-title   { color: #92400e; }
    .notice-text    { color: #78350f; }
    .closing-bg     { background-color: #f5f3ff; border: 1px solid #ddd8fa; }
    .closing-text   { color: #4c1d95; }
    .footer-sub     { color: #767676; }
    .footer-note    { color: #767676; }

    .btn-primary {
      background-color: #7c6ef0;
      color: #ffffff !important;
      text-decoration: none;
      font-weight: 500;
      border-radius: 8px;
      display: inline-block;
    }

    @media only screen and (max-width: 600px) {
      .wrapper         { width: 100% !important; padding: 16px 12px !important; }
      .card-outer      { border-radius: 12px !important; }
      .rail-cell       { display: block !important; width: 100% !important; border-radius: 0 !important; padding: 20px 20px !important; }
      .main-cell       { display: block !important; width: 100% !important; }
      .rail-meta-row   { display: flex !important; flex-wrap: wrap !important; gap: 14px !important; }
      .rail-meta-item  { flex: 1 1 80px !important; }
      .rail-heading    { font-size: 14px !important; margin-bottom: 12px !important; }
      .rail-icon-wrap  { display: none !important; }
      .rail-spacer     { display: none !important; }
      .rail-footer-row { display: none !important; }
      .main-header     { padding: 18px 20px 16px !important; }
      .main-body       { padding: 16px 20px 20px !important; }
      .main-footer     { padding: 13px 20px !important; }
    }
  </style>
</head>
<body class="email-bg">

<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="email-bg">
  <tr>
    <td align="center" style="padding: 40px 16px;" class="wrapper">

      <!-- Top label -->
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:640px; margin-bottom:14px;">
        <tr>
          <td>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0">
              <tr>
                <td style="width:6px; height:6px; border-radius:50%; background-color:#7c6ef0; vertical-align:middle;"></td>
                <td style="padding-left:8px; font-family:'DM Mono','Courier New',monospace; font-size:10px; letter-spacing:0.13em; text-transform:uppercase;" class="rail-lbl">JetLearn &mdash; Operations System</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>

      <!-- Card -->
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:640px;" class="card-outer">
        <tr valign="top">

          <!-- RAIL -->
          <td width="210" class="rail-bg rail-cell" style="padding:32px 24px; vertical-align:top; border-radius:16px 0 0 16px;">
            <p style="font-family:'DM Mono','Courier New',monospace; font-size:10px; letter-spacing:0.13em; text-transform:uppercase; margin:0 0 24px;" class="rail-brand">JetLearn</p>

            <table role="presentation" cellpadding="0" cellspacing="0" border="0" class="rail-icon-wrap" style="margin-bottom:18px;">
              <tr>
                <td style="width:32px; height:32px; border-radius:8px; text-align:center; vertical-align:middle;" class="rail-icon-bg">
                  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <rect x="3" y="3" width="10" height="10" rx="2" stroke="#7c6ef0" stroke-width="1.5"/>
                    <path d="M8 6v4M6 8h4" stroke="#7c6ef0" stroke-width="1.5" stroke-linecap="round"/>
                  </svg>
                </td>
              </tr>
            </table>

            <p style="font-size:15px; font-weight:500; line-height:1.45; margin:0 0 22px;" class="rail-heading">Account<br/>provisioning</p>

            <div class="rail-divider" style="height:1px; margin-bottom:22px;"></div>

            <div class="rail-meta-row">
              <div class="rail-meta-item rail-item" style="margin-bottom:16px;">
                <p style="font-family:'DM Mono','Courier New',monospace; font-size:9px; text-transform:uppercase; letter-spacing:0.12em; margin:0 0 4px;" class="rail-lbl">Username</p>
                <p style="font-size:12px; font-weight:500; margin:0; line-height:1.4;" class="rail-val">${username}</p>
              </div>

              <div class="rail-meta-item rail-item" style="margin-bottom:16px;">
                <p style="font-family:'DM Mono','Courier New',monospace; font-size:9px; text-transform:uppercase; letter-spacing:0.12em; margin:0 0 4px;" class="rail-lbl">Platform</p>
                <p style="font-size:12px; font-weight:500; margin:0; line-height:1.4;" class="rail-val">${platformLabel}</p>
              </div>

              <div class="rail-meta-item rail-item" style="margin-bottom:0;">
                <p style="font-family:'DM Mono','Courier New',monospace; font-size:9px; text-transform:uppercase; letter-spacing:0.12em; margin:0 0 6px;" class="rail-lbl">Status</p>
                <div style="display:inline-block; width:30px; height:30px; border-radius:50%; text-align:center; line-height:30px; font-size:11px; font-weight:500; background-color:#edfaf3; border:1px solid #7dd4a8; color:#0f6630;">NEW</div>
              </div>
            </div>

            <div class="rail-spacer" style="height:40px;"></div>

            <p class="rail-footer rail-footer-row" style="font-family:'DM Mono','Courier New',monospace; font-size:9px; letter-spacing:0.1em; text-transform:uppercase; line-height:1.6; margin:0;">Automated<br/>system alert</p>
          </td>

          <!-- MAIN -->
          <td class="main-bg main-cell" style="vertical-align:top; border-radius:0 16px 16px 0;">

            <!-- Header -->
            <div class="main-header" style="padding:28px 28px 22px; border-bottom:1px solid #f2efea;">
              <p style="display:inline-block; font-family:'DM Mono','Courier New',monospace; font-size:9px; letter-spacing:0.12em; text-transform:uppercase; border-radius:6px; padding:3px 9px; margin:0 0 8px;" class="action-tag">Account Ready</p>
              <p style="font-size:13px; line-height:1.7; margin:0;" class="header-text">
                Hi <strong class="header-strong">${username}</strong>, your JetLearn Operations account has been successfully created. You can now access the platform using the credentials below.
              </p>
            </div>

            <!-- Body -->
            <div class="main-body" style="padding:22px 28px;">

              <p style="font-family:'DM Mono','Courier New',monospace; font-size:9px; letter-spacing:0.13em; text-transform:uppercase; margin:0 0 12px;" class="section-lbl">Access credentials</p>
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:24px;">
                <tr class="detail-border" style="border-bottom:1px solid #f8f6f3;">
                  <td style="font-size:13px; padding:12px 0;" class="detail-lbl">Platform URL</td>
                  <td style="font-size:13px; font-weight:500; padding:12px 0; text-align:right;"><a href="${platformUrl}" style="text-decoration:none; color:#5546d4;">${platformUrl} &rarr;</a></td>
                </tr>
                <tr class="detail-border" style="border-bottom:1px solid #f8f6f3;">
                  <td style="font-size:13px; padding:12px 0;" class="detail-lbl">Username</td>
                  <td style="font-size:13px; font-weight:500; padding:12px 0; text-align:right;" class="detail-val">${username}</td>
                </tr>
                <tr>
                  <td style="font-size:13px; padding:12px 0;" class="detail-lbl">Temporary Password</td>
                  <td style="font-size:13px; font-weight:500; padding:12px 0; text-align:right; font-family:'DM Mono',monospace; background:#f6f4ff; padding-right:8px; border-radius:4px;" class="detail-val">${tempPassword}</td>
                </tr>
              </table>

              <!-- Action Button -->
              <div style="text-align:center; margin-bottom:24px;">
                <a href="${platformUrl}" class="btn-primary" style="padding:14px 32px; font-size:14px;">Sign in to your account</a>
              </div>

              <!-- Notice -->
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:20px; border-radius:9px; overflow:hidden;" class="notice-bg">
                <tr><td style="padding:14px 16px; border-radius:9px; border-left:3px solid #f59e0b;">
                  <p style="font-size:12px; font-weight:500; margin:0 0 4px;" class="notice-title">Security Note</p>
                  <p style="font-size:13px; line-height:1.65; margin:0;" class="notice-text">This is a temporary password. For your security, you will be required to change it immediately after your first login.</p>
                </td></tr>
              </table>

              <!-- Closing -->
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-radius:9px; overflow:hidden;" class="closing-bg">
                <tr><td style="padding:18px 20px; text-align:center; border-radius:9px;">
                  <p style="font-size:14px; font-weight:500; line-height:1.7; margin:0;" class="closing-text">Welcome aboard! We're excited to have you on the JetLearn Operations team.</p>
                </td></tr>
              </table>

            </div>

            <!-- Footer -->
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-top:1px solid #f2efea;">
              <tr>
                <td style="padding:14px 28px;" class="main-footer">
                  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
                    <tr>
                      <td style="font-size:12px; font-weight:500;" class="footer-sub">JetLearn</td>
                      <td style="text-align:right; font-family:'DM Mono','Courier New',monospace; font-size:9px; letter-spacing:0.1em; text-transform:uppercase;" class="footer-note">Automated alert</td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>

          </td>
        </tr>
      </table>

    </td>
  </tr>
</table>

</body>
</html>
    `;

    sendTrackedEmail({
      to: toEmail,
      subject: subject,
      htmlBody: htmlBody,
      jlid: ''
    });

    Logger.log(`Welcome email sent to ${toEmail}`);
    return { success: true };

  } catch (e) {
    logError('sendWelcomeEmail', e);
    return { success: false, message: e.message };
  }
}

// =============================================
// RESEND WELCOME / LOGIN EMAIL (Super Admin only)
// =============================================
function resendLoginEmail(targetUsername, actingUser) {
  try {
    if (!actingUser || !hasPermission(actingUser.role, 'send_reset_email')) {
      return { success: false, message: 'Permission denied.' };
    }

    const sheet = getOrCreateSheet(CONFIG.SHEETS.USER_PROFILES);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];

    const usernameCol = headers.indexOf('Username');
    const emailCol    = headers.indexOf('Email');
    const statusCol   = headers.indexOf('IsActive');
    const passwordCol = headers.indexOf('Password');

    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][usernameCol]).toLowerCase() === targetUsername.toLowerCase()) {
        if (!allData[i][statusCol]) {
          return { success: false, message: 'Cannot resend email to an inactive user.' };
        }

        const newTempPass = generateTempPassword();
        sheet.getRange(i + 1, passwordCol + 1).setValue(newTempPass);

        // Also reset MustChangePassword to true
        const mustChangeCol = headers.indexOf('MustChangePassword');
        if (mustChangeCol !== -1) {
          sheet.getRange(i + 1, mustChangeCol + 1).setValue(true);
        }

        SpreadsheetApp.flush();

        const emailResult = sendWelcomeEmail(String(allData[i][emailCol]), targetUsername, newTempPass);
        if (emailResult.success) {
          logAuditAction('Login Email Resent', `Super Admin '${actingUser.username}' resent login email to '${targetUsername}'`);
          return { success: true, message: `Login email resent to ${targetUsername}.` };
        }
        return { success: false, message: 'Email send failed: ' + emailResult.message };
      }
    }

    return { success: false, message: 'User not found.' };
  } catch (e) {
    logError('resendLoginEmail', e);
    return { success: false, message: e.message };
  }
}



// =============================================
// HELPER: Generate Temporary Password
// =============================================
function generateTempPassword() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789@#$!';
  let password = '';
  for (let i = 0; i < 10; i++) {
    password += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return password;
}

// =============================================
// THIS IS THE CORRECT, ENHANCED FUNCTION
// =============================================
function getUserProfile(username) {
  try {
    const profiles = getUserProfiles();
    const user = profiles.find(u => u.username === username);
    if (!user) return null;

    // --- NEW: Calculate user-specific stats ---
    let migrationCount = 0;
    let onboardingCount = 0;
    try {
      const auditLog = _getAppDataCachedSheetData(CONFIG.APP_DATA_SHEETS.AUDIT_LOG);
      if (auditLog && auditLog.length > 1) {
        // Find the index for "Action" and "Intervened By" columns
        const headers = auditLog[0].map(h => String(h).trim());
        const actionCol = headers.indexOf('Action');
        const userCol = headers.indexOf('Intervened By');

        if (actionCol > -1 && userCol > -1) {
          // Loop through all audit records
          for (let i = 1; i < auditLog.length; i++) {
            const rowUser = String(auditLog[i][userCol] || '').trim();
            const rowAction = String(auditLog[i][actionCol] || '').trim();

            // Check if the action was performed by the current user's email
            if (rowUser.toLowerCase() === user.email.toLowerCase()) {
              if (rowAction.includes('Migration')) {
                migrationCount++;
              }
              if (rowAction.includes('Onboarding')) {
                onboardingCount++;
              }
            }
          }
        }
      }
    } catch (e) {
      Logger.log(`[getUserProfile] Failed to calculate stats for ${username}: ${e.message}`);
    }
    // --- END OF NEW STATS CALCULATION ---

    return {
      username: user.username,
      email: user.email,
      role: user.role,
      lastLogin: user.lastLogin ? new Date(user.lastLogin).toLocaleString() : 'N/A',
      migrationCount: migrationCount,   // NEW
      onboardingCount: onboardingCount  // NEW
    };
  } catch (e) {
    logError('getUserProfile', e);
    return null;
  }
}


function updateOwnProfile(data) {
  try {
    const sheet = getOrCreateSheet(CONFIG.SHEETS.USER_PROFILES);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const usernameCol = headers.indexOf('Username');
    const passwordCol = headers.indexOf('Password');
    const emailCol    = headers.indexOf('Email');
    const mustChangeCol = headers.indexOf('MustChangePassword');


    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][usernameCol]).toLowerCase() === data.username.toLowerCase()) {

        // Verify current password if changing password
        if (data.newPassword) {
          if (String(allData[i][passwordCol]) !== data.currentPassword) {
            return { success: false, message: 'Current password is incorrect.' };
          }
          sheet.getRange(i + 1, passwordCol + 1).setValue(data.newPassword);
          // After a successful password change, set MustChangePassword to FALSE
          if (mustChangeCol !== -1) {
            sheet.getRange(i + 1, mustChangeCol + 1).setValue(false);
          }
        }

        // Update email
        if (data.email) {
          sheet.getRange(i + 1, emailCol + 1).setValue(data.email);
        }

        SpreadsheetApp.flush();
        logAuditAction('Profile Updated', `User '${data.username}' updated their own profile`);
        return { success: true, message: 'Profile updated successfully.' };
      }
    }
    return { success: false, message: 'User not found.' };
  } catch (e) {
    logError('updateOwnProfile', e);
    return { success: false, message: e.message };
  }
}

function getEmailLogs() {
  try {
    const data = _getAppDataCachedSheetData(CONFIG.APP_DATA_SHEETS.EMAIL_LOGS);
    if (data.length <= 1) return [];

    const headers = data[0];
    return data.slice(1).map(row => {
      const log = {};
      headers.forEach((h, i) => { log[h] = row[i]; });
      return log;
    }).reverse(); // Most recent first
  } catch (e) {
    logError('getEmailLogs', e);
    return [];
  }
}

function logAuditAction(action, notes) {
  try {
    const sheet = getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.AUDIT_LOG);
    sheet.appendRow([
      new Date(),   // Timestamp
      action,       // Action
      '',           // JLID
      '',           // Learner
      '',           // Old Teacher
      '',           // New Teacher
      '',           // Course
      'System',     // Status
      notes,        // Notes
      '',           // Session ID
      '',           // Reason
      Session.getActiveUser().getEmail() // Intervened By
    ]);
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log('logAuditAction failed: ' + e.message);
  }
}

function toggleUserStatus(targetUsername, newStatus, actingUser) {
  try {
    if (!actingUser || !hasPermission(actingUser.role, 'manage_users')) {
      return { success: false, message: 'Permission denied.' };
    }

    const sheet = getOrCreateSheet(CONFIG.SHEETS.USER_PROFILES);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const usernameCol = headers.indexOf('Username');
    const statusCol   = headers.indexOf('IsActive');

    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][usernameCol]).toLowerCase() === targetUsername.toLowerCase()) {
        sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
        SpreadsheetApp.flush();

        const action = newStatus ? 'Activated' : 'Deactivated';
        logAuditAction(`User ${action}`, `${actingUser.username} ${action.toLowerCase()} user '${targetUsername}'`);

        return { success: true, message: `User '${targetUsername}' has been ${action.toLowerCase()}.` };
      }
    }

    return { success: false, message: 'User not found.' };
  } catch (e) {
    logError('toggleUserStatus', e);
    return { success: false, message: e.message };
  }
}


// ─────────────────────────────────────────────────────────────────────────────
// Global Impact Score
// Computes a per-user performance score (0–100) from the last 30 days of
// User Activity Log entries, benchmarked against all other users.
// ─────────────────────────────────────────────────────────────────────────────
function getUserImpactScore(username) {
  try {
    if (!username) return { success: false, message: 'No username.' };

    var POINTS = {
      'Successful Login':         0,
      'Migration Email Sent':     8,
      'Migration Processed':      8,
      'Email Sent':               3,
      'Batch Email Sent':         5,
      'Audit Completed':          6,
      'Audit Submitted':          6,
      'TP Note Added':            4,
      'Upskill Task Created':     5,
      'Invoice Generated':        5,
      'Invoice Sent':             5,
      'Task Completed':           4,
      'User Created':             6,
      'WhatsApp Sent':            2,
      'Report Generated':         3
    };

    var actData = _getAppDataCachedSheetData(CONFIG.APP_DATA_SHEETS.USER_ACTIVITY_LOG);
    if (!actData || actData.length < 2) {
      return { success: true, score: 0, percentile: 0, migrationsHandled: 0,
               emailsSent: 0, auditsRun: 0, tasksCompleted: 0,
               nextMilestone: 10, insight: 'No activity data yet.', team: '' };
    }

    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 30);

    // Col 0=Timestamp, 1=Username, 2=Action, 3=Details
    var userScores = {};  // username → { points, migrations, emails, audits, tasks }

    actData.slice(1).forEach(function(row) {
      var ts     = row[0] ? new Date(row[0]) : null;
      if (!ts || ts < cutoff) return;
      var uname  = String(row[1] || '').trim().toLowerCase();
      var action = String(row[2] || '').trim();
      if (!uname) return;

      if (!userScores[uname]) userScores[uname] = { points:0, migrations:0, emails:0, audits:0, tasks:0 };
      var pts = POINTS[action] || 0;
      userScores[uname].points += pts;
      if (action.indexOf('Migration') !== -1) userScores[uname].migrations++;
      if (action.indexOf('Email') !== -1 || action.indexOf('WhatsApp') !== -1) userScores[uname].emails++;
      if (action.indexOf('Audit') !== -1) userScores[uname].audits++;
      if (action.indexOf('Task') !== -1 || action.indexOf('Invoice') !== -1) userScores[uname].tasks++;
    });

    var nameLower = username.trim().toLowerCase();
    var myData    = userScores[nameLower] || { points:0, migrations:0, emails:0, audits:0, tasks:0 };

    // Normalise to 0–100 (cap at 200 raw points = 100)
    var MAX_RAW = 200;
    var rawScore = Math.min(myData.points, MAX_RAW);
    var score    = Math.round((rawScore / MAX_RAW) * 100 * 10) / 10;

    // Percentile: proportion of other users with lower score
    var allScores = Object.values(userScores).map(function(u) { return Math.min(u.points, MAX_RAW); });
    var below     = allScores.filter(function(s) { return s < rawScore; }).length;
    var percentile = allScores.length > 1 ? Math.round((below / (allScores.length - 1)) * 100) : 100;

    // Next milestone: next multiple of 5 above score
    var nextMilestone = Math.ceil((score + 0.1) / 5) * 5;
    if (nextMilestone > 100) nextMilestone = 100;

    // Team label from User Profiles (col 11 = Team if it exists)
    var team = '';
    try {
      var profiles = _getCachedSheetData(CONFIG.SHEETS.USER_PROFILES);
      if (profiles && profiles.length > 1) {
        var hdr = profiles[0].map(function(h) { return String(h||'').trim().toLowerCase(); });
        var uCol = hdr.indexOf('username');
        var tCol = hdr.indexOf('team');
        if (uCol !== -1 && tCol !== -1) {
          for (var i = 1; i < profiles.length; i++) {
            if (String(profiles[i][uCol]||'').trim().toLowerCase() === nameLower) {
              team = String(profiles[i][tCol]||'').trim();
              break;
            }
          }
        }
      }
    } catch(te) {}

    // Personalised insight
    var insight = '';
    if (percentile >= 99)      insight = 'Top 1% of operational managers this month. Exceptional performance.';
    else if (percentile >= 90) insight = 'Top 10% this month. Your efficiency is driving team outcomes.';
    else if (percentile >= 75) insight = 'Strong performance. You are above 75% of the team this month.';
    else if (percentile >= 50) insight = 'Solid work. Push on migrations and audits to climb the leaderboard.';
    else                       insight = 'Getting started. Complete migrations and audits to boost your score.';

    // Improvement hint
    if (myData.migrations < 3)  insight += ' Try completing more migrations.';
    else if (myData.audits < 2) insight += ' Running more audits will increase your score.';

    // Build team leaderboard (all users with activity this month)
    var leaderboard = Object.keys(userScores).map(function(u) {
      var raw = Math.min(userScores[u].points, MAX_RAW);
      return {
        username   : u,
        score      : Math.round((raw / MAX_RAW) * 100 * 10) / 10,
        migrations : userScores[u].migrations,
        emails     : userScores[u].emails,
        audits     : userScores[u].audits,
        tasks      : userScores[u].tasks
      };
    }).sort(function(a, b) { return b.score - a.score; });

    return {
      success:          true,
      score:            score,
      percentile:       percentile,
      migrationsHandled: myData.migrations,
      emailsSent:       myData.emails,
      auditsRun:        myData.audits,
      tasksCompleted:   myData.tasks,
      nextMilestone:    nextMilestone,
      insight:          insight,
      team:             team,
      leaderboard:      leaderboard   // always returned — frontend shows it for Super Admin only
    };
  } catch(e) {
    Logger.log('[getUserImpactScore] Error: ' + e.message);
    return { success: false, message: e.message };
  }
}