// =============================================
// AUTHENTICATION & USER MANAGEMENT
// =============================================

function authenticateUser(username, password) {
  Logger.log('authenticateUser called for username: ' + username);

  if (!username || !password) {
    Logger.log('Missing credentials');
    return { success: false, role: ROLES.GUEST, message: 'Missing credentials' };
  }

  try {
    const userProfiles = getUserProfiles(); 
    const user = userProfiles.find(u => u.username === username);

    if (!user) {
      Logger.log('User not found: ' + username);
      return { success: false, role: ROLES.GUEST, message: 'Invalid credentials' };
    }

    if (user.password !== password) {
      Logger.log('Invalid password for user: ' + username);
      logUserActivity(username, 'Failed Login', 'Invalid credentials');
      return { success: false, role: ROLES.GUEST, message: 'Invalid credentials' };
    }

    if (!user.isActive) {
      Logger.log('Inactive user attempted login: ' + username);
      return { success: false, role: ROLES.GUEST, message: 'Account inactive' };
    }

    updateUserLastLogin(username);
    logUserActivity(username, 'Successful Login', 'User logged in');
    const mustChange = String(user.mustChangePassword).toLowerCase() === 'true';

    Logger.log('Authentication successful for user: ' + username + ', role: ' + user.role);
    return {
    success: true,
    role: user.role,
    username: username,
    permissions: PERMISSIONS[user.role] || [],
    mustChangePassword: user.mustChangePassword === true || String(user.mustChangePassword).toLowerCase() === 'true'
      };
  } catch (error) {
    Logger.log('Error in authenticateUser: ' + error.message);
    return { success: false, role: ROLES.GUEST, message: 'Authentication error' };
  }
}

function verifyUserSession(username) {
  Logger.log('verifyUserSession called for: ' + username);
  try {
      const userProfiles = getUserProfiles(); 
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
    return { success: false, message: 'Session verification error.' };
  }
}

function getUserProfiles() {
  try {
    const sheetData = _getCachedSheetData(CONFIG.SHEETS.USER_PROFILES);

    if (sheetData.length <= 1) { 
      createDefaultUsers(); 
      return getUserProfiles();
    }

    const headers = sheetData[0];
    return sheetData.slice(1).map(row => {
      const user = {};
      headers.forEach((header, i) => {
        const key = header.toLowerCase().replace(/\s/g, ''); 
        user[key] = row[i];
      });
    return {
      username: user.username,
      password: user.password,
      role: user.role,
      email: user.email,
      isActive: user.isactive,
      lastLogin: user.lastlogin,
      createdDate: user.createddate,
      mustChangePassword: user.mustchangepassword || false  // ← ADD THIS
    };

    });
  } catch (error) {
    Logger.log('Error getting user profiles: ' + error.message);
    return [];
  }
}

function createDefaultUsers() {
  try {
    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.USER_PROFILES);

    if (sheet.getLastRow() === 0 || sheet.getRange('A1').isBlank()) { 
      sheet.getRange(1, 1, 9, 9).setValues([
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
    if (!currentUser || currentUser.role !== ROLES.ADMIN || !hasPermission(currentUser.role, 'manage_users')) {
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
    if (!actingUser || actingUser.role !== 'Super Admin') {
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
    const platformUrl = 'https://jetlearn-launcher.vercel.app/';
    const platformLabel = 'JetLearn Operation System';
    const subject = `Your account is ready`;

    const htmlBody = `
    <div style="background:#f4f4f0;padding:40px 24px;font-family:Inter,Arial,sans-serif;">
      <div style="background:#ffffff;border-radius:4px;overflow:hidden;max-width:560px;margin:0 auto;">

        <div style="padding:40px 48px 0;">

          <div style="display:flex;align-items:center;gap:10px;margin-bottom:48px;">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#4a3c8a" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 2L11 13"/><path d="M22 2L15 22 11 13 2 9l20-7z"/></svg>
            <span style="font-size:15px;font-weight:600;color:#1a1a1a;">JetLearn</span>
          </div>

          <p style="font-size:26px;font-weight:600;color:#1a1a1a;margin:0 0 12px;letter-spacing:-0.5px;line-height:1.2;">You're in, ${username}.</p>
          <p style="font-size:15px;color:#6b6b6b;margin:0 0 40px;line-height:1.6;">Your JetLearn Operations account has been created. Use the details below to sign in for the first time.</p>

          <div style="border-top:1px solid #f0f0f0;border-bottom:1px solid #f0f0f0;padding:24px 0;margin-bottom:32px;">

            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:18px;">
              <span style="font-size:12px;color:#9a9a9a;text-transform:uppercase;letter-spacing:0.6px;">Platform</span>
              <a href="${platformUrl}" style="font-size:14px;color:#4a3c8a;text-decoration:none;font-weight:500;">${platformLabel}</a>
            </div>

            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:18px;">
              <span style="font-size:12px;color:#9a9a9a;text-transform:uppercase;letter-spacing:0.6px;">Username</span>
              <span style="font-size:14px;color:#1a1a1a;font-weight:500;">${username}</span>
            </div>

            <div style="display:flex;justify-content:space-between;align-items:center;">
              <span style="font-size:12px;color:#9a9a9a;text-transform:uppercase;letter-spacing:0.6px;">Password</span>
              <code style="font-size:14px;color:#1a1a1a;background:#f6f4ff;padding:4px 12px;border-radius:4px;letter-spacing:1.5px;">${tempPassword}</code>
            </div>

          </div>

          <a href="${platformUrl}" style="display:block;background:#4a3c8a;color:white;text-decoration:none;font-size:14px;font-weight:500;padding:14px 24px;border-radius:4px;text-align:center;margin-bottom:24px;">Sign in to your account</a>

          <p style="font-size:13px;color:#9a9a9a;margin:0 0 40px;line-height:1.6;">This is a temporary password. Please change it after your first login from Settings.</p>

        </div>

        <div style="background:#fafafa;border-top:1px solid #f0f0f0;padding:24px 48px;">
          <p style="font-size:12px;color:#b0b0b0;margin:0;line-height:1.7;">
            Sent by JetLearn Operations System &nbsp;·&nbsp;
            <a href="mailto:${CONFIG.EMAIL.FROM}" style="color:#b0b0b0;text-decoration:none;">${CONFIG.EMAIL.FROM}</a><br>
            If you didn't expect this email, you can safely ignore it.
          </p>
        </div>

      </div>
    </div>
    `;

    MailApp.sendEmail({
      to: toEmail,
      subject: subject,
      htmlBody: htmlBody,
      name: CONFIG.EMAIL.FROM_NAME,
      replyTo: CONFIG.EMAIL.FROM
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
    if (!actingUser || actingUser.role !== 'Super Admin') {
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

function getUserProfile(username) {
  try {
    const profiles = getUserProfiles();
    const user = profiles.find(u => u.username === username);
    if (!user) return null;
    return { username: user.username, email: user.email, role: user.role, lastLogin: user.lastLogin };
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

    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][usernameCol]).toLowerCase() === data.username.toLowerCase()) {

        // Verify current password if changing password
        if (data.newPassword) {
          if (String(allData[i][passwordCol]) !== data.currentPassword) {
            return { success: false, message: 'Current password is incorrect.' };
          }
          sheet.getRange(i + 1, passwordCol + 1).setValue(data.newPassword);
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
    const data = _getCachedSheetData(CONFIG.SHEETS.EMAIL_LOGS);
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
    const sheet = getOrCreateSheet(CONFIG.SHEETS.AUDIT_LOG);
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
    if (!actingUser || actingUser.role !== 'Super Admin') {
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
