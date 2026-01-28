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

    Logger.log('Authentication successful for user: ' + username + ', role: ' + user.role);
    return {
      success: true,
      role: user.role,
      username: username,
      permissions: PERMISSIONS[user.role] || []
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
        createdDate: user.createddate
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

function addNewUser(userData) {
  Logger.log('addNewUser called for: ' + userData.username);

  try {
    const sheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID).getSheetByName(CONFIG.SHEETS.USER_PROFILES);

    const existingUsers = getUserProfiles(); 
    if (existingUsers.some(u => u.username.toLowerCase() === userData.username.toLowerCase())) {
      return { success: false, message: 'Username already exists' };
    }

    if (!isValidEmail(userData.email)) {
      return { success: false, message: 'Invalid email address' };
    }

    if (!Object.values(ROLES).includes(userData.role)) {
      return { success: false, message: 'Invalid user role' };
    }

    sheet.appendRow([
      userData.username,
      userData.password,
      userData.role,
      userData.email,
      true, 
      '',   
      new Date(), 
      '',   
      ''    
    ]);

    logAction('User Added', '', '', '', '', '', 'Success', `New ${userData.role} user added: ${userData.username}`);

    delete _sheetDataCache[`${CONFIG.MIGRATION_SHEET_ID}_${CONFIG.SHEETS.USER_PROFILES}`];

    return { success: true, message: 'User added successfully' };
  } catch (error) {
    Logger.log('Error adding new user: ' + error.message);
    return { success: false, message: 'Error adding user: ' + error.message };
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
