// --- Global Cache for current script execution ---
let _sheetDataCache = {};
let _spreadsheetCache = {};

function _getSpreadsheet(id) {
  if (!_spreadsheetCache[id]) {
    _spreadsheetCache[id] = SpreadsheetApp.openById(id);
  }
  return _spreadsheetCache[id];
}

function _getCachedSheetData(sheetName, spreadsheetId = CONFIG.MIGRATION_SHEET_ID) {
  const cacheKey = `${spreadsheetId}_${sheetName}`;
  if (!_sheetDataCache[cacheKey]) {
    try {
      const spreadsheet = _getSpreadsheet(spreadsheetId);
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        Logger.log(`Warning: Sheet '${sheetName}' not found in spreadsheet ID: ${spreadsheetId}. Returning empty array.`);
        _sheetDataCache[cacheKey] = []; // Cache empty result to avoid repeated lookups
        return [];
      }
      _sheetDataCache[cacheKey] = sheet.getDataRange().getValues();
    } catch (e) {
      Logger.log(`Error fetching cached data for ${sheetName}: ${e.message}`);
      _sheetDataCache[cacheKey] = []; 
      return []; // Return empty array to prevent "No data returned" crash
    }
  }
  return _sheetDataCache[cacheKey];
}

function getOrCreateSheet(sheetName) {
  try {
    const spreadsheet = _getSpreadsheet(CONFIG.MIGRATION_SHEET_ID);
    let sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      Logger.log('Created new sheet: ' + sheetName);
    }

    return sheet;
  } catch (error) {
    Logger.log('Error getting/creating sheet: ' + error.message);
    throw error;
  }
}

//AI Wrapper:
function callGenerativeAIWithRetry(endpoint, payload) {
  const MAX_RETRIES = 3;
  let lastError = null;

  for (let i = 0; i < MAX_RETRIES; i++) {
    try {
      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(endpoint, options);
      const responseCode = response.getResponseCode();

      if (responseCode === 200) {
        return response; // Success
      } else if (responseCode === 429) {
        Logger.log(`[AI API] Rate limit hit (429). Retry ${i + 1}/${MAX_RETRIES}...`);
        lastError = new Error(`API rate limit exceeded. Response: ${response.getContentText()}`);
        Utilities.sleep((i + 1) * 1000); // Wait 1s, then 2s, etc.
      } else {
        throw new Error(`AI API Error (${responseCode}): ${response.getContentText()}`);
      }
    } catch (e) {
      lastError = e;
      Logger.log(`[AI API] Connection error on retry ${i + 1}/${MAX_RETRIES}: ${e.message}`);
      Utilities.sleep((i + 1) * 1000); // Also wait on connection errors
    }
  }
  
  Logger.log(`[AI API] Final failure after ${MAX_RETRIES} retries.`);
  throw lastError; // Throw the last captured error after all retries fail
}

//Formatting & Validation:
function formatDate(dateObj) {
  if (!dateObj || isNaN(dateObj.getTime())) return 'N/A';
  const options = { year: 'numeric', month: 'long', day: 'numeric' };
  return dateObj.toLocaleDateString('en-GB', options);
}


function formatDateDDMMYYYY(dateInput) {
  try {
    let dateToFormat;
    
    if (dateInput instanceof Date && !isNaN(dateInput)) {
      dateToFormat = dateInput;
    } 
    else if (typeof dateInput === 'string') {
      dateToFormat = new Date(dateInput);
      if (isNaN(dateToFormat.getTime()) && /^\d{4}-\d{2}-\d{2}/.test(dateInput)) {
        const parts = dateInput.split('T')[0].split('-');
        dateToFormat = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
      }
    }

    if (!dateToFormat || isNaN(dateToFormat.getTime())) {
      Logger.log('Could not format invalid dateInput: ' + dateInput);
      return 'N/A';
    }

    const day = String(dateToFormat.getUTCDate()).padStart(2, '0');
    const month = String(dateToFormat.getUTCMonth() + 1).padStart(2, '0');
    const year = dateToFormat.getUTCFullYear();

    return `${day}/${month}/${year}`;
  } catch (e) {
    Logger.log('Error formatting date: ' + e.message);
    return 'N/A';
  }
}

function parseSheetDate(sheetDate) {
  if (sheetDate instanceof Date && !isNaN(sheetDate)) return sheetDate;

  const directDate = new Date(sheetDate);
  if (directDate instanceof Date && !isNaN(directDate)) return directDate;
  
  if (typeof sheetDate === 'string') {
    const datePart = sheetDate.split(' ')[0].replace(/,/g, '');
    const parts = datePart.split(/[\/-]/); 
    if (parts.length === 3) {
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1; 
      let year = parseInt(parts[2], 10);

      if (year < 100) {
        year += 2000;
      }
      
      if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
          const customDate = new Date(Date.UTC(year, month, day)); 
          if (sheetDate.includes(' ')) {
              const timeString = sheetDate.split(' ')[1];
              if (timeString && timeString.includes(':')) {
                  const timeParts = timeString.split(':');
                  if (timeParts.length >= 2) {
                      customDate.setUTCHours(parseInt(timeParts[0], 10) || 0);
                      customDate.setUTCMinutes(parseInt(timeParts[1], 10) || 0);
                      if (timeParts.length === 3) {
                        customDate.setUTCSeconds(parseInt(timeParts[2], 10) || 0);
                      }
                  }
              }
          }
          if (!isNaN(customDate)) return customDate;
      }
    }
  }
  return null;
}
function isValidEmail(email) {
  if (!email || typeof email !== 'string') return false;
  const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  return emailRegex.test(email.trim());
}

function validateMigrationData(data) {
  const errors = [];

  if (!data.jlid || data.jlid.trim() === '') {
    errors.push('JLID is required');
  }

  if (!data.learner || data.learner.trim() === '') {
    errors.push('Learner name is required');
  }

  if (!data.newTeacher || data.newTeacher.trim() === '') {
    errors.push('New teacher is required');
  }

  if (!data.course || data.course.trim() === '') {
    errors.push('Course is required');
  }

  if (!data.classSessions || data.classSessions.length === 0) {
    errors.push('At least one class session (Day and Time) is required');
  } else {
    data.classSessions.forEach((session, index) => {
      if (!session.day || session.day.trim() === '') {
        errors.push(`Class Day for session ${index + 1} is required`);
      }
      if (!session.time || !session.time.match(/^\d{1,2}:\d{2}\s(AM|PM)$/i)) {
        errors.push(`Class Time for session ${index + 1} is invalid or missing. Expected format HH:MM AM/PM`);
      }
    });
  }

  if (!data.clsManager || data.clsManager.trim() === '') {
    errors.push('CLS Manager is required');
  }

  if (!data.jetGuide || data.jetGuide.trim() === '') {
    errors.push('JetGuide is required');
  }

  if (!data.startDate || data.startDate.trim() === '') {
    errors.push('Start Date is required');
  }

  if (!data.migrationType || !['Mid-Course', 'New Assignment'].includes(data.migrationType)) {
    errors.push('Valid migration type is required');
  }

  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

function getImagesAsBase64FromDrive(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const imageData = {};

    while (files.hasNext()) {
      const file = files.next();
      const blob = file.getBlob();
      const contentType = blob.getContentType();
      const base64Data = Utilities.base64Encode(blob.getBytes());
      
      // Construct the full Data URI for embedding in HTML
      imageData[file.getName()] = `data:${contentType};base64,${base64Data}`;
    }
    
    Logger.log(`Successfully processed ${Object.keys(imageData).length} images from Drive.`);
    return imageData;

  } catch (e) {
    Logger.log(`Error accessing Drive folder or processing images: ${e.message}`);
    return {}; // Return empty object on error
  }
}


function cleanJlidForZoom(jlid) {
  if (!jlid || typeof jlid !== 'string') return '';
  // This regex now removes a trailing 'C' or 'M' followed by any number of digits.
  return jlid.trim().toUpperCase().replace(/(C|M)\d*$/, '');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
