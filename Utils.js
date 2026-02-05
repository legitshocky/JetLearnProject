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


function backupDatabase() {
  const BACKUP_FOLDER_ID = '1LzKa1U1-ou6fsHzIh35jO7z5WkOydSRn'; // Create a folder and put ID here
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const fileName = `JetLearn_DB_Backup_${timestamp}`;
  
  const originalFile = DriveApp.getFileById(CONFIG.MIGRATION_SHEET_ID);
  const backupFolder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
  
  originalFile.makeCopy(fileName, backupFolder);
  Logger.log("Backup completed: " + fileName);
}

function logError(context, errorObject) {
  try {
    const sheetName = "System Errors";
    const sheet = getOrCreateSheet(sheetName);
    
    // Add headers if the sheet is new
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "User", "Context", "Message", "Stack Trace"]);
      sheet.setFrozenRows(1);
    }

    const timestamp = new Date();
    const user = Session.getActiveUser().getEmail();
    const message = errorObject.message || String(errorObject);
    const stack = errorObject.stack || "N/A";

    // Append the error details
    sheet.appendRow([timestamp, user, context, message, stack]);
    
    // Force save
    SpreadsheetApp.flush();

    // Console log as backup
    Logger.log(`[${context}] Error logged to sheet: ${message}`);

  } catch (e) {
    // If logging fails, fall back to console only
    Logger.log(`FAILED TO LOG ERROR to sheet: ${e.message}`);
    Logger.log(`Original Error (${context}): ${errorObject.message}`);
  }
}
const TIMEZONE_FRIENDLY_LABELS = {
  // Asia
  'Asia/Kolkata': 'IST',
  'Asia/Dubai': 'GST',
  'Asia/Singapore': 'SGT',
  'Asia/Bangkok': 'ICT',
  'Asia/Hong_Kong': 'HKT',
  'Asia/Tokyo': 'JST',
  'Asia/Seoul': 'KST',
  
  // Europe
  'Europe/London': 'UK Time',
  'Europe/Berlin': 'CET',
  'Europe/Paris': 'CET',
  'Europe/Amsterdam': 'CET',
  'Europe/Zurich': 'CET',
  'Europe/Athens': 'EET',
  'Europe/Moscow': 'MSK',
  
  // US/Americas
  'America/New_York': 'EST', 
  'America/Chicago': 'CST',
  'America/Denver': 'MST',
  'America/Phoenix': 'MST',
  'America/Los_Angeles': 'PST', 
  'America/Anchorage': 'AKST',
  'Pacific/Honolulu': 'HST',
  
  // Australia/Pacific
  'Australia/Sydney': 'AEST',
  'Australia/Melbourne': 'AEST',
  'Australia/Brisbane': 'AEST',
  'Australia/Adelaide': 'ACST',
  'Australia/Perth': 'AWST',
  'Pacific/Auckland': 'NZT'
};

function convertCetToLocal(timeStr, targetTzString) {
  try {
    if (!timeStr || timeStr === "TBD") return "TBD";

    // 1. Get the friendly label based on the dropdown selection text
    let label = ""; // Default to empty if no match found
    const tz = targetTzString || "";

    // --- Asia ---
    if (tz.includes("India") || tz.includes("Kolkata") || tz.includes("Chennai") || tz.includes("Mumbai") || tz.includes("New Delhi")) label = "IST";
    else if (tz.includes("Dubai") || tz.includes("Muscat") || tz.includes("Abu Dhabi")) label = "GST";
    else if (tz.includes("Singapore") || tz.includes("Kuala Lumpur")) label = "SGT";
    else if (tz.includes("Bangkok") || tz.includes("Hanoi") || tz.includes("Jakarta")) label = "ICT";
    else if (tz.includes("Hong Kong") || tz.includes("Beijing")) label = "HKT";
    else if (tz.includes("Tokyo") || tz.includes("Osaka")) label = "JST";
    else if (tz.includes("Seoul")) label = "KST";

    // --- Europe ---
    else if (tz.includes("London") || tz.includes("Dublin") || tz.includes("Edinburgh") || tz.includes("Lisbon")) label = "UK Time";
    else if (tz.includes("Brussels") || tz.includes("Paris") || tz.includes("Amsterdam") || tz.includes("Berlin") || tz.includes("Madrid") || tz.includes("Rome") || tz.includes("Vienna") || tz.includes("Stockholm")) label = "CET";
    else if (tz.includes("Athens") || tz.includes("Bucharest") || tz.includes("Cairo") || tz.includes("Jerusalem")) label = "EET";
    else if (tz.includes("Moscow")) label = "MSK";

    // --- US/Americas ---
    else if (tz.includes("Eastern") || tz.includes("New York") || tz.includes("Indiana")) label = "EST";
    else if (tz.includes("Central Time") || tz.includes("Chicago") || tz.includes("Mexico City")) label = "CST";
    else if (tz.includes("Mountain") || tz.includes("Denver") || tz.includes("Arizona")) label = "MST";
    else if (tz.includes("Pacific") || tz.includes("Los Angeles") || tz.includes("Tijuana")) label = "PST";
    else if (tz.includes("Alaska")) label = "AKST";
    else if (tz.includes("Hawaii")) label = "HST";
    else if (tz.includes("Brasilia") || tz.includes("Buenos Aires")) label = "BRT";

    // --- Australia/Pacific ---
    else if (tz.includes("Sydney") || tz.includes("Melbourne") || tz.includes("Canberra") || tz.includes("Brisbane")) label = "AEST";
    else if (tz.includes("Adelaide") || tz.includes("Darwin")) label = "ACST";
    else if (tz.includes("Perth")) label = "AWST";
    else if (tz.includes("Auckland") || tz.includes("Wellington")) label = "NZT";

    // 2. Return Input + Label (NO MATH)
    // If we found a label, append it. If not, just return the time.
    return label ? `${timeStr} ${label}` : timeStr;

  } catch (e) {
    return timeStr;
  }
}

