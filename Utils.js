// --- Global Cache for current script execution ---
let _sheetDataCache = {};
let _spreadsheetCache = {};

function _getSpreadsheet(id) {
  if (!_spreadsheetCache[id]) {
    try {
      _spreadsheetCache[id] = SpreadsheetApp.openById(id);
    } catch(e) {
      Logger.log('[_getSpreadsheet] Cannot open spreadsheet ' + id + ': ' + e.message);
      return null; // return null so callers can handle gracefully
    }
  }
  return _spreadsheetCache[id];
}

// Sheets listed here are also stored in GAS CacheService (survives across executions).
// TTL: 10 minutes. This lets a warm-up in one GAS call benefit the next call.
var _GAS_CACHE_SHEETS = ['Teacher Courses'];
var _GAS_CACHE_TTL    = 600; // seconds

function _getCachedSheetData(sheetName, spreadsheetId = CONFIG.MIGRATION_SHEET_ID) {
  const cacheKey = `${spreadsheetId}_${sheetName}`;

  // 1. In-memory cache (within same execution)
  if (_sheetDataCache[cacheKey]) return _sheetDataCache[cacheKey];

  // 2. GAS CacheService (across executions, up to 10 min)
  if (_GAS_CACHE_SHEETS.indexOf(sheetName) !== -1) {
    try {
      var gasCache = CacheService.getScriptCache();
      var cached   = gasCache.get(cacheKey);
      if (cached) {
        var parsed = JSON.parse(cached);
        _sheetDataCache[cacheKey] = parsed;
        Logger.log('[_getCachedSheetData] GAS cache HIT for ' + sheetName);
        return parsed;
      }
    } catch(ce) {
      Logger.log('[_getCachedSheetData] GAS cache read error: ' + ce.message);
    }
  }

  // 3. Cold sheet read
  try {
    const spreadsheet = _getSpreadsheet(spreadsheetId);
    if (!spreadsheet) { _sheetDataCache[cacheKey] = []; return []; }
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Warning: Sheet '${sheetName}' not found in spreadsheet ID: ${spreadsheetId}. Returning empty array.`);
      _sheetDataCache[cacheKey] = [];
      return [];
    }
    var data = sheet.getDataRange().getValues();
    _sheetDataCache[cacheKey] = data;

    // Store in GAS CacheService if applicable (max value size 100KB; skip if too large)
    if (_GAS_CACHE_SHEETS.indexOf(sheetName) !== -1) {
      try {
        var serialized = JSON.stringify(data);
        if (serialized.length < 90000) {
          CacheService.getScriptCache().put(cacheKey, serialized, _GAS_CACHE_TTL);
          Logger.log('[_getCachedSheetData] GAS cache SET for ' + sheetName + ' (' + serialized.length + ' bytes)');
        } else {
          Logger.log('[_getCachedSheetData] ' + sheetName + ' too large for GAS cache (' + serialized.length + ' bytes) — skipped');
        }
      } catch(se) {
        Logger.log('[_getCachedSheetData] GAS cache write error: ' + se.message);
      }
    }
    return data;
  } catch (e) {
    Logger.log(`Error fetching cached data for ${sheetName}: ${e.message}`);
    _sheetDataCache[cacheKey] = [];
    return [];
  }
}

// Call this during JLID load to pre-warm the Teacher Courses cache so the
// upskilling check on teacher selection is instant (no cold sheet read).
function warmTeacherCoursesCache() {
  try {
    _getCachedSheetData(CONFIG.SHEETS.TEACHER_COURSES);
    Logger.log('[warmTeacherCoursesCache] Done.');
  } catch(e) {
    Logger.log('[warmTeacherCoursesCache] Error: ' + e.message);
  }
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


// ── Live progress of a long run (migration etc.): steps are kept in the script cache so the page can poll them ──
var __PROG_KEY = '';
function _progWrite(key, mut) {
  try {
    var cache = CacheService.getScriptCache();
    var cur = cache.get(key);
    var o = cur ? JSON.parse(cur) : { steps: [], done: false };
    mut(o);
    cache.put(key, JSON.stringify(o), 1800);
  } catch (e) { /* progress is a nice-to-have, never break the run */ }
}
function _progBegin(label) {
  if (!__PROG_KEY) return;
  _progWrite(__PROG_KEY, function(o) { o.steps.push({ label: label, status: 'running', at: Date.now() }); });
}
function _progEnd(label, ok, detail) {
  if (!__PROG_KEY) return;
  _progWrite(__PROG_KEY, function(o) {
    for (var i = o.steps.length - 1; i >= 0; i--) {
      if (o.steps[i].label === label && o.steps[i].status === 'running') {
        o.steps[i].status = ok ? 'success' : 'failed';
        o.steps[i].ms = Date.now() - o.steps[i].at;
        if (detail) o.steps[i].detail = String(detail).slice(0, 300);
        return;
      }
    }
    o.steps.push({ label: label, status: ok ? 'success' : 'failed', at: Date.now(), ms: 0, detail: detail ? String(detail).slice(0, 300) : '' });
  });
}
function getMigrationProgress(jlid) {
  try {
    var cur = CacheService.getScriptCache().get('migprog_' + String(jlid || '').trim());
    if (!cur) return { success: true, steps: [], done: false, none: true };
    var o = JSON.parse(cur);
    return { success: true, steps: o.steps || [], done: !!o.done, result: o.result || '' };
  } catch (e) { return { success: false, message: e.message }; }
}

function _timelineStep(timeline, key, label, fn) {
  var started = new Date().getTime();
  _progBegin(label);
  try {
    var value = fn();
    _progEnd(label, true);
    var step = { key: key, label: label, status: 'success', durationMs: new Date().getTime() - started };
    timeline.push(step);
    // caller can set step.meta after we return — expose the step via timeline[timeline.length-1]
    return value;
  } catch (e) {
    _progEnd(label, false, e && e.message ? e.message : String(e));
    timeline.push({
      key: key,
      label: label,
      status: 'failed',
      durationMs: new Date().getTime() - started,
      detail: e && e.message ? e.message : String(e)
    });
    throw e;
  }
}

function _timelineAdd(timeline, key, label, status, started, detail, meta) {
  timeline.push({
    key: key,
    label: label,
    status: status || 'success',
    durationMs: started ? (new Date().getTime() - started) : 0,
    detail: detail || '',
    meta: meta || null
  });
}

// Opens (or creates) a sheet in the JetLearn App Data spreadsheet
function getOrCreateAppDataSheet(sheetName) {
  try {
    const spreadsheet = _getSpreadsheet(CONFIG.APP_DATA_SHEET_ID);
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      Logger.log('Created App Data sheet: ' + sheetName);
    }
    return sheet;
  } catch (error) {
    Logger.log('Error getting/creating App Data sheet: ' + error.message);
    throw error;
  }
}

// Cached read from JetLearn App Data spreadsheet
function _getAppDataCachedSheetData(sheetName) {
  return _getCachedSheetData(sheetName, CONFIG.APP_DATA_SHEET_ID);
}

// Clears the in-memory cache for an App Data sheet
function _clearAppDataCache(sheetName) {
  if (typeof _sheetDataCache !== 'undefined') {
    delete _sheetDataCache[CONFIG.APP_DATA_SHEET_ID + '_' + sheetName];
  }
}

//AI Wrapper:
// AI response cache TTL in seconds (6 hours)
var AI_CACHE_TTL = 21600;

function callGenerativeAIWithRetry(endpoint, payload) {
  const MAX_RETRIES = 3;

  // ── Cache check ──────────────────────────────────────────────────────────
  // Key = digest of endpoint + stringified payload (same prompt = same key)
  var cacheKey = 'ai_' + Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    endpoint + JSON.stringify(payload)
  ).map(function(b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');

  try {
    var scriptCache = CacheService.getScriptCache();
    var cached = scriptCache.get(cacheKey);
    if (cached) {
      Logger.log('[AI Cache] HIT — returning cached response, no token used.');
      // Return a mock response object with the cached text
      return { getResponseCode: function() { return 200; }, getContentText: function() { return cached; } };
    }
  } catch(cacheErr) {
    Logger.log('[AI Cache] Cache read error: ' + cacheErr.message);
  }
  // ─────────────────────────────────────────────────────────────────────────

  var lastError = null;
  for (var i = 0; i < MAX_RETRIES; i++) {
    try {
      var options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      var response = monitoredFetch(endpoint, options);
      var responseCode = response.getResponseCode();

      if (responseCode === 200) {
        // Store result in cache before returning
        try {
          var responseText = response.getContentText();
          CacheService.getScriptCache().put(cacheKey, responseText, AI_CACHE_TTL);
          Logger.log('[AI Cache] MISS — response cached for ' + (AI_CACHE_TTL / 3600) + 'h.');
        } catch(putErr) {
          Logger.log('[AI Cache] Cache write error: ' + putErr.message);
        }
        return response;
      } else if (responseCode === 429) {
        Logger.log('[AI API] Rate limit hit (429). Retry ' + (i+1) + '/' + MAX_RETRIES + '...');
        lastError = new Error('API rate limit exceeded. Response: ' + response.getContentText());
        Utilities.sleep((i + 1) * 1000);
      } else {
        throw new Error('AI API Error (' + responseCode + '): ' + response.getContentText());
      }
    } catch(e) {
      lastError = e;
      Logger.log('[AI API] Connection error on retry ' + (i+1) + '/' + MAX_RETRIES + ': ' + e.message);
      Utilities.sleep((i + 1) * 1000);
    }
  }

  Logger.log('[AI API] Final failure after ' + MAX_RETRIES + ' retries.');
  throw lastError;
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

// Maps HubSpot timezone dropdown text → IANA timezone ID
function _resolveIana(tzString) {
  var tz = tzString || '';
  if (tz.includes('India') || tz.includes('Kolkata') || tz.includes('Chennai') || tz.includes('Mumbai') || tz.includes('New Delhi')) return 'Asia/Kolkata';
  if (tz.includes('Dubai') || tz.includes('Muscat') || tz.includes('Abu Dhabi')) return 'Asia/Dubai';
  if (tz.includes('Singapore') || tz.includes('Kuala Lumpur')) return 'Asia/Singapore';
  if (tz.includes('Bangkok') || tz.includes('Hanoi') || tz.includes('Jakarta')) return 'Asia/Bangkok';
  if (tz.includes('Hong Kong') || tz.includes('Beijing')) return 'Asia/Hong_Kong';
  if (tz.includes('Tokyo') || tz.includes('Osaka')) return 'Asia/Tokyo';
  if (tz.includes('Seoul')) return 'Asia/Seoul';
  if (tz.includes('London') || tz.includes('Dublin') || tz.includes('Edinburgh') || tz.includes('Lisbon')) return 'Europe/London';
  if (tz.includes('Brussels') || tz.includes('Paris') || tz.includes('Amsterdam') || tz.includes('Berlin') || tz.includes('Madrid') || tz.includes('Rome') || tz.includes('Vienna') || tz.includes('Stockholm')) return 'Europe/Paris';
  if (tz.includes('Athens') || tz.includes('Bucharest') || tz.includes('Cairo') || tz.includes('Jerusalem')) return 'Europe/Athens';
  if (tz.includes('Moscow')) return 'Europe/Moscow';
  if (tz.includes('Eastern') || tz.includes('New York') || tz.includes('Indiana')) return 'America/New_York';
  if (tz.includes('Central Time') || tz.includes('Chicago') || tz.includes('Mexico City')) return 'America/Chicago';
  if (tz.includes('Mountain') || tz.includes('Denver') || tz.includes('Arizona')) return 'America/Denver';
  if (tz.includes('Pacific') || tz.includes('Los Angeles') || tz.includes('Tijuana')) return 'America/Los_Angeles';
  if (tz.includes('Alaska')) return 'America/Anchorage';
  if (tz.includes('Hawaii')) return 'Pacific/Honolulu';
  if (tz.includes('Brasilia') || tz.includes('Buenos Aires')) return 'America/Sao_Paulo';
  if (tz.includes('Sydney') || tz.includes('Melbourne') || tz.includes('Canberra') || tz.includes('Brisbane')) return 'Australia/Sydney';
  if (tz.includes('Adelaide') || tz.includes('Darwin')) return 'Australia/Adelaide';
  if (tz.includes('Perth')) return 'Australia/Perth';
  if (tz.includes('Auckland') || tz.includes('Wellington')) return 'Pacific/Auckland';
  return null;
}

function _tzLabel(tzString) {
  var iana = _resolveIana(tzString);
  if (!iana) return '';
  var map = {
    'Asia/Kolkata': 'IST', 'Asia/Dubai': 'GST', 'Asia/Singapore': 'SGT',
    'Asia/Bangkok': 'ICT', 'Asia/Hong_Kong': 'HKT', 'Asia/Tokyo': 'JST', 'Asia/Seoul': 'KST',
    'Europe/London': 'UK Time', 'Europe/Paris': 'CET', 'Europe/Athens': 'EET', 'Europe/Moscow': 'MSK',
    'America/New_York': 'ET', 'America/Chicago': 'CT', 'America/Denver': 'MT',
    'America/Los_Angeles': 'PT', 'America/Anchorage': 'AKT', 'Pacific/Honolulu': 'HST',
    'America/Sao_Paulo': 'BRT',
    'Australia/Sydney': 'AEST', 'Australia/Adelaide': 'ACST', 'Australia/Perth': 'AWST',
    'Pacific/Auckland': 'NZT'
  };
  return map[iana] || '';
}

// convertCetToLocal(timeStr, targetTzString)
// Appends the parent's timezone label to the time string.
// Time is already stored in the parent's local timezone — no conversion needed.
function convertCetToLocal(timeStr, targetTzString) {
  try {
    if (!timeStr || timeStr === "TBD") return "TBD";

    // Strip secondary suffix e.g. "6 PM;3 PM UK Time" -> "6 PM"
    timeStr = String(timeStr).split(';')[0].trim();

    var tzLabel = _tzLabel(targetTzString);
    return tzLabel ? timeStr + ' ' + tzLabel : timeStr;

  } catch (e) {
    return timeStr;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// SAFE LOGGING — never logs sensitive values in full
// ─────────────────────────────────────────────────────────────────────────────

function _maskValue(value, visibleChars) {
  if (!value) return '[empty]';
  var str = String(value);
  var show = visibleChars || 4;
  if (str.length <= show) return '****';
  return str.substring(0, show) + '...[masked]';
}
