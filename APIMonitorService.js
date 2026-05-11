// =============================================
// FILE: APIMonitorService.js
// =============================================

const API_USAGE_STATS_PREFIX = 'API_USAGE_STATS';
const API_FAILURE_LOG_PREFIX = 'API_FAILURE_LOG';
const API_FAILURE_LOG_LIMIT = 50;
const API_MONITOR_TTL_SECONDS = 172800;
const API_WRITE_RETRY_COUNT = 3;

function getApiStatsDate_() {
  const timeZone = Session.getScriptTimeZone() || 'Asia/Kolkata';
  return Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
}

function getApiUsageStatsKey_(dateKey) {
  return `${API_USAGE_STATS_PREFIX}_${dateKey || getApiStatsDate_()}`;
}

function getApiFailureLogKey_(dateKey) {
  return `${API_FAILURE_LOG_PREFIX}_${dateKey || getApiStatsDate_()}`;
}

function readMonitorCache_(key, fallbackValue) {
  try {
    const raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : fallbackValue;
  } catch (error) {
    Logger.log('[APIMonitor] Failed to read cache ' + key + ': ' + error.message);
    return fallbackValue;
  }
}

function writeMonitorCache_(key, value) {
  CacheService.getScriptCache().put(key, JSON.stringify(value), API_MONITOR_TTL_SECONDS);
}

function getApiFailureLogSheet_() {
  try {
    return getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.API_FAILURE_LOG);
  } catch (error) {
    Logger.log('[APIMonitor] Failed to access API Failure Log sheet: ' + error.message);
    return null;
  }
}

function getApiDailySummarySheet_() {
  try {
    return getOrCreateAppDataSheet(CONFIG.APP_DATA_SHEETS.API_DAILY_SUMMARY);
  } catch (error) {
    Logger.log('[APIMonitor] Failed to access API Daily Summary sheet: ' + error.message);
    return null;
  }
}

function ensureApiFailureLogHeader_(sheet) {
  if (!sheet) return;
  if (sheet.getLastRow() > 0) return;
  sheet.appendRow([
    'Timestamp',
    'Date',
    'Service',
    'Function',
    'HTTP Code',
    'URL',
    'Error Details'
  ]);
}

function ensureApiDailySummaryHeader_(sheet) {
  if (!sheet) return;
  if (sheet.getLastRow() > 0) return;
  sheet.appendRow([
    'Date',
    'Service',
    'Function',
    'Total Calls',
    'Success Calls',
    'Failed Calls',
    'Last Call At',
    'Last Status',
    'Last HTTP Code',
    'Last Error'
  ]);
}

function runWithRetries_(callback, label) {
  let lastError = null;
  for (let attempt = 1; attempt <= API_WRITE_RETRY_COUNT; attempt++) {
    try {
      return callback(attempt);
    } catch (error) {
      lastError = error;
      Utilities.sleep(150 * attempt);
    }
  }
  throw new Error(`${label || 'Operation'} failed after ${API_WRITE_RETRY_COUNT} attempts: ${lastError ? lastError.message : 'Unknown error'}`);
}

function appendPersistentApiFailureLog_(entry) {
  const sheet = getApiFailureLogSheet_();
  if (!sheet) return;

  try {
    runWithRetries_(function() {
      ensureApiFailureLogHeader_(sheet);
      sheet.appendRow([
        entry.timestamp || '',
        getApiStatsDate_(),
        entry.service || '',
        entry.func || '',
        entry.code || 0,
        entry.url || '',
        entry.details || ''
      ]);
    }, 'API failure log append');
    if (typeof _clearAppDataCache === 'function' && CONFIG && CONFIG.APP_DATA_SHEETS) {
      _clearAppDataCache(CONFIG.APP_DATA_SHEETS.API_FAILURE_LOG);
    }
  } catch (error) {
    Logger.log('[APIMonitor] Failed to append API failure log row: ' + error.message);
  }
}

function upsertApiDailySummaryRow_(dateKey, serviceName, functionName, stats) {
  const sheet = getApiDailySummarySheet_();
  if (!sheet) return;

  try {
    runWithRetries_(function() {
      ensureApiDailySummaryHeader_(sheet);
      const lastRow = sheet.getLastRow();
      const targetKey = `${dateKey}__${serviceName}__${functionName}`;
      let rowIndex = 0;

      if (lastRow > 1) {
        const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
        for (let i = 0; i < values.length; i++) {
          const rowKey = `${values[i][0]}__${values[i][1]}__${values[i][2]}`;
          if (rowKey === targetKey) {
            rowIndex = i + 2;
            break;
          }
        }
      }

      const row = [[
        dateKey,
        serviceName,
        functionName,
        stats.total || 0,
        stats.success || 0,
        stats.fail || 0,
        stats.lastCallAt || '',
        stats.lastStatus || '',
        stats.lastResponseCode || 0,
        stats.lastError || ''
      ]];

      if (rowIndex > 0) {
        sheet.getRange(rowIndex, 1, 1, row[0].length).setValues(row);
      } else {
        sheet.appendRow(row[0]);
      }
    }, 'API daily summary upsert');

    if (typeof _clearAppDataCache === 'function' && CONFIG && CONFIG.APP_DATA_SHEETS) {
      _clearAppDataCache(CONFIG.APP_DATA_SHEETS.API_DAILY_SUMMARY);
    }
  } catch (error) {
    Logger.log('[APIMonitor] Failed to upsert API daily summary row: ' + error.message);
  }
}

function withMonitorLock_(callback) {
  const lock = LockService.getScriptLock();
  const timeoutMs = 5000;
  try {
    lock.waitLock(timeoutMs);
    return callback();
  } catch (error) {
    Logger.log('[APIMonitor] Lock timeout or failure: ' + error.message);
    return callback();
  } finally {
    try {
      lock.releaseLock();
    } catch (releaseError) {
      // Ignore release failures when the lock was never acquired.
    }
  }
}

function inferServiceNameFromUrl_(url) {
  const raw = String(url || '').toLowerCase();
  if (!raw) return 'Unknown';
  if (raw.indexOf('hubapi.com') !== -1 || raw.indexOf('hubspot.com') !== -1) return 'HubSpot';
  if (raw.indexOf('slack.com') !== -1) return 'Slack';
  if (raw.indexOf('wati') !== -1) return 'WATI';
  if (raw.indexOf('openai.com') !== -1) return 'OpenAI';
  if (raw.indexOf('googleapis.com') !== -1 || raw.indexOf('generativelanguage') !== -1) return 'Gemini';
  if (raw.indexOf('exchangerate') !== -1) return 'ExchangeRate';
  return 'ExternalAPI';
}

function getCallerFunctionName_() {
  try {
    const helperNames = {
      monitoredFetch: true,
      logApiCall: true,
      getCallerFunctionName_: true,
      safeGetResponsePreview_: true
    };
    const stack = (new Error()).stack || '';
    const lines = stack.split('\n');
    for (let i = 1; i < lines.length; i++) {
      const match = lines[i].match(/at\s+([^\s(]+)/);
      if (!match) continue;
      const functionName = match[1];
      if (!helperNames[functionName]) {
        return functionName;
      }
    }
  } catch (error) {
    Logger.log('[APIMonitor] Failed to resolve caller: ' + error.message);
  }
  return 'anonymous';
}

function safeGetResponsePreview_(response) {
  try {
    if (!response || typeof response.getContentText !== 'function') return '';
    return String(response.getContentText()).substring(0, 500);
  } catch (error) {
    return '';
  }
}

/**
 * Logs a single API call, incrementing counters by service and function.
 * @param {string} serviceName
 * @param {string} functionName
 * @param {boolean} success
 * @param {number} responseCode
 * @param {string} errorDetails
 * @param {string} url
 */
function logApiCall(serviceName, functionName, success, responseCode, errorDetails, url) {
  withMonitorLock_(function() {
    const dateKey = getApiStatsDate_();
    const usageKey = getApiUsageStatsKey_(dateKey);
    const failureKey = getApiFailureLogKey_(dateKey);

    const usage = readMonitorCache_(usageKey, {
      date: dateKey,
      total: 0,
      success: 0,
      fail: 0,
      services: {}
    });

    if (!usage.services[serviceName]) {
      usage.services[serviceName] = {
        total: 0,
        success: 0,
        fail: 0,
        lastCallAt: '',
        lastStatus: 'unknown',
        lastResponseCode: 0,
        lastError: '',
        functions: {}
      };
    }

    const serviceStats = usage.services[serviceName];
    if (!serviceStats.functions[functionName]) {
      serviceStats.functions[functionName] = {
        total: 0,
        success: 0,
        fail: 0,
        lastCallAt: '',
        lastStatus: 'unknown',
        lastResponseCode: 0,
        lastError: ''
      };
    }

    const functionStats = serviceStats.functions[functionName];
    const timestamp = new Date().toISOString();

    usage.total++;
    serviceStats.total++;
    functionStats.total++;

    if (success) {
      usage.success++;
      serviceStats.success++;
      functionStats.success++;
    } else {
      usage.fail++;
      serviceStats.fail++;
      functionStats.fail++;
    }

    serviceStats.lastCallAt = timestamp;
    serviceStats.lastStatus = success ? 'success' : 'fail';
    serviceStats.lastResponseCode = responseCode || 0;
    serviceStats.lastError = success ? '' : String(errorDetails || '');

    functionStats.lastCallAt = timestamp;
    functionStats.lastStatus = success ? 'success' : 'fail';
    functionStats.lastResponseCode = responseCode || 0;
    functionStats.lastError = success ? '' : String(errorDetails || '');

    writeMonitorCache_(usageKey, usage);
    upsertApiDailySummaryRow_(dateKey, serviceName, functionName, functionStats);

    if (!success) {
      const failures = readMonitorCache_(failureKey, []);
      const failureEntry = {
        timestamp: timestamp,
        service: serviceName,
        func: functionName,
        code: responseCode || 0,
        url: String(url || '').substring(0, 250),
        details: String(errorDetails || '').substring(0, 500)
      };
      failures.unshift(failureEntry);
      writeMonitorCache_(failureKey, failures.slice(0, API_FAILURE_LOG_LIMIT));
      appendPersistentApiFailureLog_(failureEntry);
    }
  });
}

function monitoredFetch(url, options, functionName, serviceName) {
  const resolvedUrl = String(url || '');
  const resolvedFunctionName = functionName || getCallerFunctionName_();
  const resolvedServiceName = serviceName || inferServiceNameFromUrl_(resolvedUrl);

  let response;
  let responseCode = 0;
  let success = false;
  let errorDetails = '';

  try {
    response = UrlFetchApp.fetch(url, options || {});
    responseCode = response && typeof response.getResponseCode === 'function'
      ? response.getResponseCode()
      : 0;
    success = responseCode >= 200 && responseCode < 300;
    if (!success) {
      errorDetails = safeGetResponsePreview_(response) || `HTTP ${responseCode}`;
    }
    return response;
  } catch (error) {
    errorDetails = error && error.message ? error.message : String(error);
    const codeMatch = errorDetails.match(/\b(\d{3})\b/);
    responseCode = codeMatch ? Number(codeMatch[1]) : 0;
    throw error;
  } finally {
    logApiCall(
      resolvedServiceName,
      resolvedFunctionName,
      success,
      responseCode,
      errorDetails,
      resolvedUrl
    );
  }
}

function getApiUsageStats(dateKey) {
  return readMonitorCache_(getApiUsageStatsKey_(dateKey), {
    date: dateKey || getApiStatsDate_(),
    total: 0,
    success: 0,
    fail: 0,
    services: {}
  });
}

function getApiFailureLog(dateKey) {
  return readMonitorCache_(getApiFailureLogKey_(dateKey), []);
}

function resetApiStats(dateKey) {
  const resolvedDateKey = dateKey || getApiStatsDate_();
  const cache = CacheService.getScriptCache();
  cache.remove(getApiUsageStatsKey_(resolvedDateKey));
  cache.remove(getApiFailureLogKey_(resolvedDateKey));
  Logger.log('API usage and failure logs have been reset for ' + resolvedDateKey);
}

function testApiFailureLogging() {
  const testTimestamp = new Date().toISOString();
  const entry = {
    timestamp: testTimestamp,
    service: 'MonitorTest',
    func: 'testApiFailureLogging',
    code: 599,
    url: 'https://example.invalid/test-api-monitor',
    details: 'Synthetic API monitor test failure row'
  };

  appendPersistentApiFailureLog_(entry);
  upsertApiDailySummaryRow_(getApiStatsDate_(), entry.service, entry.func, {
    total: 1,
    success: 0,
    fail: 1,
    lastCallAt: testTimestamp,
    lastStatus: 'fail',
    lastResponseCode: entry.code,
    lastError: entry.details
  });

  return {
    success: true,
    message: `Test API failure log written for ${entry.service}/${entry.func}`,
    timestamp: testTimestamp,
    sheetName: CONFIG.APP_DATA_SHEETS.API_FAILURE_LOG
  };
}
