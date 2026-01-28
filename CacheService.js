const CACHE_SHEET_NAME = 'DashboardCache';

/**
 * TRIGGER FUNCTION
 * Run this every 30 minutes to update the hidden cache sheet.
 */
function updateDashboardCache() {
  Logger.log("Starting Dashboard Cache Update...");
  
  // 1. Run the slow calculation (from ReportService.js)
  const stats = calculateDashboardStats();
  
  // 2. Convert to string
  const jsonStats = JSON.stringify(stats);
  
  // 3. Save to Sheet
  const sheet = getOrCreateSheet(CACHE_SHEET_NAME);
  
  // Ensure headers exist
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 2).setValues([['Last Updated', 'JSON_Data']]);
  }
  
  // Write Data: A2 = Time, B2 = Data
  sheet.getRange('A2').setValue(new Date());
  sheet.getRange('B2').setValue(jsonStats);
  
  Logger.log("Dashboard Cache Updated Successfully.");
}

/**
 * PUBLIC FUNCTION
 * This is what the frontend calls now. It reads the fast cache.
 */
function getDashboardStatistics() {
  try {
    const sheet = getOrCreateSheet(CACHE_SHEET_NAME);
    const dataRange = sheet.getRange('B2');
    const jsonStats = dataRange.getValue();

    // If cache is empty (first run), fallback to slow calculation
    if (!jsonStats || jsonStats === "") {
      Logger.log("Cache miss. Calculating live (slow)...");
      const liveStats = calculateDashboardStats();
      // Optional: Update cache immediately
      updateDashboardCache(); 
      return liveStats;
    }

    // Return fast data
    return JSON.parse(jsonStats);

  } catch (e) {
    Logger.log("Error reading cache: " + e.message);
    // If anything breaks, fallback to slow calculation so app doesn't crash
    return calculateDashboardStats();
  }
}