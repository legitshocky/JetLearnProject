/**
 * DEBUG: Tests ONLY the 3 fixed templates.
 */
function debugTestSpecificTemplates() {
  // --- 1. ENTER YOUR PHONE NUMBER HERE ---
  const MY_PHONE = "918369118156"; // <--- REPLACE THIS (No + symbol)
  // ---------------------------------------

  const TARGET_TEMPLATES = [
    "migration_teacher_affinity"
  ];

  Logger.log("🚀 Starting Specific Template Test...");
  
  // 2. Mock Data (Ensuring all variables like {{Course}}, {{Date}}, {{Weekday}} are present)
  const mockMigration = {
    learner: "Robin",
    newTeacher: "Batman", 
    oldTeacher: "Alfred",
    course: "Crime Fighting 101",
    classLink: "https://meet.google.com/bat-cave",
    startDate: "2026-02-01", // For {{Date}}
    classSessions: [{ day: "Monday", time: "08:00 PM" }], // For {{Weekday}}
    manualTimezone: "Asia/Kolkata",
    calculatedLocalTime: "10:00 AM IST" // For {{Time}}
  };

  const mockHubSpot = {
    parentName: "Sourav Pal",
    parentContact: MY_PHONE,
    timezone: "Asia/Kolkata"
  };

  // 3. Iterate and Send
  TARGET_TEMPLATES.forEach(templateId => {
      Logger.log(`\n---------------------------------`);
      Logger.log(`Testing: ${templateId}`);

      try {
        // Generate Params using the UPDATED logic
        const params = getWatiParameters(templateId, mockMigration, mockHubSpot);
        
        Logger.log("Params Generated: " + JSON.stringify(params));

        // Send
        const res = sendWatiMessage(MY_PHONE, templateId, params);
        
        if (res.success) {
          Logger.log(`✅ SUCCESS`);
        } else {
          Logger.log(`❌ FAILED`);
        }

      } catch (e) {
        Logger.log(`🚨 CRASH: ${e.message}`);
      }

      Utilities.sleep(1000);
  });

  Logger.log("\n✅ Done.");
}