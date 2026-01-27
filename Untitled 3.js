// =================================================================
// START: PASTE THIS NEW FUNCTION INTO YOUR CODE.JS FILE
// =================================================================

/**
 * A standalone test function to send a specific WATI template message.
 * @param {string} jlid The JetLearn ID of the learner whose parent you want to message.
 */
function sendTestWatiMigrationMessage(jlid) {
  Logger.log(`--- Starting WATI Test for JLID: ${jlid} ---`);

  if (!jlid || typeof jlid !== 'string' || jlid.trim() === '') {
    Logger.log("ERROR: You must provide a valid JLID to run this test.");
    throw new Error("JLID cannot be empty. Please provide a valid JLID.");
  }

  try {
    // 1. Fetch Migration and Contact Details from HubSpot
    Logger.log("Step 1: Fetching data from HubSpot...");
    const hubspotInfo = fetchHubspotByJlid(jlid);

    if (!hubspotInfo.success) {
      throw new Error(`HubSpot Fetch Failed: ${hubspotInfo.message}`);
    }

    const parentName = hubspotInfo.data.parentName;
    const parentContact = hubspotInfo.data.parentContact;

    if (!parentContact) {
      throw new Error("HubSpot data is missing the parent's phone number ('parentContact'). Cannot send message.");
    }
    Logger.log(`Successfully fetched Parent Name: "${parentName}" and Contact: "${parentContact}"`);

    // 2. Define the Template and its Parameters
    Logger.log("Step 2: Preparing WATI template and parameters...");
    const templateName = "outbound_msg";
    
    // IMPORTANT: The 'name' must EXACTLY match the variable name in your WATI template body (e.g., {{parent}}, {{text}})
    const parameters = [
      {
        "name": "parent",
        "value": parentName || "there" // Use "there" as a fallback if parent name is missing
      },
      {
        "name": "text",
        "value": "This is a test message from the JetLearn Migration System."
      },
      {
        "name": "text1",
        "value": "Please confirm if you have received this. No other action is required."
      }
    ];
    Logger.log(`Template set to "${templateName}" with parameters: ${JSON.stringify(parameters)}`);

    // 3. Send the WhatsApp Message via WATI
    Logger.log("Step 3: Calling sendWatiMessage function...");
    const result = sendWatiMessage(parentContact, templateName, parameters);
    
    Logger.log("--- SUCCESS: WATI Test Completed ---");
    Logger.log(`Message sent successfully to ${parentContact}. WATI Response: ${JSON.stringify(result)}`);
    
    return { success: true, message: `Test message sent to ${parentContact}.` };

  } catch (error) {
    Logger.log(`--- ERROR: WATI Test Failed for JLID: ${jlid} ---`);
    Logger.log(error.message);
    // Re-throw the error so it's clearly visible in the execution logs
    throw error;
  }
}

function runMyTest() {
  // ----> REPLACE "JL12345" WITH THE REAL JLID YOU WANT TO TEST <----
  sendTestWatiMigrationMessage("JL39611449152C2"); 
}


// =================================================================
// START: PASTE THIS TEST FUNCTION INTO YOUR CODE.JS FILE
// =================================================================

/**
 * A simple, isolated test function to send a single WATI message.
 * This function does NOT depend on any other part of the application.
 */
function simpleWatiTest() {
  Logger.log("--- Starting Simple WATI Test ---");

  try {
    // --- 1. DEFINE YOUR TEST DATA HERE ---
    const testPhoneNumber = "918583831888"; // The phone number from your previous log
    const testTemplateName = "outbound_msg";
    const testParameters = [
      {
        "name": "parent",
        "value": "Apeksha" // Using a real name for the test
      },
      {
        "name": "text",
        "value": "This is a direct test from Google Apps Script."
      },
      {
        "name": "text1",
        "value": "If you receive this, the API connection is successful."
      }
    ];

    // --- 2. GET YOUR SECURE CREDENTIALS ---
    const scriptProperties = PropertiesService.getScriptProperties();
    const API_ENDPOINT_BASE = (scriptProperties.getProperty('WATI_API_ENDPOINT') || "").trim();
    const ACCESS_TOKEN = (scriptProperties.getProperty('WATI_ACCESS_TOKEN') || "").trim();

    if (!API_ENDPOINT_BASE || !ACCESS_TOKEN) {
      throw new Error("WATI API Credentials are not configured in Script Properties. Please check Project Settings.");
    }

    // --- 3. CONSTRUCT THE CORRECT URL AND PAYLOAD (THIS IS THE FIX) ---
    const FULL_API_ENDPOINT = `${API_ENDPOINT_BASE}/api/v1/sendTemplateMessages`;
    
    const payload = {
      "template_messages": [{
        "template_name": testTemplateName,
        "broadcast_name": "Google_Script_Test",
        "parameters": testParameters,
        "whatsappNumber": testPhoneNumber.replace(/\D/g, '') // Removes '+' and spaces
      }]
    };

    const options = {
      "method": "post",
      "headers": {
        "Authorization": ACCESS_TOKEN,
        "Content-Type": "application/json"
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    // --- 4. EXECUTE THE API CALL AND LOG EVERYTHING ---
    Logger.log(`Sending to Endpoint: ${FULL_API_ENDPOINT}`);
    Logger.log(`Sending Payload: ${JSON.stringify(payload)}`);

    const response = UrlFetchApp.fetch(FULL_API_ENDPOINT, options);
    const responseCode = response.getResponseCode();
    const content = response.getContentText();

    Logger.log(`WATI Response Code: ${responseCode}`);
    Logger.log(`WATI Raw Response Content: "${content}"`);

    if (responseCode !== 200) {
      throw new Error(`WATI returned a non-successful HTTP status: ${responseCode}.`);
    }

    Logger.log("--- ✅ SUCCESS: WATI Test Completed ---");
    return { success: true, response: content };

  } catch (error) {
    Logger.log(`--- ❌ ERROR: WATI Test Failed ---`);
    Logger.log(error.message);
    throw error;
  }
}

// =================================================================
// END: PASTE THIS TEST FUNCTION INTO YOUR CODE.JS FILE
// =================================================================