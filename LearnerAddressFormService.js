// ============================================================
// LEARNER ADDRESS FORM SERVICE
// Public, no-login form keyed by JLID — a unique link per learner
// (…?page=addressForm&jlid=JLxxxx) that the parent opens, sees their
// child's name confirmed, and only fills in the address/contact fields.
// Writes straight to HubSpot contact properties + a log sheet for ops.
// ============================================================

// Resolves the HubSpot contact ID associated with a deal (first match).
function _lafGetContactId(dealId, token) {
  try {
    var resp = monitoredFetch(
      'https://api.hubapi.com/crm/v3/objects/deals/' + dealId + '/associations/contacts',
      { method: 'get', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (resp.getResponseCode() !== 200) return '';
    var data = JSON.parse(resp.getContentText());
    if (!data.results || !data.results.length) return '';
    return data.results[0].id || data.results[0].toObjectId || '';
  } catch(e) {
    Logger.log('[LAF] _lafGetContactId error: ' + e.message);
    return '';
  }
}

function _lafGetLogSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.AUDIT_SHEET_ID);
  var sheet = ss.getSheetByName('Learner Address Submissions');
  if (!sheet) {
    sheet = ss.insertSheet('Learner Address Submissions');
    sheet.appendRow(['Timestamp', 'JLID', 'Learner', 'Email', 'Address', 'City', 'State/Region', 'Postal Code', 'Country', 'Deal ID']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Called by AddressForm.html on load — looks up the learner name + any
// address/email HubSpot already has, so the parent can review and correct
// rather than retype everything from scratch.
function getAddressFormContext(jlid) {
  try {
    if (!jlid) return { success: false, message: 'Missing learner reference in this link.' };
    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success || !hs.data) return { success: false, message: 'We could not find this learner. Please contact JetLearn support.' };

    var d = hs.data;
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    var existing = { address: '', city: '', state: '', zip: '', country: '', email: '' };
    if (d.dealId) {
      var contactId = _lafGetContactId(d.dealId, token);
      if (contactId) {
        try {
          var cRes = monitoredFetch(
            'https://api.hubapi.com/crm/v3/objects/contacts/' + contactId + '?properties=email,address,city,state,zip,country',
            { method: 'get', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
          );
          if (cRes.getResponseCode() === 200) {
            var p = JSON.parse(cRes.getContentText()).properties || {};
            existing = { address: p.address || '', city: p.city || '', state: p.state || '', zip: p.zip || '', country: p.country || '', email: p.email || d.parentEmail || '' };
          }
        } catch(ce) { Logger.log('[LAF] contact fetch error: ' + ce.message); }
      }
    }
    if (!existing.email) existing.email = d.parentEmail || '';

    return {
      success: true,
      jlid: jlid,
      learnerName: d.learnerName || '',
      existing: existing
    };
  } catch(e) {
    Logger.log('[LAF] getAddressFormContext ERROR: ' + e.message);
    return { success: false, message: 'Something went wrong loading this form. Please contact JetLearn support.' };
  }
}

// Called on Submit — patches the HubSpot contact and logs the submission.
function submitLearnerAddressForm(payload) {
  try {
    payload = payload || {};
    var jlid = String(payload.jlid || '').trim().toUpperCase();
    if (!jlid) return { success: false, message: 'Missing learner reference.' };
    if (!payload.email || !payload.address || !payload.city || !payload.state || !payload.postalCode || !payload.country) {
      return { success: false, message: 'Please fill in all required fields.' };
    }

    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success || !hs.data) return { success: false, message: 'We could not find this learner. Please contact JetLearn support.' };
    var d = hs.data;

    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    var contactId = d.dealId ? _lafGetContactId(d.dealId, token) : '';

    if (contactId) {
      try {
        monitoredFetch('https://api.hubapi.com/crm/v3/objects/contacts/' + contactId, {
          method: 'PATCH',
          headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
          payload: JSON.stringify({ properties: {
            email:   String(payload.email).trim(),
            address: String(payload.address).trim().substring(0, 255),
            city:    String(payload.city).trim(),
            state:   String(payload.state).trim(),
            zip:     String(payload.postalCode).trim(),
            country: String(payload.country).trim()
          }}),
          muteHttpExceptions: true
        });
      } catch(pe) {
        Logger.log('[LAF] HubSpot contact PATCH failed (non-fatal): ' + pe.message);
      }
    } else {
      Logger.log('[LAF] No HubSpot contact found for deal ' + d.dealId + ' — logged only, not patched.');
    }

    var sheet = _lafGetLogSheet();
    sheet.appendRow([
      new Date(), jlid, d.learnerName || '', payload.email, payload.address,
      payload.city, payload.state, payload.postalCode, payload.country, d.dealId || ''
    ]);

    Logger.log('[LAF] Address submitted for ' + jlid);
    return { success: true };
  } catch(e) {
    Logger.log('[LAF] submitLearnerAddressForm ERROR: ' + e.message);
    return { success: false, message: 'Something went wrong submitting your details. Please contact JetLearn support.' };
  }
}

// Builds the shareable unique-link URL for a JLID. Called from the app UI
// (e.g. a "Copy Address Link" button) to generate the link to send the parent.
function getAddressFormLink(jlid) {
  if (!jlid) return { success: false, message: 'JLID required.' };
  var url = ScriptApp.getService().getUrl() + '?page=addressForm&jlid=' + encodeURIComponent(String(jlid).trim().toUpperCase());
  return { success: true, url: url };
}
