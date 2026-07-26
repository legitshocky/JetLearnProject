// ============================================================
// LEARNER ADDRESS FORM SERVICE
// Public, no-login form keyed by JLID — a unique link per learner
// (…?page=addressForm&jlid=JLxxxx) that the parent opens, sees their
// child's name confirmed, and only fills in the address/contact fields.
// Writes straight to HubSpot contact properties + a log sheet for ops.
// ============================================================

// ── DIAGNOSTIC — read-only. Shows exactly which contact(s) are associated
// with a JLID's deal, which one the PATCH targets, and that contact's
// CURRENT address properties (fresh from HubSpot, not cached). Run from the
// Apps Script editor's function dropdown after temporarily setting the jlid
// below, then check the log.
function diagCheckAddressPatchTarget(jlid) {
  jlid = jlid || 'JL39611449152C2';
  var hs = fetchHubspotByJlid(jlid);
  if (!hs || !hs.success || !hs.data) { Logger.log('[LAF Diag] HubSpot lookup failed: ' + (hs && hs.message)); return; }
  var d = hs.data;
  Logger.log('[LAF Diag] dealId=' + d.dealId + ' parentContact(deal prop)=' + d.parentContact + ' parentName(deal prop)=' + d.parentName + ' parentEmail(deal prop)=' + d.parentEmail);

  var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');

  // Show ALL associated contacts (not just the first one _lafGetContactId picks)
  var assocResp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/deals/' + d.dealId + '/associations/contacts', {
    method: 'get', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true
  });
  Logger.log('[LAF Diag] associations HTTP ' + assocResp.getResponseCode() + ': ' + assocResp.getContentText());

  var contactId = _lafGetContactId(d.dealId, token);
  Logger.log('[LAF Diag] _lafGetContactId resolved to: ' + (contactId || '(none)'));
  if (!contactId) return;

  var contactResp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/contacts/' + contactId + '?properties=email,address,city,state,zip,country,firstname,lastname', {
    method: 'get', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true
  });
  Logger.log('[LAF Diag] contact ' + contactId + ' current properties HTTP ' + contactResp.getResponseCode() + ': ' + contactResp.getContentText());
}

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

// Run once from the Apps Script editor's function dropdown to create the
// "Learner Address Submissions" sheet immediately, without waiting for a
// real form submission.
function setupLearnerAddressSheet() {
  _lafGetLogSheet();
  return 'Learner Address Submissions sheet ready.';
}

function _lafGetLogSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.AUDIT_SHEET_ID);
  var sheet = ss.getSheetByName('Learner Address Submissions');
  if (!sheet) {
    sheet = ss.insertSheet('Learner Address Submissions');
    sheet.appendRow(['Timestamp', 'JLID', 'Learner', 'Email', 'Address', 'City', 'State/Region', 'Postal Code', 'Country', 'Deal ID', 'Kit Bridge Status']);
    sheet.getRange(1, 1, 1, 11).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Called by AddressForm.html on load — looks up the learner name + any
// address/email HubSpot already has, so the parent can review and correct
// rather than retype everything from scratch.
function getAddressFormContext(jlid) {
  try {
    // Defensively strip stray quote characters — belt-and-braces in case a
    // link ever gets copy-pasted with encoding artifacts around the JLID.
    jlid = String(jlid || '').replace(/^"+|"+$/g, '').trim();
    if (!jlid) return { success: false, message: 'Missing learner reference in this link.' };

    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success || !hs.data) return { success: false, message: 'We could not find this learner (' + jlid + '): ' + ((hs && hs.message) || 'no data returned') + '. Please contact JetLearn support.' };

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
    var jlid = String(payload.jlid || '').replace(/^"+|"+$/g, '').trim().toUpperCase();
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
        var patchResp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/contacts/' + contactId, {
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
        var patchCode = patchResp.getResponseCode();
        if (patchCode < 200 || patchCode >= 300) {
          Logger.log('[LAF] HubSpot contact PATCH REJECTED (contactId=' + contactId + ', HTTP ' + patchCode + '): ' + patchResp.getContentText());
        } else {
          Logger.log('[LAF] HubSpot contact PATCH OK (contactId=' + contactId + ', HTTP ' + patchCode + ')');
        }
      } catch(pe) {
        Logger.log('[LAF] HubSpot contact PATCH failed (non-fatal): ' + pe.message);
      }
    } else {
      Logger.log('[LAF] No HubSpot contact found for deal ' + d.dealId + ' — logged only, not patched.');
    }

    var addressText = [payload.address, payload.city, payload.state, payload.postalCode, payload.country]
      .filter(function(p) { return p && String(p).trim(); }).join(', ');

    var bridgeStatus = _bridgeAddressToKitTracking(jlid, addressText, d);

    var sheet = _lafGetLogSheet();
    sheet.appendRow([
      new Date(), jlid, d.learnerName || '', payload.email, payload.address,
      payload.city, payload.state, payload.postalCode, payload.country, d.dealId || '', bridgeStatus
    ]);

    Logger.log('[LAF] Address submitted for ' + jlid + ' (kitBridgeStatus=' + bridgeStatus + ')');
    return { success: true };
  } catch(e) {
    Logger.log('[LAF] submitLearnerAddressForm ERROR: ' + e.message);
    return { success: false, message: 'Something went wrong submitting your details. Please contact JetLearn support.' };
  }
}

// Bridges a public-form address submission into the Kit Tracking sheet —
// this is what actually advances the Kit Tracking pipeline (ADDR_STATUS
// 'Requested' → 'Received') when a parent uses the new short link instead of
// the older WATI free-text / HubSpot-form flow. Never blocks the caller —
// any failure here is logged and swallowed so the parent still sees success.
// Returns a short status string logged on the submission row for auditing.
function _bridgeAddressToKitTracking(jlid, addressText, hsData) {
  try {
    var sheet = _getKitSheet();
    if (sheet.getLastRow() < 2) return 'no_kit_sheet_rows';

    var rowIndex = _findOpenKitRowByJlid(jlid);
    if (!rowIndex) return 'no_kit_row_found_or_ambiguous';

    var now = new Date();
    sheet.getRange(rowIndex, KIT_COL.DELIVERY_ADDRESS).setValue(addressText);
    sheet.getRange(rowIndex, KIT_COL.ADDR_STATUS).setValue('Received');
    sheet.getRange(rowIndex, KIT_COL.NUDGE_STAGE_AT).setValue(now);
    sheet.getRange(rowIndex, KIT_COL.ADDR_SUBMITTED_AT).setValue(now);

    // Same confirmation WATI the internal flow sends — parent gets a
    // consistent experience regardless of which channel they used.
    var phone = _normalisePhone((hsData && hsData.parentContact) || '');
    try {
      if (phone) {
        var kitName = String(sheet.getRange(rowIndex, KIT_COL.KIT).getValue() || '').trim();
        sendWatiMessage(phone, 'kit_address_received_confirmation', [
          { name: '1', value: (hsData && hsData.parentName) || '' },
          { name: '2', value: kitName }
        ]);
      }
    } catch(we) {
      Logger.log('[LAF] bridge confirmation WATI failed (non-fatal): ' + we.message);
    }

    // Clear the pending-request cache/queue — otherwise fetchKitLearnerDetails
    // keeps showing "waiting for parent" in the Add Kit Entry modal even
    // though the address has already come in via this channel.
    try {
      if (phone) CacheService.getScriptCache().remove('KIT_ADDR_REQ_' + phone);
      _kitAddrQueueRemove(jlid);
    } catch(ce) {
      Logger.log('[LAF] bridge cache/queue cleanup failed (non-fatal): ' + ce.message);
    }

    Logger.log('[LAF] Bridged address to Kit Tracking row ' + rowIndex + ' for ' + jlid);
    return 'bridged';
  } catch(e) {
    Logger.log('[LAF] _bridgeAddressToKitTracking error: ' + e.message);
    return 'bridge_error';
  }
}

// Builds the shareable unique-link URL for a JLID. Called from the app UI
// (e.g. a "Copy Address Link" button) to generate the link to send the parent.
function getAddressFormLink(jlid, learnerName) {
  if (!jlid) return { success: false, message: 'JLID required.' };

  var cleanJlid = String(jlid).trim().toUpperCase();
  var longUrl = ScriptApp.getService().getUrl() + '?page=addressForm&r=' + encodeURIComponent(cleanJlid);
  if (learnerName) longUrl += '&name=' + encodeURIComponent(String(learnerName).trim());

  // Branded short link — a real static page hosted on Firebase Hosting
  // (jetlearn-kit-links.web.app), not a redirect. The page fetch()es learner
  // data from the GAS backend as a JSON API, so this URL stays in the address
  // bar the whole time. Safe to paste directly into WhatsApp.
  var shortUrl = 'https://jetlearn-kit-links.web.app/kit/' + encodeURIComponent(cleanJlid);
  if (learnerName) shortUrl += '?name=' + encodeURIComponent(String(learnerName).trim());

  return { success: true, url: shortUrl, longUrl: longUrl };
}