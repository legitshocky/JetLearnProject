/**
 * KitTrackingService.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Automates kit delivery follow-up via WATI WhatsApp.
 *
 * Flow:
 *   Daily trigger → sendKitFollowUps()
 *     → reads Kit Tracking sheet
 *     → sends WATI "migration_kit_fup_sent_by_us" template for overdue rows
 *     → marks Follow-up Sent + stores phone
 *
 *   WATI webhook → handleKitReply(waId, buttonText)
 *     → matches phone to sheet row
 *     → "Kit Received"     → fills Delivery Date, updates HubSpot kit status
 *     → "Not Received yet" → flags sheet, adds HubSpot deal note
 *     → "Need To Check"    → flags sheet for manual review
 *
 * Sheet: "Kit Tracking" (SHEET_ID_KIT_TRACKING)
 * Columns (1-indexed):
 *   A(1)  Sr No         B(2)  Learner Name   C(3)  Kit
 *   D(4)  Country       E(5)  Price EUR       F(6)  Site
 *   G(7)  Date of Order H(8)  Timestamp Month I(9)  ETA
 *   J(10) Delivery Date K(11) Time Taken      L(12) Reason
 *   M(13) Subscription  N(14) Roadmap         O(15) Col 15
 *   P(16) JLID          Q(17) Follow-up Sent  R(18) Follow-up Sent At
 *   S(19) Parent Response  T(20) Phone Sent To
 */

// ── Column indices (1-based for getRange) ─────────────────────────────────────
var KIT_COL = {
  SR_NO:              1,
  LEARNER_NAME:       2,
  KIT:                3,
  DATE_OF_ORDER:      7,
  ETA:                9,
  DELIVERY_DATE:      10,
  TIME_TAKEN:         11,
  JLID:               16,
  FOLLOWUP_SENT:      17,
  FOLLOWUP_SENT_AT:   18,
  PARENT_RESPONSE:    19,
  PHONE_SENT_TO:      20,
  FOLLOWUP2_SENT:     21,
  FOLLOWUP2_SENT_AT:  22,
  // Address collection columns (appended — W, X)
  DELIVERY_ADDRESS:   23,
  ADDR_STATUS:        24,  // HubSpot / Requested / Received / Verified
  // Escalation columns — Y, Z
  ESCALATED:          25,
  ESCALATED_AT:       26,
  // Refunded flag — AA(27)
  REFUNDED:           27
};

// ── HubSpot kit property map ──────────────────────────────────────────────────
// Fetch current learning_kit_cost directly from deal GET — bypasses search cache
function _getHubSpotDealKitCost(dealId, token) {
  try {
    var resp = monitoredFetch(
      'https://api.hubapi.com/crm/v3/objects/deals/' + dealId + '?properties=learning_kit_cost',
      { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (resp.getResponseCode() === 200) {
      var d = JSON.parse(resp.getContentText());
      return parseFloat((d.properties || {}).learning_kit_cost) || 0;
    }
    Logger.log('[_getHubSpotDealKitCost] HTTP ' + resp.getResponseCode());
  } catch(e) { Logger.log('[_getHubSpotDealKitCost] ERROR: ' + e.message); }
  return 0;
}

// Patch learning_kit_cost on a deal (used by add + refund)
function _patchHubSpotKitCost(dealId, token, newValue) {
  try {
    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/deals/' + dealId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ properties: { learning_kit_cost: newValue } }),
      muteHttpExceptions: true
    });
    Logger.log('[_patchHubSpotKitCost] dealId=' + dealId + ' newValue=' + newValue + ' HTTP ' + resp.getResponseCode());
    return resp.getResponseCode() === 200;
  } catch(e) { Logger.log('[_patchHubSpotKitCost] ERROR: ' + e.message); return false; }
}

function _kitPropertyForType(kitName) {
  if (!kitName) return null;
  var k = kitName.toLowerCase().trim();
  if (k.indexOf('vr')       > -1 || k.indexOf('oculus') > -1 || k.indexOf('headset') > -1) return 'vr_headset__oculus_status'; // confirmed via API
  if (k.indexOf('microbit') > -1)                                                            return 'microbit_kit_status';           // confirmed via API
  if (k.indexOf('makey')    > -1)                                                            return 'makey_makey_kit_status';         // confirmed via API (no __t_)
  if (k.indexOf('arduino')  > -1)                                                            return 'arduino_kit_status';             // confirmed via API
  return null;
}

// ── Fetch delivery address from HubSpot contact associated with a deal ───────
function _fetchContactAddress(dealId) {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    if (!token || !dealId) return {};
    var assocRes = monitoredFetch(
      'https://api.hubapi.com/crm/v3/objects/deals/' + dealId + '/associations/contacts',
      { method: 'get', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (assocRes.getResponseCode() !== 200) return {};
    var assocData = JSON.parse(assocRes.getContentText());
    if (!assocData.results || !assocData.results.length) return {};
    var contactId = assocData.results[0].id || assocData.results[0].toObjectId;
    if (!contactId) return {};
    var cRes = monitoredFetch(
      'https://api.hubapi.com/crm/v3/objects/contacts/' + contactId + '?properties=address,city,state,zip,country',
      { method: 'get', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (cRes.getResponseCode() !== 200) return {};
    var p = (JSON.parse(cRes.getContentText()).properties) || {};
    return { address: p.address || '', city: p.city || '', state: p.state || '', zip: p.zip || '', country: p.country || '' };
  } catch(e) {
    Logger.log('[KitTracking] _fetchContactAddress error: ' + e.message);
    return {};
  }
}

// ── Count line items on the deal itself ───────────────────────────────────────
// >1 subscription line item on the deal = renewed learner (added a new course/term
// to an existing deal rather than this being their first subscription line).
function _countDealLineItems(dealId) {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    if (!token || !dealId) return 0;
    var res = monitoredFetch(
      'https://api.hubapi.com/crm/v3/objects/deals/' + dealId + '/associations/line_items',
      { method: 'get', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (res.getResponseCode() !== 200) return 0;
    var data = JSON.parse(res.getContentText());
    return (data.results || []).length;
  } catch(e) {
    Logger.log('[KitTracking] _countDealLineItems error: ' + e.message);
    return 0;
  }
}

// ── Build "Delivery to …" address string for WATI ────────────────────────────
function _buildKitAddress(learnerName, addrObj, fallbackCountry) {
  var parts = [addrObj.address, addrObj.city, addrObj.state, addrObj.country || fallbackCountry]
    .filter(function(p) { return p && String(p).trim(); });
  return 'Delivery to ' + (learnerName || '') + (parts.length ? ', ' + parts.join(', ') : '');
}

// ── Update HubSpot contact address from free-text (best-effort, non-fatal) ───
function _updateHubSpotContactAddress(dealId, addressText) {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    if (!token || !dealId) return;
    var assocRes = monitoredFetch(
      'https://api.hubapi.com/crm/v3/objects/deals/' + dealId + '/associations/contacts',
      { method: 'get', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (assocRes.getResponseCode() !== 200) return;
    var assocData = JSON.parse(assocRes.getContentText());
    if (!assocData.results || !assocData.results.length) return;
    var contactId = assocData.results[0].id || assocData.results[0].toObjectId;
    if (!contactId) return;
    // Store raw text in HubSpot contact `address` field (single line)
    monitoredFetch('https://api.hubapi.com/crm/v3/objects/contacts/' + contactId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ properties: { address: String(addressText).substring(0, 255) } }),
      muteHttpExceptions: true
    });
    Logger.log('[KitTracking] HubSpot contact address updated for contactId=' + contactId);
  } catch(e) {
    Logger.log('[KitTracking] _updateHubSpotContactAddress error: ' + e.message);
  }
}

// ── Request delivery address from parent via WATI ────────────────────────────
// Call before placing a kit order when HubSpot address is blank.
// Returns { success, addressSource, address, phone, needsWati, watiSent }
function requestKitDeliveryAddress(jlid, kitName, rowIndex) {
  try {
    if (!jlid || !kitName) return { success: false, message: 'JLID and kit name required' };

    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success) return { success: false, message: 'HubSpot lookup failed for ' + jlid };

    var d           = hs.data;
    var phone       = _normalisePhone(d.parentContact || '');
    var parentName  = d.parentName  || '';
    var learnerName = d.learnerName || '';

    if (!phone) return { success: false, message: 'No phone number for ' + jlid };

    // Always reconfirm address — parents move. Never assume existing address is current.
    // Always send WATI regardless of what HubSpot contact address says.
    var wRes = sendWatiMessage(phone, 'migration_address_template', [
      { name: 'Parent',   value: parentName },
      { name: 'Kit_type', value: kitName    }
    ]);

    if (!wRes || !wRes.success) {
      return { success: false, message: 'WATI send failed: ' + (wRes && wRes.error ? wRes.error : 'Unknown') };
    }

    // Cache pending request (TTL 8 hours) — for WATI free-text reply path
    var cachePayload = JSON.stringify({
      jlid:        jlid,
      kitName:     kitName,
      learnerName: learnerName,
      parentName:  parentName,
      rowIndex:    rowIndex || 0,
      requestedAt: new Date().toISOString()
    });
    CacheService.getScriptCache().put('KIT_ADDR_REQ_' + phone, cachePayload, 28800);

    // Add to ScriptProperties queue — poll trigger can enumerate even without sheet row
    _kitAddrQueueAdd({
      jlid:        jlid,
      kitName:     kitName,
      phone:       phone,
      parentName:  parentName,
      learnerName: learnerName,
      dealId:      d.dealId || '',
      rowIndex:    rowIndex || 0,
      requestedAt: new Date().toISOString()
    });

    // Update sheet if row provided
    if (rowIndex > 0) {
      try {
        var sheet2 = _getKitSheet();
        sheet2.getRange(rowIndex, KIT_COL.ADDR_STATUS).setValue('Requested');
        sheet2.getRange(rowIndex, KIT_COL.DELIVERY_ADDRESS).setValue('');
      } catch(se2) {}
    }

    // Update HubSpot kit status → "Asked for address" (best-effort)
    try {
      var kitPropForAddr = _kitPropertyForType(kitName);
      if (kitPropForAddr && d.dealId) {
        var hsToken = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
        var addrHsProps = {};
        addrHsProps[kitPropForAddr] = 'Asked for address';
        monitoredFetch('https://api.hubapi.com/crm/v3/objects/deals/' + d.dealId, {
          method: 'PATCH',
          headers: { 'Authorization': 'Bearer ' + hsToken, 'Content-Type': 'application/json' },
          payload: JSON.stringify({ properties: addrHsProps }),
          muteHttpExceptions: true
        });
        Logger.log('[KitTracking] HubSpot kit status → "Asked for address" for ' + jlid);
      }
    } catch(hsAddrErr) {
      Logger.log('[KitTracking] HS "Asked for address" update failed (non-fatal): ' + hsAddrErr.message);
    }

    Logger.log('[KitTracking] Address request sent via WATI for ' + jlid + ' phone=' + phone);
    return { success: true, addressSource: 'wati_requested',
             phone: phone, needsWati: true, watiSent: true };

  } catch(e) {
    Logger.log('[KitTracking] requestKitDeliveryAddress ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Called from doPost when parent replies with address text ─────────────────
// Returns true if the reply was handled as an address, false if not a pending request
function handleKitAddressReply(phone, addressText) {
  try {
    Logger.log('[KitTracking] handleKitAddressReply phone=' + phone + ' text="' + String(addressText).substring(0,100) + '"');

    var sc       = CacheService.getScriptCache();
    var cacheKey = 'KIT_ADDR_REQ_' + phone;
    var cached   = sc.get(cacheKey);
    if (!cached) {
      Logger.log('[KitTracking] No pending address request for phone ' + phone);
      return false;
    }

    var meta     = JSON.parse(cached);
    var rowIndex = meta.rowIndex || 0;
    var addrText = String(addressText).trim();

    // Update sheet row
    var sheet = _getKitSheet();
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, KIT_COL.DELIVERY_ADDRESS).setValue(addrText);
      sheet.getRange(rowIndex, KIT_COL.ADDR_STATUS).setValue('Received');
    } else {
      // Find row by JLID
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        var rows = sheet.getRange(2, 1, lastRow - 1, KIT_COL.JLID).getValues();
        for (var i = 0; i < rows.length; i++) {
          if (String(rows[i][KIT_COL.JLID - 1] || '').trim() === meta.jlid) {
            sheet.getRange(i + 2, KIT_COL.DELIVERY_ADDRESS).setValue(addrText);
            sheet.getRange(i + 2, KIT_COL.ADDR_STATUS).setValue('Received');
            break;
          }
        }
      }
    }

    // Push address back to HubSpot contact (best-effort)
    try {
      var hs = fetchHubspotByJlid(meta.jlid);
      if (hs && hs.success && hs.data && hs.data.dealId) {
        _updateHubSpotContactAddress(hs.data.dealId, addrText);
      }
    } catch(hse) {
      Logger.log('[KitTracking] handleKitAddressReply: HS update failed (non-fatal): ' + hse.message);
    }

    // Send thank-you reply to parent — {{1}}=Parent {{2}}=Kit
    try {
      sendWatiMessage(phone, 'kit_address_received_confirmation', [
        { name: '1', value: meta.parentName || '' },
        { name: '2', value: meta.kitName    || '' }
      ]);
    } catch(we) {}

    sc.remove(cacheKey);
    _kitAddrQueueRemove(meta.jlid); // clean ScriptProperties queue too
    Logger.log('[KitTracking] Address captured for ' + meta.jlid + ': "' + addrText + '"');
    return true;

  } catch(e) {
    Logger.log('[KitTracking] handleKitAddressReply ERROR: ' + e.message);
    return false;
  }
}

// ── Poll HubSpot for address after parent fills the form ─────────────────────
// Called from auto-trigger (pollPendingKitAddresses) and "Check for Reply" button.
// When address found: updates sheet, sends WATI confirmation, clears cache.
// Returns { received, address, confirmationSent }
function checkKitAddressReply(jlid, rowIndex, kitName) {
  try {
    if (!jlid) return { received: false };
    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success) return { received: false, message: 'HubSpot lookup failed' };

    var d     = hs.data;
    var phone = _normalisePhone(d.parentContact || '');

    // If kitName not passed, try to read from cache or sheet
    var resolvedKit = kitName || '';
    if (!resolvedKit && phone) {
      try {
        var cachedReq = CacheService.getScriptCache().get('KIT_ADDR_REQ_' + phone);
        if (cachedReq) resolvedKit = JSON.parse(cachedReq).kitName || '';
      } catch(ce) {}
    }
    if (!resolvedKit && rowIndex && rowIndex > 0) {
      try {
        var sheetForKit = _getKitSheet();
        resolvedKit = String(sheetForKit.getRange(rowIndex, KIT_COL.KIT).getValue() || '').trim();
      } catch(se) {}
    }

    // Poll HubSpot contact for address
    var addr = d.dealId ? _fetchContactAddress(d.dealId) : {};
    var addrStr = [addr.address, addr.city, addr.state, addr.zip, addr.country]
      .filter(function(p) { return p && String(p).trim(); }).join(', ');

    if (!addrStr) {
      // Address not yet in HubSpot — still waiting for parent to submit form
      return { received: false, waiting: true };
    }

    // Address found ─ update sheet row
    if (rowIndex && rowIndex > 0) {
      try {
        var sheet = _getKitSheet();
        sheet.getRange(rowIndex, KIT_COL.DELIVERY_ADDRESS).setValue(addrStr);
        sheet.getRange(rowIndex, KIT_COL.ADDR_STATUS).setValue('Received');
      } catch(se) {
        Logger.log('[KitTracking] checkKitAddressReply: sheet update failed: ' + se.message);
      }
    }

    // Send WATI confirmation to parent (best-effort)
    // Template kit_address_received_confirmation uses positional {{1}}=Parent {{2}}=Kit
    var confirmationSent = false;
    if (phone) {
      try {
        var wRes = sendWatiMessage(phone, 'kit_address_received_confirmation', [
          { name: '1', value: d.parentName || '' },
          { name: '2', value: resolvedKit  || '' }
        ]);
        confirmationSent = !!(wRes && wRes.success);
        Logger.log('[KitTracking] Confirmation WATI sent=' + confirmationSent + ' for ' + jlid);
      } catch(we) {
        Logger.log('[KitTracking] checkKitAddressReply: WATI confirm error: ' + we.message);
      }
    }

    // Clear pending cache
    if (phone) {
      try { CacheService.getScriptCache().remove('KIT_ADDR_REQ_' + phone); } catch(ce) {}
    }

    Logger.log('[KitTracking] Address found for ' + jlid + ': "' + addrStr + '"');
    return { received: true, address: addrStr, confirmationSent: confirmationSent };

  } catch(e) {
    Logger.log('[KitTracking] checkKitAddressReply ERROR: ' + e.message);
    return { received: false, message: e.message };
  }
}

// ── ScriptProperties queue helpers ───────────────────────────────────────────
// Stores pending address requests so poll can find them even without a sheet row.
var _KIT_ADDR_QUEUE_KEY = 'KIT_ADDR_QUEUE';

function _kitAddrQueueGet() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(_KIT_ADDR_QUEUE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch(e) { return []; }
}

function _kitAddrQueueAdd(entry) {
  try {
    var queue = _kitAddrQueueGet();
    // Remove stale entry for same JLID first (re-request scenario)
    queue = queue.filter(function(q) { return q.jlid !== entry.jlid; });
    queue.push(entry);
    PropertiesService.getScriptProperties().setProperty(_KIT_ADDR_QUEUE_KEY, JSON.stringify(queue));
    Logger.log('[KitAddrQueue] Added JLID=' + entry.jlid + ' queue size=' + queue.length);
  } catch(e) {
    Logger.log('[KitAddrQueue] _kitAddrQueueAdd ERROR: ' + e.message);
  }
}

function _kitAddrQueueRemove(jlid) {
  try {
    var queue = _kitAddrQueueGet().filter(function(q) { return q.jlid !== jlid; });
    PropertiesService.getScriptProperties().setProperty(_KIT_ADDR_QUEUE_KEY, JSON.stringify(queue));
    Logger.log('[KitAddrQueue] Removed JLID=' + jlid + ' queue size=' + queue.length);
  } catch(e) {
    Logger.log('[KitAddrQueue] _kitAddrQueueRemove ERROR: ' + e.message);
  }
}

// ── Auto-poll: scan sheet rows + ScriptProperties queue, check HubSpot ───────
// Runs every 30 min via time-based trigger. No manual action needed.
function pollPendingKitAddresses() {
  Logger.log('[KitAddrPoll] Starting poll...');
  var checked = 0, found = 0;

  try {
    // ── Part 1: sheet rows with ADDR_STATUS='Requested' ──────────────────────
    var sheet   = _getKitSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var rows = sheet.getRange(2, 1, lastRow - 1, KIT_COL.ADDR_STATUS).getValues();
      rows.forEach(function(row, idx) {
        var sheetRow   = idx + 2;
        var addrStatus = String(row[KIT_COL.ADDR_STATUS - 1] || '').trim();
        var jlid       = String(row[KIT_COL.JLID - 1]        || '').trim();
        var kitName    = String(row[KIT_COL.KIT - 1]          || '').trim();
        if (addrStatus !== 'Requested' || !jlid) return;
        checked++;
        Logger.log('[KitAddrPoll] Sheet row=' + sheetRow + ' JLID=' + jlid);
        try {
          var result = checkKitAddressReply(jlid, sheetRow, kitName);
          if (result.received) {
            found++;
            _kitAddrQueueRemove(jlid); // clean up queue too if entry exists
            Logger.log('[KitAddrPoll] Sheet: address found JLID=' + jlid + ' confirm=' + result.confirmationSent);
          }
        } catch(re) { Logger.log('[KitAddrPoll] Sheet row error JLID=' + jlid + ': ' + re.message); }
        Utilities.sleep(400);
      });
    }

    // ── Part 2: ScriptProperties queue (requests made before sheet row exists) ─
    var queue = _kitAddrQueueGet();
    Logger.log('[KitAddrPoll] Queue size=' + queue.length);
    queue.forEach(function(entry) {
      checked++;
      Logger.log('[KitAddrPoll] Queue JLID=' + entry.jlid + ' kit=' + entry.kitName);
      try {
        // Poll HubSpot contact directly for address
        var addr = entry.dealId ? _fetchContactAddress(entry.dealId) : {};
        var addrStr = [addr.address, addr.city, addr.state, addr.zip, addr.country]
          .filter(function(p) { return p && String(p).trim(); }).join(', ');

        if (!addrStr) {
          Logger.log('[KitAddrPoll] Queue: still waiting for JLID=' + entry.jlid);
          return;
        }

        found++;
        Logger.log('[KitAddrPoll] Queue: address found for JLID=' + entry.jlid + ': "' + addrStr + '"');

        // Send confirmation WATI — {{1}}=Parent {{2}}=Kit
        var confirmSent = false;
        if (entry.phone) {
          try {
            var wRes = sendWatiMessage(entry.phone, 'kit_address_received_confirmation', [
              { name: '1', value: entry.parentName  || '' },
              { name: '2', value: entry.kitName     || '' }
            ]);
            confirmSent = !!(wRes && wRes.success);
            Logger.log('[KitAddrPoll] Confirmation WATI sent=' + confirmSent + ' JLID=' + entry.jlid);
          } catch(we) {
            Logger.log('[KitAddrPoll] WATI confirm error: ' + we.message);
          }
        }

        // Update sheet row if it exists now (e.g. addKitEntry ran after request)
        if (entry.rowIndex > 0) {
          try {
            var sh = _getKitSheet();
            sh.getRange(entry.rowIndex, KIT_COL.DELIVERY_ADDRESS).setValue(addrStr);
            sh.getRange(entry.rowIndex, KIT_COL.ADDR_STATUS).setValue('Received');
          } catch(se) {}
        }

        // Clear ScriptCache entry
        if (entry.phone) {
          try { CacheService.getScriptCache().remove('KIT_ADDR_REQ_' + entry.phone); } catch(ce) {}
        }

        // Remove from queue
        _kitAddrQueueRemove(entry.jlid);

      } catch(qe) {
        Logger.log('[KitAddrPoll] Queue entry error JLID=' + entry.jlid + ': ' + qe.message);
      }
      Utilities.sleep(400);
    });

    Logger.log('[KitAddrPoll] Done. Checked=' + checked + ' Found=' + found);
  } catch(e) {
    Logger.log('[KitAddrPoll] ERROR: ' + e.message);
  }
}

// ── HubSpot form webhook handler — fires instantly on form submit ─────────────
// Called from doPost when HubSpot Kit Address Form is submitted.
// Matches submission to pending queue entry by phone or email, sends WATI immediately.
function _handleKitAddressFormWebhook(payload) {
  Logger.log('[KitAddrWebhook] Processing form submission...');

  // Extract fields from submission data array
  var fields = {};
  (payload.data || []).forEach(function(f) {
    fields[String(f.name || '').toLowerCase()] = String(f.value || '').trim();
  });
  Logger.log('[KitAddrWebhook] Fields: ' + JSON.stringify(fields));

  // Build address string from submitted fields
  var addrParts = [
    fields['address'] || fields['street'] || fields['street_address'] || '',
    fields['city']    || '',
    fields['state']   || fields['county'] || '',
    fields['zip']     || fields['postcode'] || fields['postal_code'] || '',
    fields['country'] || ''
  ].filter(function(p) { return p && p.trim(); });

  // Also try a single combined address field
  if (!addrParts.length) {
    var combined = fields['full_address'] || fields['delivery_address'] || fields['address_line_1'] || '';
    if (combined) addrParts = [combined];
  }

  var addrStr = addrParts.join(', ');
  Logger.log('[KitAddrWebhook] Parsed address: "' + addrStr + '"');

  if (!addrStr) {
    Logger.log('[KitAddrWebhook] No address found in form fields — skipping');
    return;
  }

  // Match to pending queue entry by phone or email
  var submittedPhone = _normalisePhone(fields['phone'] || fields['mobilephone'] || fields['phone_number'] || '');
  var submittedEmail = fields['email'] || '';
  var queue = _kitAddrQueueGet();

  Logger.log('[KitAddrWebhook] Queue size=' + queue.length + ' phone="' + submittedPhone + '" email="' + submittedEmail + '"');

  var matched = null;

  // Match by phone first
  if (submittedPhone) {
    queue.forEach(function(q) {
      if (!matched && _normalisePhone(q.phone) === submittedPhone) matched = q;
    });
  }

  // Fallback: match by email via HubSpot lookup
  if (!matched && submittedEmail) {
    queue.forEach(function(q) {
      if (matched) return;
      try {
        var hs = fetchHubspotByJlid(q.jlid);
        if (hs && hs.success && (hs.data.parentEmail || '').toLowerCase() === submittedEmail.toLowerCase()) {
          matched = q;
        }
      } catch(le) {}
    });
  }

  // Last resort: if only one entry in queue, assume it's the one
  if (!matched && queue.length === 1) {
    matched = queue[0];
    Logger.log('[KitAddrWebhook] Single queue entry — assuming match: JLID=' + matched.jlid);
  }

  if (!matched) {
    Logger.log('[KitAddrWebhook] No matching queue entry found — poll will catch it shortly');
    return;
  }

  Logger.log('[KitAddrWebhook] Matched JLID=' + matched.jlid + ' kit=' + matched.kitName);

  // Send confirmation WATI immediately — {{1}}=Parent {{2}}=Kit
  var confirmSent = false;
  if (matched.phone) {
    try {
      var wRes = sendWatiMessage(matched.phone, 'kit_address_received_confirmation', [
        { name: '1', value: matched.parentName || '' },
        { name: '2', value: matched.kitName    || '' }
      ]);
      confirmSent = !!(wRes && wRes.success);
      Logger.log('[KitAddrWebhook] WATI confirmation sent=' + confirmSent);
    } catch(we) {
      Logger.log('[KitAddrWebhook] WATI error: ' + we.message);
    }
  }

  // Update sheet row if exists
  if (matched.rowIndex > 0) {
    try {
      var sh = _getKitSheet();
      sh.getRange(matched.rowIndex, KIT_COL.DELIVERY_ADDRESS).setValue(addrStr);
      sh.getRange(matched.rowIndex, KIT_COL.ADDR_STATUS).setValue('Received');
    } catch(se) {}
  }

  // Clear cache + queue
  if (matched.phone) {
    try { CacheService.getScriptCache().remove('KIT_ADDR_REQ_' + matched.phone); } catch(ce) {}
  }
  _kitAddrQueueRemove(matched.jlid);

  Logger.log('[KitAddrWebhook] Done. JLID=' + matched.jlid + ' address="' + addrStr + '" confirm=' + confirmSent);
}

// ── Register 1-min address poll trigger (fallback) ────────────────────────────
function setupKitAddressPollTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'pollPendingKitAddresses') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pollPendingKitAddresses').timeBased().everyMinutes(1).create();
  Logger.log('[KitTracking] 1-min kit address poll trigger created.');
}

// Installs the daily kit follow-up trigger (9 AM). Run once from the editor,
// or call again any time — it de-dupes existing triggers first.
function setupKitFollowUpTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendKitFollowUps') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendKitFollowUps').timeBased().everyDays(1).atHour(9).create();
  Logger.log('[KitTracking] Daily 9AM sendKitFollowUps trigger created.');
  return 'Daily 9AM kit follow-up trigger installed.';
}

// ── Get Kit Tracking sheet ────────────────────────────────────────────────────
function _getKitSheet() {
  var ssId = PropertiesService.getScriptProperties().getProperty('SHEET_ID_KIT_TRACKING');
  if (!ssId) {
    // Fallback: hardcoded ID from plan
    ssId = '17Jsa2Kl2AkI5SgtlITYGqb-Q-PxfzkNFzzG5HwBJp_Q';
  }
  var ss = SpreadsheetApp.openById(ssId);
  // Try tab named "Kit Tracking", fall back to first sheet
  var sheet = ss.getSheetByName('Kits') || ss.getSheetByName('Kit Tracking') || ss.getSheets()[0];
  return sheet;
}

// ── Parse DD/MM/YYYY → Date ───────────────────────────────────────────────────
function _parseDMY(str) {
  if (!str) return null;
  var s = String(str).trim();
  // Handle Date objects returned by Sheets
  if (s.indexOf('/') === -1 && s.indexOf('-') === -1) return null;
  var parts = s.indexOf('/') > -1 ? s.split('/') : s.split('-');
  if (parts.length !== 3) return null;
  // DD/MM/YYYY
  var d = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10) - 1;
  var y = parseInt(parts[2], 10);
  var dt = new Date(y, m, d);
  return isNaN(dt.getTime()) ? null : dt;
}

// ── Format Date → DD/MM/YYYY ──────────────────────────────────────────────────
function _formatDMY(date) {
  var d = date.getDate();
  var m = date.getMonth() + 1;
  var y = date.getFullYear();
  return (d < 10 ? '0' + d : d) + '/' + (m < 10 ? '0' + m : m) + '/' + y;
}

// ── Normalise phone → digits only, no leading + ───────────────────────────────
function _normalisePhone(raw) {
  if (!raw) return '';
  return String(raw).replace(/[^0-9]/g, '');
}

// ─────────────────────────────────────────────────────────────────────────────
// ALL KIT PROPERTIES  (used for multi-kit HubSpot search)
// ─────────────────────────────────────────────────────────────────────────────
var KIT_HS_PROPS = [
  'vr_headset__oculus_status',   // confirmed
  'microbit_kit_status',          // confirmed
  'makey_makey_kit_status',       // confirmed (no __t_)
  'arduino_kit_status'            // confirmed
];

// ─────────────────────────────────────────────────────────────────────────────
// In-memory cache for the �Sent� deals list � fetched once per trigger run
var _kitSentDealsCache = null;

// AUTO-FIND JLID  — searches HubSpot deals where ANY kit property = �Sent� (internal enum value)
// then matches by learner name. Fetches once per execution, reuses cached results for all rows.
// ─────────────────────────────────────────────────────────────────────────────
function _findJlidByKitStatus(learnerName, kitName) {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    if (!token) return null;

    // Use cached results if already fetched this execution
    var results = _kitSentDealsCache;
    if (!results) {
      var filterGroups = KIT_HS_PROPS.map(function(prop) {
        return { filters: [{ propertyName: prop, operator: 'EQ', value: 'Sent' }] };
      });
      var body = {
        filterGroups: filterGroups,
        properties: ['jetlearner_id', 'dealname'].concat(KIT_HS_PROPS),
        limit: 200
      };
      var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/deals/search', {
        method: 'post',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify(body),
        muteHttpExceptions: true
      });
      var respCode = resp.getResponseCode();
      var respText = resp.getContentText();
      Logger.log('[KitTracking] _findJlidByKitStatus: HTTP ' + respCode + ' (fresh fetch)');
      var data = JSON.parse(respText);
      results = (data && data.results) || [];
      _kitSentDealsCache = results; // cache for rest of run
      Logger.log('[KitTracking] _findJlidByKitStatus: ' + results.length + ' deals cached with any kit=�Sent�');
    }

    if (!results || !results.length) return null;

    var nameLower = learnerName.toLowerCase().trim();
    var match = null;

    // Pass 1 — full name match
    results.forEach(function(deal) {
      if (match) return;
      var dealName = String((deal.properties && deal.properties.dealname) || '').toLowerCase();
      var jlid     = String((deal.properties && deal.properties.jetlearner_id) || '').trim();
      if (jlid && dealName.indexOf(nameLower) > -1) match = jlid;
    });

    // Pass 2 — first name only (fallback)
    if (!match) {
      var firstName = nameLower.split(' ')[0];
      if (firstName.length > 2) {
        results.forEach(function(deal) {
          if (match) return;
          var dealName = String((deal.properties && deal.properties.dealname) || '').toLowerCase();
          var jlid     = String((deal.properties && deal.properties.jetlearner_id) || '').trim();
          if (jlid && dealName.indexOf(firstName) > -1) match = jlid;
        });
      }
    }

    if (match) {
      Logger.log('[KitTracking] _findJlidByKitStatus: matched JLID=' + match + ' for "' + learnerName + '" (kit: ' + kitName + ')');
    } else {
      Logger.log('[KitTracking] _findJlidByKitStatus: no name match for "' + learnerName + '" in ' + results.length + ' results');
    }
    return match;

  } catch (e) {
    Logger.log('[KitTracking] _findJlidByKitStatus ERROR: ' + e.message);
    return null;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// SEND KIT FOLLOW-UPS  (daily trigger at 8am)
// ─────────────────────────────────────────────────────────────────────────────
function sendKitFollowUps() {
  Logger.log('[KitTracking] sendKitFollowUps started');
  try {
    var sheet = _getKitSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('[KitTracking] No data rows found.');
      return;
    }

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var rows = sheet.getRange(2, 1, lastRow - 1, KIT_COL.REFUNDED).getValues();
    var sent = 0;

    rows.forEach(function(row, idx) {
      var sheetRow = idx + 2;

      // Refunded kits need no delivery follow-up
      if (String(row[KIT_COL.REFUNDED - 1] || '').trim().toUpperCase() === 'TRUE') return;

      var jlid          = String(row[KIT_COL.JLID - 1]          || '').trim();
      var deliveryDate  = String(row[KIT_COL.DELIVERY_DATE - 1]  || '').trim();
      var followupSent  = String(row[KIT_COL.FOLLOWUP_SENT - 1]  || '').trim();
      var etaRaw        = row[KIT_COL.ETA - 1];
      var kitName       = String(row[KIT_COL.KIT - 1]            || '').trim();
      var learnerName   = String(row[KIT_COL.LEARNER_NAME - 1]   || '').trim();

      if (deliveryDate) return;
      if (followupSent === 'TRUE' || followupSent === 'true' || followupSent === true) return;
      if (!learnerName || !kitName) return;

      // Only process 2026+ rows � skip if no order date or order date < 2026
      var orderRawCheck  = row[KIT_COL.DATE_OF_ORDER - 1];
      var orderYearCheck = null;
      if (orderRawCheck instanceof Date) orderYearCheck = orderRawCheck.getFullYear();
      else { var od = _parseDMY(String(orderRawCheck || '')); if (od) orderYearCheck = od.getFullYear(); }
      if (!orderYearCheck || orderYearCheck < 2026) return;

      // Auto-fill JLID from HubSpot if missing
      if (!jlid) {
        Logger.log('[KitTracking] Row ' + sheetRow + ': no JLID — auto-searching HubSpot for "' + learnerName + '" / ' + kitName);
        jlid = _findJlidByKitStatus(learnerName, kitName) || '';
        if (jlid) {
          sheet.getRange(sheetRow, KIT_COL.JLID).setValue(jlid);
          Logger.log('[KitTracking] Row ' + sheetRow + ': auto-filled JLID=' + jlid);
        } else {
          Logger.log('[KitTracking] Row ' + sheetRow + ': could not auto-find JLID, skipping.');
          return;
        }
      }

      // Parse ETA — skip if future
      var etaDate = null;
      if (etaRaw instanceof Date) {
        etaDate = etaRaw;
      } else {
        etaDate = _parseDMY(String(etaRaw));
      }
      if (!etaDate) {
        Logger.log('[KitTracking] Row ' + sheetRow + ': cannot parse ETA "' + etaRaw + '", skipping.');
        return;
      }
      etaDate.setHours(0, 0, 0, 0);
      if (etaDate > today) {
        Logger.log('[KitTracking] Row ' + sheetRow + ': ETA ' + _formatDMY(etaDate) + ' is future, skipping.');
        return;
      }

      // Fetch parent phone + name from HubSpot
      var hs = fetchHubspotByJlid(jlid);
      if (!hs || !hs.success || !hs.data.parentContact) {
        Logger.log('[KitTracking] Row ' + sheetRow + ': HubSpot lookup failed for ' + jlid + ' — ' + (hs && hs.message));
        return;
      }

      var phone      = _normalisePhone(hs.data.parentContact);
      var parentName = hs.data.parentName || learnerName || 'Parent';

      if (!phone) {
        Logger.log('[KitTracking] Row ' + sheetRow + ': no phone for ' + jlid + ', skipping.');
        return;
      }

      // Send WATI template
      Logger.log('[KitTracking] Sending follow-up for row ' + sheetRow + ' JLID=' + jlid + ' phone=' + phone + ' kit=' + kitName);
      var watiResult = sendWatiMessage(phone, 'migration_kit_fup_sent_by_us', [
        { name: 'Parent',   value: parentName  },
        { name: 'Kit_name', value: kitName     },
        { name: 'Learner',  value: learnerName }
      ]);
      Logger.log('[KitTracking] WATI result: ' + JSON.stringify(watiResult));

      // Mark sheet
      sheet.getRange(sheetRow, KIT_COL.FOLLOWUP_SENT).setValue('TRUE');
      sheet.getRange(sheetRow, KIT_COL.FOLLOWUP_SENT_AT).setValue(new Date());
      sheet.getRange(sheetRow, KIT_COL.PHONE_SENT_TO).setValue(phone);
      sent++;
    });

    Logger.log('[KitTracking] sendKitFollowUps done. Sent ' + sent + ' first reminders.');

    // -- 2nd reminder pass ---------------------------------------------
    // Rows where: 1st follow-up sent 2+ days ago, no response, no 2nd follow-up sent
    var sent2 = 0;
    var twoDaysAgo = new Date(today);
    twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);

    // Re-read sheet (may have been updated above)
    var rows2 = sheet.getRange(2, 1, lastRow - 1, KIT_COL.FOLLOWUP2_SENT_AT).getValues();

    rows2.forEach(function(row, idx) {
      var sheetRow = idx + 2;

      var deliveryDate   = String(row[KIT_COL.DELIVERY_DATE - 1]   || '').trim();
      var followupSent   = String(row[KIT_COL.FOLLOWUP_SENT - 1]   || '').trim();
      var followup2Sent  = String(row[KIT_COL.FOLLOWUP2_SENT - 1]  || '').trim();
      var parentResponse = String(row[KIT_COL.PARENT_RESPONSE - 1] || '').trim();
      var sentAtRaw      = row[KIT_COL.FOLLOWUP_SENT_AT - 1];
      var learnerName    = String(row[KIT_COL.LEARNER_NAME - 1]    || '').trim();
      var kitName        = String(row[KIT_COL.KIT - 1]             || '').trim();
      var jlid           = String(row[KIT_COL.JLID - 1]            || '').trim();
      var phone          = String(row[KIT_COL.PHONE_SENT_TO - 1]   || '').trim();

      if (deliveryDate)   return; // already delivered
      if (parentResponse) return; // already replied
      if (followup2Sent === 'TRUE' || followup2Sent === 'true' || followup2Sent === true) return; // already 2nd sent
      if (followupSent  !== 'TRUE' && followupSent !== 'true'  && followupSent  !== true)  return; // 1st not sent yet

      // Parse 1st follow-up timestamp
      var sentAt = (sentAtRaw instanceof Date) ? sentAtRaw : (sentAtRaw ? new Date(sentAtRaw) : null);
      if (!sentAt || isNaN(sentAt.getTime())) return;
      sentAt.setHours(0, 0, 0, 0);
      if (sentAt > twoDaysAgo) return; // less than 2 days old

      // Need phone to send
      if (!phone) {
        Logger.log('[KitTracking] 2nd reminder row ' + sheetRow + ': no phone, skipping.');
        return;
      }

      // Need HubSpot for parent name
      var parentName = learnerName || 'Parent';
      if (jlid) {
        try {
          var hs2 = fetchHubspotByJlid(jlid);
          if (hs2 && hs2.success && hs2.data.parentName) parentName = hs2.data.parentName;
        } catch(he) {}
      }

      Logger.log('[KitTracking] Sending 2nd reminder row ' + sheetRow + ' JLID=' + jlid + ' phone=' + phone);
      sendWatiMessage(phone, 'migration_kit_fup_sent_by_us', [
        { name: 'Parent',   value: parentName  },
        { name: 'Kit_name', value: kitName     },
        { name: 'Learner',  value: learnerName }
      ]);

      sheet.getRange(sheetRow, KIT_COL.FOLLOWUP2_SENT).setValue('TRUE');
      sheet.getRange(sheetRow, KIT_COL.FOLLOWUP2_SENT_AT).setValue(new Date());
      sent2++;
    });

    Logger.log('[KitTracking] sendKitFollowUps done. 2nd reminders sent: ' + sent2);

    // ── 3rd pass: Escalate to CLS after 2nd reminder unanswered 2+ days ──────
    var escalated = 0;
    var twoDaysAgo2 = new Date(today);
    twoDaysAgo2.setDate(twoDaysAgo2.getDate() - 2);

    var rows3 = sheet.getRange(2, 1, lastRow - 1, KIT_COL.ESCALATED_AT).getValues();

    rows3.forEach(function(row, idx) {
      var sheetRow = idx + 2;

      var deliveryDate   = String(row[KIT_COL.DELIVERY_DATE  - 1] || '').trim();
      var parentResponse = String(row[KIT_COL.PARENT_RESPONSE - 1] || '').trim();
      var fup2Sent       = String(row[KIT_COL.FOLLOWUP2_SENT  - 1] || '').trim();
      var fup2SentAt     = row[KIT_COL.FOLLOWUP2_SENT_AT - 1];
      var alreadyEsc     = String(row[KIT_COL.ESCALATED       - 1] || '').trim();
      var jlid           = String(row[KIT_COL.JLID            - 1] || '').trim();
      var learnerName    = String(row[KIT_COL.LEARNER_NAME    - 1] || '').trim();
      var kitName        = String(row[KIT_COL.KIT             - 1] || '').trim();

      if (deliveryDate)   return; // delivered
      if (parentResponse) return; // replied
      if (alreadyEsc === 'TRUE' || alreadyEsc === 'true') return; // already escalated
      if (fup2Sent !== 'TRUE' && fup2Sent !== 'true') return; // 2nd reminder not sent yet

      var sentAt2 = (fup2SentAt instanceof Date) ? fup2SentAt : (fup2SentAt ? new Date(fup2SentAt) : null);
      if (!sentAt2 || isNaN(sentAt2.getTime())) return;
      sentAt2.setHours(0,0,0,0);
      if (sentAt2 > twoDaysAgo2) return; // not 2 days old yet

      Logger.log('[KitTracking] Escalating row ' + sheetRow + ' JLID=' + jlid);

      // Fetch learner deal info
      var dealId = '', currentTeacher = '', parentName = learnerName, courseName = '';
      if (jlid) {
        try {
          var hs3 = fetchHubspotByJlid(jlid);
          if (hs3 && hs3.success && hs3.data) {
            dealId         = hs3.data.dealId        || '';
            currentTeacher = hs3.data.currentTeacher || '';
            parentName     = hs3.data.parentName    || learnerName;
            courseName     = hs3.data.courseName    || '';
          }
        } catch(he) { Logger.log('[KitTracking] esc HubSpot fetch err: ' + he.message); }
      }

      // Look up CLS + TP manager emails from teacher sheet
      var mgrs = _kitGetTeacherManagerEmails(currentTeacher);

      // Send escalation email
      _sendKitEscalationEmail({
        jlid: jlid, learnerName: learnerName, parentName: parentName,
        kitName: kitName, courseName: courseName, currentTeacher: currentTeacher,
        clsManagerEmail: mgrs.clsEmail, clsManagerName: mgrs.clsName,
        tpManagerEmail:  mgrs.tpEmail
      });

      // Create HubSpot task on deal
      if (dealId) {
        _createKitEscalationTask(dealId, { jlid: jlid, learnerName: learnerName, kitName: kitName, courseName: courseName, clsManagerEmail: mgrs.clsEmail });
      }

      // Update HubSpot kit status enum to "Escalated to CLS" (VR/Microbit/Makey Makey only —
      // Arduino's status property has no such option, so it's skipped for that kit).
      if (dealId) {
        try {
          var escKitProp = _kitPropertyForType(kitName);
          if (escKitProp && escKitProp !== 'arduino_kit_status') {
            var escToken = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
            var escProps = {}; escProps[escKitProp] = 'Escalated to CLS';
            monitoredFetch('https://api.hubapi.com/crm/v3/objects/deals/' + dealId, {
              method: 'PATCH',
              headers: { 'Authorization': 'Bearer ' + escToken, 'Content-Type': 'application/json' },
              payload: JSON.stringify({ properties: escProps }),
              muteHttpExceptions: true
            });
          }
        } catch(escErr) { Logger.log('[KitTracking] escalation HS status patch error: ' + escErr.message); }
      }

      // Mark escalated in sheet
      sheet.getRange(sheetRow, KIT_COL.ESCALATED).setValue('TRUE');
      sheet.getRange(sheetRow, KIT_COL.ESCALATED_AT).setValue(new Date());
      escalated++;
    });

    Logger.log('[KitTracking] sendKitFollowUps done. Escalations: ' + escalated);
  } catch (e) {
    Logger.log('[KitTracking] sendKitFollowUps ERROR: ' + e.message + '\n' + e.stack);
  }
}

// ── Kit escalation helpers ────────────────────────────────────────────────────

function _kitGetTeacherManagerEmails(teacherName) {
  var result = { clsName: '', clsEmail: '', tpName: '', tpEmail: '' };
  if (!teacherName) return result;
  try {
    var sheet = _getCachedSheetData(CONFIG.SHEETS.TEACHER_DATA);
    var nameLower = String(teacherName).trim().toLowerCase();
    for (var r = 1; r < sheet.length; r++) {
      var rowName = String(sheet[r][1] || '').trim().toLowerCase();
      if (rowName === nameLower) {
        result.tpName  = String(sheet[r][6]  || '').trim();
        result.clsName = String(sheet[r][7]  || '').trim();
        result.clsEmail = String(sheet[r][9] || '').trim();
        result.tpEmail  = String(sheet[r][10]|| '').trim();
        break;
      }
    }
  } catch(e) {
    Logger.log('[KitTracking] _kitGetTeacherManagerEmails err: ' + e.message);
  }
  return result;
}

function _sendKitEscalationEmail(d) {
  try {
    var to  = d.clsManagerEmail || CONFIG.EMAIL.MAIN_MANAGER;
    var cc  = ['sourav.pal@jet-learn.com'];
    if (d.tpManagerEmail) cc.push(d.tpManagerEmail);
    var ccStr = cc.join(',');

    var subject = '[Action Required] Kit Not Received — ' + d.learnerName + ' (' + d.jlid + ')';
    var body =
      '<p>Hi ' + (d.clsManagerName || 'Team') + ',</p>' +
      '<p>The kit follow-up sequence for the learner below has completed. ' +
      'Two WhatsApp reminders were sent but <strong>no response was received</strong>. ' +
      'The kit has not been confirmed as delivered.</p>' +
      '<table style="border-collapse:collapse;font-family:sans-serif;font-size:14px;">' +
      '<tr><td style="padding:4px 16px 4px 0"><strong>JLID</strong></td><td>' + (d.jlid || '—') + '</td></tr>' +
      '<tr><td style="padding:4px 16px 4px 0"><strong>Learner</strong></td><td>' + (d.learnerName || '—') + '</td></tr>' +
      '<tr><td style="padding:4px 16px 4px 0"><strong>Parent</strong></td><td>' + (d.parentName || '—') + '</td></tr>' +
      '<tr><td style="padding:4px 16px 4px 0"><strong>Kit</strong></td><td>' + (d.kitName || '—') + '</td></tr>' +
      '<tr><td style="padding:4px 16px 4px 0"><strong>Current Teacher</strong></td><td>' + (d.currentTeacher || '—') + '</td></tr>' +
      '<tr><td style="padding:4px 16px 4px 0"><strong>Course</strong></td><td>' + (d.courseName || '—') + '</td></tr>' +
      '</table>' +
      '<p><strong>⚠️ Recommended action:</strong> If the kit cannot be confirmed, ' +
      'please consider <strong>skipping the course module</strong> that requires this kit ' +
      'and update the learner\'s roadmap accordingly.</p>' +
      '<p>A task has been created in HubSpot on this learner\'s deal.</p>' +
      '<p>— JetLearn Operations Platform</p>';

    MailApp.sendEmail({ to: to, cc: ccStr, subject: subject, htmlBody: body,
      name: CONFIG.EMAIL.FROM_NAME || 'JetLearn Ops', from: CONFIG.EMAIL.FROM });
    Logger.log('[KitTracking] Escalation email → ' + to + ' CC: ' + ccStr);
  } catch(e) {
    Logger.log('[KitTracking] _sendKitEscalationEmail ERROR: ' + e.message);
  }
}

function _createKitEscalationTask(dealId, d) {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY') || '';
    if (!token || !dealId) return;
    var due = new Date(); due.setDate(due.getDate() + 1);
    var taskProps = {
      hs_task_subject:  '[Kit Not Received] ' + d.learnerName + ' — ' + d.kitName,
      hs_task_body:     'Kit has not been confirmed as delivered after 2 WhatsApp follow-ups.\n\n' +
                        'JLID: ' + d.jlid + '\nKit: ' + d.kitName +
                        (d.courseName ? '\nCourse: ' + d.courseName : '') +
                        '\n\nAction: Contact parent directly. If unresolved, skip the course module that requires this kit and update the learner\'s roadmap.',
      hs_task_status:   'NOT_STARTED',
      hs_task_priority: 'HIGH',
      hs_task_type:     'TODO',
      hs_timestamp:     due.getTime()
    };
    var clsOwnerId = d.clsManagerEmail
      ? _resolveHubSpotOwnerIdByEmail(d.clsManagerEmail, token) : null;
    if (clsOwnerId) taskProps.hubspot_owner_id = clsOwnerId;
    var payload = {
      properties: taskProps,
      associations: [{
        to:    { id: String(dealId) },
        types: [{ associationCategory: 'HUBSPOT_DEFINED', associationTypeId: 216 }]
      }]
    };
    monitoredFetch('https://api.hubapi.com/crm/v3/objects/tasks', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    Logger.log('[KitTracking] HubSpot task created for deal ' + dealId);
  } catch(e) {
    Logger.log('[KitTracking] _createKitEscalationTask ERROR: ' + e.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// HANDLE KIT REPLY  (called from doPost WATI webhook)
// ─────────────────────────────────────────────────────────────────────────────
function handleKitReply(waId, buttonText) {
  Logger.log('[KitTracking] handleKitReply waId=' + waId + ' btn="' + buttonText + '"');
  try {
    var normPhone = _normalisePhone(waId);
    if (!normPhone) {
      Logger.log('[KitTracking] No phone in waId, aborting.');
      return;
    }

    var sheet   = _getKitSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var rows = sheet.getRange(2, 1, lastRow - 1, KIT_COL.PHONE_SENT_TO).getValues();

    // Collect ALL pending rows for this phone (siblings share same parent phone)
    var matchedRows = [];
    rows.forEach(function(row, idx) {
      var phone        = _normalisePhone(String(row[KIT_COL.PHONE_SENT_TO - 1] || ''));
      var deliveryDate = String(row[KIT_COL.DELIVERY_DATE - 1] || '').trim();
      if (phone && phone === normPhone && !deliveryDate) matchedRows.push(idx + 2);
    });

    if (!matchedRows.length) {
      Logger.log('[KitTracking] No matching pending row for phone ' + normPhone);
      return;
    }
    Logger.log('[KitTracking] Found ' + matchedRows.length + ' pending row(s) for phone ' + normPhone);

    matchedRows.forEach(function(matchRow) {
    var dataRow      = rows[matchRow - 2];
    var jlid         = String(dataRow[KIT_COL.JLID - 1]         || '').trim();
    var kitName      = String(dataRow[KIT_COL.KIT - 1]          || '').trim();
    var learnerName  = String(dataRow[KIT_COL.LEARNER_NAME - 1] || '').trim();
    var orderDateRaw = dataRow[KIT_COL.DATE_OF_ORDER - 1];
    var normJlid = jlid.replace(/[^A-Z0-9]$/i, '').trim();
    Logger.log('[KitTracking] Processing row ' + matchRow + ' JLID=' + normJlid + ' kit=' + kitName);

    // Write parent response
    sheet.getRange(matchRow, KIT_COL.PARENT_RESPONSE).setValue(buttonText);

    if (buttonText === 'Kit Received') {
      var today = new Date();

      // Fill Delivery Date + Time Taken
      sheet.getRange(matchRow, KIT_COL.DELIVERY_DATE).setValue(_formatDMY(today));
      var orderDate = (orderDateRaw instanceof Date) ? orderDateRaw : _parseDMY(String(orderDateRaw || ''));
      if (orderDate) {
        var diffDays = Math.round((today - orderDate) / (1000 * 60 * 60 * 24));
        sheet.getRange(matchRow, KIT_COL.TIME_TAKEN).setValue(diffDays + ' days');
      }

      // Update HubSpot kit status � internal enum value is "Received" (label: "Received by the Parents")
      if (normJlid) _updateHubspotKitStatus(normJlid, kitName, 'Received');

      // Auto-reply to parent - confirm via template (session messages unreliable)
      try {
        var confirmResult = sendWatiMessage(normPhone, 'kit_received_confirmation', [
          { name: 'kit',     value: kitName     },
          { name: 'learner', value: learnerName }
        ]);
        Logger.log('[KitTracking] Confirmation template result: ' + JSON.stringify(confirmResult));
      } catch(re) { Logger.log('[KitTracking] Auto-reply error: ' + re.message); }


      Logger.log('[KitTracking] Kit confirmed received � row ' + matchRow + ' updated.');

    } else if (buttonText === 'Not Received yet') {
      // Add HubSpot note
      if (normJlid) {
        var hs1 = fetchHubspotByJlid(normJlid);
        if (hs1 && hs1.success && hs1.data.dealId) {
          _addNoteToDeal(hs1.data.dealId,
            '[Kit Follow-up] Parent replied �Not Received yet� for kit: ' + kitName +
            ' on ' + _formatDMY(new Date()) + '. Verify with logistics and update parent.');
        }
      }
      // Auto-reply via template
      try {
        sendWatiMessage(normPhone, 'kit_not_received_reply', [
          { name: 'kit',     value: kitName     },
          { name: 'learner', value: learnerName }
        ]);
      } catch(re) { Logger.log('[KitTracking] Auto-reply error: ' + re.message); }

      Logger.log('[KitTracking] Not received � row ' + matchRow + ' flagged, HS note added.');

    } else if (buttonText === 'Need To check') {
      // Add HubSpot note
      if (normJlid) {
        var hs2 = fetchHubspotByJlid(normJlid);
        if (hs2 && hs2.success && hs2.data.dealId) {
          _addNoteToDeal(hs2.data.dealId,
            '[Kit Follow-up] Parent replied �Need To Check� for kit: ' + kitName +
            ' on ' + _formatDMY(new Date()) + '. Awaiting parent confirmation in 12-24 hrs.');
        }
      }
      // Auto-reply via template
      try {
        sendWatiMessage(normPhone, 'kit_need_to_check_reply', [
          { name: 'kit',     value: kitName     },
          { name: 'learner', value: learnerName }
        ]);
      } catch(re) { Logger.log('[KitTracking] Auto-reply error: ' + re.message); }

      Logger.log('[KitTracking] Need to check � row ' + matchRow + ' flagged, HS note added.');
    }

    }); // end matchedRows.forEach
  } catch (e) {
    Logger.log('[KitTracking] handleKitReply ERROR: ' + e.message + '\n' + e.stack);
  }
}

// ── ONE-TIME DIAGNOSTIC — find correct HubSpot internal names + enum values ───
// Run from GAS editor → check Execution Log
function diagKitHubSpotProperties() {
  var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  if (!token) { Logger.log('[KitDiag] No HUBSPOT_API_KEY set'); return; }

  // 1. Check deal property names for learning kit cost
  var costProps = ['learning_kit_cost', 'learning_kits_total_cost', 'total_learning_kit_cost', 'kit_cost'];
  Logger.log('[KitDiag] === Checking Learning Kit Cost property names ===');
  costProps.forEach(function(prop) {
    Utilities.sleep(800);
    var r = monitoredFetch('https://api.hubapi.com/crm/v3/properties/deals/' + prop, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (r.getResponseCode() === 200) {
      var d = JSON.parse(r.getContentText());
      Logger.log('[KitDiag] FOUND: "' + prop + '" | label: "' + d.label + '" | type: ' + d.fieldType);
    } else {
      Logger.log('[KitDiag] NOT FOUND: "' + prop + '" (HTTP ' + r.getResponseCode() + ')');
    }
  });

  // 2. Check kit status enum values
  Logger.log('[KitDiag] === Checking kit status enum values ===');
  var kitProps = ['microbit_kit_status', 'makey_makey_kit_status', 'vr_headset__oculus_status', 'arduino_kit_status'];
  kitProps.forEach(function(prop) {
    Utilities.sleep(800);
    var r = monitoredFetch('https://api.hubapi.com/crm/v3/properties/deals/' + prop, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (r.getResponseCode() !== 200) {
      Logger.log('[KitDiag] "' + prop + '" HTTP ' + r.getResponseCode() + ' (not found)');
      return;
    }
    var d = JSON.parse(r.getContentText());
    var options = (d.options || []).map(function(o) { return '"' + o.value + '" -> ' + o.label; });
    Logger.log('[KitDiag] "' + prop + '":\n  ' + (options.length ? options.join('\n  ') : '(no options)'));
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// ONE-TIME DIAGNOSTIC � run manually to see valid enum values for kit properties
function diagKitEnumValues() {
  var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
  // Use exact property names from _kitPropertyForType + common variants
  var props = [
    'microbit_kit_status',
    'vr_headset__oculus_status',    // double underscore
    'vr_headset_oculus_status',     // single underscore variant
    'makey_makey_kit_status',       // without __t_
    'makey_makey_kit_status__t_',   // with __t_
    'arduino_kit_status'
  ];
  props.forEach(function(prop) {
    Utilities.sleep(1500); // avoid bandwidth quota
    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/properties/deals/' + prop, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code !== 200) { Logger.log('[KitEnum] ' + prop + ': HTTP ' + code + ' (not found or error)'); return; }
    var data = JSON.parse(resp.getContentText());
    var options = (data.options || []).map(function(o) { return o.label + ' ? "' + o.value + '"'; });
    Logger.log('[KitEnum] ' + prop + ':\n  ' + (options.length ? options.join('\n  ') : '(no options � text field?)'));
  });
}

// UPDATE HUBSPOT KIT STATUS  (PATCH deal property)
// ─────────────────────────────────────────────────────────────────────────────
function _updateHubspotKitStatus(jlid, kitName, statusValue) {
  try {
    var prop = _kitPropertyForType(kitName);
    if (!prop) {
      Logger.log('[KitTracking] Unknown kit type "' + kitName + '" — no HubSpot property to update.');
      return;
    }

    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success || !hs.data.dealId) {
      Logger.log('[KitTracking] Cannot update HubSpot — no dealId for ' + jlid);
      return;
    }

    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    if (!token) {
      Logger.log('[KitTracking] HUBSPOT_API_KEY not set.');
      return;
    }

    var url = 'https://api.hubapi.com/crm/v3/objects/deals/' + hs.data.dealId;
    var resp = monitoredFetch(url, {
      method: 'PATCH',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ properties: { [prop]: statusValue } }),
      muteHttpExceptions: true
    });

    var code = resp.getResponseCode();
    Logger.log('[KitTracking] HubSpot PATCH ' + prop + '="' + statusValue + '" → HTTP ' + code);
    if (code !== 200) {
      Logger.log('[KitTracking] HubSpot error body: ' + resp.getContentText().substring(0, 300));
    }
  } catch (e) {
    Logger.log('[KitTracking] _updateHubspotKitStatus ERROR: ' + e.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// GET KIT STATUS ENUM VALUES — callable from client to discover valid HubSpot values
// ─────────────────────────────────────────────────────────────────────────────
function getKitStatusEnums() {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    if (!token) return { success: false, message: 'No API key' };
    var props = ['microbit_kit_status', 'makey_makey_kit_status', 'vr_headset__oculus_status', 'arduino_kit_status'];
    var results = {};
    props.forEach(function(prop) {
      Utilities.sleep(500);
      var resp = monitoredFetch('https://api.hubapi.com/crm/v3/properties/deals/' + prop, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });
      var code = resp.getResponseCode();
      if (code === 200) {
        var d = JSON.parse(resp.getContentText());
        results[prop] = (d.options || []).map(function(o) { return { label: o.label, value: o.value }; });
      } else {
        results[prop] = 'HTTP ' + code;
      }
    });
    return { success: true, enums: results };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// ADD NOTE TO HUBSPOT DEAL
// ─────────────────────────────────────────────────────────────────────────────
function _addNoteToDeal(dealId, noteBody) {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
    if (!token || !dealId || !noteBody) return;

    // Create note via CRM v3 engagements (Notes)
    var payload = {
      properties: {
        hs_note_body:      noteBody,
        hs_timestamp:      String(new Date().getTime())
      }
    };
    var resp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/notes', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var noteId = null;
    try { noteId = JSON.parse(resp.getContentText()).id; } catch(e2) {}

    // Associate note to deal
    if (noteId) {
      monitoredFetch(
        'https://api.hubapi.com/crm/v3/objects/notes/' + noteId + '/associations/deals/' + dealId + '/note_to_deal',
        {
          method: 'put',
          headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
          payload: '{}',
          muteHttpExceptions: true
        }
      );
    }
    Logger.log('[KitTracking] Note added to deal ' + dealId + ' (noteId=' + noteId + ')');
  } catch (e) {
    Logger.log('[KitTracking] _addNoteToDeal ERROR: ' + e.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// GET KIT TRACKING DATA  (called from client dashboard)
// ─────────────────────────────────────────────────────────────────────────────
function getKitTrackingData() {
  try {
    var sheet   = _getKitSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, rows: [], stats: { total: 0, delivered: 0, awaiting: 0, notReceived: 0, overdue: 0 } };

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var raw  = sheet.getRange(2, 1, lastRow - 1, KIT_COL.REFUNDED).getValues();
    var rows = [];

    var cutoff = new Date(2026, 0, 1); // Jan 1 2026 — ignore older rows

    raw.forEach(function(r, idx) {
      var srNo         = r[KIT_COL.SR_NO - 1];
      var learnerName  = String(r[KIT_COL.LEARNER_NAME - 1]   || '').trim();
      var kit          = String(r[KIT_COL.KIT - 1]            || '').trim();
      var orderRaw     = r[KIT_COL.DATE_OF_ORDER - 1];
      var tsMonth      = String(r[7]                           || '').trim(); // col H = Timestamp month
      var etaRaw       = r[KIT_COL.ETA - 1];
      // Delivery date — handle both Date objects and DD/MM/YYYY strings
      var deliveryRaw  = r[KIT_COL.DELIVERY_DATE - 1];
      var deliveryDate = '';
      if (deliveryRaw instanceof Date && !isNaN(deliveryRaw.getTime())) {
        deliveryDate = _formatDMY(deliveryRaw);
      } else if (deliveryRaw) {
        var dStr = String(deliveryRaw).trim();
        // Strip full timestamp if present (e.g. "Fri Apr 03 2026 00:00:00 GMT+0530...")
        if (dStr.length > 10 && dStr.indexOf('/') === -1) {
          var parsed = new Date(dStr);
          deliveryDate = isNaN(parsed.getTime()) ? dStr : _formatDMY(parsed);
        } else {
          deliveryDate = dStr;
        }
      }
      var timeTaken    = String(r[KIT_COL.TIME_TAKEN - 1]     || '').trim();
      var followupSent  = String(r[KIT_COL.FOLLOWUP_SENT - 1]   || '').trim();
      var sentAt        = r[KIT_COL.FOLLOWUP_SENT_AT - 1];
      var response      = String(r[KIT_COL.PARENT_RESPONSE - 1] || '').trim();
      var jlid          = String(r[KIT_COL.JLID - 1]            || '').trim();
      var followup2Sent = String(r[KIT_COL.FOLLOWUP2_SENT - 1]  || '').trim();
      var sentAt2       = r[KIT_COL.FOLLOWUP2_SENT_AT - 1];

      if (!learnerName && !kit) return; // blank row

      // Parse order date — skip rows before Jan 2026
      var orderDate = null;
      if (orderRaw instanceof Date) {
        orderDate = orderRaw;
      } else if (orderRaw) {
        orderDate = _parseDMY(String(orderRaw));
      }
      if (!orderDate || orderDate < cutoff) return;

      // Build orderMonth label e.g. �April 2026� � always from parsed orderDate (never col H formula)
      var MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      var orderMonth = orderDate ? (MONTHS[orderDate.getMonth()] + ' ' + orderDate.getFullYear()) : '';

      // Parse ETA
      var etaStr = '';
      var etaDate = null;
      if (etaRaw instanceof Date) {
        etaDate = etaRaw;
        etaStr  = _formatDMY(etaRaw);
      } else if (etaRaw) {
        etaDate = _parseDMY(String(etaRaw));
        etaStr  = String(etaRaw).trim();
      }
      if (etaDate) etaDate.setHours(0, 0, 0, 0);

      var fupSentBool  = (followupSent  === 'TRUE' || followupSent  === 'true');
      var fup2SentBool = (followup2Sent === 'TRUE' || followup2Sent === 'true');
      var sentAtStr    = '';
      if (sentAt instanceof Date) sentAtStr = _formatDMY(sentAt);
      else if (sentAt) sentAtStr = String(sentAt).trim();
      var sentAt2Str = '';
      if (sentAt2 instanceof Date) sentAt2Str = _formatDMY(sentAt2);
      else if (sentAt2) sentAt2Str = String(sentAt2).trim();

      // Days since 2nd follow-up
      var daysSince2nd = null;
      if (fup2SentBool && sentAt2) {
        var t2 = (sentAt2 instanceof Date) ? sentAt2 : new Date(sentAt2);
        t2.setHours(0,0,0,0);
        daysSince2nd = Math.floor((today - t2) / 86400000);
      }

      // Compute status — refunded wins over everything (no ETA/overdue tracking applies)
      var isRefunded = String(r[26] || '').trim().toUpperCase() === 'TRUE';
      var status = 'pending';
      if (isRefunded) {
        status = 'refunded';
      } else if (deliveryDate || response === 'Kit Received') {
        status = 'delivered';
      } else if (response === 'Not Received yet') {
        status = 'not_received';
      } else if (response === 'Need To check') {
        status = 'need_check';
      } else if (fup2SentBool && !response) {
        // 2nd reminder sent, still no reply ? escalated
        status = 'escalated';
      } else if (fupSentBool && !response) {
        status = 'awaiting';
      } else if (!fupSentBool && etaDate && etaDate <= today) {
        status = 'overdue';
      }

      rows.push({
        rowIndex:      idx + 2,
        srNo:          srNo || (idx + 1),
        learnerName:   learnerName,
        kit:           kit,
        orderMonth:    orderMonth,
        eta:           etaStr,
        deliveryDate:  deliveryDate,
        timeTaken:     timeTaken,
        followupSent:  fupSentBool,
        sentAt:        sentAtStr,
        followup2Sent: fup2SentBool,
        sentAt2:       sentAt2Str,
        daysSince2nd:  daysSince2nd,
        response:      response,
        jlid:          jlid,
        status:        status,
        refunded:      isRefunded,
        // Extra detail fields
        country:      String(r[3]  || '').trim(),   // D: Country
        price:        String(r[4]  || '').trim(),   // E: Price EUR
        site:         String(r[5]  || '').trim(),   // F: Site
        orderDate:    (function() {
          var v = r[6];
          if (v instanceof Date) return _formatDMY(v);
          return String(v || '').trim();
        })(),                                        // G: Date of Order
        reason:       String(r[11] || '').trim(),   // L: Reason
        subscription: String(r[12] || '').trim(),   // M: Subscription
        roadmap:      String(r[13] || '').trim(),   // N: Roadmap
        sentBy:       String(r[14] || '').trim(),   // O: Sent By
        phone:        String(r[19] || '').replace(/\D/g, '')  // T: Phone Sent To
      });
    });

    // Stats
    var stats = {
      total:       rows.length,
      delivered:   rows.filter(function(r) { return r.status === 'delivered'; }).length,
      awaiting:    rows.filter(function(r) { return r.status === 'awaiting'; }).length,
      notReceived: rows.filter(function(r) { return r.status === 'not_received' || r.status === 'need_check'; }).length,
      overdue:     rows.filter(function(r) { return r.status === 'overdue'; }).length,
      escalated:   rows.filter(function(r) { return r.status === 'escalated'; }).length,
      refunded:    rows.filter(function(r) { return r.status === 'refunded'; }).length
    };

    Logger.log('[KitTracking] getKitTrackingData: ' + rows.length + ' rows, stats=' + JSON.stringify(stats));
    return { success: true, rows: rows, stats: stats };

  } catch (e) {
    Logger.log('[KitTracking] getKitTrackingData ERROR: ' + e.message);
    return { success: false, message: e.message, rows: [], stats: {} };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// RESEND KIT FOLLOW-UP  (manual resend from dashboard)
// ─────────────────────────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────
// FETCH LEARNER DETAILS FOR KIT FORM  (JLID lookup in Add Kit modal)
// ─────────────────────────────────────────────────────────────────────────────
function fetchKitLearnerDetails(jlid) {
  try {
    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success) return { success: false, message: (hs && hs.message) || 'Learner not found for JLID: ' + jlid };
    var d = hs.data;

    // Fetch delivery address from HubSpot contact
    var addrObj  = d.dealId ? _fetchContactAddress(d.dealId) : {};
    var addrStr  = [addrObj.address, addrObj.city, addrObj.state, addrObj.zip, addrObj.country]
      .filter(function(p) { return p && String(p).trim(); }).join(', ');
    var hasAddr  = !!(addrStr);

    // Check if address request already pending in cache
    var phone        = _normalisePhone(d.parentContact || '');
    var addrPending  = phone ? !!(CacheService.getScriptCache().get('KIT_ADDR_REQ_' + phone)) : false;

    // >1 subscription line item on the deal = renewed learner
    var lineItemCount = d.dealId ? _countDealLineItems(d.dealId) : 0;
    var isRenewed      = lineItemCount > 1;

    return {
      success:          true,
      isRenewedLearner: isRenewed,
      lineItemCount:    lineItemCount,
      learnerName:      d.learnerName      || '',
      subscription:     d.planName         || '',
      country:          d.country          || '',
      kitCostSoFar:     d.learningKitCost  || 0,
      dealId:           d.dealId           || '',
      phone:            phone,
      deliveryAddress:  addrStr,       // existing address shown as reference only — always reconfirm
      hasAddress:       hasAddr,        // true = address exists but still needs reconfirmation
      addressPending:   addrPending,    // WATI already sent this session, awaiting new reply
      addressStale:     hasAddr         // always treat existing address as potentially stale
    };
  } catch (e) {
    Logger.log('[KitTracking] fetchKitLearnerDetails ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// ADD KIT ENTRY  (called from client Add Kit modal)
// ─────────────────────────────────────────────────────────────────────────────
function addKitEntry(data) {
  try {
    var sheet   = _getKitSheet();
    var sheetLastRow = sheet.getLastRow();

    // Find actual last data row � scan col A from bottom, skip empty/formatted rows
    var lastDataRow = 1;
    if (sheetLastRow >= 2) {
      var colA = sheet.getRange(2, 1, sheetLastRow - 1, 1).getValues();
      for (var i = colA.length - 1; i >= 0; i--) {
        if (String(colA[i][0]).trim() !== '') { lastDataRow = i + 2; break; }
      }
    }

    // Auto Sr No = highest existing Sr No + 1
    var lastSrNo = 0;
    if (sheetLastRow >= 2) {
      var srValues = sheet.getRange(2, KIT_COL.SR_NO, sheetLastRow - 1, 1).getValues();
      srValues.forEach(function(r) {
        var v = parseInt(r[0], 10);
        if (!isNaN(v) && v > lastSrNo) lastSrNo = v;
      });
    }
    var srNo = lastSrNo + 1;

    // Auto Timestamp Month from order date  e.g. "April-26"
    var MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    var SHORT_MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var tsMonth = '';
    var orderDate = null;
    if (data.orderDate) {
      // data.orderDate arrives as "YYYY-MM-DD" from HTML date input
      var parts = data.orderDate.split('-');
      if (parts.length === 3) {
        orderDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
        tsMonth = MONTHS[orderDate.getMonth()] + '-' + String(orderDate.getFullYear()).slice(2);
        // e.g. "April-26"
      }
    }

    // Format dates DD/MM/YYYY
    var fmtDate = function(iso) {
      if (!iso) return '';
      var p = iso.split('-');
      if (p.length !== 3) return iso;
      return p[2] + '/' + p[1] + '/' + p[0];
    };

    var jlid = String(data.jlid || '').trim();

    // Write per-column to skip col H (formulated in sheet � never overwrite)
    var writeRow = lastDataRow + 1;
    var writeMap = [
      [1,  srNo],                        // A: Sr No
      [2,  data.learnerName || ''],      // B: Learner's name
      [3,  data.kit         || ''],      // C: Kit
      [4,  data.country     || ''],      // D: Country
      [5,  data.price       || ''],      // E: Price (EUR)
      [6,  data.site        || ''],      // F: Site
      [7,  fmtDate(data.orderDate)],     // G: Date of Order
      // col 8 (H) = SKIP � formula in sheet
      [9,  fmtDate(data.eta)],           // I: ETA
      [10, fmtDate(data.deliveryDate)],  // J: Delivery Date
      [12, data.reason       || ''],     // L: Reason
      [13, data.subscription || ''],     // M: Current Subscription
      [14, data.roadmap      || ''],     // N: Roadmap
      [15, data.sentBy       || ''],     // O: Name
      [16, jlid],                        // P: JLID
      [KIT_COL.DELIVERY_ADDRESS, data.deliveryAddress || ''],  // W: Delivery Address
      [KIT_COL.ADDR_STATUS,      data.deliveryAddress ? 'Verified' : '']  // X: Addr Status
    ];
    writeMap.forEach(function(pair) {
      sheet.getRange(writeRow, pair[0]).setValue(pair[1]);
    });
    Logger.log('[KitTracking] addKitEntry: wrote to row ' + writeRow + ' sr=' + srNo);
    Logger.log('[KitTracking] addKitEntry: row added sr=' + srNo + ' learner=' + data.learnerName + ' jlid=' + jlid + ' row=' + writeRow);

    // ── HubSpot updates on kit entry ──────────────────────────────────
    var price = parseFloat(data.price) || 0;
    var hsUpdated = false, hsWarning = '';
    if (!jlid) {
      hsWarning = 'No JLID — HubSpot status/cost and parent WhatsApp were skipped.';
    } else {
      try {
        var hs = fetchHubspotByJlid(jlid);
        if (!hs || !hs.success || !hs.data || !hs.data.dealId) {
          hsWarning = 'HubSpot deal not found for ' + jlid + ' — status/cost not updated.';
        }
        if (hs && hs.success && hs.data && hs.data.dealId) {
          var token    = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
          var dealId   = hs.data.dealId;
          var kitProp  = _kitPropertyForType(data.kit || '');
          var hsProps  = {};

          // Kit status → "Sent by Us" (internal enum = "Sent")
          if (kitProp) hsProps[kitProp] = 'Sent';

          // Accumulate learning_kit_cost — fetch current value directly from deal (not search cache)
          if (price > 0) {
            var existing = _getHubSpotDealKitCost(dealId, token);
            hsProps['learning_kit_cost'] = existing + price;
            Logger.log('[KitTracking] learning_kit_cost: existing=' + existing + ' + new=' + price + ' = ' + hsProps['learning_kit_cost']);
          }

          if (!kitProp) {
            hsWarning = (hsWarning ? hsWarning + ' ' : '') + 'Kit type "' + (data.kit||'') + '" has no HubSpot status property mapped.';
          }

          // PATCH 1: kit status + cost (never blocked by subscription enum issues)
          if (Object.keys(hsProps).length > 0) {
            var patchResp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/deals/' + dealId, {
              method: 'PATCH',
              headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
              payload: JSON.stringify({ properties: hsProps }),
              muteHttpExceptions: true
            });
            var patchCode = patchResp.getResponseCode();
            Logger.log('[KitTracking] addKitEntry HubSpot PATCH HTTP ' + patchCode + ' props=' + JSON.stringify(hsProps));
            if (patchCode !== 200) {
              var patchErr = patchResp.getContentText();
              Logger.log('[KitTracking] addKitEntry PATCH ERROR: ' + patchErr.substring(0, 500));
              hsWarning = (hsWarning ? hsWarning + ' ' : '') + 'HubSpot update failed (HTTP ' + patchCode + '): ' + patchErr.substring(0, 200);
            } else {
              hsUpdated = true;
            }
          }

          // PATCH 2: subscription — mapped to HubSpot's internal enum values.
          // Form dropdown uses "Yearly"/"2 Yearly"/"3 Yearly"/"Half Yearly"; HS wants
          // "Annual"/"2 Years"/"3 years"/"Half-Yearly". A bad value here only skips
          // subscription, never the status/cost above.
          if (data.subscription) {
            var _SUB_HS_MAP = {
              'yearly': 'Annual', 'annual': 'Annual', '1 year': 'Annual',
              '2 yearly': '2 Years', '2 years': '2 Years',
              '3 yearly': '3 years', '3 years': '3 years',
              '4 yearly': '4 years', '4 years': '4 years',
              'half yearly': 'Half-Yearly', 'half-yearly': 'Half-Yearly',
              'quarterly': 'Quarterly', 'monthly': 'Monthly',
              'credit transfer': 'Credit Transfer'
            };
            var subKey = String(data.subscription).toLowerCase().replace(/[-_]/g, ' ').replace(/\s+/g, ' ').trim();
            var hsSub = _SUB_HS_MAP[subKey] || (String(data.subscription).indexOf('GCSE') === 0 ? data.subscription : null);
            if (hsSub) {
              var subResp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/deals/' + dealId, {
                method: 'PATCH',
                headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
                payload: JSON.stringify({ properties: { subscription: hsSub } }),
                muteHttpExceptions: true
              });
              Logger.log('[KitTracking] addKitEntry subscription PATCH "' + hsSub + '" HTTP ' + subResp.getResponseCode());
              if (subResp.getResponseCode() !== 200) {
                Logger.log('[KitTracking] subscription PATCH ERROR: ' + subResp.getContentText().substring(0, 300));
              }
            } else {
              Logger.log('[KitTracking] subscription "' + data.subscription + '" not mappable to HS enum — skipped.');
            }
          }
        }
      } catch (hsErr) {
        hsWarning = (hsWarning ? hsWarning + ' ' : '') + 'HubSpot update error: ' + hsErr.message;
        Logger.log('[KitTracking] HubSpot kit update ERROR (non-fatal): ' + hsErr.message);
      }
    }

    // ── Send WATI kit-sent confirmation to parent ──────────────────────────
    var watiSent = false;
    var noteSaved = false;
    if (jlid) {
      try {
        var hsKitOrder = fetchHubspotByJlid(jlid);
        if (hsKitOrder && hsKitOrder.success && hsKitOrder.data) {
          var kd = hsKitOrder.data;
          var kitPhone = _normalisePhone(kd.parentContact || '');
          var kitParentName = kd.parentName || data.learnerName || '';
          // Use address passed from UI (collected from parent or HubSpot) — only fall back to HS fetch if not provided
          var deliveryAddrStr = '';
          if (data.deliveryAddress && String(data.deliveryAddress).trim()) {
            deliveryAddrStr = 'Delivery to ' + (data.learnerName || '') + ', ' + String(data.deliveryAddress).trim();
          } else {
            var contactAddr = kd.dealId ? _fetchContactAddress(kd.dealId) : {};
            deliveryAddrStr = _buildKitAddress(data.learnerName || '', contactAddr, kd.country || data.country || '');
          }

          // Format delivery date for WATI (ETA)
          var watiDeliveryDate = data.eta || '';

          if (kitPhone) {
            try {
              sendWatiMessage(kitPhone, 'migration_kit_sent_by_us_parent_information', [
                { name: 'Parent',        value: kitParentName     },
                { name: 'Kit_name',      value: data.kit || ''    },
                { name: 'Delivery_date', value: watiDeliveryDate  },
                { name: 'Address',       value: deliveryAddrStr   }
              ]);
              watiSent = true;
            } catch(we) { Logger.log('[addKitEntry] WATI error: ' + we.message); }
          }

          // Add HubSpot note
          if (kd.dealId) {
            try {
              var noteLines = [
                'Kit Order Entry',
                'Ordered: ' + (data.orderDate || ''),
                'Kit: ' + (data.kit || '') + (data.price ? ' — €' + data.price : ''),
                'ETA: ' + (data.eta || ''),
                'Reason: ' + (data.reason || ''),
                'Sent by: ' + (data.sentBy || '')
              ];
              if (data.orderNo) noteLines.push('Order No: ' + data.orderNo);
              if (data.orderLink || data.amazonLink) noteLines.push(data.orderLink || data.amazonLink);
              _addNoteToDeal(kd.dealId, noteLines.join('\n'));
              noteSaved = true;
            } catch(ne) { Logger.log('[addKitEntry] Note error: ' + ne.message); }
          }
        }
      } catch(kitE) { Logger.log('[addKitEntry] WATI/note block error: ' + kitE.message); }
    }

    if (jlid && !watiSent) {
      hsWarning = (hsWarning ? hsWarning + ' ' : '') + 'WhatsApp message to parent was not sent.';
    }

    return { success: true, srNo: srNo, watiSent: watiSent, noteSaved: noteSaved, hsUpdated: hsUpdated, warning: hsWarning || null };

  } catch (e) {
    Logger.log('[KitTracking] addKitEntry ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// SEND FOLLOW-UP BY ROW INDEX  (manual send — handles missing JLID via auto-lookup)
// ─────────────────────────────────────────────────────────────────────────────
function sendKitFollowUpByRow(rowIndex, jlidOverride) {
  var timeline = [];
  try {
    var sheet = _getKitSheet();
    var row   = sheet.getRange(rowIndex, 1, 1, KIT_COL.PHONE_SENT_TO).getValues()[0];

    var learnerName = String(row[KIT_COL.LEARNER_NAME - 1] || '').trim();
    var kitName     = String(row[KIT_COL.KIT - 1]          || '').trim();
    var jlid        = jlidOverride || String(row[KIT_COL.JLID - 1] || '').trim();

    // Auto-find JLID if missing
    if (!jlid) {
      jlid = _findJlidByKitStatus(learnerName, kitName) || '';
      if (!jlid) return { success: false, needJlid: true, message: 'Could not auto-find JLID for "' + learnerName + '". Please enter JLID manually.', timeline: timeline };
      // Save it to sheet
      sheet.getRange(rowIndex, KIT_COL.JLID).setValue(jlid);
      Logger.log('[KitTracking] sendKitFollowUpByRow: auto-filled JLID=' + jlid + ' for row ' + rowIndex);
    }

    var hsStarted = new Date().getTime();
    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success || !hs.data.parentContact) {
      _timelineAdd(timeline, 'hubspot_lookup', 'HubSpot Parent Contact Failed', 'failed', hsStarted, (hs && hs.message) || '');
      return { success: false, message: 'HubSpot lookup failed for ' + jlid + ': ' + (hs && hs.message), timeline: timeline };
    }
    _timelineAdd(timeline, 'hubspot_lookup', 'HubSpot Parent Contact Found', 'success', hsStarted, '');

    var phone      = _normalisePhone(hs.data.parentContact);
    var parentName = hs.data.parentName || hs.data && hs.data.parentName || learnerName;
    if (!phone) return { success: false, message: 'No phone number found for ' + jlid, timeline: timeline };

    var whatsappStarted = new Date().getTime();
    var watiResult = sendWatiMessage(phone, 'migration_kit_fup_sent_by_us', [
      { name: 'Parent',   value: parentName  },
      { name: 'Kit_name', value: kitName     },
      { name: 'Learner',  value: learnerName }
    ]);
    Logger.log('[KitTracking] sendKitFollowUpByRow: WATI=' + JSON.stringify(watiResult));
    _timelineAdd(timeline, 'kit_whatsapp', 'Parent WhatsApp Sent', watiResult && watiResult.success === false ? 'failed' : 'success', whatsappStarted, phone);

    var sheetStarted = new Date().getTime();
    sheet.getRange(rowIndex, KIT_COL.FOLLOWUP_SENT).setValue('TRUE');
    sheet.getRange(rowIndex, KIT_COL.FOLLOWUP_SENT_AT).setValue(new Date());
    sheet.getRange(rowIndex, KIT_COL.PHONE_SENT_TO).setValue(phone);
    _timelineAdd(timeline, 'sheet_update', 'Kit Reminder Logged', 'success', sheetStarted, '');

    return { success: true, message: 'Follow-up sent to ' + phone, jlid: jlid, timeline: timeline };
  } catch (e) {
    Logger.log('[KitTracking] sendKitFollowUpByRow ERROR: ' + e.message);
    return { success: false, message: e.message, timeline: timeline };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// RESEND KIT FOLLOW-UP  (manual resend from dashboard)
// ─────────────────────────────────────────────────────────────────────────────
function resendKitFollowUp(jlid) {
  var timeline = [];
  try {
    if (!jlid) return { success: false, message: 'No JLID provided.', timeline: timeline };
    var sheet   = _getKitSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Sheet empty.', timeline: timeline };

    var rows = sheet.getRange(2, 1, lastRow - 1, KIT_COL.PHONE_SENT_TO).getValues();
    var matchRow = -1;
    rows.forEach(function(r, idx) {
      if (matchRow > -1) return;
      if (String(r[KIT_COL.JLID - 1] || '').trim() === jlid) matchRow = idx + 2;
    });

    if (matchRow === -1) return { success: false, message: 'JLID not found in sheet: ' + jlid, timeline: timeline };

    var dataRow     = rows[matchRow - 2];
    var learnerName = String(dataRow[KIT_COL.LEARNER_NAME - 1] || '').trim();
    var kitName     = String(dataRow[KIT_COL.KIT - 1]          || '').trim();

    var hsStarted = new Date().getTime();
    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success || !hs.data.parentContact) {
      _timelineAdd(timeline, 'hubspot_lookup', 'HubSpot Parent Contact Failed', 'failed', hsStarted, (hs && hs.message) || '');
      return { success: false, message: 'HubSpot lookup failed for ' + jlid, timeline: timeline };
    }
    _timelineAdd(timeline, 'hubspot_lookup', 'HubSpot Parent Contact Found', 'success', hsStarted, '');

    var phone      = _normalisePhone(hs.data.parentContact);
    var parentName = hs.data.parentName || learnerName || 'Parent';
    if (!phone) return { success: false, message: 'No phone found for ' + jlid, timeline: timeline };

    var whatsappStarted = new Date().getTime();
    var watiResult = sendWatiMessage(phone, 'migration_kit_fup_sent_by_us', [
      { name: 'Parent',   value: parentName  },
      { name: 'Kit_name', value: kitName     },
      { name: 'Learner',  value: learnerName }
    ]);
    Logger.log('[KitTracking] Resend result: ' + JSON.stringify(watiResult));
    _timelineAdd(timeline, 'kit_whatsapp', 'Parent WhatsApp Sent', watiResult && watiResult.success === false ? 'failed' : 'success', whatsappStarted, phone);

    // Update sheet
    var sheetStarted = new Date().getTime();
    sheet.getRange(matchRow, KIT_COL.FOLLOWUP_SENT).setValue('TRUE');
    sheet.getRange(matchRow, KIT_COL.FOLLOWUP_SENT_AT).setValue(new Date());
    sheet.getRange(matchRow, KIT_COL.PHONE_SENT_TO).setValue(phone);
    _timelineAdd(timeline, 'sheet_update', 'Kit Reminder Logged', 'success', sheetStarted, '');

    return { success: true, message: 'Follow-up resent to ' + phone, timeline: timeline };
  } catch (e) {
    Logger.log('[KitTracking] resendKitFollowUp ERROR: ' + e.message);
    return { success: false, message: e.message, timeline: timeline };
  }
}

// -----------------------------------------------------------------------------
// UPDATE KIT ROW  (manual edit from dashboard � pencil button)
// -----------------------------------------------------------------------------
function updateKitRow(data) {
  try {
    var rowIndex     = parseInt(data.rowIndex, 10);
    var jlid         = String(data.jlid         || '').trim();
    var deliveryDate = String(data.deliveryDate || '').trim();
    var response     = String(data.response     || '').trim();

    if (!rowIndex || rowIndex < 2) return { success: false, message: 'Invalid row index.' };

    var sheet = _getKitSheet();
    var row   = sheet.getRange(rowIndex, 1, 1, KIT_COL.PHONE_SENT_TO).getValues()[0];
    var kitName      = String(row[KIT_COL.KIT - 1]          || '').trim();
    var learnerName  = String(row[KIT_COL.LEARNER_NAME - 1] || '').trim();
    var existingJlid = String(row[KIT_COL.JLID - 1]         || '').trim();

    // Update JLID if provided and different
    if (jlid && jlid !== existingJlid) {
      sheet.getRange(rowIndex, KIT_COL.JLID).setValue(jlid);
      Logger.log('[KitTracking] updateKitRow: JLID=' + jlid + ' row=' + rowIndex);
    }
    var effectiveJlid = jlid || existingJlid;

    // Update delivery date + compute time taken
    if (deliveryDate) {
      sheet.getRange(rowIndex, KIT_COL.DELIVERY_DATE).setValue(deliveryDate);
      var orderRaw  = row[KIT_COL.DATE_OF_ORDER - 1];
      var orderDate = (orderRaw instanceof Date) ? orderRaw : _parseDMY(String(orderRaw || ''));
      var delivDate = _parseDMY(deliveryDate);
      if (orderDate && delivDate) {
        var days = Math.round((delivDate - orderDate) / 86400000);
        sheet.getRange(rowIndex, KIT_COL.TIME_TAKEN).setValue(days + ' days');
      }
      Logger.log('[KitTracking] updateKitRow: DeliveryDate=' + deliveryDate + ' row=' + rowIndex);
    }

    // Update parent response
    if (response) {
      sheet.getRange(rowIndex, KIT_COL.PARENT_RESPONSE).setValue(response);
      Logger.log('[KitTracking] updateKitRow: Response=�' + response + '� row=' + rowIndex);
    }

    var hsStatus = 'skipped';

    // Kit Received ? update HubSpot kit status
    if (response === 'Kit Received') {
      if (!effectiveJlid) {
        hsStatus = 'no_jlid';
        Logger.log('[KitTracking] updateKitRow: Kit Received but no JLID � HubSpot NOT updated.');
      } else {
        var prop = _kitPropertyForType(kitName);
        if (!prop) {
          hsStatus = 'unknown_kit';
          Logger.log('[KitTracking] updateKitRow: unknown kit type "' + kitName + '" � cannot map to HubSpot property.');
        } else {
          try {
            var hsLookup = fetchHubspotByJlid(effectiveJlid);
            if (!hsLookup || !hsLookup.success || !hsLookup.data.dealId) {
              hsStatus = 'no_deal';
              Logger.log('[KitTracking] updateKitRow: no dealId for ' + effectiveJlid);
            } else {
              var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
              var url   = 'https://api.hubapi.com/crm/v3/objects/deals/' + hsLookup.data.dealId;
              var patchBody = {};
              patchBody[prop] = 'Received'; // internal enum value (label: "Received by the Parents")
              var resp = monitoredFetch(url, {
                method: 'PATCH',
                headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
                payload: JSON.stringify({ properties: patchBody }),
                muteHttpExceptions: true
              });
              var code = resp.getResponseCode();
              var fullBody = resp.getContentText();
              Logger.log('[KitTracking] updateKitRow: SENDING prop="' + prop + '" value="Received" dealId=' + hsLookup.data.dealId);
              Logger.log('[KitTracking] updateKitRow: HubSpot PATCH HTTP ' + code + ' FULL_BODY=' + fullBody);
              hsStatus = (code === 200) ? 'updated' : ('http_' + code + ': ' + fullBody.substring(0, 200));
            }
          } catch (hsErr) {
            hsStatus = 'error: ' + hsErr.message;
            Logger.log('[KitTracking] updateKitRow: HubSpot error: ' + hsErr.message);
          }
        }
      }
    }

    // Not Received / Need To Check ? add HubSpot note
    if ((response === 'Not Received yet' || response === 'Need To check') && effectiveJlid) {
      try {
        var hsN = fetchHubspotByJlid(effectiveJlid);
        if (hsN && hsN.success && hsN.data.dealId) {
          _addNoteToDeal(hsN.data.dealId, '[Kit Tracking] Manual update: ' + response + ' for ' + kitName + ' � ' + learnerName);
        }
      } catch (noteErr) {
        Logger.log('[KitTracking] updateKitRow: note failed: ' + noteErr.message);
      }
    }

    // Build user-facing message
    var msg = 'Sheet updated.';
    if (response === 'Kit Received') {
      if (hsStatus === 'updated')       msg = 'Sheet updated ?  HubSpot ' + prop + ' ? Received by the Parents ?';
      else if (hsStatus === 'no_jlid')  msg = 'Sheet updated ?  HubSpot SKIPPED � JLID missing. Add JLID and save again.';
      else if (hsStatus === 'no_deal')  msg = 'Sheet updated ?  HubSpot SKIPPED � deal not found for ' + effectiveJlid;
      else if (hsStatus === 'unknown_kit') msg = 'Sheet updated ?  HubSpot SKIPPED � kit type "' + kitName + '" not mapped.';
      else                              msg = 'Sheet updated ?  HubSpot failed: ' + hsStatus;
    }

    return { success: true, message: msg, hsStatus: hsStatus };
  } catch (e) {
    Logger.log('[KitTracking] updateKitRow ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// REGISTER DAILY TRIGGER  (call once from initializeSystem or manually)
function setupKitTrackingTrigger() {
  // Remove any existing duplicate triggers first
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendKitFollowUps') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('sendKitFollowUps')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  Logger.log('[KitTracking] Daily trigger registered for sendKitFollowUps at 8am.');
}

function testKaiConfirmMsg() {
  var result = sendWatiMessage('447711736472', 'kit_received_confirmation', [
    { name: 'kit',     value: 'VR Headset'   },
    { name: 'learner', value: 'Kai Thobani' }
  ]);
  Logger.log('Result: ' + JSON.stringify(result));
}


// ── addPWBEntry — called from UI "Add Entry" in Parent Will Buy mode ───────
function addPWBEntry(data) {
  try {
    var sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID_KIT_TRACKING')
                  || '17Jsa2Kl2AkI5SgtlITYGqb-Q-PxfzkNFzzG5HwBJp_Q';
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Parent_will_buy');
    if (!sheet) return { success: false, message: 'Parent_will_buy sheet not found' };

    // Fetch HubSpot data for names/phone
    var learnerName = '';
    var parentName  = '';
    var phone       = '';
    try {
      var hs = fetchHubspotByJlid((data.jlid || '').trim());
      if (hs && hs.success && hs.data) {
        learnerName = hs.data.learnerName || '';
        parentName  = hs.data.parentName  || '';
        phone       = _normalisePhone(hs.data.parentContact || hs.data.phone || '');
      }
    } catch(e) { Logger.log('[addPWBEntry] HS lookup failed: ' + e.message); }

    var kit = _getKitForCourse(data.courseName || '') || '';
    var lastRow = sheet.getLastRow();
    var srNo = lastRow < 2 ? 1 : (sheet.getRange(lastRow, 1).getValue() || 0) + 1;

    var now        = new Date();
    var entryDate  = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    var entryMonth = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM yyyy');

    sheet.appendRow([
      srNo,                      // A: Sr No
      entryDate,                 // B: Date
      entryMonth,                // C: Month
      (data.jlid || '').trim(),  // D: JLID
      learnerName,               // E: Learner Name
      parentName,                // F: Parent Name
      phone,                     // G: Parent Phone
      data.courseName || '',     // H: Course Name
      kit,                       // I: Kit
      data.courseStartDate || '',// J: Course Start Date
      data.orderLink || data.amazonLink || '', // K: Order Link
      'Pending',                 // L: Status
      '', '', '', '', '', '', '', '', // M-T: timestamps / response / escalation / interval
      data.entryBy || ''         // U: Entry By
    ]);

    // Fire initial message immediately — don't wait for 9am trigger
    try {
      var newRow    = sheet.getLastRow();
      var rowData   = sheet.getRange(newRow, 1, 1, 21).getValues()[0]; // 21 cols = A–U
      var today     = new Date(); today.setHours(0,0,0,0);
      _processPWBRow(sheet, newRow, rowData, today);
      Logger.log('[addPWBEntry] Immediate processing triggered for row ' + newRow);
    } catch(triggerErr) {
      Logger.log('[addPWBEntry] Immediate trigger failed: ' + triggerErr.message);
      // Row is saved — 9am trigger will catch it as fallback
    }

    return { success: true };
  } catch(e) {
    Logger.log('[addPWBEntry] ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Next FUP label helper ──────────────────────────────────────────────────
function _pwbNextFupLabel(r) {
  // r = raw row array (0-indexed, 21 cols A-U)
  // PWB_COL is 1-based so subtract 1
  var interval      = Number(r[19]) || 0; // T = col 20 → index 19
  var initialSentAt = r[12]; // M = col 13 → index 12
  var fup1SentAt    = r[13]; // N
  var fup2SentAt    = r[14]; // O
  var finalSentAt   = r[15]; // P
  var status        = String(r[11] || ''); // L = col 12 → index 11

  var TERMINAL = ['Order Placed', 'Kit Received', "Parent Didn't Buy - Roadmap Changed",
                  '⚠️ CLS Notified - Awaiting Response'];
  if (TERMINAL.some(function(t){ return status.indexOf(t) > -1; })) return '—';
  if (!initialSentAt) return 'Sends tonight';

  var today = new Date(); today.setHours(0,0,0,0);

  function label(sentAt, step) {
    if (!interval) return 'Unknown interval';
    var d = new Date(sentAt); d.setHours(0,0,0,0);
    var due = new Date(d); due.setDate(due.getDate() + interval);
    var diff = Math.round((due - today) / 86400000);
    var stepLabel = step + ' in ';
    if (diff < 0)  return '🔴 ' + step + ' overdue by ' + Math.abs(diff) + 'd';
    if (diff === 0) return '🟡 ' + step + ' today';
    if (diff === 1) return '🟡 ' + step + ' tomorrow';
    return '🟢 ' + step + ' in ' + diff + 'd';
  }

  if (!fup1SentAt)  return label(initialSentAt, 'FUP 1');
  if (!fup2SentAt)  return label(fup1SentAt,    'FUP 2');
  if (!finalSentAt) return label(fup2SentAt,    'Final FUP');
  return '⏳ Awaiting reply';
}

// ── getPWBEntries — returns all rows for UI table ──────────────────────────
function getPWBEntries() {
  try {
    var sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID_KIT_TRACKING')
                  || '17Jsa2Kl2AkI5SgtlITYGqb-Q-PxfzkNFzzG5HwBJp_Q';
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Parent_will_buy');
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var rows = sheet.getRange(2, 1, lastRow - 1, 25).getValues(); // A–Y
    var tz = Session.getScriptTimeZone();
    function fmtDate(v) { return v instanceof Date ? Utilities.formatDate(v, tz, 'dd/MM/yyyy') : String(v || ''); }
    function fmtTs(v)   { return v instanceof Date ? Utilities.formatDate(v, tz, 'dd/MM/yyyy HH:mm') : String(v || ''); }
    function fmtMonth(v) {
      if (v instanceof Date) return Utilities.formatDate(v, tz, 'MMMM yyyy');
      var s = String(v || '').trim();
      // If it looks like a date string (contains timezone/GMT info), parse and reformat
      if (s && (s.indexOf('GMT') > -1 || s.indexOf(':') > -1)) {
        try { return Utilities.formatDate(new Date(s), tz, 'MMMM yyyy'); } catch(e) {}
      }
      return s;
    }

    return rows.map(function(r) {
      return {
        srNo:           r[0],
        entryDate:      fmtDate(r[1]),
        entryMonth:     fmtMonth(r[2]),
        jlid:           String(r[3]  || ''),
        learnerName:    String(r[4]  || ''),
        parentName:     String(r[5]  || ''),
        parentPhone:    String(r[6]  || ''),
        courseName:     String(r[7]  || ''),
        kit:            String(r[8]  || ''),
        courseStartDate:fmtDate(r[9]),
        amazonLink:     String(r[10] || ''),
        status:         String(r[11] || 'Pending'),
        initialSentAt:  fmtTs(r[12]),
        fup1SentAt:     fmtTs(r[13]),
        fup2SentAt:     fmtTs(r[14]),
        finalFupSentAt: fmtTs(r[15]),
        parentResponse: String(r[16] || ''),
        escalated:      String(r[17] || ''),
        escalatedAt:    fmtTs(r[18]),
        interval:       Number(r[19]) || 0,
        entryBy:        String(r[20] || ''),
        emailLog:       String(r[21] || ''),
        promisedDate:   String(r[22] || ''),
        aiIntent:       String(r[23] || ''),
        responseAt:     fmtTs(r[24]),
        nextFup:        _pwbNextFupLabel(r)
      };
    }).filter(function(r) { return r.jlid; });
  } catch(e) {
    Logger.log('[getPWBEntries] ERROR: ' + e.message);
    return [];
  }
}

// ── Mark kit as refunded (logistics loss) — zeros price, sets refunded flag, subtracts from HubSpot ──
function markKitAsRefunded(rowIndex) {
  if (!rowIndex) return { success: false, message: 'No rowIndex' };
  try {
    var sheet = _getKitSheet();
    var row   = sheet.getRange(rowIndex, 1, 1, KIT_COL.REFUNDED).getValues()[0];

    var price = parseFloat(row[4]) || 0;    // col E (index 4) = price
    var jlid  = String(row[KIT_COL.JLID - 1] || '').trim(); // col P (index 15)

    // Mark refunded + zero sheet price
    sheet.getRange(rowIndex, KIT_COL.REFUNDED).setValue('TRUE');
    sheet.getRange(rowIndex, 5).setValue(0);
    Logger.log('[KitTracking] markKitAsRefunded row=' + rowIndex + ' price=' + price + ' jlid=' + jlid);

    // Subtract from HubSpot learning_kit_cost
    if (price > 0 && jlid) {
      try {
        var token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY') || '';
        if (token) {
          var hs = fetchHubspotByJlid(jlid);
          if (hs && hs.success && hs.data && hs.data.dealId) {
            var current = _getHubSpotDealKitCost(hs.data.dealId, token);
            var updated = Math.max(0, current - price);
            _patchHubSpotKitCost(hs.data.dealId, token, updated);
            Logger.log('[KitTracking] markKitAsRefunded HS: ' + current + ' - ' + price + ' = ' + updated);
          }
        }
      } catch(hsErr) { Logger.log('[KitTracking] markKitAsRefunded HS error: ' + hsErr.message); }
    }

    return { success: true, price: price };
  } catch(e) {
    Logger.log('[KitTracking] markKitAsRefunded ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── Update PWB row status manually (Order Placed / Kit Received / Course Changed) ──
function updatePWBStatus(jlid, newStatus) {
  if (!jlid || !newStatus) return { success: false, message: 'Missing jlid or status' };
  try {
    var sheet   = _getPWBSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'PWB sheet empty' };
    var data    = sheet.getRange(2, 1, lastRow - 1, PWB_COL.EMAIL_LOG).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][PWB_COL.JLID - 1] || '').trim() !== jlid.trim()) continue;
      var sheetRow = i + 2;
      sheet.getRange(sheetRow, PWB_COL.STATUS).setValue(newStatus);
      // Update HubSpot deal status to match
      try {
        var kitName = String(data[i][PWB_COL.KIT - 1] || '').trim();
        var hs = fetchHubspotByJlid(jlid);
        if (hs && hs.success && hs.data && hs.data.dealId && kitName) {
          var hsVal = (newStatus === 'Order Placed' || newStatus === 'Kit Received')
            ? PWB_HS_STATUSES.BOUGHT : PWB_HS_STATUSES.ROADMAP;
          _updateHubspotPWBStatus(hs.data.dealId, hsVal, kitName);
        }
      } catch(he) { Logger.log('[updatePWBStatus] HS: ' + he.message); }
      Logger.log('[updatePWBStatus] ' + jlid + ' → ' + newStatus);
      return { success: true };
    }
    return { success: false, message: 'JLID not found: ' + jlid };
  } catch(e) {
    Logger.log('[updatePWBStatus] ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// KIT ORDER LOGGING  (called from KitOrderService after Amazon order placed)
// Appends a new row to the Kit Tracking sheet using the correct KIT_COL map.
// ─────────────────────────────────────────────────────────────────────────────

/**
 * @param {object} data  { jlid, kitName, orderDate, deliveryDate, price, orderNo, amazonLink }
 * @param {string} learnerName
 * @param {string} country  (from HubSpot contact)
 */
function logKitOrderToSheet(data, learnerName, country) {
  try {
    var sheet        = _getKitSheet();
    var sheetLastRow = sheet.getLastRow();

    // Find actual last data row — scan col A from bottom, skip empty/formatted rows
    var lastDataRow = 1;
    if (sheetLastRow >= 2) {
      var colA = sheet.getRange(2, 1, sheetLastRow - 1, 1).getValues();
      for (var i = colA.length - 1; i >= 0; i--) {
        if (String(colA[i][0]).trim() !== '') { lastDataRow = i + 2; break; }
      }
    }

    // Auto Sr No = highest existing Sr No + 1
    var lastSrNo = 0;
    if (sheetLastRow >= 2) {
      var srValues = sheet.getRange(2, KIT_COL.SR_NO, sheetLastRow - 1, 1).getValues();
      srValues.forEach(function(r) {
        var v = parseInt(r[0], 10);
        if (!isNaN(v) && v > lastSrNo) lastSrNo = v;
      });
    }
    var srNo = lastSrNo + 1;

    // Write per-column to skip col H (formula in sheet — never overwrite)
    var writeRow = lastDataRow + 1;
    var writeMap = [
      [1,  srNo],                           // A: Sr No
      [2,  learnerName || ''],              // B: Learner's name
      [3,  data.kitName    || ''],          // C: Kit
      [4,  country         || ''],          // D: Country
      [5,  data.price      || ''],          // E: Price (EUR)
      [6,  ''],                             // F: Site
      [7,  data.orderDate  || ''],          // G: Date of Order
      // col 8 (H) = SKIP — formula in sheet
      [9,  data.deliveryDate || ''],        // I: ETA
      [10, ''],                             // J: Delivery Date
      [12, ''],                             // L: Reason
      [13, ''],                             // M: Current Subscription
      [14, ''],                             // N: Roadmap
      [15, ''],                             // O: Name (Sent By)
      [16, data.jlid       || '']          // P: JLID
    ];
    writeMap.forEach(function(pair) {
      sheet.getRange(writeRow, pair[0]).setValue(pair[1]);
    });

    Logger.log('[logKitOrderToSheet] Logged kit order for ' + (data.jlid || learnerName) +
               ' row=' + writeRow);
    return true;
  } catch(e) {
    Logger.log('[logKitOrderToSheet] ERROR: ' + e.message);
    return false;
  }
}

