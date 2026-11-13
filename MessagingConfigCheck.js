// ═══════════════════════════════════════════════════════════════════════════
// MESSAGING CONFIG CHECK
// "Is everything the app sends over WhatsApp set up in WATI?" Lists the templates the code uses and says, for each, whether it exists in
// WATI and is APPROVED (a template that is missing, pending or rejected silently fails to reach the parent).
// Run checkMessagingConfig() from the Apps Script editor; the result is logged and returned. Sends nothing.
// ═══════════════════════════════════════════════════════════════════════════
var MSG_REQUIRED_TEMPLATES = [
  { name: 'kit_address_request_link_v2',   used: 'Kit: ask the parent for the delivery address (link)' },
  { name: 'kit_address_reconfirm_v2',      used: 'Kit: re-confirm an address already on file' },
  { name: 'kit_address_received_confirmation', used: 'Kit: address received' },
  { name: 'kit_order_placed_notice_v2',    used: 'Kit: order placed notice (address + expected date)' },
  { name: 'kit_tracking_link_v1',          used: 'Kit: tracking link message' },
  { name: 'kit_tracking_link_btn_v1',      used: 'Kit: tracking link button' },
  { name: 'kit_received_confirmation',     used: 'Kit: received confirmation' },
  { name: 'kit_not_received_reply',        used: 'Kit: parent says not received' },
  { name: 'kit_need_to_check_reply',       used: 'Kit: need to check reply' },
  { name: 'migration_address_template',    used: 'Kit: fallback address request' },
  { name: 'migration_kit_sent_by_us_parent_information', used: 'Kit: parent information (sent by us)' },
  { name: 'migration_kit_fup_sent_by_us',  used: 'Kit: follow-up (sent by us)' },
  { name: 'outbound_msg',                  used: 'Kit logistics messages to parents (delivered, delay, refund...)' },
  { name: 'migration_amz_giftvoucher_',    used: 'Gift card: ask for Amazon details' },
  { name: 'amz_giftcard_sent_',            used: 'Gift card: issued (referral)' },
  { name: 'amz_giftcard_sales_pitch',      used: 'Gift card: issued (sales)' },
  { name: 'onboarding_link',               used: 'Onboarding: Jet, Set, Go!' },
  { name: 'onboarding_reminder',           used: 'Onboarding: reminder' }
];

function _mccFetchTemplates() {
  var props = PropertiesService.getScriptProperties();
  var base = (props.getProperty('WATI_API_ENDPOINT') || '').trim().replace(/\/$/, '');
  var token = (props.getProperty('WATI_ACCESS_TOKEN') || '').trim();
  if (!base || !token) throw new Error('WATI_API_ENDPOINT / WATI_ACCESS_TOKEN are not set in Script Properties.');
  if (!/^Bearer /i.test(token)) token = 'Bearer ' + token;
  var all = [];
  for (var page = 1; page <= 6; page++) {
    var resp = monitoredFetch(base + '/api/v1/getMessageTemplates?pageSize=100&pageNumber=' + page, { method: 'get', headers: { 'Authorization': token }, muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) throw new Error('WATI answered HTTP ' + resp.getResponseCode() + ' when listing templates.');
    var j = JSON.parse(resp.getContentText());
    var items = j.messageTemplates || j.templates || [];
    all = all.concat(items);
    if (items.length < 100) break;
  }
  return all;
}

function checkMessagingConfig() {
  try {
    var templates = _mccFetchTemplates();
    var byName = {};
    templates.forEach(function(t) { byName[String(t.elementName || t.name || '').trim()] = t; });
    var rows = MSG_REQUIRED_TEMPLATES.map(function(r) {
      var t = byName[r.name];
      var status = t ? String(t.status || '').toUpperCase() : '';
      var state = !t ? 'MISSING' : (status === 'APPROVED' ? 'OK' : (status || 'UNKNOWN'));
      return { template: r.name, state: state, usedFor: r.used };
    });
    var bad = rows.filter(function(r) { return r.state !== 'OK'; });
    Logger.log('[MessagingConfig] ' + templates.length + ' templates in WATI. ' + (bad.length ? bad.length + ' need attention:' : 'All required templates are approved.'));
    rows.forEach(function(r) { Logger.log((r.state === 'OK' ? 'OK       ' : r.state.padEnd(9)) + r.template + '   — ' + r.usedFor); });
    return { success: true, total: templates.length, problems: bad.length, rows: rows };
  } catch (e) {
    Logger.log('[MessagingConfig] failed: ' + e.message);
    return { success: false, message: e.message };
  }
}
