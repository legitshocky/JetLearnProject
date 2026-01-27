function debugLineItemDetails() {
  const TEST_JLID = "JL39464432157C"; 
  const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');

  console.log(`🔹 1. Finding Deal for: ${TEST_JLID}`);
  
  // A. Search for Deal
  const searchUrl = 'https://api.hubapi.com/crm/v3/objects/deals/search';
  const searchRes = UrlFetchApp.fetch(searchUrl, {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify({
      filterGroups: [{ filters: [{ propertyName: 'jetlearner_id', operator: 'EQ', value: TEST_JLID }] }],
      sorts: [{ propertyName: 'createdate', direction: 'DESCENDING' }],
      limit: 5
    }),
    muteHttpExceptions: true
  });
  
  const searchJson = JSON.parse(searchRes.getContentText());
  if (!searchJson.results || searchJson.results.length === 0) { console.error("No deal found."); return; }

  // B. Find Deal with Line Items
  let targetDeal = null;
  let itemIds = [];
  
  for (const deal of searchJson.results) {
    const assocRes = UrlFetchApp.fetch(`https://api.hubapi.com/crm/v4/objects/deals/${deal.id}/associations/line_items`, {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const assocJson = JSON.parse(assocRes.getContentText());
    if (assocJson.results && assocJson.results.length > 0) {
      targetDeal = deal;
      itemIds = assocJson.results.map(r => ({ id: r.toObjectId }));
      break;
    }
  }

  if (!targetDeal) { console.error("No line items attached to any deal."); return; }
  console.log(`   ✅ Target Deal ID: ${targetDeal.id} | Line Items: ${itemIds.length}`);

  // C. Fetch Line Item Details (Requesting properties one by one to be safe)
  console.log(`🔹 2. Fetching Details...`);
  
  const batchUrl = `https://api.hubapi.com/crm/v3/objects/line_items/batch/read`;
  const batchPayload = {
    properties: [
      'name', 'price', 'hs_createdate',
      // Check these specific names from your screenshot
      'payment_received_date___cloned_', 
      'renewal__payment_type__cloned_',
      'renewal__payment_term__cloned_',
      'full_payment_received__y_n___cloned_',
      'hs_recurring_billing_number_of_payments'
    ],
    inputs: itemIds
  };

  const batchRes = UrlFetchApp.fetch(batchUrl, {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify(batchPayload),
    muteHttpExceptions: true
  });

  const responseText = batchRes.getContentText();
  const batchJson = JSON.parse(responseText);

  // --- ERROR TRAPPING ---
  if (batchRes.getResponseCode() !== 200 || !batchJson.results) {
      console.error("❌ HUBSPOT API ERROR:");
      console.error(responseText); // This will print the specific reason
      return;
  }
  // ---------------------
  
  // D. Sort and Display
  const sortedItems = batchJson.results.sort((a, b) => {
    return new Date(a.properties.hs_createdate) - new Date(b.properties.hs_createdate);
  });

  console.log("\n📊 LINE ITEM ANALYSIS (Sorted Oldest to Newest):");

  sortedItems.forEach((item, index) => {
    const p = item.properties;
    let type = (index === 0) ? "🟢 ENROLLMENT" : "⭐ RENEWAL";
    
    // Check specific fields
    const dateVal = p.payment_received_date___cloned_;
    
    console.log(`\n${type}`);
    console.log(`   ID: ${item.id}`);
    console.log(`   Name: ${p.name}`);
    console.log(`   Created: ${new Date(p.hs_createdate).toISOString()}`);
    console.log(`   💰 Payment Date (raw): ${dateVal}`);
    
    if (dateVal && !isNaN(dateVal)) {
       console.log(`   -> Formatted: ${new Date(parseInt(dateVal)).toISOString().split('T')[0]}`);
    }
  });
  
  console.log("\n🔹 END DEBUG");
}