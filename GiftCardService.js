// ─────────────────────────────────────────────────────────────────────────────
// GIFT CARD / REFERRAL REWARD SERVICE
// Reads/writes the Referral Reward & Gift Card sheet.
// Updates HubSpot deal properties and adds notes.
// Sends WATI templates for ask + issued flows.
// ─────────────────────────────────────────────────────────────────────────────

var GC_SHEET_ID  = '17Jsa2Kl2AkI5SgtlITYGqb-Q-PxfzkNFzzG5HwBJp_Q';
var GC_SHEET_GID = 565604841;

// Column indices (1-based) — match actual sheet headers
var GC_COL = {
  LEARNER_NAME: 1,   // A
  EMAIL:        2,   // B
  AMOUNT:       3,   // C  Gift card amount (in sent currency)
  DATE:         4,   // D
  MONTHS:       5,   // E  formula in sheet — never write
  CURRENCY:     6,   // F
  EUR_AMOUNT:   7,   // G  Converted EUR amount
  WEBSITE:      8,   // H
  REASON:       9,   // I  Referral / Sales / other
  TYPE:         10,  // J  "Source" col — referral | sales
  JLID:         11,  // K  "Referred to" col
  ORDER_LINK:   12,  // L  "Received By HS link" col — Amazon order link
  ASK_SENT_AT:  13,  // M
  ISSUED_SENT_AT: 14,// N
  HS_UPDATED:   15,  // O
  ADDED_BY:     16   // P
};
var GC_LAST_COL = 16;

function _getGiftCardSheet() {
  var ss = SpreadsheetApp.openById(GC_SHEET_ID);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === GC_SHEET_GID) return sheets[i];
  }
  // Fallback: sheet named "Referral Reward & Gift Card" or first match
  throw new Error('Gift card sheet not found (gid=' + GC_SHEET_GID + ')');
}

// ── READ ──────────────────────────────────────────────────────────────────────
function getGiftCardData() {
  try {
    var sheet   = _getGiftCardSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, rows: [], stats: {} };

    var numCols = Math.max(sheet.getLastColumn(), GC_LAST_COL);
    var raw = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    var rows = [];

    raw.forEach(function(r, idx) {
      var learnerName = String(r[GC_COL.LEARNER_NAME - 1] || '').trim();
      var email       = String(r[GC_COL.EMAIL        - 1] || '').trim();
      if (!learnerName && !email) return;

      var dateRaw = r[GC_COL.DATE - 1];
      var dateStr = '';
      if (dateRaw instanceof Date && !isNaN(dateRaw.getTime())) {
        dateStr = _gcFormatDMY(dateRaw);
      } else if (dateRaw) {
        dateStr = String(dateRaw).trim();
      }

      rows.push({
        rowIndex:     idx + 2,
        learnerName:  learnerName,
        email:        email,
        amount:       String(r[GC_COL.AMOUNT    - 1] || '').trim(),
        date:         dateStr,
        months:       String(r[GC_COL.MONTHS    - 1] || '').trim(),
        currency:     String(r[GC_COL.CURRENCY  - 1] || '').trim(),
        eurAmount:    String(r[GC_COL.EUR_AMOUNT- 1] || '').trim(),
        website:      String(r[GC_COL.WEBSITE   - 1] || '').trim(),
        reason:       String(r[GC_COL.REASON    - 1] || '').trim(),
        jlid:         String(r[GC_COL.JLID      - 1] || '').trim(),
        orderLink:    String(r[GC_COL.ORDER_LINK- 1] || '').trim(),
        type:         String(r[GC_COL.TYPE      - 1] || '').trim(),
        askSentAt:    _gcFormatDateTime(r[GC_COL.ASK_SENT_AT   - 1]),
        issuedSentAt: _gcFormatDateTime(r[GC_COL.ISSUED_SENT_AT- 1]),
        hsUpdated:    String(r[GC_COL.HS_UPDATED- 1] || '').trim(),
        addedBy:      String(r[GC_COL.ADDED_BY  - 1] || '').trim()
      });
    });

    var referralCount = rows.filter(function(r) { return r.type === 'referral'; }).length;
    var salesCount    = rows.filter(function(r) { return r.type === 'sales'; }).length;
    var totalEur      = rows.reduce(function(s, r) { return s + (parseFloat(r.eurAmount) || 0); }, 0);

    return {
      success: true,
      rows:    rows,
      stats:   { total: rows.length, referral: referralCount, sales: salesCount, totalEur: totalEur.toFixed(2) }
    };
  } catch(e) {
    Logger.log('[GiftCard] getGiftCardData ERROR: ' + e.message);
    return { success: false, message: e.message, rows: [], stats: {} };
  }
}

// ── ADD ENTRY ─────────────────────────────────────────────────────────────────
// data: { jlid, learnerName, email, amount, date (YYYY-MM-DD), currency, eurAmount,
//         website, reason, type ('referral'|'sales'), orderLink, addedBy, sendWati }
function addGiftCardEntry(data) {
  try {
    var sheet    = _getGiftCardSheet();
    // Use col A (learner name) to find last real data row — avoids formula-only rows in col E
    var colAVals = sheet.getRange(2, GC_COL.LEARNER_NAME, sheet.getLastRow(), 1).getValues();
    var lastDataRow = 1;
    for (var i = colAVals.length - 1; i >= 0; i--) {
      if (String(colAVals[i][0] || '').trim()) { lastDataRow = i + 2; break; }
    }
    var writeRow = lastDataRow + 1;

    var jlid       = String(data.jlid       || '').trim().toUpperCase();
    var eurAmount  = parseFloat(data.eurAmount) || 0;
    var dateVal    = data.date ? new Date(data.date) : new Date();
    var gcType     = String(data.type || 'referral').toLowerCase();

    // Write sheet row (skip col E = MONTHS which has a formula)
    var writeMap = [
      [GC_COL.LEARNER_NAME, data.learnerName || ''],
      [GC_COL.EMAIL,        data.email        || ''],
      [GC_COL.AMOUNT,       data.amount       || ''],
      [GC_COL.DATE,         dateVal],
      // col 5 = MONTHS — skip (formula)
      [GC_COL.CURRENCY,     data.currency     || ''],
      [GC_COL.EUR_AMOUNT,   eurAmount || ''],
      [GC_COL.WEBSITE,      data.website      || ''],
      [GC_COL.REASON,       data.reason       || ''],
      [GC_COL.JLID,         jlid],
      [GC_COL.ORDER_LINK,   data.orderLink    || ''],
      [GC_COL.TYPE,         gcType],
      [GC_COL.ADDED_BY,     data.addedBy      || '']
    ];
    writeMap.forEach(function(pair) {
      sheet.getRange(writeRow, pair[0]).setValue(pair[1]);
    });
    Logger.log('[GiftCard] addGiftCardEntry: row ' + writeRow + ' jlid=' + jlid);

    var hsWarning = '', watiResult = null;

    // ── HubSpot updates ───────────────────────────────────────────────────────
    if (jlid) {
      try {
        var hs = fetchHubspotByJlid(jlid);
        if (hs && hs.success && hs.data && hs.data.dealId) {
          var token  = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
          var dealId = hs.data.dealId;

          // Fetch current deal properties
          var dealResp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/deals/' + dealId +
            '?properties=amazon_gift_vouchers_issued__month_,amazon_card_voucher_amount__in_euros_,number_of_referral_conversion', {
            method: 'get',
            headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
            muteHttpExceptions: true
          });

          var props = {};
          if (dealResp.getResponseCode() === 200) {
            props = JSON.parse(dealResp.getContentText()).properties || {};
          }

          // 1. Append month (e.g. "September 2026")
          var MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
          var newMonth = MONTHS[dateVal.getMonth()] + ' ' + dateVal.getFullYear();
          var existingMonths = String(props['amazon_gift_vouchers_issued__month_'] || '').trim();
          var updatedMonths  = existingMonths ? existingMonths + ', ' + newMonth : newMonth;

          // 2. Add EUR amount
          var existingEur  = parseFloat(props['amazon_card_voucher_amount__in_euros_']) || 0;
          var updatedEur   = existingEur + eurAmount;

          // 3. Increment referral count (only for referral type)
          var existingRef  = parseInt(props['number_of_referral_conversion']) || 0;
          var updatedRef   = gcType === 'referral' ? existingRef + 1 : existingRef;

          var patchProps = {
            'amazon_gift_vouchers_issued__month_':      updatedMonths,
            'amazon_card_voucher_amount__in_euros_':    String(updatedEur)
          };
          if (gcType === 'referral') {
            patchProps['number_of_referral_conversion'] = String(updatedRef);
          }

          var patchResp = monitoredFetch('https://api.hubapi.com/crm/v3/objects/deals/' + dealId, {
            method: 'PATCH',
            headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
            payload: JSON.stringify({ properties: patchProps }),
            muteHttpExceptions: true
          });
          var patchCode = patchResp.getResponseCode();
          Logger.log('[GiftCard] HS PATCH HTTP ' + patchCode + ' props=' + JSON.stringify(patchProps));

          if (patchCode === 200) {
            sheet.getRange(writeRow, GC_COL.HS_UPDATED).setValue('YES');
          } else {
            hsWarning = 'HS update failed HTTP ' + patchCode;
            Logger.log('[GiftCard] HS PATCH ERROR: ' + patchResp.getContentText().substring(0, 300));
          }

          // HS Note
          var noteLines = [
            '🎁 Gift Card Issued',
            'Learner: ' + (data.learnerName || '—'),
            'Type: ' + (gcType === 'referral' ? 'Referral Reward' : 'Sales Pitch'),
            'Amount: ' + (data.amount || '—') + ' ' + (data.currency || '') + ' (€' + eurAmount + ')',
            'Website: ' + (data.website || '—'),
            'Date: ' + _gcFormatDMY(dateVal),
            (data.orderLink ? 'Order Link: ' + data.orderLink : ''),
            'Added By: ' + (data.addedBy || '—')
          ].filter(Boolean);
          addNoteToHubSpotDeal(dealId, noteLines.join('\n'));

          // ── WATI — send issued message ────────────────────────────────────
          if (data.sendWati) {
            try {
              var parentPhone  = String(hs.data.parentContact || '').replace(/\D/g, '');
              var parentName   = String(hs.data.parentName    || 'Parent').trim().split(' ')[0];
              var templateName = gcType === 'referral' ? 'amz_giftcard_sent_' : 'amz_giftcard_sales_pitch';
              var watiParams   = [
                { name: 'Parent', value: parentName },
                { name: 'amount', value: (data.amount || '') + ' ' + (data.currency || '') }
              ];
              var wr = sendWatiMessage(parentPhone, templateName, watiParams);
              watiResult = wr;
              if (wr && wr.result) {
                sheet.getRange(writeRow, GC_COL.ISSUED_SENT_AT).setValue(new Date());
              }
            } catch(we) {
              Logger.log('[GiftCard] WATI issued send error: ' + we.message);
              watiResult = { success: false, message: we.message };
            }
          }
        } else {
          hsWarning = 'HubSpot deal not found for ' + jlid;
        }
      } catch(he) {
        hsWarning = 'HS error: ' + he.message;
        Logger.log('[GiftCard] HS update error: ' + he.message);
      }
    }

    // ── Email — issued (after the row exists so the template reads the saved values) ──
    var emailResult = null;
    if (data.sendEmail && jlid) {
      emailResult = gcSendIssuedEmail(writeRow, jlid, gcType);
    }

    return { success: true, rowIndex: writeRow, hsWarning: hsWarning, watiResult: watiResult, emailResult: emailResult };
  } catch(e) {
    Logger.log('[GiftCard] addGiftCardEntry ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── SEND ASK MESSAGE ──────────────────────────────────────────────────────────
// Sends migration_amz_giftvoucher_ to the parent to ask about gift card preference
function gcSendAskMessage(jlid) {
  try {
    if (!jlid) return { success: false, message: 'JLID required' };
    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success || !hs.data) return { success: false, message: 'Deal not found for ' + jlid };

    var parentPhone = String(hs.data.parentContact || '').replace(/\D/g, '');
    var parentName  = String(hs.data.parentName || 'Parent').trim().split(' ')[0];
    if (!parentPhone) return { success: false, message: 'No parent phone for ' + jlid };

    var wr = sendWatiMessage(parentPhone, 'migration_amz_giftvoucher_', [
      { name: 'Parent', value: parentName }
    ]);
    Logger.log('[GiftCard] gcSendAskMessage jlid=' + jlid + ' phone=' + parentPhone + ' result=' + JSON.stringify(wr));
    return { success: true, phone: parentPhone, watiResult: wr, parentName: parentName };
  } catch(e) {
    Logger.log('[GiftCard] gcSendAskMessage ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── SEND ISSUED MESSAGE (standalone, after entry is already in sheet) ─────────
function gcSendIssuedMessage(rowIndex, jlid, gcType) {
  try {
    var sheet = _getGiftCardSheet();
    var r = sheet.getRange(rowIndex, 1, 1, GC_LAST_COL).getValues()[0];
    var amount   = String(r[GC_COL.AMOUNT   - 1] || '').trim();
    var currency = String(r[GC_COL.CURRENCY - 1] || '').trim();

    var hs = fetchHubspotByJlid(jlid);
    if (!hs || !hs.success || !hs.data) return { success: false, message: 'Deal not found for ' + jlid };
    var parentPhone = String(hs.data.parentContact || '').replace(/\D/g, '');
    var parentName  = String(hs.data.parentName || 'Parent').trim().split(' ')[0];
    if (!parentPhone) return { success: false, message: 'No parent phone' };

    var templateName = (gcType === 'sales') ? 'amz_giftcard_sales_pitch' : 'amz_giftcard_sent_';
    var wr = sendWatiMessage(parentPhone, templateName, [
      { name: 'Parent', value: parentName },
      { name: 'amount', value: amount + ' ' + currency }
    ]);
    if (wr && wr.result) {
      sheet.getRange(rowIndex, GC_COL.ISSUED_SENT_AT).setValue(new Date());
    }
    return { success: true, watiResult: wr };
  } catch(e) {
    Logger.log('[GiftCard] gcSendIssuedMessage ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── EMAIL — referral reward ───────────────────────────────────────────────────
var GC_AMAZON_SITES = [
  { label: 'Amazon.co.uk',  host: 'amazon.co.uk' },
  { label: 'Amazon.com',    host: 'amazon.com' },
  { label: 'Amazon.ae',     host: 'amazon.ae' },
  { label: 'Amazon.in',     host: 'amazon.in' },
  { label: 'Amazon.ca',     host: 'amazon.ca' },
  { label: 'Amazon.com.au', host: 'amazon.com.au' },
  { label: 'Amazon.sg',     host: 'amazon.sg' },
  { label: 'Other',         host: '' }
];
var GC_CURRENCY_SYMBOLS = { GBP: '£', EUR: '€', USD: '$', INR: '₹', CAD: 'C$', AUD: 'A$', SGD: 'S$', AED: 'AED ' };

function _gcRender(templateName, data) {
  var t = HtmlService.createTemplateFromFile(templateName);
  t.data = data;
  return t.evaluate().getContent();
}

function _gcParentEmails(hsData) {
  return String((hsData && hsData.parentEmail) || '')
    .split(/[;,\s]+/)
    .filter(function(e) { return e && isValidEmail(e); })
    .filter(function(e, i, arr) { return arr.indexOf(e) === i; });
}

function _gcAmountDisplay(amount, currency) {
  var a = String(amount || '').replace(/\.00$/, '').trim();
  var c = String(currency || '').trim().toUpperCase();
  if (!a) return '';
  return GC_CURRENCY_SYMBOLS[c] ? GC_CURRENCY_SYMBOLS[c] + a : (a + (c ? ' ' + c : ''));
}

// "https://www.amazon.co.uk/..." / "Amazon UK" -> { host, label }; unknown -> raw text, no host
function _gcResolveSite(website) {
  var raw = String(website || '').trim();
  var m = raw.toLowerCase().match(/amazon\.[a-z.]+/);
  if (m) return { host: m[0].replace(/\.+$/, ''), label: m[0].replace(/\.+$/, '') };
  var low = raw.toLowerCase();
  if (/\buk\b|united kingdom|britain/.test(low)) return { host: 'amazon.co.uk', label: 'amazon.co.uk' };
  if (/\buae\b|dubai|emirates/.test(low))        return { host: 'amazon.ae',    label: 'amazon.ae' };
  if (/\bindia\b/.test(low))                     return { host: 'amazon.in',    label: 'amazon.in' };
  if (/\busa?\b|united states/.test(low))        return { host: 'amazon.com',   label: 'amazon.com' };
  return { host: '', label: raw };
}

function _gcSiteLinks(jlid) {
  var to = CONFIG.EMAIL.FROM;
  var subject = 'Referral reward — Amazon details (' + jlid + ')';
  return GC_AMAZON_SITES.map(function(s) {
    var body = 'Hi JetLearn team,\n\nAmazon website: ' + (s.host || '(please type it here)') +
               '\nEmail on my Amazon account: \n\nLearner ID: ' + jlid;
    return { label: s.label, href: 'mailto:' + to + '?subject=' + encodeURIComponent(subject) + '&body=' + encodeURIComponent(body) };
  });
}

// Asks the parent for their Amazon website + the email on their Amazon account.
function gcSendAskEmail(jlid) {
  try {
    jlid = String(jlid || '').trim().toUpperCase();
    if (!jlid) return { success: false, message: 'JLID required' };
    var hs = fetchHubspotByJlid(jlid, true);
    if (!hs || !hs.success || !hs.data) return { success: false, message: 'Deal not found for ' + jlid };

    var emails = _gcParentEmails(hs.data);
    if (!emails.length) return { success: false, message: 'No valid parent email for ' + jlid };

    var htmlBody = _gcRender('GiftCardAskTemplate', {
      parentName: String(hs.data.parentName || '').trim().split(/\s+/)[0],
      siteLinks:  _gcSiteLinks(jlid)
    });
    sendTrackedEmail({
      to: emails.join(','),
      subject: '🎁 Your JetLearn referral reward — two quick details',
      htmlBody: htmlBody,
      jlid: jlid
    });
    try { addNoteToHubSpotDeal(hs.data.dealId, '📧 Referral reward email sent asking for Amazon website + account email → ' + emails.join(', ')); } catch(ne) {}
    Logger.log('[GiftCard] gcSendAskEmail jlid=' + jlid + ' to=' + emails.join(','));
    return { success: true, to: emails.join(', ') };
  } catch(e) {
    Logger.log('[GiftCard] gcSendAskEmail ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// Pulls the Amazon website + account email out of a parent's reply (quoted text removed).
function _gcParseAmazonReply(body) {
  var text = String(body || '').replace(/\r/g, '');
  var cut = text.search(/\n\s*(On [^\n]{5,200}wrote:|-{2,}\s*Original Message|>)/i);
  if (cut >= 0) text = text.slice(0, cut);

  var site = '', email = '';
  var m = text.match(/Amazon website:[ \t]*([^\n]+)/i);
  if (m) {
    var cand = m[1].trim();
    if (cand && !/please type it here/i.test(cand)) site = _gcResolveSite(cand).label;
  }
  if (!site) {
    var am = text.match(/amazon\.(?:co\.uk|com\.au|com|ae|in|ca|sg|de|fr|es|it|nl|sa|eg|jp)\b/i);
    if (am) site = am[0].toLowerCase();
  }

  var em = text.match(/Email on my Amazon account:[ \t]*([^\s<>,;]+@[^\s<>,;]+)/i);
  if (em) email = em[1];
  if (!email) {
    var all = text.match(/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g) || [];
    for (var i = 0; i < all.length; i++) {
      if (!/jet-?learn/i.test(all[i])) { email = all[i]; break; }
    }
  }
  return { website: site, amazonEmail: email, snippet: text.trim().substring(0, 300) };
}

// Latest parent reply to the "Amazon details" email, so ops can see it while issuing the card.
function gcGetAmazonDetails(jlid, parentEmails, parentPhone) {
  try {
    jlid = String(jlid || '').trim().toUpperCase();
    if (!/^[A-Z0-9]{6,}$/.test(jlid)) return { success: false, message: 'Valid JLID required' };

    var emails = String(parentEmails || '').split(/[;,\s]+/).filter(function(e) { return e && isValidEmail(e); });
    var q = 'newer_than:180d (subject:("Amazon details" "' + jlid + '")';
    if (emails.length) q += ' OR (from:(' + emails.join(' OR ') + ') subject:("referral reward"))';
    q += ')';

    var candidates = [];
    GmailApp.search(q, 0, 10).forEach(function(th) {
      th.getMessages().forEach(function(m) {
        if (/jet-?learn/i.test(String(m.getFrom() || ''))) return;   // skip our own sends / team replies
        var p = _gcParseAmazonReply(m.getPlainBody());
        candidates.push({ when: m.getDate(), from: m.getFrom(), parsed: p, source: 'Email' });
      });
    });

    // WhatsApp (WATI) replies — only incoming messages that actually contain a site or email,
    // since a chat has plenty of unrelated messages.
    try {
      var phone = String(parentPhone || '').replace(/\D/g, '');
      if (phone) {
        var chat = getWatiChatData(phone, 'giftcard');
        if (chat && chat.success) {
          chat.messages.forEach(function(wm) {
            if (wm.direction !== 'in') return;
            var wp = _gcParseAmazonReply(wm.text);
            var wt = new Date(wm.timestamp);
            if (!(wp.website || wp.amazonEmail) || isNaN(wt.getTime())) return;
            candidates.push({ when: wt, from: 'WhatsApp +' + phone, parsed: wp, source: 'WhatsApp' });
          });
        }
      }
    } catch(we) { Logger.log('[GiftCard] WATI read skipped: ' + we.message); }

    if (!candidates.length) return { success: true, found: false };

    candidates.sort(function(a, b) { return b.when - a.when; });
    // Each field comes from the newest reply (email or WhatsApp) that actually gave it
    var withSite  = candidates.filter(function(c) { return c.parsed.website; })[0];
    var withEmail = candidates.filter(function(c) { return c.parsed.amazonEmail; })[0];
    var used = [withSite, withEmail].filter(Boolean);
    var latest = used.length ? used.sort(function(a, b) { return b.when - a.when; })[0] : candidates[0];
    var sources = [];
    (used.length ? used : [candidates[0]]).forEach(function(c) { if (sources.indexOf(c.source) === -1) sources.push(c.source); });
    return {
      success: true,
      found: true,
      website: withSite ? withSite.parsed.website : '',
      amazonEmail: withEmail ? withEmail.parsed.amazonEmail : '',
      snippet: latest.parsed.snippet,
      from: latest.from,
      source: sources.join(' + '),
      repliedAt: _gcFormatDateTime(latest.when)
    };
  } catch(e) {
    Logger.log('[GiftCard] gcGetAmazonDetails ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// Sends the gift-wrapped "your gift card is here" email for a recorded sheet row.
function gcSendIssuedEmail(rowIndex, jlid, gcType) {
  try {
    jlid = String(jlid || '').trim().toUpperCase();
    if (!jlid) return { success: false, message: 'JLID required' };
    var sheet = _getGiftCardSheet();
    var r = sheet.getRange(rowIndex, 1, 1, GC_LAST_COL).getValues()[0];
    var amountDisplay = _gcAmountDisplay(r[GC_COL.AMOUNT - 1], r[GC_COL.CURRENCY - 1]);
    if (!amountDisplay) return { success: false, message: 'No amount on this row' };
    var rowEmail = String(r[GC_COL.EMAIL - 1] || '').trim();

    var hs = fetchHubspotByJlid(jlid, true);
    if (!hs || !hs.success || !hs.data) return { success: false, message: 'Deal not found for ' + jlid };
    var emails = _gcParentEmails(hs.data);
    if (!emails.length && isValidEmail(rowEmail)) emails = [rowEmail];
    if (!emails.length) return { success: false, message: 'No valid parent email for ' + jlid };

    var site = _gcResolveSite(r[GC_COL.WEBSITE - 1]);
    var type = String(gcType || r[GC_COL.TYPE - 1] || 'referral').toLowerCase();
    var htmlBody = _gcRender('GiftCardIssuedTemplate', {
      parentName:    String(hs.data.parentName || '').trim().split(/\s+/)[0],
      learnerName:   String(r[GC_COL.LEARNER_NAME - 1] || '').trim().split(/\s+/)[0],
      amountDisplay: amountDisplay,
      website:       site.label,
      redeemUrl:     'https://www.' + (site.host || 'amazon.com') + '/gc/redeem',
      sentTo:        isValidEmail(rowEmail) ? rowEmail : '',
      isSales:       type === 'sales'
    });
    sendTrackedEmail({
      to: emails.join(','),
      subject: '🎁 Your ' + amountDisplay + ' Amazon gift card from JetLearn',
      htmlBody: htmlBody,
      jlid: jlid
    });
    try { addNoteToHubSpotDeal(hs.data.dealId, '📧 Gift card email sent (' + amountDisplay + ') → ' + emails.join(', ')); } catch(ne) {}
    Logger.log('[GiftCard] gcSendIssuedEmail row=' + rowIndex + ' jlid=' + jlid + ' to=' + emails.join(','));
    return { success: true, to: emails.join(', ') };
  } catch(e) {
    Logger.log('[GiftCard] gcSendIssuedEmail ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── HELPERS ───────────────────────────────────────────────────────────────────
function _gcFormatDMY(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return '';
  return ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth()+1)).slice(-2) + '/' + d.getFullYear();
}

function _gcFormatDateTime(v) {
  if (!v) return '';
  var d = (v instanceof Date) ? v : new Date(v);
  if (isNaN(d.getTime())) return String(v);
  return _gcFormatDMY(d) + ' ' + ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2);
}
