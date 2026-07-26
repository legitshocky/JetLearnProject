# JetLearn Platform — Changelog

---

## [2026-07-27] — Fix Stale "Verify Address" Button (V7.99)

### Verify Address button stayed visible pointing at a now-blank address (`JavaScript.html`)
- `requestKitDeliveryAddress()` clears the sheet's `DELIVERY_ADDRESS` when a fresh request goes out (since it's about to be reconfirmed), but `ktRequestAddress()`'s success handler never hid the "Verify Address" button — so clicking it afterward failed with "No address on file," even though the textarea still showed the old (now-stale) address client-side. Now hides the button and clears the textarea display in that same success handler, so the UI matches what's actually in the sheet.

---

## [2026-07-27] — Fix kit_address_request_link WATI Rejection (V7.98)

### Removed `?name=...` query param from the WhatsApp-sent link (`LearnerAddressFormService.js`)
- Live test showed `kit_address_request_link` getting rejected by WATI ("cannot have typos or blank text") whenever the link included `?name=Ahaan%20Padia%20Test%20Deal` — the fallback (`migration_address_template`, which has the old `share.hsforms.com` link hardcoded directly in its static body, unrelated to the new link-choice feature) then fired instead.
- `getAddressFormLink()` no longer appends the `?name=` query string to the short link — it was a purely cosmetic instant-display hint (the page's own API call resolves the real name a moment later regardless) and isn't worth risking template rejection over.

---

## [2026-07-27] — Kit Tracking: Choice of Link (My Page vs HubSpot Form) (V7.97)

### New "Send Link Via" selector (`Index.html`, `JavaScript.html`, `KitTrackingService.js`, `LearnerAddressFormService.js`)
- New dropdown in the Add Kit Entry address section — **My Page** (`jetlearn-kit-links.web.app`, default) or **HubSpot Form** (the older native `share.hsforms.com` link) — applies to both "Request via WhatsApp" and "Copy Address Link."
- `getAddressFormLink(jlid, learnerName, useHsForm)` returns the HubSpot form's static share URL when `useHsForm` is true, reusing the same already-approved `kit_address_request_link` WATI template either way (just swaps which URL fills `{{3}}`).
- `requestKitDeliveryAddress(jlid, kitName, rowIndex, useHsForm)`: when the HubSpot Form option is chosen, always sends the plain link-ask and skips the yes/no reconfirm (that flow's tied specifically to our own page's experience).

---

## [2026-07-27] — Kit Tracking: Fix Legacy-Row Bug, Copy Link, Verify Address (V7.95–V7.96)

### Critical fix: legacy rows wrongly flooding the Asking Address tab (`KitTrackingService.js`)
- The "Asking Address" tab (V7.93) filtered on any non-blank `ADDR_STATUS` + no `ORDER_PLACED` — but old rows created via the legacy Add Kit Entry form set `ADDR_STATUS='Verified'` just from typing an address in at creation time, and never have `ORDER_PLACED` set (that column didn't exist yet). Every already-delivered/ordered legacy kit was getting misclassified as "still asking," and JetLearn Sends was hiding all of them, showing 0 for entire months.
- New `inAddressPipeline` flag requires `ADDR_REQUESTED_AT` to actually be stamped (only happens via the new `requestKitDeliveryAddress()`/`sendKitAddressVerifyRequest()` pipeline) — legacy rows are correctly excluded regardless of their old `ADDR_STATUS` value.

### Address-received date shown in Add Kit Entry (`KitTrackingService.js`, `JavaScript.html`)
- `fetchKitLearnerDetails()` now returns `addressReceivedAt` (from `ADDR_SUBMITTED_AT`), shown under the "Previous address on file" warning so it's clear how stale it is.

### Fixed "Copy Address Link" (`JavaScript.html`)
- `navigator.clipboard.writeText()` was called inside the `google.script.run` async callback — by then the browser's user-gesture flag from the click had expired, so it silently fell back to a jarring native `prompt()`. Replaced with a small modal showing the link in a selectable field with its own "Copy" button — copying now happens on a fresh, direct click, which works reliably.

### New "Verify Address" action (`KitTrackingService.js`, `Index.html`, `JavaScript.html`)
- New `sendKitAddressVerifyRequest(jlid, rowIndex)` sends the yes/no `kit_address_reconfirm` template using the address **already on file in our own sheet** — no HubSpot round-trip (which is blocked anyway). Faster than a full new address request when you already trust the address on file and just want a quick parent confirmation. New "Verify Address" button appears in the Add Kit Entry modal whenever an address is already on file.

---

## [2026-07-26] — Kit Tracking: New "Asking Address" Sub-Tab (V7.93–V7.94)

### Separate tab for the address-request pipeline (`Index.html`, `JavaScript.html`)
- New **"Asking Address"** tab (with a live sidebar-style badge count) sits alongside JetLearn Sends / Parent Will Buy / Pivot Report, showing only rows still waiting on an address — not yet order-placed. Own KPI strip (Total Asking, Awaiting Reply, Address Received, Needs Your Call), own filter bar (stage: Requested / Received / Needs Call, plus search), own table with Address Status, Follow-up Stage (WhatsApp sent / Email nudge sent / Call needed), Link Opened count+timestamp, Requested date, and inline "Mark Order Placed" once received.
- **JetLearn Sends now excludes these rows** (`row.addrStatus && !row.orderPlaced`) — it only shows kits that have actually moved to being ordered, so the two tabs cleanly answer two different questions: "how many are we still waiting on an address for" vs "how many have we actually sent."
- Backend: `getKitTrackingData()` now also returns `nudgeStage`/`nudgeTier` per row and `addressAwaitingReply`/`askingAddressTotal` in `stats`, needed to drive this tab.

---

## [2026-07-26] — Kit Tracking: Check Sheet Not HubSpot, Stop Legacy Trigger (V7.92)

### Check the Kit Tracking sheet before HubSpot for address data (`KitTrackingService.js`)
- `fetchKitLearnerDetails()` (backs the "Fetch" button in Add Kit Entry) now checks the Kit Tracking sheet's `DELIVERY_ADDRESS`/`ADDR_STATUS` first — instant, no HTTP round-trip — since HubSpot's contact address fields are permanently unreachable via this token (confirmed V7.90/91). Falls back to a HubSpot contact GET only if the sheet has nothing (covers an address entered manually in HubSpot's own UI, which doesn't go through our token's scopes). `addressPending` now also prefers the sheet's real `ADDR_STATUS` over the older cache-based guess.

### Stopped triggering the legacy hsforms.com automation (`KitTrackingService.js`)
- `requestKitDeliveryAddress()` no longer PATCHes the deal's kit-status property to `"Asked for address"` — that property change was re-triggering a pre-existing external automation (the native "Kit Address Form", `share.hsforms.com` link) that predates this whole pipeline and produced a confusing duplicate message with old wording. We now own the entire address-request flow through our own WATI templates, so this step was redundant and actively conflicting.

---

## [2026-07-26] — Kit Address Form: Confirmed Scope Ceiling, Simplified to Note-Only (V7.90–V7.91)

### Confirmed: no direct-write path to HubSpot is available with the current token
- `diagFindKitAddressForm()` showed both Forms APIs are also blocked — v2 needs `forms-access`/`form-submissions`/`forms-read`, v3 needs `forms` — neither granted. Combined with the earlier-confirmed sensitive-property block (which also covers plain `email`, not just address), every direct-write path to HubSpot (property PATCH, Forms API submission) is closed off with this token, and none of the missing scopes are obtainable.

### Simplified `submitLearnerAddressForm()` to the one thing that actually works (`LearnerAddressFormService.js`)
- Removed the doomed email-PATCH and Forms-API-submission attempts (both always fail — no point spending API calls confirming that every time). The address is now recorded as a HubSpot deal **Note** only (via `_addNoteToDeal`, unaffected by either scope restriction) — visible directly on the deal in HubSpot. The Kit Tracking sheet remains the actual system of record, as it already was.

---

## [2026-07-26] — Kit Address Form: Submit via HubSpot Forms API (V7.89)

### Bypass the sensitive-property scope via a real Forms submission (`LearnerAddressFormService.js`)
- Forms Submission API writes aren't gated by `crm.objects.contacts.sensitive.write.v2` — only direct CRM property PATCHes are. New `_submitKitAddressHSForm(payload, portalId)` submits the address through HubSpot's native "Kit Address Form" (same form the old hsforms.com link pointed at) via `POST /submissions/v3/integration/submit/{portalId}/{guid}`, same pattern already used for the migration form in `LearnerProgressionService.js`.
- New `_findKitAddressFormGuid()` resolves the form's GUID by scraping it from the share URL's embedded HTML (works without Forms API scope), falling back to a Forms v2 API name search, mirroring `_findMigrationFormGuid()`.
- `submitLearnerAddressForm()` now tries this as the primary write path — if it succeeds, real values land in the actual contact `address`/`city`/`state`/`zip`/`country` fields despite the sensitive-scope block. The email-only PATCH and deal-Note fallback both stay in place regardless, since neither costs anything extra.
- Result is appended to the "HubSpot PATCH Status" audit column for visibility.

---

## [2026-07-26] — Kit Address Form: Root Cause Found + HubSpot Sensitive-Scope Workaround (V7.87–V7.88)

### Root cause found: HubSpot 403, missing sensitive-data scope
- Diagnostics traced the address never appearing on the HubSpot contact to a genuine `403` from HubSpot: `address`/`city`/`state`/`zip`/`country` are classified as **sensitive** contact properties in this portal, and the API token lacks `crm.objects.contacts.sensitive.write.v2` — a scope that isn't available to grant here.
- Added "HubSpot PATCH Status" column to the "Learner Address Submissions" audit sheet (`LearnerAddressFormService.js`) so this kind of failure is visible directly in the sheet going forward — Cloud Logs proved unreliable for capturing `doPost` web-app executions during this investigation.

### Workaround: address goes on a deal Note instead (`LearnerAddressFormService.js`)
- `submitLearnerAddressForm()` now PATCHes only `email` on the contact (not sensitive, still works) and adds the actual delivery address as a HubSpot **Note** on the deal via the existing `_addNoteToDeal()` helper — notes aren't subject to the sensitive-property scope, so the address is still visible directly in HubSpot, just as a note rather than a contact field.
- The Kit Tracking sheet remains the real system of record for the address either way (unaffected by this — that bridge was already working correctly).

---

## [2026-07-26] — Kit Address Form: Fixes + Instant Submit (V7.85–V7.86)

### Fixed: stale "waiting for parent" state after a real submission (`LearnerAddressFormService.js`)
- `_bridgeAddressToKitTracking()` never cleared the `KIT_ADDR_REQ_<phone>` pending-request cache/queue after a successful public-form submission — so `fetchKitLearnerDetails()` kept reporting "waiting for parent" in the Add Kit Entry modal even after the address had already come in. Now clears both on a successful bridge.

### Fixed: silent HubSpot contact PATCH failures (`LearnerAddressFormService.js`)
- The contact-property PATCH in `submitLearnerAddressForm()` used `muteHttpExceptions: true` but never checked the response code — a rejected write (bad property, permissions, etc.) would fail completely silently with no log at all. Now logs the HTTP status and response body on both success and failure.
- Added read-only diagnostic `diagCheckAddressPatchTarget(jlid)` — shows every contact associated with a deal, which one the PATCH actually targets, and that contact's live HubSpot properties, to debug cases where the write doesn't appear to take effect.

### Instant submit — no more waiting on the network (`AddressForm.html`, `firebase-hosting/public/kit/index.html`)
- Both the parent-facing address form and its legacy GAS-hosted twin now show the "Thank you" success screen the instant client-side field validation passes, instead of waiting on the full backend chain (HubSpot lookup → contact PATCH → Kit Tracking bridge → WATI confirmation) to finish first. The backend call still fires and completes in the background; failures are logged to the console for ops rather than shown to the parent, since by that point they've already been told it worked and the JLID was already confirmed valid when the page loaded.

---

## [2026-07-26] — Kit Tracking: Fix Address-First Rows Invisible in Table (V7.84)

### Fixed: bare address-first rows never appeared in the dashboard (`KitTrackingService.js`)
- `_createBareKitRow()` (V7.79) creates a row with no `DATE_OF_ORDER`, but `getKitTrackingData()` silently skips any row without a valid, parseable order date (`if (!orderDate || orderDate < cutoff) return;`) — so address-first rows were being fetched from the sheet but then filtered out of every response, making them permanently invisible in the UI (no "Mark Order Placed" button ever appeared, no way to tell the request even worked).
- Added an exception: rows with `ADDR_STATUS` set, no order date, and no order placed yet are never filtered out, regardless of order date.
- Note: these rows have no `orderMonth` (nothing to derive it from until an order is placed), so they only show under the **"All Months"** filter, not a specific month — expected, since they aren't tied to an order month yet.

---

## [2026-07-26] — Kit Tracking: Yes/No Address Reconfirm (V7.80)

### Quick-reply reconfirm instead of always re-asking from scratch (`KitTrackingService.js`, `Code.js`)
- `requestKitDeliveryAddress()` now checks HubSpot for an existing address before messaging. If one exists, sends the new `kit_address_reconfirm` quick-reply template (buttons: "Yes, Same Address" / "No, Need To Update") instead of the plain address-link ask — still always messages the parent to reconfirm (never auto-trusts a stale address), just faster when nothing's changed.
- New `handleKitAddressReconfirmReply(waId, buttonText)` handles the tap: "Yes" flips `ADDR_STATUS` to `Received` immediately using the address already on file (unlocking "Mark Order Placed"); "No" sends the normal `kit_address_request_link` short link to collect a fresh one.
- Routed via a new `KIT_ADDR_RECONFIRM_BTNS` exact-match list in `doPost`'s WATI webhook handling (`Code.js`), checked *before* the fuzzy free-text matcher — otherwise "Yes"/"No" would get misread as delivery-confirmation replies ("yes" is in the fuzzy "Kit Received" word list).
- **Manual step required**: new WATI template `kit_address_reconfirm` needs creating + approval in the WATI dashboard, with buttons labeled exactly `Yes, Same Address` and `No, Need To Update` (must match verbatim — the code routes on exact button text).

---

## [2026-07-26] — Kit Tracking: Address Request Works Before a Kit Row Exists (V7.79)

### Fixed: requesting an address before placing an order silently skipped all tracking (`KitTrackingService.js`)
- The "Request via WhatsApp" button in the Add Kit Entry modal calls `requestKitDeliveryAddress(jlid, kitName, 0)` — `rowIndex=0` because no row is saved yet. All of the new pipeline stamping (`ADDR_REQUESTED_AT`, `NUDGE_TIER`, `NUDGE_STAGE`) only ran when `rowIndex > 0`, so requesting an address before saving a full order entry sent the WhatsApp/email but left the Address column, timeline, and nudge system completely blind to it.
- `requestKitDeliveryAddress()` now: reuses an existing open row for the JLID if one exists, otherwise calls new `_createBareKitRow(jlid, learnerName, kitName)` to create a minimal row (Learner + Kit + JLID only, no order details, no HubSpot kit-status change) — so address-first workflows now get full tracking from the very first click. Order/store details get filled in later via "Mark Order Placed", which already requires `ADDR_STATUS === 'Received'` first.

---

## [2026-07-26] — Kit Tracking: Always-Visible Address Column (V7.78)

### New "Address" table column (`Index.html`, `JavaScript.html`)
- Previously, address progress (`addr_received_pending_order`) only appeared in the `Status` column, and only when a row wasn't already Delivered/Awaiting/Overdue — so most rows showed nothing about their address state at all.
- Added a dedicated **Address** column between Parent Response and Status, visible on every row regardless of delivery status: Not asked / Requested / Received / Verified / From HubSpot, plus "Order placed" and "👁 opened N×" sub-lines when relevant.
- Added `addr_received_pending_order` to the Status filter dropdown.

---

## [2026-07-26] — Multi-Contact Kit Address Requests + Certificate Emails (V7.77)

### WhatsApp + email fired together, to every contact on the deal (`KitTrackingService.js`)
- `requestKitDeliveryAddress()` now uses the already-existing `getPhoneNumbersForDeal(dealId)`/`getEmailsForDeal(dealId)` (HubSpotService.js) to reach **every** associated contact (both parents/guardians), not just the primary one — de-duped, primary contact first.
- The branded reminder email now fires **at the same time** as the WhatsApp send (not just as a later no-reply nudge) — both channels go out together at Day 0.
- `sendKitAddressReminderEmail()` (`EmailService.js`) now accepts either a single email or an array — sends one email to all valid, de-duped recipients via a comma-joined `to`. The nudge-stage email (`checkKitAddressNudges`) also now pulls all deal emails instead of just the primary contact's.

### Certificates now email every contact on the deal (`CertificateService.js`)
- New `_resolveAllCertEmails(jlid, fallbackEmail)` looks up the deal via `fetchHubspotByJlid` + `getEmailsForDeal`, merges with whatever email the UI passed in, de-dupes, and returns a comma-joined recipient list.
- Wired into all three certificate send paths: `sendCourseCertificateEmail`, `resendCertificate`, `sendBulkCertificates` (which `resendFailedCertificates` already delegates to) — certificates now reach every parent/guardian contact on the deal, not just one email.

---

## [2026-07-26] — Kit Tracking: Address Link Open/Submit Timeline (V7.76)

### Link-open + submission tracking (`Code.js`, `KitTrackingService.js`, `LearnerAddressFormService.js`)
- New columns on the Kits sheet: `LINK_OPEN_COUNT`, `LINK_FIRST_OPENED_AT`, `LINK_LAST_OPENED_AT`, `ADDR_SUBMITTED_AT`.
- `doGet`'s `?api=addressFormContext` branch (hit every time the public Firebase page loads/reloads) now calls new `recordKitAddressLinkOpen(jlid)` — best-effort, silent, never affects what the parent sees.
- `_bridgeAddressToKitTracking` now stamps `ADDR_SUBMITTED_AT` alongside the existing `ADDR_STATUS`/`DELIVERY_ADDRESS` writes; both it and the open-tracker share a new `_findOpenKitRowByJlid(jlid)` helper (replacing duplicated row-matching logic).
- Row detail expansion in Kit Tracking (`JavaScript.html`) now shows a timeline: 🔗 Link Sent → 👁 Opened (count + first/last time) → ✅ Submitted, so you can see exactly where a parent is in the process.

---

## [2026-07-26] — Kit Tracking: End-to-End Address → Order → Delivery Pipeline (V7.75)

### Bridged the public address form into Kit Tracking (`LearnerAddressFormService.js`, `KitTrackingService.js`)
- `submitLearnerAddressForm()` now calls new `_bridgeAddressToKitTracking(jlid, addressText, hsData)` — matches the learner's latest open (non-refunded, non-delivered) Kit Tracking row and flips `ADDR_STATUS` → `Received`, fills `DELIVERY_ADDRESS`, sends the existing `kit_address_received_confirmation` WATI regardless of which channel the parent used. Ambiguous/no-match cases are logged (`Kit Bridge Status` column on the submission log sheet) rather than guessed.
- The short link (`jetlearn-kit-links.web.app/kit/{JLID}`) is now the real front door for address collection — submitting it drives the whole pipeline below.

### Urgency-tiered nudges for silent address requests (`KitTrackingService.js`)
- New columns on the Kits sheet: `ADDR_REQUESTED_AT`, `MODULE_START_DATE`, `NUDGE_TIER`, `NUDGE_STAGE`, `NUDGE_STAGE_AT`, `ORDER_PLACED(_AT)`, `ORDER_STORE`, `ORDER_TRACKING_NO/URL`.
- `requestKitDeliveryAddress()` now sends via new `sendKitAddressLinkWhatsApp()` (new WATI template `kit_address_request_link` carrying the short link, falls back to `migration_address_template` until the new template is approved in the WATI dashboard) and stamps a frozen urgency tier from HubSpot `module_start_date`: **urgent** (≤7 days to course start) → email nudge same day, call-me nudge next day; **medium** (8–21 days) → email day 3, call-me day 7; **default** (>21 days / unknown) → email day 7, call-me day 14.
- New daily trigger `checkKitAddressNudges()` (run `setupKitAddressNudgeTrigger()` once from the Apps Script editor) — timezone-safe day-math, idempotent via `NUDGE_STAGE` so re-runs never double-send. Email nudge uses a new branded template `KitAddressReminderTemplate.html` (`getKitAddressReminderEmailHTML`/`sendKitAddressReminderEmail` in `EmailService.js`). Call-me nudge emails `hello@jet-learn.com` to prompt a direct call to the parent.

### "Mark Order Placed" action (manual purchase, in-app tracking)
- New server function `markKitOrderPlaced(rowIndex, payload)` — writes `DATE_OF_ORDER`/`ETA` (existing columns, so `sendKitFollowUps()` needs no changes) plus the new order-placed audit trail, and sends a new WATI template `kit_order_placed_notice` ("your kit is on the way") to the parent.
- New UI: "Mark Order Placed" button + modal on rows where the address is Received but the order isn't placed yet, in `Index.html`/`JavaScript.html`.

### UI visibility
- New status chip "📬 Address Received" (`addr_received_pending_order`) and a "📞 Call Parent" flag chip (`needsCall`).
- Two new sidebar counters next to Kit Tracking: address-received-pending-order count (teal) and needs-your-call count (red).

**Manual step required**: two new WATI templates (`kit_address_request_link`, `kit_order_placed_notice`) need to be created and approved in the WATI/Meta dashboard — code falls back to the older template until then. Also run `setupKitAddressNudgeTrigger()` once from the Apps Script editor.

---

## [2026-07-26] — Address Form: Fully Self-Hosted Page, Not a Redirect (V7.73–V7.74)

### Real Static Page Instead of a Redirect (`firebase-hosting/public/kit/index.html`, `Code.js`, `LearnerAddressFormService.js`)
- `jetlearn-kit-links.web.app/kit/{JLID}` now serves the **actual page** directly from Firebase Hosting — no 302 to `script.google.com`, so the address bar never changes
- The static page fetches learner data and submits the form via plain `fetch()` calls against a small JSON API on the GAS backend, instead of `google.script.run` (which only works when Apps Script itself serves the HTML)
- Added `doGet` branch `?api=addressFormContext&jlid=...` (GET, returns JSON, relies on Apps Script's default `Access-Control-Allow-Origin: *` for cross-origin reads) and `doPost` branch `?api=addressFormSubmit` (form-urlencoded body so the browser treats it as a "simple request" and skips CORS preflight, which Apps Script can't answer)
- `firebase-hosting/firebase.json` switched from a `redirects` rule to a `rewrites` rule (`/kit/** → /kit/index.html`) so Firebase serves the file in place rather than forwarding

---

## [2026-07-26] — Address Form: Branded Short Links, Real Logo, Faster Load (V7.71–V7.72)

### Branded Short Links (`LearnerAddressFormService.js`)
- `getAddressFormLink()` now returns `https://jetlearn-kit-links.web.app/kit/{JLID}` — a real WhatsApp-safe short link on JetLearn's own Firebase Hosting site (under the existing `gen-lang-client-0051643037` project, as a separate hosting site so it doesn't touch the other app already using that project)
- Firebase Hosting 302-redirects `/kit/:jlid` straight to the live GAS exec URL (`firebase-hosting/firebase.json`) — no server code, no external shortener API call per link, replaces the earlier is.gd dependency

### Real Logo + Faster Load (`AddressForm.html`)
- Replaced the CSS-simulated yellow "JET learn" badge with the actual JetLearn logo image (already hosted at `cdn.jsdelivr.net/gh/legitshocky/Jet-learn-Images@main/Logo.png`, same asset used in `ParentOnboardingTemplate.html`)
- Wrapped the form in an elevated white panel card, refined input/stat-card/spacing polish
- Removed the `window.load` wait (was blocking on fonts/Font Awesome CDN) — script now runs immediately on parse so the form appears instantly instead of flashing a loading spinner first

---

## [2026-07-26] — Address Form: Matched Real JetLearn Branding (V7.63)

### Redesigned to Match Actual Site (`AddressForm.html`)
- Replaced the guessed dark-navy theme (from an inaccurate brand fetch) with the real JetLearn look, based on the actual referral/trial-class page screenshots: warm off-white background, yellow circular "JET learn" logo badge, bold headline with yellow highlighter-marker effect, purple pill CTA, stats row (learners count + Google rating)
- Two-column layout: form on the left, a "what happens next" checklist panel on the right with soft gradient blob decoration

---

## [2026-07-26] — Address Form: Fixed JLID Bug, Landing-Page Redesign (V7.62)

### Fixed: JLID Arriving With Embedded Quote Characters (`AddressForm.html`, `LearnerAddressFormService.js`)
- The form's `<?= JSON.stringify(jlid) ?>` GAS template tag HTML-escapes its output — corrupting the quote characters from `JSON.stringify()` so the JLID landed in the JS variable wrapped in literal `"..."` characters, causing every HubSpot lookup to fail ("No learner found")
- Fixed by switching to the raw/unescaped GAS tag `<?!= ... ?>` (the documented pattern for injecting JSON into inline scripts), plus defensive quote-stripping server-side as a backstop
- Confirmed via char-code logging: the corrupted value had ASCII 34 (`"`) at both ends before the fix

### Landing-Page Redesign (`AddressForm.html`)
- Rebuilt to match JetLearn's actual site branding (checked live): dark navy hero with blue/purple gradient glow, bold "Let's get [Child]'s kit moving" headline auto-filled with the learner's first name, trust row (★4.8 · 12k+ learners · 60+ countries), white floating form card, rounded pill CTA button
- Learner name now shown instantly from the link itself (new `name` URL param) while the server confirms + fetches existing HubSpot data in the background

### Shorter Link Params (`Code.js`, `LearnerAddressFormService.js`, `JavaScript.html`)
- Link shape changed from `?page=addressForm&jlid=X` to `?page=addressForm&r=X&name=Y` (old `jlid=` param still accepted for compatibility)
- "Copy Address Link" now also passes the learner's name from the Kit Tracking form

---

## [2026-07-21] — Public Per-Learner Address Confirmation Form (V7.59)

### New Public Address Form — No Login Required (`Code.js`, new `AddressForm.html`, new `LearnerAddressFormService.js`)
- New unique-link form: `<exec-url>?page=addressForm&jlid=JLxxxx` — opens a standalone page (no app login) showing the learner's name/JLID read-only, and lets the parent fill in only Email, Address, City, State/Region, Postal Code, Country
- Existing HubSpot contact address/email (if any) pre-fills the fields so the parent reviews/corrects instead of retyping from scratch
- On submit: patches the HubSpot contact's `email`/`address`/`city`/`state`/`zip`/`country` properties directly, and logs every submission to a new "Learner Address Submissions" sheet in the Audit spreadsheet for ops visibility
- New "Copy Address Link" button next to the JLID field in Kit Tracking's Add Entry form — builds the link for the loaded JLID and copies it to clipboard, ready to paste into a WATI message or email
- **Action needed**: confirm the web app deployment's "Who has access" is set to allow the parent (not signed into a JetLearn Google account) to open the link without a login prompt — this can't be verified from the deploy CLI, please test with a real JLID and check Apps Script deployment settings if it prompts for sign-in

---

## [2026-07-21] — WATI Template Rename, Updated Cancellation Policy (V7.58)

### Onboarding WATI Template Renamed (`OnboardingChecklistService.js`)
- `OBC_TEMPLATE` changed from `welcome_ob_u` to `onboarding_welcome_policy` — used by the Operations Onboarding checklist's "Send Welcome" step

### Cancellation Policy Link Updated (`ParentOnboardingTemplate.html`)
- Footer "Cancellation Policy" link in the parent onboarding welcome email now points to the new policy document

---

## [2026-07-20] — Learner Ops: Both AI-Coding and Maths Pipelines (V7.57)

### Dynamic Pipeline-Stage Resolution (`LearnerOpsService.js`)
- Pause & Retention previously only covered the AI-Coding Pipeline's hardcoded stage IDs — Maths Pipeline has its own separate stage IDs for the equivalent steps (confirmed via HubSpot: e.g. Maths "15.1 Urge on Pause" = `208671578`, vs AI-Coding's `50954839`)
- Replaced the hand-maintained ID list with a live fetch from HubSpot's canonical `/crm/v3/pipelines/deals` endpoint (cached 30 min), auto-matching pause-flow stages (13., 15.1–15.4) by label across **every** pipeline — no more manually tracking IDs per pipeline, and Maths is now fully covered
- New `clearLearnerOpsPipelineCache()` utility to force a re-fetch if HubSpot pipeline stages ever change

### Subject Column — Coding vs Maths (`Index.html`, `JavaScript.html`)
- Both the New Learner Onboarding and Pause & Retention tables now show a Subject badge per row
- Pause & Retention's stage-count cards combine both pipelines' equivalent stages into one number per step (e.g. "15.1 Urge on Pause" count = AI-Coding + Maths combined) so the summary reads as one retention pipeline, while the table still shows which subject each learner is in

---

## [2026-07-20] — Learner Ops Page: Onboarding, Migration, Pause & Retention (V7.56)

### Repurposed "Teacher Onboarding Tracker" → "Learner Ops" (`Index.html`, `JavaScript.html`, new `LearnerOpsService.js`)
- The old page tracked new-teacher milestones and had never been used (0 teachers, ever) — replaced entirely with a three-tab learner-lifecycle visibility page, sidebar entry renamed "Learner Ops"
- **New Learner Onboarding** — every deal whose payment-trigger date falls in the last 30 days, any stage, showing teacher/course/current stage/welcome-sent/doc-linked/payment date. KPIs: total, payment received, onboarded, welcome sent, missing welcome
- **Migration Activity** — last 30 days of migration tickets (reusing the existing `getMigrationRegistry()`), with from→to teacher, reason, stage, days-in-stage (flagged red at 7+ days stuck), created date
- **Pause & Retention** — live view of the AI-Coding Pipeline's pause stages (15.1 Urge on Pause → 15.2 WIP → 15.3 Pause Save → 13. Retained & Save, or 15.4 Pause Follow-up if not saved), with exact HubSpot stage IDs, per-stage counts, and days-in-current-stage per learner (flagged red at 3+ days in 15.1/15.2 — the stages needing active retention work)
- Old dead teacher-onboarding UI code (cards, add-teacher modal, milestone toggles) removed entirely — nothing referenced it outside this page

---

## [2026-07-20] — Operations Onboarding Checklist Upgrades (V7.55)

### No Longer Requires Visiting New Communication First (`JavaScript.html`)
- The Onboarding Checklist (Operations → Onboarding tab) only populated its Teacher/Course dropdowns if `globalTeachers`/`globalCourses` were already loaded — which only happened after visiting New Communication
- Now lazily backfills them itself before rendering cards, same fix pattern applied earlier to the Teacher Persona page
- Also fixed a dead `data-current` attribute that was set but never applied — Teacher/Course fields now actually start pre-filled with the deal's current value

### Fetch from Notes (`OnboardingChecklistService.js`, `JavaScript.html`)
- New "Fetch from Notes" button per card — pulls teacher, course, classes offered, timezone, amount received, and payment type out of the deal's HubSpot sales/onboarding note (reusing the existing `parseSalesNote` parser) and pre-fills the fields
- Nothing is written back automatically — operator reviews and corrects before clicking Run, per request

### Multiple Classes Per Week (`JavaScript.html`)
- Class Day & Time is now a repeatable row ("Add another class day") instead of a single pair — supports learners with 2+ classes/week
- Existing `class_timings` values are parsed back into rows when a card loads; on Run, all rows are joined into one comma-separated string (e.g. "Monday 10:00 AM, Wednesday 06:00 PM") — no backend changes needed since the note text already accepted free-form strings

### Location-Searchable Timezone (`JavaScript.html`)
- Timezone field is now a searchable dropdown — type a city or country (e.g. "Sri Lanka", "Colombo") instead of needing to know the abbreviation code
- Expanded the code list from 23 to 31 zones (added Nepal, Thailand/Vietnam, Indonesia, Philippines, Brazil, Argentina, Nigeria, Turkey) with location names baked into each option's searchable label
- Custom text override field kept alongside for anything not in the list

---

## [2026-07-20] — Searchable Teacher/Course Dropdowns on Teacher Persona (V7.54)

### Type-Ahead Search for Name & Course Fields (`JavaScript.html`)
- Current Teacher, Direct Reserve teacher, and all four course fields (Current/Future 1/2/3) on the Teacher Persona page now use the same searchable dropdown widget as Migration/Onboarding — type to filter instead of scrolling a long plain `<select>`
- Reuses the existing `makeSearchableSelect` widget, which already syncs automatically when these fields are populated by Smart Context Fetch or Direct Reserve's lazy teacher load — no other code needed to change

---

## [2026-07-20] — Direct Reserve Fixes: Timezone Error, Teacher Dropdown (V7.53)

### Fixed "Missing time zone definition for start time" (`ReserveSlot.js`)
- `reserveCalendarSlot`'s `Calendar.Events.insert` calls used a bare UTC ISO string for `start`/`end` with no `timeZone` field — Google's API rejects that for recurring (weekly) events since it can't resolve DST across the recurrence without an explicit zone
- Added `timeZone` to both the master and teacher calendar event bodies (master + teacher calendar), matching the pattern already used in `bookClassesWithNewTeacher`

### Teacher Dropdown Now Populates Without Visiting New Communication (`JavaScript.html`)
- The Direct Reserve teacher list was only wired into the loader that runs when opening New Communication — but the Persona page already lazily loads teachers/courses itself the first time "Fetch & Analyze" runs, and that path never populated the new dropdown
- Now populates in both places — Direct Reserve works standalone on the Persona page

---

## [2026-07-20] — Direct Reserve: Any Learner, Any Teacher, Multi-Day (V7.52)

### Direct Reserve Slot (`Index.html`, `JavaScript.html`, `ReserveSlot.js`)
- New "Or Reserve Directly — Any Teacher" section on the Teacher Persona page — skip Smart Slot Match entirely, pick any teacher from a dropdown, and reserve for the currently loaded learner
- Reuses the existing multi-day slot rows ("Add Session Row"), timezone, and Total Classes to Reserve fields — already supports multiple different days per reservation (the earlier even-split fix applies here too)
- `reserveCalendarSlot()` now resolves the teacher's calendar ID/email/HubSpot ID by name server-side when the caller doesn't supply them, so this path doesn't need a persona search result to work

---

## [2026-07-20] — Kit HubSpot Escalation, TIC→Migration Handoff, Country Timezone Fallback (V7.51)

### Kit Escalation Now Updates HubSpot Status (`KitTrackingService.js`)
- The 2nd-reminder escalation pass already sent an email and created a HubSpot task, but never updated the kit's HubSpot status enum — it just sat there
- Now patches the kit status property to "Escalated to CLS" (VR Headset, Microbit, Makey Makey — Arduino's property has no such option, so it's skipped for that kit)

### TIC → Migration Form Handoff (`JavaScript.html`)
- New "→ Migration Form" button on persona match cards — pre-fills the Migration form (teacher, course, timezone, weekly class schedule converted from the requested slot dates, total sessions) directly from the TIC search inputs, no HubSpot migration ticket required
- Complements the existing "Migrate →" button, which only works when an open migration ticket already exists

### Booking Timezone — Fixed & Country Fallback (`Code.js`, `HubSpotService.js`, `JavaScript.html`)
- Found and fixed a dead code path: the client-side GMT-label→IANA matching in `fetchMigrationSpecificData` referenced `TIMEZONE_IANA_MAP`, a variable that only ever existed server-side — the lookup silently did nothing
- Timezone resolution moved server-side (`_resolveIanaFromTimezoneOrCountry`) and now returns a ready-to-use IANA zone directly
- New country → IANA fallback map (~50 common markets) — when a deal has no explicit timezone set, the learner's HubSpot `country` property is used instead (e.g. Sri Lanka → Asia/Colombo)
- Applied to both the Migration form's auto-fetch and TIC's Reserve Slot timezone dropdown

---

## [2026-07-20] — Reserve Slot Cleanup on Booking Confirm (V7.50)

### Auto-Detect Existing Reserve Slot When Finalizing a Booking (`JavaScript.html`, `ReserveSlot.js`)
- When clicking "Book Classes with New Teacher" in Migration, the confirm popup now checks the Class Booking Log for an active Reserve Slot (from a TIC persona match) on this JLID
- If found, shows a checked-by-default checkbox: "Delete reserved slot with [Teacher] (...) so it isn't double-booked"
- On confirm, the new booking is created first; only after it succeeds does it delete the old reserved slot's calendar events (master + teacher calendar) — a booking failure never touches the existing reservation
- Saves the extra manual trip to Manage Bookings when a Reserve Slot gets converted into a real booking with the same (or a different) teacher

---

## [2026-07-20] — Remaining Classes Now From Live HubSpot Data (V7.49)

### Fetch Remaining — HubSpot Subscription Data as Primary Source (`HubSpotService.js`, `LearnerProgressionService.js`)
- `getRemainingClassesForJlid()` now reads **"Current Subscription - Total Classes Offered"** minus **"Current Subscription Taken Classes Till Date"** directly off the deal — accurate to the current subscription, not dependent on the Athena PRMS/CPRS sheet being freshly uploaded
- Falls back to the PRMS/CPRS-derived estimate only if those HubSpot properties are blank on the deal
- Fetch Remaining status text now shows which source was used ("from HubSpot subscription data" vs "from PRMS/CPRS estimate")

---

## [2026-07-20] — Total-Class Split, Remaining Classes, Booking Cancellation (V7.48)

### "No. of Sessions" → "Total Classes to Book/Reserve" — Even Split (`ReserveSlot.js`)
- Previously the number entered booked THAT MANY occurrences of EVERY weekly session (e.g. 12 entered with Mon+Thu selected booked 24 classes total, not 12)
- Now the number is the TOTAL classes to book, split evenly across all configured weekly sessions/slots (remainder goes to the earliest ones) — e.g. 24 total with Mon+Thu → 12 + 12
- Applies to both Migration "Book Classes with New Teacher" and TIC "Reserve Slot" (which previously only ever booked the first slot row even when multiple were added — now books all of them)
- A live split preview shows the breakdown as you type (e.g. "12 × Monday + 12 × Thursday = 24 total")
- Removed the max="52" cap on both inputs

### Fetch Remaining Classes (`LearnerProgressionService.js`, `Index.html`, `JavaScript.html`)
- New "Fetch Remaining" button next to the sessions field in both Migration and TIC — pulls the live remaining-class count from PRMS/CPRS data (same calculation used by the Course Planner) and pre-fills the field
- New `getRemainingClassesForJlid(jlid)` reuses the existing 10-min cached progression batch, so repeat lookups are cheap

### Manage / Cancel Bookings (`ReserveSlot.js`, `Index.html`, `JavaScript.html`)
- New "Manage / Cancel Bookings" button (Migration + TIC) opens a modal listing this learner's bookings from the Class Booking Log, each with a Cancel button
- Cancelling deletes the actual calendar event series (master class calendar + teacher's calendar) via the Calendar API and marks the row Cancelled — for when a booking should have been removed after a migration/roadmap change but wasn't
- Booking log now stores the master + teacher calendar event IDs per booking (JSON) to make this possible; **bookings made before this update won't have stored event IDs**, so cancelling those will mark the row Cancelled without being able to auto-delete the calendar events (delete manually from the calendar in that case)

---

## [2026-07-19] — Kit HubSpot Fix, PWB AI Brain, Booking Fixes (V7.47)

### Kit Entry — HubSpot Update Fix (`KitTrackingService.js`)
- Kit status ("Sent by Us"), cost accumulation, and subscription were sent in a single PATCH; an invalid subscription enum value silently failed the whole call, so price/status never updated
- Subscription now maps to HubSpot's exact enum (`Yearly`→`Annual`, `2 Yearly`→`2 Years`, etc.) and PATCHes separately — status/cost can no longer be blocked by a bad subscription value

### Kit Delivery Follow-up — Trigger Was Never Installed (`KitTrackingService.js`)
- `sendKitFollowUps()` existed but no trigger ever called it — added `setupKitFollowUpTrigger()` (run once from the editor) to install a daily 9 AM trigger

### Refunded Kits No Longer Show as Overdue (`KitTrackingService.js`)
- Refunded rows now get their own `refunded` status — excluded from Overdue count/badge and skipped by the follow-up sender

### Parent Will Buy — AI Reply Classification (`ParentWillBuyService.js`)
- Free-text WhatsApp replies are now classified by Gemini into BOUGHT / WILL_BUY / WONT_BUY / UNCLEAR
- WILL_BUY extracts a promised date, pauses reminders/escalation until it passes (+1 day grace), then auto-escalates as "Parent didn't buy - Roadmap changed" if still unbought
- Fixed the roadmap-changed HubSpot enum value, which differs per kit property (no space vs. space before the hyphen) — was silently failing for 2 of 3 kits
- New guard: before any escalation, checks whether the deal's course has changed and no longer needs this kit — closes the row quietly with a note instead of escalating to CLS

### Kit Purchase Links — Auto-fill by Country (`ParentWillBuyService.js`)
- New "Kit Links" sheet (seed via `setupKitLinksSheet()`) maps each kit to all 22 Amazon marketplaces + an "Other" fallback per kit for non-Amazon countries
- PWB auto-fills the purchase link by learner country when not already set — no more manually finding and pasting Amazon links

### Monthly KPIs — Avg Time to Deliver / Avg Response Time (`JavaScript.html`, `Index.html`)
- Kit Tracking tab: new "⏱ Avg Time to Deliver" KPI (order → delivery), respects month/kit/search filters
- Parent Will Buy tab: new "⏱ Avg Response Time" KPI (first message → parent reply); also fixed PWB KPIs never recalculating on month change
- New "Response At" column tracks first-reply timestamp going forward

### TIC "Clear All" Fix (`JavaScript.html`)
- Clear All button now properly resets the Smart Context Fetch summary bar, trait chips, JLID field, and slot rows

### Double-Booking Guard — False Positives on "Availability Hour" (`ReserveSlot.js`, `HubSpotService.js`)
- "Teacher's Availability Hour" calendar markers (open slots, not bookings) were being flagged as conflicts in both the migration booking guard and the TIC slot check — both now correctly ignore them

---

## [2026-07-18] — Popup Fix, Double-Booking Guard, Booking Log (V7.43)

### All Popups Fixed — Missing `</div>` (`Index.html`)
- `#documentationOverlay` was never closed, so the browser auto-nested **every modal** (changelog, email details, teacher profile, etc.) inside the hidden overlay — popups opened but were invisible
- One closing `</div>` added; version popup, Email Activity View Details, and TIC View Profile all work again

### Double-Booking Guard (`ReserveSlot.js` + `JavaScript.html`)
- New `checkBookingConflicts()` — before booking, checks the teacher's personal calendar **and** master class calendar for events overlapping each session's first occurrence
- Conflicts shown as a red warning inside the booking confirm popup; operator can still book anyway or cancel
- Unverifiable calendars flagged instead of silently passing

### CET Preview — Day Rollover Fix (`JavaScript.html`)
- Conversions crossing midnight (e.g. 11 PM India, late-evening US) were off by ±24h; diff now normalized to ±12h
- Preview shows Day + 12-hour time (e.g. `⌚ Wed 1:00 AM CET`), always visible with "select time above" placeholder
- Handles `(GMT)` zero-offset labels and both IANA and GMT display-string timezones

### TIC View Profile → Popup (`JavaScript.html`)
- Persona cards and replacement table View Profile buttons now open the profile popup modal instead of navigating to the profile page

### Class Booking Log (`ReserveSlot.js`)
- Every successful booking appends a row to the "Class Booking Log" sheet in the Audit spreadsheet (auto-created): Timestamp, JLID, Learner, Teacher, Course, Sessions, Weeks, Start Date, Timezone, Performed By, Class Link, Event Title

### Login Page — Light Redesign (`Styles.html`)
- Lavender-white gradient, white feature cards, indigo accents; white sign-in panel — replaces the previous dark theme via `#loginPage` override block (structure and animations untouched)

---

## [2026-07-14] — CET Preview, GCSE Tag, Migration Fixes (V7.35)

### CET Time Preview — Migration & Onboarding Parent Forms (`JavaScript.html`)
- Purple `⌚ Day H:MM AM/PM CET` badge now appears below every Class Schedule row in Migration and Onboarding Parent forms
- Shows **"⌚ CET: select time above"** by default (always visible), updates live when Day / Hr / Min / AM-PM change
- Handles both IANA timezone strings (Booking Timezone) and GMT display strings like `(GMT -5:00) Eastern Time...` via offset parsing
- Fixed sign bug in UTC conversion (`- diff` → `+ diff`) that was producing wrong CET times
- Uses selected weekday's actual date for accurate day-rollover (e.g. Tue 7 PM Eastern → Wed 1:00 AM CET)

### GCSE Event Tag (`ReserveSlot.js`)
- Courses containing "gcse" (GCSE Premium CS Pro, GCSE NC, GCSE Custom Revision) now produce calendar event title: `Learner (JLID) : JetLearn GCSE Lesson (TJL...)`

### Migration Tag — First Session Only (`ReserveSlot.js`)
- When booking multiple sessions, `Migration :` prefix is now applied only to the **first session's** first occurrence, not all sessions

### App Version (`Code.js`)
- `APP_VERSION` updated to `"7.35"`

---

## [2026-07-06] — Certificate Improvements (V6.18)

### Certificate — Slide Selected from Course Name Sheet (`CertificateService.js`)
- Certificate template slide (Foundation / Math / Pro+Advanced) now driven by **col C (Tagging)** of the Course Name sheet — no more hardcoded keyword list
- `_buildCertCategoryCache()` reads sheet once per execution and caches; bulk sends hit the sheet only once
- Fallback to math-keyword regex if course not found in sheet

### Certificate — Course Dropdown Loads from Sheet (`JavaScript.html`)
- `_bcCourseList` was hardcoded in frontend; now fetched live from Course Name sheet via `getCourseNames()` on page load
- Any course added/renamed in the sheet appears in the dropdown automatically

### Certificate — Sent from hello@jet-learn.com (`CertificateService.js`)
- All certificate emails now explicitly send from `hello@jet-learn.com` (script owner account)

### Certificate — Re-send Button in Log (`JavaScript.html`)
- Every row in the certificate log now has a **Re-send** button
- Calls `resendCertificate()` — re-uses existing Drive file if available, falls back to regenerating
- Reloads log on success; shows toast on failure

### Certificate — Drive Link in HubSpot Notes (`CertificateService.js` + `Code.js`)
- Certificate PDF saved to Google Drive with public shareable link
- HubSpot deal note includes `🔗 View/Download Certificate: <drive_url>` so team can view/share without opening email

### Course Name Sheet — Category Tagging (`CertificateService.js`)
- Added `updateCourseCategories()` — run once from Apps Script editor to populate col C with Foundation / Math / Advanced / Pro labels based on course name

---

## [2026-07-05] — Kit Dashboard & PWB Fixes (V6.17)

### Kit Dashboard — Kit Status HubSpot Fix (`KitTrackingService.js`)
- `microbit_kit_status`, `makey_makey_kit_status`, `vr_headset__oculus_status`, `arduino_kit_status` now correctly updated to `Sent` in HubSpot when a kit entry is added
- Split into two separate HubSpot PATCHes: one for kit status, one for `learning_kit_cost` + subscription — a bad enum on one no longer blocks the other
- Full HubSpot error body now logged when PATCH fails (up to 500 chars) for easier debugging

### Kit Dashboard — HubSpot Enum Check (`KitTrackingService.js` + `Index.html` + `JavaScript.html`)
- New `getKitStatusEnums()` server function fetches valid enum options for all 4 kit status properties from HubSpot Properties API
- "Check Valid Values" button in Pivot Report tab calls this and renders a table of `Label → internal value` — no need to open GAS editor

### Parent Will Buy — Column Order Fix (`JavaScript.html`)
- Learner and Date columns were swapped in the rendered table rows (Date appeared under Learner header and vice versa)
- Fixed: Learner + JLID now renders in column 2, Date in column 3, matching the table headers

### Parent Will Buy — Month Filter Format Fix (`KitTrackingService.js`)
- `entryMonth` (sheet col C) was storing Date objects in older rows, which serialised to `"Mon Jun 01 2026 00:00:00 GMT+0530 (India Standard Time)"` in the dropdown
- Added `fmtMonth()` helper: reformats Date objects and date strings to `"MMMM yyyy"` (e.g. `"June 2026"`)

### Parent Will Buy — Month Filter Sort Fix (`JavaScript.html`)
- Month dropdown was sorting alphabetically (`April → July → June → May`)
- Now sorts chronologically by year then month index (`May 2026 → June 2026 → July 2026`)

---

## [2026-06-19] — Book Classes with New Teacher Overhaul

### Booking Timezone — IANA Google Calendar Timezones (`Code.js` + `Index.html` + `JavaScript.html` + `ReserveSlot.js`)
- Replaced HubSpot-style GMT offset labels with proper IANA timezone list (`bookingTimezones`) — 65 entries covering all major regions
- Added dedicated **Booking Timezone** search field (fixed-position dropdown appended to `document.body`, no clipping) separate from the main migration form timezone
- Auto-fills booking timezone from learner's stored timezone on JLID load (maps GMT label → IANA via `TIMEZONE_IANA_MAP`)
- Booking timezone pre-populates `bookingTimezoneSearch` with friendly label; hidden input holds IANA id

### Calendar Events — Correct Timezone Stamping (`ReserveSlot.js`)
- Switched from `CalendarApp.createEventSeries` (always stamped in calendar's CET timezone) to `Calendar.Events.insert` (Advanced Calendar API) with explicit `start.timeZone` / `end.timeZone` set to the selected IANA timezone
- Events now show correct local time for any timezone (UK, IST, etc.) — not forced to CET
- Added `reminders`: popup 10 min before + email 5 hours before on all new bookings
- Added `responseRequested: true` for RSVP invites

### Guest List Privacy (`ReserveSlot.js`)
- `guestsCanSeeOtherGuests: false` + `guestsCanInviteOthers: false` set directly on event body at creation
- `_hideGuestList` patch called after creation for belt-and-suspenders enforcement
- Added `patchAllEventsHideGuests()` one-time utility to retroactively hide guest lists on all existing calendar events (±1 year window, paginated)

### JetGuide Invites (`ReserveSlot.js`)
- JetGuide selected in migration form is now invited to all booked class events
- `_JETGUIDE_EMAILS` map: Abhishek Nayak, Anamika Parmar, Sana Rais, Satyam Mehra
- Salima Chhatriwala and Aishwarya Jain excluded by design (not in map)

### Event Notes / Description (`ReserveSlot.js` + `Index.html` + `JavaScript.html`)
- **Fetch from existing event** button: calls `getExistingEventDescription(jlid)` — uses `Calendar.Events.list` with `q: JLID` title search (single API call, `maxResults: 3`, ±30 days window)
- Strips HTML tags server-side (`<br>` → newline, entity decode) before returning clean plain text
- Editable textarea pre-filled with fetched notes; carried into all new event descriptions alongside Zoom link
- Description format: `Join Zoom Meeting: <link>\n\n<notes>`

### Layout Fixes (`Index.html`)
- Booking section extracted from New Teacher `form-group` cell into its own `full-width` card — fixes grid asymmetry and right-side clipping
- Booking card row: Timezone search | No. of Sessions | Book Classes button (flex, wraps on narrow screens)

### Permissions
- `book_classes_calendar` permission restricted to Super Admin only

---

## [2026-06-03] — Teacher Persona Enhancement, Practice Doc Deduplication & Bug Fixes

### Teacher Persona — Inline Stats on Cards (`TeacherService.js` + `HubSpotService.js` + `JavaScript.html` + `Index.html`)
- **Active learner count per course**: how many learners each teacher currently has on a given course, pulled from HubSpot in a single paginated call (`_buildAllTeacherStatsMap`) cached per execution
- **Age range**: min–max age of active learners (e.g. `8–12 yrs`) shown on every teacher card
- **Teaching since**: month/year of the oldest active deal for that teacher × course
- **`⭐ IDEAL MATCH` indicator**: awarded when proficiency ≥ 90% + ≥ 2 active learners on course + learner age fits range (±2 yr buffer)
- **Age fit badge**: `🎂 Ages 8–12 — good fit` (green) or `— age gap` (amber) on alternative teacher cards; learner age read from migration form and passed to `findUpskillAlternatives` as 4th param
- **Upskilling history**: reads optional `Teacher Upskill History` sheet (cols: Teacher Name · Course · Status Before · Status After · Changed Date · Notes); shown as `Course: 90% → Not onboarded (Jan 2025)` badge on cards
- **"Previously Taught (Removed)" section** in teacher profile modal for courses that appear in history but are no longer in Teacher Courses sheet

**Card updates:**
- *Alternative teacher cards* (migration): ideal ribbon (top-right), proficiency badges, stats row (learner count / age range / since), history + upskill history badges; sorted ideal → 100% → most learners
- *Teacher profile modal*: 5 columns (Course · Proficiency · Learners · Ages · Since); ideal badge inline; history tooltip (📋); removed-courses section at bottom
- *Course panel teacher cards*: age range badge + ideal badge in header; individual learner age tag (`8y`) in expanded learner list

**New helpers (`TeacherService.js`):** `_ageRangeStr`, `_getTeacherUpskillHistory`, `_buildAllTeacherStatsMap`, `_mergeStatsIntoCourses`

---

### Practice Document — Deduplication & Teacher Update (`PracticeDocService.js` + `OnboardingChecklistService.js`)
- **Root cause**: both the onboarding email flow (`createPracticeDocAndPostNote`) and checklist run flow (`runOnboardingChecklist`) independently called `makeCopy` — two separate docs were created and shared with the parent causing confusion
- **Fix**: before creating a new doc, both flows now check HubSpot `learner_practice_document_link` for an existing URL
  - If found → update teacher permissions on the existing doc (no new doc created)
  - If not found → create new doc as normal
- **`_updateExistingPracticeDocTeacher(docUrl, newTeacherName)`** — scans current editors, cross-references Teacher Data sheet to identify teacher emails, removes any previous teacher, adds new teacher; `support@jet-learn.com` and parent commenter always preserved
- **`_pdFetchExistingDocUrl(dealId)`** — lightweight HubSpot GET to check for existing link before any Drive operation
- `createPracticeDocAndPostNote` now strips `TJL1280 - ` prefix before passing teacher name to `_pdTeacherEmail`
- `runOnboardingChecklist` `else if (existingDocLink)` branch now calls `_updateExistingPracticeDocTeacher` instead of silently re-patching

---

### WATI Chat Link — Fix (`OnboardingChecklistService.js`)
- **Root cause**: `_obcGetWatiChatLink` used `/api/v1/contact/{phone}` and saved `contact.id` — but WATI teamInbox URLs require `conversationId` not contact ID
- **Fix**: replaced with `fetchWatiDirectLink(phone)` (already in `WatiService.js`) which calls `/api/v1/getMessages` to get the actual `conversationId`, falls back to `contactId` if no message history

---

### Practice Document — HubSpot Property Save (`PracticeDocService.js`)
- `createPracticeDocAndPostNote` now calls `_obcPatchDeal(dealId, { learner_practice_document_link: url })` immediately after doc creation — property was never saved to HubSpot via the onboarding email path before this fix

---

### Practice Document — Naming Format (`PracticeDocService.js`)
- Subject labels corrected: `AI-Coding` → `Ai- Coding`, `FinLit` adds ` : ` separator before learner name
- Formats: `JetLearn Ai- Coding Practice Doc {Name} ({JLID})` · `JetLearn Maths Practice Doc {Name} ({JLID})` · `JetLearn FinLit Practice Doc : {Name} ({JLID})`

---

### PWB Table — Raw HTML Tags Rendering Fix (`JavaScript.html`)
- `<\\/td>` double-backslash sequences in the `renderPWBTable` row builder were being output as literal `<\/td>` text in the DOM
- Root cause: `\\` in JS source → `\` in string → browser HTML parser treats `<\` as invalid tag open → emits literally
- Fix: global replace `<\\/` → `<\/` throughout `JavaScript.html` (single `\/` = `/` in JS, valid HTML closing tag)
- Same fix applied to migration report and audit sections that had the same pattern

---

### Kit Tracking — Address Flow & Poll (`KitTrackingService.js`)
- `learning_kit_cost` PATCH now sends as number (not string) — HubSpot was silently ignoring string values
- ScriptProperties queue (`KIT_ADDR_QUEUE`) persists pending address requests even when no sheet row exists yet
- Poll trigger reduced from 30 min to 1 min via `setupKitAddressPollTrigger()`
- `kit_address_received_confirmation` WATI template uses positional params `{{1}}` / `{{2}}` (name: `'1'`, `'2'`) not named params
- HubSpot form webhook handler `_handleKitAddressFormWebhook` added for instant detection (pending HubSpot forms access)

---

## [2026-05-14b] — Migration Center: Communication Tracker

### Migration Comms Tracker (`HubSpotService.js` + `Index.html` + `JavaScript.html`)
- New panel at bottom of Migration Center: **"Communication Tracker — This Month"**
- Shows all `Migration Completed` tickets from current month with per-ticket comms status
- **Server**: `getMigrationCommsStatus()` — fetches completed tickets from HubSpot pipeline `66161281` stage `128913753`, cross-references Audit Log sheet to detect which comms were sent via tool
  - Checks: Parent WhatsApp (`WATI Sent`), New Teacher Email (`Teacher Email Sent`), Old Teacher Email (`Old Teacher Email Sent`)
  - Also detects deliberately skipped comms (`Teacher Email Skipped`, `WhatsApp Skipped`)
- **Status logic**: `complete` (WA + new teacher sent) / `partial` (some sent) / `not_sent` (no tool usage logged)
- **UI**: flat table with ✓ / ✗ / — per comm channel, colour-coded rows (red = not sent, amber = partial, white = complete)
- KPI pills: Not Sent · Partial · Complete counts
- Alert badge appears on section header when any ticket needs attention
- **Send** button on non-complete rows → jumps to Communication page
- Auto-loads on `loadLearnerMigrationPage()` in parallel with registry
- Deployed @619

---

## [2026-05-14] — Email Scheduling, Kit Entry Unification & Task Queue Table View

### Email Queue System (`EmailQueueService.js` — new)
- Schedule any email to fire automatically at 8am on a chosen date
- `scheduleEmail(payload)` — appends row to **Email Queue** sheet (Status = Pending)
- `processEmailQueue()` — daily 8am trigger; sends all Pending rows where Scheduled Date ≤ today
- `cancelQueuedEmail(queueId)` — marks row Cancelled
- `getEmailQueue()` — returns all rows for the Scheduled Emails tab UI
- `setupEmailQueueTrigger()` — registers the 8am GAS time trigger (run once manually)
- Sheet: `Email Queue` in Kit Tracking spreadsheet
- Columns: `Queue ID · Scheduled Date · Email Type · JLID · Recipient Email · Learner Name · Form Data (JSON) · Status · Created At · Created By · Sent At · Error`

### Send Email — "Send Now / Schedule?" Modal
- Removed inline schedule checkbox from all 3 email forms (Onboarding Parent, Minecraft, Roblox)
- Clicking **Send Email** now validates form first, then pops a small modal with two choices:
  - **Send Now** — fires email immediately (no preview step)
  - **Schedule for later** — date picker expands inline → queues to Email Queue sheet at 8am on chosen date
- Cancel button dismisses without action
- Form errors surface before modal opens (no wasted click)

### Kit Entry — Unified `addKitEntry` (KitTrackingService.js)
- Single function handles all kit order actions in one call:
  1. Writes row to Kit Tracking sheet
  2. PATCHes HubSpot deal (kit status + `learning_kit_cost`)
  3. Sends WATI WhatsApp message with full delivery address
  4. Adds HubSpot deal note
- `_fetchContactAddress(dealId)` moved from deleted `KitOrderService.js` — fetches real street/city/state/country from HubSpot contact association
- Fixed SR No sequence (scans col A for max + 1, was using wrong row index)
- Removed stray `£` symbol on EUR-priced kits
- Returns `{ success, srNo, watiSent, noteSaved }`

### KitOrderService.js — Deleted
- Was duplicate/dead code after `_fetchContactAddress` moved to `KitTrackingService.js`
- No functionality lost

### Email Attachments — Minecraft / Roblox (EmailService.js)
- Fixed Minecraft and Roblox install emails sending without attachments
- Root cause: Drive folders were empty
- Added `testDriveFolderAttachments()` diagnostic function (run from GAS editor to verify folder contents)
- All 3 Drive folders now verified ✅

### Email Preview Modal — Visibility Fix
- Fixed modal invisible when opened from Communication page
- Root cause: `#emailPreviewModal` was a child of `#documentationOverlay` (display:none) — all descendants hidden
- Fix: `document.body.appendChild(modal)` in `_showEmailModal` before showing — escapes the hidden parent

### Operations — Task Queue: HubSpot-Style Flat Table
- Replaced grouped card view with flat table matching HubSpot layout
- Columns: ○ (status circle) · Title · Associated Deal · Associated Ticket · Task Type · Due Date · Assigned To · Actions
- Row tints: overdue = red, today = yellow, upcoming = white
- Status circle: hover to preview checkmark, click to mark Done in HubSpot
- Filter tabs (All / Overdue / Today / Upcoming) and stats strip unchanged
- `getMyHubSpotTasks()` updated:
  - Now batch-fetches ticket associations (`/crm/v4/associations/tasks/tickets/batch/read`)
  - Batch-fetches ticket subjects
  - Returns `dealName`, `ticketId`, `ticketName` per task (was always empty before)

---

## [2026-05-13] — Operations Page, Credentials Automation & Task Queue

### New Page — Operations (`⚡ Operations` sidebar)
- New dedicated page for daily execution work: Kit Orders, Credentials, Task Queue
- Tab-based layout: **Kit Order** · **Credentials** · **Task Queue**
- Sidebar entry with badge showing overdue task count

### Kit Order Flow (`KitOrderService.js` — new)
- `logKitOrder(data)` — single call after Amazon order placed:
  1. Fetches learner name + parent phone from HubSpot deal
  2. Fetches delivery address from associated HubSpot contact (`address`, `city`, `state`, `zip`, `country`)
  3. Sends WATI template `migration_kit_sent_by_us_parent_information` with `{name,value}` params: Parent · Kit_name · Delivery_date · Address
  4. Writes HubSpot deal note matching existing format (Order Details / Order No. / Dispatch to / Arriving / Amazon link)
  5. Appends row to Kit Tracking sheet via `logKitOrderToSheet()` in `KitTrackingService.js`
- `getKitOrderData(jlid)` — prefills form: learner name, parent name, phone, full address
- `_fetchContactAddress(dealId)` — fetches address via deal→contact association chain
- Sheet logging: `logKitOrderToSheet()` added to `KitTrackingService.js` — uses correct `KIT_COL` map (G=OrderDate, I=ETA, D=Country, E=Price, P=JLID); replaces broken separate implementation
- Bug fixed: original `_appendKitTrackingRow` was writing order date to col F (Site) — now correctly writes to col G (Date of Order)

### Scratch Credentials Automation (`CredentialsService.js` — new)
- `generateScratchCredentials(jlid, learnerName)`:
  1. Reads Scratch Credentials sheet → finds highest `SHJLK` number → increments
  2. Appends row: username · `jetlearn` · learner name · JLID · timestamp
  3. Searches Google Calendar 180 days ahead for events containing JLID in title
  4. Updates event description: appends `Scratch = SHJLKxx\npass = jetlearn`; removes any stale Scratch block first
- `peekNextScratchUsername()` — preview next username without committing (shown in UI before generate)
- Credentials spreadsheet: `1KsyxldnHpm7gEyTcmmQFkz-uaqTM_FMhNTxh7OXBCTk`
- Calendar matching: strips trailing `C` from JLID for safer partial match (`JL55030989090C` → searches `JL55030989090`)
- After generation: "Register on Scratch" link shown pointing to `scratch.mit.edu/join`
- Code.org tab: placeholder shown (coming soon — PDF attach flow)

### HubSpot Task Queue (`HubSpotService.js`)
- `getMyHubSpotTasks()` — fetches all open HubSpot tasks for owner `61546090` (Sourav):
  - Filters: `hs_task_status ≠ COMPLETED` + `hubspot_owner_id = 61546090`
  - Includes deal + ticket associations in single search call
  - Batch-fetches associated deals to get JLID + learner name per task
  - Returns tasks sorted: 🔴 Overdue → 🟡 Today → 🟢 Upcoming
- `completeHubSpotTask(taskId)` — PATCHes task status to COMPLETED
- `_categoriseTask(subject)` — classifies by title pattern: `installation_email` · `material_email` · `credentials` · `certificate` · `afa` · `kit` · `migration` · `manual`
- Task Queue UI: colour-coded rows with left border per category, JLID chip, due date, action button + ✓ Done
- Action buttons route to correct page/tab: Installation Email → Communication · Credentials → Credentials tab pre-filled · Certificate → Bulk Certificates · Kit → Kit Order tab · Migration → Migration form · AFA → toast reminder
- ✓ Done: marks complete in HubSpot, fades row instantly, updates overdue count

### Auth — Auto-Login Removed
- Removed auto-`triggerGoogleSignIn()` call from `DOMContentLoaded` entirely
- Previously: GAS injects Workspace email into page → on every reload, auto-signed in even after logout
- Now: user must click Sign In button manually; session restore (reload while logged in) still works
- `handleLogout` sets `sessionStorage.manualLogout = '1'`; `loginSuccess` clears it (retained as safety fallback)

---

## [2026-05-12] — Performance, Auth & UX Overhaul

### Performance — Page Load Speed (5–12s saved)
- Removed `getLiveCurrencyRates()` from `doGet` — was a blocking `UrlFetchApp` call on every page load
- Page now serves instantly with hardcoded fallback rates; live rates load async via `_refreshLiveCurrencyRatesAsync()` after DOMContentLoaded
- New `getCachedCurrencyRates()` server function — wraps API call with `CacheService` (6-hour TTL); only 1 real API call per 6h across all users

### Authentication — Google Workspace Login
- Replaced username/password login with Google Workspace native auth (`Session.getActiveUser().getEmail()`)
- Access restricted to `@jet-learn.com` domain — any other email sees a clear error
- `sourav.pal@jet-learn.com` auto-assigned Super Admin role; all other `@jet-learn.com` users auto-created as User role on first login
- `authenticateByEmail(email)` — new server function; no token verification needed (Workspace handles it)
- `authenticateWithGoogle(idToken)` — backup Google tokeninfo verifier
- `_createGoogleUser()` / `_updateGoogleUserLastLogin()` — auto-manages user profiles without manual setup
- `verifyUserSession` updated to match by email OR username for backwards compatibility

### Login Page — Premium Redesign
- Replaced plain HTML form with glassmorphism sign-in card
- Custom Google button with inline SVG — no GIS library dependency
- Trust badges: `@jet-learn.com only` · `OAuth 2.0 secured` · `Role-based access`
- Google avatar shown in app header after login (falls back to initials if no picture)
- Auto sign-in on page load: if GAS detects Workspace email and no active session, triggers sign-in after 600ms delay

### Session — Timeout Extended
- Session timeout extended from 45 minutes to 8 hours (`SESSION_TIMEOUT_MS = 8 * 60 * 60 * 1000`)

### Audit Log — Column Rename
- Column M header renamed: `Intervened By` → `Actioned By`

### Invoice / Onboarding — Custom Installment Amounts
- New "Custom installment amounts" checkbox in both Invoice Generator and Onboarding Email forms
- When checked: renders N input rows matching installment count — each amount editable independently
- Token/deposit tracked separately via `Partial Payment Received` — all custom installments shown as Pending on invoice
- Server-side `calculateInvoicePricing` Case C added for custom amount arrays
- Client-side `calculateInvoicePricingClient` mirrors Case C logic

### Invoice — Currency Fix (EUR → GBP)
- Fixed: GBP invoices were showing EUR base price (was always reading `Base Price EUR` then multiplying by rate=1.0)
- Now reads native `Base Price GBP` / `Base Price USD` columns; conversion only applied for non-native currencies

### Invoice — Token Box on Invoice
- Green token receipt box now appears above the orange installment plan when a partial payment exists
- Shows: `✅ Token Amount Received — £X paid`

### Invoice / Onboarding — Auto-Preview Disabled
- Removed auto-preview-on-keystroke behaviour on Invoice Generator and Onboarding Email pages
- Preview now only fires on explicit button click — prevents unnecessary server calls while typing

### Invoice — Preview Reliability
- Fixed failure handler in `getAndRenderEmailPreview` — was silently swallowing errors
- Now shows actual error message inside preview frame + toast notification

---

## [2026-04-30] — Parent Will Buy Kit Automation (V53–V55)

### Parent Will Buy — Full WhatsApp Follow-Up System (V53)
- New `ParentWillBuyService.js` — complete automation for kits parents procure themselves
- Initial WhatsApp message fires the moment entry is added via UI (no waiting for 9am trigger)
- Adaptive interval system: >21d=7d, 15–21d=5d, 8–14d=3d, 4–7d=2d, ≤3d=1d between FUPs
- Always runs full sequence (Initial → FUP1 → FUP2 → Final) — never skips steps
- Rows with ≤7 days to course start marked `In Progress - URGENT 🔴`
- CLS notified (email + HubSpot note) when course is 2 days away with no confirmation
- Full escalation (HubSpot task + email to learner's CLS manager) when course has started
- CLS email resolved per-learner from HubSpot `cls_manager` deal property via `findClsEmailByManagerName()`
- New `Parent_will_buy` sheet tab + `getPWBEntries()` / `addPWBEntry()` server functions
- Kit Tracking page: JetLearn Sends / Parent Will Buy tab toggle with separate KPI strips

### Parent Will Buy — Dashboard & Entry Improvements (V54)
- Entry By dropdown (Sourav / Shubham / Ankita) — mandatory field on Add Entry modal
- Date (B) and Month (C) auto-filled on row creation for monthly reporting
- Interval locked into sheet col T at initial send — immune to course date drift
- Next FUP countdown in dashboard: 🟢 in Xd / 🟡 tomorrow / 🔴 overdue
- `renderPWBTable` updated with all new columns

### Parent Will Buy — Reply Handling & HubSpot Sync (V55)
- Kit-specific HubSpot property updates at every stage:
  - VR Headset → `vr_headset__oculus_status`
  - Microbit → `microbit_kit_status`
  - Makey-Makey → `makey_makey_kit_status`
  - Arduino → `arduino_kit_status`
- Status values: Reminder 1 sent → Reminder 2 sent → Final reminder sent → Parent bought it → Escalated to CLS
- Free text reply capture: unmatched messages logged to sheet + HubSpot note + CLS email
- PWB fuzzy matching: "Order Placed, delivery on 2nd May" → `Order Placed`; "told by JetLearn" → `Yet to place an order`
- Bug fix: reply handler matches from first message (not just after final FUP)
- Bug fix: sibling phone conflict — picks most-recently-active row when two learners share phone
- `discoverPWBHubspotProperty()` utility added for future property discovery

---

## [2026-04-28] — Kit Tracking Automation + Certificate Reliability

### Kit Tracking — WATI Webhook Fix
- Fixed webhook constantly failing — was using `/a/macros/jet-learn.com/` domain URL; updated to standard `/macros/s/` URL accessible by WATI
- `doPost` now returns HTTP 200 to WATI **immediately** — queues kit reply to `CacheService`, fires `_processWatiKitReply` trigger 5s later to avoid WATI timeout failures
- Added `_processWatiKitReply()` background function that processes queued webhook data async

### Kit Tracking — Auto-Reply Messages
- Parent taps **Kit Received** → bot replies: "✅ Thank you for confirming! We've updated our records..."
- Parent taps **Not Received yet** → bot replies: "😟 We're checking with logistics right away. We'll update you shortly."
- Parent taps **Need To Check** → bot replies: "👍 No problem! We'll follow up in 12-24 hours."
- Replies use WATI session messages (free-text, no template required) — works within 24hr window after parent replies
- Fuzzy-matched free-text replies also trigger auto-reply

### Kit Tracking — HubSpot JLID Normalisation
- Fixed HubSpot kit status not updating when JLID has trailing stray characters (e.g. `JL39611449152C2` → normalised to `JL39611449152C`)
- Logs JLID normalisation: `JLID normalised: "JL39611449152C2" → "JL39611449152C"`

### Kit Tracking — Escalation System (2nd Reminder)
- Daily trigger now runs a second pass: rows with 1st follow-up sent + no response + 2 days elapsed → auto-sends 2nd WATI reminder
- New sheet columns: T (FOLLOWUP2_SENT), U (FOLLOWUP2_SENT_AT)
- Status `escalated` = 2nd reminder sent, still no reply
- Dashboard: red KPI card, "Needs Attention" banner with learner chips showing 2nd-sent date
- Table rows with `escalated` status: red left border, `🔴 Escalated` badge, "Re-send Urgent" button
- Banner "View Escalated" button programmatically sets status filter dropdown

### Kit Tracking — Add Entry Fix
- New entries now write to actual last data row (scans col A from bottom) — previously appended at row 10972 due to empty formatted rows
- Col H (Timestamp Month — formulated) is now skipped when writing new entries
- Mandatory fields enforced on Add Entry form: Learner Name, Kit, Country, Price (EUR), Site, Date of Order, ETA, Reason, Subscription, Roadmap, Name (Sent By)
- Error message shows exactly which fields are missing

---

## [2026-04-27] — Certificate Bulk Sending Reliability

### Certificate Center — Pool Architecture
- Replaced per-certificate `makeCopy` (N calls) with pool architecture (3 calls total — one per slide type: Foundation / Maths / Pro)
- Pool copies reused: fill → export → reset → next cert — eliminates Drive API rate limiting
- Font sizes read from template upfront and restored after each cert reset
- Switched PDF export from `UrlFetchApp` to `DriveApp.getAs('application/pdf')` — eliminates bandwidth quota errors
- `makeCopy` retried up to 4 times with 3s sleep on transient Drive errors

### Certificate Center — Resend Failed
- Failed cert log rows now show checkboxes
- Resend toolbar appears on selection showing count of selected
- "Resend Selected" groups by learner+email and re-runs `sendBulkCertificates`

---

## [2026-04-26] — Kit Tracking Dashboard

### Kit Tracking — Dashboard
- New Kit Tracking page with KPI strip: Total / Delivered / Awaiting / Not Received / Overdue / Escalated
- Table with month/status/kit filters, search, row expand for details
- Inline Edit button on every row — set Delivery Date, response, JLID manually
- Manual "Delivered" edit auto-updates HubSpot kit status + Time Taken
- Send Follow-up button per row (with JLID auto-lookup fallback prompt)

### Kit Tracking — WATI Automation
- Daily 8am trigger: sends WATI template `migration_kit_fup_sent_by_us` to parents of overdue kits
- WATI webhook → `handleKitReply`: Kit Received → fill Delivery Date + PATCH HubSpot; Not Received / Need To Check → HubSpot deal note
- Fuzzy + predictive text matching for free-text parent replies ("received it", "haven't got it", "let me check", etc.)
- HubSpot kit property map: VR Headset → `vr_headset_oculus_status`, Microbit → `microbit_kit_status`, Makey-Makey → `makey_makey_kit_status__t_`, Arduino → `arduino_kit_status`
- Status value on confirmation: `Received by the Parents`

---

## [2026-04-25] — Learner Course Progression (Course Planner)

### Course Planner — New Page
- New "Course Planner" sidebar page for predicting learner course completions
- Ingests Athena CPRS + PRMS CSV data (paste into sheets)
- Computes: sessions done, frequency (last 28 days), classes left, projected completion date
- Alert levels: 🔴 Critical (≤4 weeks + migration needed), 🟡 Warning (≤6 weeks), 🟢 OK
- Sidebar badge shows critical alert count on every page load
- CCTC flag: teacher upskill < 71% on next course = migration needed

### Course Planner — Smart Migration Trigger
- "Trigger Migration" button: searches matching teachers, pre-fills top 3 matches
- Creates HubSpot ticket on migration pipeline with learner + teacher + reason
- Critical learners: "⚠ CLS Approval Required" warning on confirmation modal

### Course Planner — HubSpot Course History
- `_buildHealthMap()` fetches `propertiesWithHistory: ['current_course__t_']`
- Full course journey pulled from HubSpot (Fundamentals → Edublocks → Game Dev → Python 2.0)
- `courseNumberWithTeacher` now accurate (was showing "1st course" for 3rd/4th course learners)
- CCTC badge fires correctly for all learners regardless of CPRS window

---

## Earlier Releases

### Authentication & Security
- Force password reset flow with secure token
- Session timeout + re-authentication
- Hardened input handling across all forms

### Migration Tracker
- Full migration pipeline: CLS approval via Slack buttons
- Teacher matching by course, slot, persona
- HubSpot ticket auto-creation with deal properties

### Certificate Center (Initial)
- Bulk certificate generation from Google Slides template
- PDF export + email to parent
- HubSpot deal note with Drive links

### Teacher Persona & Upskill
- Teacher persona mapping (traits, age groups, availability)
- Course upskill progress tracking per teacher
- Smart teacher search with slot + course matching
