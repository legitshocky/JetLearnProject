# JetLearn Platform — Changelog

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
