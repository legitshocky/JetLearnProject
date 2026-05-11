# JetLearn Platform — Changelog

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
