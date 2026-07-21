# JetLearn Platform — Changelog

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
