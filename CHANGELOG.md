# JetLearn Command Center — Changelog

---

## v9.51 — 2026-09-21

### Migration Automation

- **Class booking fix**: Automation now fetches the existing calendar event for a learner and extracts the class link (Zoom or GMeet) from its description. Previously relied solely on `zoom_masked_link` from the HubSpot contact, which caused booking to be skipped for GMeet learners.
- **Upcoming-event priority**: `getExistingEventDescription` now searches upcoming events first (now → +60 days) before falling back to recent past events, ensuring the active class type is detected correctly.
- **GMeet support**: When the existing event is a GMeet class, the meeting code is extracted and `conferenceData` is set on new events — creates the "Join with Google Meet" button in Google Calendar.
- **Event title prefixing**:
  - All recurring events get the class-type prefix (e.g. `GMEET : Learner Name (JLID) : Jetlearn ... Lesson`)
  - First occurrence of the first session gets `Migration : ` prepended (e.g. `Migration : GMEET : Learner Name ...`)
- **Event description carry-over**: Existing event description (GMeet dial-in info, Zoom link, etc.) is copied into newly booked events.
- **CCTC course update**: On course-change migrations, automation updates the deal's `current_course` to the new course (future_course_1). Ticket course fields are left as-is.
- **CCTC certificate**: `sendCertificate` flag set automatically when migration reason contains "course change".
- **CET time conversion**: Ticket class time stored in CET format (e.g. "17:00") is converted to 12h AM/PM before passing to `bookClassesWithNewTeacher`.

### Teacher Intelligence Center (TIC)

- **Ticket stage movement**: After Smart Context Fetch, if the learner's ticket is in Migration Triggered stage, an action strip appears with three buttons:
  - ⚡ Execution Pending — teacher found, no CLS required
  - 🔄 CLS Pending — teacher found, CLS intervention required
  - 🔍 TP Pending — no teacher found
- **CLS auto-highlight**: CLS Pending button is highlighted when the migration reason matches any CLS-required reason.
- **Server function**: `moveTicketStageForJlid(jlid, targetStage)` added to `HubSpotService.js`.
- **Smart Context Fetch**: Now returns `ticketId` and `ticketStage` in contextData.

### Migration Page

- **⚡ Auto Migrate button**: After fetching a learner by JLID, an "Auto Migrate" button appears that runs full automation (emails + class booking + ticket update) for one ticket, with a progress popup showing each step.

### Learner Ops — Migration Activity Tab

- **Automation stats panel**: Shows count of auto-executed, queued, failed, and skipped tickets. Lists tickets needing manual review.

### General

- Version bumped to **9.51** (GAS @952).

---

## v9.35 and earlier

See git log for prior changes.
