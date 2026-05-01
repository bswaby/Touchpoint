### ✅ [Attendance Markings (Reverse Attendance)](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Attendance%20Markings)
Team-friendly attendance entry for large events. The trick: when most people show up, default everyone to **Present** and only click the few who didn't (or flip the default for sparsely-attended events). Built for VBS, camps, conferences where anywhere clicking 2,000 people one-by-one would take forever. Multiple staff can chip away at the involvement list together with live updates every 10 seconds.

- ⚙️ **Implementation Level:** Easy
- 🧩 **Installation:** Single script. Paste into Special Content > Python, navigate to `/PyScript/TPxi_AttendanceMarkings`. Configure your first session from the UI.

<summary><strong>Features</strong></summary>

- **Reverse Attendance:** When most people attend, default everyone to **Present** and click only the absentees. 2,400 enrolled · 2,000 attended = 400 clicks instead of 2,000
- **Three Default Modes** (per config): Default Present (click absent), Default Absent (click present, for sparsely-attended events), or Default Unmarked (must explicitly mark each)
- **Team Workflow:** Multiple staff work the involvement list together. Soft "claim" indicators show who's working what (e.g., *"Working: Mary (2m ago)"*). Dashboard auto-refreshes every 10 seconds
- **Down-To-Zero Counter:** Live "Remaining: X / Y · Z% done" tracker with clickable filter pills.  All / Not Started / In Progress / Done. Click any pill to filter the list
- **Walk-In People:** Search for an existing person (smart multi-token name match.  "Be Swab" finds Ben Swaby) or create a brand-new person right from the roster screen. Auto-marks them present and adds them to the involvement with the configured member type (Member / Visitor / Prospect / Guest)
- **Per-Click DB Writes:** Each toggle writes immediately via `model.EditPersonAttendance` so teammates see updates on the next 10-second poll. No "save" button needed for individual marks
- **Finalize:** One click commits the default state to all unmarked people and locks the involvement as Done. Re-open is supported in case of mistakes
- **Smart Source Selection:** Pick by Program / Division (auto-includes all matching active involvements via `OrganizationStructure` so multi-program orgs surface correctly) or by Specific Involvements (manual picker)
- **Schedule-Aware Filtering:** Optional "Only involvements with a meeting on the selected day" toggle. Honors:
  - Legacy `OrgSchedule` recurring weekly schedule
  - `Organizations.FirstMeetingDate` / `LastMeetingDate` window
  - New TouchPoint **Scheduler / MeetingSeries** (with RRULE expansion) — works for churches that have migrated to the new model OR are running a mix of both
- **Headcount Link:** Quick "Headcount in TouchPoint ↗" button on the roster opens the meeting page so leaders can enter `HeadCount` / `NumNewVisit` / etc. via TouchPoint's native UI (TouchPoint's Python API doesn't expose those columns)
- **Configurable Leader Exclusion:** Comma-separated list of `MemberTypeId` values to exclude from the roster (e.g., `140,310,320` to skip leaders so the count reflects participants only)
- **P / A / U Status Tags:** Each person row shows a colored tag. **P** Present (green), **A** Absent (red), **U** Unmarked (grey). Hover for tooltip with full meaning
- **Audit Log:** Every mark, claim, finalize, walk-in add, and reopen is logged per session for after-the-fact review
- **Auto-Update:** Built-in version check (admin-only). Banner appears on the landing page when a new version is available; one click pulls the latest from DisplayCache and writes it back to your TouchPoint script

<hr>

<summary><strong>Pick a Saved Config</strong></summary>
<p>Each config card shows source summary, default state, and last updated. Pick one to start a session, edit one, or build a new one from scratch.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Attendance%20Markings/AM-SelectSavedConfig.png" width="700">
</p>

<summary><strong>Configure Once</strong></summary>
<p>Define source (Program/Division or Specific Involvements), default state, schedule filter, walk-in member type, and leader exclusions. Save and reuse for the whole VBS week or every Sunday.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Attendance%20Markings/AM-Settings.png" width="700">
</p>

<summary><strong>Start a Session</strong></summary>
<p>Pick the date you're taking attendance for, enter your name (so teammates see who's working what), and start. If your config doesn't follow involvement schedules, you'll also enter a global meeting time used for any involvement without a scheduled time that day.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Attendance%20Markings/AM-SetDateandName.png" width="700">
</p>

<summary><strong>Live Involvement Dashboard</strong></summary>
<p>The list splits into Not Started, In Progress, and Done. Click any pill at the top to filter. Each row shows live counts (P / A / U), who's currently working it, and the last-touch timestamp. Auto-refreshes every 10 seconds so the team stays in sync.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Attendance%20Markings/AM-InvDoneNotDone.png" width="700">
</p>

<summary><strong>Mark Attendance</strong></summary>
<p>Each person row defaults to the config's default state (e.g., Present for VBS). Click a row to toggle. Big rooms work fast.  The only people you click are the ones who don't match the default. Search by name, add walk-ins on the fly, or jump to TouchPoint's headcount entry.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Attendance%20Markings/AM-MarkingsSheet.png" width="700">
</p>

<summary><strong>Finalize When Done</strong></summary>
<p>The Finalize button tells you exactly how many unmarked people will become the default state when you click it (e.g., "Finalize (47 will become Present)"). One click commits, marks the involvement Done, and bumps the dashboard's Remaining counter down.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Attendance%20Markings/AM-FinalizeButton.png" width="700">
</p>

<summary><strong>Verify in TouchPoint</strong></summary>
<p>All marks land in TouchPoint's `Attend` table immediately, viewable in the org's standard Attendance tab. Headcount fields (`HeadCount`, `NumNewVisit`, etc.) are entered via the native TouchPoint meeting page using the link on the roster screen.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Attendance%20Markings/AM-Results.png" width="700">
</p>

<summary><strong>Installation</strong></summary>

1. **Admin > Advanced > Special Content > Python**
2. Click **Add New**, name it `TPxi_AttendanceMarkings`
3. Paste the script and Save
4. Navigate to `/PyScript/TPxi_AttendanceMarkings`
5. Click **+ New Config** to build your first session — pick source, default state, and walk-in member type, then save
6. From the landing page, pick the config and choose a date to begin taking attendance

<summary><strong>Config Options</strong></summary>

| Setting | Description |
|---------|-------------|
| **Source Type** | Program / Division (auto-includes matching active involvements) or Specific Involvements (manual picker) |
| **Program / Division Groups** | Stack multiple combos to scope the session across several areas in one go |
| **Default Attendance State** | Present (most attended), Absent (sparsely attended), or Unmarked (explicit-only). The "click less" choice depends on your event |
| **Only With Meeting** | When ON, only includes involvements with a meeting scheduled on the selected day (legacy `OrgSchedule` + `[FirstMeetingDate, LastMeetingDate]` window OR new TouchPoint Scheduler `MeetingSeries`) |
| **Exclude Member Type IDs** | Comma-separated list (e.g., `140,310,320` for leaders) to filter out of the roster |
| **Walk-Ins Allowed** | Show or hide the "+ Add Person" button on the roster |
| **Add Walk-Ins As** | Member / Visitor / Prospect / Guest / Inactive Member.  Applied automatically when adding a walk-in |
| **Exclude Involvement IDs** | Optional list to skip specific orgs when source is Program/Division |

<summary><strong>How "Reverse Attendance" Saves Time</strong></summary>

| Event | Enrolled | Attended | Old Way (clicks) | Reverse Way (clicks) | Time Saved |
|-------|----------|----------|------------------|----------------------|------------|
| VBS Day 1 | 2,400 | 2,000 | 2,000 | 400 | 80% |
| Sunday Morning Kids | 250 | 220 | 220 | 30 | 86% |
| Workshop | 200 | 30 | 30 | 170 | (use Default Absent — back to 30 clicks) |
| Wednesday Youth | 150 | 75 | 75 | 75 | 0% (default doesn't help when ~half attend; use either mode) |

The bigger the gap between *enrolled* and *attended* in either direction, the bigger the win. Pick the default that matches the gap and you click less.

<summary><strong>Team Workflow Tips</strong></summary>

- **One staff per involvement at a time.** The dashboard shows who's actively working a row ("Working: Mary 2m ago"). Soft signal — anyone can still open it, but the indicator helps you avoid stepping on toes
- **Filter to Not Started** when joining mid-session it gives you the list of involvements no one has touched yet
- **Filter to In Progress** to find rows where someone started but didn't finalize to helps clean up at the end of the day
- **Walk-ins land in the roster instantly.** If you're adding a guest, search first (smart multi-token: "Be Swab" finds Ben Swaby, "Swa, Be" works too); only create a brand-new person if no match is found
- **Finalize promptly** — until you click Finalize, the involvement stays in In Progress on the dashboard. The Finalize button preview tells you exactly how many people will get the default state, so no surprises

<summary><strong>Tips</strong></summary>

- **Default Present is right for most church events.** Most people show up; click the few who don't. For event-specific situations like a workshop where most are absent, use Default Absent for the same time-saving math in the opposite direction
- **Use schedule filtering for VBS.** Set "Only With Meeting" ON so the dashboard only shows the involvements actually meeting that day. Off-days disappear automatically; June dates show; April dates don't
- **Multi-program orgs work correctly.** Some involvements roll up under multiple Programs / Divisions (e.g., a VBS room that lives under both "VBS K-5" and "VBS DAILY GRAND TOTAL"). The picker uses `OrganizationStructure` so it surfaces under any program assignment, not just the primary
- **Per-click writes are atomic.** Two staff can click the same person within the same second.  Last write wins, and the dashboard's 10-second poll picks up whichever was last. Rare in practice; cleanup is one click if it happens
- **Headcount lives in TouchPoint.** TouchPoint's Python API doesn't expose setters for `HeadCount` / `NumNewVisit` / etc. The roster has a "Headcount in TouchPoint ↗" button that opens the meeting page so leaders enter the number in the official spot

---
*Written by [Ben Swaby](https://github.com/bswaby). These tools are free because they should be. If they've saved you time, consider [DisplayCache](https://displaycache.com) Church digital signage that integrates with TouchPoint or [TPxi Go](https://tpxigo.com) your church contacts in Outlook and on your phone.*
