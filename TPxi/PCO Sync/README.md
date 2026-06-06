### 🔄 [PCO Sync (Planning Center Online → TouchPoint)](https://github.com/bswaby/Touchpoint/tree/main/TPxi/PCO%20Sync)
One-way sync from Planning Center Online into TouchPoint: people, rosters, teams, and per-plan attendance. Worship admins schedule and take attendance in PCO Services; TouchPoint is the authoritative people database. This bridges the two so staff never have to double-enter.  Pull a service plan, match its attendees to TP people once (the link is saved as an Extra Value), and write attendance, roster, and subgroup memberships back into the corresponding TP involvement in one click. Or schedule it to run automatically and email you the summary.

- ⚙️ **Implementation Level:** Easy–Moderate
- 🧩 **Installation:** Single script. Paste into Special Content > Python, navigate to `/PyScriptForm/TPxi_PCOSync`, paste your PCO Personal Access Token, and start mapping.

<summary><strong>Features</strong></summary>

- **Three Sync Modes** — pick what fits each PCO concept:
  - **All People Sync** — singleton mapping. Walk the entire PCO People directory and reflect every matched record into one TP "PCO Directory" involvement
  - **Service Type Sync** — one PCO Service Type (e.g., "11:00 Worship Center") → one umbrella TP involvement. Optional layers: teams-as-subgroups, per-plan attendance writes
  - **Team Sync** — one PCO Team (e.g., "Band" under Wilson Hall Service) → one TP involvement. Optional layers: positions-as-subgroups, per-plan attendance
- **PCO is Source of Truth (mirror behavior):** roster sync adds AND removes. TP members whose `PCO_PersonId` is no longer in scope get removed from the involvement on the next sync. Subgroup memberships matching a current PCO position/team but no longer held also get dropped. Manually-added members (no PCO link) and unrelated subgroups are left alone — your hand-curated data stays untouched
- **Person Matching at Scale:**
  - **Proposed Matches** — scores TP candidates for every unmatched PCO record using name + email + birthdate signals. Tiered Strong / Medium / Weak. Per-row Apply or Skip Forever, bulk Apply for high-confidence tier, scoped (from a preview) or full-directory walk. Client-side cached so tier and search changes don't re-walk PCO
  - **Verify Person Link** — search a TP person, see their stored PCO link side-by-side with PCO's record (with red cells where the data disagrees and a one-line verdict). Unlink or Replace with another PCO person, all from one panel
- **Preview Before Sync:** every sync opens a preview modal showing match counts, sync-mode banner, mirror-removal banner (with red "will be removed" pill listing how many TP members + stale subgroups will drop), and per-attendee rows. Manual search-and-link for anything unmatched
- **Confirm Spells Out Every Action:** the Sync Now dialog lists adds, drops, subgroup writes, subgroup drops, and attendance writes by name before any DB write. No surprises
- **Scheduled Sync (with Email Summary):**
  - One-click install adds a managed block to TouchPoint's `ScheduledTasks` special content (matches ProspectBuilder's pattern). Per-mapping schedule editor is gated on global install — both client- and server-side
  - Per-mapping: Daily or Weekly, day-of-week + hour, notify a TouchPoint user (typeahead picker by name / username / email), include-issues toggle
  - Scheduler runner walks all mappings every invocation, fires anything whose day/hour match now and hasn't run this hour. Each fire syncs fully server-side and emails the configured user: summary counts (joined, already, subgroup writes, members removed, stale subgroups removed), optional issues list (unmatched, ambiguous matches, PCO API warnings), and a link back to PCO Sync
- **Person Data Sync (per-field, opt-in):** for each TP field (FirstName, LastName, EmailAddress, CellPhone, etc.) pick the direction (PCO → TP) and behavior (off, auto-apply, or queue-for-review). Defaults off. TouchPoint stays authoritative until you opt in. Queued changes appear in a review queue with side-by-side diff
- **Verify-After-Write:** every settings save, mapping save, and scheduler install reads back from storage and confirms the change persisted. Silent permission failures surface immediately with a clear error message
- **Diagnostics On Every Mapping:**
  - Team Mappings have a "Check PCO positions" button that walks PCO and reports exactly what's there. `5 position(s) [Lead Vocal, Backup, ...], 12 assignment(s) across 8 people. Subgroups will sync.` Or the matching red state when positions exist but nobody's assigned in PCO
  - Dashboard health panel surfaces broken mappings (deleted PCO resource, archived TP org) before staff hit Sync
  - Audit log in `PCOSync_Log_YYYYMM` captures per-sync counters (joined, dropped, subgroup adds/drops, failures, scheduler runs, mapping edits, link write/unlink, email send/fail)
- **Last-Sync + Next-Run Pills:** dashboard cards show `Synced 3h ago` (green) and, when scheduled, `Next: Sun 6:00 AM` (blue) so you always know the state at a glance

<hr>

<summary><strong>Sync Dashboard</strong></summary>
<p>One landing page for every mapping. All People at the top (when configured), then per-plan cards under day headers (auto-scrolls to today like the PCO mobile app), then Service Type and Team mapping cards. Each card shows last-synced and next-run pills, plus a one-click Preview & Sync button.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/PCO%20Sync/PCO-Dashboard.png" width="700">
</p>

<summary><strong>Settings & PCO Connection</strong></summary>
<p>Paste your PCO App ID + Secret (Personal Access Token), Test Connection, then enable per-field Person Data Sync rules and install the global Scheduled Sync block. Verify-after-write on every save catches silent permission failures.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/PCO%20Sync/PCO-Settings.png" width="700">
</p>

<summary><strong>Sync Mappings — One Place, Three Types</strong></summary>
<p>All People (singleton), Service Type Mappings (one PCO Service Type → one umbrella TP involvement, optional team subgroups + per-plan attendance), and Team Mappings (one PCO Team → one TP involvement, optional position subgroups + per-plan attendance). Per-row toggles, schedule editor, and inline diagnostics ("Check PCO positions") on every row.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/PCO%20Sync/PCO-SyncMapping.png" width="700">
</p>

<summary><strong>People Matching</strong></summary>
<p><strong>Proposed Matches</strong> scores TP candidates for every unmatched PCO record (name + email + birthdate signals) with Strong / Medium / Weak tiers and bulk Apply for high-confidence hits. <strong>Verify Person Link</strong> lets staff inspect any existing TP↔PCO link side-by-side with a one-line verdict — Unlink or Replace from one panel. Pending Data Reviews collects field-diff changes flagged by your Person Data Sync rules.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/PCO%20Sync/PCO-PeopleMatching.png" width="700">
</p>

<summary><strong>Scheduled-Sync Email</strong></summary>
<p>Every scheduled run emails the configured TouchPoint user a clean summary: joined, already member, subgroup writes, members removed (mirror), stale subgroups removed. Optionally includes the issues list.  Unmatched PCO records, ambiguous email matches, PCO API warnings, plus a deep link back to PCO Sync to act on them.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/PCO%20Sync/PCO-Email.png" width="700">
</p>

<summary><strong>Installation</strong></summary>

1. **Admin > Advanced > Special Content > Python**
2. Click **Add New**, name it `TPxi_PCOSync`
3. Paste the script and Save
4. Navigate to `/PyScriptForm/TPxi_PCOSync`
5. Open the **Settings** tab, paste your PCO **App ID + Secret** (generate one in PCO under *My Account → Applications → Personal Access Tokens*), Save, then click **Test Connection**
6. (Optional) Open **Settings → Scheduled Sync** and click **Install** to auto-add the runner to `ScheduledTasks`
7. Switch to **Sync Mappings**, click **+ Add** under any of the three sections (All People / Service Type / Team), pick the PCO resource and the TP involvement, save
8. Hit **Preview & Sync** on the Dashboard card to do your first run

<summary><strong>Sync Mode Cheat Sheet</strong></summary>

| Mode | Best For | Subgroups | Attendance |
|------|----------|-----------|------------|
| **All People** | Mirror every PCO record into a "PCO Directory" involvement so everyone has a TP shell | n/a | n/a |
| **Service Type** | One TP umbrella involvement for a whole worship service (Worship Center, Wilson Hall, etc.). Teams under that service type become subgroups | Optional (Teams as subgroups) | Optional (per-plan, marks Confirmed as Present) |
| **Team** | Each PCO team (Band, Production, Welcome, etc.) lives in its own TP involvement. Positions become subgroups | Optional (Positions as subgroups) | Optional (per-plan) |

You can use any combination. The Band team can have a dedicated Team mapping AND be reflected as a subgroup under the Service Type umbrella.

<summary><strong>How Mirror Removal Works</strong></summary>

**PCO is the source of truth for who is in the involvement.** On every sync:

- **Adds:** every PCO person matched to a TP person gets `JoinOrg`'d if not already on the roster. New position/team assignments get `AddSubGroup`'d
- **Removes:** every TP member whose `PCO_PersonId` extra value is no longer in PCO's scope (no longer on the team / service type / directory) gets `RemoveFromOrg`'d. Every TP subgroup membership whose name matches a current PCO position/team but the person no longer holds gets `RemoveSubGroup`'d
- **Untouched:** TP members without a `PCO_PersonId` (manually added) stay. Subgroups whose name doesn't match any current PCO position/team stay (e.g., a manually-added "Pyrotechnics" subgroup is never touched)

The preview modal shows a red **Mirror removal** banner with counts before you hit Sync, and the confirm dialog spells out every drop by category. Once TouchPoint-side write-back is implemented in a future version, this strict-mirror default will become a per-mapping toggle.

<summary><strong>Tips</strong></summary>

- **Match the long-tail with Proposed Matches first.** Open it from any preview to scope to that preview's unmatched (faster). For ongoing maintenance, the unscoped walk processes the whole PCO directory in one shot. Bulk Apply the Strong tier — those are confident matches
- **Use Verify Person Link when something looks off.** "Why is Alice's email wrong?" → search Alice → see her PCO record side-by-side → if red cells confirm it's the wrong PCO person, Unlink or Replace. Faster than digging through PCO and TP separately
- **Schedule different mappings at different times.** Worship teams sync Sunday at 5 AM (catches Saturday rehearsal changes). All People can run nightly. Each mapping has its own day/time, so you don't have to compromise on one schedule
- **Check PCO Positions diagnostic** is the fastest way to tell apart "subgroups aren't syncing because of a bug" vs "PCO has no positions assigned." Run it before opening a support thread. Most "missing subgroup" reports are PCO setup gaps
- **Person Data Sync defaults off for a reason.** TouchPoint stays authoritative on person fields unless you opt in. When you do, prefer queue-for-review over auto-apply for the first few weeks so you can spot any PCO-side dirty data before it lands in TP
- **Email backlinks open the right tab.** The scheduled-sync email's "Open PCO Sync" link drops you on the Dashboard. Click any failed mapping's card to see why

<summary><strong>Storage & Schema</strong></summary>

| Storage | Purpose |
|---------|---------|
| `PCOSync_Settings` | PCO App ID, Secret, last-sync timestamps |
| `PCOSync_AllPeopleMapping` | Singleton All People mapping + schedule |
| `PCOSync_PeopleMappings` | Service Type mappings keyed by PCO Service Type ID |
| `PCOSync_TeamMappings` | Team mappings keyed by PCO Team ID |
| `PCOSync_PersonRules` | Per-field Person Data Sync rules |
| `PCOSync_Log_YYYYMM` | Monthly audit log of every write, scheduled run, and email outcome |
| `ScheduledTasks` | Managed block (between markers) that calls back into the script for scheduled runs |
| `PCO_PersonId` (Person Extra Value) | The canonical TP ↔ PCO link, per person |

---
*Written by [Ben Swaby](https://github.com/bswaby). These tools are free because they should be. If they've saved you time, consider [DisplayCache](https://displaycache.com) — church digital signage that integrates with TouchPoint — or [TPxi Go](https://tpxigo.com) — your church contacts in Outlook and on your phone.*
