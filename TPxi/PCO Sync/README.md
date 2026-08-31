### 🔄 [PCO Sync (Planning Center Online ↔ TouchPoint)](https://github.com/bswaby/Touchpoint/tree/main/TPxi/PCO%20Sync)
Sync between Planning Center Online and TouchPoint: people, rosters, teams, per-plan attendance, and selected person fields back to PCO. Worship admins schedule and take attendance in PCO Services; TouchPoint is the authoritative people database. This bridges the two so staff never have to double-enter. Pull a service plan, match its attendees to TP people once (the link is saved as an Extra Value), and write attendance, roster, and subgroup memberships into the corresponding TP involvement in one click. Or schedule it to run automatically and email you the summary.

**Who owns what.** Neither system is authoritative for everything, so it is worth being precise about which owns which:

| Data | Owner | Direction |
|------|-------|-----------|
| Involvement membership (rosters, teams, subgroups) | PCO | PCO → TouchPoint, mirrored (adds and removes) |
| Per-plan attendance | PCO | PCO → TouchPoint |
| Whether a person exists at all | TouchPoint | PCO never creates a TouchPoint person. An unmatched PCO record is reported, not added |
| Person field values (name, email, phone, address, background check) | You choose, per field | Off by default. Either direction, never both for the same field |

The short version: **PCO owns the roster, TouchPoint owns the people, and you decide field by field who owns each value.**

- ⚙️ **Implementation Level:** Easy to Moderate
- 🧩 **Installation:** Single script. Paste into Special Content > Python, navigate to `/PyScriptForm/TPxi_PCOSync`, paste your PCO Personal Access Token, and start mapping.

<summary><strong>Features</strong></summary>

- **Three Sync Modes** for the PCO → TouchPoint direction:
  - **All People Sync** is a singleton mapping. Walk the entire PCO People directory and reflect every matched record into one TP "PCO Directory" involvement
  - **Service Type Sync** maps one PCO Service Type (e.g., "11:00 Worship Center") to one umbrella TP involvement. Optional layers: teams-as-subgroups, per-plan attendance writes
  - **Team Sync** maps one PCO Team (e.g., "Band" under Wilson Hall Service) to one TP involvement. Optional layers: positions-as-subgroups, per-plan attendance
- **PCO owns the roster (mirror behavior):** roster sync adds AND removes. TP members whose `PCO_PersonId` is no longer in scope get removed from the involvement on the next sync. Subgroup memberships matching a current PCO position/team but no longer held also get dropped. Manually-added members (no PCO link) and unrelated subgroups are left alone, so your hand-curated data stays untouched
- **Removals refuse to run on bad data.** If the PCO read was incomplete, or PCO returned nobody while TouchPoint has linked members, nothing is removed and the reason is reported in the UI and in the scheduled-run email. A rate-limited or failed read can never be mistaken for "PCO has nobody"
- **Conflicting mappings are flagged.** Two mappings pointing at the same TP involvement will each remove the other's people on every run. The Mappings tab detects this and names the offenders
- **Person Matching at Scale:**
  - **Proposed Matches** scores TP candidates for every unmatched PCO record using name, email, and birthdate signals, sorted into Strong / Medium / Weak tiers. Per-row Apply or Skip Forever, bulk Apply for the high-confidence tier, scoped (from a preview) or full-directory walk. Client-side cached so tier and search changes don't re-walk PCO
  - **Manual TouchPoint search** on any proposed row for when the suggestion is wrong. Shows age and gender to separate twins, parent and child sharing a name, or a remarried surname, and warns if the person you picked is already linked to someone else in PCO
  - **Verify Person Link** lets you search a TP person and see their stored PCO link side-by-side with PCO's record, with red cells where the data disagrees and a one-line verdict. Unlink or Replace with another PCO person, all from one panel
- **Person Field Sync (per field, both directions, opt-in):** choose a direction and behavior for each field. Available fields: first name (goes by), last name, email, cell phone, home phone, birthdate, gender, household address, and background check. Everything defaults to off and TouchPoint stays authoritative until you opt in
  - **PCO → TouchPoint** runs as part of any mapping sync, manual or scheduled. Auto-apply or queue for review, with side-by-side diffs in the review queue
  - **TouchPoint → PCO** is manual and deliberate: Preview, tick the rows you want, Apply. Nothing writes on a schedule
- **Preview Before Any Write:** the outbound preview reads PCO and reports exactly what an apply would write, changes nothing, and shares one code path with the real run, so what you see is what runs. Rows where the two names are unrelated are flagged as a probable bad link, with Open and Unlink right on the row
- **Every Write Is Verified:** after each write the value is read back from PCO and compared. A field whose value keeps coming back different is marked stuck and stops being retried, so a formatting mismatch can never turn into an endless rewrite
- **Blanks Are Never Pushed:** an empty TouchPoint value means missing data, not an instruction to erase what PCO has
- **Background Check Sync (TouchPoint → PCO):** publishes the latest genuine outcome from `dbo.BackgroundChecks`, skipping abandoned submissions. Sends status, completion date, expiration, and the provider's report link. Reads the provider-agnostic `ApprovalStatus`, so it works whoever your provider is. Validity window is a setting in months; leave it blank to let PCO's own expiration policy decide
- **Preview Before Sync:** every roster sync opens a preview modal showing match counts, sync-mode banner, mirror-removal banner (with a red "will be removed" pill listing how many TP members and stale subgroups will drop), and per-attendee rows. Manual search-and-link for anything unmatched
- **Confirm Spells Out Every Action:** the Sync Now dialog lists adds, drops, subgroup writes, subgroup drops, and attendance writes by name before any DB write. No surprises
- **Scheduled Sync (with Email Summary):**
  - One-click install adds a managed block to TouchPoint's `ScheduledTasks` special content (matches ProspectBuilder's pattern). The per-mapping schedule editor is gated on global install, both client- and server-side
  - Per-mapping: Daily or Weekly, day-of-week and hour, notify a TouchPoint user (typeahead picker by name / username / email), include-issues toggle
  - The runner walks all mappings every invocation and fires anything whose day and hour match now and hasn't run this hour. Each fire syncs fully server-side and emails the configured user: summary counts (joined, already member, subgroup writes, members removed, stale subgroups removed, person fields updated), optional issues list (unmatched, ambiguous matches, PCO API warnings), and a link that opens straight to the mapping it is about
- **Diagnostics (Settings tab):** a read-only health check covering credentials, what your PCO token can actually reach, how many people are linked, mapped involvements that no longer exist, conflicting mappings, scheduler state, and storage. **Copy for support** produces a pasteable summary. A 403 is reported as the permission it is, not as an outage
- **Activity Log (Settings tab):** what the tool has changed, when, and who ran it. Filter to changes only, failures only, or everything. A write whose read-back disagreed shows both values so churn is visible in history
- **Verify-After-Write everywhere else too:** every settings save, mapping save, and scheduler install reads back from storage and confirms the change persisted. Silent permission failures surface immediately with a clear error message
- **Diagnostics On Every Mapping:**
  - Team Mappings have a "Check PCO positions" button that walks PCO and reports exactly what's there: `5 position(s) [Lead Vocal, Backup, ...], 12 assignment(s) across 8 people. Subgroups will sync.` Or the matching red state when positions exist but nobody's assigned in PCO
  - Dashboard health panel surfaces broken mappings (deleted PCO resource, archived TP org) before staff hit Sync
  - Audit log in `PCOSync_Log_YYYYMM` captures per-sync counters (joined, dropped, subgroup adds/drops, failures, scheduler runs, mapping edits, link write/unlink, email send/fail, and every PCO write)
- **Addressable Tabs:** a refresh stays on the tab you were using, and links can point straight at a tab or a specific mapping
- **Last-Sync and Next-Run Pills:** dashboard cards show `Synced 3h ago` (green) and, when scheduled, `Next: Sun 6:00 AM` (blue) so you always know the state at a glance

<hr>

<summary><strong>Sync Dashboard</strong></summary>
<p>One landing page for every mapping. All People at the top (when configured), then per-plan cards under day headers (auto-scrolls to today like the PCO mobile app), then Service Type and Team mapping cards. Each card shows last-synced and next-run pills, plus a one-click Preview & Sync button.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/PCO%20Sync/PCO-Dashboard.png" width="700">
</p>

<summary><strong>Settings & PCO Connection</strong></summary>
<p>Paste your PCO App ID and Secret (Personal Access Token), Test Connection, then set per-field sync rules, configure background check validity, install the global Scheduled Sync block, and run Diagnostics. Verify-after-write on every save catches silent permission failures.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/PCO%20Sync/PCO-Settings.png" width="700">
</p>

<summary><strong>Sync Mappings, One Place for Three Types</strong></summary>
<p>All People (singleton), Service Type Mappings (one PCO Service Type to one umbrella TP involvement, optional team subgroups and per-plan attendance), and Team Mappings (one PCO Team to one TP involvement, optional position subgroups and per-plan attendance). A "Which mapping should I use?" comparison sits above them, and per-row toggles, schedule editor, and inline diagnostics ("Check PCO positions") sit on every row.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/PCO%20Sync/PCO-SyncMapping.png" width="700">
</p>

<summary><strong>People Matching</strong></summary>
<p><strong>Proposed Matches</strong> scores TP candidates for every unmatched PCO record (name, email, and birthdate signals) with Strong / Medium / Weak tiers and bulk Apply for high-confidence hits, plus a manual TouchPoint search on any row when the suggestion is wrong. <strong>Verify Person Link</strong> lets staff inspect any existing TP↔PCO link side-by-side with a one-line verdict, then Unlink or Replace from one panel. Pending Data Reviews collects field-diff changes flagged by your person field rules.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/PCO%20Sync/PCO-PeopleMatching.png" width="700">
</p>

<summary><strong>Scheduled-Sync Email</strong></summary>
<p>Every scheduled run emails the configured TouchPoint user a clean summary: joined, already member, subgroup writes, members removed (mirror), stale subgroups removed, person fields updated. Optionally includes the issues list covering unmatched PCO records, ambiguous email matches, and PCO API warnings, plus a link that opens straight to the mapping the email is about. If a removal was skipped for safety, the email says so prominently.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/PCO%20Sync/PCO-Email.png" width="700">
</p>

<summary><strong>Installation</strong></summary>

1. **Admin > Advanced > Special Content > Python**
2. Click **Add New**, name it `TPxi_PCOSync`
3. Paste the script and Save
4. Navigate to `/PyScriptForm/TPxi_PCOSync`
5. Open the **Settings** tab, paste your PCO **App ID + Secret** (generate one in PCO under *My Account → Applications → Personal Access Tokens*), Save, then click **Test Connection**
6. Run **Settings → Diagnostics** once. It confirms what your token can reach before you rely on it
7. (Optional) Open **Settings → Scheduled Sync** and click **Install** to auto-add the runner to `ScheduledTasks`
8. Switch to **Sync Mappings**, click **+ Add** under any of the three sections (All People / Service Type / Team), pick the PCO resource and the TP involvement, save
9. Hit **Preview & Sync** on the Dashboard card to do your first run

<summary><strong>Sync Mode Cheat Sheet</strong></summary>

| Mode | Best For | Subgroups | Attendance |
|------|----------|-----------|------------|
| **All People** | Mirror every PCO record into a "PCO Directory" involvement so everyone has a TP shell | n/a | n/a |
| **Service Type** | One TP umbrella involvement for a whole worship service (Worship Center, Wilson Hall, etc.). Teams under that service type become subgroups | Optional (Teams as subgroups) | Optional (per-plan, marks Confirmed as Present) |
| **Team** | Each PCO team (Band, Production, Welcome, etc.) lives in its own TP involvement. Positions become subgroups | Optional (Positions as subgroups) | Optional (per-plan) |

You can use any combination. The Band team can have a dedicated Team mapping AND be reflected as a subgroup under the Service Type umbrella.

**Give every mapping its own involvement.** A sync removes anyone in its involvement who is PCO-linked but no longer in its own PCO scope, so two mappings sharing one involvement will each keep removing the other's people. The Mappings tab flags this if it happens.

<summary><strong>How Mirror Removal Works</strong></summary>

**PCO owns involvement membership.** This is about who is on the roster, not about the people themselves: TouchPoint still owns whether a person exists, and person field values follow whatever direction you set per field. On every sync:

- **Adds:** every PCO person matched to a TP person gets `JoinOrg`'d if not already on the roster. New position/team assignments get `AddSubGroup`'d
- **Removes:** every TP member whose `PCO_PersonId` extra value is no longer in PCO's scope (no longer on the team, service type, or directory) gets `RemoveFromOrg`'d. Every TP subgroup membership whose name matches a current PCO position/team but the person no longer holds gets `RemoveSubGroup`'d
- **Untouched:** TP members without a `PCO_PersonId` (manually added) stay. Subgroups whose name doesn't match any current PCO position/team stay, so a manually-added "Pyrotechnics" subgroup is never touched
- **Refused:** if the PCO read was incomplete or returned nobody while TouchPoint has linked members, no removals run at all and the reason is reported. Additions still proceed. A skipped removal is a row that stays for another day; a wrong removal is someone quietly dropped off a serving team

The preview modal shows a red **Mirror removal** banner with counts before you hit Sync, and the confirm dialog spells out every drop by category.

<summary><strong>TouchPoint → PCO Field Sync</strong></summary>

Configured under **Settings → person data sync**. Every field starts at "No sync" and TouchPoint stays authoritative until you choose otherwise.

Direction is one way per field, never both. That is a constraint rather than a simplification: PCO only exposes record-level timestamps, so after a PCO edit there is no reliable way to tell which field changed, which makes "whichever side changed most recently wins" impossible to implement correctly. One direction per field means a value flows one way and stops, and nothing can ping-pong.

This is also why the roster being mirrored from PCO and a field being pushed to PCO are not in conflict. They are different data with different owners.

The workflow is always the same:

1. Set a field to **TouchPoint → PCO**
2. Click **Preview changes**. This reads PCO and reports exactly what an apply would write. It changes nothing
3. Review the table: person, field, TouchPoint value, PCO value, and the action. Rows where the two names are unrelated are flagged **Verify this match**, since that usually means the link is wrong rather than the name
4. Tick the rows you want. Flagged rows start unticked
5. **Apply**. Each write is read back from PCO and confirmed

Start with one low-stakes field, look at the preview, and expand from there. On a first run the preview is often more valuable as a matching audit than as a field sync: unrelated names, dead PCO links, and duplicate TouchPoint people all surface in one table.

<summary><strong>Tips</strong></summary>

- **Match the long-tail with Proposed Matches first.** Open it from any preview to scope to that preview's unmatched (faster). For ongoing maintenance, the unscoped walk processes the whole PCO directory in one shot. Bulk Apply the Strong tier, those are confident matches
- **Use Verify Person Link when something looks off.** "Why is Alice's email wrong?" leads to: search Alice, see her PCO record side-by-side, and if red cells confirm it's the wrong PCO person, Unlink or Replace. Faster than digging through PCO and TP separately
- **Schedule different mappings at different times.** Worship teams sync Sunday at 5 AM to catch Saturday rehearsal changes. All People can run nightly. Each mapping has its own day and time, so you don't have to compromise on one schedule
- **Check PCO Positions diagnostic** is the fastest way to tell apart "subgroups aren't syncing because of a bug" from "PCO has no positions assigned." Run it before opening a support thread. Most "missing subgroup" reports are PCO setup gaps
- **Person field sync defaults off for a reason.** TouchPoint stays authoritative unless you opt in. When you do, prefer queue-for-review over auto-apply for the first few weeks so you can spot PCO-side dirty data before it lands in TP
- **Names are the riskiest field to sync.** TouchPoint keeps a legal first name and a goes-by name; PCO keeps one. The sync uses the goes-by name, but preview it before applying: a name that isn't a variant of the other is usually a bad link, not a nickname
- **Background checks need a PCO permission.** A Personal Access Token inherits the permissions of the PCO user who created it, so background check access has to be granted to that user (PCO: Overview → Background checks → Settings → Background check administrators). Until then the field is skipped entirely rather than written blind
- **Email backlinks open the right place.** The scheduled-sync email links straight to the mapping it is about, and to People Matching for unmatched records

<summary><strong>Storage & Schema</strong></summary>

| Storage | Purpose |
|---------|---------|
| `PCOSync_Settings` | PCO App ID, Secret, API version overrides, background check validity, last-sync timestamps |
| `PCOSync_AllPeopleMapping` | Singleton All People mapping and schedule |
| `PCOSync_PeopleMappings` | Service Type mappings keyed by PCO Service Type ID |
| `PCOSync_TeamMappings` | Team mappings keyed by PCO Team ID |
| `PCOSync_OrgMappings` | Legacy service-type-to-involvement map |
| `PCOSync_AllPeopleSkip` | PCO Person IDs marked as having no TouchPoint equivalent |
| `PCOSync_PersonSyncRules` | Per-field direction and behavior |
| `PCOSync_PendingPersonChanges` | Field changes queued for review |
| `PCOSync_WriteChurn` | Fields whose writes keep coming back different, so they stop being retried |
| `PCOSync_Log_YYYYMM` | Monthly audit log of every write, scheduled run, and email outcome |
| `ScheduledTasks` | Managed block (between markers) that calls back into the script for scheduled runs |
| `PCO_PersonId` (Person Extra Value) | The canonical TP ↔ PCO link, per person |

---
*Written by [Ben Swaby](https://github.com/bswaby). These tools are free because they should be. If they've saved you time, consider [DisplayCache](https://displaycache.com), church digital signage that integrates with TouchPoint, or [TPxi Go](https://tpxigo.com), your church contacts in Outlook and on your phone.*
