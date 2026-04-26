### ✅ [Operations Checklists](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Operations%20Checklists)
Group-based recurring operations management for the things your team needs to do every day, week, month, or year.  Data-health checks, Sunday-morning prep, month-end finance, new-member follow-up, IT audits. Mix automated SQL checks with manual checklist items, get morning-batch reminder emails, and pull from a community marketplace of pre-built checks.

- ⚙️ **Implementation Level:** Easy
- 🧩 **Installation:** Single script. Paste into Special Content > Python, navigate to `/PyScript/TPxi_OpsChecklists`. Configure groups and roles via the Settings panel.

<summary><strong>Features</strong></summary>

- **Group-Based Architecture:** Organize checks into Groups (Data Health, Sunday Operations, Month-End Finance, IT Security, etc.). Each with its own owner, frequency, and roster of checks
- **Automated SQL Checks:** Run live SQL on demand . Green when 0 rows returned, red with the row count and a drill-down grid when issues exist
- **Saved Search Checks:** Point a check at a TouchPoint Saved Search or raw Search Builder code.  Same pass/fail UX as SQL
- **Manual Checklist Items:** Pure-checklist items for non-data tasks (building unlocked, AV tested, batch reconciled)
- **Sub-Steps with Notes:** Break any check into per-step actions and capture per-step notes that persist between runs
- **Drill-Down Person Grid:** SQL/search results render in an ag-Grid table with sortable columns, person picker, and bulk-tag actions
- **Bulk Tag from Results:** Tag everyone (or selected rows) directly from a check's results — no copy/paste between screens
- **Community Marketplace:** Browse 50+ pre-built checks across People, Involvements, Email, Finance, Facilities, General, and Archive categories. One-click install into any group
- **Marketplace Updates:** When a marketplace check gets a new revision, you get a per-item Update button plus dashboard "Update All".  Local edits are detected and protected before overwrite
- **"New" Badge:** Items added to the marketplace within the last 14 days get a green "new" badge so you can spot fresh community contributions
- **Share a Check:** Submit your own SQL/search checks back to the marketplace for community review — approved submissions become available to other churches
- **Per-Check Reminder Emails:** Assign specific people (or fall back to a Default Recipient) to receive one consolidated email when checks are due
- **One-Click MorningBatch Install:** Auto-install/uninstall the reminder block in your existing `MorningBatch` script.  No manual editing required
- **Send Reminders Now:** Trigger today's reminders on demand without waiting for the next morning-batch run
- **Calendar View:** Month-grid showing what's due each day, what was completed, and what was missed
- **Completion Log:** Per-check history of who ran it, when, and the row count returned
- **Person Picker:** Type-ahead search for assigning recipients. Handles both "First Last" and "Last, First" search patterns
- **Saved Search Autocomplete:** Type-ahead search of saved searches by name, with description preview
- **Two-Tier Permissions:** Edit Roles (manage groups, checks, library) and Complete Roles (run checks, mark steps done) configurable from Settings
- **Auto-Update:** Checks DisplayCache for new versions (Admin/Developer only) and updates in-place
- **Help System:** Built-in TOC of contextual help topics. Groups, Checks, Catalog, Reminders, Sharing
- **Defaults Catalog:** First-time-open auto-installs a curated starter set into a "Data Health" group so new users have something to run immediately

<summary><strong>Marketplace Categories</strong></summary>

| Category | Examples |
|----------|----------|
| **People** | Missing emails / cell phones / DOB, Mrs./Mr. gender mismatch, deceased still active, duplicates, family-position gaps |
| **Involvements** | Active orgs with no members, inactive orgs with current members, orgs missing leader / Entry Point |
| **Email** | Bounced / failed (last 7 days), spam / reputation failures (last 30 days), email account notifications |
| **Finance** | Reconcile contribution batches, returned-check follow-up, year-end giving statements, fund-balance review |
| **Facilities** | Verify building unlocked, AV/sound test, check-in stations ready, HVAC pre-service |
| **General** | Sensitive role audit, review user access, update church info, backup special content |
| **Archive** | Previous Member not archived, archive candidates (3.5yr no activity), deceased records review |

<hr>

<summary><strong>Group Selection</strong></summary>
<p>Pick the operations group you want to work in. Each group has its own checks, frequency, and roster.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Operations%20Checklists/OCL-GroupSelection.png" width="700">
</p>

<summary><strong>Group Detail</strong></summary>
<p>Run checks, see pass/fail at a glance, drill into details, mark sub-steps complete, and bulk-act on results.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Operations%20Checklists/OCL-Group.png" width="700">
</p>

<summary><strong>Check Library (Marketplace)</strong></summary>
<p>Browse pre-built community checks by category. Toggle on the ones you want, click Save Changes — they install into your group.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Operations%20Checklists/OCL-CheckLibrary.png" width="700">
</p>

<summary><strong>Library Updates</strong></summary>
<p>Items with a newer marketplace revision show an inline Update button. Local edits are detected and protected before overwrite.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Operations%20Checklists/OCL-CheckLibraryUpdates.png" width="700">
</p>

<summary><strong>Create Your Own Check</strong></summary>
<p>Build a custom SQL, saved-search, or manual-checklist item. Add sub-steps, set frequency, assign reminder recipients.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Operations%20Checklists/OCL-CreateCheck.png" width="700">
</p>

<summary><strong>Reminder Email</strong></summary>
<p>One consolidated morning-batch email per person — every check that's due, with direct links and counts.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Operations%20Checklists/OCL-Email.png" width="700">
</p>

<summary><strong>Installation</strong></summary>

1. **Admin > Advanced > Special Content > Python**
2. Click **Add New**, name it `TPxi_OpsChecklists`
3. Paste the script and Save
4. Navigate to `/PyScript/TPxi_OpsChecklists`
5. On first open, a curated "Data Health" group is auto-installed so you have something to run immediately
6. (Optional) Click the gear icon to set Edit/Complete roles, default reminder recipient, and enable morning-batch reminders

<summary><strong>Settings Overview</strong></summary>

| Setting | Description |
|---------|-------------|
| **Edit Roles** | Comma-separated TouchPoint roles allowed to add/edit/delete groups, checks, and library items (Admin always has access) |
| **Complete Roles** | Roles allowed to mark checks complete, toggle steps, and run results |
| **Default Recipient** | Person who receives reminders for any check without specific assignees |
| **Send Reminders on MorningBatch** | Toggle the morning-batch consolidation email on/off |
| **MorningBatch Install** | One-click add/remove of this script from your existing MorningBatch content |
| **Send Reminders Now** | Manually trigger today's reminders without waiting for morning batch |

<summary><strong>Reminder Email Setup</strong></summary>

The script can append a managed block to your existing `MorningBatch` content automatically (Settings → "Add to MorningBatch"). When morning batch runs, every check due that day collects its assigned recipients (or falls back to the Default Recipient), and each person gets one consolidated email listing every check that needs their attention — with direct links and pass/fail row counts.

If you'd rather edit MorningBatch by hand, the Settings panel shows the exact block to paste.

<summary><strong>How "New" Badges Work</strong></summary>

The dashboard banner and Manage Library button show a count of items that are uninstalled AND added to the marketplace within the last 14 days. After 14 days, items just become "uninstalled" — the badge clears automatically so you don't get permanent clutter. The 14-day window applies to both the dashboard count and the per-item ✨ `new` badge inside Manage Library.

---
*Written by [Ben Swaby](https://github.com/bswaby). These tools are free because they should be. If they've saved you time, consider [DisplayCache](https://displaycache.com) Church digital signage that integrates with TouchPoint or [TPxi Go](https://tpxigo.com) your church contacts in Outlook and on your phone.*
