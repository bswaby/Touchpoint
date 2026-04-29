### ✅ [Attendance Report Builder](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Attendance%20Builder)
Configurable historical attendance reports for any ministry area. Build a config once (source, sections, leader types, sort), pick a date range, and get daily summaries, program/division breakdowns, per-involvement detail cards, subgroup breakdowns, and enrollment ratios.  All from one report.

- ⚙️ **Implementation Level:** Easy
- 🧩 **Installation:** Single script. Paste into Special Content > Python, navigate to `/PyScript/TPxi_AttendanceBuilder`. Configure your first report from the UI.

<summary><strong>Features</strong></summary>

- **Reusable Configs:** Build a report config once (source, sections, leader rules, sort, title) and run it for any date range. Configs persist across script updates
- **Historical Accuracy:** Queries the `Attend` table for historical attendance, not current `OrganizationMembers`. People who attended in the past still show up in the report even if they've since dropped out of the involvement
- **Two Source Modes:** Pick by Program / Division (auto-includes all matching active involvements) or by Specific Involvements (manual picker). Mix multiple Program/Division pairs in one config
- **Toggleable Report Sections:** Build a focused or comprehensive report by flipping each section on/off
  - Daily Summary Table
  - Program / Division Breakdown
  - Involvement Detail Cards
  - Subgroup Breakdown (uses `MemberTags` — rooms, cabins, breakouts)
  - Enrollment Ratio (attended vs. enrolled %)
  - Separate Leaders vs Members
- **Leader Classification:** Configurable list of `MemberTypeId` values (default `140, 160`) that get counted as leaders rather than attendees, so leadership numbers don't inflate member attendance
- **Count Distinct vs Sum:** When the same person attends multiple involvements in your scope, choose whether to count them once (distinct) or per-involvement (sum). Distinct is right for unique-people reporting; sum is right for total attendance hours
- **Inactive Involvements Toggle:** Include retired/inactive involvements when reporting on past data (e.g., a closed VBS week's history). Off by default to keep current reports clean
- **Date Range:** Free-form start/end date. Run for a Sunday, a week, a month, a year.  Whatever fits the question
- **Sort Options:** Sort Involvements by name or by attendance count, applied within each section
- **Custom Title:** Override the report header per config (e.g., "VBS 2026 — Morning Session")
- **Exclude Involvements:** Optional list of involvement IDs to skip when source is Program/Division. Handy for excluding admin/leader-only orgs
- **Print Popup:** Generates a clean popup window for printing. Bypasses TouchPoint's page CSS so background colors and inline styles render correctly on every browser
- **Search Builder for Specific Involvements:** Type-ahead search of involvements by name with member count preview

<hr>

<summary><strong>Main Menu</strong></summary>
<p>Pick a saved config to generate a report, or create/edit one. Each card shows source summary and last updated.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Attendance%20Builder/ARB-Main.png" width="700">
</p>

<summary><strong>Create / Edit Config</strong></summary>
<p>Define source, report sections, leader types, sort, and title. The same screen edits an existing config or builds a new one from scratch.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Attendance%20Builder/ARB-Configuration.png" width="700">
</p>

<summary><strong>Generated Report</strong></summary>
<p>Print-optimized output combining the sections you enabled. Daily summary at the top, then breakdowns by program/division, per-involvement detail cards, optional subgroup breakdowns, and enrollment ratios.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Attendance%20Builder/ARB-Report.png" width="700">
</p>

<summary><strong>Installation</strong></summary>

1. **Admin > Advanced > Special Content > Python**
2. Click **Add New**, name it `TPxi_AttendanceBuilder`
3. Paste the script and Save
4. Navigate to `/PyScriptForm/TPxi_AttendanceBuilder`
5. Click **+ New Config** to build your first report.  Pick source, toggle sections, set leader types, and save
6. From the main menu, pick the config and choose a date range to generate

<summary><strong>Config Options</strong></summary>

| Setting | Description |
|---------|-------------|
| **Source Type** | Program / Division (auto-includes matching active involvements) or Specific Involvements (manual picker) |
| **Program / Division Groups** | Stack multiple combos (e.g., "Adults / Sunday School" + "Children / Kids Church") to scope the report across several areas |
| **Specific Involvement IDs** | Manual list when you want a precise scope, ignoring Program/Division |
| **Exclude Involvement IDs** | Optional list of Involvement IDs to skip when source is Program/Division |
| **Leader Member Type IDs** | Comma-separated `MemberTypeId` values that count as leaders (default `140, 160`). Used by "Separate Leaders" section |
| **Sort By** | Name or Attendance Count (within each section) |
| **Title** | Header text for the printed report (per-config override) |

<summary><strong>Report Sections</strong></summary>

| Section | What It Shows | Default |
|---------|---------------|---------|
| **Daily Summary Table** | Date-by-date attendance totals across all in-scope involvements | ON |
| **Program / Division Breakdown** | Sub-totals grouped by Program and Division | ON |
| **Organization Detail Cards** | One card per involvement with date-by-date attendance | ON |
| **Subgroup Breakdown** | When `MemberTags` are used (rooms, cabins, breakout groups), shows attendance per subgroup | OFF |
| **Enrollment Ratio** | Adds an "attended / enrolled = %" column to each Involvement card | OFF |
| **Separate Leaders vs Members** | Splits attendance into Leader column vs Member column using your Leader Member Type IDs | ON |
| **Count Distinct People** | When ON, a person attending two involvements counts once total. When OFF, counts twice (once per org) | OFF |
| **Include Inactive Involvements** | Pull historical data from involvements that are now inactive (e.g., past VBS weeks) | OFF |

<summary><strong>Common Use Cases</strong></summary>

| Question | Recommended Settings |
|----------|----------------------|
| **Quarterly attendance trend for an entire program** | Source = Program; sections = Daily Summary + Program/Division Breakdown; date range = 90 days |
| **VBS week recap by room** | Source = Specific Involvements; sections = Daily Summary + Involvment Detail + Subgroup Breakdown; date range = the VBS week; Include Inactive = ON if VBS invovlements are already archived |
| **Sunday school enrollment health** | Source = Program/Division; sections = Involvement Detail + Enrollment Ratio + Separate Leaders; date range = past 8 weeks |
| **Leadership engagement** | Source = Program; sections = Involvment Detail + Separate Leaders; toggle Count Distinct = ON to avoid double-counting leaders who serve in two areas |

<summary><strong>Tips</strong></summary>

- **Use one config per question.** A "Sunday Morning Health" config and a "Wednesday Night Health" config can target the same Program with different date ranges and section emphasis
- **Distinct vs Sum matters more than you'd think.** For "how many unique people attended this month?" use Distinct. For "how many seats were filled?" use Sum
- **Leader Member Type IDs vary by church.** Default `140, 160` works for most churches but verify against your `lookup.MemberType` table — the wrong IDs will misclassify your numbers
- **Inactive involvements default to OFF for a reason.** Leave it off for current-state reports; flip it on only when you specifically want past, retired involvements in the picture (e.g., last year's camp involvements that are now archived)
- **Print Popup, not Print Preview.** The print button opens a clean popup window. This is intentional — printing directly from the page strips background colors and disrupts the layout

---
*Written by [Ben Swaby](https://github.com/bswaby). These tools are free because they should be. If they've saved you time, consider [DisplayCache](https://displaycache.com) Church digital signage that integrates with TouchPoint or [TPxi Go](https://tpxigo.com) your church contacts in Outlook and on your phone.*
