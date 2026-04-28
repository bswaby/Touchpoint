### ✅ [Roll Sheet](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Roll%20Sheet)
Configurable rollsheet generator for any ministry area. Build reusable configs once, then print clean two/three/single-column rosters with the columns and data your leaders actually need — registration answers, allergies, emergency contacts, extra values, and more.

- ⚙️ **Implementation Level:** Easy
- 🧩 **Installation:** Single script. Paste into Special Content > Python, navigate to `/PyScriptForm/TPxi_RollSheet`. Configure your first rollsheet from the UI.

<summary><strong>Features</strong></summary>

- **Reusable Configs:** Build a rollsheet config once (source, columns, data sources, layout, sort) and run it any time. Configs persist across script updates
- **Two Source Modes:** Pick by Program / Division (auto-includes all matching active orgs) or by specific Involvements (search-and-add picker)
- **Per-Day Filtering:** Optional "only show involvements with a meeting scheduled on the selected day" toggle.  Uses each Involvement's `OrgSchedule` so you only print rosters for the rooms actually meeting that day
- **Date Presets:** Today, Coming Sunday, Following Sunday, Coming Wednesday or pick any custom date
- **Layouts:** Two Column (default), Three Column, or Single Column. Each layout is print-optimized with page breaks between orgs
- **Standard Columns:** Name, Age, Gender, Phone, Email, Sub-Group, Member Type.  Toggle each on/off per config
- **Name-Line Columns:** Choose which standard fields appear inline with the name vs. as a separate info line. Keeps dense rosters readable
- **Universal Data Sources:** Add unlimited custom info lines per config:
  - **Registration Questions** by wildcard pattern (e.g., `%Allergies%` matches any question containing "Allergies")
  - **Extra Values** (text, date, bit) by field name.  Pick from a dropdown of fields actually used in the org
  - **RecReg Fields** — pre-built shortcuts for Emergency Contact (name + phone), Allergies & Medical, Doctor (name + phone), and Insurance (company + policy)
- **Color-Coded Highlights:** Allergies, medical, and other safety-critical data render in colored callouts so leaders can spot them at a glance
- **Sort Options:** By Name, Age, or Gender,  Sorting is per config
- **Custom Title:** Override the rollsheet header per config (e.g., "VBS 2026 — Morning Session")
- **Footer with Sign-Off:** Optional footer shows org count, total members, and a teacher signature line
- **Sub-Group Display:** When sub-group column is on, each person's `MemberTags` (rooms, cabins, breakouts) appear inline
- **Print Popup:** Generates a clean popup window for printing.  Bypasses TouchPoint's page CSS so background colors and inline styles render correctly on every browser
- **Search Builder for Specific Orgs:** Type-ahead search of organizations by name with member count preview
- **Config Migration:** Older configs with separate medical/recreg fields are auto-migrated to the unified Data Sources format on first open.  No manual cleanup

<hr>

<summary><strong>Main Menu</strong></summary>
<p>Pick a saved config to generate a rollsheet, or create/edit one. Each card shows source, layout, and last updated.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Roll%20Sheet/RollSheet-MainMenu.png" width="700">
</p>

<summary><strong>Create / Edit Config</strong></summary>
<p>Define source, columns, data sources, layout, sort, and title. The same screen edits an existing config or builds a new one from scratch.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Roll%20Sheet/RollSheet-CreateConfig.png" width="700">
</p>

<summary><strong>Generated Rollsheet</strong></summary>
<p>Print-optimized output with the columns, data sources, and layout you configured. Page breaks between orgs, color-coded callouts for medical/safety info, and an optional teacher signature line.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Roll%20Sheet/RollSheet-Example.png" width="700">
</p>

<summary><strong>Installation</strong></summary>

1. **Admin > Advanced > Special Content > Python**
2. Click **Add New**, name it `TPxi_RollSheet`
3. Paste the script and Save
4. Navigate to `/PyScript/TPxi_RollSheet`
5. Click **+ New Config** to build your first rollsheet — pick source, columns, and data sources, then save
6. From the main menu, pick the config and choose a date to generate

<summary><strong>Config Options</strong></summary>

| Setting | Description |
|---------|-------------|
| **Source Type** | Program / Division (auto-includes matching active orgs) or Specific Involvements (manual picker) |
| **Only With Meeting** | When ON, only includes orgs with a meeting scheduled on the selected day-of-week (via `OrgSchedule`) |
| **Layout** | Two Column (default), Three Column, or Single Column |
| **Sort By** | Name, Age, or Gender (applied within each org) |
| **Standard Columns** | Toggle Age, Gender, Phone, Email, Sub-Group, Member Type on/off |
| **Name-Line Columns** | Which standard columns appear inline next to the name vs. as a separate info line |
| **Data Sources** | Unlimited custom rows. Each is a Registration Question (wildcard pattern), Extra Value, or RecReg field with a custom label and color |
| **Title** | Header text for the printed rollsheet (per-config override) |
| **Show Footer** | Toggle the org-count / member-count / signature-line footer |
| **Exclude Orgs** | Optional list of org IDs to skip when source is Program/Division |

<summary><strong>Data Source Types</strong></summary>

| Type | What It Pulls | Example Use |
|------|---------------|-------------|
| **Registration Question** | Answer text from any registration question whose label matches a `%wildcard%` pattern | `%Allergies%`, `%Photo%Release%`, `%Pickup%Authorized%` |
| **Extra Value (Text/Date/Bit)** | A specific named extra value off the person record | `T-shirt Size`, `Background Check Date`, `Photo Release` |
| **RecReg — Emergency Contact** | Standard `RecReg` emergency contact fields (name + phone) | Always-on for kids/youth events |
| **RecReg — Allergies & Medical** | Standard `RecReg` allergy and medical-condition fields, color-coded | Camps, VBS, retreats |
| **RecReg — Doctor** | Doctor name + phone | Overnight events |
| **RecReg — Insurance** | Insurance company + policy number | Off-site trips |

<summary><strong>Tips</strong></summary>

- **Use one config per use case.** A "Sunday Morning Kids" config and a "VBS Summer" config can both target the same Program but with different data sources and layouts
- **Wildcard patterns are your friend.** Instead of hard-coding question IDs, use `%Allergies%` so the same config works across registrations that ask the question with slightly different wording
- **Print Popup, not Print Preview.** The print button opens a clean popup window.  This is intentional. Trying to print directly from the page will strip background colors due to how browsers handle TouchPoint's page CSS

---
*Written by [Ben Swaby](https://github.com/bswaby). These tools are free because they should be. If they've saved you time, consider [DisplayCache](https://displaycache.com) Church digital signage that integrates with TouchPoint or [TPxi Go](https://tpxigo.com) your church contacts in Outlook and on your phone.*
