### 🧭 [Prospector](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Prospect%20Builder)
A configurable prospect management workspace for TouchPoint that combines named configurations, multi-view person-by-person follow-up, group-based outreach tracking, and a dashboard showing outreach health at a glance. Built for staff and lay leaders who need to move people from prospect to engaged member systematically, without losing track of who has been contacted, by whom, and when.

- ⚙️ **Implementation Level:** Easy
- 🧩 **Installation:** Paste-and-go Python script. All states live in Special Content (no schema changes, no admin settings to flip). Configure groups, configurations, and senders from inside the tool.
- 🔐 **Security:** Requires Edit role.

## ✨ Highlights

- **Named Configurations** – Build reusable prospect lists from an involvement or saved query. Run them as often as you want.
- **Group Management** – Define groups of involvements (by program, division, or hand-picked) to track outreach health across a ministry. Each group can have its own member-type, converted-attendance, and stale thresholds.
- **Per-Person Workspace** – List / Single / Batch views for working prospects. Track contact efforts with keyword badges, log calls/visits/emails, and resume saved sessions.
- **Prospect Sender** – Bulk-email senders with per-org tokenization, configurable role gates, and a delivery log.
- **Outreach Dashboard** – KPI cards (active prospects, touches, conversions, overdue, conversion rate) plus rolling 90-day trend charts. Scoped to your Group Management groups so the numbers reflect work you've actually scoped.
- **Activity Log** – Read-only audit trail of every processed prospect, contact effort, and group assignment.

---

<details>
<summary><strong>Prospect Management</strong></summary>
<p>Run any named configuration and work prospects person-by-person. List, Single, and Batch views; contact-effort logging; save your session and resume later.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Prospector/Pros-ProspectManagement.png" width="900" alt="Prospect Management workspace">
</p>
</details>

<details>
<summary><strong>Group Management</strong></summary>
<p>Define groups of involvements (program, division, or specific orgs) and watch the per-involvement view of who is a prospect, who is stale, and who needs a touch. Each group has its own member type, converted attendance types, and freshness thresholds.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Prospector/Pros-GroupManagement.png" width="900" alt="Group Management view">
</p>
</details>

<details>
<summary><strong>Dashboard</strong></summary>
<p>Outreach health at a glance: active prospects, touches (7d / 30d), new conversions, overdue prospects, 90-day conversion rate, and average involvements tried before conversion. Weekly touches and conversions, plus a 90-day rolling conversion-rate trend per group. All metrics are scoped to the groups configured in Group Management.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Prospector/Pros-Dashboard.png" width="900" alt="Outreach Dashboard">
</p>
</details>

---

## How the math works

The dashboard uses **per-involvement pair counts**: a person who is a prospect in two involvements counts as 2 in the denominator, and converts in just one of them counts as 1 in the numerator. This keeps each (person, involvement) treated as its own conversion scenario — closer to how staff actually think about outreach than collapsing to distinct people.

- **A conversion** is a person's **first** converted-type attendance in an involvement during the window, AND they had previously attended that same involvement with a non-converted attendance type — i.e., a real prospect → member transition in the Attend table. The converted attendance type list is configurable per group (default: AttendanceTypeId 30 = Member).
- **The 90-day rolling conversion rate** chart uses the same window definition as the headline card. Each chart point is the per-involvement rate over the 90-day window ending at that month-end (current month uses today). The "Show numbers" toggle under the chart exposes the raw conversion / prospect-involvement counts per anchor for verification.
- **Scope:** every metric is restricted to involvements covered by the groups you have configured in Group Management. People or involvements outside those groups are invisible to the dashboard.

---

## Content Storage

| Key | Purpose |
|---|---|
| `ProspectBuilder_Configs` | Named prospect configurations |
| `ProspectBuilder_Groups` | Group Management groups + assignments + efforts |
| `ProspectBuilder_Settings` | Install-wide settings (contact methods, default sender, scorecard weights) |
| `ProspectBuilder_Senders` | Saved sender configurations |
| `ProspectBuilder_SenderLog` | Sender delivery history |
| `ProspectBuilder_WorkStates` | Saved work sessions |
| `ProspectBuilder_ActivityLog` | Audit trail |
| `ProspectBuilder_GroupMetrics` | Cached per-group metrics |

No SQL schema changes. All persistence is in TouchPoint's Content table.

---

## Installation

1. In TouchPoint, go to **Admin > Advanced > Special Content > Python**
2. Click **New Python Script File**
3. Name it `TPxi_ProspectBuilder` and paste the code
4. Navigate to `/PyScript/TPxi_ProspectBuilder` to launch
5. Optionally add to the menu via **CustomReports**:
   ```xml
   <Report name="TPxi_ProspectBuilder" type="PyScript" role="Edit" />
   ```

**First run:** open the **Group Management** tab and create at least one group. The Dashboard will be empty until at least one group is configured.

---
*Written by [Ben Swaby](https://github.com/bswaby). These tools are free because they should be. If they've saved you time, consider [DisplayCache](https://displaycache.com) Church digital signage that integrates with TouchPoint or [TPxi Go](https://tpxigo.com) your church contacts in Outlook and on your phone.*
