### 📅 [Schedule List](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Schedule%20List)
The built-in calendar view in TouchPoint is complicated and hard for volunteers to quickly scan.  This script creates a simple, clean schedule view so team members can easily see who is serving and when.  It's built to be viewed from the web, printed, and emailed out.  It supports morning batch automation so schedules can be sent on a recurring basis — to the full involvement, to yourself, or to specific people filtered by subgroup.

- ⚙️ **Implementation Level:** Easy
- 🧩 **Installation:** Paste the script into Special Content > Python. Add to CustomReports for Blue Toolbar access. Configuration is done in the User Config section at the top of the script. Org-specific overrides are supported via Organization Extra Values.

<summary><strong>Schedule View</strong></summary>
<p>Clean multi-column layout grouped by date with color-coded fill status (Full, Partial, Empty).</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Schedule%20Report/SchedulerReport.png" width="700">
</p>


<summary><strong>Morning Batch Options</strong></summary>
<p>Automate schedule emails on a recurring basis with flexible filtering and delivery options.</p>

| Parameter | Required | Description |
|-----------|----------|-------------|
| `sendReport` | Yes | Set to `'y'` to trigger email send |
| `CurrentOrgId` | Yes | Organization ID to report on |
| `reportTo` | Yes | `'Involvement'`, `'Self'`, or `'PeopleId'` |
| `reportToPeopleId` | If PeopleId | One or more PeopleIds, comma-separated |
| `subgroupFilter` | No | Only include specific subgroup(s), comma-separated |
| `scheduleDays` | No | Override default 365-day window (e.g., `'7'` for next week) |
| `showManageCommitments` | No | `'false'` to hide button in email |
| `showFamilyButtons` | No | `'0'` to hide family links in email |

**Example: Send filtered subgroup report to pastor every Friday**
```python
from datetime import datetime
today = datetime.now()

if today.weekday() == 4:  # Friday
    model.Data.sendReport = 'y'
    model.Data.reportTo = 'PeopleId'
    model.Data.reportToPeopleId = '3134'
    model.Data.subgroupFilter = 'LEO - Pastor'
    model.Data.scheduleDays = '7'
    model.Data.showManageCommitments = 'false'
    model.Data.showFamilyButtons = '0'
    model.Data.CurrentOrgId = '2832'
    print(model.CallScript("TPxi_ScheduleList"))
```

<summary><strong>Configuration Options</strong></summary>
<p>All defaults are set in the User Config section of the script. Several can be overridden per-organization using Extra Values.</p>

| Setting | Default | Org Extra Value |
|---------|---------|-----------------|
| `ShowEmptySlots` | `True` | `SchedulerReportShowEmptySlots` (bool) |
| `ShowFamilyButtons` | `1` | `SchedulerReportShowFamilyButtons` (int) |
| `ShowManageCommitments` | `True` | — |
| `MinVolunteerAge` | `12` | `SchedulerReportMinVolunteerAge` (int) |
| `Title` | Org name + Schedule | `SchedulerReportTitle` (string) |
| `FromAddress` | `scheduler@church.org` | `SchedulerReportFromAddress` (string) |
| `Subject` | Org name | `SchedulerReportSubject` (string) |
| `UseMultiColumn` | `True` | — |
| `ColumnsPerRow` | `2` | — |

---
*Like this tool? [DisplayCache](https://displaycache.com) integrates directly with TouchPoint and 
helps fund continued development of tools like this one.*
