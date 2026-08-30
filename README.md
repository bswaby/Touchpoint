<div align="center">

<img src="https://fbchville.com/wp-content/uploads/2022/08/BenSwaby.jpg" alt="Ben Swaby" width="120">

### TPxi Software™ — TouchPoint® Integrated Tools

**Less Admin. More Ministry.**

58 free, open-source tools for TouchPoint®, built at a real church and used at 50+ others.

[**Browse all 58 tools →**](https://bswaby.github.io/Touchpoint/Tools.html)

[![Website](https://img.shields.io/badge/Website-tpxisoftware.com-1e40af?style=flat-square)](https://tpxisoftware.com)
[![Tools](https://img.shields.io/badge/Tools-58-22c55e?style=flat-square)](https://bswaby.github.io/Touchpoint/Tools.html)
[![Churches](https://img.shields.io/badge/Churches-50%2B-a855f7?style=flat-square)](https://github.com/bswaby/Touchpoint)
[![Stars](https://img.shields.io/github/stars/bswaby/Touchpoint?style=flat-square&color=f0a500)](https://github.com/bswaby/Touchpoint/stargazers)

</div>

---

## About

I'm Ben Swaby, Director of Technology Solutions at First Baptist Church Hendersonville. I got tired
of waiting for software to do what ministry actually needs, so I started building it myself.

Over the past few years that's become 100,000+ lines of code inside TouchPoint® — tools that track
attendance, process giving, re-engage lapsed members, manage volunteers and make sense of church
data. Along the way I wrote a [59-page SQL reference](https://bswaby.github.io/Touchpoint/DOC_SQLDocumentation.html)
for the TouchPoint® database and stay active in the developer community.

The work outgrew an evening hobby, so it now lives under [TPxi Software, LLC](https://tpxisoftware.com).
The tools in this repo stay free. The LLC funds them through two TouchPoint®-first products,
[DisplayCache™](https://displaycache.com) and [TPxi Go™](https://tpxigo.com).

Every tool here came from a real problem at a real church. None of it is theoretical.
Kingdom tools should be accessible to every church, regardless of budget.

## Install any tool in four steps

**1. Copy the script.** Open the tool's folder and copy the contents of its `.py` file.

**2. Paste it into TouchPoint®.** Go to **Admin → Advanced → Special Content → Python**, click
**Add New**, name it (`TPxi_LiveSearch`), paste, save.

**3. Put it on a menu with [Menu Editor](TPxi/Menu%20Editor).** A script in Special Content is
invisible to your staff until it's on a menu. Menu Editor adds scripts to the Blue Toolbar and the
Admin, Finance, Involvement and People menus, sets role permissions from a dropdown of your real
roles, reorders by drag-and-drop, and backs up automatically before every save. **Install this one
first** — after that, every other tool is one menu entry away.

**4. Configure it.** Most newer tools have a built-in admin UI: open the script, click into
settings, done. A few older ones have a `# CONFIGURATION` block near the top of the file instead;
the tool's README will say so.

## All 58 tools

&#9733; marks the tools churches install first. Every tool's own folder has a fuller README.
For a searchable version with a longer write-up on each one, see the
**[tool browser](https://bswaby.github.io/Touchpoint/Tools.html)**.

### People & Attendance

| Tool | What it does |
|---|---|
| &#9733; [Live Search](TPxi/Live%20Search) | Real-time member search with instant actions |
| [Weekly Attendance (WAAG 2.0)](TPxi/Weekly%20Attendance) | Advanced group attendance tracking |
| [Attendance Markings](TPxi/Attendance%20Markings) `new` | Mark a full roll present, then uncheck the few who missed |
| [Attendance Report Builder](TPxi/Attendance%20Builder) `new` | Attendance reports across programs, divisions and involvements |
| [In-Progress Registrations](TPxi/In-Progress%20Registrations) | Find and finish incomplete registrations |
| [Duplicate Involvements](TPxi/Duplicate%20Involvements) | People sitting in more than one involvement |
| &#9733; [Roll Sheet](TPxi/Roll%20Sheet) `new` | Printable roll sheets built the way teachers want them |
| [Emergency List](TPxi/Emergency%20List) | Emergency contacts and medical info for a group |
| [Anniversaries Widget](TPxi/Anniversaries) | Birthdays and membership milestones on the home page |
| [New Member Report](TPxi/New%20Member%20Report) | Track new members through onboarding |

### Finance & Giving

| Tool | What it does |
|---|---|
| &#9733; [Weekly Contribution Report](TPxi/Contribution%20Report) | The standard for weekly giving reconciliation |
| [Giving Dashboard](TPxi/Giving%20Dashboard) | Giving trends, donor movement and fund performance |
| [Statement Audit Dashboard](TPxi/Statement%20Audit) | Catch statement problems before your donors do |
| [Deposit Report](TPxi/Deposit%20Report) | End-to-end deposit reconciliation |
| [Coupon Report](TPxi/Coupon%20Report) | Coupon use and cost across ministries |
| [Envelope Number Report](TPxi/Envelope%20Number%20Report) | Giving envelope SQL report |
| [Last 4 Search](TPxi/Transaction%20Search) | Find a transaction by the last four digits |
| [Find Funds in Batch](TPxi/Find%20Funds%20in%20Batch) | Which batches contain a specific fund |
| [Fortis Fees](Finance/FortisFees) | Automated processing fee breakdown |
| [QCD-Grant Letters](TPxi/Finance%20Grant-QCD%20Letter) | Automated grant and QCD letter generation |
| &#9733; [Payment Manager](TPxi/Payment%20Manager) | Outstanding balances, receipts and follow-up |
| [Involvement with Fees](TPxi/Involvements%20with%20Fees) | Fee-based involvement tracking |

### Events & Registration

| Tool | What it does |
|---|---|
| [Day of Registration](TPxi/Day%20of%20Registration) `new` | Assign walk-ins to classes on event day |
| &#9733; [Involvement Processor](TPxi/Involvement%20Processor) `new` | Full registrant processing in one screen |
| [Registration Export](TPxi/Registration%20Export) | Registration data out, question answers included |
| [Registration Data Manager](TPxi/Registration%20Data%20Manager) | Edit registration answers in bulk |
| [FastLaneCheckIn](TPxi/FastLaneCheckIn) | Stripped-down check-in for large events |

### Outreach & Engagement

| Tool | What it does |
|---|---|
| [Communication Dashboard](TPxi/Communication%20Dashboard) | What you send, and who actually opens it |
| [Lapsed Attenders](TPxi/Lapsed%20Attenders) | Spot attendance dropping off early enough to call |
| [Prospector](TPxi/Prospector) `new` | Configurable prospect pipeline |
| [Auxiliary to Group Analytics](TPxi/Auxiliary%20to%20Group%20Analytics) | Are your events moving people into groups? |
| [TaskNote Activity Dashboard](TPxi/TaskNote%20Activity%20Dashboard) | Task and note activity across the staff |

### Ministry Insights

| Tool | What it does |
|---|---|
| &#9733; [Ministry Structure](TPxi/Ministry%20Structure) | Your whole involvement hierarchy on one page |
| &#9733; [Enterprise Reporting](TPxi/Enterprise%20Reporting) `new` | 100+ reports behind a single dashboard |
| [Program Pulse](TPxi/Program%20Pulse) `new` | What is happening in each program right now |
| [Missions Dashboard](TPxi/Mission%20Dashboard) `new` | Mission trips, teams and funding |
| [Membership Analysis](TPxi/Membership%20Analysis) | Membership demographics and trends |
| [Involvement Activity Dashboard](TPxi/Involvement%20Activity%20Dashboard) | Engagement trends by involvement |
| &#9733; [Geographic Distribution Map](TPxi/Geographic%20Distribution%20Map) | Where your people actually live |

### Volunteers, Tasks & Reports

| Tool | What it does |
|---|---|
| [Volunteer Scheduler Report](TPxi/Scheduler%20Report) | Who is serving, and where you are short |
| [Volunteer Widget](TPxi/Volunteer%20Widget) | A person's own upcoming assignments |
| [Task Runner](TPxi/Task%20Runner) `new` | Team and personal task management |
| [QuickLinks](TPxi/Quicklinks) | Permissioned quick-access menu with live counts |
| [Operations Checklists](TPxi/Operations%20Checklists) `new` | Recurring checks and reminders |
| &#9733; [Report Writer](TPxi/Report%20Writer) `new` | Build and save custom reports without SQL |

### System & Admin

| Tool | What it does |
|---|---|
| &#9733; [Menu Editor](TPxi/Menu%20Editor) `new` | Put any script on any menu, with role permissions |
| [API Explorer](TPxi/API%20Explorer) | Try TouchPoint® API calls live |
| [SQL Query Explorer](TPxi/SQL%20Query%20Explorer) | Run SQL against your database directly |
| [Email Technical Diagnostics](TPxi/Email%20Technical%20Diagnostics) | What sent, what bounced, and why |
| [Account Security Monitor](TPxi/Account%20Security%20Monitor) | Login, role and access analytics |
| [Person-Audit Detail](TPxi/Person%20Attendance%20Audit) | Everywhere a person has served and attended |
| [CSV Phone Matcher](TPxi/CSV%20Phone%20Matcher.) | Match a CSV of phone numbers to people records |
| [Link Generator](TPxi/Link%20Generator) | Pre-authenticated links, no login wall |
| [Attachment Link Downloader](TPxi/Attachment%20Link%20Generator) | Bulk document download |
| [Involvement Sync](TPxi/Involvement%20Sync) | Copy involvement settings across groups |
| [Involvement Owner Audit](TPxi/Involvement%20Notification%20Audit%20Tool) | Involvements with missing or wrong leaders |
| [User Activity](TPxi/User%20Activity) | Staff system usage and behavior |
| &#9733; [TechStatus](Python%20Scripts/TechStatus/TechStatus) | Instance health and performance monitoring |

## Running a specific ministry moment?

**VBS and camps** — [Day of Registration](TPxi/Day%20of%20Registration) ·
[Attendance Markings](TPxi/Attendance%20Markings) · [Roll Sheet](TPxi/Roll%20Sheet) ·
[Attendance Report Builder](TPxi/Attendance%20Builder)

**Year-end giving** — [Weekly Contribution Report](TPxi/Contribution%20Report) ·
[Giving Dashboard](TPxi/Giving%20Dashboard) · [Statement Audit](TPxi/Statement%20Audit) ·
[QCD-Grant Letters](TPxi/Finance%20Grant-QCD%20Letter)

**Lapsed-member re-engagement** — [Lapsed Attenders](TPxi/Lapsed%20Attenders) ·
[Prospector](TPxi/Prospector) · [Communication Dashboard](TPxi/Communication%20Dashboard) ·
[Task Runner](TPxi/Task%20Runner)

**New member onboarding** — [New Member Report](TPxi/New%20Member%20Report) ·
[Live Search](TPxi/Live%20Search) · [Anniversaries](TPxi/Anniversaries) ·
[Operations Checklists](TPxi/Operations%20Checklists)

## Support this work

Every tool here is free, with no account and no licence. If they've saved your team hours, there are
two ways to put something back.

**[DisplayCache™](https://displaycache.com)** — digital signage, presentation and room signage for
churches. Integrates with eSpace, TouchPoint®, Planning Center, OneDrive and Google Drive, and runs on
Apple TV, Fire Stick or Raspberry Pi. $10/device/month.

**[TPxi Go™](https://tpxigo.com)** — your directory, caller ID, and call and email logging, in
Outlook and on your phone. PAT-based auth, so your data stays in TouchPoint®.

## Documentation

- [TouchPoint® SQL reference](https://bswaby.github.io/Touchpoint/DOC_SQLDocumentation.html) — 59 pages on the database
- [NextGen TP concept mocks](NextGen%20TP%20Mocks) — what TouchPoint® could become
- Questions? Open an issue here, or find me in the TouchPoint® Discord

## Contributing

The most useful contributions right now:

- **Bug reports.** If something breaks in your environment, open an issue and include your TP version.
- **SQL improvements.** Got a better query? Send a PR with a short explanation of what changed.
- **Real-world feedback.** What's missing, what's confusing. That shapes the roadmap more than anything else.

---

<div align="center">

Built with coffee and a deep belief that technology should serve ministry, not complicate it.

[tpxisoftware.com](https://tpxisoftware.com) ·
[TPxi Go™](https://tpxigo.com) ·
[DisplayCache™](https://displaycache.com) ·
[TPxi Scan](https://tpxiscan.com) ·
[SQL docs](https://bswaby.github.io/Touchpoint/DOC_SQLDocumentation.html)

<sub>TouchPoint® is a registered trademark of Touchpoint Software, Inc. TPxi Software™, TPxi Go™ and
DisplayCache™ are trademarks of TPxi Software, LLC. TPxi Software is not affiliated with, endorsed
by, or sponsored by Touchpoint Software, Inc.</sub>

</div>
