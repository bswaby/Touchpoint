<div align="center">

<img src="docs/assets/ben-swaby.jpg" alt="Ben Swaby" width="120">

### TPxi Software™ — TouchPoint® Integrated Tools

**Less Admin. More Ministry.**

58 free, open-source tools for TouchPoint®, built at a real church and used at 50+ others.

[**Browse all 58 tools →**](https://bswaby.github.io/Touchpoint/tools.html)

[![Website](https://img.shields.io/badge/Website-tpxisoftware.com-1e40af?style=flat-square)](https://tpxisoftware.com)
[![Tools](https://img.shields.io/badge/Tools-58-22c55e?style=flat-square)](https://bswaby.github.io/Touchpoint/tools.html)
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

Click any tool to read what it does, then open its folder. A star marks the tools churches install
first; **new** marks a recent release.

*Prefer to search instead of scroll? The same list is
[browsable and searchable here](https://bswaby.github.io/Touchpoint/tools.html).*

### People & Attendance <sub>10 tools</sub>

<details>
<summary>&#9733; <b>Live Search</b></summary>

Type part of a name and get that person's full history on one screen. Log a note or a task without leaving the results.

**[Open Live Search](TPxi/Live%20Search)**

</details>
<details>
<summary><b>Weekly Attendance (WAAG 2.0)</b></summary>

A rebuilt weekly attendance view for groups, with the comparisons and trend lines the standard report never gave you.

**[Open Weekly Attendance (WAAG 2.0)](TPxi/Weekly%20Attendance)**

</details>
<details>
<summary><b>Attendance Markings</b> &nbsp;<code>new</code></summary>

Marks a whole roll present in one pass, then you uncheck the few who missed. About a fifth of the clicks for a full-attendance event.

**[Open Attendance Markings](TPxi/Attendance%20Markings)**

</details>
<details>
<summary><b>Attendance Report Builder</b> &nbsp;<code>new</code></summary>

Build attendance reports across programs, divisions and involvements without writing any SQL.

**[Open Attendance Report Builder](TPxi/Attendance%20Builder)**

</details>
<details>
<summary><b>In-Progress Registrations</b></summary>

Shows the registrations people started and never finished, so you can follow up before the event fills.

**[Open In-Progress Registrations](TPxi/In-Progress%20Registrations)**

</details>
<details>
<summary><b>Duplicate Involvements</b></summary>

Finds people sitting in more than one involvement when they should only be in one.

**[Open Duplicate Involvements](TPxi/Duplicate%20Involvements)**

</details>
<details>
<summary>&#9733; <b>Roll Sheet</b> &nbsp;<code>new</code></summary>

Printable roll sheets built the way your teachers actually want them: which columns, what order, how much room to write.

**[Open Roll Sheet](TPxi/Roll%20Sheet)**

</details>
<details>
<summary><b>Emergency List</b></summary>

Pulls emergency contacts and medical notes for a group into one list you can hand to a leader.

**[Open Emergency List](TPxi/Emergency%20List)**

</details>
<details>
<summary><b>Anniversaries Widget</b></summary>

Surfaces birthdays, membership anniversaries and other milestones on the home page so they don't slip past.

**[Open Anniversaries Widget](TPxi/Anniversaries)**

</details>
<details>
<summary><b>New Member Report</b></summary>

Tracks new members through onboarding: who joined, what they've connected to, and who has gone quiet.

**[Open New Member Report](TPxi/New%20Member%20Report)**

</details>

### Finance & Giving <sub>12 tools</sub>

<details>
<summary>&#9733; <b>Weekly Contribution Report</b></summary>

The weekly giving and reconciliation report most churches running these tools open every Monday morning.

**[Open Weekly Contribution Report](TPxi/Contribution%20Report)**

</details>
<details>
<summary><b>Giving Dashboard</b></summary>

Giving trends, donor movement and fund performance in one view instead of five exports.

**[Open Giving Dashboard](TPxi/Giving%20Dashboard)**

</details>
<details>
<summary><b>Statement Audit Dashboard</b></summary>

Finds the statement problems before your donors do: bad addresses, missing emails, electronic statements that never arrived.

**[Open Statement Audit Dashboard](TPxi/Statement%20Audit)**

</details>
<details>
<summary><b>Deposit Report</b></summary>

Reconciles a deposit end to end so the bank, the batch and the books agree.

**[Open Deposit Report](TPxi/Deposit%20Report)**

</details>
<details>
<summary><b>Coupon Report</b></summary>

Shows where coupons are being used across ministries and what they are costing.

**[Open Coupon Report](TPxi/Coupon%20Report)**

</details>
<details>
<summary><b>Envelope Number Report</b></summary>

A SQL report for giving envelope numbers: who has one, and who is actually using it.

**[Open Envelope Number Report](TPxi/Envelope%20Number%20Report)**

</details>
<details>
<summary><b>Last 4 Search</b></summary>

Look up a transaction by the last four digits of the card or bank account, when that's all the donor can tell you.

**[Open Last 4 Search](TPxi/Transaction%20Search)**

</details>
<details>
<summary><b>Find Funds in Batch</b></summary>

Tells you which batches contain a specific fund when you're chasing down a posting error.

**[Open Find Funds in Batch](TPxi/Find%20Funds%20in%20Batch)**

</details>
<details>
<summary><b>Fortis Fees</b></summary>

Breaks down Fortis processing fees automatically instead of by hand every month.

**[Open Fortis Fees](Finance/FortisFees)**

</details>
<details>
<summary><b>QCD-Grant Letters</b></summary>

Generates acknowledgement letters for qualified charitable distributions and donor-advised grants.

**[Open QCD-Grant Letters](TPxi/Finance%20Grant-QCD%20Letter)**

</details>
<details>
<summary>&#9733; <b>Payment Manager</b></summary>

Tracks outstanding balances on fee-based registrations, sends receipts and handles the follow-up.

**[Open Payment Manager](TPxi/Payment%20Manager)**

</details>
<details>
<summary><b>Involvement with Fees</b></summary>

Lists every involvement carrying a fee and shows what has been paid against it.

**[Open Involvement with Fees](TPxi/Involvements%20with%20Fees)**

</details>

### Events & Registration <sub>5 tools</sub>

<details>
<summary><b>Day of Registration</b> &nbsp;<code>new</code></summary>

Assign walk-ins to classes and sessions at the table on event day, without slowing the line down.

**[Open Day of Registration](TPxi/Day%20of%20Registration)**

</details>
<details>
<summary>&#9733; <b>Involvement Processor</b> &nbsp;<code>new</code></summary>

The full registrant processing workflow in one screen. Replaces several disconnected steps you used to do in order.

**[Open Involvement Processor](TPxi/Involvement%20Processor)**

</details>
<details>
<summary><b>Registration Export</b></summary>

Exports registration data, question answers included, in a shape you can actually work with.

**[Open Registration Export](TPxi/Registration%20Export)**

</details>
<details>
<summary><b>Registration Data Manager</b></summary>

Edit and clean registration answers in bulk instead of one record at a time.

**[Open Registration Data Manager](TPxi/Registration%20Data%20Manager)**

</details>
<details>
<summary><b>FastLaneCheckIn</b></summary>

A stripped-down check-in flow for large events where the standard station can't keep up.

**[Open FastLaneCheckIn](TPxi/FastLaneCheckIn)**

</details>

### Outreach & Engagement <sub>5 tools</sub>

<details>
<summary><b>Communication Dashboard</b></summary>

Shows what you're sending, who is opening it, and which ministries have gone quiet.

**[Open Communication Dashboard](TPxi/Communication%20Dashboard)**

</details>
<details>
<summary><b>Lapsed Attenders</b></summary>

Finds people whose attendance has dropped off, early enough that a phone call still makes sense.

**[Open Lapsed Attenders](TPxi/Lapsed%20Attenders)**

</details>
<details>
<summary><b>Prospector</b> &nbsp;<code>new</code></summary>

A configurable pipeline for working prospects: assign, track, follow up, close the loop.

**[Open Prospector](TPxi/Prospector)**

</details>
<details>
<summary><b>Auxiliary to Group Analytics</b></summary>

Answers whether your events and auxiliary programs are actually moving people into groups.

**[Open Auxiliary to Group Analytics](TPxi/Auxiliary%20to%20Group%20Analytics)**

</details>
<details>
<summary><b>TaskNote Activity Dashboard</b></summary>

Shows task and note activity across the staff: what is moving, and what has stalled.

**[Open TaskNote Activity Dashboard](TPxi/TaskNote%20Activity%20Dashboard)**

</details>

### Ministry Insights <sub>7 tools</sub>

<details>
<summary>&#9733; <b>Ministry Structure</b></summary>

Your whole program, division and involvement hierarchy on one page, so you can see where it doesn't line up.

**[Open Ministry Structure](TPxi/Ministry%20Structure)**

</details>
<details>
<summary>&#9733; <b>Enterprise Reporting</b> &nbsp;<code>new</code></summary>

Over 100 reports behind a single dashboard. Replaces a folder of bookmarks and half-remembered saved queries.

**[Open Enterprise Reporting](TPxi/Enterprise%20Reporting)**

</details>
<details>
<summary><b>Program Pulse</b> &nbsp;<code>new</code></summary>

A quick read on what is actually happening in each program right now.

**[Open Program Pulse](TPxi/Program%20Pulse)**

</details>
<details>
<summary><b>Missions Dashboard</b> &nbsp;<code>new</code></summary>

Tracks mission trips end to end: teams, funding and who has gone.

**[Open Missions Dashboard](TPxi/Mission%20Dashboard)**

</details>
<details>
<summary><b>Membership Analysis</b></summary>

Demographic and trend analysis of your membership: who you have, and how that is changing.

**[Open Membership Analysis](TPxi/Membership%20Analysis)**

</details>
<details>
<summary><b>Involvement Activity Dashboard</b></summary>

Engagement trends by involvement, so a leader can see their own group's direction without asking you.

**[Open Involvement Activity Dashboard](TPxi/Involvement%20Activity%20Dashboard)**

</details>
<details>
<summary>&#9733; <b>Geographic Distribution Map</b></summary>

Maps where your people live. Useful for campus, small group and outreach decisions.

**[Open Geographic Distribution Map](TPxi/Geographic%20Distribution%20Map)**

</details>

### Volunteers, Tasks & Reports <sub>6 tools</sub>

<details>
<summary><b>Volunteer Scheduler Report</b></summary>

A full report on scheduled volunteers: who is serving, where you're short, who never confirmed.

**[Open Volunteer Scheduler Report](TPxi/Scheduler%20Report)**

</details>
<details>
<summary><b>Volunteer Widget</b></summary>

Shows the logged-in person their own upcoming assignments on the home page.

**[Open Volunteer Widget](TPxi/Volunteer%20Widget)**

</details>
<details>
<summary><b>Task Runner</b> &nbsp;<code>new</code></summary>

Task management for you and your team, inside TouchPoint® rather than in yet another app.

**[Open Task Runner](TPxi/Task%20Runner)**

</details>
<details>
<summary><b>QuickLinks</b></summary>

A permissioned quick-access menu with live counts on the links that matter.

**[Open QuickLinks](TPxi/Quicklinks)**

</details>
<details>
<summary><b>Operations Checklists</b> &nbsp;<code>new</code></summary>

Recurring checks and reminders: the weekly and monthly things that get forgotten until they bite.

**[Open Operations Checklists](TPxi/Operations%20Checklists)**

</details>
<details>
<summary>&#9733; <b>Report Writer</b> &nbsp;<code>new</code></summary>

Build and save your own involvement and people reports without writing SQL.

**[Open Report Writer](TPxi/Report%20Writer)**

</details>

### System & Admin <sub>13 tools</sub>

<details>
<summary>&#9733; <b>Menu Editor</b> &nbsp;<code>new</code></summary>

Puts any script onto the People, Involvement, Finance, Admin or Blue Toolbar menus, with role permissions from a dropdown and an automatic backup before every save. Install this one first, and every other tool is one menu entry away from your staff.

**[Open Menu Editor](TPxi/Menu%20Editor)**

</details>
<details>
<summary><b>API Explorer</b></summary>

Try TouchPoint® API calls live and see what comes back before you write anything against them.

**[Open API Explorer](TPxi/API%20Explorer)**

</details>
<details>
<summary><b>SQL Query Explorer</b></summary>

Run and explore SQL against your database directly, from inside TouchPoint®.

**[Open SQL Query Explorer](TPxi/SQL%20Query%20Explorer)**

</details>
<details>
<summary><b>Email Technical Diagnostics</b></summary>

Traces email problems: what sent, what bounced, and what the receiving server actually said.

**[Open Email Technical Diagnostics](TPxi/Email%20Technical%20Diagnostics)**

</details>
<details>
<summary><b>Account Security Monitor</b></summary>

Security analytics on user accounts: logins, roles and access that doesn't look right.

**[Open Account Security Monitor](TPxi/Account%20Security%20Monitor)**

</details>
<details>
<summary><b>Person-Audit Detail</b></summary>

Shows everywhere a person has served and attended. The report you want in hand when a question comes up about someone.

**[Open Person-Audit Detail](TPxi/Person%20Attendance%20Audit)**

</details>
<details>
<summary><b>CSV Phone Matcher</b></summary>

Matches a CSV of phone numbers back to people records.

**[Open CSV Phone Matcher](TPxi/CSV%20Phone%20Matcher.)**

</details>
<details>
<summary><b>Link Generator</b></summary>

Creates pre-authenticated links, so people land where you sent them instead of at a login wall.

**[Open Link Generator](TPxi/Link%20Generator)**

</details>
<details>
<summary><b>Attachment Link Downloader</b></summary>

Bulk-downloads documents attached across records.

**[Open Attachment Link Downloader](TPxi/Attachment%20Link%20Generator)**

</details>
<details>
<summary><b>Involvement Sync</b></summary>

Copies involvement settings across many groups at once so they stay consistent.

**[Open Involvement Sync](TPxi/Involvement%20Sync)**

</details>
<details>
<summary><b>Involvement Owner Audit</b></summary>

Finds involvements with no leader, no notification target, or the wrong person still attached.

**[Open Involvement Owner Audit](TPxi/Involvement%20Notification%20Audit%20Tool)**

</details>
<details>
<summary><b>User Activity</b></summary>

System usage by staff: who is in, what they use, and what they never touch.

**[Open User Activity](TPxi/User%20Activity)**

</details>
<details>
<summary>&#9733; <b>TechStatus</b></summary>

Health and performance monitoring for your TouchPoint® instance.

**[Open TechStatus](Python%20Scripts/TechStatus/TechStatus)**

</details>

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
three ways to put something back.

**[DisplayCache™](https://displaycache.com)** — church digital signage that pulls live TouchPoint®
data onto your screens. Runs on Apple TV, Fire Stick or Raspberry Pi. $10/device/month.

**[TPxi Go™](https://tpxigo.com)** — your directory, caller ID, and call and email logging, in
Outlook and on your phone. PAT-based auth, so your data stays in TouchPoint®.

**[Give once](https://buy.stripe.com/REPLACE_ANY)** — if a subscription isn't a fit but a tool here
saved your ministry a weekend, you can chip in directly. TPxi Software, LLC is a for-profit company,
so this is support rather than a tax-deductible donation.

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
