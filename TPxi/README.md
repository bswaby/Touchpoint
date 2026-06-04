<div align="center">

<img src="https://fbchville.com/wp-content/uploads/2022/08/BenSwaby.jpg" width="120" style="border-radius: 50%;" />

# TPxi Software™ — TouchPoint® Integrated Tools

**Built by Ben Swaby** · Director of Technology Solutions, First Baptist Church Hendersonville

*Open-source TouchPoint® tools from **[TPxi Software, LLC](https://tpxisoftware.com)** · 100,000+ lines of code · 50+ tools · All free.*

[![Website](https://img.shields.io/badge/Website-tpxisoftware.com-1e40af?style=flat-square)](https://tpxisoftware.com)
[![GitHub Stars](https://img.shields.io/github/stars/bswaby/Touchpoint?style=flat-square&color=f0a500)](https://github.com/bswaby/Touchpoint/stargazers)
[![GitHub Forks](https://img.shields.io/github/forks/bswaby/Touchpoint?style=flat-square&color=3b82f6)](https://github.com/bswaby/Touchpoint/network)
[![Tools](https://img.shields.io/badge/Tools-50%2B-22c55e?style=flat-square)](https://github.com/bswaby/Touchpoint)
[![Churches](https://img.shields.io/badge/Churches%20Using-50%2B-a855f7?style=flat-square)](https://github.com/bswaby/Touchpoint)

</div>

---

## 👋 About the Builder

<img align="right" src="https://fbchville.com/wp-content/uploads/2022/08/BenSwaby.jpg" width="160" style="border-radius: 12px; margin-left: 20px; margin-bottom: 10px;" />

I'm Ben, a church technology director who got tired of waiting for software to do what ministry actually needs. So I started building it myself.

Over the past few years, I've written **100,000+ lines of code** inside TouchPoint®, creating tools that have quietly transformed how dozens of churches track attendance, process giving, re-engage lapsed members, manage volunteers, and understand their own data. I've written a **59-page SQL reference guide** for the TouchPoint® database, led regional developer events, and stay actively involved in the TouchPoint® developer community.

The work eventually outgrew an evening hobby, so it now lives under **[TPxi Software, LLC](https://tpxisoftware.com)**. The 50+ open-source TouchPoint® tools in this repo stay free; the LLC funds the work through two products built TouchPoint®-first: **DisplayCache™** and **TPxi Go™**.

Every tool in this repo was born from a real problem at a real church. None of it is theoretical. All of it is free.

My philosophy is simple: **Kingdom tools should be accessible to every church, regardless of budget.**

> *"Start where it hurts most. Code what matters most. Keep it simple, keep it working."*

<br clear="right"/>

---

## 🚀 Quick Install — 4 Steps

> 💡 Every tool folder has its own README with specifics. This is the general flow.

### 1. Copy the script
Open the tool folder, copy the contents of the `.py` (or `.html`) file.

### 2. Paste into Special Content
In TouchPoint®, go to **Admin → Advanced → Special Content → Python**, click **Add New**, name it (e.g., `TPxi_LiveSearch`), paste, save.

### 3. Surface it in your UI with **[Menu Editor](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Menu%20Editor)**
This is the step most installs skip — a script in Special Content is **invisible to your staff** until you put it on a menu. Install Menu Editor once and it becomes the gateway for every other tool.

**With Menu Editor you can, without touching raw XML:**
- 🧭 Add scripts to the **Blue Toolbar** (the menu that appears when you select people from a search)
- 📋 Add scripts to **Admin / Finance / Involvements / People** report menus
- 🛡️ Set **role-based permissions** from a dropdown of your real roles
- ↕️ **Drag-and-drop** to reorder; move items between columns
- 🛟 **Automatic backup** before every save, one-click restore from history

Install Menu Editor first, then every other tool in this repo is one menu entry away from your staff.

### 4. Configure it
Most newer tools have a **built-in admin UI** — open the script, click into its settings, and you're done. No editing the source.

A few older scripts still use an inline `# CONFIGURATION` block near the top of the file (org IDs, role names, defaults). If the tool's README says "configure these variables," that's the spot. You only need the values listed in the block; leave everything else alone.

---

## ⭐ Start Here — Community Favorites

The tools churches install first and use daily:

| Tool | What It Does in Practice |
|------|--------------------------|
| 🏆 [Live Search](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Live%20Search) | Type a name, see full history, log a note or task in under 10 seconds |
| 🏆 [Involvement Processor](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Involvement%20Processor) | Full registrant processing workflow in one place. Replaces several disconnected steps |
| 🏆 [Report Writer](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Report%20Writer) | Build and save custom involvement and user reports without writing SQL |
| 🏆 [Roll Sheet](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Roll%20Sheet) | Build printable roll sheets exactly the way your teachers want them |
| 🏆 [TechStatus](https://github.com/bswaby/Touchpoint/blob/main/Python%20Scripts/TechStatus/TechStatus) | System health and performance monitoring for your TouchPoint® instance |
| 🏆 [Enterprise Reporting](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Enterprise%20Reporting) | 100+ reports accessible from a single dashboard. Replaces a stack of bookmarks |
| 🏆 [Ministry Structure](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Ministry%20Structure) | See your entire involvement hierarchy at a glance |

---

## 🎯 Scenario Starters

Running a specific ministry moment? Start with these bundles:

**🌞 VBS / Camps**
[Day of Registration](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Day%20of%20Registration) · [Attendance Markings](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Attendance%20Markings) · [Roll Sheet](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Roll%20Sheet) · [Attendance Report Builder](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Attendance%20Builder)

**💰 Year-End Giving / Finance Close**
[Weekly Contribution Report](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Contribution%20Report) · [Giving Dashboard](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Giving%20Dashboard) · [Statement Audit](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Statement%20Audit) · [QCD-Grant Letters](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Finance%20Grant-QCD%20Letter)

**🔄 Lapsed-Member Re-Engagement**
[Lapsed Attenders](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Lapsed%20Attenders) · [Prospector](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Prospector) · [Communication Dashboard](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Communication%20Dashboard) · [Task Runner](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Task%20Runner)

**🛟 New Member Onboarding**
[New Member Report](https://github.com/bswaby/Touchpoint/tree/main/TPxi/New%20Member%20Report) · [Live Search](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Live%20Search) · [Anniversaries Widget](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Anniversaries) · [Operations Checklists](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Operations%20Checklists)

---

## 📦 All Tools by Category

> ⚡ marks tools added in the recent release cycle.

<details>
<summary><strong>👥 People & Attendance</strong> — Track who's there, contact info, milestones</summary>

| Tool | What It Does |
|------|--------------|
| [Live Search](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Live%20Search) | Real-time member search with instant actions |
| [Weekly Attendance (WAAG 2.0)](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Weekly%20Attendance) | Advanced group attendance tracking |
| [Attendance Markings](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Attendance%20Markings) ⚡ | Reverse attendance markings — 1/5 the clicks for a full-attendance event |
| [Attendance Report Builder](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Attendance%20Builder) ⚡ | Easy attendance reports across programs/divisions/orgs |
| [Roll Sheet](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Roll%20Sheet) ⚡ | Build custom printable roll sheets |
| [Emergency List](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Emergency%20List) | Critical contact and medical info management |
| [Anniversaries Widget](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Anniversaries) | Track and celebrate member milestones |
| [New Member Report](https://github.com/bswaby/Touchpoint/tree/main/TPxi/New%20Member%20Report) | Track new member onboarding comprehensively |

</details>

<details>
<summary><strong>💰 Finance & Giving</strong> — Money in, money out, reconciliation</summary>

| Tool | What It Does |
|------|--------------|
| [Weekly Contribution Report](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Contribution%20Report) | The standard for financial reporting and reconciliation |
| [Giving Dashboard](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Giving%20Dashboard) | Clarity and insight into financial stewardship |
| [Statement Audit Dashboard](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Statement%20Audit) | Work through electronic and printed statement issues |
| [Deposit Report](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Deposit%20Report) | Comprehensive deposit reconciliation |
| [Envelope Number Report](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Envelope%20Number%20Report) | Giving envelope SQL report |
| [Find Funds in Batch](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Find%20Funds%20in%20Batch) | Find which batches contain specific funds |
| [Fortis Fees](https://github.com/bswaby/Touchpoint/tree/main/Finance/FortisFees) | Automated fee breakdown |
| [QCD-Grant Letters](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Finance%20Grant-QCD%20Letter) | Automated grant and QCD letter generation |
| [Payment Manager](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Payment%20Manager) | Outstanding payment tracking, receipts, follow-up |
| [Involvement with Fees](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Involvements%20with%20Fees) | Fee-based involvement tracking |

</details>

<details>
<summary><strong>📝 Events & Registration</strong> — Sign-ups, day-of operations, post-event processing</summary>

| Tool | What It Does |
|------|--------------|
| [Day of Registration](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Day%20of%20Registration) ⚡ | Quickly assign walk-ins to classes on event day |
| [Involvement Processor](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Involvement%20Processor) ⚡ | Full registrant processing workflow in one place |
| [Registration Export](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Registration%20Export) | Easy registration data export |
| [Registration Data Manager](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Registration%20Data%20Manager) | Manage registration data at scale |
| [FastLaneCheckIn](https://github.com/bswaby/Touchpoint/tree/main/TPxi/FastLaneCheckIn) | Streamlined large-event check-in |

</details>

<details>
<summary><strong>📣 Outreach & Engagement</strong> — Reach lapsed members, prospects, and the broader community</summary>

| Tool | What It Does |
|------|--------------|
| [Communication Dashboard](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Communication%20Dashboard) | Analyze outreach patterns and effectiveness |
| [Lapsed Attenders](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Lapsed%20Attenders) | Identify and re-engage members going quiet |
| [Prospector](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Prospector) ⚡ | Configurable prospect management workflow |
| [Auxiliary to Group Analytics](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Auxiliary%20to%20Group%20Analytics) | How well are programs driving group participation? |
| [TaskNote Activity Dashboard](https://github.com/bswaby/Touchpoint/tree/main/TPxi/TaskNote%20Activity%20Dashboard) | Monitor task and note activities across the team |

</details>

<details>
<summary><strong>📊 Ministry Insights</strong> — Dashboards and reports that show what's actually happening</summary>

| Tool | What It Does |
|------|--------------|
| [Ministry Structure](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Ministry%20Structure) | Visualize your full involvement hierarchy |
| [Enterprise Reporting](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Enterprise%20Reporting) ⚡ | 100+ reports accessible from a single dashboard |
| [Program Pulse](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Program%20Pulse) ⚡ | Surface what's actually happening across programs |
| [Mission Dashboard](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Mission%20Dashboard) ⚡ | Comprehensive mission trip and activity tracking |
| [Membership Analysis](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Membership%20Analysis) | Deep dive into membership demographics and trends |
| [Involvement Activity Dashboard](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Involvement%20Activity%20Dashboard) | Track and analyze engagement trends |
| [Geographic Distribution Map](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Geographic%20Distribution%20Map) | Visualize where your members live |

</details>

<details>
<summary><strong>🙋 Volunteers, Tasks & Custom Reports</strong> — Coordinate the team and build the report you need</summary>

| Tool | What It Does |
|------|--------------|
| [Volunteer Scheduler Report](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Scheduler%20Report) | Full volunteer scheduling report |
| [Volunteer Widget](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Volunteer%20Widget) | Shows logged-in user's upcoming assignments |
| [Task Runner](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Task%20Runner) ⚡ | Self and team task management |
| [QuickLinks](https://github.com/bswaby/Touchpoint/blob/main/TPxi/Quicklinks) | Permissioned quick-access menu with count overlays |
| [Operations Checklists](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Operations%20Checklists) ⚡ | Organize recurring checks and reminders |
| [Report Writer](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Report%20Writer) ⚡ | Build and save custom involvement and user reports |

</details>

<details>
<summary><strong>🛠️ System & Admin</strong> — Keep the platform healthy and developers productive</summary>

| Tool | What It Does |
|------|--------------|
| [Menu Editor](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Menu%20Editor) ⚡ | Visually edit People, Involvement, Finance, Admin & Blue Toolbar menus |
| [API Explorer](https://github.com/bswaby/Touchpoint/tree/main/TPxi/API%20Explorer) | Explore and test TouchPoint® API queries live |
| [SQL Query Explorer](https://github.com/bswaby/Touchpoint/tree/main/TPxi/SQL%20Query%20Explorer) | Run and explore SQL queries directly |
| [Email Technical Diagnostics](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Email%20Technical%20Diagnostics) | Deep email troubleshooting dashboard |
| [Account Security Monitor](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Account%20Security%20Monitor) | Advanced security analytics |
| [Link Generator](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Link%20Generator) | Pre-authenticated link creation |
| [Attachment Link Downloader](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Attachment%20Link%20Generator) | Bulk document download |
| [Involvement Sync](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Involvement%20Sync) | Synchronize involvement settings across groups |
| [Involvement Owner Audit](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Involvement%20Notification%20Audit%20Tool) | Track involvement leadership and ownership gaps |
| [User Activity](https://github.com/bswaby/Touchpoint/tree/main/TPxi/User%20Activity) | System usage and staff behavior analysis |
| [TechStatus](https://github.com/bswaby/Touchpoint/blob/main/Python%20Scripts/TechStatus/TechStatus) | System status and performance monitoring |

</details>

---

## ☕ Support This Work

If these tools have saved your team hours (and they will), I'd love your support in return. **[TPxi Software, LLC](https://tpxisoftware.com)** funds this community work through two TouchPoint®-first products. Same builder, same philosophy, same heart as the free tools above.

### 📺 DisplayCache™ | Church Digital Signage

**[DisplayCache™](https://displaycache.com)** connects directly to TouchPoint® and pulls **real people and ministry data** onto your screens. No double entry. No stale slides. Just live, ministry-driven content.

| What You Get | Why It Matters |
|---|---|
| 🔗 Live TouchPoint® data on your screens | No more manually updating slides |
| 📺 Works on Apple TV, Fire Stick, Raspberry Pi | Use hardware you already own |
| ✝️ Built specifically for churches | Not a generic signage tool |
| 💰 $10/device/month | Fuels all 50+ free tools here |

**[→ Check Out DisplayCache™](https://displaycache.com)**

### 📱 TPxi Go™ | TouchPoint®, Wherever You Work

**[TPxi Go™](https://tpxigo.com)** brings your TouchPoint® contacts into the apps your team already lives in. Look up anyone, log calls and emails from **Outlook** or your **phone**. No tab switching, no lost context.

| What You Get | Why It Matters |
|---|---|
| 📧 Outlook add-in. Lookup + log emails inline | Stop copy/pasting between Outlook and TouchPoint® |
| 📞 iOS app with Caller ID + call logging | Know who's calling. Log the call in one tap |
| 🔍 Universal search across people, families, orgs | Same search you'd do in TouchPoint®, anywhere |
| 🔒 PAT-based auth. Your data stays in TouchPoint® | No third-party cloud storing your church's contacts |
| 💰 Affordable per-user pricing | Funds the free TPxi tools right alongside DisplayCache™ |

**[→ Check Out TPxi Go™](https://tpxigo.com)** · Available on the App Store, Outlook Add-in store, and Android (coming soon).

> Your subscription to **either** product directly funds this community work. Pick the one that fits, or both. See everything at **[tpxisoftware.com](https://tpxisoftware.com)**.

---

## 📚 Documentation & Resources

| Resource | Link |
|----------|------|
| 📖 SQL Reference Guide | **[TouchPoint® SQL Documentation](https://bswaby.github.io/Touchpoint/DOC_SQLDocumentation.html)** |
| 🎨 NextGen TP Concept Mocks | **[Visualize What TP Could Become](https://github.com/bswaby/Touchpoint/tree/main/NextGen%20TP%20Mocks)** |
| 💬 Community & Support | Open an issue in this repo or find me in the TouchPoint® Discord |

---

## 🤝 Contributing

The most valuable contributions right now are:

- **Bug reports** — If something breaks in your environment, open an issue with your TP version
- **SQL improvements** — Got a better query? Submit a PR with a clear explanation
- **Real-world feedback** — Tell me what's missing or what's confusing. That shapes the roadmap more than anything

---

<div align="center">

Built with ☕ and a deep belief that **technology should serve ministry, not complicate it.**

*Ben Swaby · [TPxi Software, LLC](https://tpxisoftware.com)*

**[TPxi Software](https://tpxisoftware.com)** · **[TPxi Go™](https://tpxigo.com)** · **[DisplayCache™](https://displaycache.com)** · **[GitHub](https://github.com/bswaby/Touchpoint)** · **[SQL Docs](https://bswaby.github.io/Touchpoint/DOC_SQLDocumentation.html)**

<sub>TouchPoint® is a registered trademark of Touchpoint Software, Inc. TPxi Software™, TPxi Go™, and DisplayCache™ are trademarks of TPxi Software, LLC. TPxi Software is not affiliated with, endorsed by, or sponsored by Touchpoint Software, Inc.</sub>

</div>
