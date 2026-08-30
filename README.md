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

## Start here

The tools churches install first and use daily.

| Tool | What it does |
|---|---|
| [Live Search](TPxi/Live%20Search) | Type a name, see full history, log a note or task in under 10 seconds |
| [Menu Editor](TPxi/Menu%20Editor) | Put any script on any menu, with role permissions and automatic backups |
| [Involvement Processor](TPxi/Involvement%20Processor) | Full registrant processing workflow in one place |
| [Report Writer](TPxi/Report%20Writer) | Build and save custom involvement and people reports without SQL |
| [Roll Sheet](TPxi/Roll%20Sheet) | Printable roll sheets built the way your teachers want them |
| [Enterprise Reporting](TPxi/Enterprise%20Reporting) | 100+ reports behind one dashboard |
| [Weekly Contribution Report](TPxi/Contribution%20Report) | The standard for weekly giving reconciliation |
| [TechStatus](Python%20Scripts/TechStatus/TechStatus) | Health and performance monitoring for your instance |

**[See all 58 tools, searchable, with a summary of each →](https://bswaby.github.io/Touchpoint/tools.html)**

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
