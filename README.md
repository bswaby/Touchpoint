## 📋 Overview

This repository is the work of one guy who loves Jesus, loves clean data, and loves seeing people come to Christ — so I built a bunch of tools to help make ministry smoother, more informed, reduce overhead, and help reach others.
---

<img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Widget%20QuickLinks/WidgetQuickLink.png" width="700">

💰 Finance</br>
- 💳 [FortisFees](https://github.com/bswaby/Touchpoint/blob/main/README.md#-fortisfees) - Automated fee breakdown and back-charging
- 📜 [QCD-Grant Letters](https://github.com/bswaby/Touchpoint/blob/main/README.md#-qcd-grantletters) - Automated letter generation for windowed envelopes
- 📜 [Payment Manager](https://github.com/bswaby/Touchpoint/blob/main/README.md#-payment-manager) - Outstanding payment tracking and management
- 📊 [Weekly Contribution Report](https://github.com/bswaby/Touchpoint/blob/main/README.md#-weekly-contribution-report) - Weekly finance tracking and reporting

📈 Reports</br>
- 📅 [Anniversaries](https://github.com/bswaby/Touchpoint/blob/main/README.md#-anniversaries) - Member anniversary tracking widget
- 📊 [Auxiliary to Group Analytics](https://github.com/bswaby/Touchpoint/blob/main/README.md#-auxiliary-to-group-analytics) - Program effectiveness analysis
- 📱 [Communication Dashboard](https://github.com/bswaby/Touchpoint/blob/main/README.md#-communication-dashboard) - Email and SMS analytics
- 🧹 [Data Quality Dashboard](https://github.com/bswaby/Touchpoint/blob/main/README.md#-data-quality-dashboard) - Database completeness monitoring
- 🧑‍🤝‍🧑 [Involvement Activity Dashboard](https://github.com/bswaby/Touchpoint/blob/main/README.md#-involvement-activity-dashboard) - Involvement activity analysis
- 📉 [Lapsed Attenders](https://github.com/bswaby/Touchpoint/blob/main/README.md#-lapsed-attenders) - Statistical attendance pattern analysis
- 🏛️ [Ministry Structure](https://github.com/bswaby/Touchpoint/blob/main/README.md#%EF%B8%8F-ministry-structure) - Program and involvement structure overview
- 📤 [Registration Export](https://github.com/bswaby/Touchpoint/blob/main/README.md#-registration-export) - Event registration data export
- 📝 [TaskNote Activity Dashboard](https://github.com/bswaby/Touchpoint/blob/main/README.md#-tasknote-activity-dashboard) - Task and note activity monitoring
- ✅ [Weekly Attendance](https://github.com/bswaby/Touchpoint/blob/main/README.md#-weekly-attendance) - Enhanced attendance reporting (WAAG 2.0)

🔧 Programs</br>
- 🌍 [Missions Dashboard](https://github.com/bswaby/Touchpoint/blob/main/README.md#-missions-dashbaord) - Mission trip management system

🛠️ Tools</br>
- 🔗 [Attachment Link Downloader](https://github.com/bswaby/Touchpoint/blob/main/README.md#-attachment-link-downloader) - Bulk document downloads
- 🔗 [Account Security Monitor](https://github.com/bswaby/Touchpoint/blob/main/README.md#-account-security-monitor) - Advanced security analytics
- 🚗 [FastLaneCheckIn](https://github.com/bswaby/Touchpoint/blob/main/README.md#-fastlanecheckin) - Streamlined event check-in system
- 🗺️ [Geographic Distribution Map](https://github.com/bswaby/Touchpoint/blob/main/README.md#%EF%B8%8F-geographic-distribution-map) - Member location mapping with demographics
- 🛂 [Involvement Owner Audit](https://github.com/bswaby/Touchpoint/blob/main/README.md#-involvement-owner-audit) - Involvement ownership tracking
- ↻  [Involvement Sync](https://github.com/bswaby/Touchpoint/blob/main/README.md#-involvement-sync) - Multi-involvement synchronization
- 🔐 [Link Generator](https://github.com/bswaby/Touchpoint/blob/main/README.md#-link-generator) - Pre-authenticated link creation
- 🖥️ [TechStatus](https://github.com/bswaby/Touchpoint/blob/main/README.md#%EF%B8%8F-techstatus) - System status monitoring
- 🗺️ [User Activity](https://github.com/bswaby/Touchpoint/blob/main/README.md#%EF%B8%8F-user-activity) - User activity level analysis

🧩 Widgets</br>
- ⚡ [Widget QuickLinks](https://github.com/bswaby/Touchpoint/blob/main/README.md#-widget-quicklinks) - Permission-based quick access links

## 💰 Finance

### 💳 [FortisFees](https://github.com/bswaby/Touchpoint/blob/main/Finance/FortisFees)
It breaks down fees by program and accounting code so accounting can use it to back-charge. The biggest note is that the fees will be close to 99% but not 100% due to the disconnect of certain charges and reversals getting back to Touchpoint.  The most prominent mention of this script is how much time it saved our finance team. Before this, it would have taken 6-8 hours a month to figure out the backcharges.  Now, with this script, they may spend 15 minutes.

Update: 20250610 - Resolved issue with ACH, improved looks, added inline comments, and flow remarks.

- ⚙️ **Implementation Level: Easy
- 🧩 **Installation: Paste in code and change the fee percentages

### 📜 [QCD-GrantLetters](https://github.com/bswaby/Touchpoint/blob/main/Finance/QCD-GrantLetters)
It automatically creates QCD and Grant letters that you can print out and stuff into windowed envelopes. In addition, you can upload these as a secure note to each person's record.  On an average week for us, it saved the finance team 2-3 hours compared to how it was done prior.  

- ⚙️ **Implementation Level: Moderate-Advanced
- 🧩 **Installation: This script is built around the fact that you have a separate batch type for Grant and QCD.  Updating the letter information will require basic knowledge of HTML, but it should be easy to implement if you follow the pattern.  The advanced side is that this is set up to print on windowed envelopes, and it can be tricky (not terrible) to get aligned based on your environment and printer settings.

<details>
<summary><strong>Date Selection</strong></summary>
<p>The date selection interface allows users to choose dates to pull letters from.  Once you have pulled the letters, it allows you to upload a note to each persons record.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/Finance/screenshots/QCDGrantScreenShot.png" width="700">
</p>
</details>

<details>
<summary><strong>Grant Letter</strong></summary>
<p>When a grant letter is detected, it will use the grant letter template and create a letter for each person made for a printable windowed envelope</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/Finance/screenshots/GrantLetter.png" width="700">
</p>
</details>

<details>
<summary><strong>QCD Letter</strong></summary>
<p>When a QCD letter is detected, it will use the QCD letter template and create a letter for each person made for a printable windowed envelope</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/Finance/screenshots/QCDLetteer.png" width="700">
</p>
</details>

### 📜 [Payment Manager](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Payment%20Manager)
The payment manager was built to help us identify outstanding payments across programs and then dive into each. When you dive in, you can make payments with receipts for cash and check, send a quick digital pay link, or resend receipts.   The biggest note is that this was coded for our environment, and while it should work, I am not 100% sure what you will run into as I haven't done any testing outside our environment for this one.

- ⚙️ **Implementation Level: Easy-Moderate
- 🧩 **Installation: There are a few config items on top of the script.  Most do not need to be edited, but at a minimum, I would say you need to edit the DEFAULT_EMAIL_SENDER.  

<details>
<summary><strong>Program Outstanding</strong></summary>
<p>Shows outstanding payments at a payment level and can either immediately show people outstanding from there or dive into a division/involvement view.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Payment%20Manager/PM_Programs.png" width="700">
</p>
</details>

<details>
<summary><strong>Division View</strong></summary>
<p>Shows a breakdown of how much is owed per divison/involvement.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Payment%20Manager/PM_Divisions.png" width="700">
</p>
</details>

<details>
<summary><strong>User Fees</strong></summary>
<p>Breakdown of user fees with some actionable steps</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Payment%20Manager/PM_UserFees.png" width="700">
</p>
</details>

<details>
<summary><strong>Payment</strong></summary>
<p>Make a cash/check payment that sends the person a receipt.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Payment%20Manager/PM_Payment.png" width="700">
</p>
</details>

### 📊 [Weekly Contribution Report](https://github.com/bswaby/Touchpoint/blob/main/Finance/Weekly%20Contribution%20Report)
Our finance team uses this primary tool to track, report, and work through finances each week.  By using the system as their process and reporting mechanism, they have helped ensure the system is right to move forward and remove some of the hocus pocus potential of changing numbers outside the system. One of the pleasant surprises for us is that it revealed some newer features in Touchpoint that we were not taking advantage of.  This helped us to start to take advantage of those and further organize/define our financial data.  

- ⚙️ **Implementation Level: Easy
- 🧩 **Installation: The Script is pasted and go. Once you are in the UI, you will complete all the configuration in the UI for you financial setup.

---
## 📈 Reports

### 📅 [Anniversaries](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Anniversaries)
It is designed as a widget on your dashboard that tracks member anniversaries, such as marriage, work, and birthdays.  We use it to track staff anniversaries, but it could easily be used for any group as it's based on a saved search.

- ⚙️ **Implementation Level: Moderate
- 🧩 **Installation: Weddings and birthdays are tracked naturally in the system and uses a "saved search" to pull the list from. Extra anniversaries are tracked using an extra value field under a user's profile, but are not required.  To install, you first create a saved search in Touchpoint and then write down the name of it.  Then, paste the code, set the parameters, add the word widget to content keywords, and save it.  From there, you need to go into Admin ~ Advanced ~ HomePage widget and add this as a widget to have it appear on the homepage.

<details>
<summary><strong>Anniversary Main Screen</strong></summary>
<p>You can click through months to show past and upcoming anniversaries.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Anniversaries/Anniversaries%20Main%20Screen.png" width="700">
</p>
</details>

<details>
<summary><strong>Quick Kudo Email Send</strong></summary>
<p>Clicking the email icon pops up a quick kudo message that people can send</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Anniversaries/Anniversaries%20Pop-Up.png" width="700">
</p>
</details>

### 📊 [Auxiliary to Group Analytics](https://github.com/bswaby/Touchpoint/blob/main/TPxi/Auxiliary%20to%20Group%20Analytics/TPxi_Auxiliary2Group.py)
The Auxiliary to Group Analytics dashboard helps churches analyze how effectively their various programs drive attendance to Groups. It tracks actual attendance patterns, identifies true conversions (people who joined Connect Groups AFTER participating in programs), and provides summary and detailed metrics to measure program success in building ongoing community connections.

- ⚙️ Implementation Level: Easy
- 🧩 **Installation: Minimal configuration required. Edit the CHURCH_PROGRAM_ID (your Group program ID) and update the PROGRAM_IDS dictionary with your specific programs to analyze. The tool uses actual attendance data from the Attend table to ensure accurate metrics.

<details>
<summary><strong>Dashboard Overview</strong></summary>
<p>Example overall and specific group summary</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Auxiliary%20to%20Group%20Analytics/A2G-Dashboard.png" width="700">
</p>
</details>

### 📱 [Communication Dashboard](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Communication%20Dashboard)
Centralized dashboard for viewing and analyzing email and SMS communications.  This has helped us further understand who is sending messages out, find issues with SMS, and allowed us to work through some strategic communication methods, such as parents getting multiple emails from various age-grade pastors covering similar subjects.

- ⚙️ **Implementation Level: Easy
- 🧩 **Installation: This is a paste-and-go Python script with no configuration needed.  

<details>
<summary><strong>Dashboard Overview</strong></summary>
<p>Summary of outgoing Email and SMS with a few KPI's</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Communication%20Dashboard/CD-DashboardOverview.png" width="700">
</p>
</details>

<details>
<summary><strong>Email Status</strong></summary>
<p>Shows overall delivery performance of email, failure type breakdown, top issues, and recent email campaigns going out.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Communication%20Dashboard/CD-EmailStats.png" width="700">
</p>
</details>

<details>
<summary><strong>SMS Stats</strong></summary>
<p>Shows KPI's, failures, and top senders around SMS messages</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Communication%20Dashboard/CD-SMSStats.png" width="700">
</p>
</details>

<details>
<summary><strong>Top Senders</strong></summary>
<p>List of top email and SMS senders.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Communication%20Dashboard/CD-TopSenders.png" width="700">
</p>
</details>

<details>
<summary><strong>Program Stats</strong></summary>
<p>Top Emails by Program</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/blob/main/TPxi/Communication%20Dashboard/CD-ProgramStats.png" width="700">
</p>
</details>

### 🧹 [Data Quality Dashboard](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Data%20Quality%20Dashboard)
Monitors the completeness of the database demographics and what might be missing as data comes in.  

- ⚙️ **Implementation Level: Easy
- 🧩 **Installation:  There are a couple of configuration options, but this is a paste-and-go script.

<details>
<summary><strong>Overview</strong></summary>
<p>Shows record count and percentages of missing data areas.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Data%20Quality%20Dashboard/DataQualityOverview1.png" width="700">
</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Data%20Quality%20Dashboard/DataQualityOverview2.png" width="700">
</p>
</details>

<details>
<summary><strong>Problem Records</strong></summary>
<p>Good way to see data missing as it's changed.  This can reveal forms or methods that are not capturing data you might need.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Data%20Quality%20Dashboard/DataQualityProblemRecords.png" width="700">
</p>
</details>

<details>
<summary><strong>Recommended Actions</strong></summary>
<p>Recommended actions with ability to export list to csv.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Data%20Quality%20Dashboard/DataQualityRecommendedActions.png" width="700">
</p>
</details>

### 🧑‍🤝‍🧑 [Involvement Activity Dashboard](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Involvement%20Activity%20Dashboard)
This is like Kenny Rogers of tools in that it helps you know "when to hold them" and "when to fold them" regarding your Involvements.  Overall, this shows you active to dormant involvements and helps you improve the cleanliness of your database. 

- ⚙️ **Implementation Level: Easy
- 🧩 **Installation:  There are a couple of configuration options, but this is a paste-and-go script.

<details>
<summary><strong>Overview</strong></summary>
<p>Overall metrics for involvement activity that can be filtered down by program/division.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Involvement%20Activity%20Dashboard/IA-Overview.png" width="700">
</p>
</details>

<details>
<summary><strong>Programs & Divisions</strong></summary>
<p>Activity level by program / division.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Involvement%20Activity%20Dashboard/IA-ProgramsAndDivisions.png" width="700">
</p>
</details>

<details>
<summary><strong>Activity Metrics</strong></summary>
<p>Using a scoring system to help break down the activity level of each involvement to show most --> to least active.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Involvement%20Activity%20Dashboard/IA-ActivityMetrics.png" width="700">
</p>
</details>

<details>
<summary><strong>Meetings</strong></summary>
<p>Meeting activity across programs..</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Involvement%20Activity%20Dashboard/IA-Meetings.png" width="700">
</p>
</details>

<details>
<summary><strong>Member Changes</strong></summary>
<p>Activity level of changes across involvements</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Involvement%20Activity%20Dashboard/IA-MemberChanges.png" width="700">
</p>
</details>

<details>
<summary><strong>Inactive Involvements</strong></summary>
<p>Report showing involvements that have a low or no activity level.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Involvement%20Activity%20Dashboard/IA-InactiveInvolvements.png" width="700">
</p>
</details>

### 📉 Lapsed Attenders
An advanced statistical analysis dashboard that identifies people whose attendance patterns have significantly deviated from their normal behavior. Uses standard deviation calculations to find members who haven't attended in longer than their typical pattern suggests, helping pastoral staff proactively reach out before people fully disconnect.
Key Capabilities:

📊 Statistical analysis using standard deviations to identify lapsed patterns</br>
🎯 Priority scoring system (URGENT/HIGH/MEDIUM) based on deviation severity</br>
📋 Bulk actions: mass tagging, task creation, and note adding</br>
👥 Contact tracking to avoid duplicate outreach efforts</br>
🔍 Advanced filtering by age groups (adults vs children)</br>
⚙️ Configurable thresholds adaptable to different church contexts</br>
📱 Mobile-responsive ag-Grid interface with visual indicators</br>

- ⚙️ Implementation Level: Moderate - Requires SQL knowledge for program ID configuration
- 🧩 Installation: Upload as Python script, configure excluded program IDs and thresholds for your church's structure

<details>
<summary><strong>Lapsed Attender Dashboard Screenshot</strong></summary>
<p>Interactive dashboard showing statistical analysis with priority levels, contact status tracking, and bulk action capabilities. Each person's deviation from their normal attendance pattern is calculated and displayed with visual priority indicators.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Lapsed%20Attenders/LA-Dashboard.png" width="700">
</p>
</details>
The key differentiator here is that this isn't just "who missed church recently" - it's "who is behaving differently than their established pattern," which is much more actionable for pastoral care!

### 🏛️ [Ministry Structure](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Ministry%20Structure)
This is one of my favorite tools for understanding the structure of programs, divisions, involvements, and involvement types, as it provides a simple and effective means of laying out the structure.

- ⚙️ **Implementation Level: Easy
- 🧩 **Installation: This is a paste-and-go Python script with no configuration needed.  

<details>
<summary><strong>Main Screen</strong></summary>
<p>Gives you a structural breakdown of your involvements.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Ministry%20Structure/MinistryStructure.png" width="700">
</p>
</details>

### 📤 [Registration Export](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Registration%20Export)
Exports event or class registration data in structured formats.

### 📝 [TaskNote Activity Dashboard](https://github.com/bswaby/Touchpoint/tree/main/TPxi/TaskNote%20Activity%20Dashboard)
I developed this to understand TaskNote activity, who has open assignments, keyword trends, and more.  Believe it or not, we are weak in digital process for Touchpoint items, so this is a stepping stone for me to start to monitor and understand what we are currently doing.

- ⚙️ **Implementation Level: Easy
- 🧩 **Installation: This is a paste-and-go Python script with no configuration needed.  

<details>
<summary><strong>Overview</strong></summary>
<p>Overall view of Task and Notes</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/TaskNote%20Activity%20Dashboard/TN-Overview.png" width="700">
</p>
</details>

<details>
<summary><strong>Keyword Trends</strong></summary>
<p>Way to see trends of keywords being used.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/TaskNote%20Activity%20Dashboard/TN-KeywordTrends.png" width="700">
</p>
</details>

<details>
<summary><strong>Completed KPI's</strong></summary>
<p>Completion ratio of tasks</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/TaskNote%20Activity%20Dashboard/TN-CompletionKPIs.png" width="700">
</p>
</details>

<details>
<summary><strong>Analytics</strong></summary>
<p>Analytics of Tasks and Notes</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/TaskNote%20Activity%20Dashboard/TN-Analytics.png" width="700">
</p>
</details>

<details>
<summary><strong>Team Workload</strong></summary>
<p>Way to see each team member's workload.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/TaskNote%20Activity%20Dashboard/TN-TeamWorkload.png" width="700">
</p>
</details>

### ✅ [Weekly Attendance](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Weekly%20Attendance)
This report is still being built, but it is Week at a Glance (WAAG) 2.0 for at least us. The most significant changes are that it compares previous periods and introduces new concepts for tracking metrics with various comparisons.  One of the concepts for WAAG 2.0 is enrollment vs attendance metrics for our connect groups (Sunday School).  This allows us to track the need for outreach or inreach for each ministry area.  Another one is tracking total attendance for either the configured fiscal year or calendar year to the previous period, which helps determine as a whole if you are are up/down.

- ⚙️ **Implementation Level: Easy-Moderate
- 🧩 **Installation: Most items of this script works, but is still being developed and might contain a few bugs.  Installation is fairly easy, but with the amount of options it can make it seem overwhelming.  

<details>
<summary><strong>KPI's</strong></summary>
<p>Main KPI's showing overall summaries.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Weekly%20Attendance/WCMainKPIs.png" width="700">
</p>
</details>

<details>
<summary><strong>Enrollment Metrics</strong></summary>
<p>Enrollment and attendance go hand-in-hand, with this showing where the group as a whole and each division stands.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Weekly%20Attendance/WCEnrollmentMetrics.png" width="700">
</p>
</details>

<details>
<summary><strong>Program Summary</strong></summary>
<p>Summary of all programs configured for the report and how they stand overall in comparison to previous time periods.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Weekly%20Attendance/WCProgramSummary.png" width="700">
</p>
</details>

<details>
<summary><strong>Program Breakdown</strong></summary>
<p>Breakdown on how each program and it's respective divisions is doing.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Weekly%20Attendance/WCConnectGroupAttendance.png" width="700">
</p>
</details>

---
## 🔧 Programs

### 🌍 [Missions Dashbaord](https://github.com/bswaby/Touchpoint/tree/main/Missions/MissionsDashboard)
The mission's program is a tool that helps the mission pastor and mission leaders oversee the trips. This tool has many features. Mission pastors can easily manage outstanding payments, background checks, passports, upcoming meetings, see who the leaders are, and more. Mission leaders can easily see their team, contact information, emergency contact information, payment status of each person, and training resources.  Their access to this is dynamically given through the mission widget that lives in their dashboard.

- ⚙️ **Implementation Level: Advanced
- 🧩 **Installation:  This is highly configured for our environment.  It is possible to use it, but several hard-coded areas throughout the script will need to be considered.  I am evaluating whether I can make this easier to implement in other environments, but it will be awhile before I add any focus back on it.

<details>
<summary><strong>Dashboard</strong></summary>
<p>Gives missions pastor a single painted glass for all his active mission trips.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/Missions/MissionsDashboard/MissionsDashboard.png" width="700">
</p>
</details>

<details>
<summary><strong>Leaders Page</strong></summary>
<p>Leaders have a single page to look at all the materials, status on team payments, verification of data, and easy way to contact.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/Missions/MissionsDashboard/MissionsLeaderPage.png" width="700">
</p>
</details>

---
## 🛠️ Tools

### 🔗 [Attachment Link Downloader](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Attachment%20Link%20Generator)
Simple download tool allows you to download all documents from an involvement registration easily. 

- ⚙️ **Implementation Level: Easy
- 🧩 **Installation: This is a paste-and-go Python script with no configuration needed.  

<details>
<summary><strong>Download attachments</strong></summary>
<p>Download attachments in bulk, individually, or selected.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Attachment%20Link%20Generator/ALG-Download.png" width="700">
</p>
</details>

### 🔗 [Account Security Monitor](https://github.com/bswaby/Touchpoint/blob/main/TPxi/Account%20Security%20Monitor/TPxi_AccountSecurityMonitor.py)
This Python dashboard provides advanced security analytics, automated threat detection, and actionable intelligence for protecting your church's digital infrastructure.

🛡️ Key Features

- ⚙️ **Implementation Level: Easy
- 🧩 **Installation: This is a paste-and-go Python script with no configuration needed.  

<details>
<summary><strong>🎯 Dashboard</strong></summary>
- Real-time KPI Dashboard - 7-day security overview with critical metrics
- Executive Summary - High-level security posture at a glance
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Account%20Security%20Monitor/Sec-Dashboard.png" width="700">
</p>
</details>

<details>
<summary><strong>🔍 High-Risk User Monitor</strong></summary>
- Privileged Account Protection - Enhanced monitoring for Admin, Finance, Developer, and API accounts
- Role-Based Risk Analysis - Automated classification by user privileges
- Targeted Attack Detection - Identifies attacks specifically targeting high-value accounts
- Account Compromise Detection - Behavioral analysis for potential breaches
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Account%20Security%20Monitor/Sec-HighSecurity1.png" width="700">
</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Account%20Security%20Monitor/Sec-HighSecurity2.png" width="700">
</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Account%20Security%20Monitor/Sec-HighSecurity3.png" width="700">
</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Account%20Security%20Monitor/Sec-HighSecurity4.png" width="700">
</p>
</details>

<details>
<summary><strong>📊 Enhanced Security Dashboard</strong></summary>
- Advanced Threat Intelligence - Pattern recognition and attack classification
- IP Reputation Analysis - Geographic tracking and threat attribution
- Automated Risk Scoring - Dynamic threat level assessment
- Compliance Reporting - Audit-ready security documentation
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Account%20Security%20Monitor/Sec-Enhanced1.png" width="700">
</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Account%20Security%20Monitor/Sec-Enhanced2.png" width="700">
</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Account%20Security%20Monitor/Sec-Enhanced3.png" width="700">
</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Account%20Security%20Monitor/Sec-Enhanced4.png" width="700">
</p>
</details>

### 🚗 [FastLaneCheckIn](https://github.com/bswaby/Touchpoint/tree/main/TPxi/FastLaneCheckIn)
This streamlined check-in interface is designed for large events (100–2500+ attendees). The core goals are:

1️⃣ Be fast for large crowds
2️⃣ Always work towards <b>zero</b>
3️⃣ Communicate clearly with both staff and attendees.

The app is intentionally minimal and is optimized for use by individuals or teams handling alphabetically segmented lines (A–F, G–L, etc). Users can select one or more Involvements with an active meeting for the day.

2️⃣ 🆕 Recent Updates
- 6/5/2025: Added Email Notification feature. When a person checks in:
  - If 18+, an email is sent directly to them.
  - If under 18, an email is sent to their parent/guardian.
  - Uses existing system email templates (must include CheckedIn in the name).
- 6/7/2025:
  - Feature 1: Ability to show specific Groups during check-in (e.g. Bus assignments, Cabin groups, Subgroups). Useful for quick communication.
  - Feature 2: Added an Outstanding Payment flag during check-in to clearly indicate unpaid balances.
  - Feature 3: Added alphabet grouping options in the UI so it no longer needs to be a configuration change.


3️⃣ Badge Limitation Note

Note: This will <b>not</b> print badges, as those modules from TP are not exposed.

4️⃣ Implementation & Installation Section
- ⚙️ **Implementation Level: Easy
- 🧩 **Installation: This is a paste-and-go Python script with just a couple of configuration items

<details>
<summary><strong>Select Groups</strong></summary>
<p>Filter by program, select email notification if needed, and then 1 or more groups.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/FastLaneCheckIn/FLC-SelectGroups.png" width="700">
</p>
</details>

<details>
<summary><strong>Check-In</strong></summary>
<p>Click Check in to quickly check-in each person.  Notice that alpha limitation and searching are also option when you segment group lines for multiple check-in clerks.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/FastLaneCheckIn/FLC-CheckIn.png" width="700">
</p>
</details>

<details>
<summary><strong>Check-Out</strong></summary>
<p>Switch to the checked in page and correct any mistakes.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/FastLaneCheckIn/FLC-CheckOut.png" width="700">
</p>
</details>

### 🗺️ [Geographic Distribution Map](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Geographic%20Distribution%20Map)
This blue toolbar visualizes the geographic spread of members or activities in Google Maps. It is similar to Touchpoint's map tool but has three advantages over Touchpoint. 

Advantages
1. You can overlay demographic data from the US Census.
2. You can click on the image to get info on the person, similar to how James Kurtz is doing it.
3. You can draw around a city, neighborhood, etc., and tag or export those people for the selected area

- ⚙️ **Implementation Level: Easy-Moderate
- 🧩 **Installation: This is an easy script to implement. The most complicated parts are getting Google Maps and Census API keys.

<details>
<summary><strong>Tag Areas</strong></summary>
<p>Ability to draw around a parameter and tag the results.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Geographic%20Distribution%20Map/GDM-Tag.png" width="700">
</p>
</details>

<details>
<summary><strong>Census Demographc Overlay</strong></summary>
<p>Overlay US census data.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Geographic%20Distribution%20Map/GDM-Demographics.png" width="700">
</p>
</details>

<details>
<summary><strong>Pop-Up</strong></summary>
<p>Click an icon to see info about the person.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Geographic%20Distribution%20Map/GDM-PopUp.png" width="700">
</p>
</details>

### 🛂 [Involvement Owner Audit](https://github.com/bswaby/Touchpoint/blob/main/Python%20Scripts/TechStatus/TechStatus)
Provides a quick method to see owners of involvements to help with auditing or a leader/staff changeover. 

- ⚙️ **Implementation Level: Easy
- 🧩 **Installation: This is a paste-and-go Python script with no configuration needed.  

<details>
<summary><strong>Involvement Owner Audit Tool</strong></summary>
<p><i>You can search by all, involvement, or person.</i></p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Involvement%20Notification%20Audit%20Tool/InvAuditTool.png" width="700">
</p>
</details>

### ↻ [Involvement Sync](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Involvement%20Sync)
This provides an easy method to sync a primary involvement to a secondary. It can be used in cases where scheduler involvement doesn't allow non-scheduled people to check in or to keep two involvements synced, in which they need slightly different print settings. For example, our worship team needs badges on all days except Sunday, and remembering to turn off badge printing keeps getting forgotten.  

- ⚙️ **Implementation Level: Easy
- 🧩 **Installation: This is a paste-and-go Python script, with the only configuration needed to schedule it to run.

<details>
<summary><strong>Involvement Sync Main Screen</strong></summary>
<p><i>Setup a new sync, edit and existing sync, or manually sync</i></p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Involvement%20Sync/InvolvementSync.png" width="700">
</p>
</details>

### 🔐 [Link Generator](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Link%20Generator)
This admin-only tool helps you create a pre-authenticated link that will "automagically" log a person into a direct link. This can help troubleshoot or get a person to where you need them.

- ⚙️ **Implementation Level: Easy
- 🧩 **Installation: This is a paste-and-go Python script with no configuration needed.  

<details>
<summary><strong>Person Selection</strong></summary>
<p>Quickly find and select a person.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Link%20Generator/LG-SelectPerson.png" width="700">
</p>
</details>

<details>
<summary><strong>Quick Selection</strong></summary>
<p>Choose from one of the many predefined selections.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Link%20Generator/LG-QuickSelection.png" width="700">
</p>
</details>

<details>
<summary><strong>Custom URL</strong></summary>
<p>Create your own link by using the custom URL option.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Link%20Generator/LG-CustomURL.png" width="700">
</p>
</details>

### 🔐 [Regisration Export](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Registration%20Export)
Export registrations from multiple or single invovlements.

- ⚙️ **Implementation Level: Easy
- 🧩 **Installation: This is a paste-and-go Python script with no configuration needed.  

<details>
<summary><strong>Registration Selection or Typed-In</strong></summary>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Registration%20Export/RE-Selection.png" width="700">
</p>
</details>

<details>
<summary><strong>Data Export</strong></summary>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Registration%20Export/RE-DataExport.png" width="700">
</p>
</details>

### 🖥️ [TechStatus](https://github.com/bswaby/Touchpoint/blob/main/Python%20Scripts/TechStatus/TechStatus)
This is a quick way to get failed prints, today's print stats, login failures, and more. If I get a question about whether we are having an issue with something like printing, I will pull this up to see what printed and what might be stuck in the queue.  

- ⚙️ **Implementation Level: Easy
- 🧩 **Installation: This is a paste-and-go Python script with no configuration needed.  

<details>
<summary><strong>Print Jobs</strong></summary>
<p>Show counts and print issues.  <i>Note:  TP clears print stats nightly.</i></p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/Python%20Scripts/TechStatus/TS-PrintJobs.png" width="700">
</p>
</details>

<details>
<summary><strong>Login Activity</strong></summary>
<p>Get additional information about login activity problems.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/Python%20Scripts/TechStatus/TS-LoginActivity.png" width="700">
</p>
</details>

<details>
<summary><strong>Script Activity</strong></summary>
<p>See what scripts are firing off.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/Python%20Scripts/TechStatus/TS-ScriptActivity.png" width="700">
</p>
</details>

<details>
<summary><strong>User Accounts</strong></summary>
<p>View, filter, and sort user accounts</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/Python%20Scripts/TechStatus/TS-UserAccounts.png" width="700">
</p>
</details>

### 🗺️ [User Activity](https://github.com/bswaby/Touchpoint/tree/main/TPxi/User%20Activity)
This is to help get an idea of user activity levels based on the activity log table.   You can even search a single person with this and see how much they worked at the church, home, mobile and an estimated amount of time they worked in the system.

- ⚙️ **Implementation Level: Easy
- 🧩 **Installation: This is a paste-and-go Python script with only one configuration option that you might want to consider: OFFICE_IP_ADDRESSES configuration option determines if people are working in the office or outside the office if you configure it with your network ip addresses.  

<details>
<summary><strong>Overview/strong></summary>
<p>Gives an overview of the number of changes in the database and who are the top active users.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/User%20Activity/UA-Overview.png" width="700">
</p>
</details>

<details>
<summary><strong>User Activity</strong></summary>
<p>Gives and overview of the level of activity, estimated time they are spending, and what areas they are working on.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/User%20Activity/UA-UserActivity.png" width="700">
</p>
</details>

<details>
<summary><strong>Stale Accounts</strong></summary>
<p>Accounts that have seen a level of inactivity.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/User%20Activity/UA-StaleAccounts.png" width="700">
</p>
</details>

<details>
<summary><strong>Password Resets</strong></summary>
<p>Report of accounts that have recently completed a password reset.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/User%20Activity/UA-PasswordResets.png" width="700">
</p>
</details>


## 🧩 Widgets

### ⚡ [Widget QuickLinks](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Widget%20QuickLinks)
This provides quick-access links for users to get to things quickly.  The links are permission-based and only show links if they have permission to it.  You can categorize your links into groups to help give focus to the link types.  Lastly, there is an advanced feature to add a count on top of the icon to give some level of need for the link.   

- ⚙️ **Implementation Level: Moderate-Advanced
- 🧩 **Installation: While the overall implementation is not overly complicated, using the count overlay and understanding how some of the configuration parameters go in can take moderate to advanced knowledge. 

<details>
<summary><strong>Example</strong></summary>
<p>The tool implementation is intermediate level, but not overly difficult</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Widget%20QuickLinks/WidgetQuickLink.png" width="700">
</p>
</details>

---
## 🔍 Getting Started

### ⚙️ Prerequisites
- Active Touchpoint account with administrative access
- Basic understanding of how to copy, paste, and set variables for customization
- Advanced understanding of HTML/CSS/Javascript/SQL to customize the interface

### 📥 Installation
Most of the code snippets have a few variables up top to configure.  From there, it's just a matter of copying and pasting it as a new script under Admin ~ Advanced ~ Special Content. 

---
## 👥 Contributing
Contributions are welcome! If you have ideas for improvements or have created your tools that might benefit others, please:
1. Fork the repository
2. Create a new branch for your feature
3. Submit a pull request with a clear description of your changes

---
## 📞 Support
If you need assistance implementing any of these tools, please:
- Open an issue in this repository
- Contact me through [contact information]

---
## 🙏 Acknowledgments
- Thanks to the Touchpoint team for creating an extensible platform
- Special thanks to all the churches who have tested and provided feedback on these tools
