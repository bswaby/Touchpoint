## 📋 Overview

This repository is the work of one guy who loves Jesus, loves clean data, and loves seeing people come to Christ — so I built a bunch of tools to help make ministry smoother, more informed, reduce overhead, and help reach others.
---

## 💰 Finance

### 💳 [FortisFees](https://github.com/bswaby/Touchpoint/blob/main/Finance/FortisFees)
It breaks down fees by program and accounting code so accounting can use it to back charge. The biggest note is that the fees will be close to 99% but not 100% due to the disconnect of certain charges and reversals getting back to Touchpoint.  

### 📜 [QCD-GrantLetters](https://github.com/bswaby/Touchpoint/blob/main/Finance/QCD-GrantLetters)
Automatically create QCD and Grant letters that you can print out and add to each person's notes.

<details>
<summary><strong>Date Selection</strong></summary>
<p>The date selection interface allows users to choose dates from a calendar view easily.  It also allow to upload a note to each persons record.</p>
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

### 📊 [Weekly Contribution Report](https://github.com/bswaby/Touchpoint/blob/main/Finance/Weekly%20Contribution%20Report)
This is the primary tool our finance team tracks, reports, and works through finances each week.

## 📈 Reports

### 📅 [Anniversaries](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Anniversaries)
It is designed to be a widget on your dashboard that tracks member anniversaries (marriage, work, and birthdays, primarily, but it can be expanded as well).

### 📱 [Communication Dashboard](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Communication%20Dashboard)
Centralized dashboard for viewing and analyzing communications.

### 🧹 [Data Quality Dashboard](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Data%20Quality%20Dashboard)
Monitors the completeness of database demographics.

### 🧑‍🤝‍🧑 [Involvement Activity Dashboard](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Involvement%20Activity%20Dashboard)
It's like Kenny Rogers in that it helps you know "when to hold them" and "when to fold them" when it comes to your Involvements.

### 🏛️ [Ministry Structure](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Ministry%20Structure)
Displays the hierarchy of program, division, organization, and organization type.

### 📤 [Registration Export](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Registration%20Export)
Exports event or class registration data in structured formats.

### 📝 [TaskNote Activity Dashboard](https://github.com/bswaby/Touchpoint/tree/main/TPxi/TaskNote%20Activity%20Dashboard)
Dashboard to start to understand TaskNote activity, who has open assignments, keyword trends, and more.

### ✅ [Weekly Attendance](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Weekly%20Attendance)
This builds on Week at a Glance (WAAG) but introduces many new concepts for tracking metrics with various comparisons.  

## 🔧 Programs

### 🌍 [Missions Dashbaord](https://github.com/bswaby/Touchpoint/tree/main/Missions/MissionsDashboard)
Mission's program is a tool to help the mission pastor and mission leaders oversee the trips. This tool has many features. Mission pastors can easily manage outstanding payments, background checks, passports, upcoming meetings, see who the leaders are, and more. Mission leaders can easily see their team, contact information, emergency contact information, payment status of each person, and training resources.  Their access to this is dynamically given through the mission widget that lives in their dashboard.

Note:  This is highly configured for our environment.  It is possible to use it, but time will need to be put in to configure it for another environment.

## 🛠️ Tools

### 🖥️ [TechStatus](https://github.com/bswaby/Touchpoint/blob/main/Python%20Scripts/TechStatus/TechStatus)
A quick way to get fail prints, today's print stats, login failures, and more

### 🔗 [Attachment Link Generator](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Attachment%20Link%20Generator)
A tool to make it easier to download documents uploaded to an Involvement

### 🚗 [FastLaneCheckIn](https://github.com/bswaby/Touchpoint/tree/main/TPxi/FastLaneCheckIn)
This is a streamlined check-in interface for large (100-2500+) events. It is made for a person to hold their phone/iPad/tablet to get people checked in quickly. Note: this doesn't print as those modules are not exposed for use.

### 🗺️ [Geographic Distribution Map](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Geographic%20Distribution%20Map)
This is a blue toolbar that visualizes the geographic spread of members or activities in google maps. It similar to Touchpoint's map tool, but has three advantages over Touchpoint. First, you can overlay demographic data from the census. Second, you can click on the image to get info on the person similar to how James Kurtz is doing it, and Third, you can draw around a city, neighborhood, etc.. and tag and/or export to CSV those within that selection area.

### 🔐 [Link Generator](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Link%20Generator)
This is an admin-only tool that helps you create a pre-authenticated link that will automagically log a person into a direct link. This can help with troubleshooting or getting a person to where you need them.

## 🧩 Widget 

### ⚡ [Widget QuickLinks](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Widget%20QuickLinks)
Quick-access widgets for frequently used tools and dashboards. These have permission, counts, and categories for displaying access links.

<details>
<summary><strong>QCD Letter</strong></summary>
<p>The tool implementation is intermediate level, but not overly difficult</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/Finance/screenshots/WidgetQuickLink.png" width="700">
</p>
</details>


## 🔍 Getting Started

### ⚙️ Prerequisites
- Active Touchpoint account with administrative access
- Basic understanding of how to copy, paste, and set variables for customization
- Advanced understanding of HTML/CSS/Javascript/SQL to customize the interface

### 📥 Installation
Most of the code snippets have a few variables up top to configure.  From there, it's just a matter of copying and pasting it as a new script under Admin ~ Advanced ~ Special Content. 

## 👥 Contributing
Contributions are welcome! If you have ideas for improvements or have created your tools that might benefit others, please:
1. Fork the repository
2. Create a new branch for your feature
3. Submit a pull request with a clear description of your changes

## 📞 Support
If you need assistance implementing any of these tools, please:
- Open an issue in this repository
- Contact me through [contact information]

## 🙏 Acknowledgments
- Thanks to the Touchpoint team for creating an extensible platform
- Special thanks to all the churches who have tested and provided feedback on these tools
