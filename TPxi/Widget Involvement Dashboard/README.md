### 📊 [Involvement Dashboard Widget](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Widget%20Involvement%20Dashboard)
The Involvement Dashboard provides church leaders a comprehensive overview of all involvements within a specific program and division. It displays basic real-time attendance metrics, meeting schedules, last meeting dates, and identifies organizations that need attention based on meeting frequency.

- ⚙️ Implementation Level: Easy  
- 🧩 **Installation**: Minimal configuration required. Edit the PROGRAM_ID and DIVISION_ID variables to match your specific program/division to monitor. Optionally: Adjust DAYS_FOR_AVERAGE (default 90 days) for attendance calculations and set DEBUG_ROLE for troubleshooting access.

<summary><strong>Dashboard Features</strong></summary>

**Key Metrics Displayed:**
- Total involvements and member count
- Involvements needing updates (14+ days since last meeting)
- Average attendance across all Involvements
- Individual org schedules with day/time formatting

**Status Indicators:**
- 🟢 **Current**: Recent meetings (within 7 days)
- 🔵 **Recent**: Meetings 7-14 days ago  
- 🟡 **Needs Update**: Meetings 14-30 days ago
- 🔴 **Overdue**: No meetings in 30+ days
- ⚫ **No Meetings**: Organizations without recorded meetings

  <summary><strong>Main Screen</strong></summary>
    <p align="center">
      <img src="https://github.com/bswaby/Touchpoint/blob/main/TPxi/Widget%20Involvement%20Dashboard/WID-Main.png" width="700">
    </p>
