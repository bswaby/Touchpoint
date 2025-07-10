### 🔍 TouchPoint Live Search

🟩 **Implementation Level: Widget & Standalone** - Deploy as homepage widget or full-page tool

📋 Overview
TPxi_LiveSearch is an advanced real-time search interface for TouchPoint that provides instant results as you type. It combines powerful search capabilities with quick actions for tasks and notes, making it the fastest way to find and interact with church member data.

✨ Features

1. **Real-Time Search Results**
- **Instant Results**: Search begins automatically after 300ms of typing
- **Multi-Category Search**: Simultaneously searches people, organizations, and keywords
- **Smart Ranking**: Most relevant results appear first
- **Visual Indicators**: Icons and colors help identify result types at a glance

2. **Comprehensive Member Views**
- **Journey Timeline**: Visual timeline of member engagement milestones
- **Family Engagement**: See entire family's involvement at a glance
- **Contact Info**: Quick access to phone numbers and email addresses
- **Member Status**: Clear indicators for active members, guests, and prospects

3. **Quick Actions**
- **Add Task**: Create tasks directly from search results
- **Add Note**: Add notes without leaving the search interface
- **View Journey**: One-click access to member engagement history
- **Family View**: Modal popup showing family engagement metrics


- ⚙️ **Implementation Level: Easy
- 🧩 **Installation: This is a paste-and-go Python script with the only configuration needed is to set the SCRIPT_NAME variable correctly due to home page requirement

```python
# Configuration
MAX_RESULTS = 18  # Results per category
SEARCH_DELAY = 300  # Milliseconds delay before search
SCRIPT_NAME = 'TPxi_LiveSearch'  # Must match your upload name
SHOW_GIVING_IN_JOURNEY = False  # Enable/disable giving in timeline
```

🖼️ Screenshots

Main Search Interface
![Live Search Main Screen](screenshots/livesearch-main.png)

Journey Timeline View
![Journey Timeline](screenshots/livesearch-journey.png)

Family Engagement Modal
![Family Engagement](screenshots/livesearch-family.png)
