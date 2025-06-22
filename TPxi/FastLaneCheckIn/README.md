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


<summary><strong>Select Groups</strong></summary>
<p>Filter by program, select email notification if needed, and then 1 or more groups.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/FastLaneCheckIn/FLC-SelectGroups.png" width="700">
</p>

<summary><strong>Check-In</strong></summary>
<p>Click Check in to quickly check-in each person.  Notice that alpha limitation and searching are also option when you segment group lines for multiple check-in clerks.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/FastLaneCheckIn/FLC-CheckIn.png" width="700">
</p>

<summary><strong>Check-Out</strong></summary>
<p>Switch to the checked in page and correct any mistakes.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/FastLaneCheckIn/FLC-CheckOut.png" width="700">
</p>

