### 📋 [Day of Registration](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Day%20of%20Registration)
A multi station tablet friendly tool for handling day of event registration at scale, such as VBS or summer camps. Multiple volunteers staff iPads simultaneously, each running their own station (Kindergarten, 1st Grade, etc.) and assigning incoming families to specific room involvements with live capacity tracking.

- ⚙️ **Implementation Level:** Easy to Moderate
- 🧩 **Installation:** Single script. Paste into Special Content > Python, navigate to `/PyScript/TPxi_DayOfRegistration`. Configure scenarios via Admin Setup.

<summary><strong>Features</strong></summary>

- **Two Modes:** Admin Setup for configuration, Station Mode for the iPad facing volunteer experience
- **Multiple Concurrent Stations:** Several iPads can run different stations at the same time, syncing every 10 seconds through TouchPoint
- **Source Involvement Per Station:** Each station pulls its waiting list from a pre registration involvement (e.g. the K-5 master roster)
- **Destinations With Live Capacity:** Each destination room has its own count and capacity. Soft caps warn the volunteer; hard caps block assignment
- **Subgroup Assignment:** Add the registrant to one or more subgroups within the destination involvement at the moment of assignment (Bus, Group color, etc.)
- **Pinned Registration Questions:** Pick which registration questions appear front and center in the volunteer info bar (allergies, grade, friend request) without having to click into a panel
- **Bulk Add Destinations:** Add multiple destination rooms at once with a default capacity, especially useful for setups with 20+ rooms
- **Find a Friend:** Search by name across the entire program (or all of TouchPoint) and see exactly which room a registrant has already been placed in
- **Walk in Support:** Search anyone in TouchPoint and assign them on the spot, even if they were not pre registered
- **Subgroup Filter on Display:** Color coded destination cards with sort options (A to Z, Most available, Most full) so volunteers can find rooms quickly
- **Find Friend Filter:** Restrict friend search to a specific program or division so unrelated involvements stay out of results
- **Undo Support:** Reverse a recent assignment with a single click, audit logged
- **Test Mode With Safe Reset:** Toggle Test Mode on a scenario, then reset only the people the script assigned during testing. Backs up first; pre existing room rosters are never touched
- **Audit Log:** Every assignment, undo, and reset writes to a per scenario log that can be inspected or cleared
- **Auto Update:** Checks DisplayCache for new versions and prompts to update in place. Saved scenarios and logs are preserved
- **Cross Browser:** Works on iPad Safari, desktop Chrome, and any modern tablet browser

<hr>

<summary><strong>Admin: Edit Scenarios</strong></summary>
<p>Top level scenario list. Each scenario configures one event (VBS K-5, Camp, etc.) with its own stations and destinations.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Day%20of%20Registration/DoR-EditScenarios.png" width="700">
</p>

<summary><strong>Admin: Edit Station</strong></summary>
<p>Configure each station with a source involvement, pinned questions, and destinations. Bulk Add quickly adds many rooms at once with a shared default capacity.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Day%20of%20Registration/DoR-EditStation.png" width="700">
</p>

<summary><strong>Station: Volunteer Name and Start</strong></summary>
<p>Volunteer enters their name (logged with each assignment for accountability) and picks the station they are running.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Day%20of%20Registration/DoR-AssignmentNameandStartRegistering.png" width="700">
</p>

<summary><strong>Station: Selected Person</strong></summary>
<p>Tap a registrant to see their info bar. Pinned registration answers (allergies, grade, friend request) appear immediately. Family members, emergency contact, OTC meds, and Authorized Checkout pull automatically from TouchPoint.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Day%20of%20Registration/DoR-SelectedPerson.png" width="700">
</p>

<summary><strong>Station: Select Destination</strong></summary>
<p>Right panel shows every room for this station with live capacity bars. Sort by alphabetical, most available, or most full. Color coded for at a glance recognition.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Day%20of%20Registration/DoR-SelectDestination.png" width="700">
</p>

<summary><strong>Station: Assign</strong></summary>
<p>One tap assignment with optional subgroup selection (e.g. which bus). The system blocks hard capped rooms automatically and warns on soft caps.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Day%20of%20Registration/DoR-Assign.png" width="700">
</p>

<summary><strong>Station: Find a Friend</strong></summary>
<p>Quick search across all program participants to see exactly which room a friend or sibling has been assigned to, so families can stay together.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Day%20of%20Registration/DoR-FindFriend.png" width="700">
</p>

<summary><strong>Station: Switch Station</strong></summary>
<p>Volunteers can hop between stations during the event without restarting (useful when a station gets backed up).</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Day%20of%20Registration/DoR-SwitchStation.png" width="700">
</p>

<hr>

<summary><strong>Installation</strong></summary>

1. **Admin > Advanced > Special Content > Python**
2. Click **Add New**, name it `TPxi_DayOfRegistration`
3. Paste the script and Save
4. Navigate to `/PyScript/TPxi_DayOfRegistration`
5. Click **Admin Setup** to create your first scenario

<summary><strong>Quick Setup</strong></summary>

1. **Create a Scenario.** Name it (e.g. "VBS K-5") and pick a Friend Finder program/division if you want friend search scoped to your event
2. **Add Stations.** One per group (e.g. "Kindergarten", "1st Grade"). Each station gets its own source involvement (the master pre registration list) and its own list of destination rooms
3. **Pin Registration Questions.** Click Load Questions on the source involvement to see every question, then check the ones volunteers should see at a glance (allergies, grade, etc.)
4. **Add Destinations.** Use Bulk Add to load a dozen rooms at once with a default capacity. Adjust individual rooms with the pencil icon if needed
5. **Open Station Mode** on each iPad. Volunteers enter their name, pick their station, and start checking people in
6. **Monitor From Anywhere.** Counts and assignments sync across devices every 10 seconds

<summary><strong>Test Mode</strong></summary>

Before going live, enable **Test Mode** on a scenario to unlock the **Reset Test Assignments** button. This lets you assign sample people during practice runs and then wipe just those test assignments without touching the actual room rosters or any pre existing members. The reset is gated by a name confirmation, lists every affected person before executing, and saves a backup snapshot to Special Content first. Disable Test Mode before the event starts.

The always available **Clear Assignment Log** button only empties the audit log; it never touches anyone's involvement membership.

<summary><strong>Storage Keys</strong></summary>

The script uses these Special Content > Text entries:

| Key | Purpose |
|-----|---------|
| `DayOfRegistration_Scenarios` | All scenario configs (stations, destinations, capacities) |
| `DayOfRegistration_Log_<scenarioId>` | Per scenario assignment audit log |
| `DayOfRegistration_Backup_<scenarioId>_<timestamp>` | Snapshot saved before Reset Test Assignments runs |

---
*Written by [Ben Swaby](https://github.com/bswaby). These tools are free because they should be. If they've saved you time, consider [DisplayCache](https://displaycache.com) Church digital signage that integrates with TouchPoint or [TPxi Go](https://tpxigo.com) your church contacts in Outlook and on your phone.*
