### ✅ [Task Runner](https://github.com/bswaby/Touchpoint/tree/main/TPxi/Task%20Runner)
An individual-first task triage view for staff who live in TouchPoint Tasks all day.  Instead of a workload chart, it answers the question *"what do I do RIGHT NOW?"* — open tasks are grouped by urgency, the top of the page tells you where to start, and one-click actions handle completion, hide-from-view, log a contact attempt, or reassign without leaving the page.  A separate Team View lets directors / staff leads check on their people without losing the personal queue underneath.

- ⚙️ **Implementation Level:** Easy → Medium (paste-and-go for personal use; Team View needs a Staff involvement configured in Settings)
- 🧩 **Installation:** Paste the script into Special Content > Python.  Add to CustomReports so staff can launch it.  All configuration lives in the in-app **⚙️ Settings** panel — no code editing required.
- 🔄 **Auto-Update:** Built in.  When a newer version is published the script will offer a one-click Update Now banner at the top of the page.

---

<summary><strong>My View — urgency buckets + Daily Focus</strong></summary>
<p>Open tasks grouped by Overdue / Today / This Week / Later / Undated.  The "Where to start" card at the top picks 3 things based on age + due date so you never have to decide what to do first.  The capacity line tells you how full your queue is and roughly how long it would take to clear at your recent completion rate.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Task%20Runner/TR-MyView.png" width="700">
</p>

<summary><strong>Per-row quick actions</strong></summary>
<p>Each row has one-click actions:</p>

- ✅ **Complete** — closes the task with an optional note
- 👁️ **Hide** — locally triages a task you don't want cluttering your view (1 day / 1 week / forever).  This is a per-user filter only; the task itself is untouched, so other people aren't affected.
- 📞 **Log Contact** — record a call / text / email / visit / meeting against the about-person.  Configurable contact methods (admin-set in Settings), and each row shows the last contact attempt with a code chip + days-since.
- 🔁 **Reassign** — search by first/last/partial name, hand the task to someone else.  Uses TouchPoint's TaskNoteMassAssign
- 👤 **Open person** — the about-person name in the row is a direct link to their profile.

<summary><strong>Team View — permission-gated drill-in</strong></summary>
<p>Admin-enabled view showing every teammate's queue at a glance: open / overdue / completed-today, plus when they last completed something.  Click a teammate to drill into their list and (if permissioned) reassign or complete on their behalf.</p>
<p align="center">
  <img src="https://github.com/bswaby/Touchpoint/raw/main/TPxi/Task%20Runner/TR-TeamView.png" width="700">
</p>

The toolbar above the team list has a **"Hide teammates with no open tasks"** toggle and sticky per user so leaders looking only for stuck queues can collapse the empty ones.

<summary><strong>Settings — admin-controlled, no code edits</strong></summary>
<p>All policy lives in <strong>⚙️ Settings</strong>.  Two storage slots (org-wide + per-user) so individual prefs don't trample org policy.</p>

| Setting | Scope | Description |
|---------|-------|-------------|
| **Contact methods** | Org | List of method codes/labels (e.g. `P` Phone, `E` Email, `T` Text, `V` Visit, `M` Meeting).  Empty by default — admin adds what fits the church.  A "Use suggested starters" button drops in a sane default set. |
| **Contact lookback (days)** | Org | How far back to scan for "last contact" indicators on each row.  Default 182 days. |
| **Staff involvement** | Org | The TouchPoint involvement that defines who counts as "staff" for Team View.  Type-ahead search. |
| **Team drill-in scope** | Org | `Off` (no Team View) / `Subgroup` (only people who share a subgroup with you) / `All staff` (any staff member). |
| **Team drill-in actions** | Org | `None` (read-only) / `Reassign only` / `Reassign + Complete`.  Belt-and-suspenders: the server re-checks this on every action — UI hiding alone is not the gate. |
| **Group team list by subgroup** | Org | Bucket teammates under their subgroup (department) headers for easier scanning of bigger teams. |
| **Hide empty teammates** | Per-user | Sticky local toggle — leaders can filter zero-task teammates out of Team View. |

<summary><strong>Multi-tenant & permission-aware</strong></summary>
<p>Designed to be safe to roll out to staff:</p>

- **Auth on drill-in.**  Even with the Team View toggle showing, the server rechecks that the drilled-in PeopleId is actually in the caller's team before returning their tasks.
- **Per-action permissions.**  Reassign-only vs Reassign-and-Complete is gated server-side; a user without permission gets a refusal even if they hand-craft an AJAX call.

<summary><strong>Why "Runner"?</strong></summary>
<p>The point of this tool isn't to admire the workload, it's to <em>run the queue down</em>.  Daily Focus tells you where to start, urgency buckets keep overdue work from drowning under a pile of "Later", Hide kills the noise of stale tickets that aren't actionable today, and Log Contact + Reassign close out a row in seconds.  The team layer is there for the leader who needs to know who's stuck, not who's busy.</p>
