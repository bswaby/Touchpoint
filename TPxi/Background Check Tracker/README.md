# TPxi Background Check Tracker

A focused utility for triaging in-flight background checks. Pulls TouchPoint's
cached status AND the live MinistrySafe status side-by-side so you can spot
sync drift fast.  Solves the common *"MinistrySafe says Pending but TouchPoint
says Overdue"* problem.

Extracted from the Compliance Dashboard's Pending BG tab and rebuilt as a
standalone, single-purpose tool with on-demand MinistrySafe API lookups.

---

## What it does

- **Lists every pending background check** from TouchPoint's cache
- **Shows the live MinistrySafe status** in a dedicated column
- **Auto-fetches live status** on page load (throttled, cached 5 minutes)
- **Highlights sync drift** — rows where TouchPoint says Pending but MinistrySafe says complete/cancelled/failed get a red left border so they stand out
- **Per-row Refresh button** — force a live MS API check, bypassing the cache
- **Per-row Resend Invite button** — appears when status is *Awaiting Applicant Action*, hits MS to re-send the applicant's invite email
- **Direct links** to the TouchPoint person profile and the MS report page
- **Settings tab** with step-by-step instructions for getting your MinistrySafe API token

## What it does NOT do

Does not write back to TouchPoint's `BackgroundChecks` table. The TP cache is
updated by TouchPoint's normal webhook / morning batch path. This utility
gives you visibility while you wait for that to catch up — it doesn't try
to fight TouchPoint over who owns that data.

---

## Installation

1. In TouchPoint: **Admin → Advanced → Special Content → Python → New Python Script File**
2. Name it `TPxi_BGCheckTracker`
3. Paste the entire contents of `TPxi_BGCheckTracker.py` and Save
4. *(Optional)* Add to **CustomReports** for menu access:
   ```xml
   <Report name="TPxi_BGCheckTracker" type="PyScript" role="Edit" />
   ```
   *CustomReport changes can take up to 24 hours to appear due to caching.*
5. Run the script (URL: `/PyScriptForm/TPxi_BGCheckTracker`)
6. Switch to the **Settings** tab and paste your MinistrySafe API token (see below)

---

## Getting your MinistrySafe API token

The Settings tab inside the script has the same instructions inline, but
for reference:

1. Sign in to [safetysystem.ministrysafe.com](https://safetysystem.ministrysafe.com) as an admin user
2. Click your name (top right) and choose **Account Settings**
3. Scroll to **API Access** (lower section of the page)
4. If a token already exists, click **Reveal** and copy it. Otherwise click **Generate API Token** first
5. Paste the token into the field on the Settings tab and click **Save Token**
6. Click **Test Connection** to confirm the token works

**Don't see API Access?** Your MinistrySafe account may not have API
permission enabled. Email `support@ministrysafe.com` to request it.
They enable it on most paid plans for free.

The token is stored in TouchPoint's `Setting` table under
`MinistrySafeApiToken`. It is never echoed back in any AJAX response
or log line — only a masked last-4-digit preview is shown on the
Settings tab.

### What the token can do

| Action | Used by |
|---|---|
| Read background check status (live, per applicant) | Refresh button + Refresh All |
| Resend the applicant invite | Resend Invite button (only when *Awaiting Applicant Action*) |

That's it. The token does NOT have write access to TouchPoint and
cannot modify any background check records.

---

## Usage

### Triaging a stuck check

The common scenario: a volunteer's background check shows "Overdue" in
TouchPoint but the staff member knows it was resubmitted recently.

1. Open the script (Pending BG Checks tab)
2. Find the person in the list (they'll be there because TouchPoint cache
   still reads `Pending`)
3. Look at the **Live MS Status** column for that row:
   - If it says *Processing at MinistrySafe* → MS is working on it. TouchPoint will catch up via webhook.
   - If it says *Awaiting Applicant Action* → the applicant hasn't completed their portion. Click **Resend Invite** to nudge them.
   - If it says *Awaiting Church Review* → MS finished and it's waiting on church staff to review/approve.
   - If the row is highlighted red → TouchPoint says Pending but MS says complete or failed. There's a sync gap; check the MS dashboard directly.

### Refreshing data

- **Per-row Refresh** (circular arrow icon) → forces a live MS lookup, bypassing the 5-minute cache
- **Refresh All Live Status** → batches a live lookup for every row, throttled 200ms between calls to be nice to MS
- **Reload List** → re-runs the underlying SQL query against TouchPoint, picks up any newly-pending checks

### Status legend

| Label | MS raw | Meaning |
|---|---|---|
| **Awaiting Church Review** | `complete` | MS finished. Action needed from church staff |
| **Processing at MinistrySafe** | `processing` | MS is still working |
| **Awaiting Applicant Action** | `awaiting_applicant` | Applicant hasn't completed their portion |
| **Cancelled** | `cancelled` | Check was cancelled |
| **Expired** | `expired` | Check expired before completion |
| **MS: <other>** | anything else | Verbatim MS response (lets you see new statuses they add) |
| **Not yet checked** | empty | We haven't called MS for this row yet |

---

## Troubleshooting

### Live MS Status column shows "API token not configured"

Switch to the Settings tab and paste your MinistrySafe API token.

### "Token rejected (HTTP 401)" on Test Connection

The token is wrong or has been revoked. Generate a new one in MinistrySafe
(steps above) and paste it on the Settings tab.

### Row shows "Error Sending - Check TouchPoint"

The TouchPoint record has no `ReportID` — meaning the original send to MS
never returned a confirmation. The row exists in TouchPoint's cache but
MS has no record. You'll need to re-send the background check request
from the person's TouchPoint profile.

### Refresh All is slow

Each MS API call takes ~100-300ms and we throttle 200ms between to avoid
hammering them. With 50 pending rows that's ~15-25 seconds. The page
stays interactive throughout; the button just shows a progress spinner.


---

## Architecture notes (for future maintainers)

### Cache strategy

Live MS responses are cached in `TPxi_BGCheckTracker_Cache` (Special
Content blob) keyed by `report_id` with a `ts` timestamp. Reads check
freshness (5-minute TTL) before deciding whether to hit MS. The per-row
Refresh button bypasses the cache (`force_refresh=True`); Refresh All
also bypasses so a click always gets fresh data.

The cache survives across users but per-`report_id` isolation means
two admins refreshing at the same time don't conflict.

### Auto-Update

Wired in via the standard `TPxi/AutoUpdate/README.md` pattern. Banner
appears at the top when DisplayCache publishes a newer version; one
click rewrites the installed PythonContent slot. Your saved settings
(API token) and cache are preserved across updates.

---

## Version

Current version is set in `APP_VERSION` near the top of
`TPxi_BGCheckTracker.py`. Bump on every release published to the
DisplayCache manifest at `scripts.displaycache.com/api/touchpoint/script-versions`.

## Author

Ben Swaby — bswaby@fbchtn.org

First Baptist Church Hendersonville
