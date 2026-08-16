ServiceDate

Tracks the last time each person did something that shows they are engaged, and stores it as a single date Extra Value you can query from Search Builder. Answers the question "who has gone quiet?" without anyone building a report.

A person's ServiceDate is the most recent date they did any one of three things:





Gave a contribution



Served in a serving or volunteer involvement



Went on a mission trip

The newest of those three wins. Because it is stored as a real date rather than a number of days, it never goes stale between runs, and days-since is calculated on the fly wherever you need it:

DATEDIFF(day, ServiceDate, GETDATE())

With this tool, admins can:





Populate the ServiceDate Extra Value for the whole database in batches that will not time out



Keep it current automatically from the morning batch, with no scheduled-time assumptions



Give staff ready-made Search Builder scripts for 90, 180, and 365 day windows



Configure what counts as serving and giving from a settings modal, with no code edits



Verify the configuration against this install's own lookup tables before the first run



Credit a spouse for household giving



See at a glance whether the morning batch is actually wired up, and fix it with one click



Clean up the legacy integer ServiceDays Extra Value if upgrading from an older version





Dashboard

The dashboard shows what is configured, whether the daily run is installed, and how many records need work. Preload populates the Extra Value for the first time, Force Full Rescan recomputes everything from scratch after a settings change, and both run in batches.







Search Builder Scripts

The Extra Value is only useful if staff can query it. The dashboard checks for a saved SQL script per time window and creates any that are missing. Staff then pick them from Search Builder > In SQL List with no SQL knowledge required.







Script



Returns





ServiceDate90



Everyone who gave, served, or went on a mission trip in the last 90 days





ServiceDate180



Same, within 180 days





ServiceDate365



Same, within 365 days

Scripts that already exist and match are left alone. Scripts you have edited yourself are detected and not overwritten unless you explicitly choose to replace them.







Two Run Modes

Morning batch timing drifts, gets skipped, and stops when a server reboots, so the script never assumes it ran yesterday. It records the timestamp of every successful run and reaches back to that point next time.







Mode



When it runs



What it does





Full



First run, after a settings change, or when the last full pass is older than FULL_RESCAN_EVERY_DAYS



Walks everyone in the lookback window. The only mode that can clear a value that has aged out





Incremental



Every other daily run



Looks only at activity since the last successful run, minus a small overlap. Forward-only, and never clears

Nothing is missed when a run is late or skipped for a week, and the daily read drops from the whole lookback window to just what changed.





Configuration

Everything is editable from the Settings modal on the dashboard and stored in Special Content, so upgrading the script does not lose your setup.







Setting



Default



Purpose





SERVE_ORG_TYPE_IDS



[145]



Organization types that count as serving. The one that most often needs changing





INCLUDE_MISSION_TRIPS



True



Count mission trips, using the built-in IsMissionTrip flag





EXCLUDE_CONTRIBUTION_TYPE_IDS



[6, 7, 8, 99]



Contribution types that are not gifts: returned, reversed, pledge, event fee





CONTRIBUTION_STATUS_ID



0



Only count posted contributions





CREDIT_SPOUSE_FOR_GIVING



True



Give a spouse credit for household giving





SCOPE_DAYS



730



How far back a full rescan looks





EV_NAME



ServiceDate



The Extra Value this maintains





SQL_LIST_WINDOWS



[90, 180, 365]



Day windows to publish as Search Builder scripts





INCREMENTAL_OVERLAP_DAYS



2



Overlap re-examined on every incremental run, covering late-entered activity





FULL_RESCAN_EVERY_DAYS



7



Force a full pass when the last one is older than this





BATCH_TRIGGER



run_servicedays



Name of the Data flag your morning batch sets





BATCH_CONTENT_NAME



MorningBatch



Which Special Content script runs this daily

The Setup check panel reads this install's own lookup.OrganizationType and lookup.ContributionType and prints your ids next to their descriptions. Worth doing before the first run: a wrong serving type fails silently, and the script simply never credits anyone for serving.





Setup





Navigate to Admin > Advanced > Special Content > Python



Click New Python Script File



Name it TPxi_ServiceDays



Paste the contents of TPxi_ServiceDays.py



Open /PyScriptForm/TPxi_ServiceDays and work through the Setup check panel



Press Create Search Builder Scripts to publish the In SQL List queries



Press Preload to populate the Extra Value for everyone



Press Add to Morning Batch so it stays current

Optionally add to CustomReports XML:

<Report name="TPxi_ServiceDays" type="PyScript" role="Admin" />





Morning Batch

The Add to Morning Batch button writes this into your MorningBatch script for you, between managed markers so it can be removed again cleanly:

try:
    Data.run_servicedays = "true"
    model.CallScript("TPxi_ServiceDays")
except Exception as e:
    print "ServiceDate batch error: " + str(e)

Two details that matter:

It must be MorningBatch, not ScheduledTasks. Those are different features. A trigger placed in ScheduledTasks looks installed and never runs daily, so ServiceDate quietly goes stale. The dashboard detects a block in the wrong place and offers to clean it up.

The try/except is not decoration. The morning batch runs several scripts in one pass, and an unhandled error in any of them can stop the ones after it. Wrapped, a bad day for ServiceDate stays a bad day for ServiceDate.

The script detects its own name from the URL, so renaming it does not break the registration.





Legacy Cleanup

Earlier versions stored an integer ServiceDays Extra Value, which went stale the moment it was written. If the dashboard finds any, it offers a batched cleanup to remove them. The date-based value replaces it entirely.



Like this tool? DisplayCache integrates directly with TouchPoint and helps fund continued development of tools like this one.
