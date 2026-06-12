#####################################################################
# Blue Toolbar -- Background Check + Training Status
#
# Shows the last Background Check and last Training entry (date +
# status) for each person selected from a Blue Toolbar search.
#
# In TouchPoint, BG checks and training rows BOTH live in the
# dbo.BackgroundChecks table; they are distinguished by ReportTypeID:
#   ReportTypeID = 3   -> Training      (BackgroundCheckCode.ReportTypeTraining)
#   ReportTypeID <> 3  -> Background Check
#
# ApprovalStatus values (lookup.BackGroundCheckApprovalCodes):
#   Pending, Approved, Not Approved
# Some BG vendors (Protect My Ministry, MinistrySafe) may also write
# "Waiting". The status badge handles both.
#
# Install
#   1. Admin > Advanced > Special Content > Python > Add New
#      Name it: BGAndTrainingStatus
#   2. Paste this code and Save.
#   3. Open the Text Content tab > CustomReports, add this line
#      inside the <CustomReports> root element:
#        <Report name="BGAndTrainingStatus" type="PyScript" role="Access" />
#      (CustomReports cache up to 24h before the toolbar entry appears.
#       You can still run the report immediately at
#       /PyScript/BGAndTrainingStatus for testing.)
#   4. From a search, check people > Blue Toolbar >
#      "BG And Training Status".
#
# Customize the role= attribute on the CustomReports line if you want
# to restrict who sees the menu entry.
#####################################################################

model.Header = 'Background Check + Training Status'

# Required: drive the report off the Blue Toolbar selection.
# @BlueToolbarTagId is the temp tag TouchPoint creates when you click
# the toolbar -- it holds the rows the user checked.
sql = """
WITH LatestBG AS (
    SELECT bc.PeopleId,
           bc.ApprovalStatus,
           bc.Updated,
           ROW_NUMBER() OVER (PARTITION BY bc.PeopleId
                              ORDER BY bc.Updated DESC, bc.Id DESC) AS rn
    FROM dbo.BackgroundChecks bc WITH (NOLOCK)
    WHERE ISNULL(bc.ReportTypeId, 0) <> 3      -- everything except training
),
LatestTraining AS (
    SELECT bc.PeopleId,
           bc.ApprovalStatus,
           bc.Updated,
           ROW_NUMBER() OVER (PARTITION BY bc.PeopleId
                              ORDER BY bc.Updated DESC, bc.Id DESC) AS rn
    FROM dbo.BackgroundChecks bc WITH (NOLOCK)
    WHERE bc.ReportTypeId = 3                  -- training only
)
SELECT p.PeopleId,
       p.FirstName,
       p.LastName,
       p.Age,
       bg.Updated         AS BGDate,
       bg.ApprovalStatus  AS BGStatus,
       tr.Updated         AS TrainingDate,
       tr.ApprovalStatus  AS TrainingStatus
FROM dbo.People p
JOIN dbo.TagPerson tp ON tp.PeopleId = p.PeopleId
LEFT JOIN LatestBG       bg ON bg.PeopleId = p.PeopleId AND bg.rn = 1
LEFT JOIN LatestTraining tr ON tr.PeopleId = p.PeopleId AND tr.rn = 1
WHERE tp.Id = @BlueToolbarTagId
ORDER BY p.LastName, p.FirstName
"""

rows = list(q.QuerySql(sql))


def fmt_date(d):
    # SQL DATETIME values come back as .NET System.DateTime in IronPython,
    # which has ToString(format) but no strftime. Try the .NET path first
    # and fall back to Python datetime for safety.
    if not d:
        return ''
    try:
        return d.ToString('MM/dd/yyyy')
    except AttributeError:
        try:
            return d.strftime('%m/%d/%Y')
        except Exception:
            return str(d)[:10]


def status_badge(status):
    # Color-code against the hardwired ApprovalStatus values in
    # lookup.BackGroundCheckApprovalCodes (Pending, Approved,
    # Not Approved). "Waiting" is included because some integrations
    # write that value even though it isn't in the lookup.
    if not status:
        return '<span style="color:#999;">&mdash;</span>'
    s = status.strip().lower()
    color = '#999'
    if s == 'approved':
        color = '#107c10'                      # green
    elif s in ('pending', 'waiting'):
        color = '#f0a000'                      # amber
    elif s == 'not approved':
        color = '#d13438'                      # red
    return '<span style="color:{0};font-weight:600;">{1}</span>'.format(color, status)


print '<style>'
print 'table.bgrep{border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;font-size:13px;}'
print 'table.bgrep th,table.bgrep td{padding:6px 10px;border-bottom:1px solid #e5e5e5;text-align:left;}'
print 'table.bgrep th{background:#2c3e50;color:#fff;}'
print 'table.bgrep tr:nth-child(even){background:#fafafa;}'
print '.muted{color:#999;}'
print '</style>'

print '<h3 style="font-family:Segoe UI,Arial,sans-serif;">Background Check + Training Status ({0} people)</h3>'.format(len(rows))

if not rows:
    print '<p>No people selected.</p>'
else:
    print '<table class="bgrep">'
    print '<thead><tr><th>PeopleId</th><th>First</th><th>Last</th><th>Age</th>'
    print '<th>BG Check</th><th>BG Status</th><th>Training</th><th>Training Status</th></tr></thead><tbody>'
    for r in rows:
        print '<tr>'
        print '<td><a href="/Person2/{0}" target="_blank">{0}</a></td>'.format(r.PeopleId)
        print '<td>{0}</td>'.format(r.FirstName or '')
        print '<td>{0}</td>'.format(r.LastName or '')
        print '<td>{0}</td>'.format(r.Age if r.Age is not None else '')
        print '<td>{0}</td>'.format(fmt_date(r.BGDate) or '<span class="muted">never</span>')
        print '<td>{0}</td>'.format(status_badge(r.BGStatus))
        print '<td>{0}</td>'.format(fmt_date(r.TrainingDate) or '<span class="muted">never</span>')
        print '<td>{0}</td>'.format(status_badge(r.TrainingStatus))
        print '</tr>'
    print '</tbody></table>'
