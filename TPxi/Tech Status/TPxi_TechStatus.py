#--------------------------------------------------------------------
# TPxi Technical Status
# Tool to help troubleshoot basic issues: print queue, kiosk printing,
# failed logins and script activity.
#
# Written By: Ben Swaby (TPxi Software, LLC)
# Email: bswaby@fbchtn.org
# Website: https://tpxisoftware.com
# GitHub: https://github.com/bswaby/Touchpoint  (50+ free tools)
# ----------------------------------------------------------------
# These tools are free because they should be.
# If they've saved you time or helped your team, and you want to
# support continued development, check out:
#
# DisplayCache(TM) - church digital signage that integrates with TouchPoint(R)
# https://displaycache.com
#
# TPxi Go(TM) - your church contacts, wherever you work.
# Look up anyone in TouchPoint(R), log calls and emails from Outlook
# or your phone. No tab switching, no lost context.
# https://tpxigo.com
# ----------------------------------------------------------------
#
# Version history
#   2.0.0  2026-11-15  Auto-update: Admins see Update now / Later when a newer
#                      version is published; the download sends this site's address, the
#                      script and its version (no person; TPxiSendHost = 0
#                      turns it off). Print Queue Details rebuilt (below).
#   2.0.1  2026-11-15  Opens on Overview; #prints, #login, #scripts or #users in
#                      the address opens that tab (it used to reopen the last
#                      tab viewed).
#   2.0.2  2026-11-15  TouchPoint's plain page title is hidden; the dashboard's
#                      own header is the only title.
#
# Print Queue Details (2026-11-15): each waiting job now shows what is in it,
# flags label designs with no fields (they can never print and jam the print
# service), and shows the printer's label size setting when TouchPoint has it
# (who and which class for Admins, the label settings and designs) instead
# of a link to a page that does not exist. A job waiting 2 minutes or more
# is "Stuck". One stuck job can hold up every station on the same print
# service; scripts cannot clear it (see the note on the Print Queue tab).
#--------------------------------------------------------------------
####REPORT INFORMATION####
#--------------------------------------------------------------------
#
#Add all this code to a single Python Script
#1. Navigate to Admin ~ Advanced ~ Special Content ~ Python Scripts
#2. Select New Python Script File and Name the File
#3. Paste in all this code and Run
#Optional: Add to navigation menu

#--------------------------------------------------------------------
####USER CONFIG FIELDS
#--------------------------------------------------------------------
model.Header = ''          # the page draws its own title; TouchPoint's is hidden in the CSS

# --- Version / Auto-update -------------------------------------------
# Admins see a banner when a newer version is published, and Update now
# installs it in place.
import json
import re
APP_VERSION = '2.0.2'                              # bump on each release published
DC_SCRIPT_ID = 'TPxi_TechnicalStatus'
# scripts.displaycache.com is the public domain for browser-side version checks.
# workers.dev is used for server-side fetches (bypasses Cloudflare Bot Fight Mode).
DC_API_BASE = 'https://scripts.displaycache.com/api/touchpoint'
DC_API_WORKER = 'https://touchpoint-scripts.bswaby.workers.dev/api/touchpoint'


def get_script_name():
    """The name this script is installed under (an admin may have renamed it):
    the posted script_name, then the URL, then the default."""
    try:
        if hasattr(Data, 'script_name') and Data.script_name:
            sn = str(Data.script_name).strip()
            if sn:
                return sn
    except:
        pass
    try:
        url = str(getattr(model, 'URL', '') or '')
        m = re.search(r'/PyScript(?:Form)?/([^/?#&]+)', url)
        if m:
            return m.group(1)
    except:
        pass
    return DC_SCRIPT_ID


def dc_headers():
    """Sent with the update download: this church's TouchPoint address and the
    script and version, so the DisplayCache traffic page can say who is on
    what. No person, nothing from the database. The admin setting
    TPxiSendHost = 0 turns it off (shared with the other TPxi scripts)."""
    h = {}
    try:
        if str(model.Setting("TPxiSendHost", "1")).strip() in ("0", "false", "False"):
            return h
    except Exception:
        pass
    try:
        host = str(model.CmsHost or "").strip()
    except Exception:
        host = ""
    if host:
        h["X-CmsHost"] = host
    h["X-TPxi-Script"] = DC_SCRIPT_ID
    h["X-TPxi-Version"] = APP_VERSION
    return h


def _ts_is_admin():
    try:
        return model.UserIsInRole('Admin')
    except:
        return False


def apply_update():
    """Fetch the published code, check it really is this script, write it
    over the installed copy, and confirm the new version landed. Admins only:
    WriteContentPython itself checks no role."""
    if not _ts_is_admin():
        return {'success': False, 'message': 'Only Admins can update this script.'}
    target_name = re.sub(r'[^\w.\- ]', '', get_script_name() or '') or DC_SCRIPT_ID
    fetch_url = DC_API_WORKER + '/scripts/' + DC_SCRIPT_ID
    try:
        new_code = model.RestGet(fetch_url, dc_headers())
    except Exception as fe:
        return {'success': False, 'message': "Couldn't reach the update server. " + str(fe) + ' | Manual install: ' + fetch_url}
    new_code = str(new_code or '')
    if len(new_code) < 500:
        return {'success': False, 'message': 'The update server sent ' + str(len(new_code)) +
                ' bytes, too small to be the script. Try again in a minute, or use the manual link: ' + fetch_url}
    m = re.search(r"APP_VERSION = '([0-9.]+)'", new_code)
    if 'TPxi Technical Status' not in new_code or not m:
        return {'success': False, 'message': "The update didn't look like Technical Status code; the update server may be "
                'returning an error page. Manual install: ' + fetch_url}
    new_version = m.group(1)
    try:
        model.WriteContentPython(target_name, new_code)
    except Exception as we:
        return {'success': False, 'message': 'Write failed: ' + str(we) + ' | Paste the latest code from ' + fetch_url +
                ' into Admin > Advanced > Special Content > Python > ' + target_name + '.'}
    try:
        written = str(model.PythonContent(target_name) or '')
        if ("APP_VERSION = '" + new_version + "'") not in written:
            return {'success': False, 'message': "Wrote to '" + target_name + "' but it does not show v" + new_version +
                    '. The script is probably installed under another name. Find it under Admin > Advanced > Special Content > Python '
                    'and paste the latest code from ' + fetch_url + ' into it.'}
    except:
        pass   # the read-back is best effort; the write itself succeeded
    return {'success': True, 'message': 'Updated ' + target_name + ' from v' + APP_VERSION + ' to v' + new_version + '. Reloading...'}


_ts_ajax_done = False
try:
    _ts_action = str(Data.action) if hasattr(Data, 'action') and Data.action else ''
except:
    _ts_action = ''
if model.HttpMethod == 'post' and _ts_action == 'apply_update':
    try:
        print(json.dumps(apply_update()))
    except Exception as e:
        print(json.dumps({'success': False, 'message': 'Update failed: ' + str(e)}))
    _ts_ajax_done = True

#--------------------------------------------------------------------
####START OF CODE.  No configuration should be needed beyond this point
#--------------------------------------------------------------------

# Print job statistics by hour
sqlKioskPrints = '''
SELECT DATEPART(HOUR, Stamp) AS [Hour], 
       COUNT(*) AS [TotalPrints] 
FROM PrintJob 
GROUP BY DATEPART(HOUR, Stamp)
ORDER BY [Hour];
'''

# Print jobs still in queue. A blank Id means the label already printed.
# Each row says what is in the job: the label settings and designs, and
# (for Admins only) who and which class. Bad JSON never breaks the page.
sqlPrintsInQueue = """
WITH jobs AS (
  SELECT pj.Id, pj.Stamp,
         DATEDIFF(minute, pj.Stamp, GETDATE()) AgeMin,
         CASE WHEN DATEDIFF(minute, pj.Stamp, GETDATE()) >= 2 THEN 1 ELSE 0 END IsStuck,
         CASE WHEN pj.Id LIKE 'PTDRY%' THEN 1 ELSE 0 END IsPractice,
         CASE WHEN pj.JsonData IS NULL OR ISJSON(pj.JsonData) = 0 THEN 1 ELSE 0 END IsOldFormat,
         {who} Who,
         (SELECT STRING_AGG(CAST(k.[key] COLLATE DATABASE_DEFAULT + ': ' + ISNULL(l.name, CONCAT('#', k.value COLLATE DATABASE_DEFAULT))
                     + CASE WHEN l.active = 0 THEN ' (inactive)' ELSE '' END
                     + CASE WHEN l.id IS NOT NULL AND fc.n = 0 THEN ' (NO FIELDS)' ELSE '' END AS varchar(max)), '; ')
            FROM OPENJSON(j.J, '$.kioskLabels') k
            LEFT JOIN CheckInLabel l ON l.id = TRY_CAST(k.value AS int)
            OUTER APPLY (SELECT COUNT(*) n FROM CheckInLabelEntry e WHERE e.labelID = l.id) fc) Designs,
         (SELECT STRING_AGG(CAST(k.[key] COLLATE DATABASE_DEFAULT + ': ' + l.name AS varchar(max)), '; ')
            FROM OPENJSON(j.J, '$.kioskLabels') k
            JOIN CheckInLabel l ON l.id = TRY_CAST(k.value AS int)
            WHERE NOT EXISTS (SELECT 1 FROM CheckInLabelEntry e WHERE e.labelID = l.id)) EmptyDesigns,
         (SELECT TOP 1 cp.Id + '/' + ISNULL(JSON_VALUE(cp.Data, '$.userName'), '')
            FROM CheckInPrinters cp WITH (NOLOCK)
            CROSS APPLY OPENJSON(CASE WHEN ISJSON(cp.Data) = 1 THEN cp.Data ELSE '{}' END, '$.printers') WITH (kioskName nvarchar(200)) pr
            WHERE REPLACE(pr.kioskName, ' ', '') COLLATE DATABASE_DEFAULT = pj.Id COLLATE DATABASE_DEFAULT) Service,
         (SELECT TOP 1 CONCAT(cp.Id, ISNULL(' (' + JSON_VALUE(cp.Data, '$.userName') + ')', ''),
                       ', label size setting ', ISNULL(CAST(pr.labelSizeIndex AS varchar(10)), 'not set'),
                       ', print service last signed in ', CONVERT(varchar(16), cp.UpdatedDate, 120))
            FROM CheckInPrinters cp WITH (NOLOCK)
            CROSS APPLY OPENJSON(CASE WHEN ISJSON(cp.Data) = 1 THEN cp.Data ELSE '{}' END, '$.printers')
                 WITH (kioskName nvarchar(200), labelSizeIndex int) pr
            WHERE REPLACE(pr.kioskName, ' ', '') COLLATE DATABASE_DEFAULT = pj.Id COLLATE DATABASE_DEFAULT) Printer,
         CONCAT('Security labels: ',
                CASE JSON_VALUE(j.J, '$.securityLabels') WHEN '0' THEN 'none' WHEN '1' THEN 'one per meeting'
                     WHEN '2' THEN 'one per child' WHEN '3' THEN 'one per family' ELSE 'not set' END,
                '. Guest labels: ', ISNULL(JSON_VALUE(j.J, '$.guestLabels'), 'not set'),
                '. Location labels: ', CASE JSON_VALUE(j.J, '$.locationLabels') WHEN '0' THEN 'never' WHEN '1' THEN 'default'
                     WHEN '2' THEN 'always' ELSE 'not set' END,
                '. QR code labels: ', ISNULL(JSON_VALUE(j.J, '$.printQrCodeLabels'), 'not set'), '.') Settings
  FROM PrintJob pj WITH (NOLOCK)
  CROSS APPLY (SELECT CASE WHEN ISJSON(pj.JsonData) = 1 THEN pj.JsonData ELSE '{}' END J) j
  WHERE LEN(LTRIM(RTRIM(ISNULL(pj.Id, '')))) > 0
)
SELECT a.*,
       CASE WHEN a.EmptyDesigns IS NOT NULL THEN 1 ELSE 0 END CannotPrint,
       b.BlockerAt, b.BlockerStation,
       CASE WHEN a.EmptyDesigns IS NULL AND b.BlockerAt IS NOT NULL THEN 1 ELSE 0 END HeldUp,
       (SELECT STRING_AGG(CAST(cp.Id + ISNULL(' (' + JSON_VALUE(cp.Data, '$.userName') + ')', '') AS varchar(max)), ', ')
          FROM CheckInPrinters cp WITH (NOLOCK) WHERE ISNULL(cp.Id, '') <> '') ListedServices
FROM jobs a
OUTER APPLY (
  SELECT TOP 1 CONCAT((DATEPART(hour, x.Stamp) + 11) % 12 + 1, ':', RIGHT('0' + CAST(DATEPART(minute, x.Stamp) AS varchar(2)), 2),
                      ':', RIGHT('0' + CAST(DATEPART(second, x.Stamp) AS varchar(2)), 2),
                      CASE WHEN DATEPART(hour, x.Stamp) < 12 THEN ' AM' ELSE ' PM' END) BlockerAt, x.Id BlockerStation
  FROM jobs x
  WHERE x.EmptyDesigns IS NOT NULL AND x.Stamp <= GETDATE()
    AND (x.Id = a.Id OR (x.Service IS NOT NULL AND x.Service = a.Service))
  ORDER BY x.Stamp
) b
ORDER BY a.Id, a.Stamp
"""
# Who and which class: Admins only (children's names).
sqlPrintJobWho = """(SELECT STRING_AGG(CAST(ISNULL(p.Name, CONCAT('#', a.peopleID)) + ISNULL(' (' + o.OrganizationName + ')', '') AS varchar(max)), '; ')
          FROM OPENJSON(j.J, '$.attendances') WITH (peopleID int, groups nvarchar(max) AS JSON) a
          OUTER APPLY OPENJSON(a.groups) WITH (groupID int) g
          LEFT JOIN People p ON p.PeopleId = a.peopleID
          LEFT JOIN Organizations o ON o.OrganizationId = g.groupID)"""
try:
    _canSeeWho = model.UserIsInRole('Admin')
except Exception:
    _canSeeWho = False
sqlPrintsInQueue = sqlPrintsInQueue.replace('{who}', sqlPrintJobWho if _canSeeWho else 'NULL')

# Failed login statistics from the last 72 hours - LIMITED to top 10
sqlFailedLoginStats = '''
SELECT TOP 10
    COUNT(Activity) AS [FailedLogins],
    Activity,
    MAX(ActivityDate) AS [LastFailedAttempt]
FROM dbo.ActivityLog 
WHERE 
    (Activity LIKE '%Failed password%' 
        OR Activity LIKE '%Invalid log%'
        OR Activity LIKE '%ForgotPassword%')
    AND CAST(ActivityDate AS Date) >= CAST(DATEADD(DAY, -3, CONVERT(DATE, GETDATE())) AS Date)
GROUP BY Activity
ORDER BY COUNT(Activity) DESC;
'''

# Last 200 failed login attempts
sqlFailedLogins = '''
SELECT 
    TOP 200
    ActivityDate,
    UserId,
    Activity,
    PeopleId,
    OrgId,
    ClientIp
FROM 
    dbo.ActivityLog 
WHERE 
    Activity LIKE '%ForgotPassword%'
    OR Activity LIKE '%failed password%'
    OR Activity LIKE '%Invalid log%'
ORDER BY ActivityDate DESC
'''

# Last 50 successful logins
sqlLogins = '''
SELECT 
    TOP 50
    ActivityDate,
    UserId,
    Activity,
    PeopleId,
    OrgId,
    ClientIp
FROM 
    dbo.ActivityLog 
WHERE 
    Activity LIKE '%logged in%'
ORDER BY ActivityDate DESC
'''

# All script executions from last 7 days
sqlScriptActivity = '''
SELECT 
    ActivityDate,
    UserId,
    Activity,
    PeopleId,
    OrgId,
    ClientIp
FROM 
    dbo.ActivityLog 
WHERE 
    Activity LIKE '%script%'
    AND ActivityDate >= DATEADD(DAY, -7, GETDATE())
ORDER BY ActivityDate DESC
'''

# User accounts query - Only execute if requested
sqlUserAccounts = '''
SELECT 
    DISTINCT u.Name, 
    u.Username, 
    u.CreationDate, 
    u.LastLoginDate, 
    u.LastActivityDate, 
    u.EmailAddress, 
    u.MFAEnabled, 
    u.MustChangePassword, 
    u.IsLockedOut,
    u.PeopleId,
    u.UserId
FROM 
    ActivityLog al
INNER JOIN Users u ON (al.UserId = u.UserId) 
LEFT JOIN UserList ul ON ul.UserId = u.UserId
ORDER BY u.Name ASC;
'''

# Execute queries for data that's always needed
Data.kioskprints = q.QuerySql(sqlKioskPrints)
Data.printsinqueue = q.QuerySql(sqlPrintsInQueue)
Data.failedloginstat = q.QuerySql(sqlFailedLoginStats)
Data.failedlogins = q.QuerySql(sqlFailedLogins)
Data.logins = q.QuerySql(sqlLogins)
Data.script = q.QuerySql(sqlScriptActivity)

# Count rows for display
Data.printsinqueueCount = len(list(Data.printsinqueue))
Data.scriptCount = len(list(Data.script))

# Calculate total failed logins
totalFailedLogins = 0
for stat in Data.failedloginstat:
    totalFailedLogins += stat.FailedLogins
Data.totalFailedLogins = totalFailedLogins

# Get login failure trends by hour for the last 24 hours
sqlFailureTrends = '''
SELECT 
    DATEPART(HOUR, ActivityDate) AS Hour,
    COUNT(*) AS FailureCount
FROM dbo.ActivityLog 
WHERE 
    (Activity LIKE '%Failed password%' 
        OR Activity LIKE '%Invalid log%'
        OR Activity LIKE '%ForgotPassword%')
    AND ActivityDate >= DATEADD(HOUR, -24, GETDATE())
GROUP BY DATEPART(HOUR, ActivityDate)
ORDER BY Hour
'''
Data.failureTrends = q.QuerySql(sqlFailureTrends)

# Get script execution statistics
sqlScriptStats = '''
SELECT 
    COUNT(DISTINCT UserId) AS UniqueUsers,
    COUNT(*) AS TotalExecutions,
    COUNT(DISTINCT CAST(ActivityDate AS DATE)) AS ActiveDays
FROM dbo.ActivityLog 
WHERE 
    Activity LIKE '%script%'
    AND ActivityDate >= DATEADD(DAY, -7, GETDATE())
'''
scriptStats = q.QuerySqlTop1(sqlScriptStats)
Data.uniqueScriptUsers = scriptStats.UniqueUsers if scriptStats else 0
Data.totalScriptExecutions = scriptStats.TotalExecutions if scriptStats else 0
Data.activeScriptDays = scriptStats.ActiveDays if scriptStats else 0

# Security Analytics - Multiple logins from different IPs (72 hours)
sqlSecurityAnalytics = '''
WITH LoginActivity AS (
    SELECT 
        UserId,
        ClientIp,
        COUNT(*) AS LoginCount,
        MIN(ActivityDate) AS FirstLogin,
        MAX(ActivityDate) AS LastLogin
    FROM dbo.ActivityLog
    WHERE 
        Activity LIKE '%logged in%'
        AND ActivityDate >= DATEADD(HOUR, -72, GETDATE())
    GROUP BY UserId, ClientIp
),
SuspiciousAccounts AS (
    SELECT 
        UserId,
        COUNT(DISTINCT ClientIp) AS UniqueIPs,
        SUM(LoginCount) AS TotalLogins
    FROM LoginActivity
    GROUP BY UserId
    HAVING COUNT(DISTINCT ClientIp) > 1 OR SUM(LoginCount) > 5
)
SELECT TOP 20
    sa.UserId,
    sa.UniqueIPs,
    sa.TotalLogins,
    u.Name AS UserName,
    u.EmailAddress,
    STUFF((
        SELECT ', ' + ClientIp + ' (' + CAST(LoginCount AS VARCHAR) + ')'
        FROM LoginActivity la
        WHERE la.UserId = sa.UserId
        ORDER BY LoginCount DESC
        FOR XML PATH('')
    ), 1, 2, '') AS IPDetails
FROM SuspiciousAccounts sa
LEFT JOIN Users u ON sa.UserId = u.UserId
ORDER BY sa.UniqueIPs DESC, sa.TotalLogins DESC
'''

Data.securityAnalytics = q.QuerySql(sqlSecurityAnalytics)

# Get login success/failure ratio for donut chart
sqlLoginRatio = '''
SELECT 
    SUM(CASE WHEN Activity LIKE '%logged in%' THEN 1 ELSE 0 END) AS SuccessCount,
    SUM(CASE WHEN Activity LIKE '%Failed password%' OR Activity LIKE '%Invalid log%' OR Activity LIKE '%ForgotPassword%' THEN 1 ELSE 0 END) AS FailureCount
FROM dbo.ActivityLog
WHERE ActivityDate >= DATEADD(HOUR, -24, GETDATE())
    AND (Activity LIKE '%logged in%' 
        OR Activity LIKE '%Failed password%' 
        OR Activity LIKE '%Invalid log%' 
        OR Activity LIKE '%ForgotPassword%')
'''
loginRatio = q.QuerySqlTop1(sqlLoginRatio)
Data.successCount = loginRatio.SuccessCount if loginRatio else 0
Data.failureCount = loginRatio.FailureCount if loginRatio else 0

# Check if user account data was requested
loadUserAccounts = getattr(model.Data, 'loadUsers', '') == "true"

# Only execute user account query if requested
if loadUserAccounts:
    Data.useraccounts = q.QuerySql(sqlUserAccounts)
else:
    Data.useraccounts = [] # Empty placeholder for template

# ===========================================
# HTML Template with Modern Design
# ===========================================
template = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>TechStatus Dashboard</title>
    <!-- Font Awesome for icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- Google Fonts -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <!-- AG Grid for User Accounts -->
    <script src="https://cdn.jsdelivr.net/npm/ag-grid-enterprise/dist/ag-grid-enterprise.js"></script>
    <style>
        .page-header{display:none!important}
        :root {
            --primary: #4F46E5;
            --primary-dark: #4338CA;
            --primary-light: #6366F1;
            --secondary: #1E293B;
            --secondary-light: #334155;
            --success: #10B981;
            --danger: #EF4444;
            --warning: #F59E0B;
            --info: #3B82F6;
            --light: #F8FAFC;
            --gray-50: #F9FAFB;
            --gray-100: #F3F4F6;
            --gray-200: #E5E7EB;
            --gray-300: #D1D5DB;
            --gray-400: #9CA3AF;
            --gray-500: #6B7280;
            --gray-600: #4B5563;
            --gray-700: #374151;
            --gray-800: #1F2937;
            --gray-900: #111827;
            --border-radius: 12px;
            --border-radius-sm: 8px;
            --box-shadow: 0 1px 3px rgba(0,0,0,0.12), 0 1px 2px rgba(0,0,0,0.06);
            --box-shadow-md: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -1px rgba(0,0,0,0.06);
            --box-shadow-lg: 0 10px 15px -3px rgba(0,0,0,0.1), 0 4px 6px -2px rgba(0,0,0,0.05);
            --transition: all 0.2s ease;
        }
        
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            line-height: 1.6;
            color: var(--gray-900);
            background-color: var(--gray-50);
            font-size: 14px;
        }
        
        /* Dashboard Container */
        .dashboard {
            max-width: 1400px;
            margin: 0 auto;
            padding: 24px;
        }
        
        /* Dashboard Header */
        .dashboard-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 32px;
            background: linear-gradient(135deg, var(--primary) 0%, var(--primary-dark) 100%);
            padding: 24px 32px;
            border-radius: var(--border-radius);
            box-shadow: var(--box-shadow-lg);
            color: white;
        }
        
        .dashboard-title {
            font-size: 32px;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 12px;
        }
        
        .dashboard-title i {
            font-size: 28px;
            opacity: 0.9;
        }
        
        .dashboard-actions {
            display: flex;
            gap: 12px;
            align-items: center;
        }
        
        /* Buttons */
        .btn {
            padding: 10px 20px;
            border: none;
            border-radius: var(--border-radius-sm);
            font-weight: 600;
            cursor: pointer;
            transition: var(--transition);
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            font-size: 14px;
            box-shadow: var(--box-shadow);
        }
        
        .btn-primary {
            background-color: white;
            color: var(--primary);
        }
        
        .btn-primary:hover {
            background-color: var(--gray-100);
            transform: translateY(-1px);
            box-shadow: var(--box-shadow-md);
        }
        
        .btn-secondary {
            background-color: rgba(255,255,255,0.2);
            color: white;
            border: 1px solid rgba(255,255,255,0.3);
        }
        
        .btn-secondary:hover {
            background-color: rgba(255,255,255,0.3);
            transform: translateY(-1px);
        }
        
        /* Cards */
        .card {
            background-color: white;
            border-radius: var(--border-radius);
            box-shadow: var(--box-shadow);
            margin-bottom: 24px;
            overflow: hidden;
            transition: var(--transition);
        }
        
        .card:hover {
            box-shadow: var(--box-shadow-md);
        }
        
        .card-header {
            padding: 20px 24px;
            background: linear-gradient(135deg, var(--secondary) 0%, var(--secondary-light) 100%);
            color: white;
            font-weight: 600;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .card-header h2 {
            font-size: 18px;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .card-header h2 i {
            opacity: 0.8;
            font-size: 20px;
        }
        
        .card-body {
            padding: 24px;
        }
        
        /* Tab Navigation */
        .tab-container {
            margin-bottom: 32px;
            background: white;
            border-radius: var(--border-radius);
            box-shadow: var(--box-shadow);
            padding: 8px;
        }
        
        .tabs {
            display: flex;
            list-style: none;
            gap: 4px;
            flex-wrap: wrap;
        }
        
        .tab {
            padding: 12px 24px;
            cursor: pointer;
            font-weight: 600;
            color: var(--gray-600);
            border-radius: var(--border-radius-sm);
            transition: var(--transition);
            position: relative;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .tab:hover {
            color: var(--primary);
            background-color: var(--gray-50);
        }
        
        .tab.active {
            color: white;
            background: linear-gradient(135deg, var(--primary) 0%, var(--primary-dark) 100%);
            box-shadow: var(--box-shadow);
        }
        
        .tab-content {
            display: none;
            animation: fadeIn 0.3s ease;
        }
        
        .tab-content.active {
            display: block;
        }
        
        @keyframes fadeIn {
            from {
                opacity: 0;
                transform: translateY(10px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
        
        /* Grid Layout */
        .grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(350px, 1fr));
            gap: 24px;
            margin-bottom: 24px;
        }
        
        /* Overview KPI Cards */
        .kpi-card {
            background: white;
            border-radius: var(--border-radius);
            padding: 24px;
            box-shadow: var(--box-shadow);
            transition: var(--transition);
            border-left: 4px solid transparent;
        }
        
        .kpi-card:hover {
            transform: translateY(-2px);
            box-shadow: var(--box-shadow-md);
        }
        
        .kpi-card.success {
            border-left-color: var(--success);
        }
        
        .kpi-card.warning {
            border-left-color: var(--warning);
        }
        
        .kpi-card.danger {
            border-left-color: var(--danger);
        }
        
        .kpi-card.info {
            border-left-color: var(--info);
        }
        
        .kpi-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
        }
        
        .kpi-title {
            font-size: 14px;
            font-weight: 600;
            color: var(--gray-600);
            text-transform: uppercase;
            letter-spacing: 0.05em;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .kpi-icon {
            width: 40px;
            height: 40px;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 20px;
        }
        
        .kpi-icon.success {
            background-color: rgba(16, 185, 129, 0.1);
            color: var(--success);
        }
        
        .kpi-icon.warning {
            background-color: rgba(245, 158, 11, 0.1);
            color: var(--warning);
        }
        
        .kpi-icon.danger {
            background-color: rgba(239, 68, 68, 0.1);
            color: var(--danger);
        }
        
        .kpi-icon.info {
            background-color: rgba(59, 130, 246, 0.1);
            color: var(--info);
        }
        
        .kpi-value {
            font-size: 32px;
            font-weight: 700;
            color: var(--gray-900);
            margin-bottom: 8px;
        }
        
        .kpi-subtitle {
            font-size: 13px;
            color: var(--gray-500);
        }
        
        /* Tables */
        .data-table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            font-size: 14px;
        }
        
        .data-table th,
        .data-table td {
            padding: 16px 20px;
            text-align: left;
            border-bottom: 1px solid var(--gray-200);
        }
        
        .data-table th {
            background-color: var(--gray-50);
            font-weight: 600;
            color: var(--gray-700);
            text-transform: uppercase;
            font-size: 12px;
            letter-spacing: 0.05em;
        }
        
        .data-table tbody tr {
            transition: var(--transition);
        }
        
        .data-table tbody tr:hover {
            background-color: var(--gray-50);
        }
        
        .data-table tbody tr:last-child td {
            border-bottom: none;
        }
        
        .data-table a {
            color: var(--primary);
            text-decoration: none;
            font-weight: 500;
        }
        
        .data-table a:hover {
            color: var(--primary-dark);
            text-decoration: underline;
        }
        
        /* Badges */
        .badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.025em;
        }
        
        .badge-success {
            background-color: #10B981;
            color: white;
        }
        
        .badge-danger {
            background-color: #EF4444;
            color: white;
        }
        
        .badge-warning {
            background-color: #F59E0B;
            color: white;
        }
        
        .badge-info {
            background-color: #3B82F6;
            color: white;
        }
        
        /* Scroll Container */
        .scroll-container {
            max-height: 600px;
            overflow-y: auto;
            border: 1px solid var(--gray-200);
            border-radius: var(--border-radius-sm);
        }
        
        .scroll-container::-webkit-scrollbar {
            width: 8px;
        }
        
        .scroll-container::-webkit-scrollbar-track {
            background: var(--gray-100);
        }
        
        .scroll-container::-webkit-scrollbar-thumb {
            background: var(--gray-400);
            border-radius: 4px;
        }
        
        .scroll-container::-webkit-scrollbar-thumb:hover {
            background: var(--gray-500);
        }
        
        /* Chart Container */
        .chart-container {
            height: 350px;
            margin-bottom: 24px;
            padding: 16px;
            background: var(--gray-50);
            border-radius: var(--border-radius-sm);
        }
        
        /* Loading Indicator */
        .loading-indicator {
            text-align: center;
            padding: 60px;
        }
        
        .loading-spinner {
            border: 4px solid var(--gray-200);
            width: 48px;
            height: 48px;
            border-radius: 50%;
            border-left-color: var(--primary);
            animation: spin 1s linear infinite;
            margin: 0 auto 24px;
        }
        
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        /* Expand Button */
        .expand-button {
            background: none;
            border: none;
            color: white;
            cursor: pointer;
            font-size: 16px;
            padding: 4px 8px;
            border-radius: 4px;
            transition: var(--transition);
            opacity: 0.8;
        }
        
        .expand-button:hover {
            opacity: 1;
            background-color: rgba(255,255,255,0.1);
        }
        
        .expandable-content {
            max-height: 300px;
            overflow-y: auto;
            transition: max-height 0.3s ease;
        }
        
        .expandable-content.expanded {
            max-height: 800px;
        }
        
        /* AG Grid Styling */
        .ag-theme-quartz {
            height: 700px;
            width: 100%;
            margin-bottom: 24px;
        }
        
        /* User Accounts Loading */
        #user-accounts-loading {
            text-align: center;
            padding: 80px 40px;
            background-color: var(--gray-50);
            border-radius: var(--border-radius);
        }
        
        #user-accounts-loading p {
            margin-bottom: 24px;
            font-size: 16px;
            color: var(--gray-600);
        }
        
        #load-users-btn {
            padding: 14px 32px;
            font-size: 16px;
            background: linear-gradient(135deg, var(--primary) 0%, var(--primary-dark) 100%);
            color: white;
        }
        
        #load-users-btn:hover {
            transform: translateY(-2px);
            box-shadow: var(--box-shadow-lg);
        }
        
        
        /* Responsive Design */
        @media (max-width: 1024px) {
            .dashboard {
                padding: 16px;
            }
            
            .grid {
                grid-template-columns: 1fr;
            }
            
            .dashboard-header {
                padding: 20px;
            }
            
            .dashboard-title {
                font-size: 24px;
            }
        }
        
        @media (max-width: 768px) {
            .dashboard-header {
                flex-direction: column;
                align-items: flex-start;
                gap: 16px;
            }
            
            .dashboard-actions {
                width: 100%;
                flex-wrap: wrap;
            }
            
            .tabs {
                overflow-x: auto;
                flex-wrap: nowrap;
                -webkit-overflow-scrolling: touch;
            }
            
            .tab {
                white-space: nowrap;
                font-size: 14px;
                padding: 10px 16px;
            }
            
            .data-table {
                font-size: 13px;
            }
            
            .data-table th,
            .data-table td {
                padding: 12px 16px;
            }
            
            .kpi-value {
                font-size: 28px;
            }
            
            .card-header h2 {
                font-size: 16px;
            }
        }
        
        /* Print Styles */
        @media print {
            .dashboard-actions,
            .tab-container,
            .expand-button,
            .theme-switch {
                display: none !important;
            }
            
            .dashboard {
                max-width: 100%;
            }
            
            .card {
                break-inside: avoid;
                box-shadow: none;
                border: 1px solid var(--gray-300);
            }
            
            .expandable-content {
                max-height: none !important;
            }
        }
        
        /* Empty State */
        .empty-state {
            text-align: center;
            padding: 60px 40px;
            color: var(--gray-500);
        }
        
        .empty-state i {
            font-size: 48px;
            color: var(--gray-400);
            margin-bottom: 16px;
        }
        
        .empty-state p {
            font-size: 16px;
        }
        
    </style>
</head>
<body>
    <div class="dashboard">
        <div id="appUpdateBanner" style="display:none;gap:12px;align-items:center;background:#EAF3FC;border:1px solid #9CC4EC;border-radius:8px;padding:10px 14px;margin-bottom:14px"></div>
        <div class="dashboard-header">
            <h1 class="dashboard-title">
                <i class="fas fa-server"></i> Technical Status Dashboard
            </h1>
            <div class="dashboard-actions">
                <button class="btn btn-secondary" onclick="window.location.reload()">
                    <i class="fas fa-sync-alt"></i> Refresh
                </button>
            </div>
        </div>
        
        <div class="tab-container">
            <ul class="tabs">
                <li class="tab active" data-tab="overview">
                    <i class="fas fa-th-large"></i> Overview
                </li>
                <li class="tab" data-tab="prints">
                    <i class="fas fa-print"></i> Print Jobs
                </li>
                <li class="tab" data-tab="login">
                    <i class="fas fa-user-lock"></i> Login Activity
                </li>
                <li class="tab" data-tab="scripts">
                    <i class="fas fa-code"></i> Script Activity
                </li>
                <li class="tab" data-tab="users">
                    <i class="fas fa-users"></i> User Accounts
                </li>
            </ul>
        </div>
        
        <!-- Overview Tab -->
        <div id="overview" class="tab-content active">
            <div class="grid">
                <!-- Print Queue Status -->
                <div class="kpi-card {{#if printsinqueueCount}}warning{{else}}success{{/if}}">
                    <div class="kpi-header">
                        <div class="kpi-title">
                            <i class="fas fa-print"></i> Print Queue Status
                        </div>
                        <div class="kpi-icon {{#if printsinqueueCount}}warning{{else}}success{{/if}}">
                            <i class="fas {{#if printsinqueueCount}}fa-exclamation-triangle{{else}}fa-check-circle{{/if}}"></i>
                        </div>
                    </div>
                    <div class="kpi-value">{{printsinqueueCount}}</div>
                    <div class="kpi-subtitle">
                        {{#if printsinqueueCount}}
                            <a href="#prints" onclick="activateTab('prints')" style="color: inherit; text-decoration: underline; cursor: pointer;">
                                Jobs pending in queue - View Details
                            </a>
                        {{else}}
                            No jobs in queue
                        {{/if}}
                    </div>
                </div>
                
                <!-- Failed Logins Summary -->
                <div class="kpi-card danger">
                    <div class="kpi-header">
                        <div class="kpi-title">
                            <i class="fas fa-shield-alt"></i> Failed Logins (72hr)
                        </div>
                        <div class="kpi-icon danger">
                            <i class="fas fa-times-circle"></i>
                        </div>
                    </div>
                    <div class="kpi-value">{{totalFailedLogins}}</div>
                    <div class="kpi-subtitle">Total failed login attempts</div>
                </div>
                
                <!-- Recent Script Activity -->
                <div class="kpi-card info">
                    <div class="kpi-header">
                        <div class="kpi-title">
                            <i class="fas fa-terminal"></i> Script Activity
                        </div>
                        <div class="kpi-icon info">
                            <i class="fas fa-code"></i>
                        </div>
                    </div>
                    <div class="kpi-value">{{scriptCount}}</div>
                    <div class="kpi-subtitle">Recent script executions</div>
                </div>
            </div>
            
            <!-- Charts Section for Overview -->
            <div class="grid">
                <div class="card">
                    <div class="card-header">
                        <h2><i class="fas fa-chart-bar"></i> Print Job Activity</h2>
                    </div>
                    <div class="card-body">
                        <div class="chart-container" id="printChartOverview" style="height: 300px;">
                            <!-- Chart will be inserted here by JavaScript -->
                        </div>
                    </div>
                </div>
                
                <div class="card">
                    <div class="card-header">
                        <h2><i class="fas fa-chart-line"></i> Login Failure Trends</h2>
                    </div>
                    <div class="card-body">
                        <div class="chart-container" id="failureTrendsChartOverview" style="height: 300px;">
                            <!-- Chart will be inserted here by JavaScript -->
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Print Jobs Tab -->
        <div id="prints" class="tab-content">
            <div class="card">
                <div class="card-header">
                    <h2><i class="fas fa-print"></i> Print Job Statistics</h2>
                </div>
                <div class="card-body">
                    <div class="chart-container" id="printChart">
                        <!-- Chart will be inserted here by JavaScript -->
                    </div>
                    <div class="scroll-container">
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th>Time Period</th>
                                    <th>Print Jobs</th>
                                    <th>Volume</th>
                                </tr>
                            </thead>
                            <tbody>
                                {{#each kioskprints}}
                                <tr>
                                    <td>{{Hour}}:00 - {{Hour}}:59</td>
                                    <td><strong>{{TotalPrints}}</strong></td>
                                    <td>
                                        {{#if TotalPrints}}
                                            <div style="background: #e5e7eb; height: 20px; width: 100%; border-radius: 4px; overflow: hidden;">
                                                <div style="background: #3b82f6; height: 100%; width: {{TotalPrints}}%;"></div>
                                            </div>
                                        {{else}}
                                            -
                                        {{/if}}
                                    </td>
                                </tr>
                                {{/each}}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
            
            <div class="card">
                <div class="card-header">
                    <h2><i class="fas fa-tasks"></i> Print Queue Details</h2>
                </div>
                <div class="card-body">
                    {{#if printsinqueueCount}}
                    <div style="background:#FFF8EA;border:1px solid #E5B567;border-radius:8px;padding:10px 14px;margin-bottom:12px;font-size:14px;line-height:1.5">
                        <b>About stuck jobs.</b> A print service asks for the labels of all its stations at once, so one job that cannot be built holds up
                        every station on that print service, not just its own. Scripts cannot remove print jobs. TouchPoint's nightly cleanup removes jobs
                        older than 12 hours, or TouchPoint support can delete them now. To free the other stations sooner, take the stuck station off its
                        print service and restart that service. Open <i>What's in it</i> to see which settings or label designs the job used.
                    </div>
                    {{/if}}
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>Station</th>
                                <th>Created</th>
                                <th>Status</th>
                            </tr>
                        </thead>
                        <tbody>
                            {{#each printsinqueue}}
                            <tr>
                                <td>
                                    <code>{{Id}}</code>
                                    {{#if IsPractice}}<div style="font-size:12px;color:#6b7280;margin-top:4px">OneVisit practice build: blocks nothing, cleared tonight</div>{{/if}}
                                </td>
                                <td>{{Stamp}}<div style="font-size:12px;color:#6b7280">{{AgeMin}} min ago</div></td>
                                <td>
                                    {{#if CannotPrint}}<span class="badge badge-danger">Cannot print</span>{{else}}{{#if HeldUp}}<span class="badge badge-info">Held up</span>{{else}}{{#if IsStuck}}<span class="badge badge-warning">Stuck in Queue</span>{{else}}<span class="badge badge-info">Waiting to print</span>{{/if}}{{/if}}{{/if}}
                                    {{#if HeldUp}}<div style="font-size:13px;margin-top:4px">Waiting behind the {{BlockerAt}} job on {{BlockerStation}} that cannot print. This one is fine and prints once that job is removed.</div>{{/if}}
                                    <details style="margin-top:6px"><summary style="cursor:pointer">What's in it</summary>
                                        <div style="font-size:13px;margin-top:6px;line-height:1.5">
                                            {{#if EmptyDesigns}}<div style="color:#9B2C2C;font-weight:600">Cannot print: {{EmptyDesigns}} has no fields. This job will hold up its print service until it is removed.</div>{{/if}}
                                            {{#if Who}}<div><b>For:</b> {{Who}}</div>{{/if}}
                                            <div><b>Printer:</b> {{#if Printer}}{{Printer}}{{else}}settings not available right now. TouchPoint keeps one set of printer settings per print server computer, and when several print services share a computer, only the last one to sign in is kept. It currently holds: {{ListedServices}}. This station prints fine; only its settings cannot be shown.{{/if}}</div>
                                            {{#if IsOldFormat}}<div>Older kiosk format; the details cannot be read here.</div>{{else}}<div>{{Settings}}</div>{{/if}}
                                            {{#if Designs}}<div><b>Label designs:</b> {{Designs}}</div>{{/if}}
                                        </div>
                                    </details>
                                </td>
                            </tr>
                            {{/each}}
                            {{#unless printsinqueue}}
                            <tr>
                                <td colspan="4" class="empty-state">
                                    <i class="fas fa-check-circle"></i>
                                    <p>No print jobs currently in queue</p>
                                </td>
                            </tr>
                            {{/unless}}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        
        <!-- Login Activity Tab -->
        <div id="login" class="tab-content">
            <!-- Charts Grid - Moved to top -->
            <div class="grid">
                <div class="card">
                    <div class="card-header">
                        <h2><i class="fas fa-chart-line"></i> Login Failure Trends (24hr)</h2>
                    </div>
                    <div class="card-body">
                        <div class="chart-container" id="failureTrendsChart">
                            <!-- Chart will be inserted here by JavaScript -->
                        </div>
                    </div>
                </div>
                
                <div class="card">
                    <div class="card-header">
                        <h2><i class="fas fa-chart-pie"></i> Login Success Rate (24hr)</h2>
                    </div>
                    <div class="card-body">
                        <div class="chart-container" id="loginRatioChart" style="height: 300px;">
                            <!-- Donut chart will be inserted here -->
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Security Analytics Card - Moved below charts -->
            <div class="card" style="border-left: 4px solid #F59E0B; margin-bottom: 24px;">
                <div class="card-header" style="background: linear-gradient(135deg, #F59E0B 0%, #D97706 100%);">
                    <h2><i class="fas fa-shield-alt"></i> Security Analytics (72hr)</h2>
                    <a href="/LastActivity" class="btn btn-secondary" style="font-size: 12px; padding: 6px 12px;">
                        <i class="fas fa-external-link-alt"></i> View Full Activity Log
                    </a>
                </div>
                <div class="card-body">
                    <p style="margin-bottom: 16px; color: #4B5563;">
                        <i class="fas fa-exclamation-triangle" style="color: #F59E0B;"></i> 
                        Accounts with multiple logins from different IPs or high login frequency
                    </p>
                    <div class="scroll-container" style="max-height: 400px;">
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th>User</th>
                                    <th>Unique IPs</th>
                                    <th>Total Logins</th>
                                    <th>IP Details</th>
                                </tr>
                            </thead>
                            <tbody>
                                {{#each securityAnalytics}}
                                <tr>
                                    <td>
                                        <strong>{{UserName}}</strong><br/>
                                        <small style="color: #6B7280;">{{UserId}}</small>
                                    </td>
                                    <td>
                                        <span class="badge badge-warning">
                                            {{UniqueIPs}}
                                        </span>
                                    </td>
                                    <td>
                                        <span class="badge badge-warning">
                                            {{TotalLogins}}
                                        </span>
                                    </td>
                                    <td style="font-size: 12px;">
                                        <code>{{IPDetails}}</code>
                                    </td>
                                </tr>
                                {{/each}}
                                {{#unless securityAnalytics}}
                                <tr>
                                    <td colspan="4" class="empty-state">
                                        <i class="fas fa-check-circle"></i>
                                        <p>No suspicious login patterns detected</p>
                                    </td>
                                </tr>
                                {{/unless}}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
            
            <div class="grid">
                <div class="card">
                    <div class="card-header">
                        <h2><i class="fas fa-exclamation-circle"></i> Top Failure Types (72hr)</h2>
                    </div>
                    <div class="card-body">
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th>Failure Type</th>
                                    <th>Count</th>
                                    <th>Last Occurrence</th>
                                </tr>
                            </thead>
                            <tbody>
                                {{#each failedloginstat}}
                                <tr>
                                    <td>{{Activity}}</td>
                                    <td>
                                        <span class="badge badge-danger">{{FailedLogins}}</span>
                                    </td>
                                    <td>{{LastFailedAttempt}}</td>
                                </tr>
                                {{/each}}
                            </tbody>
                        </table>
                        <p style="margin-top: 12px; font-size: 13px; color: #6B7280;">
                            <i class="fas fa-info-circle"></i> Showing top 10 failure types. Full data available in activity logs.
                        </p>
                    </div>
                </div>
            </div>
            
            <div class="card">
                <div class="card-header">
                    <h2><i class="fas fa-history"></i> Recent Failed Login Attempts</h2>
                </div>
                <div class="card-body">
                    <div class="scroll-container">
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th>Date/Time</th>
                                    <th>User ID</th>
                                    <th>Activity</th>
                                    <th>Person</th>
                                    <th>IP Address</th>
                                </tr>
                            </thead>
                            <tbody>
                                {{#each failedlogins}}
                                <tr>
                                    <td>{{ActivityDate}}</td>
                                    <td><code>{{UserId}}</code></td>
                                    <td style="max-width: 300px; word-wrap: break-word;">
                                        <span class="badge badge-danger" style="white-space: normal; display: inline-block; line-height: 1.4;">{{Activity}}</span>
                                    </td>
                                    <td>
                                        {{#if PeopleId}}
                                        <a href="https://myfbch.com/Person2/{{PeopleId}}" target="_blank">
                                            <i class="fas fa-external-link-alt"></i> {{PeopleId}}
                                        </a>
                                        {{else}}
                                        <span class="text-muted">-</span>
                                        {{/if}}
                                    </td>
                                    <td>
                                        <a href="https://mxtoolbox.com/SuperTool.aspx?action=ptr%3a{{ClientIp}}&run=toolpage" target="_blank">
                                            <code>{{ClientIp}}</code>
                                        </a>
                                    </td>
                                </tr>
                                {{/each}}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
            
            <div class="card">
                <div class="card-header">
                    <h2><i class="fas fa-sign-in-alt"></i> Recent Successful Logins</h2>
                </div>
                <div class="card-body">
                    <div class="scroll-container">
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th>Date/Time</th>
                                    <th>User ID</th>
                                    <th>Activity</th>
                                    <th>Person</th>
                                    <th>IP Address</th>
                                </tr>
                            </thead>
                            <tbody>
                                {{#each logins}}
                                <tr>
                                    <td>{{ActivityDate}}</td>
                                    <td><code>{{UserId}}</code></td>
                                    <td style="max-width: 300px; word-wrap: break-word;">
                                        <span class="badge badge-success" style="white-space: normal; display: inline-block; line-height: 1.4;">{{Activity}}</span>
                                    </td>
                                    <td>
                                        {{#if PeopleId}}
                                        <a href="https://myfbch.com/Person2/{{PeopleId}}" target="_blank">
                                            <i class="fas fa-external-link-alt"></i> {{PeopleId}}
                                        </a>
                                        {{else}}
                                        <span class="text-muted">-</span>
                                        {{/if}}
                                    </td>
                                    <td>
                                        <a href="https://mxtoolbox.com/SuperTool.aspx?action=ptr%3a{{ClientIp}}&run=toolpage" target="_blank">
                                            <code>{{ClientIp}}</code>
                                        </a>
                                    </td>
                                </tr>
                                {{/each}}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Script Activity Tab -->
        <div id="scripts" class="tab-content">
            <!-- Script Statistics KPIs -->
            <div class="grid" style="margin-bottom: 20px;">
                <div class="kpi-card info">
                    <div class="kpi-header">
                        <div class="kpi-title">
                            <i class="fas fa-users"></i> Unique Users
                        </div>
                        <div class="kpi-icon info">
                            <i class="fas fa-user-cog"></i>
                        </div>
                    </div>
                    <div class="kpi-value">{{uniqueScriptUsers}}</div>
                    <div class="kpi-subtitle">Users ran scripts (7 days)</div>
                </div>
                
                <div class="kpi-card info">
                    <div class="kpi-header">
                        <div class="kpi-title">
                            <i class="fas fa-terminal"></i> Total Executions
                        </div>
                        <div class="kpi-icon info">
                            <i class="fas fa-play-circle"></i>
                        </div>
                    </div>
                    <div class="kpi-value">{{totalScriptExecutions}}</div>
                    <div class="kpi-subtitle">Scripts run (7 days)</div>
                </div>
                
                <div class="kpi-card info">
                    <div class="kpi-header">
                        <div class="kpi-title">
                            <i class="fas fa-calendar-check"></i> Active Days
                        </div>
                        <div class="kpi-icon info">
                            <i class="fas fa-calendar"></i>
                        </div>
                    </div>
                    <div class="kpi-value">{{activeScriptDays}}</div>
                    <div class="kpi-subtitle">Days with activity (7 days)</div>
                </div>
            </div>
            
            <div class="card">
                <div class="card-header">
                    <h2><i class="fas fa-code"></i> Script Executions (Last 7 Days)</h2>
                </div>
                <div class="card-body">
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>Date/Time</th>
                                <th>User</th>
                                <th>Script/Activity</th>
                                <th>Person</th>
                                <th>IP Address</th>
                            </tr>
                        </thead>
                        <tbody>
                            {{#each script}}
                            <tr>
                                <td>{{ActivityDate}}</td>
                                <td><code>{{UserId}}</code></td>
                                <td>
                                    <span class="badge badge-info">{{Activity}}</span>
                                </td>
                                <td>
                                    {{#if PeopleId}}
                                    <a href="https://myfbch.com/Person2/{{PeopleId}}" target="_blank">
                                        <i class="fas fa-external-link-alt"></i> {{PeopleId}}
                                    </a>
                                    {{else}}
                                    <span class="text-muted">-</span>
                                    {{/if}}
                                </td>
                                <td>
                                    <a href="https://mxtoolbox.com/SuperTool.aspx?action=ptr%3a{{ClientIp}}&run=toolpage" target="_blank">
                                        <code>{{ClientIp}}</code>
                                    </a>
                                </td>
                            </tr>
                            {{/each}}
                            {{#unless script}}
                            <tr>
                                <td colspan="5" class="empty-state">
                                    <i class="fas fa-info-circle"></i>
                                    <p>No recent script activity found</p>
                                </td>
                            </tr>
                            {{/unless}}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        
        <!-- User Accounts Tab -->
        <div id="users" class="tab-content">
            <div class="card">
                <div class="card-header">
                    <h2><i class="fas fa-users"></i> User Account Management</h2>
                </div>
                <div class="card-body">
                    {{#if loadUserAccounts}}
                        <div id="user-accounts-grid" class="ag-theme-quartz"></div>
                    {{else}}
                        <div id="user-accounts-loading">
                            <div class="loading-spinner"></div>
                            <p>User account data is not loaded by default to improve dashboard performance.</p>
                            <p>Click the button below to load user account information.</p>
                            <a href="?loadUsers=true#users" id="load-users-btn" class="btn btn-primary">
                                <i class="fas fa-users"></i> Load User Accounts
                            </a>
                        </div>
                    {{/if}}
                </div>
            </div>
        </div>
    </div>
    
    <!-- Load Chart.js for visualizations -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js@3.9.1/dist/chart.min.js"></script>
<script>
    // User Accounts data
    {{#if loadUserAccounts}}
    const userAccountsData = [
        {{#each useraccounts}}
        {
            Name: "{{Name}}",
            UserName: "{{UserName}}",
            CreationDate: "{{CreationDate}}",
            LastLoginDate: "{{LastLoginDate}}",
            LastActivityDate: "{{LastActivityDate}}",
            EmailAddress: "{{EmailAddress}}",
            MFAEnabled: {{#if MFAEnabled}}true{{else}}false{{/if}},
            MustChangePassword: {{#if MustChangePassword}}true{{else}}false{{/if}},
            IsLockedOut: {{#if IsLockedOut}}true{{else}}false{{/if}},
            PeopleId: "{{PeopleId}}",
            UserId: {{UserId}}
        },
        {{/each}}
    ];
    {{else}}
    const userAccountsData = []; // Empty if not loaded yet
    {{/if}}
    
    // Initialize User Accounts Grid
    function initUserAccountsGrid() {
        const gridOptions = {
            rowData: userAccountsData,
            columnDefs: [
                { 
                    headerName: 'Name', 
                    field: "Name", 
                    enableRowGroup: true,
                    cellRenderer: function(params) {
                        if(params.data) {
                            return '<a href="/Person2/' + params.data.PeopleId + '" target="_blank">' + params.data.Name + '</a>';
                        } else {
                            return '';
                        }
                    }
                },
                { 
                    headerName: 'Username',
                    field: "UserName",
                    sortable: true,
                    filter: 'agTextColumnFilter'
                },
                { 
                    headerName: 'Created', 
                    field: "CreationDate", 
                    sortable: true, 
                    filter: 'agDateColumnFilter', 
                    comparator: (date1, date2) => {
                        return new Date(date1) - new Date(date2); 
                    }
                },
                { 
                    headerName: 'Last Login', 
                    field: "LastLoginDate", 
                    sortable: true, 
                    filter: 'agDateColumnFilter',
                    comparator: (date1, date2) => {
                        return new Date(date1) - new Date(date2); 
                    }
                },
                { 
                    headerName: 'Last Activity', 
                    field: "LastActivityDate", 
                    sortable: true, 
                    filter: 'agDateColumnFilter',
                    comparator: (date1, date2) => {
                        return new Date(date1) - new Date(date2); 
                    }
                },
                { 
                    headerName: 'Email', 
                    field: "EmailAddress", 
                    sortable: true, 
                    filter: 'agTextColumnFilter' 
                },
                { 
                    headerName: 'MFA', 
                    field: "MFAEnabled", 
                    sortable: true, 
                    filter: 'agTextColumnFilter',
                    cellRenderer: function(params) {
                        if(params.value) {
                            return '<span class="badge badge-success">Enabled</span>';
                        } else {
                            return '<span class="badge badge-warning">Disabled</span>';
                        }
                    }
                },
                { 
                    headerName: 'Status', 
                    field: "IsLockedOut", 
                    sortable: true, 
                    filter: 'agTextColumnFilter',
                    cellRenderer: function(params) {
                        if(params.value) {
                            return '<span class="badge badge-danger">Locked</span>';
                        } else {
                            return '<span class="badge badge-success">Active</span>';
                        }
                    }
                }
            ],
            defaultColDef: {
                flex: 1,
                minWidth: 150,
                filter: 'agMultiColumnFilter',
                menuTabs: ['filterMenuTab'],
                filterParams: {
                    comparator: (filterLocalDateAtMidnight, cellValue) => {
                        const cellDate = new Date(cellValue);
                        return cellDate - filterLocalDateAtMidnight;
                    }
                }
            },
            groupDefaultExpanded: -1,
            sideBar: {
                toolPanels: [
                    {
                        id: 'columns',
                        labelDefault: 'Columns',
                        labelKey: 'columns',
                        iconKey: 'columns',
                        toolPanel: 'agColumnsToolPanel',
                    },
                    {
                        id: 'filters',
                        labelDefault: 'Filters',
                        labelKey: 'filters',
                        iconKey: 'filter',
                        toolPanel: 'agFiltersToolPanel',
                    }
                ],
                defaultToolPanel: "",
            }
        };
        
        const gridElement = document.getElementById('user-accounts-grid');
        if (gridElement) {
            agGrid.createGrid(gridElement, gridOptions);
        }
    }
    
    // Global function to activate a specific tab
    function activateTab(tabName) {
        const tabs = document.querySelectorAll('.tab');
        const tabContents = document.querySelectorAll('.tab-content');
        
        tabs.forEach(t => t.classList.remove('active'));
        tabContents.forEach(content => content.classList.remove('active'));
        
        // Find the tab element and activate it
        const tabElement = document.querySelector(`.tab[data-tab="${tabName}"]`);
        if (tabElement) {
            tabElement.classList.add('active');
            
            // Find the content element and activate it
            const contentElement = document.getElementById(tabName);
            if (contentElement) {
                contentElement.classList.add('active');
            }
            
            // Initialize charts and grids when tab becomes visible
            if (tabName === 'overview') {
                setTimeout(() => {
                    initPrintChartOverview();
                    initFailureTrendsChartOverview();
                }, 100);
            } else if (tabName === 'prints') {
                setTimeout(initPrintChart, 100);
            } else if (tabName === 'login') {
                setTimeout(() => {
                    initFailureTrendsChart();
                    initLoginRatioChart();
                }, 100);
            } else if (tabName === 'users') {
                {{#if loadUserAccounts}}
                setTimeout(() => {
                    initUserAccountsGrid();
                }, 100);
                {{/if}}
            }
        }
    }
    
    // Tab switching
    document.addEventListener('DOMContentLoaded', function() {
        const tabs = document.querySelectorAll('.tab');
        const tabContents = document.querySelectorAll('.tab-content');
        
        // The address decides the tab: #prints, #login, #scripts or #users opens
        // that tab, and no # (or one that is not a tab) opens Overview.
        function tabFromHash() {
            const hashTab = window.location.hash.substring(1);
            return hashTab && document.querySelector(`.tab[data-tab="${hashTab}"]`) ? hashTab : 'overview';
        }
        activateTab(tabFromHash());
        window.addEventListener('hashchange', () => activateTab(tabFromHash()));
        
        // Add click event listeners to tabs
        tabs.forEach(tab => {
            tab.addEventListener('click', () => {
                const targetTab = tab.dataset.tab;
                activateTab(targetTab);
                
                // Update URL hash
                window.location.hash = targetTab;
            });
        });
        
        // Handle expand buttons for overview sections
        const expandButtons = document.querySelectorAll('.expand-button');
        expandButtons.forEach(button => {
            button.addEventListener('click', () => {
                const targetId = button.getAttribute('data-target');
                const targetElement = document.getElementById(targetId);
                if (targetElement) {
                    targetElement.classList.toggle('expanded');
                    
                    // Change the icon based on state
                    const icon = button.querySelector('i');
                    if (targetElement.classList.contains('expanded')) {
                        icon.classList.replace('fa-expand-alt', 'fa-compress-alt');
                    } else {
                        icon.classList.replace('fa-compress-alt', 'fa-expand-alt');
                    }
                }
            });
        });
        
    });
    
    // Initialize Print Chart
    function initPrintChart() {
        const printChartEl = document.getElementById('printChart');
        if (!printChartEl) return;
        
        // Get data from the table
        const hours = [];
        const counts = [];
        
        // Parse kioskprints data directly
        {{#each kioskprints}}
        hours.push('{{Hour}}:00');
        counts.push({{TotalPrints}});
        {{/each}}
        
        // Create the chart
        const ctx = document.createElement('canvas');
        printChartEl.innerHTML = '';
        printChartEl.appendChild(ctx);
        
        new Chart(ctx, {
            type: 'bar',
            data: {
                labels: hours,
                datasets: [{
                    label: 'Print Jobs',
                    data: counts,
                    backgroundColor: 'rgba(79, 70, 229, 0.7)',
                    borderColor: 'rgba(79, 70, 229, 1)',
                    borderWidth: 1,
                    borderRadius: 4
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: false
                    },
                    title: {
                        display: true,
                        text: 'Print Jobs by Hour',
                        font: {
                            size: 16,
                            weight: 'bold'
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            precision: 0
                        },
                        grid: {
                            color: 'rgba(0, 0, 0, 0.05)'
                        }
                    },
                    x: {
                        grid: {
                            display: false
                        }
                    }
                }
            }
        });
    }
    
    // Initialize Failure Trends Chart
    function initFailureTrendsChart() {
        const chartEl = document.getElementById('failureTrendsChart');
        if (!chartEl) return;
        
        // Get data from server
        const hours = [];
        const counts = [];
        
        // Parse failure trends data
        {{#each failureTrends}}
        hours.push('{{Hour}}:00');
        counts.push({{FailureCount}});
        {{/each}}
        
        // Fill in missing hours with 0
        const fullHours = [];
        const fullCounts = [];
        for (let i = 0; i < 24; i++) {
            fullHours.push(i + ':00');
            const idx = hours.indexOf(i + ':00');
            fullCounts.push(idx >= 0 ? counts[idx] : 0);
        }
        
        // Create the chart
        const ctx = document.createElement('canvas');
        chartEl.innerHTML = '';
        chartEl.appendChild(ctx);
        
        new Chart(ctx, {
            type: 'line',
            data: {
                labels: fullHours,
                datasets: [{
                    label: 'Failed Login Attempts',
                    data: fullCounts,
                    borderColor: '#EF4444',
                    backgroundColor: 'rgba(239, 68, 68, 0.1)',
                    borderWidth: 2,
                    tension: 0.3,
                    fill: true
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: false
                    },
                    title: {
                        display: true,
                        text: 'Failed Login Attempts by Hour (Last 24 Hours)',
                        font: {
                            size: 16,
                            weight: 'bold'
                        }
                    },
                    tooltip: {
                        mode: 'index',
                        intersect: false
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            precision: 0
                        },
                        title: {
                            display: true,
                            text: 'Number of Failures'
                        }
                    },
                    x: {
                        title: {
                            display: true,
                            text: 'Hour of Day'
                        }
                    }
                }
            }
        });
    }
    
    // Initialize Overview Charts
    function initPrintChartOverview() {
        const printChartEl = document.getElementById('printChartOverview');
        if (!printChartEl) return;
        
        // Get data from the table
        const hours = [];
        const counts = [];
        
        // Parse kioskprints data directly
        {{#each kioskprints}}
        hours.push('{{Hour}}:00');
        counts.push({{TotalPrints}});
        {{/each}}
        
        // Create the chart
        const ctx = document.createElement('canvas');
        printChartEl.innerHTML = '';
        printChartEl.appendChild(ctx);
        
        new Chart(ctx, {
            type: 'bar',
            data: {
                labels: hours,
                datasets: [{
                    label: 'Print Jobs',
                    data: counts,
                    backgroundColor: 'rgba(79, 70, 229, 0.7)',
                    borderColor: 'rgba(79, 70, 229, 1)',
                    borderWidth: 1,
                    borderRadius: 4
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: false
                    },
                    title: {
                        display: true,
                        text: 'Print Jobs by Hour Today',
                        font: {
                            size: 16,
                            weight: 'bold'
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            precision: 0
                        },
                        grid: {
                            color: 'rgba(0, 0, 0, 0.05)'
                        }
                    },
                    x: {
                        grid: {
                            display: false
                        }
                    }
                }
            }
        });
    }
    
    function initFailureTrendsChartOverview() {
        const chartEl = document.getElementById('failureTrendsChartOverview');
        if (!chartEl) return;
        
        // Get data from server
        const hours = [];
        const counts = [];
        
        // Parse failure trends data
        {{#each failureTrends}}
        hours.push('{{Hour}}:00');
        counts.push({{FailureCount}});
        {{/each}}
        
        // Fill in missing hours with 0
        const fullHours = [];
        const fullCounts = [];
        for (let i = 0; i < 24; i++) {
            fullHours.push(i + ':00');
            const idx = hours.indexOf(i + ':00');
            fullCounts.push(idx >= 0 ? counts[idx] : 0);
        }
        
        // Create the chart
        const ctx = document.createElement('canvas');
        chartEl.innerHTML = '';
        chartEl.appendChild(ctx);
        
        new Chart(ctx, {
            type: 'line',
            data: {
                labels: fullHours,
                datasets: [{
                    label: 'Failed Login Attempts',
                    data: fullCounts,
                    borderColor: '#EF4444',
                    backgroundColor: 'rgba(239, 68, 68, 0.1)',
                    borderWidth: 2,
                    tension: 0.3,
                    fill: true
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: false
                    },
                    title: {
                        display: true,
                        text: 'Failed Login Attempts (Last 24 Hours)',
                        font: {
                            size: 16,
                            weight: 'bold'
                        }
                    },
                    tooltip: {
                        mode: 'index',
                        intersect: false
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            precision: 0
                        },
                        title: {
                            display: true,
                            text: 'Number of Failures'
                        }
                    },
                    x: {
                        title: {
                            display: true,
                            text: 'Hour of Day'
                        }
                    }
                }
            }
        });
    }
    
    // Initialize Login Success/Failure Ratio Chart
    function initLoginRatioChart() {
        const chartEl = document.getElementById('loginRatioChart');
        if (!chartEl) return;
        
        const successCount = {{successCount}};
        const failureCount = {{failureCount}};
        const total = successCount + failureCount;
        
        if (total === 0) {
            chartEl.innerHTML = '<div class="empty-state"><i class="fas fa-info-circle"></i><p>No login data available</p></div>';
            return;
        }
        
        const ctx = document.createElement('canvas');
        chartEl.innerHTML = '';
        chartEl.appendChild(ctx);
        
        new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: ['Successful Logins', 'Failed Logins'],
                datasets: [{
                    data: [successCount, failureCount],
                    backgroundColor: ['#10B981', '#EF4444'],
                    borderWidth: 0
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            padding: 20,
                            font: {
                                size: 14
                            }
                        }
                    },
                    title: {
                        display: true,
                        text: 'Success Rate: ' + Math.round((successCount / total) * 100) + '%',
                        font: {
                            size: 18,
                            weight: 'bold'
                        },
                        padding: 20
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const label = context.label || '';
                                const value = context.parsed;
                                const percentage = Math.round((value / total) * 100);
                                return label + ': ' + value + ' (' + percentage + '%)';
                            }
                        }
                    }
                },
                cutout: '60%'
            }
        });
    }
</script>
<script>
(function(){
    // AUTO-UPDATE. Only Admins see the banner.
    var APP_VERSION = '{{appVersion}}', DC_SCRIPT_ID = 'TPxi_TechnicalStatus', IS_ADMIN = '{{isAdmin}}' === '1';
    var DC_API_BASE = 'https://scripts.displaycache.com/api/touchpoint', LATEST = '';
    var SCRIPT_NAME = (function(){ var m = (location.pathname || '').match(/\\/PyScript(?:Form)?\\/([^\\/?#]+)/); return m ? m[1] : DC_SCRIPT_ID; })();
    function newer(pub, mine){
        var a = String(pub || '').split('.'), b = String(mine || '').split('.');
        for (var i = 0; i < Math.max(a.length, b.length); i++){ var x = parseInt(a[i], 10) || 0, y = parseInt(b[i], 10) || 0; if (x !== y) return x > y; }
        return false;
    }
    function banner(){
        var b = document.getElementById('appUpdateBanner'); if (!b) return;
        b.innerHTML = '<div style="flex:1;font-size:13px;color:#0B5394"><strong>A newer Technical Status is out.</strong> You have ' + APP_VERSION +
            ', the newest is ' + LATEST + '. <span id="appUpdateSt"></span></div>' +
            '<button id="appUpdateBtn" type="button" style="white-space:nowrap;padding:6px 14px;background:#0078d4;color:#fff;border:0;border-radius:4px;cursor:pointer">Update now</button>' +
            '<button id="appUpdateLater" type="button" style="white-space:nowrap;padding:6px 14px;background:#fff;color:#0078d4;border:1px solid #0078d4;border-radius:4px;cursor:pointer">Later</button>';
        b.style.display = 'flex';
        document.getElementById('appUpdateBtn').onclick = applyUpdate;
        document.getElementById('appUpdateLater').onclick = function(){ try { sessionStorage.setItem('tsupdate.later', LATEST); } catch(e) {} b.style.display = 'none'; };
    }
    function applyUpdate(){
        if (!confirm('Update Technical Status from ' + APP_VERSION + ' to ' + LATEST + '?')) return;
        var btn = document.getElementById('appUpdateBtn'), st = document.getElementById('appUpdateSt');
        btn.disabled = true; st.textContent = 'Updating...';
        var x = new XMLHttpRequest(), path = location.pathname.replace('/PyScript/', '/PyScriptForm/');
        x.open('POST', path, true); x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        x.onload = function(){
            var t = x.responseText || '', i = t.indexOf('{'), j = t.lastIndexOf('}'), d = null;
            try { d = JSON.parse(t.slice(i, j + 1)); } catch(e) {}
            st.textContent = (d && d.message) || 'Unexpected reply from TouchPoint.';
            if (d && d.success) setTimeout(function(){ location.reload(); }, 800); else btn.disabled = false;
        };
        x.onerror = function(){ btn.disabled = false; st.textContent = 'Network error.'; };
        x.send('action=apply_update&script_name=' + encodeURIComponent(SCRIPT_NAME));
    }
    if (!IS_ADMIN) return;
    var skip = ''; try { skip = sessionStorage.getItem('tsupdate.later') || ''; } catch(e) {}
    try {
        var x = new XMLHttpRequest(); x.open('GET', DC_API_BASE + '/script-versions', true); x.timeout = 5000;
        x.onreadystatechange = function(){
            if (x.readyState !== 4 || x.status !== 200) return;
            try { var v = JSON.parse(x.responseText)[DC_SCRIPT_ID]; if (v && newer(v, APP_VERSION) && skip != v) { LATEST = v; banner(); } } catch(e) {}
        };
        x.send();
    } catch(e) {}
})();
</script>
</body>
</html>
'''

# Pass the loadUserAccounts flag and counts to the template
Data.appVersion = APP_VERSION
Data.isAdmin = '1' if _ts_is_admin() else '0'
Data.loadUserAccounts = loadUserAccounts
Data.printsinqueueCount = Data.printsinqueueCount
Data.scriptCount = Data.scriptCount
Data.totalFailedLogins = Data.totalFailedLogins
Data.uniqueScriptUsers = Data.uniqueScriptUsers
Data.totalScriptExecutions = Data.totalScriptExecutions
Data.activeScriptDays = Data.activeScriptDays
Data.failureTrends = Data.failureTrends
Data.securityAnalytics = Data.securityAnalytics
Data.successCount = Data.successCount
Data.failureCount = Data.failureCount

# Render the template with our data
if not _ts_ajax_done:
    dashboardReport = model.RenderTemplate(template)
    print(dashboardReport)
