########################################################
### Enhanced Scheduler Report
### Original: SimpleSchedulerReport
### Updates: Improved layout, error handling, multi-column display

########################################################
### Info
#Created By: Ben Swaby
#Email: bswaby@fbchtn.org

# Enhanced scheduler report with improved layout and error handling
# Features:
# - Better error handling and validation
# - Visual indicators for staffing levels
# - Duplicate volunteer handling
# - No hardcoded script names - works with any filename

# USAGE OPTIONS:

# Option 1: Bluetoolbar Menu
# - Add to Bluetoolbar via CustomReport in special content:
#   <Report name="YourScriptName" type="PyScript" role="Access" />
# - Run from menu or directly with ?CurrentOrgId=#### 

# Option 2: Morning Batch (Automated Sending)
# Create a separate morning batch script like this:
'''
# Morning Batch Scheduler
from datetime import datetime

# Send on specific days of the month (1st and 15th in this example)
if model.DateTime.Today.Day in [1, 15]:
    model.Data.sendReport = 'y'
    model.Data.reportTo = 'Involvement'  # or 'Self'
    model.Data.CurrentOrgId = '2832'     # Your involvement ID
    print(model.CallScript("YourScriptName"))

# Or send on specific weekdays (Monday=1, Sunday=7)
# if model.DateTime.Today.DayOfWeek in [1, 3, 5]:  # Mon, Wed, Fri
#     model.Data.sendReport = 'y'
#     model.Data.reportTo = 'Involvement'
#     model.Data.CurrentOrgId = '2832'
#     print(model.CallScript("YourScriptName"))

# Or send X days before the schedule date
# You would need to modify the main script to accept a lookahead parameter
'''

########################################################
### User Config Area

SendRoles = "Admin,ManageGroups"    # Roles that can send report to involvement
ScheduleDays = '365'                # Days to include in report
ShowEmptySlots = True               # Show slots with no volunteers
GroupByDate = True                  # Group time slots by date
ColumnsPerRow = 3                   # Number of columns for multi-column layout

# Email Variables
FromName = 'Scheduler'
FromAddress = 'scheduler@church.org'  # Update this!
Subject = ''  # Uses org name if empty

# Visual Settings
ColorScheme = {
    'full': '#4CAF50',      # Green for fully staffed
    'partial': '#FFC107',   # Amber for partially staffed
    'empty': '#F44336',     # Red for no volunteers
    'required': '#2196F3'   # Blue for required positions
}

########################################################
### Start of Code

from datetime import datetime

# Get current script path for self-referencing URLs
# This allows the script to work with any filename
try:
    scriptPath = model.ScriptName if hasattr(model, 'ScriptName') else 'SchedulerReport'
except:
    # Fallback - extract from URL if available
    scriptPath = 'SchedulerReport'

# Initialize variables with error handling
try:
    sendReport = model.Data.sendReport if hasattr(model.Data, 'sendReport') else 'n'
    reportTo = model.Data.reportTo if hasattr(model.Data, 'reportTo') else ''
    OrgId = str(Data.CurrentOrgId) if Data.CurrentOrgId else None
except:
    sendReport = 'n'
    reportTo = ''
    OrgId = None

# Validate OrgId
if not OrgId:
    print('<div class="alert alert-danger">Error: No organization ID provided. Please access this report from an involvement page or provide CurrentOrgId parameter.</div>')
    raise SystemExit()

# Check permissions for send to involvement
showSendToInvolvement = 'n'
try:
    for role in SendRoles.split(','):
        if model.UserIsInRole(role.strip()):
            showSendToInvolvement = 'y'
            break
except Exception as e:
    print('<div class="alert alert-warning">Warning: Could not check user roles: {}</div>'.format(e))

# Get organization info with error handling
try:
    org = model.GetOrganization(int(OrgId))
    if not org:
        raise ValueError("Organization not found")
    
    model.Header = "Schedule: {}".format(org.name)
    orgName = org.name
except Exception as e:
    print('<div class="alert alert-danger">Error: Could not load organization {}: {}</div>'.format(OrgId, e))
    raise SystemExit()

# Enhanced SQL query with duplicate handling
sql = '''
DECLARE @InvId Int = {0}
DECLARE @days int = {1}

;WITH
-- Get latest volunteer records only (handles duplicates from cancel/re-signup)
LatestVolunteers AS (
    SELECT 
        TimeSlotMeetingId, 
        PeopleId, 
        MAX(TimeSlotMeetingVolunteerId) as TimeSlotMeetingVolunteerId
    FROM TimeSlotMeetingVolunteers
    GROUP BY TimeSlotMeetingId, PeopleId
),
AllServices AS 
(
SELECT 
    tsm.MeetingDateTime as ServiceDate,
    FORMAT(tsm.MeetingDateTime, 'hh:mm tt') as ServiceTime,
    FORMAT(tsm.MeetingDateTime, 'M/d/yy h:mm tt') as ServiceDateTime,
    FORMAT(tsm.MeetingDateTime, 'dddd, MMMM dd, yyyy') AS DateOnly,
    DATEPART(WEEKDAY, tsm.MeetingDateTime) as DayOfWeek,
    tstSG.NumberVolunteersNeeded as Needed,
    tstSG.Require as [Required],
    tsmt.TimeSlotMeetingTeamId,
    tsgrpvol.TimeSlotMeetingTeamSubGroupId,
    tssg.NumberVolunteersNeeded as SubGroupsNeeded,
    CASE 
        WHEN (tsmv.VolunteerOption IS NOT NULL AND tsmv.VolunteerOption <> 'This Time Slot') 
            THEN CONCAT(p.Name2, ' (', tsmv.VolunteerOption, ')') 
        ELSE p.Name2 
    END AS Name2,
    p.Name,
    p.PeopleId,
    p.EmailAddress,
    p.CellPhone,
    CASE WHEN mt.Name IS NULL THEN tsmt.TeamName ELSE mt.Name END AS SubGroupTeam,
    tsmt.TeamName,
    m.MeetingId,
    m.Location

FROM TimeSlotMeetingTeams tsmt 
LEFT JOIN TimeSlotMeetings tsm ON tsmt.TimeSlotMeetingId = tsm.TimeSlotMeetingId
LEFT JOIN TimeSlotMeetingTeamSubGroups tssg ON tssg.TimeSlotMeetingTeamId = tsmt.TimeSlotMeetingTeamId
LEFT JOIN TimeSlotTeamSubGroups tstSG ON tstSG.TimeSlotTeamSubGroupId = tssg.TimeSlotTeamSubGroupId
LEFT JOIN Meetings m ON tsm.MeetingId = m.MeetingId
LEFT JOIN TimeSlotMeetingTeamSubGroupVolunteers tsGrpVol ON 
    (tsGrpVol.IsActive <> 0 AND tsGrpVol.TimeSlotMeetingTeamId = tsmt.TimeSlotMeetingTeamId 
        AND tsGrpVol.TimeSlotMeetingTeamSubGroupId IS NULL)
    OR (tsGrpVol.IsActive <> 0 AND tsGrpVol.TimeSlotMeetingTeamSubGroupId = tssg.TimeSlotMeetingTeamSubGroupId  
        AND tsGrpVol.TimeSlotMeetingTeamSubGroupId IS NOT NULL)
LEFT JOIN LatestVolunteers lv ON lv.TimeSlotMeetingId = tsmt.TimeSlotMeetingId 
    AND lv.PeopleId = tsGrpVol.PeopleId
LEFT JOIN TimeSlotMeetingVolunteers tsmv ON tsmv.TimeSlotMeetingVolunteerId = lv.TimeSlotMeetingVolunteerId
LEFT JOIN MemberTags mt ON mt.Id = tssg.MemberTagId AND mt.OrgId = {0}
LEFT JOIN People p ON p.PeopleId = tsGrpVol.PeopleId

WHERE 
    tsm.MeetingDateTime > GETDATE() 
    AND tsm.MeetingDateTime < DATEADD(DAY, {1}, GETDATE())
    AND OrganizationId = {0}
),
-- Aggregate volunteers by service
ServiceSummary AS (
    SELECT 
        ServiceDate,
        ServiceTime,
        ServiceDateTime,
        DateOnly,
        DayOfWeek,
        TeamName,
        SubGroupTeam,
        TimeSlotMeetingTeamId,
        Needed,
        Required,
        Location,
        COUNT(DISTINCT CASE WHEN Name2 IS NOT NULL THEN PeopleId END) as Serving,
        STRING_AGG(Name2, ', ') WITHIN GROUP (ORDER BY Name2) as Names
    FROM AllServices
    GROUP BY 
        ServiceDate, ServiceTime, ServiceDateTime, DateOnly, DayOfWeek,
        TeamName, SubGroupTeam, TimeSlotMeetingTeamId, Needed, Required, Location
)
SELECT 
    *,
    CASE 
        WHEN Needed IS NULL OR Needed = 0 THEN 100
        ELSE (CAST(Serving AS FLOAT) / CAST(Needed AS FLOAT)) * 100
    END as FillPercentage,
    CASE 
        WHEN Serving = 0 AND Required = 1 THEN 'Required - Empty'
        WHEN Serving = 0 THEN 'Empty'
        WHEN Needed IS NULL OR Needed = 0 THEN 'Full'
        WHEN Serving >= Needed THEN 'Full'
        WHEN Serving > 0 THEN 'Partial'
        ELSE 'Empty'
    END as Status
FROM ServiceSummary
WHERE SubGroupTeam NOT IN ('EMS', 'LEO')  -- Exclude special teams
ORDER BY ServiceDate, ServiceTime, SubGroupTeam
'''

# Required SQL for BlueToolbar compatibility (not used but must be present)
sql2 = '''
SELECT p.PeopleId FROM dbo.People p JOIN dbo.TagPerson tp ON tp.PeopleId = p.PeopleId AND tp.Id = @BlueToolbarTagId
'''

# CSS Styles for compact email-friendly layout
styles = '''
<style>
/* Compact table styles for email and print */
body {
    font-family: Arial, sans-serif;
    color: #333;
}

.scheduler-container {
    max-width: 800px;
    margin: 0 auto;
    padding: 5px;
}

h1 {
    text-align: center;
    color: #333;
    font-size: 20px;
    margin: 10px 0;
}

.btn {
    display: inline-block;
    padding: 6px 12px;
    text-decoration: none;
    background-color: #3AAEE0;
    color: white;
    border-radius: 4px;
    font-weight: bold;
    margin: 3px;
    font-size: 13px;
}

.action-buttons {
    text-align: center;
    margin: 10px 0;
}

/* Main schedule table */
.schedule-table {
    width: 100%;
    border-collapse: collapse;
    margin: 10px 0;
    font-size: 13px;
    line-height: 1.3;
}

.schedule-table th {
    background-color: #f8f9fa;
    padding: 6px 5px;
    text-align: left;
    font-weight: bold;
    border-bottom: 2px solid #dee2e6;
    font-size: 12px;
}

.schedule-table td {
    padding: 4px 5px;
    border-bottom: 1px solid #e9ecef;
    vertical-align: top;
}

/* Date header rows */
.date-header-row {
    background-color: #e9ecef;
    font-weight: bold;
}

.date-header-row td {
    padding: 8px 5px;
    font-size: 14px;
    border-bottom: 2px solid #dee2e6;
}

/* Status indicators */
.status-cell {
    text-align: center;
    font-weight: bold;
}

.status-full {
    color: #28a745;
}

.status-partial {
    color: #ffc107;
}

.status-empty {
    color: #dc3545;
}

.status-required {
    color: #dc3545;
    font-weight: bold;
}

/* Sign up link */
.signup-link {
    color: #3AAEE0;
    text-decoration: none;
}

.signup-link:hover {
    text-decoration: underline;
}

/* Alert styles */
.alert {
    padding: 8px;
    margin: 8px 0;
    border-radius: 4px;
    font-size: 13px;
}

.alert-info {
    background-color: #d1ecf1;
    color: #0c5460;
    border: 1px solid #bee5eb;
}

.alert-warning {
    background-color: #fff3cd;
    color: #856404;
    border: 1px solid #ffeaa7;
}

.alert-danger {
    background-color: #f8d7da;
    color: #721c24;
    border: 1px solid #f5c6cb;
}

/* Print optimization */
@media print {
    .action-buttons, .no-print {
        display: none;
    }
    
    .schedule-table {
        font-size: 11px;
    }
    
    .schedule-table td {
        padding: 2px 3px;
    }
    
    .date-header-row {
        break-inside: avoid;
    }
    
    .date-header-row td {
        padding: 5px 3px;
        font-size: 12px;
    }
}

/* Email client compatibility */
table {
    mso-table-lspace: 0pt;
    mso-table-rspace: 0pt;
}

td {
    mso-line-height-rule: exactly;
}
</style>
'''

# Build the report
template = '''{}
<div class="scheduler-container">
    <h1>{} Schedule</h1>
    
    <div class="action-buttons no-print">
        <a href="{}/OnlineReg/{}" class="btn">
            Manage My Commitments
        </a>
    </div>
'''.format(styles, orgName, model.CmsHost, OrgId)

# Execute query with error handling
try:
    results = q.QuerySql(sql.format(OrgId, ScheduleDays))
    
    if not results:
        template += '''
        <div class="alert alert-info">
            No scheduled services found for the next {} days.
        </div>
        '''.format(ScheduleDays)
    else:
        # Start table
        template += '''
        <table class="schedule-table">
            <thead>
                <tr>
                    <th width="12%">Time</th>
                    <th width="28%">Position</th>
                    <th width="8%">Need</th>
                    <th width="52%">Volunteers</th>
                </tr>
            </thead>
            <tbody>
        '''
        
        current_date = None
        
        for row in results:
            # Add date header row
            if current_date != str(row.DateOnly):
                current_date = str(row.DateOnly)
                template += '''
                <tr class="date-header-row">
                    <td colspan="4">{}</td>
                </tr>
                '''.format(current_date)
            
            # Determine status
            status_class = 'status-empty'
            status_text = '0/{}'.format(row.Needed if row.Needed else '?')
            
            if row.Status == 'Full':
                status_class = 'status-full'
                status_text = 'Full'
            elif row.Status == 'Partial':
                status_class = 'status-partial'
                status_text = '{}/{}'.format(row.Serving, row.Needed)
            elif row.Status == 'Required - Empty':
                status_class = 'status-required'
                status_text = 'NEEDED!'
            
            # Build volunteer column
            volunteer_cell = ''
            if row.Names:
                volunteer_cell = row.Names
            elif not ShowEmptySlots and row.Serving == 0:
                continue  # Skip empty slots if configured
            else:
                volunteer_cell = '<a href="{}/OnlineReg/{}" class="signup-link">Click to sign up</a>'.format(model.CmsHost, OrgId)
            
            # Add row
            template += '''
            <tr>
                <td>{}</td>
                <td>{}</td>
                <td class="status-cell {}">
                    {}
                </td>
                <td>{}</td>
            </tr>
            '''.format(row.ServiceTime, row.SubGroupTeam, status_class, status_text, volunteer_cell)
        
        # Close table
        template += '''
            </tbody>
        </table>
        '''

except Exception as e:
    template += '''
    <div class="alert alert-danger">
        Error loading schedule data: {}
    </div>
    '''.format(e)

# Close action buttons and container
template += '''
    </div>
</div>
'''

# Handle report sending
NMReport = model.RenderTemplate(template)

if sendReport == 'y' and reportTo == 'Self':
    try:
        # Add tracking
        NMReport += '{track}{tracklinks}<br />'
        
        # Email variables
        QueuedBy = model.UserPeopleId
        MailToQuery = model.UserPeopleId
        
        # Set subject
        if not Subject:
            Subject = "Schedule: {}".format(orgName)
        
        # Send email
        model.Email(MailToQuery, QueuedBy, FromAddress, FromName, Subject, NMReport)
        
        print('<div class="alert alert-info">Report sent to your email address.</div>')
    except Exception as e:
        print('<div class="alert alert-danger">Error sending email: {}</div>'.format(e))

elif sendReport == 'y' and reportTo == 'Involvement':
    if showSendToInvolvement != 'y':
        print('<div class="alert alert-warning">You do not have permission to send reports to the involvement.</div>')
    else:
        try:
            # Add tracking
            NMReport += '{track}{tracklinks}<br />'
            
            QueuedBy = model.UserPeopleId
            
            # Set subject
            if not Subject:
                Subject = "Schedule: {}".format(orgName)
            
            # Get involvement members
            sqlEmailTo = '''
            SELECT p.PeopleId, p.Name, p.EmailAddress 
            FROM OrganizationMembers om
            JOIN People p ON om.PeopleId = p.PeopleId
            WHERE om.OrganizationId = {} 
                AND om.Pending = 0 
                AND p.EmailAddress IS NOT NULL
                AND p.EmailAddress != ''
            '''.format(OrgId)
            
            recipients = q.QuerySql(sqlEmailTo)
            sent_count = 0
            error_count = 0
            
            for recipient in recipients:
                try:
                    model.Email(recipient.PeopleId, QueuedBy, FromAddress, FromName, Subject, NMReport)
                    sent_count += 1
                except:
                    error_count += 1
            
            print('''
            <div class="alert alert-info">
                Report sent to {} involvement members.
                {}
            </div>
            '''.format(sent_count, '({} errors)'.format(error_count) if error_count > 0 else ''))
            
        except Exception as e:
            print('<div class="alert alert-danger">Error sending to involvement: {}</div>'.format(e))

else:
    # Show action buttons at bottom
    action_buttons = '<div class="action-buttons no-print">'
    
    if showSendToInvolvement == 'y':
        action_buttons += '''
        <a href="/PyScript/{}?sendReport=y&reportTo=Involvement&CurrentOrgId={}" class="btn">
            Send Schedule to Involvement
        </a>
        '''.format(scriptPath, OrgId)
    
    action_buttons += '''
    <a href="/PyScript/{}?sendReport=y&reportTo=Self&CurrentOrgId={}" class="btn">
        Send Schedule to Self
    </a>
    </div>
    '''.format(scriptPath, OrgId)
    
    # Insert buttons after the title
    template = template.replace('</h1>', '</h1>\n' + action_buttons)
    
    print(template)
