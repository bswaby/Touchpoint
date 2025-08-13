########################################################
### Enhanced Scheduler Report v2.1.1
### Original: SimpleSchedulerReport
### Updates: Multi-column layout, family filtering, proper functions
########################################################

### Info
# The built-in calendar view is complicated and this attempt to crate a simple view
# so team members can easily see a full calendar.  It's built to be seen from the web,
# printed, and emailed out. 

# Created By: Ben Swaby
# Email: bswaby@fbchtn.org
# 
# Enhanced By: Heath Kouns (8/9/25)
# - Semi-colon separator for names
# - New tab for registration links
# - Family schedule management buttons
# 
# Version 2.1 Updates: Ben Swaby (8/9/25)
# - Got multi-column layout option working
# - Enhanced filter family buttons to show all family or only family members who are involved
# - Put in proper function-based code structure
# - Fixed print not printing color and shading
#
# Enhanced By: Heath Kouns (8/12/25)
# - Added MinVolunteer Age
# - Added a Configurable Title
# - Added an option to have Org Specific over-writing of the default settings for the following variables:
#   ShowEmptySlots, ShowFamilyButtons, MinVolunteerAge, Title, FromAddress, Subject 
#   AdHoc ExtraValues are in the form 'SchedulerReport' + variable name  (e.g. SchedulerReportShowEmptySlots)
#
#
#To add this to the Bluetoolbar, navigate to open CustomReport under special content 
# / text and add in the following line.  Make sure to adjust report name to what
# you called the Python report
#
#  <Report name="ScheduleList" type="PyScript" role="Access" />
#
# note: CustomReport changes can take 24 hrs to show on the site due to how the cache 
# has been implemented on the TP servers

# Add to the morning batch to automatically email
########################################################
### MORNING BATCH EXAMPLES
########################################################
'''
MORNING BATCH SETUP:

1. Create a new Python script called "SchedulerBatch"
2. Set it to run in Morning Batch
3. Use one of these examples:

# Send on specific days of month
if model.DateTime.Today.Day in [1, 15]:
    model.Data.sendReport = 'y'
    model.Data.reportTo = 'Involvement'
    model.Data.CurrentOrgId = '2832'
    print(model.CallScript("YourSchedulerScriptName"))

# Send weekly on specific days
if model.DateTime.Today.DayOfWeek in [1, 4]:  # Monday, Thursday
    model.Data.sendReport = 'y'
    model.Data.reportTo = 'Involvement'
    model.Data.CurrentOrgId = '2832'
    print(model.CallScript("YourSchedulerScriptName"))
'''
global ShowEmptySlots, ShowFamilyButtons, MinVolunteerAge, Title, FromAddress, Subject

########################################################
### User Config Area

#Default Values - Some can be overwritten with Org ExtraValue

SendRoles = "Admin,ManageGroups"    # Roles that can send report to involvement
ScheduleDays = '365'                # Days to include in report
ShowEmptySlots = True               # Show slots with no volunteers | Org ExtraValue (bool): SchedulerReportShowEmptySlots
GroupByDate = True                  # Group time slots by date
ShowFamilyButtons = 1               # 0=No buttons, 1=All family, 2=Only family in involvement | Org ExtraValue (int): SchedulerReportShowFamilyButtons
MinVolunteerAge = 12                # Minimum Age of Family Members to show buttons for | Org ExtraValue (int): SchedulerReportMinVolunteerAge (min: 1)
UseMultiColumn = True               # Use multi-column layout
ColumnsPerRow = 2                   # Number of columns (2 or 3 recommended)
Title = ''                          # Uses '{org name} + Schedule' if empty | or can be overwritten with Org ExtraValue (string): SchedulerReportTitle

# Email Variables
FromName = 'Scheduler'
FromAddress = 'scheduler@church.org'   # Update this! | Org ExtraValue (string): SchedulerReportFromAddress
Subject = ''                  # Uses org name if empty | Org ExtraValue (string): SchedulerReportSubject

########################################################
### Start of Code

# Initialize global variables
def initialize_variables():
    """Initialize all global variables with error handling"""
    global sendReport, reportTo, OrgId, scriptPath, showSendToInvolvement
    
    # Get script path
    try:
        scriptPath = model.ScriptName if hasattr(model, 'ScriptName') else 'SchedulerReport'
    except:
        scriptPath = 'SchedulerReport'
    
    # Initialize variables
    try:
        sendReport = model.Data.sendReport if hasattr(model.Data, 'sendReport') else 'n'
        reportTo = model.Data.reportTo if hasattr(model.Data, 'reportTo') else ''
        OrgId = str(Data.CurrentOrgId) if Data.CurrentOrgId else None
    except:
        sendReport = 'n'
        reportTo = ''
        OrgId = None
    
    # Check permissions
    showSendToInvolvement = 'n'
    try:
        for role in SendRoles.split(','):
            if model.UserIsInRole(role.strip()):
                showSendToInvolvement = 'y'
                break
    except:
        pass
    
    return sendReport, reportTo, OrgId, scriptPath, showSendToInvolvement

def validate_organization():
    """Validate organization ID and get org info"""
    global orgName
    
    if not OrgId:
        print('<div class="alert alert-danger">Error: No organization ID provided.</div>')
        raise SystemExit()
    
    try:
        org = model.GetOrganization(int(OrgId))
        if not org:
            raise ValueError("Organization not found")
        
        model.Header = "Schedule: {}".format(org.name)
        orgName = org.name
        return org
    except Exception as e:
        print('<div class="alert alert-danger">Error: Could not load organization {}: {}</div>'.format(OrgId, e))
        raise SystemExit()


# Override default User Config settings with Org Specfic ExtraValues if they exist.
def pull_org_extra_values():
    global ShowEmptySlots, ShowFamilyButtons, MinVolunteerAge, FromAddress, Subject, Title
   
    if Title == '':
       Title = orgName + " Schedule"
       
    evsql = """
    SELECT 
    *
    FROM OrganizationExtra
    WHERE OrganizationId = @OrgId
    AND Field IN (
        'SchedulerReportShowEmptySlots',
        'SchedulerReportShowFamilyButtons',
        'SchedulerReportMinVolunteerAge',
        'SchedulerReportFromAddress',
        'SchedulerReportSubject',
        'SchedulerReportTitle'
        )
    """
    
    extra_values = q.QuerySql(evsql, {'OrgId': OrgId})
    
    for row in extra_values:
        if row.Field == 'SchedulerReportShowEmptySlots':
            ShowEmptySlots = row.BitValue

        elif row.Field == 'SchedulerReportShowFamilyButtons':
            ShowFamilyButtons = row.IntValue

        elif row.Field == 'SchedulerReportMinVolunteerAge' and row.IntValue != 0:
            MinVolunteerAge = row.IntValue

        elif row.Field == 'SchedulerReportFromAddress' and row.Data != None:
            FromAddress = row.Data

        elif row.Field == 'SchedulerReportSubject' and row.Data != None:
            Subject = row.Data

        elif row.Field == 'SchedulerReportTitle' and row.Data != None:
            Title = row.Data

    return

def get_schedule_sql():
    """Return the main schedule query"""
    return '''
DECLARE @InvId Int = {0}
DECLARE @days int = {1}

;WITH
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
    FORMAT(tsm.MeetingDateTime, 'h:mm tt') as ServiceTime,
    FORMAT(tsm.MeetingDateTime, 'M/d/yy h:mm tt') as ServiceDateTime,
    FORMAT(tsm.MeetingDateTime, 'dddd, MMMM dd, yyyy') AS DateOnly,
    DATEPART(WEEKDAY, tsm.MeetingDateTime) as DayOfWeek,
    tssg.NumberVolunteersNeeded as Needed,
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
	AND (tssg.IsDeleted = 'False' OR tssg.IsDeleted IS NULL OR tssg.TimeSlotMeetingTeamSubGroupId IS NULL)
    AND OrganizationId = {0}
),
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
        STRING_AGG(Name2, '; ') WITHIN GROUP (ORDER BY Name2) as Names
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
WHERE SubGroupTeam NOT IN ('EMS', 'LEO')
ORDER BY ServiceDate, ServiceTime, SubGroupTeam
'''

def get_family_sql():
    """Return SQL for family members based on ShowFamilyButtons setting"""
    if ShowFamilyButtons == 2:
        # Only show family members who are in this involvement
        return '''
        SELECT DISTINCT p.PeopleId AS FamilyMemberID, p.name AS FamilyMemberName
        FROM dbo.People AS p
        JOIN People p1 ON p1.FamilyId = p.FamilyId
        JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
        WHERE p1.PeopleId = {0} 
            AND om.OrganizationId = {1}
            AND om.MemberTypeId NOT IN (230) -- Exclude inactive members
            AND NOT (p.IsDeceased = 1) 
            AND NOT (p.ArchivedFlag = 1)
            AND p.PeopleId != {0}
            AND (p.Age >= {2} OR (p.Age IS NULL))
        ORDER BY p.name
        '''
    else:
        # Show all family members (ShowFamilyButtons == 1)
        return '''
        SELECT DISTINCT p.PeopleId AS FamilyMemberID, p.name AS FamilyMemberName
        FROM dbo.People AS p
        JOIN People p1 ON p1.FamilyId = p.FamilyId
        WHERE p1.PeopleId = {0} 
            AND NOT (p.IsDeceased = 1) 
            AND NOT (p.ArchivedFlag = 1)
            AND p.PeopleId != {0}
            AND (p.Age >= {2} OR (p.Age IS NULL))
        ORDER BY p.name
        '''

def get_styles():
    """Return CSS styles based on layout choice"""
    base_styles = '''
<style>
body {
    font-family: Arial, sans-serif;
    color: #333;
}

.scheduler-container {
    max-width: 1200px;
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

.btn-small {
    display: inline-block;
    padding: 3px 6px;
    text-decoration: none;
    background-color: #3AAEE0;
    color: white;
    border-radius: 4px;
    font-weight: lighter;
    margin: 1px;
    font-size: 10px;
}

.action-buttons {
    text-align: center;
    margin: 10px 0;
}

.action-buttons-small {
    text-align: center;
    font-style: italic;
    margin: 5px 0;
}

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

.alert-danger {
    background-color: #f8d7da;
    color: #721c24;
    border: 1px solid #f5c6cb;
}

.print-header {
    display: none;
}
'''

    if UseMultiColumn:
        # Multi-column styles
        return base_styles + '''
/* Multi-column layout */
.date-section {
    margin: 10px 0;
    clear: both;
}

.date-header {
    background-color: #e9ecef;
    padding: 5px 8px;
    font-weight: bold;
    font-size: 14px;
    margin-bottom: 6px;
    border-radius: 3px;
}

/* Grid layout styles */
.time-grid {
    background: #f8f9fa;
    border: 1px solid #dee2e6;
    border-radius: 3px;
    padding: 8px;
}

.time-column-container {
    display: flex;
    gap: 20px;
}

.time-column {
    flex: 1;
    min-width: 0;
}

.time-group {
    margin-bottom: 3px;
}

.schedule-line {
    font-size: 11px;
    line-height: 1.3;
    padding: 1px 0;
    word-wrap: break-word;
}

.schedule-line strong {
    font-weight: bold;
    color: #495057;
}

.status-cell {
    font-size: 11px;
    font-weight: bold;
}

.status-full { color: #28a745; }
.status-partial { color: #ffc107; }
.status-empty { color: #dc3545; }
.status-required { color: #dc3545; }

.volunteer-list {
    font-size: 11px;
    line-height: 1.3;
}

.signup-link {
    color: #3AAEE0;
    text-decoration: none;
    font-size: 11px;
}

.signup-link:hover {
    text-decoration: underline;
}

@media (max-width: 768px) {
    /* Force single column on mobile */
    .time-column-container {
        flex-direction: column !important;
        gap: 5px !important;
    }
    
    .time-column {
        width: 100% !important;
        margin-bottom: 0;
    }
    
    .time-group {
        margin-bottom: 5px;
    }
}

@media print {
    .action-buttons, .action-buttons-small, .no-print {
        display: none !important;
    }
    
    /* Hide h1 but show print header */
    h1 {
        display: none !important;
    }
    
    .print-header {
        display: block !important;
        font-size: 16px;
        font-weight: bold;
        text-align: center;
        margin-bottom: 10px;
    }
    
    /* Maintain grid layout for print */
    .time-column-container {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        gap: 15px !important;
    }
    
    .time-column {
        flex: 1 1 auto !important;
        width: auto !important;
        break-inside: avoid;
        page-break-inside: avoid;
    }
    
    .time-group {
        break-inside: avoid;
        page-break-inside: avoid;
    }
    
    .date-section {
        break-before: auto;
        page-break-before: auto;
    }
    
    .time-label { font-size: 11px !important; }
    .schedule-line { font-size: 10px !important; }
    
    /* Force color and background printing */
    * {
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
        color-adjust: exact !important;
    }
    
    /* Ensure backgrounds print */
    .time-grid {
        background: #f8f9fa !important;
        -webkit-print-color-adjust: exact !important;
    }
    
    .date-header {
        background-color: #e9ecef !important;
        -webkit-print-color-adjust: exact !important;
    }
}
</style>
'''
    else:
        # Table layout styles
        return base_styles + '''
/* Table layout */
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

.date-header-row {
    background-color: #e9ecef;
    font-weight: bold;
}

.date-header-row td {
    padding: 8px 5px;
    font-size: 14px;
    border-bottom: 2px solid #dee2e6;
}

.status-cell {
    text-align: center;
    font-weight: bold;
}

.status-full { color: #28a745; }
.status-partial { color: #ffc107; }
.status-empty { color: #dc3545; }
.status-required { color: #dc3545; }

.signup-link {
    color: #3AAEE0;
    text-decoration: none;
}

.signup-link:hover {
    text-decoration: underline;
}

@media print {
    .action-buttons, .action-buttons-small, .no-print {
        display: none;
    }
    
    /* Hide h1 but show print header */
    h1 {
        display: none !important;
    }
    
    .print-header {
        display: block !important;
        font-size: 16px;
        font-weight: bold;
        text-align: center;
        margin-bottom: 10px;
    }
    
    .schedule-table {
        font-size: 11px;
    }
    
    .schedule-table td {
        padding: 2px 3px;
    }
    
    /* Force color and background printing */
    * {
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
        color-adjust: exact !important;
    }
    
    /* Ensure backgrounds print */
    .date-header-row {
        background-color: #e9ecef !important;
        -webkit-print-color-adjust: exact !important;
    }
    
    .schedule-table th {
        background-color: #f8f9fa !important;
        -webkit-print-color-adjust: exact !important;
    }
    
    /* Ensure status colors print */
    .status-full { color: #28a745 !important; }
    .status-partial { color: #ffc107 !important; }
    .status-empty { color: #dc3545 !important; }
    .status-required { color: #dc3545 !important; font-weight: bold !important; }
}
</style>
'''

def build_family_buttons():
    """Build family member schedule buttons"""
    if ShowFamilyButtons == 0:
        return ''
    
    try:
        family_sql = get_family_sql()
        family = q.QuerySql(family_sql.format(model.UserPeopleId, OrgId, MinVolunteerAge))
        
        if not family:
            return ''
        
        buttons_html = '<div class="action-buttons-small no-print">'
        buttons_html += '<span style="font-size: 11px; color: #666;">Manage schedules for: </span>'
        
        for row in family:
            buttons_html += '''
            <a href="{}/OnlineReg/{}?peopleId={}" class="btn-small" target="_blank" rel="noopener noreferrer">
                {}
            </a>
            '''.format(model.CmsHost, OrgId, row.FamilyMemberID, row.FamilyMemberName)
        
        buttons_html += '</div>'
        return buttons_html
        
    except Exception as e:
        return '<div class="alert alert-danger">Error loading family info: {}</div>'.format(e)

def build_schedule_table(results):
    """Build schedule in table format"""
    html = '''
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
            html += '''
            <tr class="date-header-row">
                <td colspan="4">{}</td>
            </tr>
            '''.format(current_date)
        
        # Skip empty slots if configured
        if not ShowEmptySlots and row.Serving == 0:
            continue
        
        # Determine status
        status_class, status_text = get_status_info(row)
        
        # Build volunteer column
        volunteer_cell = row.Names if row.Names else '<a href="{}/OnlineReg/{}" class="signup-link" target="_blank">Click to sign up</a>'.format(model.CmsHost, OrgId)
        
        # Add row
        html += '''
        <tr>
            <td>{}</td>
            <td>{}</td>
            <td class="status-cell {}">{}</td>
            <td>{}</td>
        </tr>
        '''.format(row.ServiceTime, row.SubGroupTeam, status_class, status_text, volunteer_cell)
    
    html += '</tbody></table>'
    return html

def build_schedule_columns(results):
    """Build schedule in multi-column format organized by time"""
    html = ''
    current_date = None
    date_data = {}
    date_sort_keys = {}  # Store actual dates for sorting
    
    # Group by date and time
    for row in results:
        date_key = str(row.DateOnly)
        if date_key not in date_data:
            date_data[date_key] = {}
            date_sort_keys[date_key] = row.ServiceDate  # Store actual date for sorting
        
        time_key = row.ServiceTime
        if time_key not in date_data[date_key]:
            date_data[date_key][time_key] = []
        
        # Skip empty slots if configured
        if not ShowEmptySlots and row.Serving == 0:
            continue
            
        date_data[date_key][time_key].append(row)
    
    # Build output for each date, sorted by actual date
    sorted_dates = sorted(date_data.keys(), key=lambda x: date_sort_keys[x])
    for date_header in sorted_dates:
        times = date_data[date_header]
        if times:  # Only show dates with data
            html += build_date_section_grid(date_header, times)
    
    return html

def build_date_section_grid(date_header, times_dict):
    """Build a date section with time-based grid layout"""
    html = '''
    <div class="date-section">
        <div class="date-header">{}</div>
        <div class="time-grid">
    '''.format(date_header)
    
    # Convert time strings to sortable format and sort chronologically
    time_list = []
    for time_str in times_dict.keys():
        # Parse time string (e.g., "7:30 AM" or "11:00 AM")
        try:
            from datetime import datetime
            dt = datetime.strptime(time_str, '%I:%M %p')
            time_list.append((dt, time_str))
        except:
            # Fallback if parsing fails
            time_list.append((time_str, time_str))
    
    # Sort by actual time value
    time_list.sort(key=lambda x: x[0])
    sorted_times = [t[1] for t in time_list]
    
    # Calculate how many columns we can fit
    items_per_column = max(1, len(sorted_times) // ColumnsPerRow)
    if len(sorted_times) % ColumnsPerRow != 0:
        items_per_column += 1
    
    html += '<div class="time-column-container">'
    
    col_count = 0
    for i, time in enumerate(sorted_times):
        if i % items_per_column == 0:
            if i > 0:
                html += '</div>'
            html += '<div class="time-column">'
            col_count += 1
        
        html += '<div class="time-group">'
        
        for slot in times_dict[time]:
            status_class, status_text = get_status_info(slot)
            
            # Build compact single line with time inline
            if slot.Names:
                line_content = '<strong>{}</strong> - {} <span class="{}">({})</span> - {}'.format(
                    time, slot.SubGroupTeam, status_class, status_text, slot.Names)
            else:
                line_content = '<strong>{}</strong> - {} <span class="{}">({})</span> - <a href="{}/OnlineReg/{}" class="signup-link" target="_blank">Sign up</a>'.format(
                    time, slot.SubGroupTeam, status_class, status_text, model.CmsHost, OrgId)
            
            html += '<div class="schedule-line">{}</div>'.format(line_content)
        
        html += '</div>'
    
    html += '</div></div></div></div>'
    return html

def build_date_section(date_header, slots):
    """Build a date section with time slots"""
    html = '''
    <div class="date-section">
        <div class="date-header">{}</div>
        <div class="time-slots-container">
    '''.format(date_header)
    
    for slot in slots:
        status_class, status_text = get_status_info(slot)
        
        volunteer_html = slot.Names if slot.Names else '<a href="{}/OnlineReg/{}" class="signup-link" target="_blank">Sign up</a>'.format(model.CmsHost, OrgId)
        
        html += '''
        <div class="time-slot">
            <div class="slot-header">
                <span class="slot-time">{}</span> - <span class="slot-team">{}</span>
                <span class="status-cell {}">{}</span>
            </div>
            <div class="volunteer-list">{}</div>
        </div>
        '''.format(slot.ServiceTime, slot.SubGroupTeam, status_class, status_text, volunteer_html)
    
    html += '''
        </div>
    </div>
    '''
    return html

def get_status_info(row):
    """Determine status class and text for a row"""
    if row.Status == 'Full':
        return 'status-full', 'Full'
    elif row.Status == 'Partial':
        return 'status-partial', '{}/{}'.format(row.Serving, row.Needed)
    elif row.Status == 'Required - Empty':
        return 'status-required', 'NEEDED!'
    else:
        return 'status-empty', '0/{}'.format(row.Needed if row.Needed else '?')

def send_email(to_involvement=False):
    """Handle email sending"""
    global NMReport, UseMultiColumn
    
    # Force single column for email
    original_use_multi_column = UseMultiColumn
    UseMultiColumn = False
    
    try:
        # Add tracking
        NMReport += '{track}{tracklinks}<br />'
        
        # Set subject
        email_subject = Subject if Subject else "Schedule: {}".format(orgName)
        
        if to_involvement:
            # Send to involvement
            QueuedBy = model.UserPeopleId
            
            # Get involvement members
            sql = '''
            SELECT p.PeopleId, p.Name, p.EmailAddress 
            FROM OrganizationMembers om
            JOIN People p ON om.PeopleId = p.PeopleId
            WHERE om.OrganizationId = {} 
                AND om.Pending = 0 
                AND p.EmailAddress IS NOT NULL
                AND p.EmailAddress != ''
            '''.format(OrgId)
            
            recipients = q.QuerySql(sql)
            sent_count = 0
            error_count = 0
            
            for recipient in recipients:
                try:
                    model.Email(recipient.PeopleId, QueuedBy, FromAddress, FromName, email_subject, NMReport)
                    sent_count += 1
                except:
                    error_count += 1
            
            message = 'Report sent to {} involvement members.'.format(sent_count)
            if error_count > 0:
                message += ' ({} errors)'.format(error_count)
            
            print('<div class="alert alert-info">{}</div>'.format(message))
        else:
            # Send to self
            QueuedBy = model.UserPeopleId
            MailToQuery = model.UserPeopleId
            
            model.Email(MailToQuery, QueuedBy, FromAddress, FromName, email_subject, NMReport)
            print('<div class="alert alert-info">Report sent to your email address.</div>')
            
    except Exception as e:
        print('<div class="alert alert-danger">Error sending email: {}</div>'.format(e))

def main():
    """Main execution function"""
    global NMReport
    
    # Initialize
    sendReport, reportTo, OrgId, scriptPath, showSendToInvolvement = initialize_variables()
    
    # Validate org
    org = validate_organization()
    
    pull_org_extra_values()
    
    # Get styles
    styles = get_styles()
    
    # Build initial template
    template = '''{}
    <div class="scheduler-container">
    <div class="print-header">{} </div>
    <h1>{} </h1>
    
    <div class="action-buttons no-print">
        <a href="{}/OnlineReg/{}" class="btn" target="_blank" rel="noopener noreferrer">
            Manage My Commitments
        </a>
    </div>
    {}
'''.format(styles, Title, Title, model.CmsHost, OrgId, build_family_buttons())
    
    # Execute query
    try:
        sql = get_schedule_sql()
        results = q.QuerySql(sql.format(OrgId, ScheduleDays))
        
        if not results:
            template += '''
            <div class="alert alert-info">
                No scheduled services found for the next {} days.
            </div>
            '''.format(ScheduleDays)
        else:
            # Build schedule based on layout choice
            if UseMultiColumn:
                template += build_schedule_columns(results)
            else:
                template += build_schedule_table(results)
    
    except Exception as e:
        template += '''
        <div class="alert alert-danger">
            Error loading schedule data: {}
        </div>
        '''.format(e)
    
    # Close container
    template += '</div>'
    
    # Required SQL for BlueToolbar
    sql2 = '''
    SELECT p.PeopleId FROM dbo.People p JOIN dbo.TagPerson tp ON tp.PeopleId = p.PeopleId AND tp.Id = @BlueToolbarTagId
    '''
    
    # Render template
    NMReport = model.RenderTemplate(template)
    
    # Handle report sending
    if sendReport == 'y' and reportTo == 'Self':
        send_email(to_involvement=False)
    elif sendReport == 'y' and reportTo == 'Involvement':
        if showSendToInvolvement != 'y':
            print('<div class="alert alert-danger">You do not have permission to send to involvement.</div>')
        else:
            send_email(to_involvement=True)
    else:
        # Show send buttons
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
        
        # Insert buttons after title
        template = template.replace('</h1>', '</h1>\n' + action_buttons)
        
        print(template)

# Execute main function
main()
