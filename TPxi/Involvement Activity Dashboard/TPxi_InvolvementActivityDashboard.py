#roles=Admin
import datetime
import json
import traceback

model.Title = "Involvement Actiity Dashboard"

# ==============================================================
# CONFIGURATION VARIABLES - Edit these as needed
# ==============================================================
# Number of days to consider for inactive/active calculations
INACTIVE_DAYS = 90      # Organizations with no activity in this many days will be marked inactive
RECENT_DAYS = 30        # Consider activities within this many days as "recent"
NEW_ORG_DAYS = 180      # Organizations created within this many days are considered "new"
MEMBER_CHANGE_DAYS = 45 # Consider member changes within this many days as activity
WEIGHT_MEMBER_CHANGES = 1.5  # Weight for member changes in member-management orgs

# Color scheme for charts and indicators
COLOR_ACTIVE = "#28a745"     # Green for active
COLOR_INACTIVE = "#dc3545"   # Red for inactive
COLOR_WARNING = "#ffc107"    # Yellow for warning
COLOR_INFO = "#17a2b8"       # Blue for info
COLOR_NEUTRAL = "#6c757d"    # Gray for neutral

# ==============================================================
# UTILITY FUNCTIONS
# ==============================================================
def format_number(number):
    """Format a number with commas for thousands"""
    try:
        if number is None:
            return "0"
        return "{:,}".format(number)
    except:
        return "0"

def format_date(date_obj):
    """Format a date as MM/DD/YYYY"""
    try:
        if isinstance(date_obj, basestring):
            # Try to convert string to date
            date_formats = [
                '%Y-%m-%d %H:%M:%S.%f',
                '%Y-%m-%d %H:%M:%S',
                '%Y-%m-%d',
                '%m/%d/%Y %H:%M:%S',
                '%m/%d/%Y %I:%M:%S %p',  # AM/PM format
                '%m/%d/%Y'
            ]
            
            for fmt in date_formats:
                try:
                    date_obj = datetime.datetime.strptime(date_obj, fmt)
                    return date_obj.strftime('%m/%d/%Y')
                except ValueError:
                    continue
                
            # If nothing worked, check if it's just a date part
            try:
                date_part = date_obj.split(' ')[0]
                return datetime.datetime.strptime(date_part, '%m/%d/%Y').strftime('%m/%d/%Y')
            except:
                return date_obj
        else:
            # Handle datetime objects directly
            return date_obj.strftime('%m/%d/%Y')
    except:
        return date_obj

def format_meeting_date(meeting_date):
    """Safely format a meeting date for chart use"""
    try:
        # First check if it's already a datetime object
        if isinstance(meeting_date, datetime.datetime):
            return meeting_date.strftime('%Y-%m-%d')
            
        # Try parsing the string
        date_formats = [
            '%Y-%m-%d %H:%M:%S.%f',
            '%Y-%m-%d %H:%M:%S',
            '%Y-%m-%d',
            '%m/%d/%Y %H:%M:%S',
            '%m/%d/%Y %I:%M:%S %p',  # AM/PM format
            '%m/%d/%Y'
        ]
            
        for fmt in date_formats:
            try:
                return datetime.datetime.strptime(meeting_date, fmt).strftime('%Y-%m-%d')
            except ValueError:
                continue
                
        # Try extracting date part as last resort
        parts = meeting_date.split(' ')[0]
        return datetime.datetime.strptime(parts, '%m/%d/%Y').strftime('%Y-%m-%d')
    except:
        # Add debugging output
        print("<!-- DEBUG: Failed to parse date: " + str(meeting_date) + " -->")
        return None

def get_date_days_ago(days):
    """Get a date N days ago from today"""
    return (datetime.datetime.now() - datetime.timedelta(days=days)).strftime('%Y-%m-%d')

def is_active(last_activity_date):
    """Determine if an organization is active based on its last activity date"""
    if not last_activity_date:
        return False
    
    try:
        if isinstance(last_activity_date, basestring):
            last_activity_date = datetime.datetime.strptime(last_activity_date, '%Y-%m-%d %H:%M:%S.%f')
        
        # Current date for comparison
        now = datetime.datetime.now()
        
        # Future dates are considered active
        if last_activity_date > now:
            return True
            
        days_since_last_activity = (now - last_activity_date).days
        return days_since_last_activity <= INACTIVE_DAYS
    except:
        return False

def get_activity_level(activity_score):
    """Determine activity level based on activity score"""
    if activity_score >= 7:
        return "high"
    elif activity_score >= 3:
        return "moderate"
    elif activity_score > 0:
        return "low"
    else:
        return "inactive"

def generate_sql_structure_query():
    """Generate the SQL query for organization structure data with enhanced activity metrics"""
    return """
    SELECT 
        p.Name AS Program,
        d.Name AS Division,
        o.OrganizationStatusId AS OrgStatusId,
        CASE WHEN o.OrganizationStatusId = 30 THEN 'Active' ELSE 'Inactive' END AS OrgStatus,
        o.OrganizationName AS Organization,
        o.MemberCount AS Members,
        o.CreatedDate,
        -- Add a field to detect member-management organizations based on org type and meeting patterns
        CASE WHEN (o.OrganizationTypeId IN (SELECT Id FROM lookup.OrganizationType 
                                           WHERE Description LIKE '%Group%' 
                                              OR Description LIKE '%Team%'
                                              OR Description LIKE '%Committee%'
                                              OR Description LIKE '%Roster%')) 
             OR ((SELECT COUNT(*) FROM Meetings m WHERE m.OrganizationId = o.OrganizationId) = 0)
             THEN 1 ELSE 0 END AS IsMemberManagement,
        
        (SELECT COUNT(*) FROM OrganizationMembers om 
         WHERE om.OrganizationId = o.OrganizationId 
         AND om.InactiveDate IS NOT NULL) AS Previous,
        (SELECT COUNT(*) FROM OrganizationMembers om 
         JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
         JOIN lookup.AttendType att ON mt.AttendanceTypeId = att.Id
         WHERE om.OrganizationId = o.OrganizationId 
         AND att.Description LIKE '%Visit%') AS Visitors,
        (SELECT COUNT(*) FROM Meetings m 
         WHERE m.OrganizationId = o.OrganizationId) AS Meetings,
        (SELECT MAX(m.MeetingDate) FROM Meetings m 
         WHERE m.OrganizationId = o.OrganizationId) AS LastMeetingDate,
        (SELECT COUNT(*) FROM Meetings m 
         WHERE m.OrganizationId = o.OrganizationId 
         AND m.MeetingDate >= DATEADD(day, -{0}, GETDATE())) AS RecentMeetings,
        -- New activity metrics
        (SELECT MAX(om.EnrollmentDate) FROM OrganizationMembers om 
         WHERE om.OrganizationId = o.OrganizationId) AS LastMemberAddedDate,
        (SELECT COUNT(*) FROM OrganizationMembers om 
         WHERE om.OrganizationId = o.OrganizationId 
         AND om.EnrollmentDate >= DATEADD(day, -{1}, GETDATE())) AS NewMembersCount,
        (SELECT COUNT(*) FROM OrganizationMembers om 
         WHERE om.OrganizationId = o.OrganizationId 
         AND om.InactiveDate >= DATEADD(day, -{1}, GETDATE())) AS RecentInactiveCount,
        -- detect member-management organizations without using MeetingFreq
        CASE WHEN (o.OrganizationTypeId IN (SELECT Id FROM lookup.OrganizationType 
                                           WHERE Description LIKE '%Group%' 
                                              OR Description LIKE '%Team%'
                                              OR Description LIKE '%Committee%'
                                              OR Description LIKE '%Roster%')) 
             OR ((SELECT COUNT(*) FROM Meetings m WHERE m.OrganizationId = o.OrganizationId) = 0)
             THEN 1 ELSE 0 END AS IsMemberManagement,
        p.Id AS ProgId,
        d.Id AS DivId,
        o.OrganizationId AS OrgId,
        ot.Code AS OrgType,
        ot.Description AS OrgTypeDesc
    FROM Organizations o
    JOIN Division d ON o.DivisionId = d.Id
    JOIN Program p ON d.ProgId = p.Id
    LEFT JOIN lookup.OrganizationType ot ON o.OrganizationTypeId = ot.Id
    ORDER BY p.Name, d.Name, o.OrganizationName
    """.format(RECENT_DAYS, MEMBER_CHANGE_DAYS)

def generate_sql_activity_detail_query():
    """Generate SQL query for detailed activity metrics"""
    return """
    SELECT 
        o.OrganizationId,
        o.OrganizationName,
        o.MemberCount,
        o.CreatedDate,
        (SELECT MAX(m.MeetingDate) FROM Meetings m 
         WHERE m.OrganizationId = o.OrganizationId) AS LastMeetingDate,
        (SELECT COUNT(*) FROM Meetings m 
         WHERE m.OrganizationId = o.OrganizationId 
         AND m.MeetingDate >= DATEADD(day, -{0}, GETDATE())) AS RecentMeetings,
        (SELECT MAX(om.EnrollmentDate) FROM OrganizationMembers om 
         WHERE om.OrganizationId = o.OrganizationId) AS LastMemberAddedDate,
        (SELECT COUNT(*) FROM OrganizationMembers om 
         WHERE om.OrganizationId = o.OrganizationId 
         AND om.EnrollmentDate >= DATEADD(day, -{1}, GETDATE())) AS NewMembersCount,
        (SELECT COUNT(*) FROM OrganizationMembers om 
         WHERE om.OrganizationId = o.OrganizationId 
         AND om.InactiveDate IS NOT NULL 
         AND om.InactiveDate <= GETDATE()) AS RecentInactiveCount,
        (SELECT MAX(om.EnrollmentDate) FROM OrganizationMembers om 
         WHERE om.OrganizationId = o.OrganizationId) AS LastEnrollmentDate,
        (SELECT MAX(om.InactiveDate) FROM OrganizationMembers om 
         WHERE om.OrganizationId = o.OrganizationId) AS LastInactiveDate,
        (SELECT MAX(om.InactiveDate) FROM OrganizationMembers om 
         WHERE om.OrganizationId = o.OrganizationId) AS LastMemberInactiveDate,
        CASE
            WHEN o.CreatedDate >= DATEADD(day, -{2}, GETDATE()) THEN 1
            ELSE 0
        END AS IsNewOrg
    FROM Organizations o
    WHERE o.OrganizationStatusId = 30
    ORDER BY o.OrganizationName
    """.format(RECENT_DAYS, MEMBER_CHANGE_DAYS, NEW_ORG_DAYS)

def generate_sql_involvement_type_query():
    """Generate SQL query for involvement type distribution"""
    return """
    SELECT 
        ot.Code AS OrgType, 
        ot.Description AS OrgTypeDesc,
        COUNT(*) AS OrgCount,
        SUM(CASE WHEN o.OrganizationStatusId = 30 THEN 1 ELSE 0 END) AS ActiveCount,
        SUM(CASE WHEN o.OrganizationStatusId <> 30 THEN 1 ELSE 0 END) AS InactiveCount,
        SUM(o.MemberCount) AS TotalMembers
    FROM Organizations o
    LEFT JOIN lookup.OrganizationType ot ON o.OrganizationTypeId = ot.Id
    GROUP BY ot.Code, ot.Description
    ORDER BY OrgCount DESC
    """

def generate_sql_meetings_query():
    """Generate SQL query for meeting data"""
    return """
    SELECT TOP 100
        m.MeetingId,
        m.OrganizationId,
        o.OrganizationName,
        m.MeetingDate,
        m.NumPresent,
        m.Description,
        m.Location,
        m.HeadCount,
        m.DidNotMeet
    FROM Meetings m
    JOIN Organizations o ON m.OrganizationId = o.OrganizationId
    WHERE m.MeetingDate <= GETDATE()
    ORDER BY m.MeetingDate DESC
    """

def generate_sql_member_changes_query():
    """Generate SQL query for recent member changes"""
    return '''
    SELECT TOP 2500
        om.OrganizationId,
        o.OrganizationName,
        p.Name AS PersonName,
        om.EnrollmentDate,
        om.InactiveDate,
        CASE 
            WHEN om.InactiveDate IS NOT NULL AND om.InactiveDate <= GETDATE() THEN 'Inactive'
            ELSE 'Active'
        END AS MemberStatus,
        mt.Description AS MemberType
    FROM OrganizationMembers om
    JOIN Organizations o ON om.OrganizationId = o.OrganizationId
    JOIN People p ON om.PeopleId = p.PeopleId
    LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
    WHERE (om.EnrollmentDate >= DATEADD(day, -{0}, GETDATE())
        OR (om.InactiveDate IS NOT NULL AND om.InactiveDate <= GETDATE()))
    ORDER BY 
        CASE 
            WHEN om.EnrollmentDate >= DATEADD(day, -{0}, GETDATE()) THEN om.EnrollmentDate
            ELSE om.InactiveDate
        END DESC
    '''.format(MEMBER_CHANGE_DAYS)

def generate_sql_inactive_orgs_query():
    """Generate SQL query for orgs without recent activity"""
    return '''
    -- Organizations with no recent activity
    SELECT
        o.OrganizationId,
        o.OrganizationName,
        p.Name AS Program,
        d.Name AS Division,
        p.Id AS ProgId,
        d.Id AS DivId,
        o.MemberCount,
        o.CreatedDate,
        ot.Description AS OrgTypeDesc,
        ot.Code AS OrgType,
        (SELECT MAX(m.MeetingDate) FROM Meetings m WHERE m.OrganizationId = o.OrganizationId) AS LastMeetingDate,
        (SELECT MAX(om.EnrollmentDate) FROM OrganizationMembers om WHERE om.OrganizationId = o.OrganizationId) AS LastMemberAddedDate,
        (SELECT MAX(om.InactiveDate) FROM OrganizationMembers om 
         WHERE om.OrganizationId = o.OrganizationId 
         AND om.InactiveDate <= GETDATE()) AS LastMemberInactiveDate
    FROM Organizations o
    JOIN Division d ON o.DivisionId = d.Id
    JOIN Program p ON d.ProgId = p.Id
    LEFT JOIN lookup.OrganizationType ot ON o.OrganizationTypeId = ot.Id
    WHERE o.OrganizationStatusId = 30
    AND (
        -- Most recent meeting is older than threshold or doesn't exist
        (SELECT MAX(m.MeetingDate) FROM Meetings m WHERE m.OrganizationId = o.OrganizationId) IS NULL 
        OR (SELECT MAX(m.MeetingDate) FROM Meetings m WHERE m.OrganizationId = o.OrganizationId) < DATEADD(day, -{0}, GETDATE())
    )
    AND (
        -- Most recent member add is older than threshold or doesn't exist
        (SELECT MAX(om.EnrollmentDate) FROM OrganizationMembers om WHERE om.OrganizationId = o.OrganizationId) IS NULL
        OR (SELECT MAX(om.EnrollmentDate) FROM OrganizationMembers om WHERE om.OrganizationId = o.OrganizationId) < DATEADD(day, -{0}, GETDATE())
    )
    AND (
        -- Most recent member inactivation is older than threshold or doesn't exist
        (SELECT MAX(om.InactiveDate) FROM OrganizationMembers om 
         WHERE om.OrganizationId = o.OrganizationId 
         AND om.InactiveDate <= GETDATE()) IS NULL
        OR (SELECT MAX(om.InactiveDate) FROM OrganizationMembers om 
            WHERE om.OrganizationId = o.OrganizationId 
            AND om.InactiveDate <= GETDATE()) < DATEADD(day, -{0}, GETDATE())
    )
    -- Only include orgs that are older than the threshold
    AND (o.CreatedDate IS NULL OR o.CreatedDate < DATEADD(day, -{0}, GETDATE()))
    ORDER BY o.MemberCount DESC
    '''.format(INACTIVE_DAYS)
    
# Add a new SQL query function to get recent member status changes for inactive orgs:
def generate_sql_inactive_member_changes_query():
    """Generate SQL query for recent member changes in inactive orgs"""
    return """
    SELECT TOP 100
        om.OrganizationId,
        o.OrganizationName,
        p.Name AS PersonName,
        om.EnrollmentDate,
        om.InactiveDate,
        CASE 
            WHEN om.InactiveDate IS NOT NULL THEN 'Inactive'
            ELSE 'Active'
        END AS MemberStatus,
        mt.Description AS MemberType
    FROM OrganizationMembers om
    JOIN Organizations o ON om.OrganizationId = o.OrganizationId
    JOIN People p ON om.PeopleId = p.PeopleId
    LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
    WHERE o.OrganizationId IN (
        SELECT o.OrganizationId
        FROM Organizations o
        WHERE o.OrganizationStatusId = 30
        AND DATEDIFF(day, COALESCE(
            (SELECT MAX(m.MeetingDate) FROM Meetings m WHERE m.OrganizationId = o.OrganizationId),
            (SELECT MAX(om.EnrollmentDate) FROM OrganizationMembers om WHERE om.OrganizationId = o.OrganizationId),
            (SELECT MAX(om.InactiveDate) FROM OrganizationMembers om WHERE om.OrganizationId = o.OrganizationId),
            o.CreatedDate
        ), GETDATE()) > {0}
    )
    ORDER BY 
        CASE 
            WHEN om.InactiveDate IS NOT NULL THEN om.InactiveDate
            ELSE om.EnrollmentDate
        END DESC
    """.format(INACTIVE_DAYS)

def convert_to_datetime(date_input):
    """
    Convert various date input formats to a datetime object
    
    Args:
        date_input: A date in various possible formats
    
    Returns:
        datetime object or None
    """
    if date_input is None:
        return None
    
    # If it's already a datetime object, return it
    if isinstance(date_input, datetime.datetime):
        return date_input
    
    # List of possible date formats to try
    date_formats = [
        '%Y-%m-%d %H:%M:%S.%f',  # Typical datetime format with microseconds
        '%Y-%m-%d %H:%M:%S',     # Datetime without microseconds
        '%Y-%m-%d',              # Date only
        '%m/%d/%Y %H:%M:%S',     # MM/DD/YYYY with time
        '%m/%d/%Y %I:%M:%S %p',  # MM/DD/YYYY with AM/PM time format
        '%m/%d/%Y'               # MM/DD/YYYY
    ]
    
    # If it's a string, try to parse it
    if isinstance(date_input, basestring):
        # Remove any 'Never' markers that might have been added for debugging
        if 'Never' in date_input:
            return None
            
        for fmt in date_formats:
            try:
                dt = datetime.datetime.strptime(date_input.strip(), fmt)
                # Check if date is in the future and handle appropriately
                if dt > datetime.datetime.now():
                    # You could either:
                    # 1. Return None for future dates (ignore them)
                    # return None
                    # 2. Or return the date as is (which will consider the org active)
                    return dt
                return dt
            except ValueError:
                continue
    
    return None


def find_most_recent_date(date_options):
    """
    Find the most recent date from a list of possible date options
    
    Args:
        date_options: A list of tuples (date_str, date_type)
    
    Returns:
        A tuple of (most_recent_date_str, most_recent_date_type) or None
    """
    valid_dates = []
    
    for date_str, date_type in date_options:
        if not date_str or date_str == 'Never' or date_str == 'Unknown':
            continue
        
        try:
            date_obj = convert_to_datetime(date_str)
            if date_obj:
                valid_dates.append((date_obj, date_str, date_type))
        except:
            continue
    
    if not valid_dates:
        return None
    
    # Sort by most recent date (descending order)
    valid_dates.sort(key=lambda x: x[0], reverse=True)
    
    # Return the most recent date information
    return (valid_dates[0][1], valid_dates[0][2])

# ==============================================================
# MAIN FUNCTION
# ==============================================================
def main():
    # Get URL parameters for filtering/view options
    view = model.Data.view if hasattr(model.Data, "view") else "overview"
    program_filter = model.Data.program if hasattr(model.Data, "program") else None
    division_filter = model.Data.division if hasattr(model.Data, "division") else None
    org_type_filter = model.Data.orgtype if hasattr(model.Data, "orgtype") else None
    
    try:
        # Execute queries
        org_structure_data = q.QuerySql(generate_sql_structure_query())
        activity_detail_data = q.QuerySql(generate_sql_activity_detail_query())
        involvement_type_data = q.QuerySql(generate_sql_involvement_type_query())
        meetings_data = q.QuerySql(generate_sql_meetings_query())
        member_changes_data = q.QuerySql(generate_sql_member_changes_query())
        inactive_orgs_data = q.QuerySql(generate_sql_inactive_orgs_query())
    
        # Calculate activity scores and add to data
        enhance_activity_data(org_structure_data, activity_detail_data)
    
        # Process data and render dashboard
        render_dashboard(view, org_structure_data, activity_detail_data, involvement_type_data, 
                         meetings_data, member_changes_data, inactive_orgs_data, 
                         program_filter, division_filter, org_type_filter)
    except Exception as e:
        # Print any errors
        print "<h2>Error</h2>"
        print "<p>An error occurred: " + str(e) + "</p>"
        print "<pre>"
        traceback.print_exc()
        print "</pre>"

def enhance_activity_data(org_structure_data, activity_detail_data):
    """Calculate comprehensive activity scores and enhance the data with activity metrics"""
    # Create a lookup of detailed activity data by org ID
    activity_lookup = {}
    for item in activity_detail_data:
        if hasattr(item, 'OrganizationId'):
            activity_lookup[item.OrganizationId] = item
    
    # Current date for calculations
    now = datetime.datetime.now()
    
    # Calculate combined activity metrics for each organization
    for org in org_structure_data:
        if not hasattr(org, 'OrgId'):
            continue
            
        # Start with zero score
        activity_score = 0
        last_activity_date = None
        
        # Check if we have detailed activity data for this org
        detail = activity_lookup.get(org.OrgId)
        
        # Factor 1: Recent meetings
        if hasattr(org, 'RecentMeetings') and org.RecentMeetings:
            activity_score += min(org.RecentMeetings * 2, 10)  # Up to 10 points for meetings
            
        # Factor 2: Recent member additions
        if hasattr(org, 'NewMembersCount') and org.NewMembersCount:
            # Check if this is a member management organization
            if hasattr(org, 'IsMemberManagement') and org.IsMemberManagement:
                # Higher weight for member-management orgs
                activity_score += min(org.NewMembersCount * 2.5, 10)  # Up to 10 points
            else:
                # Normal weight for regular orgs
                activity_score += min(org.NewMembersCount, 5)  # Up to 5 points
            
        # Factor 3: Recent member status changes
        if hasattr(org, 'RecentInactiveCount') and org.RecentInactiveCount:
            # Check if this is a member management organization
            if hasattr(org, 'IsMemberManagement') and org.IsMemberManagement:
                # Higher weight for member-management orgs
                activity_score += min(org.RecentInactiveCount * 2.0, 6)  # Up to 6 points
            else:
                # Normal weight for regular orgs
                activity_score += min(org.RecentInactiveCount, 3)  # Up to 3 points
        
        # Factor 4: New organization bonus
        if detail and hasattr(detail, 'IsNewOrg') and detail.IsNewOrg:
            activity_score += 5  # Bonus points for new organizations

        # Compensate for lack of meetings in member-management orgs
        if hasattr(org, 'IsMemberManagement') and org.IsMemberManagement:
            if not hasattr(org, 'RecentMeetings') or org.RecentMeetings == 0:
                activity_score += 2  # Bonus points to compensate for lack of meetings
        
        # Calculate most recent activity date
        dates_to_check = []
        
        # Last meeting date
        if hasattr(org, 'LastMeetingDate') and org.LastMeetingDate:
            try:
                meeting_date = convert_to_datetime(org.LastMeetingDate)
                if meeting_date:
                    dates_to_check.append(meeting_date)
            except:
                pass
        
        # Last member added date
        if hasattr(org, 'LastMemberAddedDate') and org.LastMemberAddedDate:
            try:
                member_added_date = convert_to_datetime(org.LastMemberAddedDate)
                if member_added_date:
                    dates_to_check.append(member_added_date)
            except:
                pass
            
        # Last member inactivated date  
        if hasattr(org, 'LastMemberInactiveDate') and org.LastMemberInactiveDate:
            try:
                member_inactive_date = convert_to_datetime(org.LastMemberInactiveDate)
                if member_inactive_date:
                    dates_to_check.append(member_inactive_date)
            except:
                pass
                
        # Creation date as fallback
        if hasattr(org, 'CreatedDate') and org.CreatedDate:
            try:
                created_date = convert_to_datetime(org.CreatedDate)
                if created_date:
                    dates_to_check.append(created_date)
            except:
                pass

        # Add this to the enhance_activity_data function
        if hasattr(org, 'IsMemberManagement') and org.IsMemberManagement:
            # For membership management orgs, give more weight to member changes
            if hasattr(org, 'NewMembersCount') and org.NewMembersCount:
                activity_score += min(org.NewMembersCount * WEIGHT_MEMBER_CHANGES, 10)
                
            if hasattr(org, 'RecentInactiveCount') and org.RecentInactiveCount:
                activity_score += min(org.RecentInactiveCount * WEIGHT_MEMBER_CHANGES, 5)
                
            # Reduce the penalty for no meetings
            if not hasattr(org, 'RecentMeetings') or not org.RecentMeetings:
                activity_score += 2  # Add points to compensate for lack of meetings
                
        # Use the most recent date
        if dates_to_check:
            last_activity_date = max(dates_to_check)
            
            # Calculate days since last activity
            if last_activity_date:
                days_since = (now - last_activity_date).days
                
                # Adjust score based on recency
                if days_since <= RECENT_DAYS:
                    activity_score += 3  # Bonus for very recent activity
                elif days_since <= INACTIVE_DAYS:
                    activity_score += 1  # Small bonus for somewhat recent activity
                else:
                    activity_score = max(0, activity_score - 3)  # Penalty for old activity
        
        # Add calculated fields to the org object
        org.ActivityScore = activity_score
        org.ActivityLevel = get_activity_level(activity_score)
        org.LastActivityDate = last_activity_date
        org.IsActive = activity_score > 0

def render_dashboard(view, org_structure_data, activity_detail_data, involvement_type_data, 
                    meetings_data, member_changes_data, inactive_orgs_data,
                    program_filter, division_filter, org_type_filter):
    """Render the dashboard based on the current view and filters"""
    
    # Build lookups for program and division names from org_structure_data
    program_lookup = {}  # id -> name
    division_lookup = {}  # id -> name
    
    for item in org_structure_data:
        if hasattr(item, 'ProgId') and hasattr(item, 'Program'):
            program_lookup[str(item.ProgId)] = item.Program
        if hasattr(item, 'DivId') and hasattr(item, 'Division'):
            division_lookup[str(item.DivId)] = item.Division
    
    # Apply filters to main data
    filtered_data = filter_data(org_structure_data, program_filter, division_filter, org_type_filter)
    
    # Create a filtered version of inactive_orgs_data
    filtered_inactive_data = []
    
    if inactive_orgs_data:
        if program_filter or division_filter or org_type_filter:
            for item in inactive_orgs_data:
                include = True
                
                # Program filter
                if program_filter and include:
                    prog_name = program_lookup.get(str(program_filter))
                    if prog_name and hasattr(item, 'Program') and item.Program != prog_name:
                        include = False
                
                # Division filter
                if division_filter and include:
                    div_name = division_lookup.get(str(division_filter))
                    if div_name and hasattr(item, 'Division') and item.Division != div_name:
                        include = False
                
                # Org Type filter
                if org_type_filter and include:
                    if hasattr(item, 'OrgType') and item.OrgType != org_type_filter:
                        include = False
                
                if include:
                    filtered_inactive_data.append(item)
        else:
            # No filters applied
            filtered_inactive_data = inactive_orgs_data
    
    # Print CSS and JavaScript for dashboard
    print_dashboard_header()
    
    # Print navigation
    print_navigation(view)
    
    # Print filters
    print_filter_controls(org_structure_data, involvement_type_data, program_filter, division_filter, org_type_filter)
    
    # Render appropriate view
    if view == "overview":
        render_overview(filtered_data, involvement_type_data, filtered_inactive_data)
    elif view == "programs":
        render_programs_view(filtered_data)
    elif view == "activity":
        render_activity_view(filtered_data, activity_detail_data)
    elif view == "meetings":
        render_meetings_view(meetings_data, filtered_data)
    elif view == "members":
        render_members_view(member_changes_data, filtered_data)
    elif view == "inactive":
        render_inactive_view(filtered_inactive_data, filtered_data)
    else:
        render_overview(filtered_data, involvement_type_data, filtered_inactive_data)

def filter_data(data, program_filter, division_filter, org_type_filter):
    """Filter data based on selected filters with support for different data structures"""
    if not data:
        return []
        
    filtered_data = []
    
    # First, create lookups for program and division names by ID
    program_names = {}  # id -> name
    division_names = {}  # id -> name
    
    # Build lookups from the data itself
    for item in data:
        # For program lookup
        if hasattr(item, 'ProgId') and hasattr(item, 'Program'):
            program_names[str(item.ProgId)] = item.Program
            
        # For division lookup
        if hasattr(item, 'DivId') and hasattr(item, 'Division'):
            division_names[str(item.DivId)] = item.Division
    
    for item in data:
        include = True
        
        # Handle Program filter - check by ID or name
        if program_filter:
            if hasattr(item, 'ProgId'):
                # Standard filtering by ID
                if str(item.ProgId) != str(program_filter):
                    include = False
            elif hasattr(item, 'Program'):
                # Filter by name for inactive orgs
                prog_name = program_names.get(str(program_filter))
                if prog_name and item.Program != prog_name:
                    include = False
        
        # Handle Division filter - check by ID or name
        if division_filter and include:
            if hasattr(item, 'DivId'):
                # Standard filtering by ID
                if str(item.DivId) != str(division_filter):
                    include = False
            elif hasattr(item, 'Division'):
                # Filter by name for inactive orgs
                div_name = division_names.get(str(division_filter))
                if div_name and item.Division != div_name:
                    include = False
        
        # Handle Org Type filter
        if org_type_filter and include:
            type_match = False
            
            if hasattr(item, 'OrgType'):
                # Direct match by code
                if item.OrgType == org_type_filter:
                    type_match = True
            
            # Try alternative field names that might exist in inactive orgs data
            if not type_match and hasattr(item, 'OrgTypeDesc'):
                # This is a common field in the inactive orgs data
                if item.OrgTypeDesc == org_type_filter:
                    type_match = True
                    
            if not type_match and hasattr(item, 'OrgType'):
                include = False
        
        if include:
            filtered_data.append(item)
    
    return filtered_data

def get_program_name(program_id, data):
    """Helper function to get program name from id"""
    for item in data:
        if hasattr(item, 'ProgId') and str(item.ProgId) == str(program_id) and hasattr(item, 'Program'):
            return item.Program
    return None

def get_division_name(division_id, data):
    """Helper function to get division name from id"""
    for item in data:
        if hasattr(item, 'DivId') and str(item.DivId) == str(division_id) and hasattr(item, 'Division'):
            return item.Division
    return None

def org_type_matches(org_type_desc, org_type_filter, data):
    """Helper function to check if org type description matches filter"""
    # First try direct match
    if org_type_desc == org_type_filter:
        return True
    
    # Then try to match by finding the description for the code
    for item in data:
        if hasattr(item, 'OrgType') and item.OrgType == org_type_filter and hasattr(item, 'OrgTypeDesc'):
            if org_type_desc == item.OrgTypeDesc:
                return True
    
    return False

def print_dashboard_header():
    """Print CSS and JavaScript for dashboard"""
    print """
    <style>
        /* Dashboard styles */
        .dashboard-container {
            font-family: Arial, sans-serif;
            max-width: 1200px;
            margin: 0 auto;
        }
        
        .card {
            background-color: #fff;
            border-radius: 5px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            margin-bottom: 20px;
            padding: 20px;
        }
        
        .card-header {
            border-bottom: 1px solid #eee;
            font-size: 18px;
            font-weight: bold;
            margin-bottom: 15px;
            padding-bottom: 10px;
        }
        
        .stat-grid {
            display: flex;
            flex-wrap: wrap;
            gap: 15px;
            margin-bottom: 20px;
        }
        
        .stat-box {
            background-color: #f8f9fa;
            border-left: 4px solid #007bff;
            border-radius: 4px;
            flex: 1;
            min-width: 200px;
            padding: 15px;
        }
        
        .stat-title {
            color: #6c757d;
            font-size: 14px;
            margin-bottom: 5px;
        }
        
        .stat-value {
            font-size: 24px;
            font-weight: bold;
        }
        
        .nav-tabs {
            border-bottom: 1px solid #dee2e6;
            display: flex;
            list-style: none;
            margin-bottom: 20px;
            padding-left: 0;
        }
        
        .nav-tabs li {
            margin-bottom: -1px;
        }
        
        .nav-tabs a {
            background-color: transparent;
            border: 1px solid transparent;
            border-top-left-radius: 4px;
            border-top-right-radius: 4px;
            color: #007bff;
            display: block;
            padding: 10px 15px;
            text-decoration: none;
        }
        
        .nav-tabs a:hover {
            border-color: #e9ecef #e9ecef #dee2e6;
        }
        
        .nav-tabs a.active {
            background-color: #fff;
            border-color: #dee2e6 #dee2e6 #fff;
            color: #495057;
            font-weight: bold;
        }
        
        .filter-controls {
            background-color: #f8f9fa;
            border-radius: 4px;
            margin-bottom: 20px;
            padding: 15px;
        }
        
        .filter-controls select {
            margin-right: 10px;
            padding: 5px;
        }
        
        .filter-controls button {
            background-color: #007bff;
            border: none;
            border-radius: 4px;
            color: white;
            cursor: pointer;
            padding: 6px 12px;
        }
        
        .filter-controls button:hover {
            background-color: #0069d9;
        }
        
        table {
            border-collapse: collapse;
            width: 100%;
        }
        
        th, td {
            border: 1px solid #dee2e6;
            padding: 8px;
            text-align: left;
        }
        
        th {
            background-color: #f8f9fa;
        }
        
        tr:nth-child(even) {
            background-color: #f8f9fa;
        }
        
        .status-active {
            background-color: #d4edda;
            border-radius: 4px;
            color: #155724;
            display: inline-block;
            font-size: 12px;
            padding: 3px 8px;
        }
        
        .status-inactive {
            background-color: #f8d7da;
            border-radius: 4px;
            color: #721c24;
            display: inline-block;
            font-size: 12px;
            padding: 3px 8px;
        }
        
        .status-warning {
            background-color: #fff3cd;
            border-radius: 4px;
            color: #856404;
            display: inline-block;
            font-size: 12px;
            padding: 3px 8px;
        }
        
        .chart-container {
            height: 300px;
            margin-bottom: 20px;
            width: 100%;
        }
        
        .progress {
            background-color: #e9ecef;
            border-radius: 4px;
            height: 20px;
            margin-bottom: 10px;
            overflow: hidden;
        }
        
        .progress-bar {
            background-color: #007bff;
            color: white;
            height: 100%;
            line-height: 20px;
            text-align: center;
        }
        
        .activity-high {
            background-color: #28a745;
        }
        
        .activity-moderate {
            background-color: #ffc107;
        }
        
        .activity-low {
            background-color: #dc3545;
        }
        
        .activity-inactive {
            background-color: #6c757d;
        }
        
        /* Activity gauge styles */
        .activity-gauge {
            display: inline-block;
            width: 100px;
            height: 10px;
            background-color: #e9ecef;
            border-radius: 5px;
            position: relative;
            overflow: hidden;
        }
        
        .activity-gauge-fill {
            height: 100%;
            border-radius: 5px;
        }
        
        .activity-level-high {
            color: #28a745;
        }
        
        .activity-level-moderate {
            color: #ffc107;
        }
        
        .activity-level-low {
            color: #dc3545;
        }
        
        .activity-level-inactive {
            color: #6c757d;
        }
        
        .badge {
            display: inline-block;
            padding: 0.25em 0.6em;
            font-size: 75%;
            font-weight: 700;
            line-height: 1;
            text-align: center;
            white-space: nowrap;
            vertical-align: baseline;
            border-radius: 0.25rem;
        }
        
        .badge-primary {
            color: #fff;
            background-color: #007bff;
        }
        
        .badge-secondary {
            color: #fff;
            background-color: #6c757d;
        }
        
        .badge-success {
            color: #fff;
            background-color: #28a745;
        }
        
        .badge-danger {
            color: #fff;
            background-color: #dc3545;
        }
        
        .badge-warning {
            color: #212529;
            background-color: #ffc107;
        }
        
        .badge-info {
            color: #fff;
            background-color: #17a2b8;
        }
    </style>
    
    <!-- Include Google Charts library -->
    <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
    <script type="text/javascript">
        google.charts.load('current', {'packages':['corechart', 'bar', 'gauge']});
        
        // Function to apply filters
        function applyFilters() {
            var program = document.getElementById('program-filter').value;
            var division = document.getElementById('division-filter').value;
            var orgType = document.getElementById('orgtype-filter').value;
            var view = window.location.search.match(/[?&]view=([^&]+)/) 
                ? window.location.search.match(/[?&]view=([^&]+)/)[1] 
                : 'overview';
                
            var url = window.location.pathname + '?view=' + view;
            
            if (program) url += '&program=' + program;
            if (division) url += '&division=' + division;
            if (orgType) url += '&orgtype=' + orgType;
            
            window.location.href = url;
        }
        
        // Function to reset filters
        function resetFilters() {
            var view = window.location.search.match(/[?&]view=([^&]+)/) 
                ? window.location.search.match(/[?&]view=([^&]+)/)[1] 
                : 'overview';
                
            window.location.href = window.location.pathname + '?view=' + view;
        }
    </script>
    <script type="text/javascript">
        // Function to apply filters
        function applyFilters() {
            var program = document.getElementById('program-filter').value;
            var division = document.getElementById('division-filter').value;
            var orgType = document.getElementById('orgtype-filter').value;
            var view = window.location.search.match(/[?&]view=([^&]+)/) 
                ? window.location.search.match(/[?&]view=([^&]+)/)[1] 
                : 'overview';
                
            var url = window.location.pathname + '?view=' + view;
            
            if (program) url += '&program=' + program;
            if (division) url += '&division=' + division;
            if (orgType) url += '&orgtype=' + orgType;
            
            // Add program_sort parameter if on programs view
            if (view === 'programs') {
                var programSortRadios = document.getElementsByName('program_sort');
                for (var i = 0; i < programSortRadios.length; i++) {
                    if (programSortRadios[i].checked) {
                        url += '&program_sort=' + programSortRadios[i].value;
                        break;
                    }
                }
            }
            
            window.location.href = url;
        }
    </script>
    """

def print_navigation(current_view):
    """Print navigation tabs for different views"""
    tabs = [
        {"id": "overview", "name": "Overview"},
        {"id": "programs", "name": "Programs & Divisions"},
        {"id": "activity", "name": "Activity Metrics"},
        {"id": "meetings", "name": "Meetings"},
        {"id": "members", "name": "Member Changes"},
        {"id": "inactive", "name": "Inactive Involvements"}
    ]
    
    print '<div class="dashboard-container">'
    print """<h1>Involvement Activity Dashboard
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="85 75 230 130" style="width: 60px; height: 30px; margin-left: -4px; vertical-align: middle;">
        <!-- Text portion - TP -->
        <text x="100" y="120" font-family="Arial, sans-serif" font-weight="bold" font-size="60" fill="#333333">TP</text>
        
        <!-- Circular element -->
        <g transform="translate(190, 107)">
          <!-- Outer circle -->
          <circle cx="0" cy="0" r="13.5" fill="#0099FF"/>
          
          <!-- White middle circle -->
          <circle cx="0" cy="0" r="10.5" fill="white"/>
          
          <!-- Inner circle -->
          <circle cx="0" cy="0" r="7.5" fill="#0099FF"/>
          
          <!-- X crossing through the circles -->
          <path d="M-9 -9 L9 9 M-9 9 L9 -9" stroke="white" stroke-width="1.8" stroke-linecap="round"/>
        </g>
        
        <!-- Single "i" letter to the right -->
        <text x="206" y="105" font-family="Arial, sans-serif" font-weight="bold" font-size="14" fill="#0099FF">i</text>
      </svg></h1>"""
    
    print '<ul class="nav-tabs">'
    for tab in tabs:
        active_class = ' class="active"' if tab["id"] == current_view else ""
        print '<li><a href="?view={0}"{1}>{2}</a></li>'.format(tab["id"], active_class, tab["name"])
    print '</ul>'

def print_filter_controls(org_structure_data, involvement_type_data, program_filter, division_filter, org_type_filter):
    """Print filter controls for the dashboard"""
    # Get unique programs, divisions, and org types for filters
    programs = {}
    divisions = {}
    org_types = {}
    
    for item in org_structure_data:
        if not item.ProgId in programs:
            programs[item.ProgId] = item.Program
        
        if not item.DivId in divisions:
            divisions[item.DivId] = item.Division
        
        if hasattr(item, 'OrgType') and item.OrgType and not item.OrgType in org_types:
            org_types[item.OrgType] = item.OrgTypeDesc
    
    print '<div class="filter-controls">'
    print '<form>'
    print '<input type="hidden" name="view" value="{0}">'.format(
        model.Data.view if hasattr(model.Data, "view") else "overview")
        
    # Program Filter dropdown
    print '<select id="program-filter" name="program">'
    print '<option value="">All Programs</option>'
    for prog_id, prog_name in sorted(programs.items(), key=lambda x: x[1].lower()):
        selected = ' selected' if program_filter and str(prog_id) == str(program_filter) else ''
        print '<option value="{0}"{1}>{2}</option>'.format(prog_id, selected, prog_name)
    print '</select>'
    
    # Division filter
    print '<select id="division-filter" name="division">'
    print '<option value="">All Divisions</option>'
    for div_id, div_name in sorted(divisions.items(), key=lambda x: x[1].lower()):
        selected = ' selected' if division_filter and str(div_id) == str(division_filter) else ''
        print '<option value="{0}"{1}>{2}</option>'.format(div_id, selected, div_name)
    print '</select>'
    
    # Organization Type filter
    print '<select id="orgtype-filter" name="orgtype">'
    print '<option value="">All Types</option>'
    for type_code, type_desc in sorted(org_types.items(), key=lambda x: x[1].lower()):
        if type_code:  # Only include non-empty org types
            selected = ' selected' if org_type_filter and type_code == org_type_filter else ''
            print '<option value="{0}"{1}>{2}</option>'.format(type_code, selected, type_desc)
    print '</select>'
    
    # Program Sort option (only show when on programs view)
    current_view = model.Data.view if hasattr(model.Data, "view") else "overview"
    if current_view == "programs":
        program_sort = model.Data.program_sort if hasattr(model.Data, "program_sort") else "count"  # Changed default to "count"
        alpha_selected = ' checked' if program_sort == "alpha" else ''
        count_selected = ' checked' if program_sort == "count" or not program_sort else ''  # Make count selected by default
        
        print '<div style="margin-top: 10px;">'
        print '<label style="margin-right: 15px;">Sort: </label>'
        print '<label style="margin-right: 10px;"><input type="radio" name="program_sort" value="alpha"{0}> Alpha</label>'.format(alpha_selected)
        print '<label><input type="radio" name="program_sort" value="count"{0}> Inv. Count</label>'.format(count_selected)
        print '</div>'

    print '<button type="button" onclick="applyFilters()">Apply Filters</button>'
    print '<button type="button" onclick="resetFilters()">Reset Filters</button>'
    print '</form>'
    print '</div>'

def render_overview(org_data, involvement_type_data, inactive_orgs_data):
    """Render the overview dashboard with enhanced activity metrics"""
    # Calculate summary statistics
    total_programs = len(set([item.ProgId for item in org_data if hasattr(item, 'ProgId') and item.ProgId is not None]))
    total_divisions = len(set([item.DivId for item in org_data if hasattr(item, 'DivId') and item.DivId is not None]))
    total_orgs = len(org_data)
    
    # Safely sum members with None check
    total_members = sum(item.Members or 0 for item in org_data 
                        if hasattr(item, 'Members') and item.Members is not None)
            
    # Count by activity level
    high_activity = sum(1 for item in org_data 
                       if hasattr(item, 'ActivityLevel') and item.ActivityLevel == "high")
    
    moderate_activity = sum(1 for item in org_data 
                          if hasattr(item, 'ActivityLevel') and item.ActivityLevel == "moderate")
    
    low_activity = sum(1 for item in org_data 
                      if hasattr(item, 'ActivityLevel') and item.ActivityLevel == "low")
    
    inactive_orgs = sum(1 for item in org_data 
                       if hasattr(item, 'ActivityLevel') and item.ActivityLevel == "inactive")
    
    without_recent_activity = len(inactive_orgs_data)
    
    # Print summary statistics
    print '<div class="stat-grid">'
    print_stat_box("Total Programs", total_programs, COLOR_INFO)
    print_stat_box("Total Divisions", total_divisions, COLOR_INFO)
    print_stat_box("Total Involvements", total_orgs, COLOR_INFO)
    print_stat_box("Total Members", total_members, COLOR_INFO)
    print_stat_box("High Activity", high_activity, COLOR_ACTIVE)
    print_stat_box("Moderate Activity", moderate_activity, COLOR_WARNING)
    print_stat_box("Low Activity", low_activity, COLOR_INACTIVE)
    print_stat_box("Inactive Involvements", inactive_orgs, COLOR_NEUTRAL)
    print '</div>'
    
    # Activity distribution chart
    print '<div class="card">'
    print '<div class="card-header">Activity Level Distribution</div>'
    
    # Create data for the chart
    activity_data = [
        ["Activity Level", "Count"],
        ["High", high_activity],
        ["Moderate", moderate_activity],
        ["Low", low_activity],
        ["Inactive", inactive_orgs]
    ]
    
    print '<div class="chart-container" id="activity-distribution-chart"></div>'
    
    # Print chart JavaScript
    print """
    <script type="text/javascript">
        google.charts.setOnLoadCallback(drawActivityChart);
        
        function drawActivityChart() {
            var data = google.visualization.arrayToDataTable(%s);
            
            var options = {
                title: 'Involvement Activity Distribution',
                pieHole: 0.4,
                colors: ['%s', '%s', '%s', '%s'],
                legend: { position: 'right' }
            };
            
            var chart = new google.visualization.PieChart(document.getElementById('activity-distribution-chart'));
            chart.draw(data, options);
        }
    </script>
    """ % (json.dumps(activity_data), COLOR_ACTIVE, COLOR_WARNING, COLOR_INACTIVE, COLOR_NEUTRAL)
    
    print '</div>'
    
    # Print organization type distribution
    print '<div class="card">'
    print '<div class="card-header">Involvement Type Distribution</div>'
    
    # Create data for the chart
    chart_data = []
    for item in involvement_type_data:
        if hasattr(item, 'OrgType') and item.OrgType:
            chart_data.append([
                item.OrgTypeDesc, 
                item.OrgCount, 
                item.ActiveCount, 
                item.InactiveCount, 
                item.TotalMembers
            ])
    
    if chart_data:
        print '<div class="chart-container" id="type-distribution-chart"></div>'
        
        # Fix the JavaScript code by using triple quotes and proper escaping
        print """
        <script type="text/javascript">
            google.charts.setOnLoadCallback(drawTypeDistributionChart);
            
            function drawTypeDistributionChart() {
                var data = new google.visualization.DataTable();
                data.addColumn('string', 'Involvement Type');
                data.addColumn('number', 'Count');
                data.addColumn('number', 'Active');
                data.addColumn('number', 'Inactive');
                data.addColumn('number', 'Members');
                
                data.addRows(%s);
                
                var view = new google.visualization.DataView(data);
                view.setColumns([0, 1]);
                
                var options = {
                    title: 'Organization Types',
                    bars: 'horizontal',
                    legend: { position: 'none' },
                    height: 300
                };
                
                var chart = new google.visualization.BarChart(document.getElementById('type-distribution-chart'));
                chart.draw(view, options);
            }
        </script>
        """ % json.dumps(chart_data)
    else:
        print '<p>No organization type data available.</p>'
    
    print '</div>'
        
    # Show orgs without recent activity
    print '<div class="card">'
    print '<div class="card-header">Involvements Without Recent Activity (Last {0} days)</div>'.format(INACTIVE_DAYS)
    
    if inactive_orgs_data:
        print '<table>'
        print '<thead>'
        print '<tr>'
        print '<th>Involvement</th>'
        print '<th>Program</th>'
        print '<th>Division</th>'
        print '<th>Members</th>'
        print '<th>Type</th>'
        print '<th>Last Activity</th>'
        print '</tr>'
        print '</thead>'
        print '<tbody>'
        
        # Show only the first 10
        max_orgs_to_show = 10
        for i, org in enumerate(inactive_orgs_data):
            if i >= max_orgs_to_show:
                break
                
            # Format dates as strings for display and comparison
            created_date_str = None
            if hasattr(org, 'CreatedDate') and org.CreatedDate:
                try:
                    created_date_str = format_date(org.CreatedDate)
                except:
                    pass
            
            # Last Meeting Date
            last_meeting_str = None
            if hasattr(org, 'LastMeetingDate') and org.LastMeetingDate:
                try:
                    if isinstance(org.LastMeetingDate, basestring) and 'NULL' in org.LastMeetingDate.upper():
                        last_meeting_str = None
                    else:
                        last_meeting_str = format_date(org.LastMeetingDate)
                except:
                    pass
            
            # Last Member Added Date
            last_member_added_str = None
            if hasattr(org, 'LastMemberAddedDate') and org.LastMemberAddedDate:
                try:
                    if isinstance(org.LastMemberAddedDate, basestring) and 'NULL' in org.LastMemberAddedDate.upper():
                        last_member_added_str = None
                    else:
                        last_member_added_str = format_date(org.LastMemberAddedDate)
                except:
                    pass
            
            # Last Member Inactive Date
            last_inactive_str = None
            if hasattr(org, 'LastMemberInactiveDate') and org.LastMemberInactiveDate:
                try:
                    if isinstance(org.LastMemberInactiveDate, basestring) and 'NULL' in org.LastMemberInactiveDate.upper():
                        last_inactive_str = None
                    else:
                        last_inactive_str = format_date(org.LastMemberInactiveDate)
                except:
                    pass
            
            # Use a dictionary to store dates with their description
            activity_dates = {}
            
            if created_date_str:
                activity_dates[created_date_str] = ("Created", created_date_str)
            
            if last_meeting_str:
                activity_dates[last_meeting_str] = ("Last Meeting", last_meeting_str)
            
            if last_member_added_str:
                activity_dates[last_member_added_str] = ("Last Member Added", last_member_added_str)
            
            if last_inactive_str:
                activity_dates[last_inactive_str] = ("Last Member Inactive", last_inactive_str)
            
            # Find the most recent date by sorting the date strings
            if activity_dates:
                most_recent_date = sorted(activity_dates.keys(), reverse=True)[0]
                activity_type, date_str = activity_dates[most_recent_date]
                last_activity = "{0}: {1}".format(activity_type, date_str)
            else:
                last_activity = "No activity recorded"
            
            # Check if it's a new organization
            is_new = False
            if created_date_str:
                today = datetime.datetime.now()
                cutoff_date = today - datetime.timedelta(days=NEW_ORG_DAYS)
                cutoff_date_str = format_date(cutoff_date)
                is_new = created_date_str >= cutoff_date_str
            
            print '<tr>'
            print '<td><a href="/Org/{0}" target="_blank">{1}</a>{2}</td>'.format(
                org.OrganizationId, 
                org.OrganizationName,
                ' <span class="badge badge-info">New</span>' if is_new else '')
            print '<td>{0}</td>'.format(org.Program)
            print '<td>{0}</td>'.format(org.Division)
            print '<td>{0}</td>'.format(org.MemberCount)
            print '<td>{0}</td>'.format(org.OrgType if hasattr(org, 'OrgType') else '-')
            print '<td>{0}</td>'.format(last_activity)
            print '</tr>'
        
        print '</tbody>'
        print '</table>'
        
        if len(inactive_orgs_data) > 10:
            print '<p><em>Showing 10 of {0} involvements. View the "Inactive Involvements" tab for the complete list.</em></p>'.format(len(inactive_orgs_data))
    else:
        print '<p>No involvements without recent activity.</p>'
    
    print '</div>'

def render_activity_view(org_data, activity_detail_data):
    """Render the activity metrics view with combined activity indicators"""
    
    # Create a lookup of detailed activity data by org ID
    activity_lookup = {}
    for item in activity_detail_data:
        if hasattr(item, 'OrganizationId'):
            activity_lookup[item.OrganizationId] = item
    
    # Count organizations by activity score ranges
    activity_ranges = {
        "Highly Active (Score 7+)": 0,
        "Moderately Active (Score 3-6)": 0,
        "Low Activity (Score 1-2)": 0,
        "Inactive (Score 0)": 0
    }
    
    for org in org_data:
        if not hasattr(org, 'ActivityScore'):
            continue
            
        if org.ActivityScore >= 7:
            activity_ranges["Highly Active (Score 7+)"] += 1
        elif org.ActivityScore >= 3:
            activity_ranges["Moderately Active (Score 3-6)"] += 1
        elif org.ActivityScore > 0:
            activity_ranges["Low Activity (Score 1-2)"] += 1
        else:
            activity_ranges["Inactive (Score 0)"] += 1
    
    # Print summary statistics
    print '<div class="stat-grid">'
    print_stat_box("Highly Active", activity_ranges["Highly Active (Score 7+)"], COLOR_ACTIVE)
    print_stat_box("Moderately Active", activity_ranges["Moderately Active (Score 3-6)"], COLOR_WARNING)
    print_stat_box("Low Activity", activity_ranges["Low Activity (Score 1-2)"], COLOR_INACTIVE)
    print_stat_box("Inactive", activity_ranges["Inactive (Score 0)"], COLOR_NEUTRAL)
    print '</div>'
    
    # Activity score distribution chart
    print '<div class="card">'
    print '<div class="card-header">Activity Score Distribution</div>'
    
    # Create data for the chart
    activity_data = [
        ["Activity Level", "Count"],
        ["Highly Active", activity_ranges["Highly Active (Score 7+)"]],
        ["Moderately Active", activity_ranges["Moderately Active (Score 3-6)"]],
        ["Low Activity", activity_ranges["Low Activity (Score 1-2)"]],
        ["Inactive", activity_ranges["Inactive (Score 0)"]]
    ]
    
    print '<div class="chart-container" id="activity-score-chart"></div>'
    
    # Print chart JavaScript
    print """
    <script type="text/javascript">
        google.charts.setOnLoadCallback(drawActivityScoreChart);
        
        function drawActivityScoreChart() {
            var data = google.visualization.arrayToDataTable(%s);
            
            var options = {
                title: 'Involvement Activity Score Distribution',
                is3D: true,
                colors: ['%s', '%s', '%s', '%s'],
                legend: { position: 'right' }
            };
            
            var chart = new google.visualization.PieChart(document.getElementById('activity-score-chart'));
            chart.draw(data, options);
        }
    </script>
    """ % (json.dumps(activity_data), COLOR_ACTIVE, COLOR_WARNING, COLOR_INACTIVE, COLOR_NEUTRAL)
    
    print '</div>'
    
    # Activity Metrics Table
    print '<div class="card">'
    print '<div class="card-header">Involvement Activity Metrics</div>'
    
    print '<div style="margin-bottom: 15px;">'
    print '<p>The activity score is calculated based on the following factors:</p>'
    print '<ul>'
    print '<li><strong>Meetings:</strong> Recent meetings within the past {0} days (up to 10 points)</li>'.format(RECENT_DAYS)
    print '<li><strong>Member Additions:</strong> New members within the past {0} days (up to 10 points, weighted higher for member-management organizations)</li>'.format(MEMBER_CHANGE_DAYS)
    print '<li><strong>Member Status Changes:</strong> Status changes within the past {0} days (up to 6 points, weighted higher for member-management organizations)</li>'.format(MEMBER_CHANGE_DAYS)
    print '<li><strong>New Involvement:</strong> Created within the past {0} days (5 points bonus)</li>'.format(NEW_ORG_DAYS)
    print '<li><strong>Member-Management Bonus:</strong> Involvements that primarily manage people without regular meetings receive additional points</li>'
    print '</ul>'
    print '</div>'
    
    # Sort organizations by activity score (descending)
    sorted_orgs = sorted(org_data, key=lambda x: getattr(x, 'ActivityScore', 0), reverse=True)
    
    print '<table>'
    print '<thead>'
    print '<tr>'
    print '<th>Involvement</th>'
    print '<th>Members</th>'
    print '<th>Recent Meetings</th>'
    print '<th>Member Changes</th>'
    print '<th>Created Date</th>'
    print '<th>Activity Score</th>'
    print '<th>Activity Level</th>'
    print '</tr>'
    print '</thead>'
    print '<tbody>'
    
    for org in sorted_orgs:
        if not hasattr(org, 'OrgId') or not hasattr(org, 'ActivityScore'):
            continue
            
        activity_score = org.ActivityScore
        activity_level = org.ActivityLevel
        
        # Determine activity level class and style
        if activity_level == "high":
            level_class = "activity-level-high"
            bar_color = COLOR_ACTIVE
            bar_width = "100%"
        elif activity_level == "moderate":
            level_class = "activity-level-moderate"
            bar_color = COLOR_WARNING
            bar_width = "66%"
        elif activity_level == "low":
            level_class = "activity-level-low"
            bar_color = COLOR_INACTIVE
            bar_width = "33%"
        else:
            level_class = "activity-level-inactive"
            bar_color = COLOR_NEUTRAL
            bar_width = "0%"
        
        # Get recent meetings count
        recent_meetings = org.RecentMeetings if hasattr(org, 'RecentMeetings') else 0
        
        # Get member change counts
        new_members = org.NewMembersCount if hasattr(org, 'NewMembersCount') else 0
        inactive_changes = org.RecentInactiveCount if hasattr(org, 'RecentInactiveCount') else 0
        member_changes = new_members + inactive_changes
        
        # Get created date
        created_date = format_date(org.CreatedDate) if hasattr(org, 'CreatedDate') and org.CreatedDate else '-'
        
        # Check if it's a new organization
        is_new_org = False
        if hasattr(org, 'CreatedDate') and org.CreatedDate:
            try:
                if isinstance(org.CreatedDate, basestring):
                    created_date_obj = datetime.datetime.strptime(org.CreatedDate, '%Y-%m-%d %H:%M:%S.%f')
                else:
                    created_date_obj = org.CreatedDate
                
                days_since_creation = (datetime.datetime.now() - created_date_obj).days
                is_new_org = days_since_creation <= NEW_ORG_DAYS
            except:
                pass
        
        print '<tr>'
        print '<td><a href="/Org/{0}" target="_blank">{1}</a>{2}</td>'.format(
            org.OrgId, 
            org.Organization,
            ' <span class="badge badge-info">New</span>' if is_new_org else '')
        print '<td>{0}</td>'.format(org.Members if hasattr(org, 'Members') else 0)
        print '<td>{0}</td>'.format(recent_meetings)
        print '<td>{0}</td>'.format(member_changes)
        print '<td>{0}</td>'.format(created_date)
        print '<td>{0}</td>'.format(activity_score)
        print '<td class="{0}">'.format(level_class)
        print '<div class="activity-gauge">'
        print '<div class="activity-gauge-fill" style="width: {0}; background-color: {1};"></div>'.format(bar_width, bar_color)
        print '</div>'
        print ' {0}'.format(activity_level.capitalize())
        print '</td>'
        print '</tr>'
    
    print '</tbody>'
    print '</table>'
    
    print '</div>'
    
    print '</div>'  # Close dashboard-container

def render_members_view(member_changes_data, filtered_data):
    """Render the member changes view"""
    
    # Count total recent changes
    #total_enrollments = sum(1 for item in member_changes_data if hasattr(item, 'InactiveDate') and not item.InactiveDate)
    #total_inactivations = sum(1 for item in member_changes_data if hasattr(item, 'InactiveDate') and item.InactiveDate)
    total_enrollments = sum(1 for item in member_changes_data 
                            if hasattr(item, 'EnrollmentDate') 
                            and item.MemberStatus == 'Active')
    total_inactivations = sum(1 for item in member_changes_data 
                               if hasattr(item, 'InactiveDate') 
                               and item.MemberStatus == 'Inactive')
    
    # Print summary statistics
    print '<div class="stat-grid">'
    print_stat_box("Recent Enrollments", total_enrollments, COLOR_ACTIVE)
    print_stat_box("Recent Inactivations", total_inactivations, COLOR_WARNING)
    print_stat_box("Total Changes", total_enrollments + total_inactivations, COLOR_INFO)
    print '</div>'
    
    # Recent member changes table
    print '<div class="card">'
    print '<div class="card-header">Recent Member Changes (Last {0} days)</div>'.format(MEMBER_CHANGE_DAYS)
    
    if member_changes_data:
        # Add a scrollable div around the table
        print '''
        <div style="max-height: 700px; overflow-y: auto;">
        <table>
        '''
        print '<thead>'
        print '<tr>'
        print '<th>Date</th>'
        print '<th>Member</th>'
        print '<th>Involvement</th>'
        print '<th>Change Type</th>'
        print '<th>Member Type</th>'
        print '</tr>'
        print '</thead>'
        print '<tbody>'
        
        for change in member_changes_data:
            # Determine if this is an enrollment or inactivation
            if hasattr(change, 'InactiveDate') and change.InactiveDate:
                change_type = "Inactive"
                change_date = format_date(change.InactiveDate)
                badge_class = "badge-warning"
            else:
                change_type = "Enrolled"
                change_date = format_date(change.EnrollmentDate)
                badge_class = "badge-success"
            
            print '<tr>'
            print '<td>{0}</td>'.format(change_date)
            print '<td>{0}</td>'.format(change.PersonName)
            print '<td><a href="/Org/{0}" target="_blank">{1}</a></td>'.format(change.OrganizationId, change.OrganizationName)
            print '<td><span class="badge {0}">{1}</span></td>'.format(badge_class, change_type)
            print '<td>{0}</td>'.format(change.MemberType if hasattr(change, 'MemberType') else '-')
            print '</tr>'
        
        print '</tbody>'
        print '</table>'
        print '</div>'  # Close scrollable div
        
        print '<p><em>Showing {0} recent member changes.</em></p>'.format(len(member_changes_data))
    else:
        print '<p>No recent member changes found.</p>'
    
    print '</div>'

    
    # Organizations with most member activity
    print '<div class="card">'
    print '<div class="card-header">Involvements with Most Member Activity</div>'
    
    # Group and count changes by organization
    org_activity = {}
    for change in member_changes_data:
        if not hasattr(change, 'OrganizationId'):
            continue
            
        org_id = change.OrganizationId
        org_name = change.OrganizationName if hasattr(change, 'OrganizationName') else "Unknown Org"
        
        if not org_id in org_activity:
            org_activity[org_id] = {
                "name": org_name,
                "enrollments": 0,
                "inactivations": 0
            }
        
        if hasattr(change, 'InactiveDate') and change.InactiveDate:
            org_activity[org_id]["inactivations"] += 1
        else:
            org_activity[org_id]["enrollments"] += 1
    
    # Convert to list and sort by total activity
    org_activity_list = []
    for org_id, data in org_activity.items():
        total = data["enrollments"] + data["inactivations"]
        org_activity_list.append({
            "id": org_id,
            "name": data["name"],
            "enrollments": data["enrollments"],
            "inactivations": data["inactivations"],
            "total": total
        })
    
    org_activity_list = sorted(org_activity_list, key=lambda x: x["total"], reverse=True)
    
    # Create data for the chart
    if org_activity_list:
        chart_data = []
        for org in org_activity_list[:15]:  # Show top 15
            chart_data.append([
                org["name"],
                org["enrollments"],
                org["inactivations"]
            ])
        
        print '<div class="chart-container" id="member-activity-chart"></div>'
        
        # Print chart JavaScript
        print """
        <script type="text/javascript">
            google.charts.setOnLoadCallback(drawMemberActivityChart);
            
            function drawMemberActivityChart() {
                var data = new google.visualization.DataTable();
                data.addColumn('string', 'Organization');
                data.addColumn('number', 'Enrollments');
                data.addColumn('number', 'Inactivations');
                
                data.addRows(%s);
                
                var options = {
                    title: 'Involvements with Most Member Activity',
                    isStacked: true,
                    bars: 'horizontal',
                    colors: ['%s', '%s'],
                    legend: { position: 'top' }
                };
                
                var chart = new google.visualization.BarChart(document.getElementById('member-activity-chart'));
                chart.draw(data, options);
            }
        </script>
        """ % (json.dumps(chart_data), COLOR_ACTIVE, COLOR_WARNING)
        
        # Show table of involvements with most activity
        print '<table>'
        print '<thead>'
        print '<tr>'
        print '<th>Involvements</th>'
        print '<th>Enrollments</th>'
        print '<th>Inactivations</th>'
        print '<th>Total Activity</th>'
        print '</tr>'
        print '</thead>'
        print '<tbody>'
        
        for org in org_activity_list[:20]:  # Show top 20
            print '<tr>'
            print '<td><a href="/Org/{0}" target="_blank">{1}</a></td>'.format(org["id"], org["name"])
            print '<td>{0}</td>'.format(org["enrollments"])
            print '<td>{0}</td>'.format(org["inactivations"])
            print '<td>{0}</td>'.format(org["total"])
            print '</tr>'
        
        print '</tbody>'
        print '</table>'
    else:
        print '<p>No involvement member activity data available.</p>'
    
    print '</div>'
    
    print '</div>'  # Close dashboard-container
    
def render_inactive_view(inactive_orgs_data, filtered_data):
    """Render the inactive organizations view with more detailed analysis"""
    try:
        # Calculate statistics
        if inactive_orgs_data:
            total_inactive = len(inactive_orgs_data)
            
            # Use or 0 to handle None values
            inactive_with_members = sum(1 for org in inactive_orgs_data if (org.MemberCount or 0) > 0)
            
            # Use or 0 to prevent None + long error
            members_affected = sum(org.MemberCount or 0 for org in inactive_orgs_data)
            
            # Group by creation date to see how many are actually new
            recent_creation_count = 0
            for org in inactive_orgs_data:
                if hasattr(org, 'CreatedDate') and org.CreatedDate:
                    try:
                        created_date_str = format_date(org.CreatedDate)
                        
                        # Check if it's within NEW_ORG_DAYS from today
                        # We'll need to use the string comparison approach here to avoid type issues
                        today = datetime.datetime.now()
                        cutoff_date = today - datetime.timedelta(days=NEW_ORG_DAYS)
                        cutoff_date_str = format_date(cutoff_date)
                        
                        if created_date_str >= cutoff_date_str:
                            recent_creation_count += 1
                    except:
                        pass
            
            # Print summary
            print '<div class="stat-grid">'
            print_stat_box("Total Inactive", total_inactive, COLOR_INACTIVE)
            print_stat_box("With Members", inactive_with_members, COLOR_WARNING)
            print_stat_box("Members Affected", members_affected, COLOR_WARNING)
            #print_stat_box("Recently Created", recent_creation_count, COLOR_INFO)
            print '</div>'
            
            # Additional inactive reasons analysis
            print '<div class="card">'
            print '<div class="card-header">Inactivity Analysis</div>'
            
            print '<p>This view shows involvements that have had no activity (meetings, member changes) for at least {0} days.</p>'.format(INACTIVE_DAYS)
            
            # Group by inactive reasons
            inactive_reasons = {
                "No Meetings Ever": 0,
                "No Recent Meetings": 0,
                "No Recent Member Changes": 0
            }
            
            for org in inactive_orgs_data:
                # For each org, we need to determine the reason for inactivity
                # We'll check all possible reasons and choose the most appropriate one
                has_meeting_history = hasattr(org, 'LastMeetingDate') and org.LastMeetingDate
                has_recent_member_changes = False
                
                # Check for recently added members
                if hasattr(org, 'LastMemberAddedDate') and org.LastMemberAddedDate:
                    try:
                        last_add_date = convert_to_datetime(org.LastMemberAddedDate)
                        if last_add_date:
                            days_since_add = (datetime.datetime.now() - last_add_date).days
                            if days_since_add <= MEMBER_CHANGE_DAYS:
                                has_recent_member_changes = True
                    except:
                        pass
                        
                # Check for recently inactivated members
                if hasattr(org, 'LastMemberInactiveDate') and org.LastMemberInactiveDate:
                    try:
                        last_inactive_date = convert_to_datetime(org.LastMemberInactiveDate)
                        if last_inactive_date:
                            days_since_inactive = (datetime.datetime.now() - last_inactive_date).days
                            if days_since_inactive <= MEMBER_CHANGE_DAYS:
                                has_recent_member_changes = True
                    except:
                        pass
                
                # Now determine the primary reason for inactivity
                if not has_meeting_history:
                    inactive_reasons["No Meetings Ever"] += 1
                elif not has_recent_member_changes:
                    inactive_reasons["No Recent Member Changes"] += 1
                else:
                    inactive_reasons["No Recent Meetings"] += 1
            
            # Create chart data
            chart_data = [
                ["Reason", "Count"],
                ["No Meetings Ever", inactive_reasons["No Meetings Ever"]],
                ["No Recent Meetings", inactive_reasons["No Recent Meetings"]],
                ["No Recent Member Changes", inactive_reasons["No Recent Member Changes"]]
            ]
            
            print '<div class="chart-container" id="inactive-reasons-chart"></div>'
            
            # Print chart JavaScript
            print """
            <script type="text/javascript">
                google.charts.setOnLoadCallback(drawInactiveReasonsChart);
                
                function drawInactiveReasonsChart() {
                    var data = google.visualization.arrayToDataTable(%s);
                    
                    var options = {
                        title: 'Reasons for Inactivity',
                        pieHole: 0.4,
                        colors: ['%s', '%s', '%s'],
                        legend: { position: 'right' }
                    };
                    
                    var chart = new google.visualization.PieChart(document.getElementById('inactive-reasons-chart'));
                    chart.draw(data, options);
                }
            </script>
            """ % (json.dumps(chart_data), COLOR_INACTIVE, COLOR_WARNING, COLOR_INFO)
            
            print '</div>'
            
            # Show table of inactive orgs
            print """<div class="card">
                        <div class="card-header">Inactive Involvements</div>
                            <table>
                                <thead>
                                    <tr>
                                        <th>Involvement</th>
                                        <th>Program</th>
                                        <th>Division</th>
                                        <th>Members</th>
                                        <th>Type</th>
                                        <th>Created Date</th>
                                        <th>Last Activity</th>
                                    </tr>
                                </thead>
                                <tbody>"""
            
            # Sort by member count (descending)
            sorted_orgs = sorted(inactive_orgs_data, key=lambda x: x.MemberCount or 0, reverse=True)
            
            for org in sorted_orgs:
                # Format dates as strings for display and comparison
                
                # Created Date
                created_date_str = None
                if hasattr(org, 'CreatedDate') and org.CreatedDate:
                    try:
                        created_date_str = format_date(org.CreatedDate)
                    except:
                        pass
                
                # Last Meeting Date
                last_meeting_str = None
                if hasattr(org, 'LastMeetingDate') and org.LastMeetingDate:
                    try:
                        if isinstance(org.LastMeetingDate, basestring) and 'NULL' in org.LastMeetingDate.upper():
                            last_meeting_str = None
                        else:
                            last_meeting_str = format_date(org.LastMeetingDate)
                    except:
                        pass
                
                # Last Member Added Date
                last_member_added_str = None
                if hasattr(org, 'LastMemberAddedDate') and org.LastMemberAddedDate:
                    try:
                        if isinstance(org.LastMemberAddedDate, basestring) and 'NULL' in org.LastMemberAddedDate.upper():
                            last_member_added_str = None
                        else:
                            last_member_added_str = format_date(org.LastMemberAddedDate)
                    except:
                        pass
                
                # Last Member Inactive Date
                last_inactive_str = None
                if hasattr(org, 'LastMemberInactiveDate') and org.LastMemberInactiveDate:
                    try:
                        if isinstance(org.LastMemberInactiveDate, basestring) and 'NULL' in org.LastMemberInactiveDate.upper():
                            last_inactive_str = None
                        else:
                            last_inactive_str = format_date(org.LastMemberInactiveDate)
                    except:
                        pass
                
                # Use a dictionary to store dates with their description
                # We'll use the formatted date string as both the key and part of the value
                activity_dates = {}
                
                if created_date_str:
                    activity_dates[created_date_str] = ("Created", created_date_str)
                
                if last_meeting_str:
                    activity_dates[last_meeting_str] = ("Last Meeting", last_meeting_str)
                
                if last_member_added_str:
                    activity_dates[last_member_added_str] = ("Last Member Added", last_member_added_str)
                
                if last_inactive_str:
                    activity_dates[last_inactive_str] = ("Last Member Inactive", last_inactive_str)
                
                # Find the most recent date by sorting the date strings
                # This works because MM/DD/YYYY format will sort correctly as strings
                if activity_dates:
                    most_recent_date = sorted(activity_dates.keys(), reverse=True)[0]
                    activity_type, date_str = activity_dates[most_recent_date]
                    last_activity = "{0}: {1}".format(activity_type, date_str)
                else:
                    last_activity = "No activity recorded"
                
                # Check if it's a new organization - use the same approach as above to avoid datetime math
                is_new = False
                if created_date_str:
                    today = datetime.datetime.now()
                    cutoff_date = today - datetime.timedelta(days=NEW_ORG_DAYS)
                    cutoff_date_str = format_date(cutoff_date)
                    is_new = created_date_str >= cutoff_date_str
                
                print '<tr>'
                print '<td><a href="/Org/{0}" target="_blank">{1}</a>{2}</td>'.format(
                    org.OrganizationId, 
                    org.OrganizationName, 
                    ' <span class="badge badge-info">New</span>' if is_new else '')
                print '<td>{0}</td>'.format(org.Program)
                print '<td>{0}</td>'.format(org.Division)
                print '<td>{0}</td>'.format(org.MemberCount)
                print '<td>{0}</td>'.format(org.OrgType if hasattr(org, 'OrgType') else '-')
                print '<td>{0}</td>'.format(created_date_str if created_date_str else '-')
                print '<td>{0}</td>'.format(last_activity)
                print '</tr>'
            
            print '</tbody>'
            print '</table>'
            
            print '</div>'
        else:
            print '<p>No inactive involvements found.</p>'
    except Exception as e:
        print '<div class="card">'
        print '<div class="card-header">Error</div>'
        print '<p>An error occurred while rendering the inactive involvements view: {0}</p>'.format(str(e))
        print '<pre>'
        traceback.print_exc()
        print '</pre>'
        print '</div>'
    
    print '</div>'  # Close dashboard-container



def render_programs_view(org_data):
    """Render the programs & divisions view with enhanced activity metrics"""
    # Organize data by program and division
    program_data = {}
    
    for item in org_data:
        if not hasattr(item, 'ProgId') or item.ProgId is None:
            continue  # Skip items without valid ProgId
            
        if not item.ProgId in program_data:
            program_data[item.ProgId] = {
                "name": item.Program if hasattr(item, 'Program') else "Unknown Program",
                "divisions": {},
                "total_orgs": 0,
                "active_orgs": 0,
                "inactive_orgs": 0,
                "high_activity": 0,
                "moderate_activity": 0,
                "low_activity": 0,
                "total_members": 0
            }
        
        prog = program_data[item.ProgId]
        prog["total_orgs"] += 1
        
        # Add activity-based counts
        if hasattr(item, 'ActivityLevel'):
            if item.ActivityLevel == "high":
                prog["high_activity"] += 1
                prog["active_orgs"] += 1
            elif item.ActivityLevel == "moderate":
                prog["moderate_activity"] += 1
                prog["active_orgs"] += 1
            elif item.ActivityLevel == "low":
                prog["low_activity"] += 1
                prog["active_orgs"] += 1
            else:
                prog["inactive_orgs"] += 1
        elif hasattr(item, 'OrgStatusId') and item.OrgStatusId == 30:
            prog["active_orgs"] += 1
        else:
            prog["inactive_orgs"] += 1
        
        # Use or 0 to handle None values for member counts
        prog["total_members"] += item.Members or 0 if hasattr(item, 'Members') else 0
        
        if not hasattr(item, 'DivId') or item.DivId is None:
            continue  # Skip items without valid DivId
            
        if not item.DivId in prog["divisions"]:
            prog["divisions"][item.DivId] = {
                "name": item.Division if hasattr(item, 'Division') else "Unknown Division",
                "orgs": [],
                "total_orgs": 0,
                "active_orgs": 0,
                "inactive_orgs": 0,
                "high_activity": 0,
                "moderate_activity": 0,
                "low_activity": 0,
                "total_members": 0
            }
        
        div = prog["divisions"][item.DivId]
        div["orgs"].append(item)
        div["total_orgs"] += 1
        
        # Add activity-based counts at division level
        if hasattr(item, 'ActivityLevel'):
            if item.ActivityLevel == "high":
                div["high_activity"] += 1
                div["active_orgs"] += 1
            elif item.ActivityLevel == "moderate":
                div["moderate_activity"] += 1
                div["active_orgs"] += 1
            elif item.ActivityLevel == "low":
                div["low_activity"] += 1
                div["active_orgs"] += 1
            else:
                div["inactive_orgs"] += 1
        elif hasattr(item, 'OrgStatusId') and item.OrgStatusId == 30:
            div["active_orgs"] += 1
        else:
            div["inactive_orgs"] += 1
        
        if hasattr(item, 'Members') and item.Members is not None:
            div["total_members"] += item.Members
    
    # Get the sort parameter
    program_sort = model.Data.program_sort if hasattr(model.Data, "program_sort") else "count"  # Changed default to "count"


    # Print program and division summary
    print '<div class="card">'
    print '<div class="card-header">Programs and Divisions Summary</div>'
    
    if program_data:
        # Sort programs by total organizations
        #sorted_programs = sorted(program_data.items(), key=lambda x: x[1]["total_orgs"], reverse=True)
        # Sort programs alphabetically by name
        #sorted_programs = sorted(program_data.items(), key=lambda x: x[1]["name"].lower())
        # Sort programs based on selected sort option
        if program_sort == "alpha":
            sorted_programs = sorted(program_data.items(), key=lambda x: x[1]["name"].lower())
        else:
            # Default to count
            sorted_programs = sorted(program_data.items(), key=lambda x: x[1]["total_orgs"], reverse=True)
                
        
        for prog_id, prog in sorted_programs:
            active_percent = (float(prog["active_orgs"]) / prog["total_orgs"] * 100) if prog["total_orgs"] > 0 else 0
            
            print '<div style="margin-bottom: 20px;">'
            print '<h3>{0}</h3>'.format(prog["name"])
            print '<div class="stat-grid">'
            print_stat_box("Involvements", prog["total_orgs"], COLOR_INFO)
            print_stat_box("Members", prog["total_members"], COLOR_INFO)
            print_stat_box("High Activity", prog["high_activity"], COLOR_ACTIVE)
            print_stat_box("Moderate", prog["moderate_activity"], COLOR_WARNING)
            print_stat_box("Low", prog["low_activity"], COLOR_INACTIVE)
            print_stat_box("Inactive", prog["inactive_orgs"], COLOR_NEUTRAL)
            print '</div>'
            
            print '<div class="progress">'
            print '<div class="progress-bar activity-high" style="width: {0}%">{1}</div>'.format(
                (float(prog["high_activity"]) / prog["total_orgs"] * 100) if prog["total_orgs"] > 0 else 0,
                prog["high_activity"]
            )
            print '<div class="progress-bar activity-moderate" style="width: {0}%">{1}</div>'.format(
                (float(prog["moderate_activity"]) / prog["total_orgs"] * 100) if prog["total_orgs"] > 0 else 0,
                prog["moderate_activity"]
            )
            print '<div class="progress-bar activity-low" style="width: {0}%">{1}</div>'.format(
                (float(prog["low_activity"]) / prog["total_orgs"] * 100) if prog["total_orgs"] > 0 else 0,
                prog["low_activity"]
            )
            print '</div>'
            
            # Sort divisions by total organizations
            sorted_divisions = sorted(prog["divisions"].items(), key=lambda x: x[1]["total_orgs"], reverse=True)
            
            print '<table>'
            print '<thead>'
            print '<tr>'
            print '<th>Division</th>'
            print '<th>Involvements</th>'
            print '<th>Members</th>'
            print '<th>High</th>'
            print '<th>Moderate</th>'
            print '<th>Low</th>'
            print '<th>Inactive</th>'
            print '<th>Activity Distribution</th>'
            print '</tr>'
            print '</thead>'
            print '<tbody>'
            
            for div_id, div in sorted_divisions:
                print '<tr>'
                print '<td>{0}</td>'.format(div["name"])
                print '<td>{0}</td>'.format(div["total_orgs"])
                print '<td>{0}</td>'.format(div["total_members"])
                print '<td>{0}</td>'.format(div["high_activity"])
                print '<td>{0}</td>'.format(div["moderate_activity"])
                print '<td>{0}</td>'.format(div["low_activity"])
                print '<td>{0}</td>'.format(div["inactive_orgs"])
                print '<td><div class="progress" style="margin-bottom: 0;">'
                
                # Calculate percentages for the progress bar segments
                high_pct = (float(div["high_activity"]) / div["total_orgs"] * 100) if div["total_orgs"] > 0 else 0
                mod_pct = (float(div["moderate_activity"]) / div["total_orgs"] * 100) if div["total_orgs"] > 0 else 0
                low_pct = (float(div["low_activity"]) / div["total_orgs"] * 100) if div["total_orgs"] > 0 else 0
                
                print '<div class="progress-bar activity-high" style="width: {0}%"></div>'.format(high_pct)
                print '<div class="progress-bar activity-moderate" style="width: {0}%"></div>'.format(mod_pct)
                print '<div class="progress-bar activity-low" style="width: {0}%"></div>'.format(low_pct)
                print '</div></td>'
                print '</tr>'
            
            print '</tbody>'
            print '</table>'
            
            # Now list organizations in a collapsible section
            print '<div style="margin-top: 10px;">'
            print '<button onclick="toggleOrgs_{0}()" style="background: none; border: none; cursor: pointer; text-decoration: underline; color: #007bff;">Show/Hide Involvements</button>'.format(prog_id)
            print '<div id="orgs_{0}" style="display: none; margin-top: 10px;">'.format(prog_id)
            print '<table>'
            print '<thead>'
            print '<tr>'
            print '<th>Involvement</th>'
            print '<th>Status</th>'
            print '<th>Members</th>'
            print '<th>Previous</th>'
            print '<th>Visitors</th>'
            print '<th>Activity Level</th>'
            print '<th>Last Activity</th>'
            print '<th>Type</th>'
            print '</tr>'
            print '</thead>'
            print '<tbody>'
            
            # Sort organizations by member count
            all_orgs = []
            for div_id, div in prog["divisions"].items():
                all_orgs.extend(div["orgs"])
            
            sorted_orgs = sorted(all_orgs, key=lambda x: x.Members or 0 if hasattr(x, 'Members') else 0, reverse=True)
            
            for org in sorted_orgs:
                # Get activity level
                activity_level = "Unknown"
                activity_class = "status-inactive"
                
                if hasattr(org, 'ActivityLevel'):
                    activity_level = org.ActivityLevel
                    if activity_level == "high":
                        activity_class = "status-active"
                    elif activity_level == "moderate":
                        activity_class = "status-warning"
                    elif activity_level == "low":
                        activity_class = "status-inactive"
                    else:
                        activity_class = "status-inactive"
                elif hasattr(org, 'OrgStatusId') and org.OrgStatusId == 30:
                    activity_level = "Active"
                    activity_class = "status-active"
                else:
                    activity_level = "Inactive"
                    activity_class = "status-inactive"
                
                # Determine last activity date
                last_activity = "Unknown"
                if hasattr(org, 'LastActivityDate') and org.LastActivityDate:
                    last_activity = format_date(org.LastActivityDate)
                elif hasattr(org, 'LastMeetingDate') and org.LastMeetingDate:
                    last_activity = format_date(org.LastMeetingDate)
                
                print '<tr>'
                print '<td><a href="/Org/{0}" target="_blank">{1}</a></td>'.format(org.OrgId, org.Organization)
                print '<td><span class="{0}">{1}</span></td>'.format(activity_class, activity_level.capitalize())
                print '<td>{0}</td>'.format(org.Members if hasattr(org, 'Members') else 0)
                print '<td>{0}</td>'.format(org.Previous if hasattr(org, 'Previous') else 0)
                print '<td>{0}</td>'.format(org.Visitors if hasattr(org, 'Visitors') else 0)
                
                # Activity level gauge
                print '<td>'
                if hasattr(org, 'ActivityScore'):
                    print '<div class="activity-gauge">'
                    
                    if activity_level == "high":
                        width = "100%"
                        color = COLOR_ACTIVE
                    elif activity_level == "moderate":
                        width = "66%"
                        color = COLOR_WARNING
                    elif activity_level == "low":
                        width = "33%"
                        color = COLOR_INACTIVE
                    else:
                        width = "0%"
                        color = COLOR_NEUTRAL
                    
                    print '<div class="activity-gauge-fill" style="width: {0}; background-color: {1};"></div>'.format(width, color)
                    print '</div>'
                    print ' {0} ({1})'.format(activity_level.capitalize(), org.ActivityScore)
                else:
                    print activity_level.capitalize()
                print '</td>'
                
                print '<td>{0}</td>'.format(last_activity)
                print '<td>{0}</td>'.format(org.OrgType if hasattr(org, 'OrgType') else '-')
                print '</tr>'
            
            print '</tbody>'
            print '</table>'
            print '</div>'
            print '</div>'
            
            # Add JavaScript to toggle the organizations
            print """
            <script>
                function toggleOrgs_{0}() {{
                    var orgsDiv = document.getElementById('orgs_{0}');
                    if (orgsDiv.style.display === 'none') {{
                        orgsDiv.style.display = 'block';
                    }} else {{
                        orgsDiv.style.display = 'none';
                    }}
                }}
            </script>
            """.format(prog_id)
            
            print '</div>'
    else:
        print '<p>No program data available.</p>'
    
    print '</div>'
    
    print '</div>'  # Close dashboard-container

def render_meetings_view(meetings_data, org_data):
    """Render the meetings view with enhanced activity analysis"""
    print("<!-- DEBUG: Total meetings passed: " + str(len(meetings_data)) + " -->")
    # Group organizations by type and count meetings
    org_types = {}
    
    for org in org_data:
        if hasattr(org, 'OrgType') and org.OrgType and hasattr(org, 'Meetings'):
            if not org.OrgType in org_types:
                org_types[org.OrgType] = {
                    "desc": org.OrgTypeDesc if hasattr(org, 'OrgTypeDesc') else org.OrgType,
                    "count": 0,
                    "meetings": 0,
                    "orgs": 0,
                    "active_orgs": 0
                }
            
            org_types[org.OrgType]["orgs"] += 1
            
            # Count active orgs (those with activity level high or moderate)
            if hasattr(org, 'ActivityLevel') and (org.ActivityLevel == "high" or org.ActivityLevel == "moderate"):
                org_types[org.OrgType]["active_orgs"] += 1
            
            # Use or 0 to handle None values for meetings
            org_types[org.OrgType]["meetings"] += org.Meetings or 0
    
    # Meetings by organization type
    print '<div class="card">'
    print '<div class="card-header">Meetings by Involvement Type</div>'
    
    if org_types:
        # Convert to sorted list for easier display
        org_type_list = []
        for type_code, data in org_types.items():
            # Calculate average meetings per org
            avg_meetings = float(data["meetings"]) / data["orgs"] if data["orgs"] > 0 else 0
            
            org_type_list.append({
                "code": type_code,
                "desc": data["desc"],
                "orgs": data["orgs"],
                "active_orgs": data["active_orgs"],
                "meetings": data["meetings"],
                "avg_meetings": avg_meetings
            })
        
        # Sort by number of meetings
        org_type_list = sorted(org_type_list, key=lambda x: x["meetings"], reverse=True)
        
        # Chart data
        chart_data = []
        for type_data in org_type_list:
            chart_data.append([
                type_data["desc"], 
                type_data["meetings"], 
                type_data["avg_meetings"],
                type_data["active_orgs"]
            ])
        
        print '<div class="chart-container" id="org-type-meetings-chart"></div>'
        
        # Print chart JavaScript in parts
        print """
        <script type="text/javascript">
            google.charts.setOnLoadCallback(drawOrgTypeMeetingsChart);
            
            function drawOrgTypeMeetingsChart() {
                var data = new google.visualization.DataTable();
                data.addColumn('string', 'Involvement Type');
                data.addColumn('number', 'Total Meetings');
                data.addColumn('number', 'Avg per Org');
                data.addColumn('number', 'Active Orgs');
                
                data.addRows(%s);
                
                var view = new google.visualization.DataView(data);
                view.setColumns([0, 1, 2]);
                
                var options = {
                    title: 'Meetings by Involvement Type',
                    bars: 'horizontal',
                    series: {
                        0: { targetAxisIndex: 0 },
                        1: { targetAxisIndex: 1 }
                    },
                    vAxes: {
                        0: { title: 'Meetings' },
                        1: { title: 'Avg per Org' }
                    }
                };
                
                var chart = new google.visualization.BarChart(document.getElementById('org-type-meetings-chart'));
                chart.draw(view, options);
            }
        </script>
        """ % json.dumps(chart_data)
        
        # Show table of meetings by org type
        print '<table>'
        print '<thead>'
        print '<tr>'
        print '<th>Involvement Type</th>'
        print '<th>Involvements</th>'
        print '<th>Active Orgs</th>'
        print '<th>Total Meetings</th>'
        print '<th>Avg Meetings per Org</th>'
        print '<th>Activity Ratio</th>'
        print '</tr>'
        print '</thead>'
        print '<tbody>'
        
        for type_data in org_type_list:
            activity_ratio = float(type_data["active_orgs"]) / type_data["orgs"] * 100 if type_data["orgs"] > 0 else 0
            
            print '<tr>'
            print '<td>{0}</td>'.format(type_data["desc"])
            print '<td>{0}</td>'.format(type_data["orgs"])
            print '<td>{0}</td>'.format(type_data["active_orgs"])
            print '<td>{0}</td>'.format(type_data["meetings"])
            print '<td>{0:.1f}</td>'.format(type_data["avg_meetings"])
            print '<td><div class="progress" style="margin-bottom: 0;">'
            print '<div class="progress-bar" style="width: {0:.1f}%">{0:.1f}%</div>'.format(activity_ratio)
            print '</div></td>'
            print '</tr>'
        
        print '</tbody>'
        print '</table>'
    else:
        print '<p>No involvement type data available.</p>'
    
    print '</div>'

    
    # Recent meetings list
    print '<div class="card">'
    print '<div class="card-header">Recent Meetings</div>'
    
    if meetings_data:
        print '<table>'
        print '<thead>'
        print '<tr>'
        print '<th>Date</th>'
        print '<th>Involvements</th>'
        print '<th>Description</th>'
        print '<th>Location</th>'
        print '<th>Attendance</th>'
        print '</tr>'
        print '</thead>'
        print '<tbody>'
        
        # Show only the first 25
        max_meetings_to_show = 25
        for i, meeting in enumerate(meetings_data):
            if i >= max_meetings_to_show:
                break
            meeting_date = format_date(meeting.MeetingDate) if hasattr(meeting, 'MeetingDate') else '-'
            did_not_meet = hasattr(meeting, 'DidNotMeet') and meeting.DidNotMeet
            
            print '<tr>'
            print '<td>{0}</td>'.format(meeting_date)
            print '<td><a href="/Org/{0}" target="_blank">{1}</a></td>'.format(
                meeting.OrganizationId, meeting.OrganizationName)
            print '<td>{0}</td>'.format(meeting.Description if hasattr(meeting, 'Description') else '-')
            print '<td>{0}</td>'.format(meeting.Location if hasattr(meeting, 'Location') else '-')
            
            if did_not_meet:
                print '<td><span class="status-inactive">Did Not Meet</span></td>'
            else:
                attendance = meeting.NumPresent if hasattr(meeting, 'NumPresent') else meeting.HeadCount if hasattr(meeting, 'HeadCount') else '-'
                print '<td>{0}</td>'.format(attendance)
            
            print '</tr>'
        
        print '</tbody>'
        print '</table>'
        
        if len(meetings_data) > 20:
            print '<p><em>Showing 20 of {0} recent meetings.</em></p>'.format(len(meetings_data))
    else:
        print '<p>No recent meetings data available.</p>'
    
    print '</div>'
    
    print '</div>'  # Close dashboard-container

def print_stat_box(title, value, color):
    """Print a stat box with the given title, value, and color"""
    print """
    <div class="stat-box" style="border-left-color: {2};">
        <div class="stat-title">{0}</div>
        <div class="stat-value">{1}</div>
    </div>
    """.format(title, format_number(value), color)



# ==============================================================
# EXECUTE MAIN FUNCTION
# ==============================================================
main()
