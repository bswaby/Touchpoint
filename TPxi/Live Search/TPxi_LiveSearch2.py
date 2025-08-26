###############################################################
# TouchPoint Advanced Quick Search - Widget Compatible Version
###############################################################
# This script provides an advanced search interface for TouchPoint
# with real-time results as you type, plus quick task/note creation
# Fully compatible with homepage widgets and standalone use

# Upload Instructions:
# 1. Click Admin ~ Advanced ~ Special Content ~ Python
# 2. Click New Python Script File  
# 3. Name the script (remember this exact name)
# 4. Paste this code and update SCRIPT_NAME below
# 5. Test standalone first, then add as widget
###############################################################################
#### NOTE:  Make sure to update the script name below to what you name it######
###############################################################################
# For Homepage Widget:
# 1. Add keyword "widget" to the script content keywords
# 2. Admin → Advanced → Homepage Widget → Add New Widget
# 3. Select "Python Script" and choose your script name

# written by: Ben Swaby
# email: bswaby@fbchtn.org
# last updated: 06/11/2025
# Version: Widget Compatible with Namespace Isolation

# ::START:: Configuration
MAX_RESULTS = 18  # Maximum number of results to display per category
SEARCH_DELAY = 300  # Milliseconds to wait after typing before searching

# CRITICAL: Set this to your exact script name in TouchPoint
SCRIPT_NAME = 'TPxi_LiveSearch2'  # ⭐ CHANGE THIS TO YOUR ACTUAL SCRIPT NAME

# Giving visibility settings
SHOW_GIVING_IN_JOURNEY = False  # Global setting to hide giving events in journey timeline

# ::START:: Initialize Variables
search_term = ""
ajax_mode = False
embed_mode = False
add_task_mode = False
add_note_mode = False
script_name = SCRIPT_NAME

# ::START:: Detect Script Context
# Determine the actual script name from various sources
script_name_detected = False

try:
    # Method 1: Try to get script name from URL path (works for standalone)
    if hasattr(model.Data, "PATH_INFO") and model.Data.PATH_INFO:
        path_parts = model.Data.PATH_INFO.split('/')
        if len(path_parts) > 2 and path_parts[1] == "PyScript":
            script_name = path_parts[2]
            script_name_detected = True
    
    # Method 2: Check for manual script name override in query parameters
    if not script_name_detected and hasattr(model.Data, "ScriptName") and model.Data.ScriptName:
        script_name = str(model.Data.ScriptName)
        script_name_detected = True
    
    # Method 3: Fallback to configured default
    if not script_name_detected:
        script_name = SCRIPT_NAME
        
except Exception as e:
    script_name = SCRIPT_NAME

# ::START:: Load Keywords Data
try:
    # Load keywords once when the script starts
    preloaded_keywords = []
    
    # Fetch keywords from the database
    keywords_sql = """
    SELECT KeywordId, Code, Description 
    FROM Keyword
    WHERE IsActive = 1
        AND Description IS NOT NULL
        AND Description <> ''
    ORDER BY Description
    """
    
    keywords = q.QuerySql(keywords_sql)
    
    if keywords and len(keywords) > 0:
        for kw in keywords:
            # Add each keyword to our array
            key_id = kw.KeywordId
            key_code = kw.Code or ""
            key_desc = kw.Description or ""
            
            preloaded_keywords.append({
                "KeywordId": key_id,
                "Code": key_code,
                "Description": key_desc
            })
            
    # Convert to a JSON string for embedding in JavaScript
    keywords_json = "["
    for i, kw in enumerate(preloaded_keywords):
        if i > 0:
            keywords_json += ","
        # Manually build JSON object
        code_safe = str(kw["Code"]).replace('"', '\\"').replace("\\", "\\\\")
        desc_safe = str(kw["Description"]).replace('"', '\\"').replace("\\", "\\\\")
        
        keywords_json += '{"KeywordId":' + str(kw["KeywordId"]) + ',"Code":"' + code_safe + '","Description":"' + desc_safe + '"}'
    keywords_json += "]"
except Exception as e:
    # Log error but continue
    import traceback
    keywords_json = "[]"

# ::START:: Process Query Parameters
try:
    if hasattr(model.Data, "q"):
        search_term = str(model.Data.q)
    if hasattr(model.Data, "ajax") and model.Data.ajax == "1":
        ajax_mode = True
    if hasattr(model.Data, "embed") and model.Data.embed == "1":
        embed_mode = True
        model.Header = ""  # Clear the header when in embed mode
    if hasattr(model.Data, "add_task") and model.Data.add_task == "1":
        add_task_mode = True
    if hasattr(model.Data, "add_note") and model.Data.add_note == "1":
        add_note_mode = True
except Exception as e:
    # Print any errors
    import traceback
    print "<h2>Error Processing Parameters</h2>"
    print "<p>An error occurred: " + str(e) + "</p>"

# ::START:: AJAX Handler for Person Journey
# Handle AJAX requests for person journey data - must be before any HTML output
import json

# Function to handle AJAX responses without sys.exit()
def send_ajax_response(response_data):
    """Send JSON response and stop further execution"""
    print json.dumps(response_data)
    return True  # Signal that we've handled the request

# Variable to track if we've handled an AJAX request
ajax_handled = False

# Check if this is an AJAX request early to avoid any HTML output
if hasattr(model.Data, "action") and model.Data.action in ["get_person_journey", "get_family_journey"]:
    ajax_mode = True
    action = model.Data.action
    
    if action == "get_person_journey":
        try:
            people_id = int(model.Data.peopleId) if hasattr(model.Data, "peopleId") else 0
            if not people_id:
                ajax_handled = send_ajax_response({'error': 'No person ID provided'})
            
            # Get basic person info
            person_sql = """
            SELECT 
                p.PeopleId,
                p.Name2 AS Name,
                p.Age,
                p.EmailAddress,
                p.CellPhone,
                ms.Description AS MemberStatus
            FROM People p
            LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
            WHERE p.PeopleId = {0}
            """.format(people_id)
            
            if not ajax_handled:
                person_info = q.QuerySqlTop1(person_sql)
                if not person_info:
                    ajax_handled = send_ajax_response({'error': 'Person not found'})
            
            # Check if user can view giving data
            can_view_giving = False
            finance_roles = ["Finance", "FinanceAdmin", "Admin"]
            for role in finance_roles:
                if model.UserIsInRole(role):
                    can_view_giving = True
                    break
            
            # Test actual database access to Contribution table
            if can_view_giving:
                try:
                    # Try a simple query to test permissions
                    test_sql = "SELECT TOP 1 ContributionId FROM Contribution"
                    q.QuerySqlTop1(test_sql)
                    # If we get here, user has database permissions
                except:
                    # User has role but not database permissions
                    can_view_giving = False
            
            # Build giving section based on permissions AND global setting
            if can_view_giving and SHOW_GIVING_IN_JOURNEY:
                giving_section = """
                UNION ALL
                
                -- 6. First contribution
                SELECT 
                    TOP 1 c.ContributionDate AS EventDate,
                    'First Contribution' AS EventType,
                    'Started giving' AS Description,
                    'giving' AS Category,
                    6 AS SortOrder
                FROM Contribution c
                WHERE c.PeopleId = {0}
                AND c.ContributionTypeId != 99
                AND c.ContributionAmount > 0
                ORDER BY c.ContributionDate""".format(people_id)
            else:
                giving_section = ""
            
            # Get journey events (comprehensive version from EngagementIndex)
            events_sql = """
            WITH JourneyEvents AS (
                -- 1. Person added to system
                SELECT 
                    p.CreatedDate AS EventDate,
                    'Added to System' AS EventType,
                    'First contact/registration' AS Description,
                    'system' AS Category,
                    1 AS SortOrder
                FROM People p WHERE p.PeopleId = {0}
                
                UNION ALL
                
                -- 2. First program joined (organization membership)
                SELECT 
                    TOP 1 om.CreatedDate AS EventDate,
                    'First Program Joined' AS EventType,
                    o.OrganizationName AS Description,
                    'program' AS Category,
                    2 AS SortOrder
                FROM OrganizationMembers om
                JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                WHERE om.PeopleId = {0}
                AND o.OrganizationStatusId = 30
                ORDER BY om.CreatedDate
                
                UNION ALL
                
                -- 3. First attendance
                SELECT 
                    TOP 1 a.MeetingDate AS EventDate,
                    'First Attendance' AS EventType,
                    o.OrganizationName AS Description,
                    'attendance' AS Category,
                    3 AS SortOrder
                FROM Attend a
                JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                WHERE a.PeopleId = {0}
                AND a.AttendanceFlag = 1
                ORDER BY a.MeetingDate
                
                UNION ALL
                
                -- 4. Small group joined (using Connect Groups program)
                SELECT 
                    TOP 1 om.CreatedDate AS EventDate,
                    'Small Group Joined' AS EventType,
                    o.OrganizationName AS Description,
                    'smallgroup' AS Category,
                    4 AS SortOrder
                FROM OrganizationMembers om
                JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                JOIN Division d ON o.DivisionId = d.Id
                WHERE om.PeopleId = {0}
                AND o.OrganizationStatusId = 30
                AND d.ProgId IN (1128)  -- Connect Groups program ID
                ORDER BY om.CreatedDate
                
                UNION ALL
                
                -- 5. Started serving (multiple detection methods)
                SELECT TOP 1 EventDate, EventType, Description, Category, SortOrder
                FROM (
                    -- AttendType-based serving (when they serve in classrooms)
                    SELECT 
                        MIN(a.MeetingDate) AS EventDate,
                        'Started Serving' AS EventType,
                        o.OrganizationName + ' (Volunteer)' AS Description,
                        'serving' AS Category,
                        5 AS SortOrder
                    FROM Attend a
                    JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                    WHERE a.PeopleId = {0}
                    AND a.AttendanceFlag = 1
                    AND a.AttendanceTypeId IN (10, 20)  -- AttendType IDs for serving roles
                    GROUP BY o.OrganizationName
                    
                    UNION ALL
                    
                    -- Member type based serving (leadership roles)
                    SELECT 
                        om.CreatedDate AS EventDate,
                        'Started Serving' AS EventType,
                        o.OrganizationName + ' (' + ISNULL(mt.Description, 'Leadership') + ')' AS Description,
                        'serving' AS Category,
                        5 AS SortOrder
                    FROM OrganizationMembers om
                    JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                    LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                    WHERE om.PeopleId = {0}
                    AND om.MemberTypeId > 100  -- High member type IDs indicate leadership
                ) AS ServingEvents
                ORDER BY EventDate
                
                {1}
            )
            SELECT * FROM JourneyEvents
            WHERE EventDate IS NOT NULL
            ORDER BY EventDate, SortOrder
            """.format(people_id, giving_section)
            
            events = q.QuerySql(events_sql)
            
            # Get last attendance and recent activity stats
            recent_activity_sql = """
            SELECT 
                -- Last attendance
                (SELECT TOP 1 CONVERT(varchar, a.MeetingDate, 101) + ' - ' + o.OrganizationName
                 FROM Attend a
                 JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                 WHERE a.PeopleId = {0} AND a.AttendanceFlag = 1
                 ORDER BY a.MeetingDate DESC) AS LastAttendance,
                
                -- Last registration/signup
                (SELECT TOP 1 CONVERT(varchar, om.EnrollmentDate, 101) + ' - ' + o.OrganizationName
                 FROM OrganizationMembers om
                 JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                 WHERE om.PeopleId = {0}
                 ORDER BY om.EnrollmentDate DESC) AS LastSignup,
                
                -- Attendance count last 90 days
                (SELECT COUNT(*)
                 FROM Attend a
                 WHERE a.PeopleId = {0} 
                 AND a.AttendanceFlag = 1
                 AND a.MeetingDate >= DATEADD(day, -90, GETDATE())) AS AttendanceCount90Days,
                
                -- Days since last attendance
                (SELECT DATEDIFF(day, MAX(a.MeetingDate), GETDATE())
                 FROM Attend a
                 WHERE a.PeopleId = {0} AND a.AttendanceFlag = 1) AS DaysSinceLastAttendance,
                
                -- Currently serving count
                (SELECT COUNT(DISTINCT om.OrganizationId)
                 FROM OrganizationMembers om
                 WHERE om.PeopleId = {0}
                 AND om.MemberTypeId > 100
                 AND om.InactiveDate IS NULL) AS CurrentlyServingCount
            """.format(people_id)
            
            recent_activity = q.QuerySqlTop1(recent_activity_sql)
            
            # Convert events to list
            events_list = []
            icon_map = {
                'system': {'icon': 'fa-user-plus', 'color': '#6c757d'},
                'program': {'icon': 'fa-sitemap', 'color': '#ffc107'},
                'attendance': {'icon': 'fa-calendar-check-o', 'color': '#007bff'},
                'smallgroup': {'icon': 'fa-users', 'color': '#28a745'},
                'serving': {'icon': 'fa-hands-helping', 'color': '#17a2b8'},
                'giving': {'icon': 'fa-gift', 'color': '#e83e8c'}
            }
            
            for event in events:
                # Skip giving events based on global setting or permissions
                if event.Category == 'giving' and (not SHOW_GIVING_IN_JOURNEY or not can_view_giving):
                    continue
                    
                icon_info = icon_map.get(event.Category, {'icon': 'fa-circle', 'color': '#6c757d'})
                date_str = str(event.EventDate).split(' ')[0] if event.EventDate else ''
                
                events_list.append({
                    'date': date_str,
                    'event': event.EventType,
                    'description': event.Description,
                    'type': event.Category,
                    'icon': icon_info['icon'],
                    'color': icon_info['color']
                })
            
            if ajax_handled:
                # Stop processing if we already sent an error
                pass
            else:
                # Calculate more current engagement score
                score = 10  # Base score
                
                # Recent attendance (most important for current engagement)
                if recent_activity and recent_activity.AttendanceCount90Days:
                    attendance_90 = recent_activity.AttendanceCount90Days
                    if attendance_90 >= 9:  # Weekly attendance
                        score += 40
                    elif attendance_90 >= 6:  # 2x per month
                        score += 30
                    elif attendance_90 >= 3:  # Monthly
                        score += 20
                    elif attendance_90 >= 1:  # Occasional
                        score += 10
                
                # Days since last attendance (recency matters)
                if recent_activity and recent_activity.DaysSinceLastAttendance is not None:
                    days_since = recent_activity.DaysSinceLastAttendance
                    if days_since <= 7:
                        score += 15  # Very recent
                    elif days_since <= 14:
                        score += 10
                    elif days_since <= 30:
                        score += 5
                    elif days_since > 90:
                        score -= 10  # Penalty for long absence
                
                # Currently serving
                if recent_activity and recent_activity.CurrentlyServingCount and recent_activity.CurrentlyServingCount > 0:
                    score += 20
                
                # In a program
                if any(e['type'] == 'program' for e in events_list):
                    score += 10
                
                # Small group participation
                if any(e['type'] == 'smallgroup' for e in events_list):
                    score += 15
                
                # Cap score at 100
                score = min(score, 100)
                score = max(score, 0)
                
                # Format recent activity data
                last_attendance = recent_activity.LastAttendance if recent_activity else None
                last_signup = recent_activity.LastSignup if recent_activity else None
                attendance_90_days = recent_activity.AttendanceCount90Days if recent_activity else 0
                days_since_attendance = recent_activity.DaysSinceLastAttendance if recent_activity else None
                serving_count = recent_activity.CurrentlyServingCount if recent_activity else 0
                
                response = {
                    'person': {
                        'name': person_info.Name,
                        'age': person_info.Age or 0,
                        'email': person_info.EmailAddress or '',
                        'phone': person_info.CellPhone or '',
                        'member_status': person_info.MemberStatus or 'Unknown',
                        'engagement_score': score,
                        'can_view_giving': can_view_giving  # Include permission info
                    },
                    'journey': events_list,
                    'insights': {
                        'journey_length': str(len(events_list)) + ' events',
                        'entry_point': events_list[1]['description'] if len(events_list) > 1 else 'Not yet engaged',
                        'total_events': len(events_list)
                    },
                    'recent_activity': {
                        'last_attendance': last_attendance,
                        'last_signup': last_signup,
                        'attendance_90_days': attendance_90_days,
                        'days_since_attendance': days_since_attendance,
                        'currently_serving': serving_count,
                        'engagement_status': 'Active' if days_since_attendance and days_since_attendance <= 30 else 'Inactive' if days_since_attendance and days_since_attendance > 90 else 'Occasional'
                    }
                }
            
            if not ajax_handled:
                ajax_handled = send_ajax_response(response)
            
        except Exception as e:
            if not ajax_handled:
                ajax_handled = send_ajax_response({'error': str(e)})
    
    elif action == "get_family_journey" and not ajax_handled:
        try:
            people_id = int(model.Data.peopleId) if hasattr(model.Data, "peopleId") else 0
            if not people_id:
                ajax_handled = send_ajax_response({'error': 'No person ID provided'})
            
            # Check if user can view giving data
            can_view_giving = False
            finance_roles = ["Finance", "FinanceAdmin", "Admin"]
            for role in finance_roles:
                if model.UserIsInRole(role):
                    can_view_giving = True
                    break
            
            # Test actual database access to Contribution table
            if can_view_giving:
                try:
                    # Try a simple query to test permissions
                    test_sql = "SELECT TOP 1 ContributionId FROM Contribution"
                    q.QuerySqlTop1(test_sql)
                    # If we get here, user has database permissions
                except:
                    # User has role but not database permissions
                    can_view_giving = False
            
            # Get family members
            family_sql = """
            SELECT 
                f.PeopleId,
                f.Name2 AS Name,
                f.Age,
                f.PositionInFamilyId,
                fp.Description AS FamilyPosition,
                ms.Description AS MemberStatus,
                p.FamilyId
            FROM People p
            JOIN People f ON p.FamilyId = f.FamilyId
            LEFT JOIN lookup.FamilyPosition fp ON f.PositionInFamilyId = fp.Id
            LEFT JOIN lookup.MemberStatus ms ON f.MemberStatusId = ms.Id
            WHERE p.PeopleId = {0}
            AND f.IsDeceased = 0
            AND f.ArchivedFlag = 0
            ORDER BY f.PositionInFamilyId, f.Age DESC
            """.format(people_id)
            
            if not ajax_handled:
                family_members_data = q.QuerySql(family_sql)
                if not family_members_data:
                    ajax_handled = send_ajax_response({'error': 'Family not found'})
            
            # Process each family member
            family_members = []
            family_id = 0
            
            for member in family_members_data:
                family_id = member.FamilyId
                
                # Get comprehensive journey events for this member (matching individual timeline logic)
                # Build the complete SQL query directly
                if can_view_giving and SHOW_GIVING_IN_JOURNEY:
                    member_events_sql = """
                    WITH MemberJourney AS (
                        -- System: Person added
                        SELECT 
                            p.CreatedDate AS EventDate,
                            'Added to System' AS EventType,
                            'Profile created in database' AS Description,
                            'system' AS Category,
                            1 AS SortOrder
                        FROM People p
                        WHERE p.PeopleId = {0}
                        
                        UNION ALL
                        
                        -- Program: First program joined
                        SELECT TOP 1
                            om.EnrollmentDate AS EventDate,
                            'Joined Program' AS EventType,
                            ISNULL(p.Name, 'Unknown Program') AS Description,
                            'program' AS Category,
                            2 AS SortOrder
                        FROM OrganizationMembers om
                        JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                        LEFT JOIN Division d ON o.DivisionId = d.Id
                        LEFT JOIN Program p ON d.ProgId = p.Id
                        WHERE om.PeopleId = {0}
                        AND o.OrganizationStatusId = 30
                        AND om.EnrollmentDate IS NOT NULL
                        ORDER BY om.EnrollmentDate ASC
                        
                        UNION ALL
                        
                        -- Attendance: First attendance
                        SELECT TOP 1
                            a.MeetingDate AS EventDate,
                            'First Attendance' AS EventType,
                            o.OrganizationName AS Description,
                            'attendance' AS Category,
                            3 AS SortOrder
                        FROM Attend a
                        JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                        WHERE a.PeopleId = {0}
                        AND a.AttendanceFlag = 1
                        ORDER BY a.MeetingDate ASC
                        
                        UNION ALL
                        
                        -- Small Group: Connect Groups (program ID 1128)
                        SELECT TOP 1
                            om.EnrollmentDate AS EventDate,
                            'Joined Small Group' AS EventType,
                            o.OrganizationName AS Description,
                            'smallgroup' AS Category,
                            4 AS SortOrder
                        FROM OrganizationMembers om
                        JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                        LEFT JOIN Division d ON o.DivisionId = d.Id
                        WHERE om.PeopleId = {0}
                        AND d.ProgId = 1128
                        AND om.EnrollmentDate IS NOT NULL
                        ORDER BY om.EnrollmentDate ASC
                        
                        UNION ALL
                        
                        -- Serving: AttendType-based serving
                        SELECT TOP 1
                            a.MeetingDate AS EventDate,
                            'Started Serving' AS EventType,
                            o.OrganizationName + ' (' + at.Description + ')' AS Description,
                            'serving' AS Category,
                            6 AS SortOrder
                        FROM Attend a
                        JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                        LEFT JOIN lookup.AttendType at ON a.AttendanceTypeId = at.Id
                        WHERE a.PeopleId = {0}
                        AND a.AttendanceFlag = 1
                        AND a.AttendanceTypeId IN (10, 20)
                        ORDER BY a.MeetingDate ASC
                        
                        UNION ALL
                        
                        -- Serving: Member type based serving (MemberTypeId > 100)
                        SELECT TOP 1
                            om.EnrollmentDate AS EventDate,
                            'Leadership Role' AS EventType,
                            o.OrganizationName + ' (' + ISNULL(mt.Description, 'Leader') + ')' AS Description,
                            'serving' AS Category,
                            6 AS SortOrder
                        FROM OrganizationMembers om
                        JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                        LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                        WHERE om.PeopleId = {0}
                        AND om.MemberTypeId > 100
                        AND om.EnrollmentDate IS NOT NULL
                        ORDER BY om.EnrollmentDate ASC
                        
                        UNION ALL
                        
                        -- First contribution (only for users with Finance permissions)
                        SELECT TOP 1
                            c.ContributionDate AS EventDate,
                            'Started Giving' AS EventType,
                            'First contribution of $' + CAST(c.ContributionAmount AS VARCHAR(20)) AS Description,
                            'giving' AS Category,
                            5 AS SortOrder
                        FROM Contribution c
                        WHERE c.PeopleId = {0}
                        AND c.ContributionTypeId != 99
                        ORDER BY c.ContributionDate ASC
                    )
                    SELECT TOP 10 EventDate, EventType, Description, Category, SortOrder
                    FROM MemberJourney
                    WHERE EventDate IS NOT NULL
                    ORDER BY EventDate ASC, SortOrder ASC
                    """.format(member.PeopleId)
                else:
                    # Query without giving section
                    member_events_sql = """
                    WITH MemberJourney AS (
                        -- System: Person added
                        SELECT 
                            p.CreatedDate AS EventDate,
                            'Added to System' AS EventType,
                            'Profile created in database' AS Description,
                            'system' AS Category,
                            1 AS SortOrder
                        FROM People p
                        WHERE p.PeopleId = {0}
                        
                        UNION ALL
                        
                        -- Program: First program joined
                        SELECT TOP 1
                            om.EnrollmentDate AS EventDate,
                            'Joined Program' AS EventType,
                            ISNULL(p.Name, 'Unknown Program') AS Description,
                            'program' AS Category,
                            2 AS SortOrder
                        FROM OrganizationMembers om
                        JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                        LEFT JOIN Division d ON o.DivisionId = d.Id
                        LEFT JOIN Program p ON d.ProgId = p.Id
                        WHERE om.PeopleId = {0}
                        AND o.OrganizationStatusId = 30
                        AND om.EnrollmentDate IS NOT NULL
                        ORDER BY om.EnrollmentDate ASC
                        
                        UNION ALL
                        
                        -- Attendance: First attendance
                        SELECT TOP 1
                            a.MeetingDate AS EventDate,
                            'First Attendance' AS EventType,
                            o.OrganizationName AS Description,
                            'attendance' AS Category,
                            3 AS SortOrder
                        FROM Attend a
                        JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                        WHERE a.PeopleId = {0}
                        AND a.AttendanceFlag = 1
                        ORDER BY a.MeetingDate ASC
                        
                        UNION ALL
                        
                        -- Small Group: Connect Groups (program ID 1128)
                        SELECT TOP 1
                            om.EnrollmentDate AS EventDate,
                            'Joined Small Group' AS EventType,
                            o.OrganizationName AS Description,
                            'smallgroup' AS Category,
                            4 AS SortOrder
                        FROM OrganizationMembers om
                        JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                        LEFT JOIN Division d ON o.DivisionId = d.Id
                        WHERE om.PeopleId = {0}
                        AND d.ProgId = 1128
                        AND om.EnrollmentDate IS NOT NULL
                        ORDER BY om.EnrollmentDate ASC
                        
                        UNION ALL
                        
                        -- Serving: AttendType-based serving
                        SELECT TOP 1
                            a.MeetingDate AS EventDate,
                            'Started Serving' AS EventType,
                            o.OrganizationName + ' (' + at.Description + ')' AS Description,
                            'serving' AS Category,
                            6 AS SortOrder
                        FROM Attend a
                        JOIN Organizations o ON a.OrganizationId = o.OrganizationId
                        LEFT JOIN lookup.AttendType at ON a.AttendanceTypeId = at.Id
                        WHERE a.PeopleId = {0}
                        AND a.AttendanceFlag = 1
                        AND a.AttendanceTypeId IN (10, 20)
                        ORDER BY a.MeetingDate ASC
                        
                        UNION ALL
                        
                        -- Serving: Member type based serving (MemberTypeId > 100)
                        SELECT TOP 1
                            om.EnrollmentDate AS EventDate,
                            'Leadership Role' AS EventType,
                            o.OrganizationName + ' (' + ISNULL(mt.Description, 'Leader') + ')' AS Description,
                            'serving' AS Category,
                            6 AS SortOrder
                        FROM OrganizationMembers om
                        JOIN Organizations o ON om.OrganizationId = o.OrganizationId
                        LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
                        WHERE om.PeopleId = {0}
                        AND om.MemberTypeId > 100
                        AND om.EnrollmentDate IS NOT NULL
                        ORDER BY om.EnrollmentDate ASC
                    )
                    SELECT TOP 10 EventDate, EventType, Description, Category, SortOrder
                    FROM MemberJourney
                    WHERE EventDate IS NOT NULL
                    ORDER BY EventDate ASC, SortOrder ASC
                    """.format(member.PeopleId)
                
                member_events = q.QuerySql(member_events_sql)
                
                # Convert events with comprehensive icon mapping
                journey_events = []
                icon_map = {
                    'system': {'icon': 'fa-user-plus', 'color': '#6c757d'},
                    'program': {'icon': 'fa-users', 'color': '#17a2b8'},
                    'attendance': {'icon': 'fa-calendar-check-o', 'color': '#007bff'},
                    'smallgroup': {'icon': 'fa-users', 'color': '#28a745'},
                    'serving': {'icon': 'fa-hands-helping', 'color': '#fd7e14'},
                    'giving': {'icon': 'fa-heart', 'color': '#dc3545'}
                }
                
                for event in member_events:
                    # Skip giving events based on global setting or permissions
                    if event.Category == 'giving' and (not SHOW_GIVING_IN_JOURNEY or not can_view_giving):
                        continue
                        
                    icon_info = icon_map.get(event.Category, {'icon': 'fa-circle', 'color': '#6c757d'})
                    date_str = str(event.EventDate).split(' ')[0] if event.EventDate else ''
                    
                    journey_events.append({
                        'date': date_str,
                        'event': event.EventType,
                        'description': event.Description,
                        'type': event.Category,
                        'icon': icon_info['icon'],
                        'color': icon_info['color']
                    })
                
                # Calculate simple score
                score = 20  # Base
                if len(journey_events) > 1:
                    score += 30  # Has activity
                
                family_members.append({
                    'person': {
                        'people_id': member.PeopleId,
                        'name': member.Name,
                        'age': member.Age or 0,
                        'position': member.FamilyPosition or 'Family Member',
                        'member_status': member.MemberStatus or 'Unknown',
                        'engagement_score': score
                    },
                    'journey': journey_events,
                    'insights': {
                        'journey_length': str(len(journey_events)) + ' events',
                        'entry_point': journey_events[0]['description'] if journey_events else 'Not engaged',
                        'total_events': len(journey_events)
                    }
                })
            
            # Calculate family stats
            total_members = len(family_members)
            engaged_members = sum(1 for m in family_members if m['person']['engagement_score'] >= 40)
            avg_score = sum(m['person']['engagement_score'] for m in family_members) / total_members if total_members > 0 else 0
            
            response = {
                'family_info': {
                    'family_id': family_id,
                    'total_members': total_members,
                    'engaged_members': engaged_members,
                    'avg_engagement': int(avg_score),
                    'timeline_start': '2023-01-01',
                    'timeline_end': str(model.DateTime).split(' ')[0]
                },
                'family_members': family_members
            }
            
            if not ajax_handled:
                ajax_handled = send_ajax_response(response)
            
        except Exception as e:
            if not ajax_handled:
                ajax_handled = send_ajax_response({'error': str(e)})

# Stop processing if we handled an AJAX request
if ajax_handled:
    # Don't process any more of the script
    pass
else:
    # ::START:: Form Submission Handler
    # Process task/note form submissions
    if model.HttpMethod == 'post':
        try:
            # Get common form fields
            action_type = ""
            about_person_id = 0
            assignee_id = None
            message = ""
            
            if hasattr(model.Data, "action_type"):
                action_type = model.Data.action_type
            if hasattr(model.Data, "about_person_id") and model.Data.about_person_id:
                about_person_id = int(model.Data.about_person_id)
            if hasattr(model.Data, "assignee_id") and model.Data.assignee_id:
                assignee_id = int(model.Data.assignee_id)
            if hasattr(model.Data, "message"):
                message = model.Data.message
            
            # Get keywords (supports multiple selections)
            keywords = []
            for i in range(20):  # Allow up to 20 keywords
                keyword_field = "keyword_" + str(i)
                if hasattr(model.Data, keyword_field) and getattr(model.Data, keyword_field):
                    keywords.append(int(getattr(model.Data, keyword_field)))
            
            # ::STEP:: Process Task Creation
            if action_type == "task" and about_person_id > 0:
                # Get due date if it exists
                due_date = None
                
                if hasattr(model.Data, "due_date") and model.Data.due_date:
                    # Try to parse the date string to a date object
                    try:
                        # Get the date string from the form
                        date_string = str(model.Data.due_date).strip()
                        
                        # Only proceed if we have a non-empty string
                        if date_string:
                            # Try parsing with model.ParseDate - TouchPoint's date parser
                            due_date = model.ParseDate(date_string)
                    except Exception as date_error:
                        due_date = None
                
                # Create the task
                owner_id = model.UserPeopleId  # Current user
                
                # Create the task using TouchPoint API with correct parameters
                try:
                    task_id = model.CreateTaskNote(
                        owner_id, 
                        about_person_id, 
                        assignee_id, 
                        None,  # roleId 
                        0,     # isNote as integer 0
                        message,  # instructions for task
                        "",    # notes can be empty for a simple task
                        due_date,  # Now passing a proper date object
                        keywords, 
                        True   # sendEmails
                    )
                    
                    # Return success message
                    print('<div class="tpls-alert tpls-alert-success">Task created successfully!</div>')
                    
                except Exception as task_error:
                    print('<div class="tpls-alert tpls-alert-danger">Error creating task: {}</div>'.format(task_error))
            
            # ::STEP:: Process Note Creation    
            elif action_type == "note" and about_person_id > 0:
                # Create the note
                owner_id = model.UserPeopleId  # Current user
                
                # Create the note using TouchPoint API
                note_id = model.CreateTaskNote(
                    owner_id, 
                    about_person_id, 
                    assignee_id, 
                    None,  # roleId
                    1,  # isNote - use 1 for true
                    "",  # instructions can be empty for a note
                    message,  # notes contains the actual note content
                    None,  # due_date (not needed for notes)
                    keywords,
                    False  # sendEmails (not needed for notes)
                )
                
                # Return success message
                print('<div class="tpls-alert tpls-alert-success">Note added successfully!</div>')
        
        except Exception as e:
            # Print error message with more detailed debugging
            import traceback
            print('<div class="tpls-alert tpls-alert-danger">Error: ' + str(e) + '</div>')

    # ::START:: Iframe Resize Handler
    # Send postMessage to parent window for iframe resizing - Enhanced for layout control
    if embed_mode:
        resize_script = """
    <script>
    // Conservative height management for iframe
        function tplsNotifyParentOfHeight() {
            try {
                var container = document.querySelector('.tpls-search-container');
                var height = 400; // Default safe height
                
                if (container) {
                    height = Math.max(
                        container.scrollHeight + 20,
                        container.offsetHeight + 20,
                        400 // Minimum height
                    );
                }
                
                // Check if any modals are open and adjust
                var modals = document.querySelectorAll('.tpls-modal');
                var modalHeight = 0;
                
                modals.forEach(function(modal) {
                    if (modal && modal.style && modal.style.display === 'block') {
                        var modalContent = modal.querySelector('.tpls-modal-content');
                        if (modalContent) {
                            modalHeight = Math.max(modalHeight, modalContent.offsetHeight + 200);
                        }
                    }
                });
                
                if (modalHeight > 0) {
                    height = Math.max(height, modalHeight);
                }
                
                // Cap maximum height to prevent huge widgets
                height = Math.min(height, 800);
                
                if (window.parent && window.parent !== window) {
                    window.parent.postMessage({
                        type: 'resize',
                        height: height,
                        modalHeight: modalHeight,
                        source: 'tpls-widget'
                    }, '*');
                }
            } catch (e) {
                console.error('Error sending height to parent:', e);
            }
        }
    
        // Targeted footer cleanup - fix the specific footers shown in the image
        function cleanupFooters() {
            try {
                // Target the specific footer elements visible in the screenshots
                var footerKillSelectors = [
                    '#footer.hidden-print',      // The footer div with id="footer" class="hidden-print"
                    '#contact-footer.hidden-print', // The contact footer div
                    '#footer',                   // Just the footer id
                    '#contact-footer',           // Just the contact footer id
                    '.hidden-print'              // Any element with hidden-print class outside our container
                ];
                
                footerKillSelectors.forEach(function(selector) {
                    try {
                        var elements = document.querySelectorAll(selector);
                        elements.forEach(function(el) {
                            if (el && !el.closest('.tpls-search-container')) {
                                el.style.display = 'none !important';
                                el.style.visibility = 'hidden !important';
                                el.style.height = '0px !important';
                                el.style.overflow = 'hidden !important';
                                el.style.position = 'absolute !important';
                                el.style.left = '-9999px !important';
                            }
                        });
                    } catch (e) {
                        // Silently continue if selector fails
                    }
                });
                
                // Target TouchPoint branded footers by background color and content
                var brandedElements = document.querySelectorAll('div[style*="background-color: rgb(74, 144, 164)"], div[style*="background-color:#4a90a4"]');
                brandedElements.forEach(function(el) {
                    if (el && !el.closest('.tpls-search-container')) {
                        var text = (el.textContent || '').toLowerCase();
                        if (text.includes('touchpoint') || text.includes('first baptist') || text.includes('powered by')) {
                            el.style.display = 'none !important';
                            el.style.visibility = 'hidden !important';
                            el.style.height = '0px !important';
                            el.style.position = 'absolute !important';
                            el.style.left = '-9999px !important';
                        }
                    }
                });
                
            } catch (e) {
                // Silently handle errors to prevent breaking the widget
            }
        }

    document.addEventListener('DOMContentLoaded', function() {
        // Enhanced mutation observer for catching dynamic footers
        const observer = new MutationObserver(function(mutations) {
            var shouldCleanup = false;
            
            mutations.forEach(function(mutation) {
                // Check for added nodes that might be footers
                if (mutation.addedNodes.length > 0) {
                    mutation.addedNodes.forEach(function(node) {
                        if (node.nodeType === 1) { // Element node
                            var text = node.textContent || '';
                            var id = node.id || '';
                            var className = node.className || '';
                            
                            // Check if this looks like a footer
                            if (text.toLowerCase().includes('touchpoint') ||
                                text.toLowerCase().includes('first baptist') ||
                                id.includes('footer') ||
                                className.includes('footer') ||
                                node.tagName === 'FOOTER') {
                                console.log('Detected footer addition:', node.tagName, id, className, text.substring(0, 50));
                                shouldCleanup = true;
                            }
                        }
                    });
                }
                
                // Check for style changes that might create footers
                if (mutation.type === 'attributes' && mutation.attributeName === 'style') {
                    var target = mutation.target;
                    var style = target.style;
                    if (style.backgroundColor === 'rgb(74, 144, 164)' || 
                        style.backgroundColor === 'rgb(91, 109, 115)') {
                        console.log('Detected footer styling on:', target.tagName, target.className);
                        shouldCleanup = true;
                    }
                }
            });
            
            if (shouldCleanup) {
                // Immediate cleanup
                setTimeout(cleanupFooters, 10);
                // Backup cleanup
                setTimeout(cleanupFooters, 100);
            }
            
            // Always update height on content changes
            setTimeout(tplsNotifyParentOfHeight, 50);
        });
        
        // Observe with comprehensive settings
        observer.observe(document.body, { 
            childList: true, 
            subtree: true,
            attributes: true,
            attributeFilter: ['style', 'class', 'id'],
            characterData: true
        });
        
        // Initial setup
        cleanupFooters();
        tplsNotifyParentOfHeight();
        
        // Aggressive initial cleanup schedule
        setTimeout(cleanupFooters, 50);
        setTimeout(cleanupFooters, 200);
        setTimeout(cleanupFooters, 500);
        setTimeout(cleanupFooters, 1000);
        setTimeout(cleanupFooters, 2000);
        
        // Periodic maintenance with more frequent checks
        setInterval(function() {
            cleanupFooters();
            tplsNotifyParentOfHeight();
        }, 1000); // Every second instead of every 2 seconds
        
        // Quick initial updates
        setTimeout(tplsNotifyParentOfHeight, 100);
        setTimeout(tplsNotifyParentOfHeight, 300);
    });
    
    // Make functions globally available
    window.tplsNotifyParentOfHeight = tplsNotifyParentOfHeight;
    window.cleanupFooters = cleanupFooters;
    </script>
    """
    else:
        resize_script = ""

    # Set page title only if not in embed mode
    if not embed_mode:
        model.Header = "Quick Search"

    # ::START:: Main HTML Output
    # Main HTML template - skip if we already handled an AJAX request
    if not ajax_handled and not ajax_mode and not add_task_mode and not add_note_mode:
        # In embed mode, add styles to hide TouchPoint UI elements and fix positioning
        if embed_mode:
            embed_styles = """
        <style>
        /* ULTRA AGGRESSIVE FOOTER HIDING - Kill everything that looks like a footer */
        body > div:last-child,
        body > footer,
        html > footer,
        .wrapper > div:last-child,
        .content-wrapper > div:last-child,
        div[id*="footer"],
        #footer\\.hidden-print,
        .footer.hidden-print,
        div[class*="footer"][class*="hidden"],
        div[class*="hidden-print"],
        [style*="background-color: rgb(74, 144, 164)"],
        [style*="background-color: rgb(91, 109, 115)"],
        [style*="background-color: #4a90a4"],
        [style*="background-color: #5b6d73"],
        [style*="background-color:#4a90a4"],
        [style*="background-color:#5b6d73"],
        [style*="background: #4a90a4"],
        [style*="background: #5b6d73"],
        [style*="background:#4a90a4"],
        [style*="background:#5b6d73"],
        div[style*="color: white"],
        div[style*="color:white"],
        div[style*="text-align: center"]:last-child,
        div[style*="text-align:center"]:last-child,
        .navbar, 
        .breadcrumb,
        footer,
        .box-header,
        .box-footer,
        .main-sidebar,
        .main-header,
        .content-header,
        .sidebar,
        #main-footer,
        .dropdown-toggle,
        .btn-toolbar,
        .footer,
        .main-footer,
        .content-header h1,
        .control-sidebar,
        .page-footer,
        .wrapper > footer,
        div[class*="footer"],
        div[id*="footer"],
        .footer-info,
        .page-footer,
        body > footer,
        html > footer,
        div.footer,
        footer.footer,
        .touchpoint-footer,
        .footer-wrapper,
        .site-footer,
        #footer,
        .footer-content,
        .footer-bar,
        .footer-section,
        .site-info,
        .powered-by,
        div[style*="background"] footer,
        div[style*="background-color"] footer,
        .footer-powered,
        .footer-brand,
        .footer-links,
        [class*="footer-"],
        [id*="footer-"],
        .container-fluid .row:last-child,
        .content-wrapper > .row:last-child,
        .content > .row:last-child {
            display: none !important;
            height: 0 !important;
            width: 0 !important;
            visibility: hidden !important;
            margin: 0 !important;
            padding: 0 !important;
            position: absolute !important;
            top: -99999px !important;
            left: -99999px !important;
            opacity: 0 !important;
            z-index: -9999 !important;
            overflow: hidden !important;
            max-height: 0 !important;
        }
        
        /* Kill any element that might be a footer by position */
        body > *:last-child:not(.tpls-search-container),
        .wrapper > *:last-child,
        .content-wrapper > *:last-child,
        .container-fluid > *:last-child,
        body > div:last-of-type:not(.tpls-search-container) {
            display: none !important;
        }
        
        /* WIDGET CONTAINMENT - Force everything to stay within bounds */
        body, html {
            padding: 0 !important;
            margin: 0 !important;
            overflow: hidden !important;
            background: transparent !important;
            position: relative !important;
            max-width: 100% !important;
            width: 100% !important;
            box-sizing: border-box !important;
            height: auto !important;
            max-height: 800px !important;
        }
        
        /* Container constraints */
        .content-wrapper, 
        .right-side, 
        .main-footer {
            margin: 0 !important;
            background: transparent !important;
            position: relative !important;
            max-width: 100% !important;
            width: 100% !important;
            overflow: hidden !important;
            box-sizing: border-box !important;
        }
        
        .container-fluid {
            padding: 0 !important;
            margin: 0 !important;
            background: transparent !important;
            position: relative !important;
            max-width: 100% !important;
            width: 100% !important;
            overflow: hidden !important;
            box-sizing: border-box !important;
        }
        
        .content {
            padding: 0 !important;
            margin: 0 !important;
            background: transparent !important;
            position: relative !important;
            max-width: 100% !important;
            width: 100% !important;
            overflow: hidden !important;
            box-sizing: border-box !important;
        }
        
        .box {
            border: none !important;
            box-shadow: none !important;
            margin: 0 !important;
            background: transparent !important;
            position: relative !important;
            max-width: 100% !important;
            width: 100% !important;
            overflow: hidden !important;
            box-sizing: border-box !important;
        }
        
        .box-body {
            padding: 0 !important;
            margin: 0 !important;
            background: transparent !important;
            position: relative !important;
            max-width: 100% !important;
            width: 100% !important;
            overflow: hidden !important;
            box-sizing: border-box !important;
        }
        
        .wrapper {
            background: transparent !important;
            min-height: auto !important;
            height: auto !important;
            position: relative !important;
            margin: 0 !important;
            padding: 0 !important;
            max-width: 100% !important;
            width: 100% !important;
            overflow: hidden !important;
            box-sizing: border-box !important;
        }
        
        .row {
            margin: 0 !important;
            background: transparent !important;
            position: relative !important;
            max-width: 100% !important;
            width: 100% !important;
            overflow: hidden !important;
            box-sizing: border-box !important;
        }
        
        [class^="col-"] {
            padding: 0 !important;
            margin: 0 !important;
            background: transparent !important;
            position: relative !important;
            max-width: 100% !important;
            width: 100% !important;
            overflow: hidden !important;
            box-sizing: border-box !important;
        }
        
        /* Force search container to be the ONLY visible content */
        .tpls-search-container {
            position: relative !important;
            margin: 0 auto !important;
            padding: 15px !important;
            z-index: 999 !important;
            max-width: 100% !important;
            width: 100% !important;
            box-sizing: border-box !important;
            overflow: hidden !important;
            background: transparent !important;
            border: none !important;
            border-radius: 8px !important;
        }
        
        /* Ensure results are properly contained within widget bounds */
        .tpls-results-container {
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
            clear: both !important;
            max-width: 100% !important;
            overflow: hidden !important;
            box-sizing: border-box !important;
        }
        
        .tpls-results-section {
            width: 100% !important;
            margin-bottom: 15px !important;
            max-width: 100% !important;
            overflow: hidden !important;
            box-sizing: border-box !important;
        }
        
        .tpls-result-item {
            width: 100% !important;
            margin: 0 0 10px 0 !important;
            box-sizing: border-box !important;
            max-width: 100% !important;
            overflow: hidden !important;
        }
        
        /* Prevent any content from extending beyond widget bounds */
        * {
            max-width: 100% !important;
            box-sizing: border-box !important;
        }
        
        /* Aggressive last-child hiding */
        body > *:last-child:not(.tpls-search-container) {
            display: none !important;
        }
        
        /* Hide anything that might break the layout */
        body > div:not(.tpls-search-container):not([class*="tpls"]) {
            display: none !important;
        }
        </style>
        """
        else:
            embed_styles = ""
            
        # Properly format the embed parameter for JavaScript
        embed_param_js = '"&embed=1"' if embed_mode else '""'

        # Build the complete HTML using simple string concatenation to avoid format issues
        html_output = embed_styles + """
    <style>
    /* TouchPoint Live Search Styles - Namespaced and constrained for widget */
    .tpls-search-container {
        max-width: 100%;
        width: 100%;
        margin: 0 auto;
        font-family: Arial, sans-serif;
        background: transparent !important;
        position: relative;
        z-index: 999;
        padding: 20px;
        box-sizing: border-box;
        clear: both;
        overflow: hidden;
        border: none;
        border-radius: 8px;
        box-shadow: none;
    }
    .tpls-search-box {
        position: relative;
        margin-bottom: 15px;
        width: 100%;
        max-width: 100%;
        box-sizing: border-box;
    }
    .tpls-search-input {
        width: 100%;
        max-width: 100%;
        padding: 12px 50px 12px 15px;
        font-size: 16px;
        border: 1px solid #ccc;
        border-radius: 4px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        background: white !important;
        box-sizing: border-box;
    }
    .tpls-search-button {
        position: absolute;
        right: 5px;
        top: 5px;
        padding: 8px 15px;
        background: #0B3D4C;
        color: white;
        border: none;
        border-radius: 4px;
        cursor: pointer;
    }
    .tpls-search-loading {
        display: none;
        position: absolute;
        right: 60px;
        top: 12px;
    }
    .tpls-results-container {
        margin-top: 15px;
        width: 100%;
        max-width: 100%;
        overflow: hidden;
        background: transparent !important;
        clear: both;
        box-sizing: border-box;
    }
    .tpls-results-section {
        margin-bottom: 20px;
        width: 100%;
        max-width: 100%;
        clear: both;
        box-sizing: border-box;
        overflow: hidden;
    }
    .tpls-section-heading {
        border-bottom: 2px solid #0B3D4C;
        padding-bottom: 8px;
        margin-bottom: 12px;
        color: #0B3D4C;
        width: 100%;
        max-width: 100%;
        font-size: 18px;
        font-weight: bold;
        box-sizing: border-box;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }
    .tpls-result-item {
        padding: 12px;
        margin-bottom: 12px;
        border: 1px solid #eee;
        border-radius: 4px;
        transition: all 0.2s;
        width: 100%;
        max-width: 100%;
        box-sizing: border-box;
        overflow: visible;  /* Changed from hidden to visible */
        background: rgba(255, 255, 255, 0.9) !important;
        position: relative;
        clear: both;
    }
    .tpls-result-item:hover {
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    .tpls-result-name {
        font-weight: bold;
        margin-bottom: 8px;
        font-size: 16px;
        position: relative;
        overflow: visible;
        width: 100%;
        display: block;  /* Changed from flex to block */
    }
    .tpls-result-name a {
        color: #0B3D4C;
        text-decoration: none;
        display: inline-block;
        line-height: 1.4;
        word-wrap: break-word;
        overflow-wrap: break-word;
        white-space: normal;
    }
    .tpls-result-name a:hover {
        text-decoration: underline;
    }
    .tpls-result-actions {
        display: flex;
        align-items: center;
        gap: 5px;
        flex-shrink: 0;
        margin-top: 10px;  /* Add space above buttons */
        flex-wrap: nowrap;
        min-width: 0;
        margin-left: 0;
    }
    /* Responsive layout for smaller screens */
    @media (max-width: 768px) {
        .tpls-result-name a {
            margin-bottom: 0;
            margin-right: 0;
        }
        .tpls-result-actions {
            justify-content: flex-start;
        }
    }
    /* Very small screens */
    @media (max-width: 320px) {
        .tpls-action-button {
            padding: 4px 8px;
            font-size: 10px;
        }
        .tpls-action-button i {
            font-size: 11px;
        }
    }
    /* Fix for TouchPoint box-content class conflicts */
    .box-content .tpls-result-item {
        overflow: visible !important;
    }
    .box-content .tpls-result-name {
        overflow: visible !important;
    }
    .box-content .tpls-result-actions {
        position: relative !important;
        overflow: visible !important;
        margin-left: 0 !important;
        flex-shrink: 0 !important;
    }
    .box-content .tpls-action-button {
        position: relative !important;
        z-index: 100 !important;
    }
    
    /* Handle all widths - buttons always below content */
    @media (max-width: 600px) {
        .tpls-result-actions {
            width: 100%;
            justify-content: flex-start;
        }
    }
    .tpls-result-meta {
        font-size: 14px;
        color: #666;
        margin-bottom: 10px;
        width: 100%;
        word-wrap: break-word;
    }
    .tpls-result-meta a {
        color: #0B3D4C;
        text-decoration: none;
    }
    .tpls-result-meta a:hover {
        text-decoration: underline;
    }
    .tpls-phone-link {
        color: #0B3D4C;
        text-decoration: none;
    }
    .tpls-phone-link:hover {
        text-decoration: underline;
    }
    .tpls-no-results {
        padding: 20px;
        text-align: center;
        color: #777;
        background: #f9f9f9;
        border-radius: 4px;
    }
    @keyframes tpls-spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
    .tpls-spinner {
        border: 3px solid #f3f3f3;
        border-top: 3px solid #0B3D4C;
        border-radius: 50%;
        width: 20px;
        height: 20px;
        animation: tpls-spin 1s linear infinite;
    }
    .tpls-action-button {
        display: inline-flex;
        align-items: center;
        margin: 0 2px;
        padding: 6px 10px;
        background: #0B3D4C !important;
        color: white !important;
        border: 1px solid #0B3D4C !important;
        border-radius: 4px;
        font-size: 11px;
        cursor: pointer;
        text-decoration: none !important;
        font-weight: normal;
        line-height: 1.2;
        vertical-align: middle;
        white-space: nowrap;
        position: relative;
        z-index: 10;
        transition: all 0.2s ease;
        flex-shrink: 0;
    }
    .tpls-action-button:hover {
        background: #0a2933 !important;
        color: white !important;
        text-decoration: none !important;
        border-color: #0a2933 !important;
        transform: translateY(-1px);
    }
    .tpls-action-button:focus {
        background: #0a2933 !important;
        color: white !important;
        text-decoration: none !important;
        outline: 2px solid #4a90a4;
    }
    .tpls-action-button i {
        margin-right: 3px;
        font-size: 12px;
    }
    /* Journey button special styling */
    .tpls-journey-btn {
        background: #28a745 !important;
        border-color: #28a745 !important;
    }
    .tpls-journey-btn:hover {
        background: #218838 !important;
        border-color: #1e7e34 !important;
    }
    /* Modal styles - namespaced */
    .tpls-modal {
        display: none;
        position: fixed;
        z-index: 99999;
        left: 0;
        top: 0;
        width: 100%;
        height: 100%;
        overflow: auto;
        background-color: rgba(0,0,0,0.5);
    }
    .tpls-modal-content {
        background-color: #fefefe;
        margin: 10% auto;
        padding: 20px;
        border: 1px solid #888;
        width: 80%;
        max-width: 600px;
        border-radius: 5px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    .tpls-modal-close {
        color: #aaa;
        float: right;
        font-size: 28px;
        font-weight: bold;
        cursor: pointer;
    }
    .tpls-modal-close:hover,
    .tpls-modal-close:focus {
        color: black;
        text-decoration: none;
    }
    .tpls-modal-title {
        margin-top: 0;
        color: #0B3D4C;
    }
    .tpls-form-group {
        margin-bottom: 15px;
    }
    .tpls-form-group label {
        display: block;
        margin-bottom: 5px;
        font-weight: bold;
    }
    .tpls-form-control {
        width: 100%;
        padding: 8px;
        border: 1px solid #ddd;
        border-radius: 4px;
        box-sizing: border-box;
    }
    .tpls-search-input-wrapper {
        position: relative;
        margin-bottom: 10px;
    }
    .tpls-search-results {
        position: absolute;
        top: 100%;
        left: 0;
        right: 0;
        background: white;
        border: 1px solid #ddd;
        border-radius: 0 0 4px 4px;
        max-height: 200px;
        overflow-y: auto;
        z-index: 10;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        display: none;
    }
    .tpls-search-result-item {
        padding: 8px 12px;
        cursor: pointer;
        border-bottom: 1px solid #eee;
    }
    .tpls-search-result-item:hover {
        background-color: #f5f5f5;
    }
    .tpls-search-result-item:last-child {
        border-bottom: none;
    }
    .tpls-search-clear {
        position: absolute;
        right: 60px;
        top: 12px;
        color: #999;
        font-size: 24px;
        cursor: pointer;
        display: none;
    }
    .tpls-search-clear:hover {
        color: #333;
    }
    .tpls-form-actions {
        text-align: right;
        margin-top: 20px;
    }
    .tpls-btn {
        padding: 8px 15px;
        border: none;
        border-radius: 4px;
        cursor: pointer;
    }
    .tpls-btn-primary {
        background-color: #0B3D4C;
        color: white;
    }
    .tpls-btn-secondary {
        background-color: #ccc;
        color: #333;
        margin-right: 10px;
    }
    .tpls-keyword-container {
        max-height: 150px;
        overflow-y: auto;
        border: 1px solid #ddd;
        border-radius: 4px;
        padding: 10px;
    }
    .tpls-keyword-item {
        margin-bottom: 5px;
    }
    .tpls-alert {
        padding: 15px;
        margin-bottom: 20px;
        border: 1px solid transparent;
        border-radius: 4px;
    }
    .tpls-alert-success {
        color: #3c763d;
        background-color: #dff0d8;
        border-color: #d6e9c6;
    }
    .tpls-alert-danger {
        color: #a94442;
        background-color: #f2dede;
        border-color: #ebccd1;
    }
    .tpls-match-highlight {
        background-color: #FFEB3B;
        padding: 0 3px;
        border-radius: 2px;
    }
    .tpls-loading-indicator {
        display: none;
        text-align: center;
        padding: 20px;
    }
    </style>
    
    <div class="tpls-search-container">
        <div class="tpls-search-box">
            <input type="text" id="tplsSearchInput" class="tpls-search-input" value=\"""" + search_term + """\" 
                   placeholder="Search by name, phone, or involvement...">
            <span id="tplsClearSearch" class="tpls-search-clear" style="display: none;">x</span>
            <div class="tpls-search-loading" id="tplsSearchLoading">
                <div class="tpls-spinner"></div>
            </div>
            <button type="button" id="tplsSearchButton" class="tpls-search-button">
                <i class="fa fa-search"></i>
            </button>
        </div>
        
        <div id="tplsLoadingIndicator" class="tpls-loading-indicator">
            <div class="tpls-spinner" style="display: inline-block;"></div>
            <p>Loading results...</p>
        </div>
        
        <div id="tplsResultsContainer" class="tpls-results-container">
            <!-- Results will be loaded here -->
        </div>
    </div>
    
    <!-- Task Modal -->
    <div id="tplsTaskModal" class="tpls-modal">
        <div class="tpls-modal-content">
            <span class="tpls-modal-close" id="tplsCloseTaskModal">&times;</span>
            <h3 class="tpls-modal-title">Add Task</h3>
            <form id="tplsTaskForm" method="post" action="/PyScriptForm/""" + script_name + """">
                <input type="hidden" name="action_type" value="task">
                <input type="hidden" id="tplsTaskAboutPersonId" name="about_person_id" value="">
                
                <div class="tpls-form-group">
                    <label for="tplsTaskAssigneeSearch">Assign To (Search by name):</label>
                    <div class="tpls-search-input-wrapper">
                        <input type="text" id="tplsTaskAssigneeSearch" class="tpls-form-control" placeholder="Search for a person...">
                        <div id="tplsAssigneeSearchResults" class="tpls-search-results"></div>
                        <input type="hidden" id="tplsTaskAssigneeId" name="assignee_id" value="">
                    </div>
                    <div id="tplsSelectedAssignee" style="margin-top: 5px; font-style: italic;"></div>
                </div>
                
                <div class="tpls-form-group">
                    <label for="tplsTaskDueDate">Due Date (optional):</label>
                    <input type="date" id="tplsTaskDueDate" name="due_date" class="tpls-form-control">
                </div>
                
                <div class="tpls-form-group">
                    <label>Keywords:</label>
                    <div id="tplsKeywordsLoadingIndicator" style="text-align: center;">
                        <div class="tpls-spinner" style="display: inline-block;"></div>
                        <p>Loading keywords...</p>
                    </div>
                    <div class="tpls-keyword-container" id="tplsTaskKeywords" style="display: none;">
                        <!-- Will be populated via JavaScript -->
                    </div>
                </div>
                
                <div class="tpls-form-group">
                    <label for="tplsTaskMessage">Message:</label>
                    <textarea id="tplsTaskMessage" name="message" rows="5" class="tpls-form-control"></textarea>
                </div>
                
                <div class="tpls-form-actions">
                    <button type="button" class="tpls-btn tpls-btn-secondary" id="tplsCancelTaskBtn">Cancel</button>
                    <button type="submit" class="tpls-btn tpls-btn-primary">Save Task</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- Note Modal -->
    <div id="tplsNoteModal" class="tpls-modal">
        <div class="tpls-modal-content">
            <span class="tpls-modal-close" id="tplsCloseNoteModal">&times;</span>
            <h3 class="tpls-modal-title">Add Note</h3>
            <form id="tplsNoteForm" method="post" action="/PyScriptForm/""" + script_name + """">
                <input type="hidden" name="action_type" value="note">
                <input type="hidden" id="tplsNoteAboutPersonId" name="about_person_id" value="">
                
                <div class="tpls-form-group">
                    <label>Keywords:</label>
                    <div id="tplsNoteKeywordsLoadingIndicator" style="text-align: center;">
                        <div class="tpls-spinner" style="display: inline-block;"></div>
                        <p>Loading keywords...</p>
                    </div>
                    <div class="tpls-keyword-container" id="tplsNoteKeywords" style="display: none;">
                        <!-- Will be populated via JavaScript -->
                    </div>
                </div>
                
                <div class="tpls-form-group">
                    <label for="tplsNoteMessage">Message:</label>
                    <textarea id="tplsNoteMessage" name="message" rows="5" class="tpls-form-control"></textarea>
                </div>
                
                <div class="tpls-form-actions">
                    <button type="button" class="tpls-btn tpls-btn-secondary" id="tplsCancelNoteBtn">Cancel</button>
                    <button type="submit" class="tpls-btn tpls-btn-primary">Save Note</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- Person Journey Modal -->
    <div id="personJourneyModal" style="display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; 
         background: rgba(0,0,0,0.5); z-index: 1000; overflow-y: auto;">
        <div style="position: relative; width: 90%; max-width: 800px; margin: 50px auto; 
             background: white; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.3);">
            <div style="padding: 20px; border-bottom: 1px solid #dee2e6; display: flex; justify-content: space-between; align-items: center;">
                <h4 style="margin: 0; color: #2c3e50;" id="journeyModalTitle">
                    <i class="fa fa-users"></i> Family Engagement Journey
                </h4>
                <button type="button" onclick="tplsClosePersonJourney()" 
                        style="background: none; border: none; font-size: 24px; cursor: pointer; color: #999;">
                    <span>&times;</span>
                </button>
            </div>
            <div id="personJourneyContent" style="max-height: 70vh; overflow-y: auto;">
                Loading...
            </div>
            <div style="padding: 15px 20px; border-top: 1px solid #dee2e6; text-align: right;">
                <button type="button" class="tpls-btn tpls-btn-secondary" onclick="tplsClosePersonJourney()" style="margin-right: 10px;">Close</button>
                <button type="button" class="tpls-btn tpls-btn-primary" onclick="tplsShowIndividualJourney()" id="familyJourneyBtn" style="margin-right: 10px;">
                    <i class="fa fa-user"></i> Individual Timeline
                </button>
                <button type="button" class="tpls-btn tpls-btn-primary" onclick="tplsViewFullProfile()" id="viewProfileBtn">
                    <i class="fa fa-user"></i> View Full Profile
                </button>
            </div>
        </div>
    </div>
    
    <script>
    // TouchPoint Live Search - Namespaced JavaScript to avoid conflicts
    (function() {
        'use strict';
        
        // ::START:: Configuration - Scoped to this function
        var TPLS_CONFIG = {
            SCRIPT_NAME: '""" + script_name + """',
            SEARCH_DELAY: """ + str(SEARCH_DELAY) + """,
            EMBED_PARAM: """ + embed_param_js + """
        };
        
        // Preloaded keywords from the server
        var tplsPreloadedKeywords = """ + keywords_json + """;
        
        // Fallback keywords in case the preloaded list is empty
        var tplsDefaultKeywords = [
            {KeywordId: 79, Description: "Note", Code: "N"},
            {KeywordId: 115, Description: "Email", Code: "E"},
            {KeywordId: 73, Description: "Phone", Code: "P"},
            {KeywordId: 82, Description: "Prayer", Code: "R"},
            {KeywordId: 126, Description: "Task", Code: "T"},
            {KeywordId: 99, Description: "PastorTouch", Code: "PT"},
            {KeywordId: 75, Description: "Visit", Code: "V"}
        ];
        
        // ::START:: Helper Functions
        function tplsEscapeHtml(unsafe) {
            return unsafe
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;")
                .replace(/'/g, "&#039;");
        }
        
        function tplsFormatPhoneForSearch(phone) {
            if (!phone) return '';
            return phone.replace(/\\D/g, '');
        }
        
        function tplsGetScriptUrl() {
            return '/PyScript/' + TPLS_CONFIG.SCRIPT_NAME;
        }
        
        function tplsGetPyScriptFormUrl() {
            return '/PyScriptForm/' + TPLS_CONFIG.SCRIPT_NAME;
        }
        
        // ::START:: Search Functions
        function tplsPerformSearch() {
            var searchInput = document.getElementById('tplsSearchInput');
            var resultsContainer = document.getElementById('tplsResultsContainer');
            var searchLoading = document.getElementById('tplsSearchLoading');
            var loadingIndicator = document.getElementById('tplsLoadingIndicator');
            
            if (!searchInput || !resultsContainer) {
                console.error('Search elements not found');
                return;
            }
            
            var searchTerm = searchInput.value.trim();
            
            // Don't search if term is too short
            if (searchTerm.length < 2) {
                resultsContainer.innerHTML = '';
                return;
            }
            
            // Show loading indicators
            if (searchLoading) searchLoading.style.display = 'block';
            if (loadingIndicator) loadingIndicator.style.display = 'block';
            resultsContainer.style.display = 'none';
            
            // Construct the proper URL for AJAX search
            var searchUrl = tplsGetScriptUrl() + '?q=' + encodeURIComponent(searchTerm) + '&ajax=1' + TPLS_CONFIG.EMBED_PARAM;
            
            // Make AJAX request
            var xhr = new XMLHttpRequest();
            xhr.open('GET', searchUrl, true);
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4) {
                    if (searchLoading) searchLoading.style.display = 'none';
                    if (loadingIndicator) loadingIndicator.style.display = 'none';
                    resultsContainer.style.display = 'block';
                    
                    if (xhr.status === 200) {
                        resultsContainer.innerHTML = xhr.responseText;
                        
                        // Debug: Check if buttons are in the response
                        console.log('Search response length:', xhr.responseText.length);
                        console.log('Response contains task buttons:', xhr.responseText.includes('tpls-add-task-btn'));
                        console.log('Response contains note buttons:', xhr.responseText.includes('tpls-add-note-btn'));
                        
                        // GENTLE post-search footer cleanup with safety checks
                        setTimeout(function() {
                            try {
                                console.log('Post-search gentle cleanup...');
                                
                                // Target specific footer elements including user-requested IDs
                                var specificFooterSelectors = [
                                    'div[id*="footer"][class*="hidden-print"]',
                                    '.footer.hidden-print',
                                    '#contact-footer',  // User requested
                                    '#footer'           // User requested
                                ];
                                
                                specificFooterSelectors.forEach(function(selector) {
                                    try {
                                        var elements = document.querySelectorAll(selector);
                                        if (elements && elements.length > 0) {
                                            elements.forEach(function(el) {
                                                if (el && el.parentNode && !el.closest('.tpls-search-container')) {
                                                    console.log('Gently removing post-search footer:', selector);
                                                    el.style.display = 'none';  // Hide instead of remove for safety
                                                }
                                            });
                                        }
                                    } catch (e) {
                                        // Silently ignore individual selector errors
                                    }
                                });
                                
                                // Also hide any divs with these specific IDs in the results container
                                var resultsFooters = resultsContainer.querySelectorAll('#contact-footer, #footer');
                                resultsFooters.forEach(function(el) {
                                    if (el) {
                                        el.style.display = 'none';
                                    }
                                });
                            } catch (e) {
                                console.error('Post-search cleanup error:', e);
                            }
                        }, 100);
                        
                        // Set up action buttons for new results
                        setTimeout(function() {
                            try {
                                var taskButtons = document.querySelectorAll('.tpls-add-task-btn');
                                var noteButtons = document.querySelectorAll('.tpls-add-note-btn');
                                console.log('Found task buttons after search:', taskButtons.length);
                                console.log('Found note buttons after search:', noteButtons.length);
                                
                                tplsSetupActionButtons();
                                
                                // Trigger resize after content is loaded
                                if (typeof tplsNotifyParentOfHeight === 'function') {
                                    tplsNotifyParentOfHeight();
                                }
                            } catch (e) {
                                console.error('Button setup error:', e);
                            }
                        }, 150);
                    } else {
                        resultsContainer.innerHTML = '<div class="tpls-no-results">An error occurred. Please try again.</div>';
                    }
                }
            };
            xhr.send();
        }
        
        function tplsSearchPeople(term) {
            var resultsContainer = document.getElementById('tplsAssigneeSearchResults');
            if (!resultsContainer) return;
            
            if (term.length < 2) {
                resultsContainer.innerHTML = '';
                resultsContainer.style.display = 'none';
                return;
            }
            
            resultsContainer.innerHTML = '<div style="padding: 10px; text-align: center;"><div class="tpls-spinner" style="display: inline-block;"></div> Searching...</div>';
            resultsContainer.style.display = 'block';
            
            var peopleSearchUrl = tplsGetScriptUrl() + '?q=' + encodeURIComponent(term) + '&ajax=1&people_search=1';
            
            var xhr = new XMLHttpRequest();
            xhr.open('GET', peopleSearchUrl, true);
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4) {
                    if (xhr.status === 200) {
                        try {
                            var html = '';
                            var tempContainer = document.createElement('div');
                            tempContainer.innerHTML = xhr.responseText;
                            
                            var peopleItems = tempContainer.querySelectorAll('.tpls-result-item');
                            var peopleTally = 0;
                            
                            peopleItems.forEach(function(item) {
                                var nameLink = item.querySelector('.tpls-result-name a');
                                if (nameLink) {
                                    var href = nameLink.getAttribute('href');
                                    var idMatch = href.match(/\\/Person2\\/(\\d+)/);
                                    
                                    if (idMatch && idMatch[1]) {
                                        var personId = idMatch[1];
                                        var personName = nameLink.textContent.trim();
                                        
                                        html += '<div class="tpls-search-result-item" data-id="' + personId + '" data-name="' + tplsEscapeHtml(personName) + '">';
                                        html += tplsEscapeHtml(personName);
                                        html += '</div>';
                                        peopleTally++;
                                    }
                                }
                            });
                            
                            if (peopleTally === 0) {
                                html = '<div class="tpls-search-result-item">No results found</div>';
                            }
                            
                            resultsContainer.innerHTML = html;
                            
                            // Add click handlers to results
                            var resultItems = resultsContainer.querySelectorAll('.tpls-search-result-item');
                            resultItems.forEach(function(item) {
                                item.addEventListener('click', function() {
                                    var id = this.getAttribute('data-id');
                                    var name = this.getAttribute('data-name');
                                    
                                    if (id && name) {
                                        var assigneeIdField = document.getElementById('tplsTaskAssigneeId');
                                        var selectedAssigneeDiv = document.getElementById('tplsSelectedAssignee');
                                        var searchField = document.getElementById('tplsTaskAssigneeSearch');
                                        
                                        if (assigneeIdField) assigneeIdField.value = id;
                                        if (selectedAssigneeDiv) selectedAssigneeDiv.textContent = 'Selected: ' + name;
                                        if (searchField) searchField.value = '';
                                        resultsContainer.style.display = 'none';
                                    }
                                });
                            });
                        } catch (e) {
                            console.error('Error processing people search results:', e);
                            resultsContainer.innerHTML = '<div class="tpls-search-result-item">Error processing results</div>';
                        }
                    } else {
                        resultsContainer.innerHTML = '<div class="tpls-search-result-item">Error loading results</div>';
                    }
                }
            };
            xhr.send();
        }
        
        // ::START:: Keyword Functions
        function tplsLoadKeywords() {
            var keywordsLoadingIndicator = document.getElementById('tplsKeywordsLoadingIndicator');
            var noteKeywordsLoadingIndicator = document.getElementById('tplsNoteKeywordsLoadingIndicator');
            var taskKeywords = document.getElementById('tplsTaskKeywords');
            var noteKeywords = document.getElementById('tplsNoteKeywords');
            
            if (keywordsLoadingIndicator) keywordsLoadingIndicator.style.display = 'block';
            if (noteKeywordsLoadingIndicator) noteKeywordsLoadingIndicator.style.display = 'block';
            if (taskKeywords) taskKeywords.style.display = 'none';
            if (noteKeywords) noteKeywords.style.display = 'none';
            
            var keywordsToUse = [];
            try {
                if (tplsPreloadedKeywords && tplsPreloadedKeywords.length > 0) {
                    keywordsToUse = tplsPreloadedKeywords;
                } else {
                    keywordsToUse = tplsDefaultKeywords;
                }
            } catch (e) {
                keywordsToUse = tplsDefaultKeywords;
                console.error("Error processing keywords, using defaults:", e);
            }
            
            tplsPopulateKeywords(keywordsToUse);
            
            if (keywordsLoadingIndicator) keywordsLoadingIndicator.style.display = 'none';
            if (noteKeywordsLoadingIndicator) noteKeywordsLoadingIndicator.style.display = 'none';
            if (taskKeywords) taskKeywords.style.display = 'block';
            if (noteKeywords) noteKeywords.style.display = 'block';
        }
        
        function tplsPopulateKeywords(keywords) {
            var taskKeywordsContainer = document.getElementById('tplsTaskKeywords');
            var noteKeywordsContainer = document.getElementById('tplsNoteKeywords');
            
            if (!taskKeywordsContainer || !noteKeywordsContainer) return;
            
            keywords.sort(function(a, b) {
                return a.Description.localeCompare(b.Description);
            });
            
            var taskKeywordsHtml = '';
            var noteKeywordsHtml = '';
            
            keywords.forEach(function(keyword, index) {
                var keywordHtml = '<div class="tpls-keyword-item">';
                keywordHtml += '<label>';
                keywordHtml += '<input type="checkbox" name="keyword_' + index + '" value="' + keyword.KeywordId + '"> ';
                keywordHtml += tplsEscapeHtml(keyword.Description);
                if (keyword.Code) {
                    keywordHtml += ' (' + tplsEscapeHtml(keyword.Code) + ')';
                }
                keywordHtml += '</label>';
                keywordHtml += '</div>';
                
                taskKeywordsHtml += keywordHtml;
                noteKeywordsHtml += keywordHtml;
            });
            
            taskKeywordsContainer.innerHTML = taskKeywordsHtml || '<p>No keywords available</p>';
            noteKeywordsContainer.innerHTML = noteKeywordsHtml || '<p>No keywords available</p>';
        }
        
        // ::START:: Modal Functions
        function tplsSetupModals() {
            var taskModal = document.getElementById('tplsTaskModal');
            var closeTaskModal = document.getElementById('tplsCloseTaskModal');
            var cancelTaskBtn = document.getElementById('tplsCancelTaskBtn');
            
            var noteModal = document.getElementById('tplsNoteModal');
            var closeNoteModal = document.getElementById('tplsCloseNoteModal');
            var cancelNoteBtn = document.getElementById('tplsCancelNoteBtn');
            
            if (closeTaskModal) {
                closeTaskModal.onclick = function() {
                    if (taskModal) taskModal.style.display = "none";
                };
            }
            
            if (cancelTaskBtn) {
                cancelTaskBtn.onclick = function() {
                    if (taskModal) taskModal.style.display = "none";
                };
            }
            
            if (closeNoteModal) {
                closeNoteModal.onclick = function() {
                    if (noteModal) noteModal.style.display = "none";
                };
            }
            
            if (cancelNoteBtn) {
                cancelNoteBtn.onclick = function() {
                    if (noteModal) noteModal.style.display = "none";
                };
            }
            
            // Close modal when clicking outside
            window.onclick = function(event) {
                if (event.target === taskModal && taskModal) {
                    taskModal.style.display = "none";
                }
                if (event.target === noteModal && noteModal) {
                    noteModal.style.display = "none";
                }
            };
            
            // Setup assignee search
            var assigneeSearchInput = document.getElementById('tplsTaskAssigneeSearch');
            var assigneeSearchTimeout = null;
            
            if (assigneeSearchInput) {
                assigneeSearchInput.addEventListener('input', function() {
                    clearTimeout(assigneeSearchTimeout);
                    var term = this.value.trim();
                    
                    assigneeSearchTimeout = setTimeout(function() {
                        tplsSearchPeople(term);
                    }, 300);
                });
            }
            
            // Close search results when clicking outside
            document.addEventListener('click', function(e) {
                var searchWrapper = e.target.closest('.tpls-search-input-wrapper');
                if (!searchWrapper) {
                    var assigneeResults = document.getElementById('tplsAssigneeSearchResults');
                    if (assigneeResults) assigneeResults.style.display = 'none';
                }
            });
        }
        
        function tplsSetupActionButtons() {
            // Set up task buttons
            var taskButtons = document.querySelectorAll('.tpls-add-task-btn');
            console.log('Found', taskButtons.length, 'task buttons'); // Debug
            
            taskButtons.forEach(function(btn) {
                // Remove any existing event listeners to prevent duplicates
                btn.replaceWith(btn.cloneNode(true));
            });
            
            // Re-query after cloning
            taskButtons = document.querySelectorAll('.tpls-add-task-btn');
            taskButtons.forEach(function(btn) {
                btn.addEventListener('click', function(e) {
                    e.preventDefault();
                    var personId = this.getAttribute('data-person-id');
                    var personName = this.getAttribute('data-person-name');
                    
                    console.log('Task button clicked for:', personName); // Debug
                    
                    var taskModal = document.getElementById('tplsTaskModal');
                    var taskAboutPersonId = document.getElementById('tplsTaskAboutPersonId');
                    var modalTitle = document.querySelector('#tplsTaskModal .tpls-modal-title');
                    
                    if (taskAboutPersonId) taskAboutPersonId.value = personId;
                    if (taskModal) taskModal.style.display = 'block';
                    if (modalTitle) modalTitle.textContent = 'Add Task for ' + personName;
                    
                    // Clear previous values
                    var assigneeId = document.getElementById('tplsTaskAssigneeId');
                    var selectedAssignee = document.getElementById('tplsSelectedAssignee');
                    var dueDate = document.getElementById('tplsTaskDueDate');
                    var message = document.getElementById('tplsTaskMessage');
                    
                    if (assigneeId) assigneeId.value = '';
                    if (selectedAssignee) selectedAssignee.textContent = '';
                    if (dueDate) dueDate.value = '';
                    if (message) message.value = '';
                    
                    // Uncheck all keywords
                    var checkboxes = document.querySelectorAll('#tplsTaskKeywords input[type="checkbox"]');
                    checkboxes.forEach(function(checkbox) {
                        checkbox.checked = false;
                    });
                    
                    // Trigger resize for iframe
                    if (typeof tplsNotifyParentOfHeight === 'function') {
                        setTimeout(tplsNotifyParentOfHeight, 100);
                    }
                });
            });
            
            // Set up note buttons
            var noteButtons = document.querySelectorAll('.tpls-add-note-btn');
            console.log('Found', noteButtons.length, 'note buttons'); // Debug
            
            noteButtons.forEach(function(btn) {
                // Remove any existing event listeners to prevent duplicates
                btn.replaceWith(btn.cloneNode(true));
            });
            
            // Re-query after cloning
            noteButtons = document.querySelectorAll('.tpls-add-note-btn');
            noteButtons.forEach(function(btn) {
                btn.addEventListener('click', function(e) {
                    e.preventDefault();
                    var personId = this.getAttribute('data-person-id');
                    var personName = this.getAttribute('data-person-name');
                    
                    console.log('Note button clicked for:', personName); // Debug
                    
                    var noteModal = document.getElementById('tplsNoteModal');
                    var noteAboutPersonId = document.getElementById('tplsNoteAboutPersonId');
                    var modalTitle = document.querySelector('#tplsNoteModal .tpls-modal-title');
                    
                    if (noteAboutPersonId) noteAboutPersonId.value = personId;
                    if (noteModal) noteModal.style.display = 'block';
                    if (modalTitle) modalTitle.textContent = 'Add Note for ' + personName;
                    
                    // Clear previous values
                    var message = document.getElementById('tplsNoteMessage');
                    if (message) message.value = '';
                    
                    // Uncheck all keywords
                    var checkboxes = document.querySelectorAll('#tplsNoteKeywords input[type="checkbox"]');
                    checkboxes.forEach(function(checkbox) {
                        checkbox.checked = false;
                    });
                    
                    // Trigger resize for iframe
                    if (typeof tplsNotifyParentOfHeight === 'function') {
                        setTimeout(tplsNotifyParentOfHeight, 100);
                    }
                });
            });
        }
        
        // ::START:: Form Submission
        function tplsSetupFormSubmission() {
            var taskForm = document.getElementById('tplsTaskForm');
            var noteForm = document.getElementById('tplsNoteForm');
            
            if (taskForm) {
                taskForm.action = tplsGetPyScriptFormUrl();
                
                taskForm.addEventListener('submit', function(e) {
                    e.preventDefault();
                    
                    var form = this;
                    var formData = new FormData(form);
                    var submitButton = form.querySelector('button[type="submit"]');
                    
                    if (submitButton) {
                        submitButton.disabled = true;
                        submitButton.innerHTML = '<div class="tpls-spinner" style="display: inline-block; width: 15px; height: 15px;"></div> Saving...';
                    }
                    
                    var xhr = new XMLHttpRequest();
                    xhr.open('POST', form.action, true);
                    xhr.onreadystatechange = function() {
                        if (xhr.readyState === 4) {
                            if (submitButton) {
                                submitButton.disabled = false;
                                submitButton.innerHTML = 'Save Task';
                            }
                            
                            if (xhr.status === 200) {
                                var taskModal = document.getElementById('tplsTaskModal');
                                if (taskModal) taskModal.style.display = 'none';
                                
                                var resultsContainer = document.getElementById('tplsResultsContainer');
                                if (resultsContainer) {
                                    var successDiv = document.createElement('div');
                                    successDiv.className = 'tpls-alert tpls-alert-success';
                                    successDiv.innerHTML = 'Task created successfully!';
                                    resultsContainer.insertBefore(successDiv, resultsContainer.firstChild);
                                    
                                    setTimeout(function() {
                                        if (successDiv.parentNode) {
                                            successDiv.parentNode.removeChild(successDiv);
                                        }
                                    }, 3000);
                                }
                            } else {
                                alert('Error saving task. Please try again.');
                            }
                        }
                    };
                    xhr.send(formData);
                });
            }
            
            if (noteForm) {
                noteForm.action = tplsGetPyScriptFormUrl();
                
                noteForm.addEventListener('submit', function(e) {
                    e.preventDefault();
                    
                    var form = this;
                    var formData = new FormData(form);
                    var submitButton = form.querySelector('button[type="submit"]');
                    
                    if (submitButton) {
                        submitButton.disabled = true;
                        submitButton.innerHTML = '<div class="tpls-spinner" style="display: inline-block; width: 15px; height: 15px;"></div> Saving...';
                    }
                    
                    var xhr = new XMLHttpRequest();
                    xhr.open('POST', form.action, true);
                    xhr.onreadystatechange = function() {
                        if (xhr.readyState === 4) {
                            if (submitButton) {
                                submitButton.disabled = false;
                                submitButton.innerHTML = 'Save Note';
                            }
                            
                            if (xhr.status === 200) {
                                var noteModal = document.getElementById('tplsNoteModal');
                                if (noteModal) noteModal.style.display = 'none';
                                
                                var resultsContainer = document.getElementById('tplsResultsContainer');
                                if (resultsContainer) {
                                    var successDiv = document.createElement('div');
                                    successDiv.className = 'tpls-alert tpls-alert-success';
                                    successDiv.innerHTML = 'Note added successfully!';
                                    resultsContainer.insertBefore(successDiv, resultsContainer.firstChild);
                                    
                                    setTimeout(function() {
                                        if (successDiv.parentNode) {
                                            successDiv.parentNode.removeChild(successDiv);
                                        }
                                    }, 3000);
                                }
                            } else {
                                alert('Error saving note. Please try again.');
                            }
                        }
                    };
                    xhr.send(formData);
                });
            }
        }
        
        // ::START:: Person Journey Functions
        var tplsCurrentPersonId = null;
        
        function tplsShowPersonJourney(peopleId) {
            console.log('tplsShowPersonJourney called with ID:', peopleId);
            tplsCurrentPersonId = peopleId;
            
            var modal = document.getElementById('personJourneyModal');
            if (!modal) {
                console.error('Modal not found!');
                alert('Modal not found. Please check the page.');
                return;
            }
            
            modal.style.display = 'block';
            console.log('Modal should be visible now');
            
            // Update modal header to individual view
            document.getElementById('journeyModalTitle').innerHTML = 
                '<i class="fa fa-map-o"></i> Individual Engagement Journey';
            
            // Update button to show family option
            document.getElementById('familyJourneyBtn').innerHTML = 
                '<i class="fa fa-users"></i> Family Timeline';
            document.getElementById('familyJourneyBtn').onclick = tplsShowFamilyJourney;
            
            document.getElementById('personJourneyContent').innerHTML = '<div style="text-align: center; padding: 40px;"><i class="fa fa-spinner fa-spin fa-2x"></i><br><br>Loading individual engagement journey...</div>';
            
            // Load individual timeline by default
            setTimeout(function() {
                tplsLoadPersonJourney(peopleId);
            }, 100);
        }
        
        function tplsLoadPersonJourney(peopleId) {
            // Make AJAX call to get real journey data
            var xhr = new XMLHttpRequest();
            xhr.open('POST', tplsGetPyScriptFormUrl(), true);
            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
            
            // Add timeout for the request
            xhr.timeout = 10000; // 10 seconds
            
            xhr.ontimeout = function() {
                console.error('Request timed out');
                document.getElementById('personJourneyContent').innerHTML = 
                    '<div class="tpls-alert tpls-alert-danger"><h4>Request Timed Out</h4><p>The request took too long. Please try again.</p></div>';
            };
            
            xhr.onerror = function() {
                console.error('Network error');
                document.getElementById('personJourneyContent').innerHTML = 
                    '<div class="tpls-alert tpls-alert-danger"><h4>Network Error</h4><p>Unable to connect to the server. Please check your connection.</p></div>';
            };
            
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4) {
                    if (xhr.status === 200) {
                        try {
                            var data = JSON.parse(xhr.responseText);
                            if (data.error) {
                                document.getElementById('personJourneyContent').innerHTML = 
                                    '<div class="tpls-alert tpls-alert-danger"><h4>Error</h4><p>' + data.error + '</p></div>';
                                return;
                            }
                            
                            var journeyHtml = tplsGenerateJourneyTimeline(data);
                            document.getElementById('personJourneyContent').innerHTML = journeyHtml;
                        } catch (e) {
                            console.error('Error parsing response:', e);
                            console.error('Response text:', xhr.responseText);
                            document.getElementById('personJourneyContent').innerHTML = 
                                '<div class="tpls-alert tpls-alert-danger"><h4>Error</h4><p>Failed to parse journey data. The server response was invalid.</p></div>';
                        }
                    } else {
                        console.error('AJAX error:', xhr.status);
                        var errorMsg = 'Failed to load journey data (Error ' + xhr.status + ')';
                        if (xhr.status === 502) {
                            errorMsg = 'Server error. Please ensure the script name is correct and try again.';
                        } else if (xhr.status === 404) {
                            errorMsg = 'Script not found. Please check the configuration.';
                        }
                        document.getElementById('personJourneyContent').innerHTML = 
                            '<div class="tpls-alert tpls-alert-danger"><h4>Error</h4><p>' + errorMsg + '</p></div>';
                    }
                }
            };
            
            xhr.send('action=get_person_journey&peopleId=' + peopleId);
        }
        
        function tplsGenerateJourneyTimeline(data) {
            // data now contains real information from TouchPoint
            var person = data.person;
            var journey = data.journey;
            var insights = data.insights;
            var recentActivity = data.recent_activity || {};
            
            var html = '<div style="padding: 20px;">';
            
            // Person header
            html += '<div style="text-align: center; margin-bottom: 30px; padding: 20px; background: #f8f9fa; border-radius: 8px;">';
            html += '<h4 style="margin: 0; color: #2c3e50;">Engagement Journey for ' + person.name + '</h4>';
            html += '<p style="margin: 5px 0 0 0; color: #6c757d;">Age: ' + (person.age || 'Unknown') + ' | Status: ' + (person.member_status || 'Unknown') + '</p>';
            
            // Show recent activity status
            if (recentActivity.engagement_status) {
                var statusColor = recentActivity.engagement_status === 'Active' ? '#28a745' : 
                                 recentActivity.engagement_status === 'Inactive' ? '#dc3545' : '#ffc107';
                html += '<p style="margin: 5px 0 0 0;"><span style="background: ' + statusColor + '; color: white; padding: 3px 8px; border-radius: 12px; font-size: 12px;">' + recentActivity.engagement_status + '</span>';
                if (recentActivity.last_attendance) {
                    html += ' <small style="color: #6c757d;">Last seen: ' + recentActivity.last_attendance + '</small>';
                }
                html += '</p>';
            }
            
            html += '</div>';
            
            // Recent Activity Stats and Journey insights
            html += '<div class="row" style="margin-bottom: 30px;">';
            
            // Recent Activity Card
            html += '<div class="col-md-6">';
            html += '<div class="panel panel-default" style="margin-bottom: 20px;">';
            html += '<div class="panel-heading" style="background: #f8f9fa; border-bottom: 2px solid #007bff;">';
            html += '<h5 style="margin: 0; color: #0056b3;"><i class="fa fa-clock-o"></i> Recent Activity (90 days)</h5>';
            html += '</div>';
            html += '<div class="panel-body" style="padding: 15px;">';
            
            // Attendance in last 90 days
            if (recentActivity.attendance_90_days !== undefined) {
                html += '<div style="margin-bottom: 10px;">';
                html += '<strong>Attendance:</strong> ' + recentActivity.attendance_90_days + ' times';
                if (recentActivity.attendance_90_days >= 9) {
                    html += ' <span style="color: #28a745;">(Weekly)</span>';
                } else if (recentActivity.attendance_90_days >= 6) {
                    html += ' <span style="color: #28a745;">(Bi-weekly)</span>';
                } else if (recentActivity.attendance_90_days >= 3) {
                    html += ' <span style="color: #ffc107;">(Monthly)</span>';
                } else if (recentActivity.attendance_90_days > 0) {
                    html += ' <span style="color: #dc3545;">(Occasional)</span>';
                }
                html += '</div>';
            }
            
            // Last signup
            if (recentActivity.last_signup) {
                html += '<div style="margin-bottom: 10px;">';
                html += '<strong>Last Signup:</strong> ' + recentActivity.last_signup;
                html += '</div>';
            }
            
            // Currently serving
            if (recentActivity.currently_serving !== undefined && recentActivity.currently_serving > 0) {
                html += '<div style="margin-bottom: 10px;">';
                html += '<strong>Currently Serving:</strong> ' + recentActivity.currently_serving + ' role(s)';
                html += '</div>';
            }
            
            // Days since last attendance
            if (recentActivity.days_since_attendance !== undefined && recentActivity.days_since_attendance !== null) {
                html += '<div>';
                html += '<strong>Days Since Last Visit:</strong> ' + recentActivity.days_since_attendance;
                if (recentActivity.days_since_attendance > 90) {
                    html += ' <span style="color: #dc3545;">(Follow up needed)</span>';
                }
                html += '</div>';
            }
            
            html += '</div>';
            html += '</div>';
            html += '</div>';
            
            // Journey Summary Card
            html += '<div class="col-md-6">';
            html += '<div class="panel panel-default" style="margin-bottom: 20px;">';
            html += '<div class="panel-heading" style="background: #f8f9fa; border-bottom: 2px solid #28a745;">';
            html += '<h5 style="margin: 0; color: #28a745;"><i class="fa fa-line-chart"></i> Engagement Summary</h5>';
            html += '</div>';
            html += '<div class="panel-body" style="padding: 15px;">';
            
            html += '<div style="margin-bottom: 10px;">';
            html += '<strong>Current Score:</strong> <span style="font-size: 24px; font-weight: bold; color: #2c3e50;">' + person.engagement_score + '</span> / 100';
            html += '<div style="background: #e9ecef; height: 10px; border-radius: 5px; margin-top: 5px;">';
            html += '<div style="background: ' + (person.engagement_score >= 60 ? '#28a745' : person.engagement_score >= 40 ? '#ffc107' : '#dc3545') + '; height: 100%; width: ' + person.engagement_score + '%; border-radius: 5px;"></div>';
            html += '</div>';
            html += '<small style="color: #6c757d;">' + tplsGetEngagementLevel(person.engagement_score) + '</small>';
            html += '</div>';
            
            html += '<div style="margin-bottom: 10px;">';
            html += '<strong>Journey Length:</strong> ' + insights.journey_length;
            html += '</div>';
            
            html += '<div>';
            html += '<strong>Entry Point:</strong> ' + insights.entry_point;
            html += '</div>';
            
            html += '</div>';
            html += '</div>';
            html += '</div>';
            
            html += '</div>';
            
            // Timeline
            html += '<div style="position: relative;">';
            
            // Timeline line
            html += '<div style="position: absolute; left: 30px; top: 0; bottom: 0; width: 2px; background: #dee2e6;"></div>';
            
            if (journey.length === 0) {
                html += '<div style="padding-left: 70px; text-align: center;">';
                html += '<p style="color: #6c757d; font-style: italic;">No journey events found for this person.</p>';
                html += '</div>';
            } else {
                journey.forEach(function(item, index) {
                html += '<div style="position: relative; margin-bottom: 25px; padding-left: 70px;">';
                
                // Timeline dot
                html += '<div style="position: absolute; left: 20px; top: 5px; width: 20px; height: 20px; ';
                html += 'border-radius: 50%; background: ' + item.color + '; display: flex; align-items: center; justify-content: center;">';
                html += '<i class="fa ' + item.icon + '" style="color: white; font-size: 10px;"></i>';
                html += '</div>';
                
                // Event content
                html += '<div style="background: white; border: 1px solid #dee2e6; border-radius: 8px; padding: 15px; ';
                html += 'box-shadow: 0 1px 3px rgba(0,0,0,0.1);">';
                html += '<div style="display: flex; justify-content: between; align-items: start;">';
                html += '<div style="flex: 1;">';
                html += '<h6 style="margin: 0 0 5px 0; color: ' + item.color + '; font-weight: bold;">' + item.event + '</h6>';
                html += '<p style="margin: 0 0 5px 0; color: #2c3e50;">' + item.description + '</p>';
                html += '<small style="color: #6c757d;"><i class="fa fa-calendar-o"></i> ' + item.date + '</small>';
                html += '</div>';
                html += '<div style="margin-left: 15px;">';
                html += '<span class="badge" style="background: ' + item.color + '; color: white; padding: 4px 8px; font-size: 11px;">';
                html += item.type.toUpperCase();
                html += '</span>';
                html += '</div>';
                html += '</div>';
                html += '</div>';
                html += '</div>';
                });
            }
            
            html += '</div>';
            
            // Journey insights at bottom
            html += '<div style="margin-top: 30px; padding: 20px; background: #e6f7ff; border-radius: 8px; border-left: 4px solid #007bff;">';
            html += '<h5 style="color: #0056b3; margin: 0 0 10px 0;"><i class="fa fa-lightbulb-o"></i> Journey Insights</h5>';
            html += '<ul style="margin: 0; padding-left: 20px; color: #2c3e50;">';
            
            if (journey.length > 0) {
                html += '<li><strong>Total Events:</strong> ' + insights.total_events + ' milestone events recorded</li>';
                html += '<li><strong>Journey Duration:</strong> ' + insights.journey_length + ' from first contact to latest activity</li>';
                html += '<li><strong>Entry Point:</strong> ' + insights.entry_point + '</li>';
                html += '<li><strong>Current Status:</strong> ' + tplsGetEngagementLevel(person.engagement_score) + ' (score: ' + person.engagement_score + ')</li>';
                
                // Check for different event types
                var hasAttendance = journey.some(function(e) { return e.type === 'attendance'; });
                var hasServing = journey.some(function(e) { return e.type === 'serving'; });
                var hasProgram = journey.some(function(e) { return e.type === 'program'; });
                var hasSmallGroup = journey.some(function(e) { return e.type === 'smallgroup'; });
                
                if (hasAttendance && hasServing) {
                    html += '<li><strong>Active Participant:</strong> Has progressed from attendance to serving</li>';
                }
                if (hasProgram) {
                    html += '<li><strong>Program Engagement:</strong> Actively involved in organized programs</li>';
                }
                if (hasSmallGroup) {
                    html += '<li><strong>Small Group:</strong> Connected in small group community</li>';
                }
                
                // Add recent activity insights
                if (recentActivity.attendance_90_days >= 9) {
                    html += '<li><strong>Consistent Attendance:</strong> Attending weekly shows strong commitment</li>';
                } else if (recentActivity.days_since_attendance > 90) {
                    html += '<li><strong>Re-engagement Needed:</strong> Has not attended in over 90 days</li>';
                }
            } else {
                html += '<li><strong>No Events:</strong> This person has no recorded engagement events yet</li>';
                html += '<li><strong>Opportunity:</strong> Consider reaching out to encourage initial engagement</li>';
            }
            
            html += '</ul>';
            html += '</div>';
            
            html += '</div>';
            
            return html;
        }
        
        function tplsGetEngagementLevel(score) {
            if (score >= 80) return 'Highly engaged';
            if (score >= 60) return 'Engaged';
            if (score >= 40) return 'Moderately engaged';
            if (score >= 20) return 'Low engagement';
            return 'Not engaged';
        }
        
        function tplsClosePersonJourney() {
            console.log('tplsClosePersonJourney called');
            var modal = document.getElementById('personJourneyModal');
            if (modal) {
                modal.style.display = 'none';
                console.log('Modal closed');
            }
            tplsCurrentPersonId = null;
        }
        
        function tplsViewFullProfile() {
            if (tplsCurrentPersonId) {
                window.open('/Person2/' + tplsCurrentPersonId, '_blank');
            }
        }
        
        // Individual Journey Function
        function tplsShowIndividualJourney() {
            if (!tplsCurrentPersonId) {
                alert('No person selected');
                return;
            }
            
            // Update modal header to indicate individual view
            document.getElementById('journeyModalTitle').innerHTML = 
                '<i class="fa fa-map-o"></i> Individual Engagement Journey';
            
            // Update button state
            document.getElementById('familyJourneyBtn').innerHTML = 
                '<i class="fa fa-users"></i> Family Timeline';
            document.getElementById('familyJourneyBtn').onclick = tplsShowFamilyJourney;
            
            // Show loading state
            document.getElementById('personJourneyContent').innerHTML = 
                '<div style="text-align: center; padding: 40px;"><i class="fa fa-spinner fa-spin fa-2x"></i><br><br>Loading individual engagement journey...</div>';
            
            // Load individual journey data
            tplsLoadPersonJourney(tplsCurrentPersonId);
        }
        
        // Family Journey Functions
        function tplsShowFamilyJourney() {
            if (!tplsCurrentPersonId) {
                alert('No person selected');
                return;
            }
            
            // Update modal header to indicate family view
            document.getElementById('journeyModalTitle').innerHTML = 
                '<i class="fa fa-users"></i> Family Engagement Journey';
            
            // Update button state
            document.getElementById('familyJourneyBtn').innerHTML = 
                '<i class="fa fa-user"></i> Individual Timeline';
            document.getElementById('familyJourneyBtn').onclick = tplsShowIndividualJourney;
            
            // Show loading state
            document.getElementById('personJourneyContent').innerHTML = 
                '<div style="text-align: center; padding: 40px;"><i class="fa fa-spinner fa-spin fa-2x"></i><br><br>Loading family engagement journey...</div>';
            
            // Load family journey data
            var xhr = new XMLHttpRequest();
            xhr.open('POST', tplsGetPyScriptFormUrl(), true);
            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
            
            // Add timeout for the request
            xhr.timeout = 10000; // 10 seconds
            
            xhr.ontimeout = function() {
                console.error('Family request timed out');
                document.getElementById('personJourneyContent').innerHTML = 
                    '<div class="tpls-alert tpls-alert-danger"><h4>Request Timed Out</h4><p>The request took too long. Please try again.</p></div>';
            };
            
            xhr.onerror = function() {
                console.error('Network error loading family');
                document.getElementById('personJourneyContent').innerHTML = 
                    '<div class="tpls-alert tpls-alert-danger"><h4>Network Error</h4><p>Unable to connect to the server. Please check your connection.</p></div>';
            };
            
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4) {
                    if (xhr.status === 200) {
                        try {
                            var data = JSON.parse(xhr.responseText);
                            if (data.error) {
                                document.getElementById('personJourneyContent').innerHTML = 
                                    '<div class="tpls-alert tpls-alert-danger"><h4>Error</h4><p>' + data.error + '</p></div>';
                                return;
                            }
                            
                            var familyHtml = tplsGenerateFamilyTimeline(data);
                            document.getElementById('personJourneyContent').innerHTML = familyHtml;
                            
                        } catch (e) {
                            console.error('Error parsing family response:', e);
                            console.error('Family response text:', xhr.responseText);
                            document.getElementById('personJourneyContent').innerHTML = 
                                '<div class="tpls-alert tpls-alert-danger"><h4>Error</h4><p>Failed to parse family journey data. The server response was invalid.</p></div>';
                        }
                    } else {
                        console.error('Family AJAX error:', xhr.status);
                        var errorMsg = 'Failed to load family journey data (Error ' + xhr.status + ')';
                        if (xhr.status === 502) {
                            errorMsg = 'Server error. Please ensure the script name is correct and try again.';
                        } else if (xhr.status === 404) {
                            errorMsg = 'Script not found. Please check the configuration.';
                        }
                        document.getElementById('personJourneyContent').innerHTML = 
                            '<div class="tpls-alert tpls-alert-danger"><h4>Error</h4><p>' + errorMsg + '</p></div>';
                    }
                }
            };
            
            xhr.send('action=get_family_journey&peopleId=' + tplsCurrentPersonId);
        }
        
        function tplsGenerateFamilyTimeline(data) {
            var family = data.family_info;
            var members = data.family_members;
            
            var html = '<div style="padding: 20px;">';
            
            // Family header
            html += '<div style="text-align: center; margin-bottom: 30px; padding: 20px; background: #f8f9fa; border-radius: 8px;">';
            html += '<h4 style="margin: 0; color: #2c3e50;">Family Engagement Timeline</h4>';
            html += '<p style="margin: 5px 0 0 0; color: #6c757d;">' + family.total_members + ' family members | ';
            html += family.engaged_members + ' engaged | Average score: ' + family.avg_engagement + '</p>';
            html += '<p style="margin: 5px 0 0 0; color: #6c757d;">Timeline from ' + family.timeline_start + ' to ' + family.timeline_end + '</p>';
            html += '</div>';
            
            // Family summary cards
            html += '<div class="row" style="margin-bottom: 30px;">';
            html += '<div class="col-md-3">';
            html += '<div class="stat-card" style="text-align: center; padding: 15px;">';
            html += '<h5 style="color: #007bff; margin: 0;">Family Members</h5>';
            html += '<p style="font-size: 24px; font-weight: bold; margin: 5px 0; color: #2c3e50;">' + family.total_members + '</p>';
            html += '<small style="color: #6c757d;">Total people</small>';
            html += '</div>';
            html += '</div>';
            html += '<div class="col-md-3">';
            html += '<div class="stat-card" style="text-align: center; padding: 15px;">';
            html += '<h5 style="color: #28a745; margin: 0;">Engaged</h5>';
            html += '<p style="font-size: 24px; font-weight: bold; margin: 5px 0; color: #2c3e50;">' + family.engaged_members + '</p>';
            html += '<small style="color: #6c757d;">Score 40+</small>';
            html += '</div>';
            html += '</div>';
            html += '<div class="col-md-3">';
            html += '<div class="stat-card" style="text-align: center; padding: 15px;">';
            html += '<h5 style="color: #ffc107; margin: 0;">Avg Score</h5>';
            html += '<p style="font-size: 24px; font-weight: bold; margin: 5px 0; color: #2c3e50;">' + family.avg_engagement + '</p>';
            html += '<small style="color: #6c757d;">' + tplsGetEngagementLevel(family.avg_engagement) + '</small>';
            html += '</div>';
            html += '</div>';
            html += '<div class="col-md-3">';
            html += '<div class="stat-card" style="text-align: center; padding: 15px;">';
            html += '<h5 style="color: #dc3545; margin: 0;">Timeline</h5>';
            html += '<p style="font-size: 14px; font-weight: bold; margin: 5px 0; color: #2c3e50;">2+ years</p>';
            html += '<small style="color: #6c757d;">Family journey</small>';
            html += '</div>';
            html += '</div>';
            html += '</div>';
            
            // Individual member timelines
            members.forEach(function(member, index) {
                var person = member.person;
                var journey = member.journey;
                var insights = member.insights;
                
                // Member header
                html += '<div style="border: 2px solid #dee2e6; border-radius: 8px; margin-bottom: 25px; overflow: hidden;">';
                html += '<div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 15px;">';
                html += '<div style="display: flex; justify-content: space-between; align-items: center;">';
                html += '<div>';
                html += '<h5 style="margin: 0; font-weight: bold;">' + person.name + '</h5>';
                html += '<small>Age: ' + (person.age || 'Unknown') + ' | ' + (person.position || 'Family Member') + '</small>';
                html += '</div>';
                html += '<div style="text-align: right;">';
                html += '<div style="background: rgba(255,255,255,0.2); padding: 8px 12px; border-radius: 20px;">';
                html += '<strong>Score: ' + person.engagement_score + '</strong>';
                html += '</div>';
                html += '<small>' + tplsGetEngagementLevel(person.engagement_score) + '</small>';
                html += '</div>';
                html += '</div>';
                html += '</div>';
                
                // Member timeline
                html += '<div style="padding: 20px; background: #fafafa;">';
                
                if (journey.length === 0) {
                    html += '<div style="text-align: center; color: #6c757d; font-style: italic; padding: 20px;">';
                    html += '<i class="fa fa-calendar-times-o fa-2x" style="margin-bottom: 10px;"></i><br>';
                    html += 'No engagement events recorded yet';
                    html += '</div>';
                } else {
                    // Mini timeline for each member
                    html += '<div style="position: relative; padding-left: 20px;">';
                    html += '<div style="position: absolute; left: 10px; top: 0; bottom: 0; width: 2px; background: #ddd;"></div>';
                    
                    journey.forEach(function(event, eventIndex) {
                        html += '<div style="position: relative; margin-bottom: 15px; padding-left: 30px;">';
                        
                        // Mini timeline dot
                        html += '<div style="position: absolute; left: -25px; top: 3px; width: 12px; height: 12px; ';
                        html += 'border-radius: 50%; background: ' + event.color + '; border: 2px solid white; box-shadow: 0 0 0 2px ' + event.color + ';"></div>';
                        
                        // Event content (compact)
                        html += '<div style="background: white; border: 1px solid #e9ecef; border-radius: 4px; padding: 8px 12px;">';
                        html += '<div style="display: flex; justify-content: space-between; align-items: center;">';
                        html += '<div>';
                        html += '<strong style="color: ' + event.color + '; font-size: 13px;">' + event.event + '</strong>';
                        html += '<div style="font-size: 12px; color: #6c757d; margin-top: 2px;">' + event.description + '</div>';
                        html += '</div>';
                        html += '<div style="font-size: 11px; color: #999; text-align: right;">';
                        html += '<div>' + event.date + '</div>';
                        html += '<div style="background: ' + event.color + '; color: white; padding: 2px 6px; border-radius: 10px; margin-top: 2px;">';
                        html += event.type.toUpperCase();
                        html += '</div>';
                        html += '</div>';
                        html += '</div>';
                        html += '</div>';
                        html += '</div>';
                    });
                    
                    html += '</div>';
                    
                    // Member insights
                    html += '<div style="margin-top: 15px; padding: 10px; background: #e7f3ff; border-radius: 4px; border-left: 3px solid #007bff;">';
                    html += '<div style="font-size: 12px; color: #0056b3;">';
                    html += '<strong>Journey:</strong> ' + insights.journey_length + ' | ';
                    html += '<strong>Entry:</strong> ' + insights.entry_point + ' | ';
                    html += '<strong>Events:</strong> ' + insights.total_events;
                    html += '</div>';
                    html += '</div>';
                }
                
                html += '</div>';
                html += '</div>';
            });
            
            // Enhanced Family insights
            html += '<div style="margin-top: 30px; padding: 20px; background: #e6f7ff; border-radius: 8px; border-left: 4px solid #007bff;">';
            html += '<h5 style="color: #0056b3; margin: 0 0 15px 0;"><i class="fa fa-users"></i> Family Journey Insights</h5>';
            html += '<div class="row">';
            html += '<div class="col-md-6">';
            html += '<ul style="margin: 0; padding-left: 20px; color: #2c3e50; font-size: 14px;">';
            html += '<li><strong>Family Size:</strong> ' + family.total_members + ' members tracked</li>';
            html += '<li><strong>Engagement Rate:</strong> ' + Math.round((family.engaged_members / family.total_members) * 100) + '% of family engaged</li>';
            
            // Find first engager
            var firstEngager = null;
            var earliestDate = '9999-12-31';
            members.forEach(function(member) {
                if (member.journey.length > 0 && member.journey[0].date < earliestDate) {
                    earliestDate = member.journey[0].date;
                    firstEngager = member.person.name;
                }
            });
            if (firstEngager) {
                html += '<li><strong>First Engager:</strong> ' + firstEngager + ' led family engagement</li>';
            }
            
            // Calculate engagement patterns
            var adultCount = 0;
            var childCount = 0;
            var engagedAdults = 0;
            var engagedChildren = 0;
            members.forEach(function(member) {
                if (member.person.age >= 18) {
                    adultCount++;
                    if (member.person.engagement_score >= 40) engagedAdults++;
                } else if (member.person.age > 0) {
                    childCount++;
                    if (member.person.engagement_score >= 40) engagedChildren++;
                }
            });
            
            if (adultCount > 0) {
                html += '<li><strong>Adult Engagement:</strong> ' + engagedAdults + ' of ' + adultCount + ' adults engaged (' + Math.round((engagedAdults/adultCount)*100) + '%)</li>';
            }
            if (childCount > 0) {
                html += '<li><strong>Child Engagement:</strong> ' + engagedChildren + ' of ' + childCount + ' children engaged (' + Math.round((engagedChildren/childCount)*100) + '%)</li>';
            }
            
            html += '</ul>';
            html += '</div>';
            html += '<div class="col-md-6">';
            html += '<ul style="margin: 0; padding-left: 20px; color: #2c3e50; font-size: 14px;">';
            
            // Check for family patterns
            var hasServingFamily = members.some(function(m) { 
                return m.journey.some(function(e) { return e.type === 'serving'; }); 
            });
            var allAttending = members.every(function(m) { 
                return m.journey.some(function(e) { return e.type === 'attendance'; }); 
            });
            var hasSmallGroup = members.some(function(m) { 
                return m.journey.some(function(e) { return e.type === 'smallgroup'; }); 
            });
            
            if (allAttending && family.total_members > 1) {
                html += '<li><strong>Full Family Engagement:</strong> All members attend regularly</li>';
            }
            if (hasServingFamily) {
                var servingCount = members.filter(function(m) { 
                    return m.journey.some(function(e) { return e.type === 'serving'; }); 
                }).length;
                html += '<li><strong>Serving Family:</strong> ' + servingCount + ' member(s) actively serve</li>';
            }
            if (hasSmallGroup) {
                var smallGroupCount = members.filter(function(m) { 
                    return m.journey.some(function(e) { return e.type === 'smallgroup'; }); 
                }).length;
                html += '<li><strong>Small Group:</strong> ' + smallGroupCount + ' member(s) in small groups</li>';
            }
            
            html += '<li><strong>Journey Span:</strong> ' + tplsCalculateDaysBetween(family.timeline_start, family.timeline_end) + ' days of engagement</li>';
            html += '</ul>';
            html += '</div>';
            html += '</div>';
            
            // Additional insights section
            html += '<div style="margin-top: 20px; padding: 15px; background: #f0f8ff; border-radius: 6px;">';
            html += '<h6 style="color: #0056b3; margin: 0 0 10px 0;"><i class="fa fa-lightbulb-o"></i> Family Engagement Patterns</h6>';
            html += '<div style="font-size: 13px; color: #495057;">';
            
            // Calculate family engagement momentum
            var recentActivity = 0;
            var currentDate = new Date();
            members.forEach(function(member) {
                member.journey.forEach(function(event) {
                    var eventDate = new Date(event.date);
                    var daysDiff = (currentDate - eventDate) / (1000 * 60 * 60 * 24);
                    if (daysDiff <= 90) recentActivity++;
                });
            });
            
            if (recentActivity >= members.length * 2) {
                html += '<p><i class="fa fa-fire" style="color: #ff6b6b;"></i> <strong>High Activity Family:</strong> Multiple recent engagement touchpoints across family members indicate strong momentum.</p>';
            } else if (recentActivity > 0) {
                html += '<p><i class="fa fa-arrow-up" style="color: #51cf66;"></i> <strong>Active Family:</strong> Recent engagement shows ongoing connection with the church.</p>';
            } else {
                html += '<p><i class="fa fa-clock-o" style="color: #feca57;"></i> <strong>Re-engagement Opportunity:</strong> No recent activity - consider reaching out to reconnect.</p>';
            }
            
            // Family engagement recommendations
            if (engagedAdults > 0 && childCount > engagedChildren) {
                html += '<p><i class="fa fa-child" style="color: #5c7cfa;"></i> <strong>Children Ministry Opportunity:</strong> Parents are engaged but children could benefit from more program involvement.</p>';
            }
            if (hasServingFamily && !allAttending) {
                html += '<p><i class="fa fa-users" style="color: #845ef7;"></i> <strong>Serving Leaders:</strong> Some family members serve while others may need encouragement to attend regularly.</p>';
            }
            if (family.avg_engagement < 40) {
                html += '<p><i class="fa fa-star-half-o" style="color: #fd7e14;"></i> <strong>Growth Potential:</strong> Family engagement score indicates opportunities for deeper involvement.</p>';
            }
            
            html += '</div>';
            html += '</div>';
            html += '</div>';
            
            html += '</div>';
            
            return html;
        }
        
        function tplsCalculateDaysBetween(start, end) {
            try {
                var startDate = new Date(start);
                var endDate = new Date(end);
                var diffTime = Math.abs(endDate - startDate);
                var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                return diffDays;
            } catch (e) {
                return 'Unknown';
            }
        }
        
        // Make person journey functions globally available
        window.tplsShowPersonJourney = tplsShowPersonJourney;
        window.tplsClosePersonJourney = tplsClosePersonJourney;
        window.tplsShowIndividualJourney = tplsShowIndividualJourney;
        window.tplsViewFullProfile = tplsViewFullProfile;
        
        // ::START:: Main Initialization
        function tplsInitialize() {
            // DEFENSIVE footer elimination for embed mode
            if (""" + ("true" if embed_mode else "false") + """) {
                function safeFooterCleanup() {
                    try {
                        console.log('Running safe footer cleanup...');
                        
                        // Only target very specific footer elements, with safety checks
                        var specificFooters = [
                            'div[id*="footer"][class*="hidden-print"]',
                            '.footer.hidden-print',
                            'footer[class*="hidden-print"]'
                        ];
                        
                        var removed = 0;
                        specificFooters.forEach(function(selector) {
                            try {
                                var elements = document.querySelectorAll(selector);
                                if (elements && elements.length > 0) {
                                    elements.forEach(function(el) {
                                        if (el && el.parentNode && !el.closest('.tpls-search-container')) {
                                            console.log('Safely removing footer:', selector);
                                            el.parentNode.removeChild(el);
                                            removed++;
                                        }
                                    });
                                }
                            } catch (e) {
                                // Silently ignore errors for individual selectors
                            }
                        });
                        
                        // Only check last child if it's safe to do so
                        try {
                            if (document.body && document.body.children) {
                                var lastChild = document.body.lastElementChild;
                                if (lastChild && 
                                    !lastChild.classList.contains('tpls-search-container') && 
                                    !lastChild.querySelector('.tpls-search-container') &&
                                    lastChild.textContent) {
                                    
                                    var text = lastChild.textContent.toLowerCase();
                                    if (text.includes('first baptist church hendersonville') && 
                                        text.includes('touchpoint')) {
                                        console.log('Removing footer last child');
                                        lastChild.parentNode.removeChild(lastChild);
                                        removed++;
                                    }
                                }
                            }
                        } catch (e) {
                            // Silently ignore
                        }
                        
                        if (removed > 0) {
                            console.log('Safe cleanup removed', removed, 'elements');
                        }
                        
                    } catch (e) {
                        console.error('Safe footer cleanup error:', e);
                    }
                }
                
                // Run cleanup with delays, but much less aggressively
                setTimeout(safeFooterCleanup, 500);
                setTimeout(safeFooterCleanup, 2000);
                setTimeout(safeFooterCleanup, 5000);
                
                // Set up a less aggressive interval
                setInterval(safeFooterCleanup, 10000); // Every 10 seconds instead of constantly
            }
            
            var searchInput = document.getElementById('tplsSearchInput');
            var searchButton = document.getElementById('tplsSearchButton');
            var clearButton = document.getElementById('tplsClearSearch');
            var searchTimeout = null;
            
            if (!searchInput || !searchButton) {
                console.error('Essential search elements not found');
                return;
            }
            
            // Clear button functionality
            if (clearButton) {
                searchInput.addEventListener('input', function() {
                    clearButton.style.display = this.value.length > 0 ? 'block' : 'none';
                });
                
                clearButton.style.display = searchInput.value.length > 0 ? 'block' : 'none';
                
                clearButton.addEventListener('click', function() {
                    searchInput.value = '';
                    clearButton.style.display = 'none';
                    searchInput.focus();
                    
                    var resultsContainer = document.getElementById('tplsResultsContainer');
                    if (resultsContainer) resultsContainer.innerHTML = '';
                });
            }
            
            // Initial search if we have a value
            if (searchInput.value.trim().length > 1) {
                tplsPerformSearch();
            }
            
            // Search as you type with a delay
            searchInput.addEventListener('input', function() {
                clearTimeout(searchTimeout);
                searchTimeout = setTimeout(tplsPerformSearch, TPLS_CONFIG.SEARCH_DELAY);
            });
            
            // Search when the button is clicked
            searchButton.addEventListener('click', function() {
                tplsPerformSearch();
            });
            
            // Search when Enter is pressed
            searchInput.addEventListener('keypress', function(e) {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    tplsPerformSearch();
                }
            });
            
            // Setup other components
            tplsSetupModals();
            tplsLoadKeywords();
            tplsSetupFormSubmission();
        }
        
        // Initialize when DOM is ready
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', tplsInitialize);
        } else {
            tplsInitialize();
        }
        
        // Fallback initialization for widget contexts
        setTimeout(function() {
            var searchInput = document.getElementById('tplsSearchInput');
            if (searchInput && !searchInput.hasAttribute('data-tpls-initialized')) {
                searchInput.setAttribute('data-tpls-initialized', 'true');
                tplsInitialize();
            }
        }, 1000);
        
    })(); // End of IIFE - All JavaScript is now namespaced
    </script>
    
    """ + resize_script

        print html_output

    # ::START:: AJAX Search Handler
    # If AJAX mode, perform the search with namespaced CSS classes
    if not ajax_handled and ajax_mode and search_term:
        try:
            # Helper function to normalize phone numbers for searching
            def normalize_phone(phone_str):
                if not phone_str:
                    return ""
                # Remove all non-digit characters
                return ''.join(c for c in phone_str if c.isdigit())
            
            # Helper function to check if the search term is a phone number
            def is_phone_search(term):
                # A phone search is when we have at least 3 digits in the search term
                digits = ''.join(c for c in term if c.isdigit())
                return len(digits) >= 3
            
            # Helper function to highlight the matching part of text
            def highlight_match(text, term):
                if not text or not term:
                    return text
                
                # For phone numbers, handle differently
                if is_phone_search(term):
                    normalized_term = normalize_phone(term)
                    normalized_text = normalize_phone(text)
                    
                    # If the normalized term is in the normalized text
                    if normalized_term in normalized_text:
                        # Try to format the displayed text with highlighting
                        result = ""
                        term_digits_used = 0
                        for c in text:
                            if c.isdigit() and term_digits_used < len(normalized_term) and c == normalized_term[term_digits_used]:
                                result += '<span class="tpls-match-highlight">' + c + '</span>'
                                term_digits_used += 1
                            else:
                                result += c
                        return result
                
                # For regular text, do case-insensitive search
                import re
                try:
                    pattern = re.compile(re.escape(term), re.IGNORECASE)
                    return pattern.sub(lambda m: '<span class="tpls-match-highlight">' + m.group(0) + '</span>', text)
                except:
                    return text
            
            is_phone_number_search = is_phone_search(search_term)
            normalized_search = normalize_phone(search_term)
            
            # Check if this is a special people search for assignee selection
            if hasattr(model.Data, "people_search") and model.Data.people_search == "1":
                # ::STEP:: People Search for Assignee
                people_search_sql = """
            SELECT TOP {1}
                p.PeopleId, p.Name, p.Age, 
                p.PrimaryAddress, p.PrimaryCity, p.PrimaryState, p.PrimaryZip
            FROM People p
            WHERE p.Name LIKE '%{0}%'
                AND p.DeceasedDate IS NULL
                AND p.ArchivedFlag = 0
            ORDER BY 
                CASE WHEN p.Name LIKE '{0}%' THEN 0 ELSE 1 END,
                p.Name
            """.format(search_term.replace("'", "''"), MAX_RESULTS)
                
                people = q.QuerySql(people_search_sql)
                print "<div class='tpls-results-section'>"
                print "<h3 class='tpls-section-heading'>People</h3>"
                
                if people and len(people) > 0:
                    for person in people:
                        # Get formatted address
                        address = getattr(person, 'PrimaryAddress', '') or ''
                        
                        print """
                        <div class="tpls-result-item">
                            <div class="tpls-result-name">
                                <a href="/Person2/{0}" target="_blank">{1}</a>
                            </div>
                            <div class="tpls-result-meta">
                                <span>Age: {2}</span>
                                <div><i class="fa fa-home"></i> {3}</div>
                            </div>
                        </div>
                        """.format(person.PeopleId, person.Name, person.Age or "", address)
                else:
                    print "<div class='tpls-no-results'>No people found matching your search.</div>"
                
                print "</div>"
            else:
                # ::STEP:: Main Search Results - People
                print "<div class='tpls-results-section'>"
                print "<h3 class='tpls-section-heading'>People</h3>"
                
                # Try to use a simple query for people
                people = []
                try:
                    # If it's a phone number search
                    if is_phone_number_search:
                        # Try SQL search for phone numbers
                        people_sql = """
                    SELECT TOP {1}
                        p.PeopleId, p.Name, p.EmailAddress, p.CellPhone, 
                        p.HomePhone, p.Age, ms.Description AS MemberStatus
                    FROM People p
                    LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
                    WHERE 
                        (p.CellPhone LIKE '%{0}%' OR p.HomePhone LIKE '%{0}%')
                        AND p.DeceasedDate IS NULL
                        AND p.ArchivedFlag = 0
                    ORDER BY p.Name
                        """.format(normalized_search, MAX_RESULTS)
                    else:
                        # First try an exact match
                        name_query = "na='{0}'".format(search_term.replace("'", "''"))
                        people = q.QueryList(name_query, "Name")
                        
                        # If no results, try a more fuzzy search
                        if not people or len(people) == 0:
                            name_query = "na LIKE '%{0}%'".format(search_term.replace("'", "''"))
                            people = q.QueryList(name_query, "Name")
                            
                    # Fallback to SQL if needed
                    if not people or len(people) == 0:
                        people_sql = """
                    SELECT TOP {1}
                        p.PeopleId, p.Name, p.EmailAddress, p.CellPhone, 
                        p.HomePhone, p.Age, ms.Description AS MemberStatus
                    FROM People p
                    LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
                    WHERE 
                        (p.Name LIKE '%{0}%' OR p.CellPhone LIKE '%{0}%' OR p.HomePhone LIKE '%{0}%')
                        AND p.DeceasedDate IS NULL
                        AND p.ArchivedFlag = 0
                    ORDER BY 
                        CASE WHEN p.Name LIKE '{0}%' THEN 0 ELSE 1 END,
                        p.Name
                        """.format(search_term.replace("'", "''"), MAX_RESULTS)
                        people = q.QuerySql(people_sql)
                except Exception as e:
                    # Fallback to direct SQL
                    try:
                        people_sql = """
                    SELECT TOP {1}
                        p.PeopleId, p.Name, p.EmailAddress, p.CellPhone, 
                        p.HomePhone, p.Age, ms.Description AS MemberStatus
                    FROM People p
                    LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
                    WHERE 
                        (p.Name LIKE '%{0}%' OR p.CellPhone LIKE '%{0}%' OR p.HomePhone LIKE '%{0}%')
                        AND p.DeceasedDate IS NULL
                        AND p.ArchivedFlag = 0
                    ORDER BY 
                        CASE WHEN p.Name LIKE '{0}%' THEN 0 ELSE 1 END,
                        p.Name
                        """.format(search_term.replace("'", "''"), MAX_RESULTS)
                        people = q.QuerySql(people_sql)
                    except:
                        people = []
                
                if people and len(people) > 0:
                    for person in people:
                        # Get person attributes safely
                        person_id = getattr(person, 'PeopleId', '')
                        person_name = getattr(person, 'Name', '')
                        person_email = getattr(person, 'EmailAddress', '') or ''
                        person_cell = getattr(person, 'CellPhone', '') or ''
                        person_home = getattr(person, 'HomePhone', '') or ''
                        person_age = getattr(person, 'Age', '') or ''
                        member_status = getattr(person, 'MemberStatus', '') or ''
                    
                        # Format phone numbers if available
                        formatted_cell = ""
                        if person_cell:
                            try:
                                formatted_cell = model.FmtPhone(person_cell)
                                # Check if this is a phone search and highlight if matching
                                if is_phone_number_search and normalized_search in normalize_phone(person_cell):
                                    formatted_cell = highlight_match(formatted_cell, search_term)
                            except:
                                formatted_cell = person_cell
                    
                        formatted_home = ""
                        if person_home:
                            try:
                                formatted_home = model.FmtPhone(person_home)
                                # Check if this is a phone search and highlight if matching
                                if is_phone_number_search and normalized_search in normalize_phone(person_home):
                                    formatted_home = highlight_match(formatted_home, search_term)
                            except:
                                formatted_home = person_home
                    
                        # Highlight name if it matches the search term (and not a phone search)
                        displayed_name = person_name
                        if not is_phone_number_search:
                            displayed_name = highlight_match(person_name, search_term)
                    
                        # Generate the result item HTML with proper button structure
                        print '<div class="tpls-result-item">'
                        print '    <div class="tpls-result-name">'
                        print '        <a href="/Person2/{0}" target="_blank">{1}</a>'.format(person_id, displayed_name)
                        print '        <div class="tpls-result-actions">'
                        print '            <button class="tpls-action-button tpls-journey-btn" onclick="tplsShowPersonJourney({0})" data-person-id="{0}" data-person-name="{1}">'.format(person_id, person_name.replace('"', '&quot;'))
                        print '                <i class="fa fa-map-o"></i> Journey'
                        print '            </button>'
                        print '            <button class="tpls-action-button tpls-add-task-btn" data-person-id="{0}" data-person-name="{1}">'.format(person_id, person_name.replace('"', '&quot;'))
                        print '                <i class="fa fa-tasks"></i> Task'
                        print '            </button>'
                        print '            <button class="tpls-action-button tpls-add-note-btn" data-person-id="{0}" data-person-name="{1}">'.format(person_id, person_name.replace('"', '&quot;'))
                        print '                <i class="fa fa-sticky-note"></i> Note'
                        print '            </button>'
                        print '        </div>'
                        print '    </div>'
                        print '    <div class="tpls-result-meta">'
                    
                        if member_status:
                            print "        <span>{0}</span>".format(member_status)
                    
                        if person_age:
                            print "        <span>  Age: {0}</span>".format(person_age)
                    
                        print "    </div>"
                    
                        # Contact information
                        if person_email or person_cell or person_home:
                            print "    <div class='tpls-result-meta'>"
                        
                            if person_email:
                                print "        <div><i class='fa fa-envelope'></i> {0}</div>".format(person_email)
                        
                            if person_cell:
                                print "        <div><i class='fa fa-mobile'></i> <a href='tel:{0}' class='tpls-phone-link'>{1}</a></div>".format(
                                normalize_phone(person_cell), 
                        formatted_cell
                                )
                        
                            if person_home and person_home != person_cell:
                                print "        <div><i class='fa fa-phone'></i> <a href='tel:{0}' class='tpls-phone-link'>{1}</a></div>".format(
                                normalize_phone(person_home), 
                        formatted_home
                                )
                        
                        print "    </div>"
                    
                        print "</div>"
                    
                else:
                    print "<div class='tpls-no-results'>No people found matching your search.</div>"
                
                print "</div>"
                
                # ::STEP:: Main Search Results - Organizations
                print "<div class='tpls-results-section'>"
                print "<h3 class='tpls-section-heading'>Involvements</h3>"
                
                # Simple query for organizations
                orgs = []
                try:
                    orgs_sql = """
                SELECT TOP {1}
                    o.OrganizationId, o.OrganizationName, o.MemberCount,
                    p.Name AS ProgramName, d.Name AS DivisionName
                FROM Organizations o
                LEFT JOIN Division d ON o.DivisionId = d.Id
                LEFT JOIN Program p ON d.ProgId = p.Id
                WHERE o.OrganizationName LIKE '%{0}%'
                AND o.OrganizationStatusId = 30
                ORDER BY 
                    CASE WHEN o.OrganizationName LIKE '{0}%' THEN 0 ELSE 1 END,
                    o.OrganizationName
                    """.format(search_term.replace("'", "''"), MAX_RESULTS)
                    orgs = q.QuerySql(orgs_sql)
                except Exception as e:
                    orgs = []
            
                if orgs and len(orgs) > 0:
                    for org in orgs:
                        # Get org attributes safely
                        org_id = getattr(org, 'OrganizationId', '')
                        org_name = getattr(org, 'OrganizationName', '')
                        member_count = getattr(org, 'MemberCount', 0) or 0
                        program_name = getattr(org, 'ProgramName', '') or ''
                        division_name = getattr(org, 'DivisionName', '') or ''
                    
                        # Highlight organization name if it matches the search term
                        displayed_org_name = highlight_match(org_name, search_term)
                    
                        print """
                        <div class="tpls-result-item">
                        <div class="tpls-result-name">
                        <a href="/Org/{0}" target="_blank">{1}</a>
                        </div>
                        <div class="tpls-result-meta">
                        """.format(org_id, displayed_org_name)
                    
                        if program_name or division_name:
                            path_parts = []
                            if program_name:
                                path_parts.append(program_name)
                            if division_name:
                                path_parts.append(division_name)
                        
                            if path_parts:
                                print "<span>{0}</span>".format(" &gt; ".join(path_parts))
                    
                        print """
                        <span>  Members: {0}</span>
                        </div>
                        </div>
                        """.format(member_count)
                else:
                    print "<div class='tpls-no-results'>No involvements found matching your search.</div>"
            
                    print "</div>"
        
        except Exception as e:
            # Display any errors in a user-friendly way
            import traceback
            print "<div style='color:red; padding:15px; border:1px solid red; border-radius:4px; margin-top:20px;'>"
            print "<strong>Error performing search:</strong> " + str(e)
            print "</div>"
