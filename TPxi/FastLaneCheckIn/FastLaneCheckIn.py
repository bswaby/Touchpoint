"""
FastLane Check-In System for Touchpoint

This script provides a fast, efficient check-in system for high-volume events (100-2500+ people).
Features:
- Meeting selection for current day
- Optimized for speed with bulk actions
- Multiple person selection for efficient check-in
- Quick alpha-based filtering and search
- Family-based check-in
- One-click check-in process
- Real-time stats and metrics
- Check-in correction functionality

NOTE: this does not print as those options are not exposed for me to access

--Upload Instructions Start--
To upload code to Touchpoint, use the following steps:
1. Click Admin ~ Advanced ~ Special Content ~ Python
2. Click New Python Script File
3. Name the Python FastLaneCheckIn (case sensitive) and paste all this code.  If you name it something else, then just update variabl below.
4. Test and optionally add to menu
--Upload Instructions End--

Written By: Ben Swaby
Email: bswaby@fbchtn.org
"""


import traceback
import json
import datetime
import re
from System import DateTime
from System.Collections.Generic import List

# Constants and configuration
Script_Name = "FastLaneCheckIn"
PAGE_SIZE = 50  # Number of people to show per page - smaller for faster loading
DATE_FORMAT = "M/d/yyyy"
ATTEND_FLAG = 1  # Present flag for attendance
ALPHA_BLOCKS = [
    "A-C", "D-F", "G-I", "J-L", "M-O", "P-R", "S-U", "V-Z", "All"
]

# Helper functions - defined first to avoid syntax errors
def get_script_name():
    """Get the current script name from the URL"""
    script_name = Script_Name #"RapidCheckIn"  # Default name
    if hasattr(model, 'Request') and hasattr(model.Request, 'RawUrl'):
        url_parts = model.Request.RawUrl.split('/')
        if len(url_parts) > 2:
            script_name = url_parts[-1].split('?')[0]  # Get the last part of the URL without query params
    return script_name

def print_header(title):
    """Print the HTML header"""
    print """
    <div class="row">
        <div class="col-md-12">
            <h2>{0}</h2>
            <hr>
        </div>
    </div>
    """.format(title)

def get_org_name(check_in_manager, org_id):
    """Get the name of an organization by ID"""
    for meeting in check_in_manager.all_meetings_today:
        if str(meeting.org_id) == str(org_id):
            return meeting.org_name
    return "Organization {0}".format(org_id)
    
def render_alpha_filters(current_filter):
    """Render the alpha filter buttons"""
    buttons = []
    for alpha_block in ALPHA_BLOCKS:
        active = 'active' if alpha_block == current_filter else ''
        buttons.append("""
            <label class="btn btn-default {0}">
                <input type="radio" name="alpha_filter" value="{1}" {2} onchange="this.form.submit()"> {1}
            </label>
        """.format(active, alpha_block, 'checked' if alpha_block == current_filter else ''))
    
    return """
        <div class="btn-group btn-group-sm" data-toggle="buttons">
            {0}
        </div>
    """.format("\n".join(buttons))

# Helper classes
class MeetingInfo:
    """Class to store and manage meeting information"""
    def __init__(self, meeting_id, org_id, meeting_date, org_name, location, program_name=None):
        self.meeting_id = meeting_id
        self.org_id = org_id
        self.meeting_date = meeting_date
        self.org_name = org_name
        self.location = location
        self.program_name = program_name  # Added program_name

class PersonInfo:
    """Class to store person information for check-in"""
    def __init__(self, people_id, name, family_id=None, org_ids=None, checked_in=False):
        self.people_id = people_id
        self.name = name
        self.family_id = family_id
        self.org_ids = org_ids or []
        self.checked_in = checked_in

class CheckInManager:
    """Manages check-in operations and state"""
    
    def __init__(self, model, q):
        self.model = model
        self.q = q
        self.today = self.model.DateTime.Date
        self.selected_meetings = []
        self.all_meetings_today = []
        self.last_check_in_time = None  # Track last check-in time for stats refresh
        
    def get_todays_meetings(self):
        """Get all meetings scheduled for today including program information"""
        today = self.model.DateTime.Date
        
        # First approach: Try using SQL directly to ensure we can find meetings
        sql = """
            SELECT m.MeetingId, m.OrganizationId, m.MeetingDate, o.OrganizationName, m.Location,
                   os.Program as ProgramName
            FROM Meetings m
            JOIN Organizations o ON m.OrganizationId = o.OrganizationId
            LEFT JOIN OrganizationStructure os ON m.OrganizationId = os.OrgId
            WHERE CONVERT(date, m.MeetingDate) = CONVERT(date, GETDATE())
            AND m.DidNotMeet = 0
            ORDER BY o.OrganizationName
        """
        sql_results = self.q.QuerySql(sql)
        
        # Try to find meetings that exist today
        meetings = []
        for result in sql_results:
            meeting = MeetingInfo(
                result.MeetingId,
                result.OrganizationId,
                result.MeetingDate,
                result.OrganizationName,
                result.Location or "",
                result.ProgramName if hasattr(result, 'ProgramName') else None
            )
            meetings.append(meeting)
        
        self.all_meetings_today = meetings
        return meetings
    
    def get_people_by_filter(self, alpha_filter, search_term="", meeting_ids=None, page=1, page_size=PAGE_SIZE):
        """Get people by alpha filter or search term who are members of the selected orgs"""
        if not meeting_ids:
            print "<!-- DEBUG: No meeting IDs provided to get_people_by_filter -->"
            return [], 0
        
        # Debug meeting IDs
        for meeting_id in meeting_ids:
            print "<!-- DEBUG: Checking meeting ID: " + str(meeting_id) + " -->"
        
        # Ensure meetings are loaded
        if not self.all_meetings_today:
            print "<!-- DEBUG: Loading meetings for get_people_by_filter -->"
            self.get_todays_meetings()
            
        # Get org IDs from meeting IDs
        org_ids = []
        for meeting_id in meeting_ids:
            meeting_id_str = str(meeting_id)
            for m in self.all_meetings_today:
                if str(m.meeting_id) == meeting_id_str:
                    org_ids.append(str(m.org_id))
        
        # If no org IDs found, try direct approach with meeting IDs
        if not org_ids:
            return self.get_people_by_meeting_ids(meeting_ids, alpha_filter, search_term, page, page_size)
            
        org_ids_str = ",".join(org_ids)
        
        # Build SQL for count query first
        count_sql = """
            SELECT COUNT(DISTINCT p.PeopleId) AS TotalCount
            FROM People p
            JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
            WHERE om.OrganizationId IN ({0})
            AND (p.IsDeceased IS NULL OR p.IsDeceased = 0)
            AND (om.InactiveDate IS NULL OR om.InactiveDate > GETDATE())
        """.format(org_ids_str)
        
        # Apply alpha filter if not "All"
        if alpha_filter and alpha_filter != "All":
            first_letter = alpha_filter[0]
            last_letter = alpha_filter[2]
            count_sql += " AND p.LastName >= '{0}' AND p.LastName <= '{1}z'".format(first_letter, last_letter)
        
        # Apply search term if provided
        if search_term:
            count_sql += " AND (p.Name LIKE '%{0}%' OR p.FirstName LIKE '%{0}%' OR p.LastName LIKE '%{0}%')".format(search_term)
        
        # Check if already checked in
        count_sql += """
            AND NOT EXISTS (
                SELECT 1 FROM Attend a 
                WHERE a.PeopleId = p.PeopleId 
                AND a.OrganizationId IN ({0}) 
                AND CONVERT(date, a.MeetingDate) = CONVERT(date, GETDATE())
                AND a.AttendanceFlag = 1
            )
        """.format(org_ids_str)
        
        # Get total count first
        count_result = self.q.QuerySqlTop1(count_sql)
        total_count = count_result.TotalCount if hasattr(count_result, 'TotalCount') else 0
        
        # Start building the SQL query for the actual data
        sql = """
            SELECT DISTINCT p.PeopleId, p.Name, p.FamilyId, p.LastName, p.FirstName
            FROM People p
            JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
            WHERE om.OrganizationId IN ({0})
            AND (p.IsDeceased IS NULL OR p.IsDeceased = 0)
            AND (om.InactiveDate IS NULL OR om.InactiveDate > GETDATE())
        """.format(org_ids_str)
        
        # Apply alpha filter if not "All"
        if alpha_filter and alpha_filter != "All":
            first_letter = alpha_filter[0]
            last_letter = alpha_filter[2]
            sql += " AND p.LastName >= '{0}' AND p.LastName <= '{1}z'".format(first_letter, last_letter)
        
        # Apply search term if provided
        if search_term:
            sql += " AND (p.Name LIKE '%{0}%' OR p.FirstName LIKE '%{0}%' OR p.LastName LIKE '%{0}%')".format(search_term)
        
        # Check if already checked in
        sql += """
            AND NOT EXISTS (
                SELECT 1 FROM Attend a 
                WHERE a.PeopleId = p.PeopleId 
                AND a.OrganizationId IN ({0}) 
                AND CONVERT(date, a.MeetingDate) = CONVERT(date, GETDATE())
                AND a.AttendanceFlag = 1
            )
        """.format(org_ids_str)
        
        # Add pagination
        sql += """ 
            ORDER BY p.LastName, p.FirstName
            OFFSET {0} ROWS
            FETCH NEXT {1} ROWS ONLY
        """.format((page - 1) * page_size, page_size)
        
        results = self.q.QuerySql(sql)
        
        # Convert to PersonInfo objects
        people = []
        for result in results:
            # Get the organizations this person is a member of
            person_org_ids = self.get_person_org_ids(result.PeopleId, org_ids)
            
            person = PersonInfo(
                result.PeopleId,
                result.Name,
                result.FamilyId,
                person_org_ids
            )
            people.append(person)
            
        return people, total_count
            
    def get_people_by_meeting_ids(self, meeting_ids, alpha_filter="All", search_term="", page=1, page_size=PAGE_SIZE):
        """Alternative method to get people directly by meeting IDs when org mapping fails"""
        if not meeting_ids:
            return [], 0
                
        # Remove duplicates from meeting_ids
        unique_meeting_ids = []
        for m_id in meeting_ids:
            if m_id not in unique_meeting_ids:
                unique_meeting_ids.append(m_id)
        
        meeting_ids = unique_meeting_ids
        meeting_ids_str = ",".join([str(m) for m in meeting_ids])
        
        # First get count
        count_sql = """
            SELECT COUNT(DISTINCT p.PeopleId) AS TotalCount
            FROM People p
            JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
            JOIN Meetings m ON m.OrganizationId = om.OrganizationId
            WHERE m.MeetingId IN ({0})
            AND (p.IsDeceased IS NULL OR p.IsDeceased = 0)
            AND (om.InactiveDate IS NULL OR om.InactiveDate > GETDATE())
        """.format(meeting_ids_str)
        
        # Apply alpha filter if not "All"
        if alpha_filter and alpha_filter != "All":
            if len(alpha_filter) >= 3:  # Make sure the alpha filter has enough characters
                first_letter = alpha_filter[0]
                last_letter = alpha_filter[2]
                count_sql += " AND p.LastName >= '{0}' AND p.LastName <= '{1}z'".format(first_letter, last_letter)
        
        # Apply search term if provided
        if search_term:
            count_sql += " AND (p.Name LIKE '%{0}%' OR p.FirstName LIKE '%{0}%' OR p.LastName LIKE '%{0}%')".format(search_term)
        
        # Check if already checked in
        count_sql += """
            AND NOT EXISTS (
                SELECT 1 FROM Attend a 
                WHERE a.PeopleId = p.PeopleId 
                AND a.MeetingId IN ({0})
                AND a.AttendanceFlag = 1
            )
        """.format(meeting_ids_str)
        
        # Get total count
        try:
            count_result = self.q.QuerySqlTop1(count_sql)
            total_count = count_result.TotalCount if hasattr(count_result, 'TotalCount') else 0
        except Exception as e:
            print "<div style='color:red;'>Error in count query: " + str(e) + "</div>"
            print "<div style='color:gray;'>SQL: " + count_sql + "</div>"
            total_count = 0
        
        # FIX: When using DISTINCT, include ORDER BY columns in the SELECT list
        sql = """
            SELECT DISTINCT p.PeopleId, p.Name, p.FamilyId, p.LastName, p.FirstName
            FROM People p
            JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
            JOIN Meetings m ON m.OrganizationId = om.OrganizationId
            WHERE m.MeetingId IN ({0})
            AND (p.IsDeceased IS NULL OR p.IsDeceased = 0)
            AND (om.InactiveDate IS NULL OR om.InactiveDate > GETDATE())
        """.format(meeting_ids_str)
        
        # Apply alpha filter if not "All"
        if alpha_filter and alpha_filter != "All":
            if len(alpha_filter) >= 3:  # Make sure the alpha filter has enough characters
                first_letter = alpha_filter[0]
                last_letter = alpha_filter[2]
                sql += " AND p.LastName >= '{0}' AND p.LastName <= '{1}z'".format(first_letter, last_letter)
        
        # Apply search term if provided
        if search_term:
            sql += " AND (p.Name LIKE '%{0}%' OR p.FirstName LIKE '%{0}%' OR p.LastName LIKE '%{0}%')".format(search_term)
        
        # Check if already checked in - using meeting IDs this time
        sql += """
            AND NOT EXISTS (
                SELECT 1 FROM Attend a 
                WHERE a.PeopleId = p.PeopleId 
                AND a.MeetingId IN ({0})
                AND a.AttendanceFlag = 1
            )
        """.format(meeting_ids_str)
        
        # Add pagination
        sql += """ 
            ORDER BY p.LastName, p.FirstName
            OFFSET {0} ROWS
            FETCH NEXT {1} ROWS ONLY
        """.format((page - 1) * page_size, page_size)
        
        try:
            results = self.q.QuerySql(sql)
        except Exception as e:
            print "<div style='color:red;'>Error in people query: " + str(e) + "</div>"
            print "<div style='color:gray;'>SQL: " + sql + "</div>"
            return [], total_count
        
        # Get org IDs from meetings for displaying purposes
        org_ids_by_meeting = {}
        if results:  # Only do this if we found people
            try:
                meeting_to_org_sql = """
                    SELECT MeetingId, OrganizationId 
                    FROM Meetings 
                    WHERE MeetingId IN ({0})
                """.format(meeting_ids_str)
                meeting_org_results = self.q.QuerySql(meeting_to_org_sql)
                
                # Build a mapping from meeting ID to org ID
                for result in meeting_org_results:
                    org_ids_by_meeting[str(result.MeetingId)] = str(result.OrganizationId)
            except Exception as e:
                print "<div style='color:red;'>Error getting org IDs: " + str(e) + "</div>"
        
        # Convert to PersonInfo objects
        people = []
        for result in results:
            try:
                # Get the organizations linked to selected meetings
                all_org_ids = org_ids_by_meeting.values()
                
                # Get the organizations this person is a member of
                person_org_ids = self.get_person_org_ids(result.PeopleId, all_org_ids)
                
                if person_org_ids:  # Only include if they're a member of at least one org
                    person = PersonInfo(
                        result.PeopleId,
                        result.Name,
                        result.FamilyId,
                        person_org_ids
                    )
                    people.append(person)
            except Exception as e:
                print "<div style='color:red;'>Error processing person: " + str(e) + "</div>"
                
        return people, total_count
    
    def get_person_org_ids(self, people_id, org_ids):
        """Get the organization IDs that this person is a member of"""
        if not org_ids:
            return []
            
        org_ids_str = ",".join([str(oid) for oid in org_ids])
        
        sql = """
            SELECT OrganizationId 
            FROM OrganizationMembers 
            WHERE PeopleId = {0}
            AND OrganizationId IN ({1})
            AND (InactiveDate IS NULL OR InactiveDate > GETDATE())
        """.format(people_id, org_ids_str)
        
        results = self.q.QuerySql(sql)
        org_list = [str(result.OrganizationId) for result in results]
        
        return org_list
    
    def get_family_members(self, family_id, org_ids):
        """Get family members who are members of the selected orgs"""
        if not family_id or not org_ids:
            return []
            
        org_ids_str = ",".join([str(oid) for oid in org_ids])
        
        sql = """
            SELECT DISTINCT p.PeopleId, p.Name
            FROM People p
            JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
            WHERE p.FamilyId = {0}
            AND p.PeopleId != {1}
            AND om.OrganizationId IN ({2})
            AND (p.IsDeceased IS NULL OR p.IsDeceased = 0)
            AND (om.InactiveDate IS NULL OR om.InactiveDate > GETDATE())
            AND NOT EXISTS (
                SELECT 1 FROM Attend a 
                WHERE a.PeopleId = p.PeopleId 
                AND a.OrganizationId IN ({2}) 
                AND CONVERT(date, a.MeetingDate) = CONVERT(date, GETDATE())
                AND a.AttendanceFlag = 1
            )
            ORDER BY p.Name
        """.format(family_id, self.model.Data.selected_person_id, org_ids_str)
        
        results = self.q.QuerySql(sql)
        
        # Convert to PersonInfo objects
        family_members = []
        for result in results:
            # Get the organizations this person is a member of
            person_org_ids = self.get_person_org_ids(result.PeopleId, org_ids)
            
            person = PersonInfo(
                result.PeopleId,
                result.Name,
                family_id,
                person_org_ids
            )
            family_members.append(person)
            
        return family_members
    
    def bulk_check_in(self, person_ids, meeting_ids):
        """Check in multiple people at once - optimized for speed"""
        if not person_ids or not meeting_ids:
            return False
            
        success = True
        
        # Process each meeting and person combination directly
        for meeting_id in meeting_ids:
            for person_id in person_ids:
                try:
                    # CRITICAL FIX: Cast to integers
                    result = self.model.EditPersonAttendance(int(meeting_id), int(person_id), True)
                    print "<!-- DEBUG: Bulk check-in result for " + str(person_id) + ": " + str(result) + " -->"
                    if "Success" not in str(result):
                        success = False
                except Exception as e:
                    success = False
                    print "<!-- DEBUG: Bulk check-in error: " + str(e) + " -->"
        
        return success
        
    # EditPersonAttendance API to check-in 
    def check_in_person(self, people_id, meeting_id):
        """Simplified direct check-in using just people_id and meeting_id"""
        try:
            # Convert IDs to integers
            people_id_int = int(people_id)
            meeting_id_int = int(meeting_id)
            
            # Direct API call
            result = self.model.EditPersonAttendance(meeting_id_int, people_id_int, True)
            
            # Check if successful
            return "Success" in str(result)
        except Exception as e:
            print "<div style='color:red'>Error checking in person: " + str(e) + "</div>"
            return False
    
    # Simplified people list rendering function - just displays each person with a direct check-in button
    def render_simplified_people_list(people, check_in_manager, script_name, meeting_ids, alpha_filter, search_term, current_page):
        """Render people list with direct check-in buttons"""
        people_list_html = []
        
        for person in people:
            # Get primary meeting ID
            meeting_id = meeting_ids[0] if meeting_ids else ""
            
            # Create item HTML with direct check-in button
            item_html = """
            <div class="list-group-item">
                <div class="row">
                    <div class="col-xs-8">
                        <span class="person-name">{1}</span>
                    </div>
                    <div class="col-xs-4 text-right">
                        <form method="post" action="/PyScriptForm/{2}" style="display:inline;" onsubmit="showLoading()">
                            <input type="hidden" name="step" value="check_in">
                            <input type="hidden" name="action" value="single_direct_check_in">
                            <input type="hidden" name="person_id" value="{0}">
                            <input type="hidden" name="meeting_id" value="{3}">
                            <input type="hidden" name="view_mode" value="not_checked_in">
                            <input type="hidden" name="alpha_filter" value="{4}">
                            <input type="hidden" name="search_term" value="{5}">
                            <input type="hidden" name="page" value="{6}">
                            <button type="submit" class="btn btn-xs btn-success" onclick="showLoading()">
                                <i class="fa fa-check"></i> Check In
                            </button>
                        </form>
                    </div>
                </div>
            </div>
            """.format(
                person.people_id,  # {0}
                person.name,       # {1}
                script_name,       # {2}
                meeting_id,        # {3}
                alpha_filter,      # {4}
                search_term,       # {5}
                current_page       # {6}
            )
            
            people_list_html.append(item_html)
        
        return "\n".join(people_list_html)
    

    
    def remove_check_in(self, people_id, meeting_ids):
        """Remove check-in for a person using the Meeting API"""
        if not people_id or not meeting_ids:
            return False
            
        success = True
        
        # Use direct approach
        for meeting_id in meeting_ids:
            try:
                # CRITICAL FIX: Cast to integers
                result = self.model.EditPersonAttendance(int(meeting_id), int(people_id), False)
                print "DEBUG: Remove check-in result: " + str(result) 
                if "Success" not in str(result):
                    success = False
            except Exception as e:
                success = False
                print "DEBUG: Remove check-in error: " + str(e) 
                        
        return success

    def remove_check_in(self, people_id, meeting_ids):
        """Remove check-in for a person using the Meeting API with improved stats refresh"""
        if not people_id or not meeting_ids:
            return False
            
        success = True
        
        # Use direct approach
        for meeting_id in meeting_ids:
            try:
                # CRITICAL FIX: Cast to integers
                result = self.model.EditPersonAttendance(int(meeting_id), int(people_id), False)
                print "<!-- DEBUG: Remove check-in result: " + str(result) + " -->"
                if "Success" not in str(result):
                    success = False
            except Exception as e:
                success = False
                print "<!-- DEBUG: Remove check-in error: " + str(e) + " -->"
                
        # Force refresh of stats after removal
        self.last_check_in_time = self.model.DateTime
                        
        return success
    
    def get_check_in_stats(self, meeting_ids):
        """Get check-in statistics for selected meetings with forced refresh"""
        if not meeting_ids:
            return {"checked_in": 0, "not_checked_in": 0, "total": 0}
            
        # Make sure meetings are loaded
        if not self.all_meetings_today:
            self.get_todays_meetings()
            
        meeting_ids_str = ",".join([str(m) for m in meeting_ids])
        
        # Get organization IDs for the meetings
        sql = """
            SELECT DISTINCT OrganizationId
            FROM Meetings
            WHERE MeetingId IN ({0})
        """.format(meeting_ids_str)
        
        org_results = self.q.QuerySql(sql)
        org_ids = [str(r.OrganizationId) for r in org_results]
        
        if not org_ids:
            return {"checked_in": 0, "not_checked_in": 0, "total": 0}
            
        org_ids_str = ",".join(org_ids)
        
        # Get total eligible members (fast count query)
        sql = """
            SELECT COUNT(DISTINCT p.PeopleId) AS TotalCount
            FROM People p
            JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
            WHERE om.OrganizationId IN ({0})
            AND (p.IsDeceased IS NULL OR p.IsDeceased = 0)
            AND (om.InactiveDate IS NULL OR om.InactiveDate > GETDATE())
        """.format(org_ids_str)
        
        total_result = self.q.QuerySqlTop1(sql)
        total = total_result.TotalCount if hasattr(total_result, 'TotalCount') else 0
        
        # Get checked-in count (directly from meetings)
        # Include current date/time to force SQL cache refresh
        now = self.model.DateTime.ToString("yyyy-MM-dd HH:mm:ss.fff")
        sql = """
            SELECT COUNT(DISTINCT a.PeopleId) AS CheckedInCount
            FROM Attend a
            WHERE a.MeetingId IN ({0})
            AND a.AttendanceFlag = 1
            -- Force SQL to recognize this as a unique query to avoid caching
            AND '{1}' = '{1}'
        """.format(meeting_ids_str, now)
        
        checked_in_result = self.q.QuerySqlTop1(sql)
        checked_in = checked_in_result.CheckedInCount if hasattr(checked_in_result, 'CheckedInCount') else 0
        
        return {
            "checked_in": checked_in,
            "not_checked_in": total - checked_in,
            "total": total
        }
    

    def get_checked_in_people(self, meeting_ids, page=1, page_size=PAGE_SIZE):
        """Get people who have already checked in"""
        if not meeting_ids:
            return [], 0
            
        meeting_ids_str = ",".join([str(m) for m in meeting_ids])
        
        # First get total count
        count_sql = """
            SELECT COUNT(DISTINCT p.PeopleId) AS TotalCount
            FROM People p
            JOIN Attend a ON p.PeopleId = a.PeopleId
            WHERE a.MeetingId IN ({0})
            AND a.AttendanceFlag = 1
        """.format(meeting_ids_str)
        
        count_result = self.q.QuerySqlTop1(count_sql)
        total_count = count_result.TotalCount if hasattr(count_result, 'TotalCount') else 0
        
        # Then get paginated data
        sql = """
            SELECT DISTINCT p.PeopleId, p.Name, p.FamilyId, p.LastName, p.FirstName
            FROM People p
            JOIN Attend a ON p.PeopleId = a.PeopleId
            WHERE a.MeetingId IN ({0})
            AND a.AttendanceFlag = 1
            ORDER BY p.LastName, p.FirstName
            OFFSET {1} ROWS
            FETCH NEXT {2} ROWS ONLY
        """.format(meeting_ids_str, (page - 1) * page_size, page_size)
        
        results = self.q.QuerySql(sql)
        
        # Get organization IDs for the meetings
        org_sql = """
            SELECT DISTINCT OrganizationId
            FROM Meetings
            WHERE MeetingId IN ({0})
        """.format(meeting_ids_str)
        
        org_results = self.q.QuerySql(org_sql)
        org_ids = [str(r.OrganizationId) for r in org_results]
        
        # Convert to PersonInfo objects
        people = []
        for result in results:
            # Get the organizations this person is checked into
            person_org_ids = self.get_person_attended_org_ids(result.PeopleId, org_ids)
            
            person = PersonInfo(
                result.PeopleId,
                result.Name,
                result.FamilyId,
                person_org_ids,
                True
            )
            people.append(person)
            
        return people, total_count
        

    
    def get_person_attended_org_ids(self, people_id, org_ids):
        """Get the organization IDs that this person has attended today"""
        if not org_ids:
            return []
            
        org_ids_str = ",".join([str(oid) for oid in org_ids])
        
        sql = """
            SELECT OrganizationId 
            FROM Attend 
            WHERE PeopleId = {0}
            AND OrganizationId IN ({1})
            AND CONVERT(date, MeetingDate) = CONVERT(date, GETDATE())
            AND AttendanceFlag = 1
        """.format(people_id, org_ids_str)
        
        results = self.q.QuerySql(sql)
        return [str(result.OrganizationId) for result in results]
        
    def direct_check_in(self, people_id, meeting_id):
        """Direct check-in using model.EditPersonAttendance API"""
        try:
            # Convert to integers
            people_id_int = int(people_id)
            meeting_id_int = int(meeting_id)
            
            # Make the API call
            result = self.model.EditPersonAttendance(meeting_id_int, people_id_int, True)
            
            # Return success based on the result
            return "Success" in str(result)
        except Exception as e:
            print "<div style='color:red;'>Check-in error: " + str(e) + "</div>"
            return False

    # Add this to your CheckInManager class to enable stats refresh after actions
    def __init__(self, model, q):
        self.model = model
        self.q = q
        self.today = self.model.DateTime.Date
        self.selected_meetings = []
        self.all_meetings_today = []
        self.last_check_in_time = None  # Track last check-in time for stats refresh
    
    # Modify the get_check_in_stats method to use the latest data
    def get_check_in_stats(self, meeting_ids):
        """Get up-to-date check-in statistics for selected meetings"""
        if not meeting_ids:
            return {"checked_in": 0, "not_checked_in": 0, "total": 0}
            
        # Make sure meetings are loaded
        if not self.all_meetings_today:
            self.get_todays_meetings()
            
        meeting_ids_str = ",".join([str(m) for m in meeting_ids])
        
        # Get organization IDs for the meetings
        sql = """
            SELECT DISTINCT OrganizationId
            FROM Meetings
            WHERE MeetingId IN ({0})
        """.format(meeting_ids_str)
        
        org_results = self.q.QuerySql(sql)
        org_ids = [str(r.OrganizationId) for r in org_results]
        
        if not org_ids:
            return {"checked_in": 0, "not_checked_in": 0, "total": 0}
            
        org_ids_str = ",".join(org_ids)
        
        # Get total eligible members (fast count query)
        sql = """
            SELECT COUNT(DISTINCT p.PeopleId) AS TotalCount
            FROM People p
            JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
            WHERE om.OrganizationId IN ({0})
            AND (p.IsDeceased IS NULL OR p.IsDeceased = 0)
            AND (om.InactiveDate IS NULL OR om.InactiveDate > GETDATE())
        """.format(org_ids_str)
        
        total_result = self.q.QuerySqlTop1(sql)
        total = total_result.TotalCount if hasattr(total_result, 'TotalCount') else 0
        
        # Get checked-in count (directly from meetings)
        sql = """
            SELECT COUNT(DISTINCT a.PeopleId) AS CheckedInCount
            FROM Attend a
            WHERE a.MeetingId IN ({0})
            AND a.AttendanceFlag = 1
        """.format(meeting_ids_str)
        
        checked_in_result = self.q.QuerySqlTop1(sql)
        checked_in = checked_in_result.CheckedInCount if hasattr(checked_in_result, 'CheckedInCount') else 0
        
        return {
            "checked_in": checked_in,
            "not_checked_in": total - checked_in,
            "total": total
        }

# Rendering functions
def render_meeting_selection(check_in_manager):
    """Render the meeting selection page"""
    meetings = check_in_manager.get_todays_meetings()
    
    # Get the current script name from the URL
    script_name = get_script_name()
    
    # Get all unique programs for filtering
    programs = set()
    for meeting in meetings:
        # This assumes you have program information in the meeting object
        # If not available directly, you could query it from the organization
        if hasattr(meeting, 'program_name') and meeting.program_name:
            programs.add(meeting.program_name)
    
    programs = sorted(list(programs))
    
    # Get selected program filter if any
    selected_program = getattr(check_in_manager.model.Data, 'program_filter', '')
    
    # Start the page with modern styling
    print """<!DOCTYPE html>
    <html>
    <head>
        <title>FastLane Check-In System</title>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            body { 
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                margin: 0; 
                padding: 20px;
                color: #333;
                background-color: #f9f9f9;
            }
            h2 { 
                margin-top: 0;
                color: #2c3e50;
                font-weight: 500;
            }
            hr {
                border: 0;
                height: 1px;
                background-color: #e0e0e0;
                margin: 15px 0;
            }
            .btn { 
                display: inline-block; 
                padding: 8px 16px; 
                margin-bottom: 0; 
                font-size: 14px; 
                font-weight: 400; 
                line-height: 1.42857143; 
                text-align: center; 
                white-space: nowrap; 
                vertical-align: middle; 
                cursor: pointer; 
                border: 1px solid transparent; 
                border-radius: 4px;
                transition: all 0.3s;
            }
            .btn-primary { color: #fff; background-color: #3498db; border-color: #2980b9; }
            .btn-primary:hover { background-color: #2980b9; }
            .btn-default { color: #333; background-color: #fff; border-color: #ccc; }
            .btn-default:hover { background-color: #f5f5f5; }
            .btn-success { color: #fff; background-color: #2ecc71; border-color: #27ae60; }
            .btn-success:hover { background-color: #27ae60; }
            .btn-lg { padding: 12px 20px; font-size: 16px; }
            .btn-sm { padding: 5px 10px; font-size: 12px; }
            .form-control { 
                display: block; 
                width: 100%; 
                height: 38px; 
                padding: 8px 12px; 
                font-size: 14px;
                color: #555; 
                background-color: #fff; 
                border: 1px solid #ddd; 
                border-radius: 4px;
                box-shadow: inset 0 1px 1px rgba(0,0,0,.075);
                transition: border-color ease-in-out .15s,box-shadow ease-in-out .15s;
            }
            .form-control:focus {
                border-color: #3498db;
                outline: 0;
                box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(52,152,219,.6);
            }
            .well { 
                min-height: 20px; 
                padding: 19px; 
                margin-bottom: 20px; 
                background-color: #fff;
                border: 1px solid #e3e3e3; 
                border-radius: 4px; 
                box-shadow: 0 1px 3px rgba(0,0,0,.05);
            }
            .checkbox { 
                position: relative; 
                display: block; 
                margin-top: 10px; 
                margin-bottom: 10px; 
            }
            .checkbox label {
                padding-left: 25px;
                cursor: pointer;
            }
            .checkbox input[type=checkbox] {
                position: absolute;
                margin-left: -20px;
                margin-top: 2px;
                cursor: pointer;
            }
            .panel { 
                margin-bottom: 20px; 
                background-color: #fff; 
                border: 1px solid #ddd; 
                border-radius: 4px;
                box-shadow: 0 1px 3px rgba(0,0,0,.05);
            }
            .panel-heading { 
                padding: 12px 15px; 
                background-color: #f7f7f7;
                border-bottom: 1px solid #ddd; 
                border-top-left-radius: 3px;
                border-top-right-radius: 3px; 
            }
            .panel-body { padding: 15px; }
            .panel-title { 
                margin-top: 0; 
                margin-bottom: 0; 
                font-size: 16px; 
                color: #333;
                font-weight: 500;
            }
            .alert { 
                padding: 15px; 
                margin-bottom: 20px; 
                border: 1px solid transparent; 
                border-radius: 4px; 
            }
            .alert-warning { color: #8a6d3b; background-color: #fcf8e3; border-color: #faebcc; }
            .alert-info { color: #31708f; background-color: #d9edf7; border-color: #bce8f1; }
            .row { 
                display: flex; 
                flex-wrap: wrap; 
                margin-right: -15px; 
                margin-left: -15px; 
            }
            .col-md-6 { 
                position: relative; 
                min-height: 1px; 
                padding-right: 15px; 
                padding-left: 15px; 
                flex: 0 0 50%; 
                max-width: 50%; 
            }
            .col-md-12 { 
                position: relative; 
                min-height: 1px; 
                padding-right: 15px; 
                padding-left: 15px; 
                flex: 0 0 100%; 
                max-width: 100%; 
            }
            @media (max-width: 768px) {
                .col-md-6 {
                    flex: 0 0 100%;
                    max-width: 100%;
                }
            }
            .form-group {
                margin-bottom: 15px;
            }
            .header {
                background-color: #3498db;
                color: #fff;
                padding: 15px;
                margin: -20px -20px 20px -20px;
                border-bottom: 3px solid #2980b9;
            }
            .header h2 {
                margin: 0;
                color: #fff;
            }
            .loading {
                display: none;
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background-color: rgba(255,255,255,0.8);
                z-index: 1000;
                text-align: center;
                padding-top: 20%;
                font-size: 18px;
            }
            .loading-spinner {
                border: 6px solid #f3f3f3;
                border-top: 6px solid #3498db;
                border-radius: 50%;
                width: 50px;
                height: 50px;
                animation: spin 2s linear infinite;
                margin: 0 auto 20px auto;
            }
            @keyframes spin {
                0% {{ transform: rotate(0deg); }}
                100% {{ transform: rotate(360deg); }}
            }
        </style>
    </head>
    <body>
    <div class="loading" id="loadingIndicator">
        <div class="loading-spinner"></div>
        <div>Loading...</div>
    </div>
    
    <!-- Fix: Make sure the header is visible -->
    <div style="background-color:#007bff; color:#fff; padding:10px; margin:-20px -20px 20px -20px;">
        <div style="display:flex; justify-content:space-between; align-items:center;">
            <div>
                <h2 style="margin:0; font-size:20px;">FastLane Check-In<svg xmlns="http://www.w3.org/2000/svg" viewBox="85 75 230 130" style="width: 60px; height: 30px; margin-left: -4px; vertical-align: middle;">
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
                    <text x="206" y="105" font-family="Arial, sans-serif" font-weight="bold" font-size="14" fill="#0099FF">si</text>
                  </svg></h2>
                <div style="font-size:12px; font-style:italic; margin-top:2px;">Life moves pretty fast. So do our lines!</div>
            </div>
        </div>
    </div>
    """
    
    # # Add debug checkbox to switch date for testing
    # print """
    # <div class="well">
    #     <form method="post" action="/PyScriptForm/{0}" onsubmit="showLoading()">
    #         <input type="hidden" name="step" value="choose_meetings">
    #         <div class="form-group">
    #             <label for="test_date">Select a different date:</label>
    #             <input type="date" class="form-control" id="test_date" name="test_date">
    #         </div>
    #         <button type="submit" class="btn btn-sm btn-default">Use Alternative Date</button>
    #     </form>
    # </div>
    # """.format(script_name)
    
    # if hasattr(check_in_manager.model.Data, 'test_date') and check_in_manager.model.Data.test_date:
    #     print """
    #     <div class="alert alert-info">
    #         <strong>Testing Mode:</strong> Using date {0} for meetings
    #     </div>
    #     """.format(check_in_manager.model.Data.test_date)
    
    if not meetings:
        print """
        <div class="alert alert-warning">
            <h4>No Meetings Today</h4>
            <p>There are no active meetings scheduled for today.</p>
            <p>Things to check:</p>
            <ul>
                <li>Make sure meetings are created for today's date</li>
                <li>Check that organizations hosting the meetings are Active</li>
                <li>Verify the meetings are not marked as "Did Not Meet"</li>
            </ul>
            <p>You can use the date picker above to test with another date where you know meetings exist.</p>
        </div>
        """
        print """
        </body>
        </html>
        """
        return
    
    # Add Program Filtering
    if programs:
        print """
        <div class="well">
            <form method="post" action="/PyScriptForm/{0}" onsubmit="showLoading()">
                <input type="hidden" name="step" value="choose_meetings">
                <div class="form-group">
                    <label for="program_filter">Filter by Program:</label>
                    <select class="form-control" id="program_filter" name="program_filter" onchange="this.form.submit()">
                        <option value="">All Programs</option>
        """.format(script_name)
        
        for program in programs:
            selected = ' selected="selected"' if program == selected_program else ''
            print '<option value="{0}"{1}>{0}</option>'.format(program, selected)
        
        print """
                    </select>
                </div>
            </form>
        </div>
        """
        
    # Continue with meeting selection display
    print """
    <div class="well">
        <form method="post" action="/PyScriptForm/{0}" onsubmit="showLoading()">
            <input type="hidden" name="step" value="check_in">
            
            <h4>Available Meetings for {1}</h4>
            
            <p>
                <button type="button" class="btn btn-sm btn-primary" id="selectAllBtn">Select All</button>
                <button type="button" class="btn btn-sm btn-default" id="deselectAllBtn">Deselect All</button>
            </p>
            
            <div class="row" style="max-height: 400px; overflow-y: auto;">
    """.format(script_name, check_in_manager.today.ToString(DATE_FORMAT))
    
    # Group meetings by organization for cleaner display
    org_groups = {}
    for meeting in meetings:
        # Apply program filter if set
        if selected_program and hasattr(meeting, 'program_name') and meeting.program_name != selected_program:
            continue
            
        if meeting.org_name not in org_groups:
            org_groups[meeting.org_name] = []
        org_groups[meeting.org_name].append(meeting)
    
    # Display meetings grouped by organization
    for org_name, org_meetings in sorted(org_groups.items()):
        print """
            <div class="col-md-6">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h4 class="panel-title">{0}</h4>
                    </div>
                    <div class="panel-body">
        """.format(org_name)
        
        for meeting in org_meetings:
            meeting_time = meeting.meeting_date.ToString("h:mm tt")
            print """
                        <div class="checkbox">
                            <label>
                                <input type="checkbox" name="meeting_id" value="{0}" class="meeting-checkbox">
                                {1} at {2} {3}
                            </label>
                        </div>
            """.format(
                meeting.meeting_id,
                org_name,
                meeting_time,
                "({0})".format(meeting.location) if meeting.location else ""
            )
        
        print """
                    </div>
                </div>
            </div>
        """
    
    print """
            </div>
            
            <div class="form-group" style="margin-top: 20px;">
                <button type="submit" class="btn btn-lg btn-success" id="start-checkin-btn">
                    <i class="fa fa-check-circle"></i> Start Check-In
                </button>
            </div>
        </form>
    </div>
    
    <script>
        // Show loading indicator during form submission
        function showLoading() {
            document.getElementById('loadingIndicator').style.display = 'block';
        }
        
        // Enable/disable the start button based on selections
        function updateStartButton() {
            var checkedCount = document.querySelectorAll('.meeting-checkbox:checked').length;
            document.getElementById('start-checkin-btn').disabled = checkedCount === 0;
        }
        
        // Set up event listeners after the DOM is loaded
        document.addEventListener('DOMContentLoaded', function() {
            // Set up checkbox listeners
            var checkboxes = document.querySelectorAll('.meeting-checkbox');
            checkboxes.forEach(function(checkbox) {
                checkbox.addEventListener('change', updateStartButton);
            });
            
            // Set up Select All button
            document.getElementById('selectAllBtn').addEventListener('click', function() {
                checkboxes.forEach(function(checkbox) {
                    checkbox.checked = true;
                });
                updateStartButton();
            });
            
            // Set up Deselect All button
            document.getElementById('deselectAllBtn').addEventListener('click', function() {
                checkboxes.forEach(function(checkbox) {
                    checkbox.checked = false;
                });
                updateStartButton();
            });
            
            // Initial check
            updateStartButton();
        });
    </script>
    </body>
    </html>
    """

def create_pagination(page, total_pages, script_name, meeting_ids, view_mode, alpha_filter, search_term):
    """Create a pagination control"""
    if total_pages <= 1:
        return ""
        
    pagination = ['<ul class="pagination">']
    
    # Previous button
    if page > 1:
        pagination.append(
            '<li><form method="post" action="/PyScriptForm/{0}" style="display:inline;"><input type="hidden" name="step" value="check_in"><input type="hidden" name="page" value="{1}"><input type="hidden" name="view_mode" value="{2}"><input type="hidden" name="alpha_filter" value="{3}"><input type="hidden" name="search_term" value="{4}">{5}<button type="submit" class="btn btn-link" style="margin:-6px;padding:6px 12px;text-decoration:none;" onclick="showLoading()">&laquo;</button></form></li>'.format(
                script_name, page - 1, view_mode, alpha_filter, search_term, 
                "\n".join(['<input type="hidden" name="meeting_id" value="{0}">'.format(m) for m in meeting_ids])
            )
        )
    else:
        pagination.append('<li class="disabled"><a href="#">&laquo;</a></li>')
        
    # Page numbers
    start_page = max(1, page - 2)
    end_page = min(total_pages, page + 2)
    
    # Always show page 1
    if start_page > 1:
        pagination.append(
            '<li><form method="post" action="/PyScriptForm/{0}" style="display:inline;"><input type="hidden" name="step" value="check_in"><input type="hidden" name="page" value="1"><input type="hidden" name="view_mode" value="{1}"><input type="hidden" name="alpha_filter" value="{2}"><input type="hidden" name="search_term" value="{3}">{4}<button type="submit" class="btn btn-link" style="margin:-6px;padding:6px 12px;text-decoration:none;" onclick="showLoading()">1</button></form></li>'.format(
                script_name, view_mode, alpha_filter, search_term, 
                "\n".join(['<input type="hidden" name="meeting_id" value="{0}">'.format(m) for m in meeting_ids])
            )
        )
        if start_page > 2:
            pagination.append('<li class="disabled"><a href="#">...</a></li>')
            
    # Page links
    for i in range(start_page, end_page + 1):
        if i == page:
            pagination.append('<li class="active"><a href="#">{0}</a></li>'.format(i))
        else:
            pagination.append(
                '<li><form method="post" action="/PyScriptForm/{0}" style="display:inline;"><input type="hidden" name="step" value="check_in"><input type="hidden" name="page" value="{1}"><input type="hidden" name="view_mode" value="{2}"><input type="hidden" name="alpha_filter" value="{3}"><input type="hidden" name="search_term" value="{4}">{5}<button type="submit" class="btn btn-link" style="margin:-6px;padding:6px 12px;text-decoration:none;" onclick="showLoading()">{1}</button></form></li>'.format(
                    script_name, i, view_mode, alpha_filter, search_term, 
                    "\n".join(['<input type="hidden" name="meeting_id" value="{0}">'.format(m) for m in meeting_ids])
                )
            )
            
    # Always show last page
    if end_page < total_pages:
        if end_page < total_pages - 1:
            pagination.append('<li class="disabled"><a href="#">...</a></li>')
        pagination.append(
            '<li><form method="post" action="/PyScriptForm/{0}" style="display:inline;"><input type="hidden" name="step" value="check_in"><input type="hidden" name="page" value="{1}"><input type="hidden" name="view_mode" value="{2}"><input type="hidden" name="alpha_filter" value="{3}"><input type="hidden" name="search_term" value="{4}">{5}<button type="submit" class="btn btn-link" style="margin:-6px;padding:6px 12px;text-decoration:none;" onclick="showLoading()">{1}</button></form></li>'.format(
                script_name, total_pages, view_mode, alpha_filter, search_term, 
                "\n".join(['<input type="hidden" name="meeting_id" value="{0}">'.format(m) for m in meeting_ids])
            )
        )
        
    # Next button
    if page < total_pages:
        pagination.append(
            '<li><form method="post" action="/PyScriptForm/{0}" style="display:inline;"><input type="hidden" name="step" value="check_in"><input type="hidden" name="page" value="{1}"><input type="hidden" name="view_mode" value="{2}"><input type="hidden" name="alpha_filter" value="{3}"><input type="hidden" name="search_term" value="{4}">{5}<button type="submit" class="btn btn-link" style="margin:-6px;padding:6px 12px;text-decoration:none;" onclick="showLoading()">&raquo;</button></form></li>'.format(
                script_name, page + 1, view_mode, alpha_filter, search_term, 
                "\n".join(['<input type="hidden" name="meeting_id" value="{0}">'.format(m) for m in meeting_ids])
            )
        )
    else:
        pagination.append('<li class="disabled"><a href="#">&raquo;</a></li>')
        
    pagination.append('</ul>')
    return "".join(pagination)


def render_fastlane_check_in(check_in_manager):
    """Render the FastLane check-in page with fixed alignment and branding"""
    # Get the current script name
    script_name = get_script_name()
    
    # Get parameters
    meeting_ids = []
    if hasattr(check_in_manager.model.Data, 'meeting_id'):
        if isinstance(check_in_manager.model.Data.meeting_id, list):
            meeting_ids = [str(m) for m in check_in_manager.model.Data.meeting_id]
        else:
            meeting_ids = [str(check_in_manager.model.Data.meeting_id)]
            
    # Get other parameters
    alpha_filter = getattr(check_in_manager.model.Data, 'alpha_filter', 'All')
    search_term = getattr(check_in_manager.model.Data, 'search_term', '')
    view_mode = getattr(check_in_manager.model.Data, 'view_mode', 'not_checked_in')
    
    # Handle page
    try:
        page_value = getattr(check_in_manager.model.Data, 'page', 1)
        current_page = int(page_value) if page_value else 1
    except (ValueError, TypeError):
        current_page = 1

    # Force a stats refresh - always get latest data
    # This ensures the counter is updated after check-in removals
    check_in_manager.last_check_in_time = check_in_manager.model.DateTime
    
    # Get check-in statistics with forced refresh
    stats = check_in_manager.get_check_in_stats(meeting_ids)
    
    # Create inputs for meeting IDs (for other forms) - MOVED UP
    meeting_id_inputs = ""
    for meeting_id in meeting_ids:
        meeting_id_inputs += '<input type="hidden" name="meeting_id" value="{0}">'.format(meeting_id)
    
    # Process form submission and prepare flash message
    success_message = ""
    flash_message = ""
    flash_name = ""
    if hasattr(check_in_manager.model.Data, 'action'):
        action = check_in_manager.model.Data.action
        
        if action == 'single_direct_check_in':
            # Get person ID and meeting ID directly
            person_id = getattr(check_in_manager.model.Data, 'person_id', None)
            meeting_id = getattr(check_in_manager.model.Data, 'meeting_id', None)
            person_name = getattr(check_in_manager.model.Data, 'person_name', "")
            
            if person_id and meeting_id:
                # Direct API call - simplified for reliability
                try:
                    # Force refresh of statistics
                    check_in_manager.last_check_in_time = check_in_manager.model.DateTime
                    
                    result = check_in_manager.model.EditPersonAttendance(int(meeting_id), int(person_id), True)
                    success = "Success" in str(result)
                    
                    if success:
                        flash_message = "Successfully checked in"
                        flash_name = person_name
                except Exception as e:
                    flash_message = "Error checking in: {0}".format(str(e))
        
        elif action == 'remove_check_in' and hasattr(check_in_manager.model.Data, 'person_id'):
            # Single person check-in removal
            person_id = check_in_manager.model.Data.person_id
            meeting_id = getattr(check_in_manager.model.Data, 'meeting_id', None)
            person_name = getattr(check_in_manager.model.Data, 'person_name', "")
            
            # Direct API call
            if person_id and meeting_id:
                try:
                    # Force refresh of statistics
                    check_in_manager.last_check_in_time = check_in_manager.model.DateTime
                    
                    result = check_in_manager.model.EditPersonAttendance(int(meeting_id), int(person_id), False)
                    success = "Success" in str(result)
                    
                    if success:
                        flash_message = "Successfully removed check-in for"
                        flash_name = person_name
                except Exception as e:
                    flash_message = "Error removing check-in: {0}".format(str(e))
    
    # Get people based on view mode
    people = []
    total_count = 0
    
    if view_mode == "checked_in":
        people, total_count = check_in_manager.get_checked_in_people(meeting_ids, current_page, PAGE_SIZE)
    else:
        people, total_count = check_in_manager.get_people_by_filter(alpha_filter, search_term, meeting_ids, current_page, PAGE_SIZE)
    
    # Calculate pagination
    total_pages = (total_count + PAGE_SIZE - 1) // PAGE_SIZE
    
    # Create pagination controls
    pagination = create_pagination(current_page, total_pages, script_name, meeting_ids, view_mode, alpha_filter, search_term)
    
    # Render alpha filters - compact version
    alpha_filters_html = """
    <div style="margin-bottom:10px;">
        <div style="font-weight:bold; margin-bottom:5px;">Filter:</div>
        <div>
    """
    
    for alpha_block in ALPHA_BLOCKS:
        is_active = alpha_block == alpha_filter
        color = "#337ab7" if is_active else "#6c757d"
        bg_color = "#e6f2fa" if is_active else "#f8f9fa"
        
        alpha_filters_html += """
        <form method="post" action="/PyScriptForm/{0}" style="display:inline-block; margin:2px;" onsubmit="document.getElementById('loadingIndicator').style.display='block';">
            <input type="hidden" name="step" value="check_in">
            <input type="hidden" name="view_mode" value="{1}">
            <input type="hidden" name="alpha_filter" value="{2}">
            <input type="hidden" name="search_term" value="{3}">
            {4}
            <button type="submit" style="padding:4px 8px; font-size:12px; color:{5}; background-color:{6}; border:1px solid #dee2e6; border-radius:3px; cursor:pointer; min-width:42px;">{2}</button>
        </form>
        """.format(
            script_name,        # {0}
            view_mode,          # {1}
            alpha_block,        # {2}
            search_term,        # {3}
            meeting_id_inputs,  # {4}
            color,              # {5}
            bg_color            # {6}
        )
    
    alpha_filters_html += """
        </div>
    </div>
    """
    
    # Generate the people list HTML - more compact layout with inline check buttons
    people_list_html = []
    
    # Determine number of columns based on viewport - updated for better fit
    people_list_html.append("""
    <div class="people-grid" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(280px, 1fr)); gap:10px; margin-bottom:20px;">
    """)
    
    for person in people:
        # Use the primary meeting ID for simplicity
        meeting_id = meeting_ids[0] if meeting_ids else ""
        
        if view_mode == "checked_in":
            # For checked-in people, show "remove" button with improved alignment
            item_html = """
            <div style="padding:8px; background-color:#fff; border:1px solid #ddd; border-radius:4px; display:flex; align-items:center;">
                <div style="flex:1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-size:15px;">{1}</div>
                <form method="post" action="/PyScriptForm/{2}" style="margin:0; padding:0; flex-shrink:0;" onsubmit="document.getElementById('loadingIndicator').style.display='block';">
                    <input type="hidden" name="step" value="check_in">
                    <input type="hidden" name="action" value="remove_check_in">
                    <input type="hidden" name="person_id" value="{0}">
                    <input type="hidden" name="person_name" value="{1}">
                    <input type="hidden" name="meeting_id" value="{3}">
                    <input type="hidden" name="view_mode" value="checked_in">
                    <input type="hidden" name="alpha_filter" value="{4}">
                    <input type="hidden" name="search_term" value="{5}">
                    <input type="hidden" name="page" value="{6}">
                    <button type="submit" style="padding:3px 8px; font-size:12px; background-color:#dc3545; color:#fff; border:1px solid #dc3545; border-radius:3px; cursor:pointer; min-width:60px; height:28px; vertical-align:middle;">Remove</button>
                </form>
            </div>
            """.format(
                person.people_id,  # {0}
                person.name,       # {1}
                script_name,       # {2}
                meeting_id,        # {3}
                alpha_filter,      # {4}
                search_term,       # {5}
                current_page       # {6}
            )
        else:
            # For not-checked-in people, show "check in" button with improved alignment
            item_html = """
            <div style="padding:8px; background-color:#fff; border:1px solid #ddd; border-radius:4px; display:flex; align-items:center;">
                <div style="flex:1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-size:15px;">{1}</div>
                <form method="post" action="/PyScriptForm/{2}" style="margin:0; padding:0; flex-shrink:0;" onsubmit="document.getElementById('loadingIndicator').style.display='block';">
                    <input type="hidden" name="step" value="check_in">
                    <input type="hidden" name="action" value="single_direct_check_in">
                    <input type="hidden" name="person_id" value="{0}">
                    <input type="hidden" name="person_name" value="{1}">
                    <input type="hidden" name="meeting_id" value="{3}">
                    <input type="hidden" name="view_mode" value="not_checked_in">
                    <input type="hidden" name="alpha_filter" value="{4}">
                    <input type="hidden" name="search_term" value="{5}">
                    <input type="hidden" name="page" value="{6}">
                    <button type="submit" style="padding:3px 8px; font-size:12px; background-color:#28a745; color:#fff; border:1px solid #28a745; border-radius:3px; cursor:pointer; min-width:60px; height:28px; vertical-align:middle;">Check In</button>
                </form>
            </div>
            """.format(
                person.people_id,  # {0}
                person.name,       # {1}
                script_name,       # {2}
                meeting_id,        # {3}
                alpha_filter,      # {4}
                search_term,       # {5}
                current_page       # {6}
            )
        
        people_list_html.append(item_html)
    
    people_list_html.append("</div>")
    
    # Combine all list items
    people_list_html_str = "\n".join(people_list_html)
    
    # Create the compact stats bar
    stats_bar = """
    <div style="display:flex; justify-content:space-between; background-color:#f8f9fa; border:1px solid #dee2e6; border-radius:4px; padding:8px; margin-bottom:10px; font-size:13px;">
        <div><strong style="color:#28a745;">{0}</strong> Checked In</div>
        <div><strong style="color:#007bff;">{1}</strong> Not Checked In</div>
        <div><strong>{2}</strong> Total</div>
    </div>
    """.format(stats["checked_in"], stats["not_checked_in"], stats["total"])
    
    # Create flash message for 2-second display
    flash_html = ""
    if flash_message and flash_name:
        flash_html = """
        <div id="flash-message" style="position:fixed; top:60px; left:50%; transform:translateX(-50%); padding:10px 20px; background-color:rgba(40,167,69,0.9); color:white; border-radius:4px; z-index:9999; box-shadow:0 2px 10px rgba(0,0,0,0.2);">
            <strong>{0}</strong> {1}
        </div>
        <script>
            setTimeout(function() {{
                var flashMessage = document.getElementById('flash-message');
                if (flashMessage) {{
                    flashMessage.style.opacity = '0';
                    flashMessage.style.transition = 'opacity 0.5s ease';
                    setTimeout(function() {{
                        flashMessage.style.display = 'none';
                    }}, 500);
                }}
            }}, 2000);
        </script>
        """.format(flash_name, flash_message)
    
    # Simple HTML with FastLane branding and compact layout
    print """<!DOCTYPE html>
<html>
<head>
    <title>FastLane Check-In</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="font-family:Arial,sans-serif; margin:0; padding:10px; color:#333; background-color:#f5f5f5;">
    <div id="loadingIndicator" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background-color:rgba(255,255,255,0.8); z-index:1000; text-align:center; padding-top:20%;">
        <div style="border:4px solid #f3f3f3; border-top:4px solid #3498db; border-radius:50%; width:40px; height:40px; margin:0 auto 15px auto;"></div>
        <div>Processing...</div>
    </div>
    
    <!-- Flash message for check-in confirmation -->
    {0}
    
    <!-- Updated FastLane header with tagline -->
    <div style="background-color:#007bff; color:#fff; padding:10px; margin:-10px -10px 10px -10px;">
        <div style="display:flex; justify-content:space-between; align-items:center;">
            <div>
                <h2 style="margin:0; font-size:20px;">FastLane Check-In<svg xmlns="http://www.w3.org/2000/svg" viewBox="85 75 230 130" style="width: 60px; height: 30px; margin-left: -4px; vertical-align: middle;">
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
                    <text x="206" y="105" font-family="Arial, sans-serif" font-weight="bold" font-size="14" fill="#0099FF">si</text>
                  </svg></h2>
                <div style="font-size:12px; font-style:italic; margin-top:2px;">Life moves pretty fast. So do our lines!</div>
            </div>
            <a href="/PyScript/{1}?step=choose_meetings" style="color:white; text-decoration:none; background-color:rgba(255,255,255,0.2); padding:5px 8px; border-radius:3px; font-size:13px;">Back</a>
        </div>
    </div>
    
    <!-- Compact stats bar -->
    {2}
    
    <!-- Main content -->
    <div style="background-color:#fff; border:1px solid #ddd; border-radius:4px; overflow:hidden; margin-bottom:10px;">
        <!-- New header with toggle in the title bar -->
        <div style="padding:8px 10px; background-color:#f8f9fa; border-bottom:1px solid #ddd; display:flex; justify-content:space-between; align-items:center;">
            <h3 style="margin:0; font-size:16px; flex-grow:1;">{3}</h3>
            
            <!-- New toggle switch for view mode instead of radio buttons -->
            <div style="display:flex; align-items:center;">
                <form method="post" action="/PyScriptForm/{1}" style="margin:0; padding:0;" onsubmit="document.getElementById('loadingIndicator').style.display='block';">
                    <input type="hidden" name="step" value="check_in">
                    <input type="hidden" name="view_mode" value="{7}">
                    <input type="hidden" name="alpha_filter" value="{4}">
                    <input type="hidden" name="search_term" value="{5}">
                    {6}
                    <button type="submit" style="padding:4px 8px; font-size:12px; color:#fff; background-color:#007bff; border:1px solid #007bff; border-radius:3px;">
                        Switch to {8}
                    </button>
                </form>
            </div>
        </div>
        
        <div style="padding:10px;">
            <!-- Search box -->
            <form method="post" action="/PyScriptForm/{1}" style="margin-bottom:10px;" onsubmit="document.getElementById('loadingIndicator').style.display='block';">
                <input type="hidden" name="step" value="check_in">
                <input type="hidden" name="view_mode" value="{3}">
                {6}
                <div style="display:flex;">
                    <input type="text" name="search_term" value="{5}" placeholder="Search by name..." style="flex-grow:1; height:36px; padding:4px 10px; border:1px solid #ddd; border-radius:4px 0 0 4px; font-size:14px;">
                    <button type="submit" style="height:36px; padding:0 12px; background-color:#007bff; color:white; border:1px solid #007bff; border-radius:0 4px 4px 0; cursor:pointer; white-space:nowrap;" onclick="document.getElementById('loadingIndicator').style.display='block';">
                        Search
                    </button>
                </div>
            </form>
            
            <!-- Alpha filters - compact -->
            {9}
            
            <!-- People grid layout with fixed alignment -->
            {10}
            
            <!-- Pagination -->
            <div style="text-align:center;">
                {11}
            </div>
        </div>
    </div>
    
    <script>
        function showLoading() {{
            document.getElementById('loadingIndicator').style.display = 'block';
        }}
    </script>
</body>
</html>
""".format(
        flash_html,               # {0}
        script_name,              # {1}
        stats_bar,                # {2}
        "People to Check In" if view_mode == "not_checked_in" else "Checked In People", # {3}
        alpha_filter,             # {4}
        search_term,              # {5}
        meeting_id_inputs,        # {6}
        "checked_in" if view_mode == "not_checked_in" else "not_checked_in", # {7} - opposite mode for toggle
        "Checked In List" if view_mode == "not_checked_in" else "Check-In Page", # {8} - toggle button text
        alpha_filters_html,       # {9}
        people_list_html_str,     # {10}
        pagination                # {11}
    )
    
    return True  # Indicate successful rendering

def render_checked_in_view(check_in_manager):
    """Render the checked-in people view"""
    # Similar to render_check_in_page but displays already checked-in people
    # This allows removing individual check-ins if needed
    return render_check_in_page(check_in_manager)
    
def render_family_check_in(check_in_manager):
    """Render the family check-in page"""
    # Get the current script name
    script_name = get_script_name()
    
    # Get parameters
    meeting_ids = []
    if hasattr(check_in_manager.model.Data, 'meeting_id'):
        if isinstance(check_in_manager.model.Data.meeting_id, list):
            meeting_ids = [str(m) for m in check_in_manager.model.Data.meeting_id]
        else:
            meeting_ids = [str(check_in_manager.model.Data.meeting_id)]
            
    person_id = getattr(check_in_manager.model.Data, 'person_id', None)
    family_id = getattr(check_in_manager.model.Data, 'family_id', None)
    
    if not person_id:
        # Redirect back to check-in page if no person selected
        return render_check_in_page(check_in_manager)
        
    # Set the selected person ID for use in family member query
    check_in_manager.model.Data.selected_person_id = person_id
    
    # Get person info
    if hasattr(check_in_manager.model.Data, 'org_ids'):
        org_ids = check_in_manager.model.Data.org_ids.split(',')
    else:
        org_ids = []
        
    # Process form submission for family check-in
    success_message = ""
    if hasattr(check_in_manager.model.Data, 'action') and check_in_manager.model.Data.action == 'family_check_in_submit':
        # Get family member IDs to check in
        family_member_ids = []
        if hasattr(check_in_manager.model.Data, 'family_member_id'):
            if isinstance(check_in_manager.model.Data.family_member_id, list):
                family_member_ids = [str(m) for m in check_in_manager.model.Data.family_member_id]
            else:
                family_member_ids = [str(check_in_manager.model.Data.family_member_id)]
                
        # Also check in the main person
        if getattr(check_in_manager.model.Data, 'check_in_person', False):
            family_member_ids.append(person_id)
            
        # Do the check-in
        if family_member_ids:
            success = check_in_manager.bulk_check_in(family_member_ids, meeting_ids)
            if success:
                success_message = """
                <div class="alert alert-success">
                    <strong>Success!</strong> {0} family members have been checked in.
                </div>
                """.format(len(family_member_ids))
                
                # Redirect back to check-in page
                return render_check_in_page(check_in_manager)
        
    # Get family members
    family_members = []
    if family_id:
        family_members = check_in_manager.get_family_members(family_id, org_ids)
        
    # Get person name
    person_name = ""
    sql = "SELECT Name FROM People WHERE PeopleId = {0}".format(person_id)
    result = check_in_manager.q.QuerySqlTop1(sql)
    if result and hasattr(result, 'Name'):
        person_name = result.Name
        
    # Start the page with modern styling
    print """<!DOCTYPE html>
    <html>
    <head>
        <title>Rapid Check-In System - Family Check-In</title>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            body { 
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                margin: 0; 
                padding: 20px;
                color: #333;
                background-color: #f9f9f9;
            }
            h2 { 
                margin-top: 0;
                color: #2c3e50;
                font-weight: 500;
            }
            hr {
                border: 0;
                height: 1px;
                background-color: #e0e0e0;
                margin: 15px 0;
            }
            .btn { 
                display: inline-block; 
                padding: 8px 16px; 
                margin-bottom: 0; 
                font-size: 14px; 
                font-weight: 400; 
                line-height: 1.42857143; 
                text-align: center; 
                white-space: nowrap; 
                vertical-align: middle; 
                cursor: pointer; 
                border: 1px solid transparent; 
                border-radius: 4px;
                transition: all 0.3s;
            }
            .btn-primary { color: #fff; background-color: #3498db; border-color: #2980b9; }
            .btn-primary:hover { background-color: #2980b9; }
            .btn-default { color: #333; background-color: #fff; border-color: #ccc; }
            .btn-default:hover { background-color: #f5f5f5; }
            .btn-success { color: #fff; background-color: #2ecc71; border-color: #27ae60; }
            .btn-success:hover { background-color: #27ae60; }
            .btn-lg { padding: 12px 20px; font-size: 16px; }
            .btn-sm { padding: 5px 10px; font-size: 12px; }
            .well { 
                min-height: 20px; 
                padding: 19px; 
                margin-bottom: 20px; 
                background-color: #fff;
                border: 1px solid #e3e3e3; 
                border-radius: 4px; 
                box-shadow: 0 1px 3px rgba(0,0,0,.05);
            }
            .checkbox { 
                position: relative; 
                display: block; 
                margin-top: 10px; 
                margin-bottom: 10px; 
            }
            .checkbox label {
                padding-left: 25px;
                cursor: pointer;
            }
            .checkbox input[type=checkbox] {
                position: absolute;
                margin-left: -20px;
                margin-top: 2px;
                cursor: pointer;
            }
            .panel { 
                margin-bottom: 20px; 
                background-color: #fff; 
                border: 1px solid #ddd; 
                border-radius: 4px;
                box-shadow: 0 1px 3px rgba(0,0,0,.05);
            }
            .panel-heading { 
                padding: 12px 15px; 
                background-color: #f7f7f7;
                border-bottom: 1px solid #ddd; 
                border-top-left-radius: 3px;
                border-top-right-radius: 3px; 
            }
            .panel-body { padding: 15px; }
            .panel-title { 
                margin-top: 0; 
                margin-bottom: 0; 
                font-size: 16px; 
                color: #333;
                font-weight: 500;
            }
            .alert { 
                padding: 15px; 
                margin-bottom: 20px; 
                border: 1px solid transparent; 
                border-radius: 4px; 
            }
            .alert-warning { color: #8a6d3b; background-color: #fcf8e3; border-color: #faebcc; }
            .alert-info { color: #31708f; background-color: #d9edf7; border-color: #bce8f1; }
            .alert-success { color: #3c763d; background-color: #dff0d8; border-color: #d6e9c6; }
            .row { 
                display: flex; 
                flex-wrap: wrap; 
                margin-right: -15px; 
                margin-left: -15px; 
            }
            .col-md-6 { 
                position: relative; 
                min-height: 1px; 
                padding-right: 15px; 
                padding-left: 15px; 
                flex: 0 0 50%; 
                max-width: 50%; 
            }
            .col-md-12 { 
                position: relative; 
                min-height: 1px; 
                padding-right: 15px; 
                padding-left: 15px; 
                flex: 0 0 100%; 
                max-width: 100%; 
            }
            @media (max-width: 768px) {
                .col-md-6 {
                    flex: 0 0 100%;
                    max-width: 100%;
                }
            }
            .form-group {
                margin-bottom: 15px;
            }
            .header {
                background-color: #3498db;
                color: #fff;
                padding: 15px;
                margin: -20px -20px 20px -20px;
                border-bottom: 3px solid #2980b9;
            }
            .header h2 {
                margin: 0;
                color: #fff;
            }
            .loading {
                display: none;
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background-color: rgba(255,255,255,0.8);
                z-index: 1000;
                text-align: center;
                padding-top: 20%;
                font-size: 18px;
            }
            .loading-spinner {
                border: 6px solid #f3f3f3;
                border-top: 6px solid #3498db;
                border-radius: 50%;
                width: 50px;
                height: 50px;
                animation: spin 2s linear infinite;
                margin: 0 auto 20px auto;
            }
            @keyframes spin {
                0% {{ transform: rotate(0deg); }}
                100% {{ transform: rotate(360deg); }}
            }
        </style>
    </head>
    <body>
    <div class="loading" id="loadingIndicator">
        <div class="loading-spinner"></div>
        <div>Processing...</div>
    </div>
    <div class="header">
        <h2>Rapid Check-In System</h2>
    </div>
    """
    
    # Create form for hidden fields
    meeting_id_inputs = ""
    for meeting_id in meeting_ids:
        meeting_id_inputs += '<input type="hidden" name="meeting_id" value="{0}">'.format(meeting_id)
        
    # Display family check-in form
    print """
    <div class="well">
        <form method="post" action="/PyScriptForm/{0}" onsubmit="showLoading()">
            <input type="hidden" name="step" value="check_in">
            <input type="hidden" name="action" value="family_check_in_submit">
            <input type="hidden" name="person_id" value="{1}">
            <input type="hidden" name="family_id" value="{2}">
            {3}
            
            <h3>Family Check-In for {4}</h3>
            
            <div class="checkbox">
                <label>
                    <input type="checkbox" name="check_in_person" value="1" checked>
                    Check in {4}
                </label>
            </div>
    """.format(script_name, person_id, family_id, meeting_id_inputs, person_name)
    
    if family_members:
        print "<hr>"
        print "<h4>Family Members</h4>"
        print "<p>Select family members to check in:</p>"
        
        for member in family_members:
            # Create list of orgs this person belongs to
            org_list = []
            for org_id in member.org_ids:
                org_name = get_org_name(check_in_manager, org_id)
                org_list.append(org_name)
                
            org_text = ""
            if org_list:
                org_text = "<small class='text-muted'>({0})</small>".format(", ".join(org_list))
                
            print """
            <div class="checkbox">
                <label>
                    <input type="checkbox" name="family_member_id" value="{0}" checked>
                    {1} {2}
                </label>
            </div>
            """.format(member.people_id, member.name, org_text)
    else:
        print """
        <div class="alert alert-info">
            <p>No other family members found that can be checked in.</p>
        </div>
        """
        
    print """
            <hr>
            
            <div class="row">
                <div class="col-md-6">
                    <a href="/PyScriptForm/{0}?step=check_in&{1}" class="btn btn-default" onclick="showLoading()">
                        <i class="fa fa-arrow-left"></i> Back to Check-In
                    </a>
                </div>
                <div class="col-md-6 text-right">
                    <button type="submit" class="btn btn-success" onclick="showLoading()">
                        <i class="fa fa-check-circle"></i> Complete Check-In
                    </button>
                </div>
            </div>
        </form>
    </div>
    
    <script>
        // Show loading indicator during form submission
        function showLoading() {
            document.getElementById('loadingIndicator').style.display = 'block';
        }
    </script>
    </body>
    </html>
    """.format(script_name, "&".join(["meeting_id={0}".format(m) for m in meeting_ids]))
    

# HTML template 
def render_simplified_check_in_page(check_in_manager):
    """Render a simplified check-in page with direct check-in buttons and no selection checkboxes"""
    # Get the current script name
    script_name = get_script_name()
    
    # Get parameters
    meeting_ids = []
    if hasattr(check_in_manager.model.Data, 'meeting_id'):
        if isinstance(check_in_manager.model.Data.meeting_id, list):
            meeting_ids = [str(m) for m in check_in_manager.model.Data.meeting_id]
        else:
            meeting_ids = [str(check_in_manager.model.Data.meeting_id)]
            
    # Get other parameters
    alpha_filter = getattr(check_in_manager.model.Data, 'alpha_filter', 'All')
    search_term = getattr(check_in_manager.model.Data, 'search_term', '')
    view_mode = getattr(check_in_manager.model.Data, 'view_mode', 'not_checked_in')
    
    # Handle page
    try:
        page_value = getattr(check_in_manager.model.Data, 'page', 1)
        current_page = int(page_value) if page_value else 1
    except (ValueError, TypeError):
        current_page = 1

    # Get check-in statistics
    stats = check_in_manager.get_check_in_stats(meeting_ids)
    
    # Process form submission
    success_message = ""
    if hasattr(check_in_manager.model.Data, 'action'):
        action = check_in_manager.model.Data.action
        
        if action == 'single_direct_check_in':
            # Get person ID and meeting ID directly
            person_id = getattr(check_in_manager.model.Data, 'person_id', None)
            meeting_id = getattr(check_in_manager.model.Data, 'meeting_id', None)
            
            if person_id and meeting_id:
                # Direct API call - simplified for reliability
                try:
                    result = check_in_manager.model.EditPersonAttendance(int(meeting_id), int(person_id), True)
                    success = "Success" in str(result)
                    
                    if success:
                        success_message = """
                        <div class="alert alert-success">
                            <strong>Success!</strong> Person has been checked in.
                        </div>
                        """
                    else:
                        success_message = """
                        <div class="alert alert-danger">
                            <strong>Error!</strong> Could not check in the person. Result: {0}
                        </div>
                        """.format(str(result))
                except Exception as e:
                    success_message = """
                    <div class="alert alert-danger">
                        <strong>Error!</strong> Exception: {0}
                    </div>
                    """.format(str(e))
        
        elif action == 'remove_check_in' and hasattr(check_in_manager.model.Data, 'person_id'):
            # Single person check-in removal
            person_id = check_in_manager.model.Data.person_id
            meeting_id = getattr(check_in_manager.model.Data, 'meeting_id', None)
            
            # Direct API call
            if person_id and meeting_id:
                try:
                    result = check_in_manager.model.EditPersonAttendance(int(meeting_id), int(person_id), False)
                    success = "Success" in str(result)
                    
                    if success:
                        success_message = """
                        <div class="alert alert-success">
                            <strong>Success!</strong> Check-in has been removed.
                        </div>
                        """
                except Exception:
                    success_message = """
                    <div class="alert alert-danger">
                        <strong>Error!</strong> Could not remove check-in.
                    </div>
                    """
    
    # Get people based on view mode
    people = []
    total_count = 0
    
    if view_mode == "checked_in":
        people, total_count = check_in_manager.get_checked_in_people(meeting_ids, current_page, PAGE_SIZE)
    else:
        people, total_count = check_in_manager.get_people_by_filter(alpha_filter, search_term, meeting_ids, current_page, PAGE_SIZE)
    
    # Calculate pagination
    total_pages = (total_count + PAGE_SIZE - 1) // PAGE_SIZE
    
    # Create pagination controls
    pagination = create_pagination(current_page, total_pages, script_name, meeting_ids, view_mode, alpha_filter, search_term)
    
    # Render alpha filters
    alpha_filters_html = render_alpha_filters(alpha_filter)
    
    # Generate the people list HTML based on view mode
    if view_mode == "checked_in":
        # For checked-in people, show "remove" button
        people_list_html = []
        for person in people:
            # Use the primary meeting ID for simplicity
            meeting_id = meeting_ids[0] if meeting_ids else ""
            
            item_html = """
            <div class="list-group-item">
                <div class="row">
                    <div class="col-xs-8">
                        <span class="person-name">{1}</span>
                    </div>
                    <div class="col-xs-4 text-right">
                        <form method="post" action="/PyScriptForm/{2}" style="display:inline;" onsubmit="showLoading()">
                            <input type="hidden" name="step" value="check_in">
                            <input type="hidden" name="action" value="remove_check_in">
                            <input type="hidden" name="person_id" value="{0}">
                            <input type="hidden" name="meeting_id" value="{3}">
                            <input type="hidden" name="view_mode" value="checked_in">
                            <input type="hidden" name="alpha_filter" value="{4}">
                            <input type="hidden" name="search_term" value="{5}">
                            <input type="hidden" name="page" value="{6}">
                            <button type="submit" class="btn btn-xs btn-danger" onclick="showLoading()">
                                <i class="fa fa-times"></i> Remove
                            </button>
                        </form>
                    </div>
                </div>
            </div>
            """.format(
                person.people_id,  # {0}
                person.name,       # {1}
                script_name,       # {2}
                meeting_id,        # {3}
                alpha_filter,      # {4}
                search_term,       # {5}
                current_page       # {6}
            )
            people_list_html.append(item_html)
    else:
        # For not-checked-in people, show "check in" button
        people_list_html = []
        for person in people:
            # Use the primary meeting ID for simplicity
            meeting_id = meeting_ids[0] if meeting_ids else ""
            
            item_html = """
            <div style="padding:8px; background-color:#fff; border:1px solid #ddd; border-radius:4px; display:flex; align-items:center;">
                <div style="flex:1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-size:15px;">{1}</div>
                <form method="post" action="/PyScriptForm/{2}" style="margin:0; padding:0; flex-shrink:0;" onsubmit="document.getElementById('loadingIndicator').style.display='block';">
                    <input type="hidden" name="step" value="check_in">
                    <input type="hidden" name="action" value="single_direct_check_in">
                    <input type="hidden" name="person_id" value="{0}">
                    <input type="hidden" name="person_name" value="{1}">
                    <input type="hidden" name="meeting_id" value="{3}">
                    <input type="hidden" name="view_mode" value="not_checked_in">
                    <input type="hidden" name="alpha_filter" value="{4}">
                    <input type="hidden" name="search_term" value="{5}">
                    <input type="hidden" name="page" value="{6}">
                    <button type="submit" style="padding:3px 8px; font-size:12px; background-color:#28a745; color:#fff; border:1px solid #28a745; border-radius:3px; cursor:pointer; min-width:60px; height:28px; vertical-align:middle;">Check In</button>
                </form>
            </div>
            """.format(
                person.people_id,  # {0}
                person.name,       # {1}
                script_name,       # {2}
                meeting_id,        # {3}
                alpha_filter,      # {4}
                search_term,       # {5}
                current_page       # {6}
            )
            people_list_html.append(item_html)
    
    # Combine all list items
    people_list_html_str = "\n".join(people_list_html)
    
    # Create inputs for meeting IDs (for other forms)
    meeting_id_inputs = ""
    for meeting_id in meeting_ids:
        meeting_id_inputs += '<input type="hidden" name="meeting_id" value="{0}">'.format(meeting_id)
    
    # HTML for the page - FIXED CSS
    print """<!DOCTYPE html>
    <html>
    <head>
        <title>Rapid Check-In System</title>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            body { 
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                margin: 0; 
                padding: 20px;
                color: #333;
                background-color: #f9f9f9;
            }
            h2 { 
                margin-top: 0;
                color: #2c3e50;
                font-weight: 500;
            }
            .btn { 
                display: inline-block; 
                padding: 8px 16px; 
                margin-bottom: 0; 
                font-size: 14px; 
                font-weight: 400; 
                line-height: 1.42857143; 
                text-align: center; 
                white-space: nowrap; 
                vertical-align: middle; 
                cursor: pointer; 
                border: 1px solid transparent; 
                border-radius: 4px;
            }
            .btn-xs {
                padding: 1px 5px;
                font-size: 12px;
                line-height: 1.5;
                border-radius: 3px;
            }
            .btn-primary { color: #fff; background-color: #3498db; border-color: #2980b9; }
            .btn-default { color: #333; background-color: #fff; border-color: #ccc; }
            .btn-success { color: #fff; background-color: #2ecc71; border-color: #27ae60; }
            .btn-danger { color: #fff; background-color: #e74c3c; border-color: #c0392b; }
            .person-name {
                font-size: 16px;
                line-height: 28px;
            }
        </style>
    </head>
    <body>
    <div class="loading" id="loadingIndicator" style="display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background-color: rgba(255,255,255,0.8); z-index: 1000; text-align: center; padding-top: 20%;">
        <div class="loading-spinner" style="border: 6px solid #f3f3f3; border-top: 6px solid #3498db; border-radius: 50%; width: 50px; height: 50px; animation: spin 2s linear infinite; margin: 0 auto 20px auto;"></div>
        <div>Processing...</div>
    </div>
    
    <div class="header" style="background-color: #3498db; color: #fff; padding: 15px; margin: -20px -20px 20px -20px; border-bottom: 3px solid #2980b9;">
        <h2 style="margin: 0; color: #fff;">Rapid Check-In System</h2>
    </div>
    
    <!-- Success message -->
    {0}
    
    <!-- Stats boxes -->
    <div class="row" style="display: flex; flex-wrap: wrap; margin-right: -15px; margin-left: -15px;">
        <div class="col-xs-4" style="position: relative; min-height: 1px; padding-right: 15px; padding-left: 15px; flex: 0 0 33.33333%; max-width: 33.33333%;">
            <div class="stats-box checked-in-box" style="padding: 15px; text-align: center; background-color: #d9f2e6; border-left: 4px solid #2ecc71; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,.05); margin-bottom: 15px;">
                <div class="number" style="font-size: 24px; font-weight: bold; margin-bottom: 5px;">{1}</div>
                <div class="label" style="font-size: 14px; color: #888;">Checked In</div>
            </div>
        </div>
        <div class="col-xs-4" style="position: relative; min-height: 1px; padding-right: 15px; padding-left: 15px; flex: 0 0 33.33333%; max-width: 33.33333%;">
            <div class="stats-box not-checked-in-box" style="padding: 15px; text-align: center; background-color: #e6f2fa; border-left: 4px solid #3498db; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,.05); margin-bottom: 15px;">
                <div class="number" style="font-size: 24px; font-weight: bold; margin-bottom: 5px;">{2}</div>
                <div class="label" style="font-size: 14px; color: #888;">Not Checked In</div>
            </div>
        </div>
        <div class="col-xs-4" style="position: relative; min-height: 1px; padding-right: 15px; padding-left: 15px; flex: 0 0 33.33333%; max-width: 33.33333%;">
            <div class="stats-box total-box" style="padding: 15px; text-align: center; background-color: #f9f9f9; border-left: 4px solid #95a5a6; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,.05); margin-bottom: 15px;">
                <div class="number" style="font-size: 24px; font-weight: bold; margin-bottom: 5px;">{3}</div>
                <div class="label" style="font-size: 14px; color: #888;">Total</div>
            </div>
        </div>
    </div>

    <!-- Main content -->
    <div class="panel panel-default" style="margin-bottom: 20px; background-color: #fff; border: 1px solid #ddd; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,.05);">
        <div class="panel-heading" style="padding: 12px 15px; background-color: #f7f7f7; border-bottom: 1px solid #ddd; border-top-left-radius: 3px; border-top-right-radius: 3px;">
            <div class="row" style="display: flex; flex-wrap: wrap; margin-right: -15px; margin-left: -15px;">
                <div class="col-xs-8" style="position: relative; min-height: 1px; padding-right: 15px; padding-left: 15px; flex: 0 0 66.66666%; max-width: 66.66666%;">
                    <h3 class="panel-title" style="margin-top: 0; margin-bottom: 0; font-size: 16px; color: #333; font-weight: 500;">{4}</h3>
                </div>
                <div class="col-xs-4 text-right" style="position: relative; min-height: 1px; padding-right: 15px; padding-left: 15px; flex: 0 0 33.33333%; max-width: 33.33333%; text-align: right;">
                    <a href="/PyScript/{5}?step=choose_meetings" class="btn btn-sm btn-default" onclick="showLoading()" style="padding: 5px 10px; font-size: 12px;">
                        <i class="fa fa-arrow-left"></i> Back to Meetings
                    </a>
                </div>
            </div>
        </div>
        
        <div class="panel-body" style="padding: 15px;">
            <!-- View toggle buttons -->
            <div class="row" style="display: flex; flex-wrap: wrap; margin-right: -15px; margin-left: -15px;">
                <div class="col-xs-12" style="position: relative; min-height: 1px; padding-right: 15px; padding-left: 15px; flex: 0 0 100%; max-width: 100%;">
                    <form method="post" action="/PyScriptForm/{5}" class="form-inline" onsubmit="showLoading()">
                        <input type="hidden" name="step" value="check_in">
                        {6}
                        <div class="btn-group" data-toggle="buttons" style="margin-bottom: 15px; position: relative; display: inline-block; vertical-align: middle;">
                            <label class="btn btn-default {7}" style="position: relative; float: left;">
                                <input type="radio" name="view_mode" value="not_checked_in" {8} onchange="this.form.submit()"> People to Check In
                            </label>
                            <label class="btn btn-default {9}" style="position: relative; float: left;">
                                <input type="radio" name="view_mode" value="checked_in" {10} onchange="this.form.submit()"> Checked In People
                            </label>
                        </div>
                    </form>
                </div>
            </div>
            
            <!-- Search and filters -->
            <form method="post" action="/PyScriptForm/{5}" onsubmit="showLoading()">
                <input type="hidden" name="step" value="check_in">
                <input type="hidden" name="view_mode" value="{11}">
                {6}
                
                <div class="row" style="display: flex; flex-wrap: wrap; margin-right: -15px; margin-left: -15px;">
                    <div class="col-xs-12" style="position: relative; min-height: 1px; padding-right: 15px; padding-left: 15px; flex: 0 0 100%; max-width: 100%;">
                        <div class="form-group" style="margin-bottom: 15px;">
                            <div class="input-group" style="position: relative; display: table; border-collapse: separate; width: 100%;">
                                <input type="text" class="form-control" name="search_term" value="{12}" placeholder="Search by name..." style="display: table-cell; position: relative; z-index: 2; float: left; width: 100%; margin-bottom: 0; height: 38px; padding: 8px 12px; font-size: 14px; line-height: 1.42857143; color: #555; background-color: #fff; background-image: none; border: 1px solid #ddd; border-radius: 4px; border-top-right-radius: 0; border-bottom-right-radius: 0; box-shadow: inset 0 1px 1px rgba(0,0,0,.075);">
                                <span class="input-group-btn" style="position: relative; font-size: 0; white-space: nowrap; width: 1%; vertical-align: middle; display: table-cell;">
                                    <button class="btn btn-primary" type="submit" onclick="showLoading()" style="border-top-left-radius: 0; border-bottom-left-radius: 0;">
                                        <i class="fa fa-search"></i> Search
                                    </button>
                                </span>
                            </div>
                        </div>
                    </div>
                </div>
                
                <div class="row" style="display: flex; flex-wrap: wrap; margin-right: -15px; margin-left: -15px;">
                    <div class="col-xs-12" style="position: relative; min-height: 1px; padding-right: 15px; padding-left: 15px; flex: 0 0 100%; max-width: 100%;">
                        <div class="form-group" style="margin-bottom: 15px;">
                            <label style="display: inline-block; max-width: 100%; margin-bottom: 5px; font-weight: 500;">Filter by Last Name:</label>
                            {13}
                        </div>
                    </div>
                </div>
            </form>
            
            <!-- People list -->
            <div class="list-group" style="padding-left: 0; margin-bottom: 20px;">
                {14}
            </div>
            
            <!-- Pagination -->
            <div class="text-center" style="text-align: center;">
                {15}
            </div>
        </div>
    </div>
    
    <script>
        // Show loading indicator during form submission
        function showLoading() {{
            document.getElementById('loadingIndicator').style.display = 'block';
        }}
    </script>
    </body>
    </html>
    """.format(
        success_message,            # {0}
        stats["checked_in"],        # {1}
        stats["not_checked_in"],    # {2}
        stats["total"],             # {3}
        "People to Check In" if view_mode == "not_checked_in" else "Checked In People", # {4}
        script_name,                # {5}
        meeting_id_inputs,          # {6}
        "active" if view_mode == "not_checked_in" else "", # {7}
        "checked" if view_mode == "not_checked_in" else "", # {8}
        "active" if view_mode == "checked_in" else "", # {9}
        "checked" if view_mode == "checked_in" else "", # {10}
        view_mode,                  # {11}
        search_term,                # {12}
        alpha_filters_html,         # {13}
        people_list_html_str,       # {14}
        pagination                  # {15}
    )
    
    return True  # Indicate successful rendering

# Simplified version with properly escaped curly braces for Python 2.7
def render_basic_check_in_page(check_in_manager):
    """Render basic check-in page with minimal CSS and properly escaped curly braces"""
    # Get the current script name
    script_name = get_script_name()
    
    # Get parameters
    meeting_ids = []
    if hasattr(check_in_manager.model.Data, 'meeting_id'):
        if isinstance(check_in_manager.model.Data.meeting_id, list):
            meeting_ids = [str(m) for m in check_in_manager.model.Data.meeting_id]
        else:
            meeting_ids = [str(check_in_manager.model.Data.meeting_id)]
            
    # Get other parameters
    alpha_filter = getattr(check_in_manager.model.Data, 'alpha_filter', 'All')
    search_term = getattr(check_in_manager.model.Data, 'search_term', '')
    view_mode = getattr(check_in_manager.model.Data, 'view_mode', 'not_checked_in')
    
    # Handle page
    try:
        page_value = getattr(check_in_manager.model.Data, 'page', 1)
        current_page = int(page_value) if page_value else 1
    except (ValueError, TypeError):
        current_page = 1

    # Get check-in statistics
    stats = check_in_manager.get_check_in_stats(meeting_ids)
    
    # Process form submission
    success_message = ""
    if hasattr(check_in_manager.model.Data, 'action'):
        action = check_in_manager.model.Data.action
        
        if action == 'single_direct_check_in':
            # Get person ID and meeting ID directly
            person_id = getattr(check_in_manager.model.Data, 'person_id', None)
            meeting_id = getattr(check_in_manager.model.Data, 'meeting_id', None)
            
            if person_id and meeting_id:
                # Direct API call - simplified for reliability
                try:
                    result = check_in_manager.model.EditPersonAttendance(int(meeting_id), int(person_id), True)
                    success = "Success" in str(result)
                    
                    if success:
                        success_message = """
                        <div style="padding:15px; margin-bottom:20px; border:1px solid #d6e9c6; border-radius:4px; background-color:#dff0d8; color:#3c763d;">
                            <strong>Success!</strong> Person has been checked in.
                        </div>
                        """
                    else:
                        success_message = """
                        <div style="padding:15px; margin-bottom:20px; border:1px solid #ebccd1; border-radius:4px; background-color:#f2dede; color:#a94442;">
                            <strong>Error!</strong> Could not check in the person.
                        </div>
                        """
                except Exception as e:
                    success_message = """
                    <div style="padding:15px; margin-bottom:20px; border:1px solid #ebccd1; border-radius:4px; background-color:#f2dede; color:#a94442;">
                        <strong>Error!</strong> Exception: {0}
                    </div>
                    """.format(str(e))
        
        elif action == 'remove_check_in' and hasattr(check_in_manager.model.Data, 'person_id'):
            # Single person check-in removal
            person_id = check_in_manager.model.Data.person_id
            meeting_id = getattr(check_in_manager.model.Data, 'meeting_id', None)
            
            # Direct API call
            if person_id and meeting_id:
                try:
                    result = check_in_manager.model.EditPersonAttendance(int(meeting_id), int(person_id), False)
                    success = "Success" in str(result)
                    
                    if success:
                        success_message = """
                        <div style="padding:15px; margin-bottom:20px; border:1px solid #d6e9c6; border-radius:4px; background-color:#dff0d8; color:#3c763d;">
                            <strong>Success!</strong> Check-in has been removed.
                        </div>
                        """
                except Exception:
                    success_message = """
                    <div style="padding:15px; margin-bottom:20px; border:1px solid #ebccd1; border-radius:4px; background-color:#f2dede; color:#a94442;">
                        <strong>Error!</strong> Could not remove check-in.
                    </div>
                    """
    
    # Get people based on view mode
    people = []
    total_count = 0
    
    if view_mode == "checked_in":
        people, total_count = check_in_manager.get_checked_in_people(meeting_ids, current_page, PAGE_SIZE)
    else:
        people, total_count = check_in_manager.get_people_by_filter(alpha_filter, search_term, meeting_ids, current_page, PAGE_SIZE)
    
    # Calculate pagination
    total_pages = (total_count + PAGE_SIZE - 1) // PAGE_SIZE
    
    # Create pagination controls
    pagination = create_pagination(current_page, total_pages, script_name, meeting_ids, view_mode, alpha_filter, search_term)
    
    # Render alpha filters
    alpha_filters_html = render_alpha_filters(alpha_filter)
    
    # Generate the people list HTML based on view mode
    people_list_html = []
    for person in people:
        # Use the primary meeting ID for simplicity
        meeting_id = meeting_ids[0] if meeting_ids else ""
        
        if view_mode == "checked_in":
            # For checked-in people, show "remove" button
            item_html = """
            <div style="position:relative; display:block; padding:10px 15px; margin-bottom:10px; background-color:#fff; border:1px solid #ddd; border-radius:4px;">
                <div style="width:100%; overflow:hidden;">
                    <div style="float:left; width:66.6%; padding-right:15px;">
                        <span style="font-size:16px; line-height:28px;">{1}</span>
                    </div>
                    <div style="float:right; width:33.3%; text-align:right;">
                        <form method="post" action="/PyScriptForm/{2}" style="display:inline;" onsubmit="document.getElementById('loadingIndicator').style.display='block';">
                            <input type="hidden" name="step" value="check_in">
                            <input type="hidden" name="action" value="remove_check_in">
                            <input type="hidden" name="person_id" value="{0}">
                            <input type="hidden" name="meeting_id" value="{3}">
                            <input type="hidden" name="view_mode" value="checked_in">
                            <input type="hidden" name="alpha_filter" value="{4}">
                            <input type="hidden" name="search_term" value="{5}">
                            <input type="hidden" name="page" value="{6}">
                            <button type="submit" style="padding:5px 10px; font-size:12px; color:#fff; background-color:#d9534f; border:1px solid #d43f3a; border-radius:3px; cursor:pointer;">
                                Remove
                            </button>
                        </form>
                    </div>
                </div>
            </div>
            """.format(
                person.people_id,  # {0}
                person.name,       # {1}
                script_name,       # {2}
                meeting_id,        # {3}
                alpha_filter,      # {4}
                search_term,       # {5}
                current_page       # {6}
            )
        else:
            # For not-checked-in people, show "check in" button
            item_html = """
            <div style="position:relative; display:block; padding:10px 15px; margin-bottom:10px; background-color:#fff; border:1px solid #ddd; border-radius:4px;">
                <div style="width:100%; overflow:hidden;">
                    <div style="float:left; width:66.6%; padding-right:15px;">
                        <span style="font-size:16px; line-height:28px;">{1}</span>
                    </div>
                    <div style="float:right; width:33.3%; text-align:right;">
                        <form method="post" action="/PyScriptForm/{2}" style="display:inline;" onsubmit="document.getElementById('loadingIndicator').style.display='block';">
                            <input type="hidden" name="step" value="check_in">
                            <input type="hidden" name="action" value="single_direct_check_in">
                            <input type="hidden" name="person_id" value="{0}">
                            <input type="hidden" name="meeting_id" value="{3}">
                            <input type="hidden" name="view_mode" value="not_checked_in">
                            <input type="hidden" name="alpha_filter" value="{4}">
                            <input type="hidden" name="search_term" value="{5}">
                            <input type="hidden" name="page" value="{6}">
                            <button type="submit" style="padding:5px 10px; font-size:12px; color:#fff; background-color:#5cb85c; border:1px solid #4cae4c; border-radius:3px; cursor:pointer;">
                                Check In
                            </button>
                        </form>
                    </div>
                </div>
            </div>
            """.format(
                person.people_id,  # {0}
                person.name,       # {1}
                script_name,       # {2}
                meeting_id,        # {3}
                alpha_filter,      # {4}
                search_term,       # {5}
                current_page       # {6}
            )
        
        people_list_html.append(item_html)
    
    # Combine all list items
    people_list_html_str = "\n".join(people_list_html)
    
    # Create inputs for meeting IDs (for other forms)
    meeting_id_inputs = ""
    for meeting_id in meeting_ids:
        meeting_id_inputs += '<input type="hidden" name="meeting_id" value="{0}">'.format(meeting_id)
    
    # Simple HTML without complex CSS or curly braces in JavaScript
    print """<!DOCTYPE html>
    <html>
    <head>
        <title>Rapid Check-In System</title>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="font-family:Arial,sans-serif; margin:0; padding:20px; color:#333; background-color:#f9f9f9;">
        <div id="loadingIndicator" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background-color:rgba(255,255,255,0.8); z-index:1000; text-align:center; padding-top:20%;">
            <div style="border:6px solid #f3f3f3; border-top:6px solid #3498db; border-radius:50%; width:50px; height:50px; margin:0 auto 20px auto;"></div>
            <div>Processing...</div>
        </div>
        
        <div style="background-color:#3498db; color:#fff; padding:15px; margin:-20px -20px 20px -20px; border-bottom:3px solid #2980b9;">
            <h2 style="margin:0; color:#fff;">Rapid Check-In System</h2>
        </div>
        
        <!-- Success message -->
        {0}
        
        <!-- Stats boxes -->
        <div style="margin-bottom:20px; overflow:hidden;">
            <div style="float:left; width:31%; margin:0 1%; padding:15px; text-align:center; background-color:#d9f2e6; border-left:4px solid #2ecc71; border-radius:4px;">
                <div style="font-size:24px; font-weight:bold; margin-bottom:5px;">{1}</div>
                <div style="font-size:14px; color:#888;">Checked In</div>
            </div>
            <div style="float:left; width:31%; margin:0 1%; padding:15px; text-align:center; background-color:#e6f2fa; border-left:4px solid #3498db; border-radius:4px;">
                <div style="font-size:24px; font-weight:bold; margin-bottom:5px;">{2}</div>
                <div style="font-size:14px; color:#888;">Not Checked In</div>
            </div>
            <div style="float:left; width:31%; margin:0 1%; padding:15px; text-align:center; background-color:#f9f9f9; border-left:4px solid #95a5a6; border-radius:4px;">
                <div style="font-size:24px; font-weight:bold; margin-bottom:5px;">{3}</div>
                <div style="font-size:14px; color:#888;">Total</div>
            </div>
        </div>
    
        <!-- Main content -->
        <div style="margin-bottom:20px; background-color:#fff; border:1px solid #ddd; border-radius:4px; padding:0;">
            <div style="padding:12px 15px; background-color:#f7f7f7; border-bottom:1px solid #ddd;">
                <div style="overflow:hidden;">
                    <div style="float:left; width:60%;">
                        <h3 style="margin:0; font-size:16px; color:#333;">{4}</h3>
                    </div>
                    <div style="float:right; width:40%; text-align:right;">
                        <a href="/PyScript/{5}?step=choose_meetings" style="padding:5px 10px; background-color:#f0f0f0; border:1px solid #ddd; border-radius:4px; text-decoration:none; color:#333; display:inline-block;" onclick="document.getElementById('loadingIndicator').style.display='block';">
                            Back to Meetings
                        </a>
                    </div>
                </div>
            </div>
            
            <div style="padding:15px;">
                <!-- View toggle buttons -->
                <div style="margin-bottom:20px;">
                    <form method="post" action="/PyScriptForm/{5}" style="display:inline-block;" onsubmit="document.getElementById('loadingIndicator').style.display='block';">
                        <input type="hidden" name="step" value="check_in">
                        {6}
                        <span style="display:inline-block;">
                            <label style="display:inline-block; padding:6px 12px; margin-right:-1px; background-color:{7}; border:1px solid #ccc; border-top-left-radius:4px; border-bottom-left-radius:4px;">
                                <input type="radio" name="view_mode" value="not_checked_in" {8} onchange="this.form.submit()" style="margin-right:5px;"> People to Check In
                            </label>
                            <label style="display:inline-block; padding:6px 12px; background-color:{9}; border:1px solid #ccc; border-top-right-radius:4px; border-bottom-right-radius:4px;">
                                <input type="radio" name="view_mode" value="checked_in" {10} onchange="this.form.submit()" style="margin-right:5px;"> Checked In People
                            </label>
                        </span>
                    </form>
                </div>
                
                <!-- Search and filters -->
                <form method="post" action="/PyScriptForm/{5}" onsubmit="document.getElementById('loadingIndicator').style.display='block';">
                    <input type="hidden" name="step" value="check_in">
                    <input type="hidden" name="view_mode" value="{11}">
                    {6}
                    
                    <div style="margin-bottom:15px;">
                        <div style="display:table; width:100%;">
                            <div style="display:table-cell; width:100%;">
                                <input type="text" name="search_term" value="{12}" placeholder="Search by name..." style="width:100%; height:38px; padding:6px 12px; border:1px solid #ddd; border-radius:4px 0 0 4px; box-sizing:border-box;">
                            </div>
                            <div style="display:table-cell; vertical-align:top; width:1%;">
                                <button type="submit" style="height:38px; padding:6px 12px; background-color:#337ab7; color:white; border:1px solid #2e6da4; border-radius:0 4px 4px 0; cursor:pointer; white-space:nowrap;" onclick="document.getElementById('loadingIndicator').style.display='block';">
                                    Search
                                </button>
                            </div>
                        </div>
                    </div>
                    
                    <div style="margin-bottom:15px;">
                        <label style="display:block; margin-bottom:5px; font-weight:bold;">Filter by Last Name:</label>
                        {13}
                    </div>
                </form>
                
                <!-- People list -->
                <div style="margin-bottom:20px;">
                    {14}
                </div>
                
                <!-- Pagination -->
                <div style="text-align:center;">
                    {15}
                </div>
            </div>
        </div>
        
        <script>
            function showLoading() {{
                document.getElementById('loadingIndicator').style.display = 'block';
            }}
        </script>
    </body>
    </html>
    """.format(
            success_message,            # {0}
            stats["checked_in"],        # {1}
            stats["not_checked_in"],    # {2}
            stats["total"],             # {3}
            "People to Check In" if view_mode == "not_checked_in" else "Checked In People", # {4}
            script_name,                # {5}
            meeting_id_inputs,          # {6}
            "#eee" if view_mode == "not_checked_in" else "#fff", # {7}
            "checked" if view_mode == "not_checked_in" else "", # {8}
            "#fff" if view_mode == "not_checked_in" else "#eee", # {9}
            "checked" if view_mode == "checked_in" else "", # {10}
            view_mode,                  # {11}
            search_term,                # {12}
            alpha_filters_html,         # {13}
            people_list_html_str,       # {14}
            pagination                  # {15}
        )
        
    return True  # Indicate successful rendering


# Main execution
try:

    # Set up the check-in manager
    check_in_manager = CheckInManager(model, q)
    
    # Clean input data to handle duplicated values in form submission
    clean_step = getattr(model.Data, 'step', 'choose_meetings')
    if ',' in str(clean_step):
        clean_step = str(clean_step).split(',')[0].strip()
        
    clean_action = getattr(model.Data, 'action', None)
    if clean_action and ',' in str(clean_action):
        clean_action = str(clean_action).split(',')[0].strip()
        
    clean_person_id = getattr(model.Data, 'person_id', None)
    if clean_person_id and ',' in str(clean_person_id):
        clean_person_id = str(clean_person_id).split(',')[0].strip()
        
    clean_meeting_id = getattr(model.Data, 'meeting_id', None)
    if clean_meeting_id and ',' in str(clean_meeting_id):
        clean_meeting_id = str(clean_meeting_id).split(',')[0].strip()
    
    # Store cleaned values
    model.Data.cleaned_step = clean_step
    model.Data.cleaned_action = clean_action
    model.Data.cleaned_person_id = clean_person_id
    model.Data.cleaned_meeting_id = clean_meeting_id
    
    # Special handling for direct check-in
    if clean_action == 'single_direct_check_in':
        # Do direct API call
        if clean_person_id and clean_meeting_id:
            try:
                # Direct API call
                result = check_in_manager.model.EditPersonAttendance(int(clean_meeting_id), int(clean_person_id), True)
                # Update last check-in time to force stats refresh
                check_in_manager.last_check_in_time = check_in_manager.model.DateTime
            except Exception as e:
                print "Error: " + str(e) + ""
    
    # Determine which page to render
    if clean_step == 'check_in':
        render_fastlane_check_in(check_in_manager)
    else:
        render_meeting_selection(check_in_manager)
        
except Exception as e:
    # Print any errors
    import traceback
    print "Error"
    print "An error occurred: {0}".format(str(e))
    print ""
    traceback.print_exc()
    print ""
    
    # Link to go back
    print ("""Back to Check-In""")
