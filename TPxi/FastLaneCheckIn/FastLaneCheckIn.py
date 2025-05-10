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
PAGE_SIZE = 500  # Number of people to show per page - smaller for faster loading
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
    """Create a pagination control with proper layout"""
    if total_pages <= 1:
        return ""
        
    # Create hidden inputs for meeting IDs
    meeting_id_inputs = ""
    for meeting_id in meeting_ids:
        meeting_id_inputs += '<input type="hidden" name="meeting_id" value="{0}">'.format(meeting_id)
    
    # Build pagination HTML with proper styling
    pagination = ['<div class="pagination" style="display: inline-block;">']
    
    # Previous button
    if page > 1:
        pagination.append(
            '<form method="post" action="/PyScriptForm/{0}" style="display:inline-block; margin:0 2px;">'
            '<input type="hidden" name="step" value="check_in">'
            '<input type="hidden" name="page" value="{1}">'
            '<input type="hidden" name="view_mode" value="{2}">'
            '<input type="hidden" name="alpha_filter" value="{3}">'
            '<input type="hidden" name="search_term" value="{4}">'
            '{5}'
            '<button type="submit" style="padding:6px 12px; border:1px solid #ddd; background-color:#f8f9fa; border-radius:4px; cursor:pointer; margin:0;" onclick="showLoading()">&laquo;</button>'
            '</form>'.format(
                script_name, page - 1, view_mode, alpha_filter, search_term, meeting_id_inputs
            )
        )
    else:
        pagination.append(
            '<span style="display:inline-block; padding:6px 12px; border:1px solid #ddd; background-color:#f8f9fa; border-radius:4px; color:#999; margin:0 2px;">&laquo;</span>'
        )
        
    # Page numbers
    start_page = max(1, page - 2)
    end_page = min(total_pages, page + 2)
    
    # Always show page 1
    if start_page > 1:
        pagination.append(
            '<form method="post" action="/PyScriptForm/{0}" style="display:inline-block; margin:0 2px;">'
            '<input type="hidden" name="step" value="check_in">'
            '<input type="hidden" name="page" value="1">'
            '<input type="hidden" name="view_mode" value="{1}">'
            '<input type="hidden" name="alpha_filter" value="{2}">'
            '<input type="hidden" name="search_term" value="{3}">'
            '{4}'
            '<button type="submit" style="padding:6px 12px; border:1px solid #ddd; background-color:#f8f9fa; border-radius:4px; cursor:pointer; margin:0;" onclick="showLoading()">1</button>'
            '</form>'.format(
                script_name, view_mode, alpha_filter, search_term, meeting_id_inputs
            )
        )
        if start_page > 2:
            pagination.append(
                '<span style="display:inline-block; padding:6px 12px; margin:0 2px;">...</span>'
            )
            
    # Page links
    for i in range(start_page, end_page + 1):
        if i == page:
            pagination.append(
                '<span style="display:inline-block; padding:6px 12px; border:1px solid #007bff; background-color:#007bff; color:white; border-radius:4px; margin:0 2px;">{0}</span>'.format(i)
            )
        else:
            pagination.append(
                '<form method="post" action="/PyScriptForm/{0}" style="display:inline-block; margin:0 2px;">'
                '<input type="hidden" name="step" value="check_in">'
                '<input type="hidden" name="page" value="{1}">'
                '<input type="hidden" name="view_mode" value="{2}">'
                '<input type="hidden" name="alpha_filter" value="{3}">'
                '<input type="hidden" name="search_term" value="{4}">'
                '{5}'
                '<button type="submit" style="padding:6px 12px; border:1px solid #ddd; background-color:#f8f9fa; border-radius:4px; cursor:pointer; margin:0;" onclick="showLoading()">{1}</button>'
                '</form>'.format(
                    script_name, i, view_mode, alpha_filter, search_term, meeting_id_inputs
                )
            )
            
    # Always show last page
    if end_page < total_pages:
        if end_page < total_pages - 1:
            pagination.append(
                '<span style="display:inline-block; padding:6px 12px; margin:0 2px;">...</span>'
            )
        pagination.append(
            '<form method="post" action="/PyScriptForm/{0}" style="display:inline-block; margin:0 2px;">'
            '<input type="hidden" name="step" value="check_in">'
            '<input type="hidden" name="page" value="{1}">'
            '<input type="hidden" name="view_mode" value="{2}">'
            '<input type="hidden" name="alpha_filter" value="{3}">'
            '<input type="hidden" name="search_term" value="{4}">'
            '{5}'
            '<button type="submit" style="padding:6px 12px; border:1px solid #ddd; background-color:#f8f9fa; border-radius:4px; cursor:pointer; margin:0;" onclick="showLoading()">{1}</button>'
            '</form>'.format(
                script_name, total_pages, view_mode, alpha_filter, search_term, meeting_id_inputs
            )
        )
        
    # Next button
    if page < total_pages:
        pagination.append(
            '<form method="post" action="/PyScriptForm/{0}" style="display:inline-block; margin:0 2px;">'
            '<input type="hidden" name="step" value="check_in">'
            '<input type="hidden" name="page" value="{1}">'
            '<input type="hidden" name="view_mode" value="{2}">'
            '<input type="hidden" name="alpha_filter" value="{3}">'
            '<input type="hidden" name="search_term" value="{4}">'
            '{5}'
            '<button type="submit" style="padding:6px 12px; border:1px solid #ddd; background-color:#f8f9fa; border-radius:4px; cursor:pointer; margin:0;" onclick="showLoading()">&raquo;</button>'
            '</form>'.format(
                script_name, page + 1, view_mode, alpha_filter, search_term, meeting_id_inputs
            )
        )
    else:
        pagination.append(
            '<span style="display:inline-block; padding:6px 12px; border:1px solid #ddd; background-color:#f8f9fa; border-radius:4px; color:#999; margin:0 2px;">&raquo;</span>'
        )
        
    pagination.append('</div>')
    return "\n".join(pagination)


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
            # If it's a comma-separated string, split it
            if ',' in str(check_in_manager.model.Data.meeting_id):
                meeting_ids = str(check_in_manager.model.Data.meeting_id).split(',')
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
            
            print "<!-- DEBUG: single_direct_check_in called with person_id={0}, meeting_id={1} -->".format(person_id, meeting_id)
            
            if person_id and meeting_id:
                # Check if meeting_id contains commas (multiple meetings)
                if ',' in str(meeting_id):
                    # This is the issue - we need to handle only a single meeting ID
                    print "<!-- DEBUG: Multiple meeting IDs detected: {0} -->".format(meeting_id)
                    # Just use the first meeting ID in the list
                    meeting_id = str(meeting_id).split(',')[0].strip()
                    print "<!-- DEBUG: Using first meeting ID instead: {0} -->".format(meeting_id)
                
                # Direct API call - simplified for reliability
                try:
                    # Force refresh of statistics
                    check_in_manager.last_check_in_time = check_in_manager.model.DateTime
                    
                    result = check_in_manager.model.EditPersonAttendance(int(meeting_id), int(person_id), True)
                    success = "Success" in str(result)
                    
                    print "<!-- DEBUG: EditPersonAttendance result: {0} -->".format(result)
                    
                    if success:
                        flash_message = "Successfully checked in"
                        flash_name = person_name
                except Exception as e:
                    flash_message = "Error checking in: {0}".format(str(e))
                    print "<!-- DEBUG: Check-in error: {0} -->".format(str(e))
        
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
    <div class="people-grid" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:10px; margin-bottom:20px;">
    """)
    
    for person in people:
        # Get the organizations this person is a member of
        person_org_names = []
        for org_id in person.org_ids:
            # Find the org name for this ID
            for meeting in check_in_manager.all_meetings_today:
                if str(meeting.org_id) == str(org_id):
                    if meeting.org_name not in person_org_names:
                        person_org_names.append(meeting.org_name)
                    break
        
        # Create a representation of the organizations this person belongs to
        org_display = ""
        if person_org_names:
            org_display = '<div style="font-size:12px; color:#666; margin-top:2px;">' + ', '.join(person_org_names) + '</div>'
        
        # For disambiguation in case of duplicate names
        unique_identifier = ""
        if person.people_id:
            # Convert to string first before trying to slice
            people_id_str = str(person.people_id)
            # Get last 4 characters if long enough
            last_digits = people_id_str[-4:] if len(people_id_str) >= 4 else people_id_str
            unique_identifier = '<span style="font-size:10px; color:#999;"> (ID: {0})</span>'.format(last_digits)
                
        if view_mode == "checked_in":
            # For checked-in people, we need to determine which meetings they're checked into
            # Get the person's attended meetings
            attended_meeting_ids = []
            
            try:
                # First get all meeting IDs this person has attended today
                sql = """
                    SELECT DISTINCT MeetingId
                    FROM Attend
                    WHERE PeopleId = {0}
                    AND CONVERT(date, MeetingDate) = CONVERT(date, GETDATE())
                    AND AttendanceFlag = 1
                """.format(person.people_id)
                
                results = check_in_manager.q.QuerySql(sql)
                
                # Extract the meeting IDs from the results
                for result in results:
                    attended_meeting_id = str(result.MeetingId)
                    # Only include meetings that were selected in the UI
                    if attended_meeting_id in meeting_ids:
                        attended_meeting_ids.append(attended_meeting_id)
            except Exception as e:
                print "<!-- DEBUG: Error getting attended meeting IDs: {0} -->".format(str(e))
            
            # If we couldn't get attended meetings from SQL, fall back to the first meeting ID
            if not attended_meeting_ids and meeting_ids:
                attended_meeting_ids = [meeting_ids[0]]
                
            # Start the HTML for this person
            item_html = """
            <div style="padding:8px; background-color:#fff; border:1px solid #ddd; border-radius:4px;">
                <div style="display:flex; align-items:center; justify-content:space-between;">
                    <div style="flex:1; overflow:hidden; text-overflow:ellipsis;">
                        <div style="font-size:15px;">{1}{2}</div>
                        {3}
                    </div>
            """.format(
                person.people_id,       # {0}
                person.name,            # {1}
                unique_identifier,      # {2}
                org_display,            # {3}
            )
            
            # If we have multiple attended meetings, show a button for each
            if len(attended_meeting_ids) > 1:
                item_html += """
                    <div style="display:flex; flex-wrap:wrap; justify-content:flex-end; gap:5px;">
                """
                
                # Add a remove button for each attended meeting
                for meeting_id in attended_meeting_ids:
                    # Get meeting info for display
                    meeting_info = None
                    for m in check_in_manager.all_meetings_today:
                        if str(m.meeting_id) == meeting_id:
                            meeting_info = m
                            break
                    
                    # Use the org name as the button label if we have meeting info
                    button_label = "Remove"
                    if meeting_info:
                        button_label = "Remove " + meeting_info.org_name.split(' ')[0]
                    
                    # Add the remove button
                    item_html += """
                    <form method="post" action="/PyScriptForm/{4}" style="margin:0; padding:0;" onsubmit="document.getElementById('loadingIndicator').style.display='block';">
                        <input type="hidden" name="step" value="check_in">
                        <input type="hidden" name="action" value="remove_check_in">
                        <input type="hidden" name="person_id" value="{0}">
                        <input type="hidden" name="person_name" value="{1}">
                        <input type="hidden" name="meeting_id" value="{5}">
                        <input type="hidden" name="view_mode" value="checked_in">
                        <input type="hidden" name="alpha_filter" value="{6}">
                        <input type="hidden" name="search_term" value="{7}">
                        <input type="hidden" name="page" value="{8}">
                        <button type="submit" style="padding:3px 8px; font-size:12px; background-color:#dc3545; color:#fff; border:1px solid #dc3545; border-radius:3px; cursor:pointer; min-width:60px; height:28px; vertical-align:middle;">{9}</button>
                    </form>
                    """.format(
                        person.people_id,       # {0}
                        person.name,            # {1}
                        "",                     # {2} - not used
                        "",                     # {3} - not used
                        script_name,            # {4}
                        meeting_id,             # {5}
                        alpha_filter,           # {6}
                        search_term,            # {7}
                        current_page,           # {8}
                        button_label            # {9}
                    )
                    
                # Close the container
                item_html += """
                    </div>
                """
            else:
                # Just one meeting - use a single remove button
                meeting_id = attended_meeting_ids[0] if attended_meeting_ids else (meeting_ids[0] if meeting_ids else "")
                
                item_html += """
                    <form method="post" action="/PyScriptForm/{4}" style="margin:0; padding:0; flex-shrink:0;" onsubmit="document.getElementById('loadingIndicator').style.display='block';">
                        <input type="hidden" name="step" value="check_in">
                        <input type="hidden" name="action" value="remove_check_in">
                        <input type="hidden" name="person_id" value="{0}">
                        <input type="hidden" name="person_name" value="{1}">
                        <input type="hidden" name="meeting_id" value="{5}">
                        <input type="hidden" name="view_mode" value="checked_in">
                        <input type="hidden" name="alpha_filter" value="{6}">
                        <input type="hidden" name="search_term" value="{7}">
                        <input type="hidden" name="page" value="{8}">
                        <button type="submit" style="padding:3px 8px; font-size:12px; background-color:#dc3545; color:#fff; border:1px solid #dc3545; border-radius:3px; cursor:pointer; min-width:60px; height:28px; vertical-align:middle;">Remove</button>
                    </form>
                """.format(
                    person.people_id,       # {0}
                    person.name,            # {1}
                    "",                     # {2} - not used
                    "",                     # {3} - not used
                    script_name,            # {4}
                    meeting_id,             # {5}
                    alpha_filter,           # {6}
                    search_term,            # {7}
                    current_page            # {8}
                )
            
            # Close the containers
            item_html += """
                </div>
            </div>
            """

        else:
            # For not-checked-in people, show "check in" button with individual meeting buttons
            # Create a container for this person
            item_html = """
            <div style="padding:8px; background-color:#fff; border:1px solid #ddd; border-radius:4px;">
                <div style="display:flex; align-items:center; justify-content:space-between;">
                    <div style="flex:1; overflow:hidden; text-overflow:ellipsis;">
                        <div style="font-size:15px;">{1}{2}</div>
                        {3}
                    </div>
                    <div style="display:flex; flex-wrap:wrap; justify-content:flex-end; gap:5px;">
            """.format(
                person.people_id,       # {0}
                person.name,            # {1}
                unique_identifier,      # {2}
                org_display,            # {3}
            )
        
            # Debug information
            print "<!-- DETAILED DEBUG FOR: {0} (ID: {1}) -->".format(person.name, person.people_id)
            print "<!-- Person org IDs: {0} -->".format([str(org_id) for org_id in person.org_ids])
            print "<!-- Meeting IDs: {0} -->".format(meeting_ids)
            
            # FIXED SECTION: Create a check-in button for each meeting that this person should attend
            # Find the correct meetings for this person's organizations
            person_meeting_ids = []
            
            # Map from person's org_ids to meeting_ids that are active today
            for org_id in person.org_ids:
                for meeting in check_in_manager.all_meetings_today:
                    if (str(meeting.org_id) == str(org_id) and 
                        str(meeting.meeting_id) in meeting_ids):
                        person_meeting_ids.append(str(meeting.meeting_id))
                        print "<!-- Found matching meeting ID {0} for org {1} -->".format(
                            meeting.meeting_id, org_id)
            
            # If no specific meetings found, use first available meeting
            if not person_meeting_ids and meeting_ids:
                person_meeting_ids = [meeting_ids[0]]
                print "<!-- No specific meetings found, using first meeting ID: {0} -->".format(meeting_ids[0])
            
            # Add check-in buttons for each appropriate meeting
            for meeting_id in person_meeting_ids:
                # Find meeting info for button label
                meeting_name = "Check In"
                for meeting in check_in_manager.all_meetings_today:
                    if str(meeting.meeting_id) == str(meeting_id):
                        if len(person_meeting_ids) > 1:
                            # If multiple buttons, use org name to differentiate
                            meeting_name = meeting.org_name.split(' ')[0]
                        break
                
                print "<!-- Adding check-in button for meeting ID: {0} ({1}) -->".format(meeting_id, meeting_name)
                
                # Create a single check-in button for this specific meeting
                item_html += """
                <form method="post" action="/PyScriptForm/{4}" style="margin:0; padding:0;" onsubmit="document.getElementById('loadingIndicator').style.display='block';">
                    <input type="hidden" name="step" value="check_in">
                    <input type="hidden" name="action" value="single_direct_check_in">
                    <input type="hidden" name="person_id" value="{0}">
                    <input type="hidden" name="person_name" value="{1}">
                    <input type="hidden" name="meeting_id" value="{5}">
                    <input type="hidden" name="view_mode" value="not_checked_in">
                    <input type="hidden" name="alpha_filter" value="{6}">
                    <input type="hidden" name="search_term" value="{7}">
                    <input type="hidden" name="page" value="{8}">
                    <button type="submit" style="padding:3px 8px; font-size:12px; background-color:#28a745; color:#fff; border:1px solid #28a745; border-radius:3px; cursor:pointer; min-width:60px; height:28px; vertical-align:middle;">{9}</button>
                </form>
                """.format(
                    person.people_id,  # {0}
                    person.name,       # {1}
                    "",                # {2} - not used
                    "",                # {3} - not used
                    script_name,       # {4}
                    meeting_id,        # {5} - specific meeting ID
                    alpha_filter,      # {6}
                    search_term,       # {7}
                    current_page,      # {8}
                    meeting_name       # {9} - button label
                )
            
            # Close the containers
            item_html += """
                    </div>
                </div>
            </div>
            """
        
        people_list_html.append(item_html)
    
    people_list_html.append("</div>")
    
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
        "\n".join(people_list_html), # {10}
        pagination                # {11}
    )
    
    return True  # Indicate successful rendering


def render_checked_in_view(check_in_manager):
    """Render the checked-in people view"""
    # Similar to render_check_in_page but displays already checked-in people
    # This allows removing individual check-ins if needed
    return render_check_in_page(check_in_manager)
    

    

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
    
    # New: Handle selected_org_ids for multi-meeting check-in
    selected_org_ids = []
    if hasattr(model.Data, 'selected_org_ids'):
        org_ids_str = str(model.Data.selected_org_ids)
        if org_ids_str:
            selected_org_ids = org_ids_str.split(',')
    
    # Store cleaned values
    model.Data.cleaned_step = clean_step
    model.Data.cleaned_action = clean_action
    model.Data.cleaned_person_id = clean_person_id
    model.Data.cleaned_meeting_id = clean_meeting_id
    model.Data.selected_org_ids = selected_org_ids
    
    # Process different actions
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
    
    # New action type for multi-meeting check-in
    elif clean_action == 'multi_direct_check_in':
        # Get person ID
        if not clean_person_id:
            print "<div style='color:red;'>Error: No person ID provided</div>"
        else:
            try:
                # Simple approach: Just use the meeting IDs directly from the form
                success = True
                check_in_count = 0
                
                # Get all meetings selected on the meetings page
                all_selected_meeting_ids = []
                if hasattr(model.Data, 'meeting_id'):
                    if isinstance(model.Data.meeting_id, list):
                        all_selected_meeting_ids = [m for m in model.Data.meeting_id if m]  # Filter out empty values
                    elif model.Data.meeting_id:  # Check if not empty
                        all_selected_meeting_ids = [model.Data.meeting_id]
                
                print "<!-- DEBUG: Person ID: {0}, Meeting IDs: {1} -->".format(clean_person_id, all_selected_meeting_ids)
                
                # Check person into each meeting directly
                for meeting_id in all_selected_meeting_ids:
                    if meeting_id:  # Skip empty meeting IDs
                        try:
                            # Convert to integers and use direct API call
                            person_id_int = int(clean_person_id)
                            meeting_id_int = int(meeting_id)
                            
                            # Skip if zeros to prevent SQL errors
                            if person_id_int <= 0 or meeting_id_int <= 0:
                                continue
                                
                            result = model.EditPersonAttendance(meeting_id_int, person_id_int, True)
                            
                            if "Success" in str(result):
                                check_in_count += 1
                            else:
                                print "<!-- DEBUG: Check-in failed: {0} -->".format(result)
                                success = False
                        except Exception as e:
                            print "<!-- DEBUG: Check-in error for meeting {0}: {1} -->".format(meeting_id, str(e))
                            success = False
                
                # Update last check-in time to force stats refresh
                check_in_manager.last_check_in_time = check_in_manager.model.DateTime
                
                if check_in_count > 0:
                    print "<!-- DEBUG: Successfully checked in to {0} meetings -->".format(check_in_count)
                else:
                    print "<div style='color:red;'>No check-ins were completed. Please try again.</div>"
                    
            except Exception as e:
                print "<div style='color:red;'>Error during check-in: " + str(e) + "</div>"
                import traceback
                print "<!-- DEBUG: Full error details: " + traceback.format_exc() + " -->"
    
    # Handle check-in removal
    elif clean_action == 'remove_check_in':
        if clean_person_id and hasattr(model.Data, 'meeting_id'):
            success = True
            try:
                meeting_id = model.Data.meeting_id
                if isinstance(meeting_id, list):
                    for mid in meeting_id:
                        result = check_in_manager.model.EditPersonAttendance(int(mid), int(clean_person_id), False)
                        if "Success" not in str(result):
                            success = False
                else:
                    result = check_in_manager.model.EditPersonAttendance(int(meeting_id), int(clean_person_id), False)
                    if "Success" not in str(result):
                        success = False
                
                # Update last check-in time to force stats refresh
                check_in_manager.last_check_in_time = check_in_manager.model.DateTime
                
                if not success:
                    print "<div style='color:red;'>Some check-in removals failed</div>"
            except Exception as e:
                print "<div style='color:red;'>Error: " + str(e) + "</div>"
    
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
