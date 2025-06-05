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
- Parent email notifications for attendees under 18

NOTE: this does not print badges as those options are not exposed for me to access

--Upload Instructions Start--
To upload code to Touchpoint, use the following steps:
1. Click Admin ~ Advanced ~ Special Content ~ Python
2. Click New Python Script File
3. Name the Python FastLaneCheckIn (case sensitive) and paste all this code.  If you name it something else, then just update variable below.
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
Script_Name = "FastLaneCheckIn2"
PAGE_SIZE = 500  # Number of people to show per page - smaller for faster loading
DATE_FORMAT = "M/d/yyyy"
ATTEND_FLAG = 1  # Present flag for attendance
ALPHA_BLOCKS = [
    "A-C", "D-F", "G-I", "J-L", "M-O", "P-R", "S-U", "V-Z", "All"
]

# Email configuration constants
PARENT_EMAIL_DELAY = False  # Set to True to queue emails for batch processing
DEFAULT_EMAIL_TEMPLATE = "CheckInParentNotification2"  # Default email template name

"""
Email Template Variables
Use standard TouchPoint replacement codes plus one custom addition:

{CheckInPerson} - Name of the person who checked in (CUSTOM - works for both adults and children)

All other standard TouchPoint variables work normally:
{name} - Recipient's name (parent for parent emails, adult for adult emails)  
{first} - Recipient's first name
{orgname} - Organization/meeting name
{today} - Today's date
{ChurchName}, {ChurchAddress}, {ChurchPhone} - Church info from settings

Example template that works for both parents and adults:
"Dear {name}, {CheckInPerson} has been checked in to {orgname} at {today}."
"""

# ::START:: Helper Functions
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

# ::START:: Helper Classes
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
    def __init__(self, people_id, name, family_id=None, org_ids=None, checked_in=False, age=None):
        self.people_id = people_id
        self.name = name
        self.family_id = family_id
        self.org_ids = org_ids or []
        self.checked_in = checked_in
        self.age = age  # Added age field

# ::START:: Email Manager Class
class EmailManager:
    """Manages email notifications for check-ins"""
    
    def __init__(self, model, q):
        self.model = model
        self.q = q
        self.email_queue = []  # Store emails to send in batch
        
    def get_available_email_templates(self):
        """Get list of available email templates for parent notifications"""
        # ::STEP:: Query for email templates
        sql = """
            SELECT Name, Title
            FROM Content
            WHERE TypeID IN (2,7)  -- Email templates
            AND (Name LIKE '%CheckedIn%')-- OR Name LIKE '%Parent%' OR Title LIKE '%Check%')
            ORDER BY Name
        """
        try:
            results = self.q.QuerySql(sql)
            templates = []
            for result in results:
                templates.append({
                    'name': result.Name,
                    'title': result.Title or result.Name
                })
            return templates
        except:
            # Return default template if query fails
            return [{'name': DEFAULT_EMAIL_TEMPLATE, 'title': 'Check-In Parent Notification'}]
            
    
    def get_parent_emails(self, people_id, age):
        """Get parent email addresses for a person under 18"""
        # ::STEP:: Check age and get parent emails
        print "<!-- DEBUG: Getting parent emails for person {0}, age {1} -->".format(people_id, age)
        
        if age >= 18:
            print "<!-- DEBUG: Person is 18 or older, no parent email needed -->"
            return []
            
        sql = """
            SELECT DISTINCT p2.EmailAddress, p2.Name, p2.PeopleId
            FROM People p1
            JOIN People p2 ON p1.FamilyId = p2.FamilyId
            WHERE p1.PeopleId = {0}
            AND p2.PositionInFamilyId IN (10, 20)  -- Primary and Secondary Adults
            AND p2.EmailAddress IS NOT NULL
            AND p2.EmailAddress != ''
            AND (p2.DoNotMailFlag IS NULL OR p2.DoNotMailFlag = 0)
        """.format(people_id)
        
        try:
            results = self.q.QuerySql(sql)
            parents = []
            for result in results:
                print "<!-- DEBUG: Found parent: {0} ({1}) -->".format(result.Name, result.EmailAddress)
                parents.append({
                    'email': result.EmailAddress,
                    'name': result.Name
                })
            
            if not parents:
                print "<!-- DEBUG: No parents found with valid emails -->"
                
            return parents
        except Exception as e:
            print "<!-- DEBUG: Error getting parent emails: {0} -->".format(str(e))
            return []
    
    def queue_parent_email(self, people_id, person_name, age, meeting_name, template_name):
        """Queue a parent email notification"""
        # ::STEP:: Add email to queue
        if age >= 18:
            return False
            
        parents = self.get_parent_emails(people_id, age)
        if parents:
            self.email_queue.append({
                'people_id': people_id,
                'person_name': person_name,
                'meeting_name': meeting_name,
                'parents': parents,
                'template': template_name
            })
            return True
        return False
        
    def send_parent_email(self, parent_email, parent_name, child_name, meeting_name, template_name):
        try:
            print "<!-- DEBUG: NEW send_parent_email called for {0} -->".format(parent_name)
            
            # Get parent's PeopleId
            sql = "SELECT TOP 1 PeopleId FROM People WHERE EmailAddress = '{0}'".format(parent_email.replace("'", "''"))
            result = self.q.QuerySqlTop1(sql)
            
            if not result or not hasattr(result, 'PeopleId'):
                print "<!-- DEBUG: ERROR - Could not find parent in database! -->".format()
                return False
                
            parent_people_id = result.PeopleId
            print "<!-- DEBUG: Parent PeopleId: {0} -->".format(parent_people_id)
            
            # For parents, we need the child's name in the email
            # Since TouchPoint's {name} will be the parent's name, we need a workaround
            
            if template_name and template_name != 'generic':
                print "<!-- DEBUG: Creating custom parent email -->".format()
                
                try:
                    # Create a custom email for parents since we need child's name
                    subject = "Check-In Notification"
                    
                    # Get organization name for the email
                    org_name = meeting_name
                    current_time = self.model.DateTime.ToString("h:mm tt")
                    current_date = self.model.DateTime.ToString("MMMM d, yyyy")
                    church_name = self.model.Setting('ChurchName', 'Our Church')
                    
                    # Create email body with child's name
                    body = """
                    <p>Dear {0},</p>
                    
                    <p>{1} has been successfully checked in for {2}.</p>
                    
                    <p>This week:</p>
                    <p>Be in prayer</p>
                    <p>Here is the agenda</p>
                    <p>Watch for notifications</p>
                    
                    <p>If you have any questions or concerns, please contact the children's ministry team.</p>
                    
                    <p>Thank you,</p>
                    <p>{3} Children's Ministry</p>
                    """.format(
                        parent_name,
                        child_name,  # This is what we needed - the child's name
                        org_name,
                        church_name
                    )
                    
                    self.model.Email(
                        "peopleids={0}".format(parent_people_id),
                        parent_people_id,
                        parent_email,
                        parent_name,
                        subject,
                        body
                    )
                    
                    print "<!-- DEBUG: Custom parent email sent with child name -->".format()
                    return True
                    
                except Exception as e:
                    print "<!-- DEBUG: Custom parent email failed: {0} -->".format(str(e))
                    # Fall through to generic email
            
            # Send generic email as fallback
            print "<!-- DEBUG: Sending generic parent email -->".format()
            subject = "Check-In Notification"
            body = """
            <p>Dear {0},</p>
            <p>Your child {1} has been checked in to {2} at {3}.</p>
            <p>Thank you,<br>
            Church Check-In System</p>
            """.format(
                parent_name,
                child_name,
                meeting_name,
                self.model.DateTime.ToString("h:mm tt")
            )
            
            self.model.Email(
                "peopleids={0}".format(parent_people_id),
                parent_people_id,
                parent_email,
                parent_name,
                subject,
                body
            )
            
            print "<!-- DEBUG: Generic parent email sent -->".format()
            return True
            
        except Exception as e:
            print "<!-- DEBUG: Parent email completely failed: {0} -->".format(str(e))
            return False
        
    def send_adult_email(self, person_email, person_name, meeting_name, template_name):
        """Send email notification to an adult who checked in"""
        try:
            print "<!-- DEBUG: NEW send_adult_email called -->"
            print "<!-- DEBUG: To: {0} ({1}), Template: {2} -->".format(person_name, person_email, template_name)
            
            # Get the person's PeopleId
            sql = "SELECT TOP 1 PeopleId FROM People WHERE EmailAddress = '{0}'".format(person_email.replace("'", "''"))
            result = self.q.QuerySqlTop1(sql)
            
            if not result or not hasattr(result, 'PeopleId'):
                print "<!-- DEBUG: ERROR - Could not find person in database! -->"
                return False
                
            person_people_id = result.PeopleId
            print "<!-- DEBUG: Person PeopleId: {0} -->".format(person_people_id)
            
            if template_name and template_name != 'generic':
                print "<!-- DEBUG: Sending TouchPoint email for adult -->".format()
                
                try:
                    # Use TouchPoint's standard email system
                    # The template should use {CheckInPerson} and we'll tell the user to change it to {name}
                    self.model.EmailContent(
                        "peopleids={0}".format(person_people_id),
                        person_people_id,
                        person_email,
                        person_name,
                        template_name
                    )
                    
                    print "<!-- DEBUG: Standard TouchPoint adult email sent -->".format()
                    return True
                    
                except Exception as e:
                    print "<!-- DEBUG: TouchPoint adult email failed: {0} -->".format(str(e))
                    # Fall through to generic email
            
            # Send generic email as fallback
            print "<!-- DEBUG: Sending generic adult email -->".format()
            subject = "Check-In Confirmation"
            body = """
            <p>Dear {0},</p>
            <p>You have been successfully checked in to {1} at {2}.</p>
            <p>Thank you,<br>
            Church Check-In System</p>
            """.format(
                person_name,
                meeting_name,
                self.model.DateTime.ToString("h:mm tt")
            )
            
            self.model.Email(
                "peopleids={0}".format(person_people_id),
                person_people_id,
                person_email,
                person_name,
                subject,
                body
            )
            
            print "<!-- DEBUG: Generic adult email sent -->".format()
            return True
            
        except Exception as e:
            print "<!-- DEBUG: Adult email completely failed: {0} -->".format(str(e))
            return False
    
    def queue_adult_email(self, people_id, person_name, person_email, meeting_name, template_name):
        """Queue an adult email notification"""
        # ::STEP:: Add adult email to queue
        self.email_queue.append({
            'people_id': people_id,
            'person_name': person_name,
            'person_email': person_email,
            'meeting_name': meeting_name,
            'template': template_name,
            'type': 'adult'
        })
        return True
    
    def send_queued_emails(self):
        """Send all queued emails (both parent and adult)"""
        # ::STEP:: Process email queue
        if not self.email_queue:
            return 0
            
        sent_count = 0
        for email_data in self.email_queue:
            try:
                if email_data.get('type') == 'adult':
                    # Send to adult
                    if self.send_adult_email(
                        email_data['person_email'],
                        email_data['person_name'],
                        email_data['meeting_name'],
                        email_data['template']
                    ):
                        sent_count += 1
                else:
                    # Send to parents
                    for parent in email_data['parents']:
                        if self.send_parent_email(
                            parent['email'],
                            parent['name'],
                            email_data['person_name'],
                            email_data['meeting_name'],
                            email_data['template']
                        ):
                            sent_count += 1
            except Exception as e:
                print "<!-- DEBUG: Error sending queued email: {0} -->".format(str(e))
                
        # Clear the queue
        self.email_queue = []
        return sent_count
    
# ::START:: Check-In Manager Class
class CheckInManager:
    """Manages check-in operations and state"""
    
    def __init__(self, model, q):
        self.model = model
        self.q = q
        self.today = self.model.DateTime.Date
        self.selected_meetings = []
        self.all_meetings_today = []
        self.last_check_in_time = None  # Track last check-in time for stats refresh
        self.email_manager = EmailManager(model, q)  # Initialize email manager
        
    def get_todays_meetings(self):
        """Get all meetings scheduled for today including program information"""
        # ::STEP:: Query today's meetings
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
            AND o.DivisionId = os.DivId
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
        
    def debug_meeting_and_org_info(self, people_id, meeting_id):
        """Debug function to see meeting and organization details"""
        print "<!-- === DEBUG MEETING AND ORG INFO === -->"
        
        # Get meeting details
        meeting_sql = """
            SELECT m.MeetingId, m.OrganizationId, m.MeetingDate, m.DidNotMeet,
                   o.OrganizationName, o.OrganizationStatusId
            FROM Meetings m
            JOIN Organizations o ON m.OrganizationId = o.OrganizationId
            WHERE m.MeetingId = {0}
        """.format(meeting_id)
        
        meeting_info = self.q.QuerySqlTop1(meeting_sql)
        if meeting_info:
            print "<!-- Meeting: ID={0}, OrgID={1}, Date={2}, DidNotMeet={3}, OrgName='{4}', OrgStatus={5} -->".format(
                meeting_info.MeetingId, meeting_info.OrganizationId, meeting_info.MeetingDate,
                meeting_info.DidNotMeet, meeting_info.OrganizationName, meeting_info.OrganizationStatusId)
        else:
            print "<!-- ERROR: No meeting found with ID {0} -->".format(meeting_id)
            return
        
        # Check person's membership in this organization
        membership_sql = """
            SELECT om.PeopleId, om.OrganizationId, om.MemberTypeId, om.EnrollmentDate, om.InactiveDate,
                   p.Name, mt.Description as MemberType
            FROM OrganizationMembers om
            JOIN People p ON om.PeopleId = p.PeopleId
            LEFT JOIN lookup.MemberType mt ON om.MemberTypeId = mt.Id
            WHERE om.PeopleId = {0} AND om.OrganizationId = {1}
        """.format(people_id, meeting_info.OrganizationId)
        
        membership_info = self.q.QuerySqlTop1(membership_sql)
        if membership_info:
            print "<!-- Membership: Person='{0}', MemberType='{1}', Enrolled={2}, Inactive={3} -->".format(
                membership_info.Name, membership_info.MemberType, 
                membership_info.EnrollmentDate, membership_info.InactiveDate)
        else:
            print "<!-- ERROR: Person {0} is NOT a member of organization {1} -->".format(people_id, meeting_info.OrganizationId)
        
        # Check current user permissions
        print "<!-- Current User: {0} -->".format(self.model.UserName)
        
        print "<!-- === END DEBUG === -->"
    
    def get_person_age(self, people_id):
        """Get a person's age"""
        # ::STEP:: Query person's age
        sql = """
            SELECT Age FROM People WHERE PeopleId = {0}
        """.format(people_id)
        
        try:
            result = self.q.QuerySqlTop1(sql)
            if result and hasattr(result, 'Age') and result.Age is not None:
                return int(result.Age)
        except:
            pass
        return 99  # Default to adult if age unknown
    
    def get_people_by_filter(self, alpha_filter, search_term="", meeting_ids=None, page=1, page_size=PAGE_SIZE):
        """Get people by alpha filter or search term who are members of the selected orgs"""
        # ::STEP:: Filter and retrieve people
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
        
        # Start building the SQL query for the actual data - NOW INCLUDING AGE
        sql = """
            SELECT DISTINCT p.PeopleId, p.Name, p.FamilyId, p.LastName, p.FirstName, p.Age
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
            
            # Get age, defaulting to 99 if not available
            age = result.Age if hasattr(result, 'Age') and result.Age is not None else 99
            
            person = PersonInfo(
                result.PeopleId,
                result.Name,
                result.FamilyId,
                person_org_ids,
                False,  # not checked in
                age
            )
            people.append(person)
            
        return people, total_count
            
    def get_people_by_meeting_ids(self, meeting_ids, alpha_filter="All", search_term="", page=1, page_size=PAGE_SIZE):
        """Alternative method to get people directly by meeting IDs when org mapping fails"""
        # ::STEP:: Get people by meeting IDs
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
        
        # FIX: When using DISTINCT, include ORDER BY columns in the SELECT list - NOW INCLUDING AGE
        sql = """
            SELECT DISTINCT p.PeopleId, p.Name, p.FamilyId, p.LastName, p.FirstName, p.Age
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
                
                # Get age, defaulting to 99 if not available
                age = result.Age if hasattr(result, 'Age') and result.Age is not None else 99
                
                if person_org_ids:  # Only include if they're a member of at least one org
                    person = PersonInfo(
                        result.PeopleId,
                        result.Name,
                        result.FamilyId,
                        person_org_ids,
                        False,  # not checked in
                        age
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
            SELECT DISTINCT p.PeopleId, p.Name, p.Age
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
            
            # Get age, defaulting to 99 if not available
            age = result.Age if hasattr(result, 'Age') and result.Age is not None else 99
            
            person = PersonInfo(
                result.PeopleId,
                result.Name,
                family_id,
                person_org_ids,
                False,  # not checked in
                age
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
        # ::STEP:: Perform check-in
        try:
            print "<!-- check_in_person START: people_id={0}, meeting_id={1} -->".format(people_id, meeting_id)
            
            # Convert IDs to integers
            people_id_int = int(people_id)
            meeting_id_int = int(meeting_id)
            
            print "<!-- check_in_person: Converted to integers: people_id={0}, meeting_id={1} -->".format(people_id_int, meeting_id_int)
            
            # Check attendance BEFORE the API call
            pre_check_sql = """
                SELECT AttendanceFlag
                FROM Attend 
                WHERE PeopleId = {0} AND MeetingId = {1}
                AND AttendanceFlag = 1
            """.format(people_id_int, meeting_id_int)
            
            pre_result = self.q.QuerySqlTop1(pre_check_sql)
            already_checked_in = pre_result is not None
            
            print "<!-- Pre-check: Already checked in = {0} -->".format(already_checked_in)
            
            # Make the API call
            print "<!-- Calling model.EditPersonAttendance({0}, {1}, True) -->".format(meeting_id_int, people_id_int)
            
            result = self.model.EditPersonAttendance(meeting_id_int, people_id_int, True)
            
            print "<!-- EditPersonAttendance raw result: '{0}' -->".format(str(result))
            print "<!-- Result type: {0} -->".format(type(result).__name__)
            
            # Check attendance AFTER the API call to see if it worked
            post_check_sql = """
                SELECT AttendanceFlag, CreatedDate
                FROM Attend 
                WHERE PeopleId = {0} AND MeetingId = {1}
                AND AttendanceFlag = 1
            """.format(people_id_int, meeting_id_int)
            
            post_result = self.q.QuerySqlTop1(post_check_sql)
            now_checked_in = post_result is not None
            
            print "<!-- Post-check: Now checked in = {0} -->".format(now_checked_in)
            
            if post_result:
                print "<!-- Attendance record: Flag={0}, Created={1} -->".format(
                    post_result.AttendanceFlag, post_result.CreatedDate)
            
            # Determine success based on actual database state, not API response string
            if now_checked_in:
                if already_checked_in:
                    print "<!-- SUCCESS: Person was already checked in -->".format()
                else:
                    print "<!-- SUCCESS: Person was successfully checked in -->".format()
                return True
            else:
                print "<!-- FAILURE: Person is still not checked in after API call -->".format()
                
                # Try to understand why it failed by checking the result string
                result_str = str(result).lower()
                if 'already' in result_str:
                    print "<!-- API says already checked in, but database doesn't show it -->".format()
                elif 'error' in result_str or 'fail' in result_str:
                    print "<!-- API returned error: {0} -->".format(result_str)
                else:
                    print "<!-- Unknown API response: {0} -->".format(result_str)
                
                return False
            
        except Exception as e:
            print "<!-- ERROR in check_in_person: {0} -->".format(str(e))
            import traceback
            print "<!-- TRACEBACK: {0} -->".format(traceback.format_exc().replace('\n', ' | '))
            return False
    
    def check_in_person_with_email(self, people_id, meeting_id, person_name, email_template):
        """Check in a person and send appropriate email notification"""
        # ::STEP:: Check-in with notification
        try:
            print "<!-- check_in_person_with_email START -->"
            
            # Add this debug call
            self.debug_meeting_and_org_info(people_id, meeting_id)
            
            # Perform the check-in
            success = self.check_in_person(people_id, meeting_id)
            print "<!-- Parameters: people_id={0}, meeting_id={1}, person_name={2}, email_template={3} -->".format(
                people_id, meeting_id, person_name, email_template)
            
            # Perform the check-in
            success = self.check_in_person(people_id, meeting_id)
            print "<!-- Check-in API result: {0} -->".format(success)
            
            if not success:
                print "<!-- Check-in failed, not sending email -->"
                return False
                
            # Check if we should send email - CRITICAL CHECK
            if not email_template or email_template == 'none':
                print "<!-- No email template selected (template='{0}'), skipping email -->".format(email_template)
                return success
                
            print "<!-- Email template is '{0}', proceeding with email -->".format(email_template)
            
            # Get person's details
            sql = """
                SELECT 
                    PeopleId,
                    Name,
                    Age,
                    EmailAddress,
                    FamilyId,
                    DATEDIFF(hour, BDate, GETDATE())/8766.0 AS CalculatedAge
                FROM People 
                WHERE PeopleId = {0}
            """.format(people_id)
            
            person_result = self.q.QuerySqlTop1(sql)
            
            if not person_result:
                print "<!-- ERROR: Could not find person with ID {0} -->".format(people_id)
                return success
                
            # Get age and email
            age = 99  # Default to adult
            if hasattr(person_result, 'Age') and person_result.Age is not None:
                age = int(person_result.Age)
            elif hasattr(person_result, 'CalculatedAge') and person_result.CalculatedAge is not None:
                age = int(person_result.CalculatedAge)
                
            person_email = person_result.EmailAddress if hasattr(person_result, 'EmailAddress') else None
            
            print "<!-- Person details: Age={0}, Email={1}, FamilyId={2} -->".format(
                age, person_email, person_result.FamilyId if hasattr(person_result, 'FamilyId') else 'None')
            
            # Get meeting name
            meeting_name = "the event"
            for meeting in self.all_meetings_today:
                if str(meeting.meeting_id) == str(meeting_id):
                    meeting_name = meeting.org_name
                    break
            
            print "<!-- Meeting name: {0} -->".format(meeting_name)
            
            # Determine who to send email to
            email_sent = False
            
            if age < 18:
                print "<!-- Person is UNDER 18 (age={0}), looking for parents -->".format(age)
                
                # Get parents
                parents = self.email_manager.get_parent_emails(people_id, age)
                print "<!-- Found {0} parents with email addresses -->".format(len(parents))
                
                if not parents:
                    print "<!-- No parents found with valid email addresses -->".format()
                else:
                    # Send to each parent
                    for i, parent in enumerate(parents):
                        print "<!-- Sending to parent {0}: {1} ({2}) -->".format(
                            i+1, parent['name'], parent['email'])
                        
                        if self.email_manager.send_parent_email(
                            parent['email'],
                            parent['name'],
                            person_name,
                            meeting_name,
                            email_template
                        ):
                            email_sent = True
                            print "<!-- Email successfully sent to parent -->".format()
                        else:
                            print "<!-- Failed to send email to parent -->".format()
            else:
                print "<!-- Person is 18+ (age={0}), sending to them directly -->".format(age)
                
                if not person_email:
                    print "<!-- ERROR: Adult has no email address -->".format()
                else:
                    print "<!-- Sending email to: {0} ({1}) -->".format(person_name, person_email)
                    
                    if self.email_manager.send_adult_email(
                        person_email,
                        person_name,
                        meeting_name,
                        email_template
                    ):
                        email_sent = True
                        print "<!-- Email successfully sent to adult -->".format()
                    else:
                        print "<!-- Failed to send email to adult -->".format()
            
            print "<!-- check_in_person_with_email COMPLETE - email_sent={0} -->".format(email_sent)
            return success
            
        except Exception as e:
            print "<!-- ERROR in check_in_person_with_email: {0} -->".format(str(e))
            import traceback
            print "<!-- Traceback: {0} -->".format(traceback.format_exc().replace('\n', ' '))
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
        # ::STEP:: Calculate statistics
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
        
        # Then get paginated data - NOW INCLUDING AGE
        sql = """
            SELECT DISTINCT p.PeopleId, p.Name, p.FamilyId, p.LastName, p.FirstName, p.Age
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
            
            # Get age, defaulting to 99 if not available
            age = result.Age if hasattr(result, 'Age') and result.Age is not None else 99
            
            person = PersonInfo(
                result.PeopleId,
                result.Name,
                result.FamilyId,
                person_org_ids,
                True,  # checked in
                age
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

# ::START:: Rendering Functions
# Rendering functions
def render_meeting_selection(check_in_manager):
    """Render the meeting selection page with email template selection"""
    # ::STEP:: Display meeting selection UI
    meetings = check_in_manager.get_todays_meetings()
    
    # Get available email templates
    email_templates = check_in_manager.email_manager.get_available_email_templates()
    
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
    
    # Get selected email template
    selected_email_template = getattr(check_in_manager.model.Data, 'email_template', DEFAULT_EMAIL_TEMPLATE)
    
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
            .email-template-section {
                background-color: #f0f8ff;
                border: 1px solid #b0d4ff;
                border-radius: 4px;
                padding: 15px;
                margin-bottom: 20px;
            }
            .email-template-section h4 {
                margin-top: 0;
                color: #0066cc;
            }
            .help-text {
                font-size: 12px;
                color: #666;
                margin-top: 5px;
                font-style: italic;
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
            
            <!-- Email Template Selection Section -->
            <div class="email-template-section">
                <h4>Parent Email Notifications</h4>
                <div class="form-group">
                    <label for="email_template">Email Template for Parents (when checking in children under 18):</label>
                    <select class="form-control" id="email_template" name="email_template">
                        <option value="none" {1}>No Email Notification</option>
                        <option value="generic" {2}>Generic Email (System Generated)</option>
    """.format(script_name, 
              'selected="selected"' if selected_email_template == 'none' else '',
              'selected="selected"' if selected_email_template == 'generic' else '')
    
    # Add available email templates
    for template in email_templates:
        selected = 'selected="selected"' if template['name'] == selected_email_template else ''
        print '<option value="{0}" {2}>{1}</option>'.format(
            template['name'], 
            template['title'],
            selected
        )
    
    print """
                    </select>
                    <div class="help-text">
                        Notifications will be sent automatically when checking in:
                        <ul style="margin:5px 0; padding-left:20px;">
                            <li>For children (under 18): Email sent to parents/guardians</li>
                            <li>For adults (18+): Email sent directly to them</li>
                        </ul>
                    </div>                      
                </div>
            </div>
            
            <h4>Available Meetings for {0}</h4>
            
            <p>
                <button type="button" class="btn btn-sm btn-primary" id="selectAllBtn">Select All</button>
                <button type="button" class="btn btn-sm btn-default" id="deselectAllBtn">Deselect All</button>
            </p>
            
            <div class="row" style="max-height: 400px; overflow-y: auto;">
    """.format(check_in_manager.today.ToString(DATE_FORMAT))
    
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
    
    # Get email template selection
    email_template = getattr(model.Data, 'email_template', 'none')
    email_template_input = '<input type="hidden" name="email_template" value="{0}">'.format(email_template)
    
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
            '{6}'
            '<button type="submit" style="padding:6px 12px; border:1px solid #ddd; background-color:#f8f9fa; border-radius:4px; cursor:pointer; margin:0;" onclick="showLoading()">&laquo;</button>'
            '</form>'.format(
                script_name, page - 1, view_mode, alpha_filter, search_term, meeting_id_inputs, email_template_input
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
            '{5}'
            '<button type="submit" style="padding:6px 12px; border:1px solid #ddd; background-color:#f8f9fa; border-radius:4px; cursor:pointer; margin:0;" onclick="showLoading()">1</button>'
            '</form>'.format(
                script_name, view_mode, alpha_filter, search_term, meeting_id_inputs, email_template_input
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
                '{6}'
                '<button type="submit" style="padding:6px 12px; border:1px solid #ddd; background-color:#f8f9fa; border-radius:4px; cursor:pointer; margin:0;" onclick="showLoading()">{1}</button>'
                '</form>'.format(
                    script_name, i, view_mode, alpha_filter, search_term, meeting_id_inputs, email_template_input
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
            '{6}'
            '<button type="submit" style="padding:6px 12px; border:1px solid #ddd; background-color:#f8f9fa; border-radius:4px; cursor:pointer; margin:0;" onclick="showLoading()">{1}</button>'
            '</form>'.format(
                script_name, total_pages, view_mode, alpha_filter, search_term, meeting_id_inputs, email_template_input
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
            '{6}'
            '<button type="submit" style="padding:6px 12px; border:1px solid #ddd; background-color:#f8f9fa; border-radius:4px; cursor:pointer; margin:0;" onclick="showLoading()">&raquo;</button>'
            '</form>'.format(
                script_name, page + 1, view_mode, alpha_filter, search_term, meeting_id_inputs, email_template_input
            )
        )
    else:
        pagination.append(
            '<span style="display:inline-block; padding:6px 12px; border:1px solid #ddd; background-color:#f8f9fa; border-radius:4px; color:#999; margin:0 2px;">&raquo;</span>'
        )
        
    pagination.append('</div>')
    return "\n".join(pagination)

# ::START:: AJAX Processing
def process_ajax_check_in(check_in_manager):
    """Process AJAX check-in requests without full page reload"""
    # ::STEP:: Handle AJAX check-in request
    try:
        # Get parameters
        person_id = getattr(check_in_manager.model.Data, 'person_id', None)
        meeting_id = getattr(check_in_manager.model.Data, 'meeting_id', None)
        person_name = getattr(check_in_manager.model.Data, 'person_name', '')
        email_template = getattr(check_in_manager.model.Data, 'email_template', 'none')
        
        print "<!-- DEBUG AJAX: person_id={0}, meeting_id={1}, email_template={2} -->".format(
            person_id, meeting_id, email_template)
        
        # Validate parameters
        if not person_id or not meeting_id:
            print '{{"success":false,"reason":"missing_parameters"}}'
            return True  # Return True to indicate AJAX was handled
            
        # Clean any comma-separated values
        if ',' in str(person_id):
            person_id = str(person_id).split(',')[0].strip()
            
        if ',' in str(meeting_id):
            meeting_id = str(meeting_id).split(',')[0].strip()
        
        # Convert to integers
        person_id_int = int(person_id)
        meeting_id_int = int(meeting_id)
        
        print "<!-- DEBUG AJAX: Calling check_in_person_with_email -->".format()
        
        # Perform the check-in with email notification
        success = check_in_manager.check_in_person_with_email(
            person_id_int, 
            meeting_id_int, 
            person_name, 
            email_template
        )
        
        print "<!-- DEBUG AJAX: check_in_person_with_email returned: {0} -->".format(success)
        
        # Force refresh of statistics
        check_in_manager.last_check_in_time = check_in_manager.model.DateTime
        
        # Process any queued emails if batch mode
        if PARENT_EMAIL_DELAY:
            sent_count = check_in_manager.email_manager.send_queued_emails()
            print "<!-- DEBUG AJAX: Sent {0} queued emails -->".format(sent_count)
        
        if success:
            print '{{"success":true}}'
        else:
            print '{{"success":false,"reason":"check_in_failed"}}'
            
        return True  # Return True to indicate AJAX was handled
            
    except Exception as e:
        print "<!-- ERROR AJAX: {0} -->".format(str(e))
        import traceback
        print "<!-- TRACEBACK AJAX: {0} -->".format(traceback.format_exc().replace('\n', ' '))
        print '{{"success":false,"reason":"general_error"}}'
        return True  # Return True to indicate AJAX was handled

def render_fastlane_check_in(check_in_manager):
    """Render the FastLane check-in page with fixed alignment and branding"""
    # ::STEP:: Display check-in UI
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
    email_template = getattr(check_in_manager.model.Data, 'email_template', 'none')
    
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
    
    # Add email template input
    email_template_input = '<input type="hidden" name="email_template" value="{0}">'.format(email_template)
    
    # ::STEP:: Process form actions
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
                
                # Direct API call with email notification
                try:
                    # Force refresh of statistics
                    check_in_manager.last_check_in_time = check_in_manager.model.DateTime
                    
                    result = check_in_manager.check_in_person_with_email(
                        int(person_id), 
                        int(meeting_id), 
                        person_name,
                        email_template
                    )
                    
                    if result:
                        flash_message = "Successfully checked in"
                        flash_name = person_name
                        
                        # Process any queued emails if batch mode
                        if PARENT_EMAIL_DELAY:
                            sent_count = check_in_manager.email_manager.send_queued_emails()
                            if sent_count > 0:
                                flash_message += " (parent notified)"
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
    
    # ::STEP:: Get people list
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
    
    # ::STEP:: Render alpha filters
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
            {5}
            <button type="submit" style="padding:4px 8px; font-size:12px; color:{6}; background-color:{7}; border:1px solid #dee2e6; border-radius:3px; cursor:pointer; min-width:42px;">{2}</button>
        </form>
        """.format(
            script_name,        # {0}
            view_mode,          # {1}
            alpha_block,        # {2}
            search_term,        # {3}
            meeting_id_inputs,  # {4}
            email_template_input, # {5}
            color,              # {6}
            bg_color            # {7}
        )
    
    alpha_filters_html += """
        </div>
    </div>
    """
    
    # ::STEP:: Generate people list
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
        
        # Show age indicator for minors
        age_indicator = ""
        if hasattr(person, 'age') and person.age is not None and person.age < 18:
            age_indicator = '<span style="font-size:11px; color:#ff6b6b; margin-left:8px;">(Age: {0})</span>'.format(person.age)
        
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
            <div id="person-{0}" class="person-card" style="padding:8px; background-color:#fff; border:1px solid #ddd; border-radius:4px;">
                <div style="display:flex; align-items:center; justify-content:space-between;">
                    <div style="flex:1; overflow:hidden; text-overflow:ellipsis;">
                        <div style="font-size:15px;">{1}{2}{3}</div>
                        {4}
                    </div>
            """.format(
                person.people_id,       # {0}
                person.name,            # {1}
                unique_identifier,      # {2}
                age_indicator,          # {3}
                org_display,            # {4}
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
                    <form method="post" action="/PyScriptForm/{5}" style="margin:0; padding:0;" onsubmit="document.getElementById('loadingIndicator').style.display='block';">
                        <input type="hidden" name="step" value="check_in">
                        <input type="hidden" name="action" value="remove_check_in">
                        <input type="hidden" name="person_id" value="{0}">
                        <input type="hidden" name="person_name" value="{1}">
                        <input type="hidden" name="meeting_id" value="{6}">
                        <input type="hidden" name="view_mode" value="checked_in">
                        <input type="hidden" name="alpha_filter" value="{7}">
                        <input type="hidden" name="search_term" value="{8}">
                        <input type="hidden" name="page" value="{9}">
                        {10}
                        <button type="submit" style="padding:3px 8px; font-size:12px; background-color:#dc3545; color:#fff; border:1px solid #dc3545; border-radius:3px; cursor:pointer; min-width:60px; height:28px; vertical-align:middle;">{11}</button>
                    </form>
                    """.format(
                        person.people_id,       # {0}
                        person.name,            # {1}
                        "",                     # {2} - not used
                        "",                     # {3} - not used
                        "",                     # {4} - not used
                        script_name,            # {5}
                        meeting_id,             # {6}
                        alpha_filter,           # {7}
                        search_term,            # {8}
                        current_page,           # {9}
                        email_template_input,   # {10}
                        button_label            # {11}
                    )
                    
                # Close the container
                item_html += """
                    </div>
                """
            else:
                # Just one meeting - use a single remove button
                meeting_id = attended_meeting_ids[0] if attended_meeting_ids else (meeting_ids[0] if meeting_ids else "")
                
                item_html += """
                    <form method="post" action="/PyScriptForm/{5}" style="margin:0; padding:0; flex-shrink:0;" onsubmit="document.getElementById('loadingIndicator').style.display='block';">
                        <input type="hidden" name="step" value="check_in">
                        <input type="hidden" name="action" value="remove_check_in">
                        <input type="hidden" name="person_id" value="{0}">
                        <input type="hidden" name="person_name" value="{1}">
                        <input type="hidden" name="meeting_id" value="{6}">
                        <input type="hidden" name="view_mode" value="checked_in">
                        <input type="hidden" name="alpha_filter" value="{7}">
                        <input type="hidden" name="search_term" value="{8}">
                        <input type="hidden" name="page" value="{9}">
                        {10}
                        <button type="submit" style="padding:3px 8px; font-size:12px; background-color:#dc3545; color:#fff; border:1px solid #dc3545; border-radius:3px; cursor:pointer; min-width:60px; height:28px; vertical-align:middle;">Remove</button>
                    </form>
                """.format(
                    person.people_id,       # {0}
                    person.name,            # {1}
                    "",                     # {2} - not used
                    "",                     # {3} - not used
                    "",                     # {4} - not used
                    script_name,            # {5}
                    meeting_id,             # {6}
                    alpha_filter,           # {7}
                    search_term,            # {8}
                    current_page,           # {9}
                    email_template_input    # {10}
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
            <div id="person-{0}" class="person-card" style="padding:8px; background-color:#fff; border:1px solid #ddd; border-radius:4px;">
                <div style="display:flex; align-items:center; justify-content:space-between;">
                    <div style="flex:1; overflow:hidden; text-overflow:ellipsis;">
                        <div style="font-size:15px;">{1}{2}{3}</div>
                        {4}
                    </div>
                    <div style="display:flex; flex-wrap:wrap; justify-content:flex-end; gap:5px;">
            """.format(
                person.people_id,       # {0}
                person.name,            # {1}
                unique_identifier,      # {2}
                age_indicator,          # {3}
                org_display,            # {4}
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
                
                # Create an inline AJAX check-in button with email support
                item_html += """
                <button onclick="(function(personId, meetingId, personName, emailTemplate) {{
                    // Mark this person's card as processing
                    var personCard = document.getElementById('person-' + personId);
                    if (personCard) {{
                        personCard.style.opacity = '0.7';
                        personCard.style.backgroundColor = '#f1f8ff';
                    }}
                    
                    // Create and show mini flash message
                    var miniFlash = document.createElement('div');
                    miniFlash.style.position = 'fixed';
                    miniFlash.style.bottom = '20px';
                    miniFlash.style.right = '20px';
                    miniFlash.style.backgroundColor = 'rgba(40,167,69,0.9)';
                    miniFlash.style.color = 'white';
                    miniFlash.style.padding = '10px 15px';
                    miniFlash.style.borderRadius = '4px';
                    miniFlash.style.boxShadow = '0 2px 10px rgba(0,0,0,0.2)';
                    miniFlash.style.zIndex = '9999';
                    miniFlash.style.transition = 'opacity 0.3s ease';
                    miniFlash.innerHTML = '<span>Processing...</span>';
                    document.body.appendChild(miniFlash);
                    
                    // Prepare form data
                    var formData = new FormData();
                    formData.append('step', 'check_in');
                    formData.append('action', 'ajax_check_in');
                    formData.append('person_id', personId);
                    formData.append('person_name', personName);
                    formData.append('meeting_id', meetingId);
                    formData.append('email_template', emailTemplate);
                    
                    // Create XMLHttpRequest
                    var xhr = new XMLHttpRequest();
                    xhr.open('POST', '/PyScriptForm/{5}', true);
                    
                    xhr.onload = function() {{
                        // Always hide the card if status is 200 - regardless of the response content
                        if (xhr.status === 200) {{
                            // Hide the person card immediately
                            if (personCard) {{
                                personCard.style.display = 'none';
                            }}
                            
                            // Show success message
                            miniFlash.innerHTML = '<span>' + personName + ' checked in!</span>';
                            
                            // Update the stats
                            var checkedInEl = document.getElementById('stat-checked-in');
                            var notCheckedInEl = document.getElementById('stat-not-checked-in');
                            
                            if (checkedInEl && notCheckedInEl) {{
                                var checkedIn = parseInt(checkedInEl.innerText || '0');
                                var notCheckedIn = parseInt(notCheckedInEl.innerText || '0');
                                
                                checkedInEl.innerText = (checkedIn + 1).toString();
                                notCheckedInEl.innerText = Math.max(0, notCheckedIn - 1).toString();
                            }}
                        }} else {{
                            // Only show error if HTTP status is not 200
                            miniFlash.style.backgroundColor = 'rgba(220,53,69,0.9)';
                            miniFlash.innerHTML = '<span>Error checking in ' + personName + '</span>';
                            
                            // Reset the person card
                            if (personCard) {{
                                personCard.style.opacity = '1';
                                personCard.style.backgroundColor = '#fff';
                            }}
                        }}
                        
                        // Remove flash after 1.5 seconds
                        setTimeout(function() {{
                            miniFlash.style.opacity = '0';
                            setTimeout(function() {{
                                if (document.body.contains(miniFlash)) {{
                                    document.body.removeChild(miniFlash);
                                }}
                            }}, 300);
                        }}, 1500);
                    }};
                    
                    xhr.send(formData);
                    return false;
                }})('{0}', '{6}', '{1}', '{10}'); return false;" style="padding:3px 8px; font-size:12px; background-color:#28a745; color:#fff; border:1px solid #28a745; border-radius:3px; cursor:pointer; min-width:60px; height:28px; vertical-align:middle;">{9}</button>
                """.format(
                    person.people_id,  # {0}
                    person.name.replace("'", "\\'").replace('"', '\\"'),  # {1} - escape quotes for JavaScript
                    "",                # {2} - not used
                    "",                # {3} - not used
                    "",                # {4} - not used
                    script_name,       # {5}
                    meeting_id,        # {6} - specific meeting ID
                    alpha_filter,      # {7}
                    search_term,       # {8}
                    meeting_name,      # {9} - button label
                    email_template     # {10} - email template selection
                )
            
            # Close the containers
            item_html += """
                    </div>
                </div>
            </div>
            """
        
        people_list_html.append(item_html)
    
    people_list_html.append("</div>")
    
    # ::STEP:: Create stats bar
    # Create the compact stats bar
    stats_bar = """
    <div style="display:flex; justify-content:space-between; background-color:#f8f9fa; border:1px solid #dee2e6; border-radius:4px; padding:8px; margin-bottom:10px; font-size:13px;">
        <div><strong id="stat-checked-in" style="color:#28a745;">{0}</strong> Checked In</div>
        <div><strong id="stat-not-checked-in" style="color:#007bff;">{1}</strong> Not Checked In</div>
        <div><strong id="stat-total">{2}</strong> Total</div>
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
    
    # ::STEP:: Render HTML page
    # Simple HTML with FastLane branding and compact layout
    print """<!DOCTYPE html>
<html>
<head>
    <title>FastLane Check-In</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        .person-card {{
            transition: all 0.3s ease;
        }}
        .person-card.processing {{
            opacity: 0.7;
            background-color: #f1f8ff !important;
        }}
        .person-card.checked-in {{
            transform: translateX(100%);
            opacity: 0;
        }}
        .mini-flash {{
            position: fixed;
            bottom: 20px;
            right: 20px;
            background-color: rgba(40,167,69,0.9);
            color: white;
            padding: 10px 15px;
            border-radius: 4px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            z-index: 9999;
            transition: opacity 0.3s ease;
        }}
        .mini-flash .success {{
            color: white;
        }}
        .mini-flash .error {{
            color: #ffcccc;
        }}
    </style>
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
                    {13}
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
                {13}
                <div style="display:flex;">
                    <input type="text" name="search_term" value="{5}" placeholder="Search by name..." style="flex-grow:1; height:36px; padding:4px 10px; border:1px solid #ddd; border-radius:4px 0 0 4px; font-size:14px;">
                    <button type="submit" style="height:36px; padding:0 12px; background-color:#007bff; color:white; border:1px solid #007bff; border-radius:0 4px 4px 0; cursor:pointer; white-space:nowrap;" onclick="document.getElementById('loadingIndicator').style.display='block';">
                        Search
                    </button>
                </div>
            </form>
            
            <!-- Alpha filters - compact -->
            {9}
            
            <!-- Parent email notification indicator -->
            {10}
            
            <!-- People grid layout with fixed alignment -->
            {11}
            
            <!-- Pagination -->
            <div style="text-align:center;">
                {12}
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
        # Check for both not_checked_in and the displayed text
        "checked_in" if (view_mode == "not_checked_in" or view_mode == "") else "not_checked_in", # {7} - opposite mode for toggle
        # Check for both not_checked_in and the displayed text
        "Checked In List" if (view_mode == "not_checked_in" or view_mode == "") else "Check-In Page", # {8} - toggle button text
        alpha_filters_html,       # {9}
        # Parent email notification indicator
        '<div style="background-color:#e7f5ff; border:1px solid #4dabf7; border-radius:3px; padding:5px 8px; margin-bottom:10px; font-size:12px;"><strong>Parent Notifications:</strong> ' + (email_template if email_template != 'none' else 'Disabled') + '</div>' if email_template != 'none' else '', # {10}
        "\n".join(people_list_html), # {11}
        pagination,               # {12}
        email_template_input      # {13}
    )
    print """<!-- DEBUG: view_mode_raw='{0}' -->""".format(view_mode)
    
    # Send any queued parent emails if in batch mode
    if PARENT_EMAIL_DELAY and check_in_manager.email_manager.email_queue:
        sent_count = check_in_manager.email_manager.send_queued_emails()
        if sent_count > 0:
            print "<!-- DEBUG: Sent {0} parent notification emails -->".format(sent_count)
    
    return True  # Indicate successful rendering

def render_checked_in_view(check_in_manager):
    """Render the checked-in people view"""
    # Similar to render_check_in_page but displays already checked-in people
    # This allows removing individual check-ins if needed
    return render_fastlane_check_in(check_in_manager)

# ::START:: Updated Main Function
def main():
    """Main entry point for the FastLane Check-In system"""
    try:
        # ::STEP:: Initialize system
        check_in_manager = CheckInManager(model, q)
        
        # ::STEP:: Clean input data
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
        
        # Get email template selection
        email_template = getattr(model.Data, 'email_template', 'none')
        
        # Store cleaned values
        model.Data.cleaned_step = clean_step
        model.Data.cleaned_action = clean_action
        model.Data.cleaned_person_id = clean_person_id
        model.Data.cleaned_meeting_id = clean_meeting_id
        model.Data.email_template = email_template

        # ::STEP:: Handle AJAX requests - FIXED VERSION
        # Check for AJAX request and handle it, then return early
        if clean_action == 'ajax_check_in':
            ajax_handled = process_ajax_check_in(check_in_manager)
            if ajax_handled:
                return  # Exit normally without sys.exit()
        
        # ::STEP:: Process other form actions (rest of your existing code)
        if clean_action == 'single_direct_check_in':
            # Do direct API call with email notification
            if clean_person_id and clean_meeting_id:
                try:
                    person_name = getattr(model.Data, 'person_name', '')
                    # Direct API call with email
                    result = check_in_manager.check_in_person_with_email(
                        int(clean_person_id), 
                        int(clean_meeting_id), 
                        person_name,
                        email_template
                    )
                    # Update last check-in time to force stats refresh
                    check_in_manager.last_check_in_time = check_in_manager.model.DateTime
                    
                    # Process any queued emails if batch mode
                    if PARENT_EMAIL_DELAY:
                        check_in_manager.email_manager.send_queued_emails()
                except Exception as e:
                    print "<div style='color:red;'>Error: " + str(e) + "</div>"
        
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
                    person_name = getattr(model.Data, 'person_name', '')
                    
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
                                
                                # Check in with email notification
                                result = check_in_manager.check_in_person_with_email(
                                    person_id_int, 
                                    meeting_id_int, 
                                    person_name,
                                    email_template
                                )
                                
                                if result:
                                    check_in_count += 1
                                else:
                                    print "<!-- DEBUG: Check-in failed for meeting {0} -->".format(meeting_id)
                                    success = False
                            except Exception as e:
                                print "<!-- DEBUG: Check-in error for meeting {0}: {1} -->".format(meeting_id, str(e))
                                success = False
                    
                    # Update last check-in time to force stats refresh
                    check_in_manager.last_check_in_time = check_in_manager.model.DateTime
                    
                    # Process any queued emails if batch mode
                    if PARENT_EMAIL_DELAY:
                        sent_count = check_in_manager.email_manager.send_queued_emails()
                        if sent_count > 0:
                            print "<!-- DEBUG: Sent {0} parent notification emails -->".format(sent_count)
                    
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
        
        # ::STEP:: Render appropriate page
        # Determine which page to render
        if clean_step == 'check_in':
            render_fastlane_check_in(check_in_manager)
        else:
            render_meeting_selection(check_in_manager)
            
    except Exception as e:
        # ::STEP:: Handle errors
        import traceback
        print "<h2>Error</h2>"
        print "<p>An error occurred: {0}</p>".format(str(e))
        print "<pre>"
        traceback.print_exc()
        print "</pre>"
        
        # Link to go back
        script_name = get_script_name()
        print '<a href="/PyScript/{0}">Back to Check-In</a>'.format(script_name)

# ::START:: Script Entry Point
# Call the main function when the script runs
if __name__ == "__main__":
    main()
else:
    # TouchPoint doesn't use __main__, so always call main()
    main()
