"""
FastLane Check-In System for Touchpoint - Enhanced Version

This script provides a fast, efficient check-in system for high-volume events (100-2500+ people).

Features:
- Meeting selection for current day with program filtering
- Optimized for speed with bulk actions and AJAX check-ins
- Multiple person selection for efficient check-in
- Quick alpha-based filtering and search
- Family-based check-in capabilities
- One-click check-in process with visual feedback
- Real-time stats and metrics with auto-refresh
- Check-in correction functionality
- Parent email notifications for attendees under 18
- Subgroup filtering support for targeted check-ins
- Outstanding balance warnings for financial tracking
- Mobile-responsive design
- Uniquely scoped CSS to prevent conflicts with TouchPoint ChMS

--Upload Instructions Start--
To upload code to Touchpoint, use the following steps:
1. Click Admin ~ Advanced ~ Special Content ~ Python
2. Click New Python Script File
3. Name the Python FastLaneCheckIn (case sensitive) and paste all this code
4. Test and optionally add to menu
--Upload Instructions End--

Configuration Options:
- ENABLE_SUBGROUP_FILTERING: Enable/disable subgroup features
- ENABLE_BALANCE_DISPLAY: Enable/disable balance warnings  
- BALANCE_WARNING_THRESHOLD: Minimum balance to show warning
- PARENT_EMAIL_DELAY: Queue emails for batch processing
- DEFAULT_EMAIL_TEMPLATE: Default template for notifications
- PAGE_SIZE: Number of people to show per page

Written By: Ben Swaby
Email: bswaby@fbchtn.org
Enhanced with CSS fixes and improved workflow structure
"""

import traceback
import json
import datetime
import re
from System import DateTime
from System.Collections.Generic import List

# ::START:: Configuration and Constants
# Configuration constants - modify these for your church's needs
#Script_Name = "FastLaneCheckIn2"
PAGE_SIZE = 500  # Number of people to show per page - smaller for faster loading
DATE_FORMAT = "M/d/yyyy"
ATTEND_FLAG = 1  # Present flag for attendance

# Alpha filtering blocks for name search - all possible configurations
ALPHA_BLOCKS_CONFIG = {
    3: ["A-H", "I-P", "Q-Z", "All"],
    4: ["A-F", "G-L", "M-R", "S-Z", "All"],
    5: ["A-E", "F-J", "K-O", "P-T", "U-Z", "All"],
    6: ["A-D", "E-H", "I-L", "M-P", "Q-T", "U-Z", "All"],
    7: ["A-C", "D-F", "G-I", "J-M", "N-Q", "R-U", "V-Z", "All"],
    8: ["A-C", "D-F", "G-I", "J-L", "M-O", "P-R", "S-U", "V-Z", "All"]
}

# Default alpha block configuration
DEFAULT_ALPHA_BLOCK_COUNT = 4
ALPHA_BLOCKS = ALPHA_BLOCKS_CONFIG[DEFAULT_ALPHA_BLOCK_COUNT]

# Feature Configuration - Set to False to disable features
ENABLE_SUBGROUP_FILTERING = True  # Set to False to disable subgroup features
ENABLE_BALANCE_DISPLAY = True     # Set to False to disable balance features
BALANCE_WARNING_THRESHOLD = 0.01  # Minimum balance to show warning
SUBGROUP_DISPLAY_LIMIT = 3        # Maximum subgroups to show in person card

# Email Configuration
PARENT_EMAIL_DELAY = False  # Set to True to queue emails for batch processing
DEFAULT_EMAIL_TEMPLATE = "CheckInParentNotification"  # Default email template name

# CSS prefix to prevent conflicts with TouchPoint ChMS
CSS_PREFIX = "flci"  # FastLane Check-In prefix

# ::START:: Helper Functions
def get_script_name():
    """Get the current script name from the URL"""
    # ::STEP:: Extract script name from request
    script_name = "FastLaneCheckIn"  # Default fallback
    try:
        if hasattr(model, 'Request'):
            if hasattr(model.Request, 'RawUrl'):
                url_parts = model.Request.RawUrl.split('/')
                # Look for PyScript or PyScriptForm in the URL
                for i, part in enumerate(url_parts):
                    if part in ['PyScript', 'PyScriptForm'] and i + 1 < len(url_parts):
                        script_name = url_parts[i + 1].split('?')[0]
                        break
            elif hasattr(model.Request, 'Path'):
                # Alternative method using Path
                path_parts = model.Request.Path.split('/')
                for i, part in enumerate(path_parts):
                    if part in ['PyScript', 'PyScriptForm'] and i + 1 < len(path_parts):
                        script_name = path_parts[i + 1].split('?')[0]
                        break
    except Exception as e:
        # If all else fails, use the default
        pass
    return script_name

def print_header(title):
    """Print the HTML header with scoped CSS"""
    # ::STEP:: Output page header
    print """
    <div class="{0}-row">
        <div class="{0}-col-12">
            <h2 class="{0}-title">{1}</h2>
            <hr class="{0}-divider">
        </div>
    </div>
    """.format(CSS_PREFIX, title)

def get_org_name(check_in_manager, org_id):
    """Get the name of an organization by ID"""
    # ::STEP:: Look up organization name
    for meeting in check_in_manager.all_meetings_today:
        if str(meeting.org_id) == str(org_id):
            return meeting.org_name
    return "Organization {0}".format(org_id)
    
def render_alpha_filters(current_filter, alpha_block_count=DEFAULT_ALPHA_BLOCK_COUNT):
    """Render the alpha filter buttons with scoped CSS"""
    # ::STEP:: Generate alpha filter buttons
    # Get the appropriate alpha blocks based on selection
    alpha_blocks = ALPHA_BLOCKS_CONFIG.get(alpha_block_count, ALPHA_BLOCKS_CONFIG[DEFAULT_ALPHA_BLOCK_COUNT])
    
    buttons = []
    for alpha_block in alpha_blocks:
        active = '{0}-active'.format(CSS_PREFIX) if alpha_block == current_filter else ''
        buttons.append("""
            <label class="{0}-btn {0}-btn-default {1}">
                <input type="radio" name="alpha_filter" value="{2}" {3} onchange="this.form.submit()"> {2}
            </label>
        """.format(CSS_PREFIX, active, alpha_block, 'checked' if alpha_block == current_filter else ''))
    
    return """
        <div class="{0}-btn-group" data-toggle="buttons">
            {1}
        </div>
    """.format(CSS_PREFIX, "\n".join(buttons))

def get_scoped_css():
    """Generate uniquely scoped CSS to prevent conflicts with TouchPoint ChMS"""
    # ::STEP:: Create scoped CSS styles
    return """
    <style>
        /* FastLane Check-In Scoped Styles */
        .{0}-container {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            color: #333;
            background-color: #f9f9f9;
            margin: 0;
            padding: 0;
        }}
        
        .{0}-title {{
            margin-top: 0;
            color: #2c3e50;
            font-weight: 500;
            font-size: 24px;
        }}
        
        .{0}-divider {{
            border: 0;
            height: 1px;
            background-color: #e0e0e0;
            margin: 15px 0;
        }}
        
        .{0}-btn {{
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
            text-decoration: none;
            background-color: #fff;
            color: #333;
            border-color: #ccc;
        }}
        
        .{0}-btn:hover {{
            background-color: #f5f5f5;
        }}
        
        .{0}-btn-primary {{
            color: #fff;
            background-color: #3498db;
            border-color: #2980b9;
        }}
        
        .{0}-btn-primary:hover {{
            background-color: #2980b9;
        }}
        
        .{0}-btn-success {{
            color: #fff;
            background-color: #2ecc71;
            border-color: #27ae60;
        }}
        
        .{0}-btn-success:hover {{
            background-color: #27ae60;
        }}
        
        .{0}-btn-danger {{
            color: #fff;
            background-color: #dc3545;
            border-color: #dc3545;
        }}
        
        .{0}-btn-danger:hover {{
            background-color: #c82333;
        }}
        
        .{0}-btn-lg {{
            padding: 12px 20px;
            font-size: 16px;
        }}
        
        .{0}-btn-sm {{
            padding: 5px 10px;
            font-size: 12px;
        }}
        
        .{0}-btn-group {{
            display: inline-block;
        }}
        
        .{0}-btn-group .{0}-btn {{
            margin-right: 2px;
        }}
        
        .{0}-active {{
            background-color: #337ab7 !important;
            color: #fff !important;
            border-color: #337ab7 !important;
        }}
        
        .{0}-form-control {{
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
            box-sizing: border-box;
        }}
        
        .{0}-form-control:focus {{
            border-color: #3498db;
            outline: 0;
            box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(52,152,219,.6);
        }}
        
        .{0}-well {{
            min-height: 20px;
            padding: 19px;
            margin-bottom: 20px;
            background-color: #fff;
            border: 1px solid #e3e3e3;
            border-radius: 4px;
            box-shadow: 0 1px 3px rgba(0,0,0,.05);
        }}
        
        .{0}-checkbox {{
            position: relative;
            display: block;
            margin-top: 10px;
            margin-bottom: 10px;
        }}
        
        .{0}-checkbox label {{
            padding-left: 25px;
            cursor: pointer;
            display: block;
        }}
        
        .{0}-checkbox input[type=checkbox] {{
            position: absolute;
            margin-left: -20px;
            margin-top: 2px;
            cursor: pointer;
        }}
        
        .{0}-panel {{
            margin-bottom: 20px;
            background-color: #fff;
            border: 1px solid #ddd;
            border-radius: 4px;
            box-shadow: 0 1px 3px rgba(0,0,0,.05);
        }}
        
        .{0}-panel-heading {{
            padding: 12px 15px;
            background-color: #f7f7f7;
            border-bottom: 1px solid #ddd;
            border-top-left-radius: 3px;
            border-top-right-radius: 3px;
        }}
        
        .{0}-panel-body {{
            padding: 15px;
        }}
        
        .{0}-panel-title {{
            margin-top: 0;
            margin-bottom: 0;
            font-size: 16px;
            color: #333;
            font-weight: 500;
        }}
        
        .{0}-alert {{
            padding: 15px;
            margin-bottom: 20px;
            border: 1px solid transparent;
            border-radius: 4px;
        }}
        
        .{0}-alert-warning {{
            color: #8a6d3b;
            background-color: #fcf8e3;
            border-color: #faebcc;
        }}
        
        .{0}-alert-info {{
            color: #31708f;
            background-color: #d9edf7;
            border-color: #bce8f1;
        }}
        
        .{0}-alert-success {{
            color: #3c763d;
            background-color: #dff0d8;
            border-color: #d6e9c6;
        }}
        
        .{0}-row {{
            display: flex;
            flex-wrap: wrap;
            margin-right: -15px;
            margin-left: -15px;
        }}
        
        .{0}-col-6 {{
            position: relative;
            min-height: 1px;
            padding-right: 15px;
            padding-left: 15px;
            flex: 0 0 50%;
            max-width: 50%;
        }}
        
        .{0}-col-12 {{
            position: relative;
            min-height: 1px;
            padding-right: 15px;
            padding-left: 15px;
            flex: 0 0 100%;
            max-width: 100%;
        }}
        
        @media (max-width: 768px) {{
            .{0}-col-6 {{
                flex: 0 0 100%;
                max-width: 100%;
            }}
        }}
        
        .{0}-form-group {{
            margin-bottom: 15px;
        }}
        
        .{0}-loading {{
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(255,255,255,0.8);
            z-index: 10000;
            text-align: center;
            padding-top: 20%;
            font-size: 18px;
        }}
        
        .{0}-loading-spinner {{
            border: 6px solid #f3f3f3;
            border-top: 6px solid #3498db;
            border-radius: 50%;
            width: 50px;
            height: 50px;
            animation: {0}-spin 2s linear infinite;
            margin: 0 auto 20px auto;
        }}
        
        @keyframes {0}-spin {{
            0% {{ transform: rotate(0deg); }}
            100% {{ transform: rotate(360deg); }}
        }}
        
        .{0}-header {{
            background-color: #007bff;
            color: #fff;
            padding: 15px;
            margin: -20px -20px 20px -20px;
            border-radius: 0;
        }}
        
        .{0}-header-flex {{
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        
        .{0}-header-title {{
            margin: 0;
            font-size: 24px;
            color: #fff;
        }}
        
        .{0}-header-tagline {{
            font-size: 14px;
            font-style: italic;
            margin-top: 5px;
            color: #fff;
            opacity: 0.9;
        }}
        
        .{0}-stats-bar {{
            display: flex;
            justify-content: space-between;
            background-color: #f8f9fa;
            border: 1px solid #dee2e6;
            border-radius: 4px;
            padding: 10px;
            margin-bottom: 15px;
            font-size: 14px;
        }}
        
        .{0}-stat-item {{
            text-align: center;
        }}
        
        .{0}-stat-number {{
            font-size: 18px;
            font-weight: bold;
            display: block;
        }}
        
        .{0}-stat-checked {{
            color: #28a745;
        }}
        
        .{0}-stat-not-checked {{
            color: #007bff;
        }}
        
        .{0}-stat-total {{
            color: #6c757d;
        }}
        
        .{0}-people-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
            gap: 10px;
            margin-bottom: 20px;
        }}
        
        .{0}-person-card {{
            padding: 10px;
            background-color: #fff;
            border: 1px solid #ddd;
            border-radius: 4px;
            transition: all 0.3s ease;
        }}
        
        .{0}-person-card:hover {{
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}
        
        .{0}-person-card.{0}-processing {{
            opacity: 0.7;
            background-color: #f1f8ff !important;
        }}
        
        .{0}-person-card.{0}-checked-in {{
            transform: translateX(100%);
            opacity: 0;
        }}
        
        .{0}-person-card-flex {{
            display: flex;
            align-items: center;
            justify-content: space-between;
        }}
        
        .{0}-person-info {{
            flex: 1;
            overflow: hidden;
            text-overflow: ellipsis;
        }}
        
        .{0}-person-name {{
            font-size: 15px;
            font-weight: 500;
            margin-bottom: 2px;
        }}
        
        .{0}-person-details {{
            font-size: 12px;
            color: #666;
            margin-top: 2px;
        }}
        
        .{0}-person-buttons {{
            display: flex;
            flex-wrap: wrap;
            justify-content: flex-end;
            gap: 5px;
        }}
        
        .{0}-age-indicator {{
            font-size: 11px;
            color: #ff6b6b;
            margin-left: 8px;
        }}
        
        .{0}-balance-indicator {{
            background-color: #dc3545;
            color: white;
            padding: 2px 6px;
            border-radius: 10px;
            font-size: 10px;
            margin-left: 8px;
            font-weight: bold;
        }}
        
        .{0}-balance-warning {{
            background-color: #ffe6e6 !important;
            border-color: #ff9999 !important;
        }}
        
        .{0}-subgroup-display {{
            font-size: 11px;
            color: #007bff;
            margin-top: 2px;
        }}
        
        .{0}-unique-id {{
            font-size: 10px;
            color: #999;
        }}
        
        .{0}-flash-message {{
            position: fixed;
            top: 60px;
            left: 50%;
            transform: translateX(-50%);
            padding: 12px 20px;
            background-color: rgba(40,167,69,0.9);
            color: white;
            border-radius: 4px;
            z-index: 9999;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            transition: opacity 0.5s ease;
        }}
        
        .{0}-mini-flash {{
            position: fixed;
            bottom: 20px;
            right: 20px;
            background-color: rgba(40,167,69,0.9);
            color: white;
            padding: 10px 15px;
            border-radius: 4px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            z-index: 10000;
            transition: opacity 0.3s ease;
        }}
        
        .{0}-pagination {{
            display: flex;
            justify-content: center;
            align-items: center;
            margin: 20px 0;
        }}
        
        .{0}-pagination-btn {{
            padding: 6px 12px;
            margin: 0 2px;
            border: 1px solid #ddd;
            background-color: #f8f9fa;
            border-radius: 4px;
            cursor: pointer;
            color: #333;
            text-decoration: none;
            display: inline-block;
        }}
        
        .{0}-pagination-btn:hover {{
            background-color: #e9ecef;
        }}
        
        .{0}-pagination-current {{
            background-color: #007bff !important;
            color: white !important;
            border-color: #007bff !important;
        }}
        
        .{0}-pagination-disabled {{
            color: #999 !important;
            cursor: not-allowed !important;
        }}
        
        .{0}-search-container {{
            margin-bottom: 15px;
        }}
        
        .{0}-search-flex {{
            display: flex;
        }}
        
        .{0}-search-input {{
            flex-grow: 1;
            height: 38px;
            padding: 8px 12px;
            border: 1px solid #ddd;
            border-radius: 4px 0 0 4px;
            font-size: 14px;
        }}
        
        .{0}-search-btn {{
            height: 38px;
            padding: 0 15px;
            background-color: #007bff;
            color: white;
            border: 1px solid #007bff;
            border-radius: 0 4px 4px 0;
            cursor: pointer;
            white-space: nowrap;
        }}
        
        .{0}-feature-section {{
            background-color: #f0f8ff;
            border: 1px solid #b0d4ff;
            border-radius: 4px;
            padding: 15px;
            margin-bottom: 20px;
        }}
        
        .{0}-feature-title {{
            margin-top: 0;
            color: #0066cc;
            font-size: 16px;
        }}
        
        .{0}-help-text {{
            font-size: 12px;
            color: #666;
            margin-top: 5px;
            font-style: italic;
        }}
        
        .{0}-subgroup-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 8px;
            margin-top: 10px;
        }}
        
        .{0}-subgroup-item {{
            display: flex;
            align-items: center;
            padding: 8px;
            background-color: #f8f9fa;
            border-radius: 4px;
            cursor: pointer;
            transition: background-color 0.2s;
        }}
        
        .{0}-subgroup-item:hover {{
            background-color: #e9ecef;
        }}
        
        .{0}-content-panel {{
            background-color: #fff;
            border: 1px solid #ddd;
            border-radius: 4px;
            overflow: hidden;
            margin-bottom: 15px;
        }}
        
        .{0}-content-header {{
            padding: 10px 15px;
            background-color: #f8f9fa;
            border-bottom: 1px solid #ddd;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        
        .{0}-content-body {{
            padding: 15px;
        }}
        
        .{0}-toggle-button {{
            padding: 4px 8px;
            font-size: 12px;
            color: #fff;
            background-color: #007bff;
            border: 1px solid #007bff;
            border-radius: 3px;
            cursor: pointer;
        }}
        
        .{0}-alpha-filters {{
            margin-bottom: 10px;
        }}
        
        .{0}-alpha-title {{
            font-weight: bold;
            margin-bottom: 5px;
            font-size: 14px;
        }}
        
        .{0}-alpha-buttons {{
            display: flex;
            flex-wrap: wrap;
            gap: 2px;
        }}
        
        .{0}-alpha-btn {{
            padding: 4px 8px;
            font-size: 12px;
            border: 1px solid #dee2e6;
            border-radius: 3px;
            cursor: pointer;
            min-width: 42px;
            text-align: center;
            background-color: #f8f9fa;
            color: #6c757d;
        }}
        
        .{0}-alpha-btn.{0}-alpha-active {{
            background-color: #e6f2fa;
            color: #337ab7;
            border-color: #337ab7;
        }}
        
        .{0}-notification-indicator {{
            background-color: #e7f5ff;
            border: 1px solid #4dabf7;
            border-radius: 3px;
            padding: 8px 12px;
            margin-bottom: 10px;
            font-size: 12px;
        }}
        
        .{0}-back-link {{
            color: white;
            text-decoration: none;
            background-color: rgba(255,255,255,0.2);
            padding: 6px 10px;
            border-radius: 3px;
            font-size: 13px;
        }}
        
        .{0}-back-link:hover {{
            background-color: rgba(255,255,255,0.3);
            color: white;
            text-decoration: none;
        }}
        
        /* Mobile responsive adjustments */
        @media (max-width: 576px) {{
            .{0}-people-grid {{
                grid-template-columns: 1fr;
            }}
            
            .{0}-stats-bar {{
                flex-direction: column;
                text-align: center;
            }}
            
            .{0}-stat-item {{
                margin-bottom: 5px;
            }}
            
            .{0}-header-flex {{
                flex-direction: column;
                text-align: center;
            }}
            
            .{0}-alpha-buttons {{
                justify-content: center;
            }}
        }}
    </style>
    """.format(CSS_PREFIX)

# ::START:: Helper Classes
class MeetingInfo:
    """Class to store and manage meeting information"""
    # ::STEP:: Initialize meeting data structure
    def __init__(self, meeting_id, org_id, meeting_date, org_name, location, program_name=None):
        self.meeting_id = meeting_id
        self.org_id = org_id
        self.meeting_date = meeting_date
        self.org_name = org_name
        self.location = location
        self.program_name = program_name

class PersonInfo:
    """Class to store person information for check-in"""
    # ::STEP:: Initialize person data structure
    def __init__(self, people_id, name, family_id=None, org_ids=None, checked_in=False, age=None, subgroups=None, balance=0.0):
        self.people_id = people_id
        self.name = name
        self.family_id = family_id
        self.org_ids = org_ids or []
        self.checked_in = checked_in
        self.age = age
        self.subgroups = subgroups or {}
        self.balance = balance

# ::START:: Email Manager Class
class EmailManager:
    """Manages email notifications for check-ins"""
    
    def __init__(self, model, q):
        # ::STEP:: Initialize email manager
        self.model = model
        self.q = q
        self.email_queue = []
        
    def get_available_email_templates(self):
        """Get list of available email templates for parent notifications"""
        # ::STEP:: Query for email templates
        sql = """
            SELECT Name, Title
            FROM Content
            WHERE TypeID IN (2,7)
            AND (Name LIKE '%CheckedIn%') 
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
        except Exception as e:
            print "<!-- DEBUG: Error getting email templates: {0} -->".format(str(e))
            return [{'name': DEFAULT_EMAIL_TEMPLATE, 'title': 'Check-In Parent Notification'}]
    
    def get_parent_emails(self, people_id, age):
        """Get parent email addresses for a person under 18"""
        # ::STEP:: Check age and get parent emails
        if age >= 18:
            return []
            
        sql = """
            SELECT DISTINCT p2.EmailAddress, p2.Name, p2.PeopleId
            FROM People p1
            JOIN People p2 ON p1.FamilyId = p2.FamilyId
            WHERE p1.PeopleId = {0}
            AND p2.PositionInFamilyId IN (10, 20)
            AND p2.EmailAddress IS NOT NULL
            AND p2.EmailAddress != ''
            AND (p2.DoNotMailFlag IS NULL OR p2.DoNotMailFlag = 0)
        """.format(people_id)
        
        try:
            results = self.q.QuerySql(sql)
            parents = []
            for result in results:
                parents.append({
                    'email': result.EmailAddress,
                    'name': result.Name
                })
            return parents
        except Exception as e:
            print "<!-- DEBUG: Error getting parent emails: {0} -->".format(str(e))
            return []
    
    def send_parent_email(self, parent_email, parent_name, child_name, meeting_name, template_name):
        """Send email notification to a parent"""
        # ::STEP:: Send parent notification email
        try:
            sql = "SELECT TOP 1 PeopleId FROM People WHERE EmailAddress = '{0}'".format(parent_email.replace("'", "''"))
            result = self.q.QuerySqlTop1(sql)
            
            if not result or not hasattr(result, 'PeopleId'):
                return False
                
            parent_people_id = result.PeopleId
            
            if template_name and template_name != 'generic' and template_name != 'none':
                # Use TouchPoint email template
                try:
                    self.model.EmailContent(
                        "peopleids={0}".format(parent_people_id),
                        parent_people_id,
                        parent_email,
                        parent_name,
                        template_name
                    )
                    return True
                except Exception as e:
                    print "<!-- DEBUG: TouchPoint email failed: {0} -->".format(str(e))
            
            # Fallback to generic email
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
            return True
            
        except Exception as e:
            print "<!-- DEBUG: Parent email failed: {0} -->".format(str(e))
            return False
        
    def send_adult_email(self, person_email, person_name, meeting_name, template_name):
        """Send email notification to an adult who checked in"""
        # ::STEP:: Send adult notification email
        try:
            sql = "SELECT TOP 1 PeopleId FROM People WHERE EmailAddress = '{0}'".format(person_email.replace("'", "''"))
            result = self.q.QuerySqlTop1(sql)
            
            if not result or not hasattr(result, 'PeopleId'):
                return False
                
            person_people_id = result.PeopleId
            
            if template_name and template_name != 'generic' and template_name != 'none':
                try:
                    self.model.EmailContent(
                        "peopleids={0}".format(person_people_id),
                        person_people_id,
                        person_email,
                        person_name,
                        template_name
                    )
                    return True
                except Exception as e:
                    print "<!-- DEBUG: TouchPoint adult email failed: {0} -->".format(str(e))
            
            # Fallback to generic email
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
            return True
            
        except Exception as e:
            print "<!-- DEBUG: Adult email failed: {0} -->".format(str(e))
            return False
    
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
                    if self.send_adult_email(
                        email_data['person_email'],
                        email_data['person_name'],
                        email_data['meeting_name'],
                        email_data['template']
                    ):
                        sent_count += 1
                else:
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
                
        self.email_queue = []
        return sent_count

# ::START:: Check-In Manager Class
class CheckInManager:
    """Manages check-in operations and state"""
    
    def __init__(self, model, q):
        # ::STEP:: Initialize check-in manager
        self.model = model
        self.q = q
        self.today = self.model.DateTime.Date
        self.selected_meetings = []
        self.all_meetings_today = []
        self.last_check_in_time = None
        self.email_manager = EmailManager(model, q)
        
    def get_todays_meetings(self):
        """Get all meetings scheduled for today including program information"""
        # ::STEP:: Query today's meetings
        today = self.model.DateTime.Date
        
        sql = """
            SELECT DISTINCT m.MeetingId, m.OrganizationId, m.MeetingDate, o.OrganizationName, 
                   ISNULL(m.Location, '') as Location,
                   os.Program as ProgramName,
                   m.DidNotMeet,
                   o.OrganizationStatusId
            FROM Meetings m
            JOIN Organizations o ON m.OrganizationId = o.OrganizationId
            LEFT JOIN OrganizationStructure os ON m.OrganizationId = os.OrgId
            WHERE CONVERT(date, m.MeetingDate) = CONVERT(date, GETDATE())
            ORDER BY o.OrganizationName
        """
        
        try:
            sql_results = self.q.QuerySql(sql)
            
            meetings = []
            meeting_ids_seen = set()
            
            for result in sql_results:
                # Skip duplicates
                if result.MeetingId in meeting_ids_seen:
                    continue
                meeting_ids_seen.add(result.MeetingId)
                
                # Apply filters
                did_not_meet = getattr(result, 'DidNotMeet', None)
                org_status = getattr(result, 'OrganizationStatusId', None)
                
                # Skip if explicitly marked as "Did Not Meet"
                if did_not_meet == 1:
                    continue
                
                # Skip if organization is explicitly inactive (status 0)
                if org_status == 0:
                    continue
                
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
            
        except Exception as e:
            # Don't print debug info that could interfere with AJAX
            return []

    def handle_ajax_subgroups(check_in_manager, model, q):
        """Handle AJAX subgroup requests with clean JSON output"""
        import json
        
        try:
            # Get meeting IDs from form data
            meeting_ids = []
            if hasattr(model.Data, 'meeting_id'):
                if isinstance(model.Data.meeting_id, list):
                    meeting_ids = [str(m) for m in model.Data.meeting_id]
                else:
                    meeting_ids = [str(model.Data.meeting_id)]
            
            # Initialize meetings if needed - no debug output
            if not check_in_manager.all_meetings_today:
                check_in_manager.all_meetings_today = []
                today = model.DateTime.Date
                
                sql = """
                    SELECT DISTINCT m.MeetingId, m.OrganizationId, m.MeetingDate, o.OrganizationName, 
                           ISNULL(m.Location, '') as Location,
                           os.Program as ProgramName
                    FROM Meetings m
                    JOIN Organizations o ON m.OrganizationId = o.OrganizationId
                    LEFT JOIN OrganizationStructure os ON m.OrganizationId = os.OrgId
                    WHERE CONVERT(date, m.MeetingDate) = CONVERT(date, GETDATE())
                    AND (m.DidNotMeet IS NULL OR m.DidNotMeet = 0)
                    AND (o.OrganizationStatusId IS NULL OR o.OrganizationStatusId != 0)
                """
                
                results = q.QuerySql(sql)
                for result in results:
                    meeting = MeetingInfo(
                        result.MeetingId,
                        result.OrganizationId,
                        result.MeetingDate,
                        result.OrganizationName,
                        result.Location or "",
                        result.ProgramName if hasattr(result, 'ProgramName') else None
                    )
                    check_in_manager.all_meetings_today.append(meeting)
            
            # Get organizations for selected meetings
            org_ids = []
            for meeting_id in meeting_ids:
                for meeting in check_in_manager.all_meetings_today:
                    if str(meeting.meeting_id) == str(meeting_id):
                        org_ids.append(str(meeting.org_id))
                        break
            
            # Get unique subgroups
            all_subgroups = []
            seen_subgroups = set()
            
            for org_id in org_ids:
                sql = """
                    SELECT Id, Name
                    FROM MemberTags
                    WHERE OrgId = {0}
                    AND Name IS NOT NULL
                    AND Name != ''
                    ORDER BY Name
                """.format(org_id)
                
                results = q.QuerySql(sql)
                for result in results:
                    if hasattr(result, 'Name') and result.Name:
                        clean_name = str(result.Name).strip()
                        if clean_name and clean_name not in seen_subgroups:
                            seen_subgroups.add(clean_name)
                            all_subgroups.append(clean_name)
            
            # Sort alphabetically
            all_subgroups.sort()
            
            # Return ONLY JSON
            return json.dumps({"subgroups": all_subgroups})
            
        except Exception as e:
            # Return error as JSON
            return json.dumps({"error": str(e), "subgroups": []})
        
    def get_person_age(self, people_id):
        """Get a person's age"""
        # ::STEP:: Query person's age
        sql = """
            SELECT COALESCE(Age, DATEDIFF(year, BDate, GETDATE())) as CalculatedAge 
            FROM People 
            WHERE PeopleId = {0}
        """.format(people_id)
        
        try:
            result = self.q.QuerySqlTop1(sql)
            if result and hasattr(result, 'CalculatedAge') and result.CalculatedAge is not None:
                return int(result.CalculatedAge)
        except Exception as e:
            print "<!-- DEBUG: Error getting age for person {0}: {1} -->".format(people_id, str(e))
        return 99
    
    def get_people_by_filter(self, alpha_filter, search_term="", meeting_ids=None, page=1, page_size=PAGE_SIZE):
        """Get people by alpha filter or search term who are members of the selected orgs"""
        # ::STEP:: Filter and retrieve people with subgroup and balance support
        if not meeting_ids:
            return [], 0
        
        if not self.all_meetings_today:
            self.get_todays_meetings()
            
        # Get organization IDs from meeting IDs
        org_ids = []
        for meeting_id in meeting_ids:
            meeting_id_str = str(meeting_id)
            for m in self.all_meetings_today:
                if str(m.meeting_id) == meeting_id_str:
                    org_ids.append(str(m.org_id))
        
        # FIX: Check if org_ids is empty before proceeding
        if not org_ids:
            print "<!-- DEBUG: No org_ids found for meetings: {0} -->".format(meeting_ids)
            # Return empty result instead of proceeding with empty IN clause
            return [], 0
            
        org_ids_str = ",".join(org_ids)
        
        # Count query
        count_sql = """
            SELECT COUNT(DISTINCT p.PeopleId) AS TotalCount
            FROM People p
            JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
            WHERE om.OrganizationId IN ({0})
            AND (p.IsDeceased IS NULL OR p.IsDeceased = 0)
            AND (om.InactiveDate IS NULL OR om.InactiveDate > GETDATE())
        """.format(org_ids_str)
        
        # Add alpha filter
        if alpha_filter and alpha_filter != "All":
            first_letter = alpha_filter[0]
            last_letter = alpha_filter[2]
            count_sql += " AND p.LastName >= '{0}' AND p.LastName <= '{1}z'".format(first_letter, last_letter)
        
        # Add search filter
        if search_term:
            search_term_safe = search_term.replace("'", "''")
            count_sql += " AND (p.Name LIKE '%{0}%' OR p.FirstName LIKE '%{0}%' OR p.LastName LIKE '%{0}%')".format(search_term_safe)
        
        # Exclude already checked in
        count_sql += """
            AND NOT EXISTS (
                SELECT 1 FROM Attend a 
                WHERE a.PeopleId = p.PeopleId 
                AND a.OrganizationId IN ({0}) 
                AND CONVERT(date, a.MeetingDate) = CONVERT(date, GETDATE())
                AND a.AttendanceFlag = 1
            )
        """.format(org_ids_str)
        
        try:
            count_result = self.q.QuerySqlTop1(count_sql)
            total_count = count_result.TotalCount if hasattr(count_result, 'TotalCount') else 0
        except Exception as e:
            print "<!-- DEBUG: Error in count query: {0} -->".format(str(e))
            total_count = 0
        
        # Main query - same as count but with data retrieval
        sql = """
            SELECT DISTINCT p.PeopleId, p.Name, p.FamilyId, p.LastName, p.FirstName, p.Age
            FROM People p
            JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
            WHERE om.OrganizationId IN ({0})
            AND (p.IsDeceased IS NULL OR p.IsDeceased = 0)
            AND (om.InactiveDate IS NULL OR om.InactiveDate > GETDATE())
        """.format(org_ids_str)
        
        # Add same filters as count query
        if alpha_filter and alpha_filter != "All":
            first_letter = alpha_filter[0]
            last_letter = alpha_filter[2]
            sql += " AND p.LastName >= '{0}' AND p.LastName <= '{1}z'".format(first_letter, last_letter)
        
        if search_term:
            search_term_safe = search_term.replace("'", "''")
            sql += " AND (p.Name LIKE '%{0}%' OR p.FirstName LIKE '%{0}%' OR p.LastName LIKE '%{0}%')".format(search_term_safe)
        
        sql += """
            AND NOT EXISTS (
                SELECT 1 FROM Attend a 
                WHERE a.PeopleId = p.PeopleId 
                AND a.OrganizationId IN ({0}) 
                AND CONVERT(date, a.MeetingDate) = CONVERT(date, GETDATE())
                AND a.AttendanceFlag = 1
            )
        """.format(org_ids_str)
        
        sql += """ 
            ORDER BY p.LastName, p.FirstName
            OFFSET {0} ROWS
            FETCH NEXT {1} ROWS ONLY
        """.format((page - 1) * page_size, page_size)
        
        try:
            results = self.q.QuerySql(sql)
        except Exception as e:
            print "<!-- DEBUG: Error in main query: {0} -->".format(str(e))
            return [], total_count
        
        people = []
        for result in results:
            try:
                person_org_ids = self.get_person_org_ids(result.PeopleId, org_ids)
                age = result.Age if hasattr(result, 'Age') and result.Age is not None else 99
                
                # Get person's MemberTags for display (not filtering)
                person_subgroups = {}
                if ENABLE_SUBGROUP_FILTERING:
                    for org_id in person_org_ids:
                        person_subgroups[org_id] = self.get_person_subgroups(result.PeopleId, org_id)
                
                person_balance = 0.0
                if ENABLE_BALANCE_DISPLAY:
                    person_balance = self.get_person_balance(result.PeopleId, person_org_ids)
                
                person = PersonInfo(
                    result.PeopleId,
                    result.Name,
                    result.FamilyId,
                    person_org_ids,
                    False,
                    age,
                    person_subgroups,
                    person_balance
                )
                people.append(person)
            except Exception as e:
                print "<!-- DEBUG: Error processing person {0}: {1} -->".format(result.PeopleId, str(e))
                
        return people, total_count
            
    def get_people_by_meeting_ids(self, meeting_ids, alpha_filter="All", search_term="", page=1, page_size=PAGE_SIZE):
        """Alternative method to get people directly by meeting IDs when org mapping fails"""
        # ::STEP:: Get people by meeting IDs
        if not meeting_ids:
            return [], 0
        
        # Ensure we have valid meeting IDs
        unique_meeting_ids = [str(m) for m in meeting_ids if m]
        if not unique_meeting_ids:
            return [], 0
            
        meeting_ids_str = ",".join(unique_meeting_ids)
        
        # Count query
        count_sql = """
            SELECT COUNT(DISTINCT p.PeopleId) AS TotalCount
            FROM People p
            JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
            JOIN Meetings m ON m.OrganizationId = om.OrganizationId
            WHERE m.MeetingId IN ({0})
            AND (p.IsDeceased IS NULL OR p.IsDeceased = 0)
            AND (om.InactiveDate IS NULL OR om.InactiveDate > GETDATE())
        """.format(meeting_ids_str)
        
        if alpha_filter and alpha_filter != "All" and len(alpha_filter) >= 3:
            first_letter = alpha_filter[0]
            last_letter = alpha_filter[2]
            count_sql += " AND p.LastName >= '{0}' AND p.LastName <= '{1}z'".format(first_letter, last_letter)
        
        if search_term:
            search_term_safe = search_term.replace("'", "''")
            count_sql += " AND (p.Name LIKE '%{0}%' OR p.FirstName LIKE '%{0}%' OR p.LastName LIKE '%{0}%')".format(search_term_safe)
        
        count_sql += """
            AND NOT EXISTS (
                SELECT 1 FROM Attend a 
                WHERE a.PeopleId = p.PeopleId 
                AND a.MeetingId IN ({0})
                AND a.AttendanceFlag = 1
            )
        """.format(meeting_ids_str)
        
        try:
            count_result = self.q.QuerySqlTop1(count_sql)
            total_count = count_result.TotalCount if hasattr(count_result, 'TotalCount') else 0
        except Exception as e:
            print "<!-- DEBUG: Error in count query: {0} -->".format(str(e))
            total_count = 0
        
        # Main query
        sql = """
            SELECT DISTINCT p.PeopleId, p.Name, p.FamilyId, p.LastName, p.FirstName, p.Age
            FROM People p
            JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
            JOIN Meetings m ON m.OrganizationId = om.OrganizationId
            WHERE m.MeetingId IN ({0})
            AND (p.IsDeceased IS NULL OR p.IsDeceased = 0)
            AND (om.InactiveDate IS NULL OR om.InactiveDate > GETDATE())
        """.format(meeting_ids_str)
        
        if alpha_filter and alpha_filter != "All" and len(alpha_filter) >= 3:
            first_letter = alpha_filter[0]
            last_letter = alpha_filter[2]
            sql += " AND p.LastName >= '{0}' AND p.LastName <= '{1}z'".format(first_letter, last_letter)
        
        if search_term:
            search_term_safe = search_term.replace("'", "''")
            sql += " AND (p.Name LIKE '%{0}%' OR p.FirstName LIKE '%{0}%' OR p.LastName LIKE '%{0}%')".format(search_term_safe)
        
        sql += """
            AND NOT EXISTS (
                SELECT 1 FROM Attend a 
                WHERE a.PeopleId = p.PeopleId 
                AND a.MeetingId IN ({0})
                AND a.AttendanceFlag = 1
            )
        """.format(meeting_ids_str)
        
        sql += """ 
            ORDER BY p.LastName, p.FirstName
            OFFSET {0} ROWS
            FETCH NEXT {1} ROWS ONLY
        """.format((page - 1) * page_size, page_size)
        
        try:
            results = self.q.QuerySql(sql)
        except Exception as e:
            print "<!-- DEBUG: Error in people query: {0} -->".format(str(e))
            return [], total_count
        
        # Get org mapping
        org_ids_by_meeting = {}
        if results:
            try:
                meeting_to_org_sql = """
                    SELECT MeetingId, OrganizationId 
                    FROM Meetings 
                    WHERE MeetingId IN ({0})
                """.format(meeting_ids_str)
                meeting_org_results = self.q.QuerySql(meeting_to_org_sql)
                
                for result in meeting_org_results:
                    org_ids_by_meeting[str(result.MeetingId)] = str(result.OrganizationId)
            except Exception as e:
                print "<!-- DEBUG: Error getting org IDs: {0} -->".format(str(e))
        
        people = []
        for result in results:
            try:
                all_org_ids = org_ids_by_meeting.values()
                person_org_ids = self.get_person_org_ids(result.PeopleId, all_org_ids)
                age = result.Age if hasattr(result, 'Age') and result.Age is not None else 99
                
                if person_org_ids:
                    person = PersonInfo(
                        result.PeopleId,
                        result.Name,
                        result.FamilyId,
                        person_org_ids,
                        False,
                        age
                    )
                    people.append(person)
            except Exception as e:
                print "<!-- DEBUG: Error processing person: {0} -->".format(str(e))
                
        return people, total_count
    
    def get_person_org_ids(self, people_id, org_ids):
        """Get the organization IDs that this person is a member of"""
        # ::STEP:: Query person's organization memberships
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
        
        try:
            results = self.q.QuerySql(sql)
            return [str(result.OrganizationId) for result in results]
        except Exception as e:
            print "<!-- DEBUG: Error getting person org IDs: {0} -->".format(str(e))
            return []
    
    def check_in_person(self, people_id, meeting_id):
        """Simplified direct check-in using just people_id and meeting_id"""
        # ::STEP:: Perform check-in
        try:
            people_id_int = int(people_id)
            meeting_id_int = int(meeting_id)
            
            # Check if already checked in
            pre_check_sql = """
                SELECT AttendanceFlag
                FROM Attend 
                WHERE PeopleId = {0} AND MeetingId = {1}
                AND AttendanceFlag = 1
            """.format(people_id_int, meeting_id_int)
            
            pre_result = self.q.QuerySqlTop1(pre_check_sql)
            already_checked_in = pre_result is not None
            
            # Call TouchPoint API
            result = self.model.EditPersonAttendance(meeting_id_int, people_id_int, True)
            
            # Verify check-in was successful
            post_check_sql = """
                SELECT AttendanceFlag, CreatedDate
                FROM Attend 
                WHERE PeopleId = {0} AND MeetingId = {1}
                AND AttendanceFlag = 1
            """.format(people_id_int, meeting_id_int)
            
            post_result = self.q.QuerySqlTop1(post_check_sql)
            now_checked_in = post_result is not None
            
            return now_checked_in
            
        except Exception as e:
            print "<!-- DEBUG: Error in check_in_person: {0} -->".format(str(e))
            return False
    
    def check_in_person_with_email(self, people_id, meeting_id, person_name, email_template):
        """Check in a person and send appropriate email notification"""
        # ::STEP:: Check-in with notification
        try:
            # Perform check-in
            success = self.check_in_person(people_id, meeting_id)
            
            if not success:
                return False
                
            if not email_template or email_template == 'none':
                return success
                
            # Get person details for email
            sql = """
                SELECT 
                    PeopleId,
                    Name,
                    Age,
                    EmailAddress,
                    FamilyId,
                    DATEDIFF(year, BDate, GETDATE()) AS CalculatedAge
                FROM People 
                WHERE PeopleId = {0}
            """.format(people_id)
            
            person_result = self.q.QuerySqlTop1(sql)
            
            if not person_result:
                return success
                
            age = 99
            if hasattr(person_result, 'Age') and person_result.Age is not None:
                age = int(person_result.Age)
            elif hasattr(person_result, 'CalculatedAge') and person_result.CalculatedAge is not None:
                age = int(person_result.CalculatedAge)
                
            person_email = person_result.EmailAddress if hasattr(person_result, 'EmailAddress') else None
            
            # Get meeting name
            meeting_name = "the event"
            for meeting in self.all_meetings_today:
                if str(meeting.meeting_id) == str(meeting_id):
                    meeting_name = meeting.org_name
                    break
            
            # Send appropriate email
            if age < 18:
                # Send to parents
                parents = self.email_manager.get_parent_emails(people_id, age)
                for parent in parents:
                    self.email_manager.send_parent_email(
                        parent['email'],
                        parent['name'],
                        person_name,
                        meeting_name,
                        email_template
                    )
            else:
                # Send to adult
                if person_email:
                    self.email_manager.send_adult_email(
                        person_email,
                        person_name,
                        meeting_name,
                        email_template
                    )
            
            return success
            
        except Exception as e:
            print "<!-- DEBUG: Error in check_in_person_with_email: {0} -->".format(str(e))
            return False
    
    def get_check_in_stats(self, meeting_ids):
        """Get check-in statistics for selected meetings with forced refresh"""
        # ::STEP:: Calculate statistics
        if not meeting_ids:
            return {"checked_in": 0, "not_checked_in": 0, "total": 0}
            
        if not self.all_meetings_today:
            self.get_todays_meetings()
            
        meeting_ids_str = ",".join([str(m) for m in meeting_ids])
        
        # Get organization IDs
        sql = """
            SELECT DISTINCT OrganizationId
            FROM Meetings
            WHERE MeetingId IN ({0})
        """.format(meeting_ids_str)
        
        try:
            org_results = self.q.QuerySql(sql)
            org_ids = [str(r.OrganizationId) for r in org_results]
        except Exception as e:
            print "<!-- DEBUG: Error getting org IDs for stats: {0} -->".format(str(e))
            return {"checked_in": 0, "not_checked_in": 0, "total": 0}
        
        if not org_ids:
            return {"checked_in": 0, "not_checked_in": 0, "total": 0}
            
        org_ids_str = ",".join(org_ids)
        
        # Get total members
        sql = """
            SELECT COUNT(DISTINCT p.PeopleId) AS TotalCount
            FROM People p
            JOIN OrganizationMembers om ON p.PeopleId = om.PeopleId
            WHERE om.OrganizationId IN ({0})
            AND (p.IsDeceased IS NULL OR p.IsDeceased = 0)
            AND (om.InactiveDate IS NULL OR om.InactiveDate > GETDATE())
        """.format(org_ids_str)
        
        try:
            total_result = self.q.QuerySqlTop1(sql)
            total = total_result.TotalCount if hasattr(total_result, 'TotalCount') else 0
        except Exception as e:
            print "<!-- DEBUG: Error getting total count: {0} -->".format(str(e))
            total = 0
        
        # Get checked in count
        sql = """
            SELECT COUNT(DISTINCT a.PeopleId) AS CheckedInCount
            FROM Attend a
            WHERE a.MeetingId IN ({0})
            AND a.AttendanceFlag = 1
            AND CONVERT(date, a.MeetingDate) = CONVERT(date, GETDATE())
        """.format(meeting_ids_str)
        
        try:
            checked_in_result = self.q.QuerySqlTop1(sql)
            checked_in = checked_in_result.CheckedInCount if hasattr(checked_in_result, 'CheckedInCount') else 0
        except Exception as e:
            print "<!-- DEBUG: Error getting checked in count: {0} -->".format(str(e))
            checked_in = 0
        
        return {
            "checked_in": checked_in,
            "not_checked_in": total - checked_in,
            "total": total
        }

    def get_checked_in_people(self, meeting_ids, page=1, page_size=PAGE_SIZE):
        """Get people who have already checked in"""
        # ::STEP:: Query checked-in people
        if not meeting_ids:
            return [], 0
            
        meeting_ids_str = ",".join([str(m) for m in meeting_ids])
        
        # Count query
        count_sql = """
            SELECT COUNT(DISTINCT p.PeopleId) AS TotalCount
            FROM People p
            JOIN Attend a ON p.PeopleId = a.PeopleId
            WHERE a.MeetingId IN ({0})
            AND a.AttendanceFlag = 1
            AND CONVERT(date, a.MeetingDate) = CONVERT(date, GETDATE())
        """.format(meeting_ids_str)
        
        try:
            count_result = self.q.QuerySqlTop1(count_sql)
            total_count = count_result.TotalCount if hasattr(count_result, 'TotalCount') else 0
        except Exception as e:
            print "<!-- DEBUG: Error getting checked in count: {0} -->".format(str(e))
            total_count = 0
        
        # Main query
        sql = """
            SELECT DISTINCT p.PeopleId, p.Name, p.FamilyId, p.LastName, p.FirstName, p.Age
            FROM People p
            JOIN Attend a ON p.PeopleId = a.PeopleId
            WHERE a.MeetingId IN ({0})
            AND a.AttendanceFlag = 1
            AND CONVERT(date, a.MeetingDate) = CONVERT(date, GETDATE())
            ORDER BY p.LastName, p.FirstName
            OFFSET {1} ROWS
            FETCH NEXT {2} ROWS ONLY
        """.format(meeting_ids_str, (page - 1) * page_size, page_size)
        
        try:
            results = self.q.QuerySql(sql)
        except Exception as e:
            print "<!-- DEBUG: Error getting checked in people: {0} -->".format(str(e))
            return [], total_count
        
        # Get org IDs
        org_sql = """
            SELECT DISTINCT OrganizationId
            FROM Meetings
            WHERE MeetingId IN ({0})
        """.format(meeting_ids_str)
        
        try:
            org_results = self.q.QuerySql(org_sql)
            org_ids = [str(r.OrganizationId) for r in org_results]
        except Exception as e:
            print "<!-- DEBUG: Error getting org IDs: {0} -->".format(str(e))
            org_ids = []
        
        people = []
        for result in results:
            try:
                person_org_ids = self.get_person_attended_org_ids(result.PeopleId, org_ids)
                age = result.Age if hasattr(result, 'Age') and result.Age is not None else 99
                
                person = PersonInfo(
                    result.PeopleId,
                    result.Name,
                    result.FamilyId,
                    person_org_ids,
                    True,
                    age
                )
                people.append(person)
            except Exception as e:
                print "<!-- DEBUG: Error processing checked in person: {0} -->".format(str(e))
            
        return people, total_count
    
    def get_person_attended_org_ids(self, people_id, org_ids):
        """Get the organization IDs that this person has attended today"""
        # ::STEP:: Query person's attendance today
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
        
        try:
            results = self.q.QuerySql(sql)
            return [str(result.OrganizationId) for result in results]
        except Exception as e:
            print "<!-- DEBUG: Error getting attended org IDs: {0} -->".format(str(e))
            return []

    # ::START:: Subgroup Management Methods (Updated for MemberTags)
    def get_org_subgroups(self, org_id):
        """Get all subgroups (MemberTags) for an organization"""
        # ::STEP:: Query organization subgroups from MemberTags table
        sql = """
            SELECT Id, Name
            FROM MemberTags
            WHERE OrgId = {0}
            AND Name IS NOT NULL
            AND Name != ''
            ORDER BY Name
        """.format(org_id)
        
        try:
            results = self.q.QuerySql(sql)
            subgroups = []
            for result in results:
                if hasattr(result, 'Name') and result.Name:
                    # Clean the name
                    clean_name = str(result.Name).strip()
                    if clean_name:
                        subgroups.append({
                            'id': result.Id,
                            'name': clean_name
                        })
            # Remove debug output that could interfere with JSON
            # print "<!-- DEBUG: Found {0} MemberTags for org {1} -->".format(len(subgroups), org_id)
            return subgroups
        except Exception as e:
            # Don't print debug output during AJAX calls
            # print "<!-- DEBUG: Error querying MemberTags: {0} -->".format(str(e))
            return []
    
    def get_person_subgroups(self, people_id, org_id):
        """Get subgroups (MemberTags) this person belongs to for a specific organization"""
        # ::STEP:: Query person's subgroups from OrgMemMemTags
        sql = """
            SELECT mt.Id, mt.Name, ommt.IsLeader
            FROM OrgMemMemTags ommt
            JOIN MemberTags mt ON ommt.MemberTagId = mt.Id
            WHERE ommt.PeopleId = {0}
            AND ommt.OrgId = {1}
            ORDER BY mt.Name
        """.format(people_id, org_id)
        
        try:
            results = self.q.QuerySql(sql)
            subgroups = []
            for result in results:
                subgroup_name = result.Name
                if hasattr(result, 'IsLeader') and result.IsLeader:
                    subgroup_name += " (Leader)"
                subgroups.append(subgroup_name)
            return subgroups
        except Exception as e:
            print "<!-- DEBUG: Error getting person MemberTags: {0} -->".format(str(e))
            return []
    
    def get_enabled_subgroups(self, meeting_ids):
        """Get enabled subgroups for selected meetings"""
        # ::STEP:: Get enabled subgroups from form data - these are display filters, not membership filters
        enabled_subgroups = getattr(self.model.Data, 'enabled_subgroups', [])
        if isinstance(enabled_subgroups, str):
            enabled_subgroups = enabled_subgroups.split(',')
        return [sg.strip() for sg in enabled_subgroups if sg.strip()]

    # ::START:: Balance Management Methods  
    def get_person_balance(self, people_id, org_ids):
        """Get outstanding balance for a person across selected organizations"""
        # ::STEP:: Query outstanding balances from TransactionSummary
        if not org_ids:
            return 0.0
            
        org_ids_str = ",".join([str(oid) for oid in org_ids])
        
        sql = """
            SELECT ISNULL(SUM(ts.IndDue), 0) AS Balance
            FROM TransactionSummary ts
            WHERE ts.PeopleId = {0}
            AND ts.OrganizationId IN ({1})
            AND ts.IndDue <> 0
            AND ts.IsLatestTransaction = 1
        """.format(people_id, org_ids_str)
        
        try:
            result = self.q.QuerySqlTop1(sql)
            if result and hasattr(result, 'Balance'):
                return float(result.Balance) if result.Balance else 0.0
        except Exception as e:
            print "<!-- DEBUG: Error getting balance for person {0}: {1} -->".format(people_id, str(e))
        
        return 0.0
    
    def format_currency(self, amount):
        """Format amount as currency"""
        # ::STEP:: Format currency display
        if amount == 0:
            return ""
        return "${0:,.2f}".format(abs(amount))

# ::START:: Rendering Functions
def render_meeting_selection(check_in_manager):
    """Render the meeting selection page with email template selection"""
    # ::STEP:: Display meeting selection UI
    meetings = check_in_manager.get_todays_meetings()
    email_templates = check_in_manager.email_manager.get_available_email_templates()
    script_name = get_script_name()
    
    # ::STEP:: Get all subgroups for all organizations
    all_org_subgroups = {}
    if ENABLE_SUBGROUP_FILTERING:
        for meeting in meetings:
            org_id = str(meeting.org_id)
            if org_id not in all_org_subgroups:
                subgroups = check_in_manager.get_org_subgroups(org_id)
                all_org_subgroups[org_id] = [sg['name'] for sg in subgroups]
    
    # Convert to JSON for JavaScript
    import json
    org_subgroups_json = json.dumps(all_org_subgroups)
    
    # FIX: Get program filtering options - prevent duplicates
    programs = set()
    for meeting in meetings:
        if hasattr(meeting, 'program_name') and meeting.program_name:
            program_name = str(meeting.program_name).strip()
            if program_name:  # Only add non-empty program names
                programs.add(program_name)
    
    programs = sorted(list(programs))
    selected_email_template = getattr(check_in_manager.model.Data, 'email_template', DEFAULT_EMAIL_TEMPLATE)
    
    # Get selected alpha block count
    selected_alpha_count = getattr(check_in_manager.model.Data, 'alpha_block_count', DEFAULT_ALPHA_BLOCK_COUNT)
    try:
        selected_alpha_count = int(selected_alpha_count)
        if selected_alpha_count not in ALPHA_BLOCKS_CONFIG:
            selected_alpha_count = DEFAULT_ALPHA_BLOCK_COUNT
    except:
        selected_alpha_count = DEFAULT_ALPHA_BLOCK_COUNT
    
    # Start HTML with scoped CSS
    print """<!DOCTYPE html>
<html>
<head>
    <title>FastLane Check-In System</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    {0}
    <script>
        // Define global helper functions
        function getPyScriptAddress() {{
            var path = window.location.pathname;
            return path.replace("/PyScript/", "/PyScriptForm/");
        }}
        function showLoading() {{
            document.getElementById('loadingIndicator').style.display = 'block';
        }}
    </script>
</head>
<body>
    <div class="{1}-container" style="padding: 20px;">
        <div class="{1}-loading" id="loadingIndicator">
            <div class="{1}-loading-spinner"></div>
            <div>Loading...</div>
        </div>
        
        <!-- Header -->
        <div class="{1}-header">
            <div class="{1}-header-flex">
                <div>
                    <h2 class="{1}-header-title">FastLane Check-In</h2>
                    <div class="{1}-header-tagline">Life moves pretty fast. So do our lines!</div>
                </div>
            </div>
        </div>
    """.format(get_scoped_css(), CSS_PREFIX)
    
    if not meetings:
        print """
        <div class="{0}-alert {0}-alert-warning">
            <h4>No Meetings Today</h4>
            <p>There are no active meetings scheduled for today.</p>
            <p>Things to check:</p>
            <ul>
                <li>Make sure meetings are created for today's date</li>
                <li>Check that organizations hosting the meetings are Active</li>
                <li>Verify the meetings are not marked as "Did Not Meet"</li>
            </ul>
        </div>
        </div>
        </body>
        </html>
        """.format(CSS_PREFIX)
        return
    
    # Main meeting selection form
    print """
    <div class="{0}-well">
        <form method="post" action="" onsubmit="this.action = getPyScriptAddress(); showLoading();">
            <input type="hidden" name="step" value="check_in">
            
            <!-- Program Filtering Section -->
    """.format(CSS_PREFIX)
    
    # Program filtering - client-side only
    if programs:
        print """
            <div class="{0}-feature-section">
                <h4 class="{0}-feature-title">Program Filter</h4>
                <div class="{0}-form-group">
                    <label for="program_filter">Filter by Program:</label>
                    <select class="{0}-form-control" id="program_filter" name="program_filter" onchange="updateMeetingDisplay()">
                        <option value="">All Programs</option>
        """.format(CSS_PREFIX)
        
        for program in programs:
            print '<option value="{0}">{0}</option>'.format(program)
        
        print """
                    </select>
                </div>
            </div>
        """.format(CSS_PREFIX)
    
    print """
            <!-- Email Template Selection Section -->
            <div class="{0}-feature-section">
                <h4 class="{0}-feature-title">Parent Email Notifications</h4>
                <div class="{0}-form-group">
                    <label for="email_template">Email Template for Parents (when checking in children under 18):</label>
                    <select class="{0}-form-control" id="email_template" name="email_template">
                        <option value="none" {2}>No Email Notification</option>
                        <option value="generic" {3}>Generic Email (System Generated)</option>
    """.format(CSS_PREFIX, script_name, 
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
                    <div class="{0}-help-text">
                        Notifications will be sent automatically when checking in:
                        <ul style="margin:5px 0; padding-left:20px;">
                            <li>For children (under 18): Email sent to parents/guardians</li>
                            <li>For adults (18+): Email sent directly to them</li>
                        </ul>
                    </div>                      
                </div>
            </div>
            
            <!-- Alpha Filter Configuration Section -->
            <div class="{0}-feature-section">
                <h4 class="{0}-feature-title">Name Filter Configuration</h4>
                <div class="{0}-form-group">
                    <label for="alpha_block_count">Number of alphabet groups for filtering:</label>
                    <select class="{0}-form-control" id="alpha_block_count" name="alpha_block_count">
    """.format(CSS_PREFIX)
    
    # Add options for each configuration
    for count in sorted(ALPHA_BLOCKS_CONFIG.keys()):
        blocks = ALPHA_BLOCKS_CONFIG[count][:-1]  # Exclude "All" from display
        preview = ", ".join(blocks)
        selected = 'selected="selected"' if count == selected_alpha_count else ''
        print '<option value="{0}" {2}>{0} groups ({1})</option>'.format(count, preview, selected)
    
    print """
                    </select>
                    <div class="{0}-help-text">
                        Choose fewer groups for smaller congregations or more groups for larger ones.
                    </div>                      
                </div>
            </div>
            
            <!-- Subgroup Selection Section -->
            <div class="{0}-feature-section" style="display: {1};">
                <h4 class="{0}-feature-title">Subgroup Display Filter (Optional)</h4>
                <p class="{0}-help-text">
                    Select which subgroups you want to highlight when displayed on each person's card. 
                    This does NOT filter who appears in the list - it only controls which subgroup badges are shown.
                </p>
                <div id="subgroup-container">
                    <p style="color:#666; font-style:italic;">Select meetings above to see available subgroups.</p>
                </div>
            </div>
            
            <h4>Available Meetings for {2}</h4>
            
            <p>
                <button type="button" class="{0}-btn {0}-btn-sm {0}-btn-primary" id="selectAllBtn">Select All</button>
                <button type="button" class="{0}-btn {0}-btn-sm" id="deselectAllBtn">Deselect All</button>
            </p>
            
            <div class="{0}-row" id="meetings-container" style="max-height: 400px; overflow-y: auto;">
    """.format(
        CSS_PREFIX,
        "block" if ENABLE_SUBGROUP_FILTERING else "none",
        check_in_manager.today.ToString(DATE_FORMAT)
    )
    
    # Generate meetings display - ALL meetings, no server-side filtering
    meetings_by_org_id = {}  # Track meetings by org ID to prevent duplicates
    
    for meeting in meetings:
        # Use org_id as the key to prevent duplicates
        org_key = str(meeting.org_id)
        if org_key not in meetings_by_org_id:
            meetings_by_org_id[org_key] = {
                'name': meeting.org_name,
                'program': meeting.program_name if hasattr(meeting, 'program_name') and meeting.program_name else '',
                'meetings': []
            }
        meetings_by_org_id[org_key]['meetings'].append(meeting)
    
    # Now render the meetings
    for org_id, org_data in sorted(meetings_by_org_id.items(), key=lambda x: x[1]['name']):
        org_name = org_data['name']
        org_program = org_data['program']
        org_meetings = org_data['meetings']
        
        # Add data-program attribute for client-side filtering
        print """
            <div class="{0}-col-6" data-program="{2}">
                <div class="{0}-panel">
                    <div class="{0}-panel-heading">
                        <h4 class="{0}-panel-title">{1}</h4>
                    </div>
                    <div class="{0}-panel-body">
        """.format(CSS_PREFIX, org_name, org_program)
        
        # Sort meetings by time and deduplicate
        seen_times = set()
        for meeting in sorted(org_meetings, key=lambda m: m.meeting_date):
            meeting_time = meeting.meeting_date.ToString("h:mm tt")
            meeting_key = "{0}_{1}".format(meeting.meeting_id, meeting_time)
            
            if meeting_key not in seen_times:
                seen_times.add(meeting_key)
                print """
                            <div class="{0}-checkbox">
                                <label>
                                    <input type="checkbox" name="meeting_id" value="{1}" class="meeting-checkbox" data-org-id="{2}">
                                    {3} at {4} {5}
                                </label>
                            </div>
                """.format(
                    CSS_PREFIX,
                    meeting.meeting_id,
                    meeting.org_id,
                    org_name,
                    meeting_time,
                    "({0})".format(meeting.location) if meeting.location else ""
                )
        
        print """
                    </div>
                </div>
            </div>
        """.format(CSS_PREFIX)
    
    print """
            </div>
            
            <div class="{0}-form-group" style="margin-top: 20px;">
                <button type="submit" class="{0}-btn {0}-btn-lg {0}-btn-success" id="start-checkin-btn">
                    <i class="fa fa-check-circle"></i> Start Check-In
                </button>
            </div>
        </form>
    </div>
    
    <script>
        // Pre-loaded subgroups data
        var orgSubgroups = {1};
        
        // Show loading indicator during form submission
        function showLoading() {{
            document.getElementById('loadingIndicator').style.display = 'block';
        }}
        
        // Enable/disable the start button based on selections
        function updateStartButton() {{
            var checkedCount = document.querySelectorAll('.meeting-checkbox:checked').length;
            document.getElementById('start-checkin-btn').disabled = checkedCount === 0;
        }}
        
        // Update meetings display when program filter changes - CLIENT SIDE ONLY
        function updateMeetingDisplay() {{
            var programFilter = document.getElementById('program_filter').value;
            var allPanels = document.querySelectorAll('[data-program]');
            
            allPanels.forEach(function(panel) {{
                if (programFilter === '' || panel.getAttribute('data-program') === programFilter) {{
                    panel.style.display = 'block';
                }} else {{
                    panel.style.display = 'none';
                }}
            }});
            
            // Uncheck any hidden meetings
            var hiddenCheckboxes = document.querySelectorAll('[data-program]:not([style*="display: block"]) .meeting-checkbox');
            hiddenCheckboxes.forEach(function(cb) {{
                cb.checked = false;
            }});
            
            // Update the start button state
            updateStartButton();
            
            // Update subgroups if enabled
            if ({2}) {{
                updateSubgroups();
            }}
        }}
        
        // Display subgroups for selection
        function displaySubgroups(subgroups) {{
            var container = document.getElementById('subgroup-container');
            if (!container) return;
            
            if (subgroups.length === 0) {{
                container.innerHTML = '<p style="color:#666; font-style:italic;">No subgroups found for selected meetings.</p>';
                return;
            }}
            
            var html = '<div style="margin-top:10px;"><strong>Available Subgroups (for display):</strong></div>';
            html += '<div class="{0}-subgroup-grid">';
            
            subgroups.forEach(function(subgroupName) {{
                html += '<label class="{0}-subgroup-item">';
                html += '<input type="checkbox" name="enabled_subgroups" value="' + subgroupName + '" style="margin-right:8px;">';
                html += '<span>' + subgroupName + '</span>';
                html += '</label>';
            }});
            
            html += '</div>';
            container.innerHTML = html;
        }}
        
        // Update subgroups when meetings change
        function updateSubgroups() {{
            if (!{2}) {{
                return;
            }}
            
            var selectedOrgIds = [];
            var checkboxes = document.querySelectorAll('.meeting-checkbox:checked');
            checkboxes.forEach(function(cb) {{
                var orgId = cb.getAttribute('data-org-id');
                if (orgId && selectedOrgIds.indexOf(orgId) === -1) {{
                    selectedOrgIds.push(orgId);
                }}
            }});
            
            // Collect all unique subgroups from selected organizations
            var allSubgroups = [];
            var seenSubgroups = {{}};
            
            selectedOrgIds.forEach(function(orgId) {{
                var subgroups = orgSubgroups[orgId] || [];
                subgroups.forEach(function(subgroup) {{
                    if (!seenSubgroups[subgroup]) {{
                        seenSubgroups[subgroup] = true;
                        allSubgroups.push(subgroup);
                    }}
                }});
            }});
            
            // Sort alphabetically
            allSubgroups.sort();
            
            // Display the subgroups
            if (allSubgroups.length > 0) {{
                displaySubgroups(allSubgroups);
            }} else {{
                document.getElementById('subgroup-container').innerHTML = '<p style="color:#666; font-style:italic;">No subgroups found for selected meetings.</p>';
            }}
        }}
        
        // Set up event listeners after the DOM is loaded
        document.addEventListener('DOMContentLoaded', function() {{
            // Set up checkbox listeners
            var checkboxes = document.querySelectorAll('.meeting-checkbox');
            
            checkboxes.forEach(function(checkbox) {{
                checkbox.addEventListener('change', function() {{
                    updateStartButton();
                    if ({2}) {{
                        updateSubgroups();
                    }}
                }});
            }});
            
            // Set up Select All button
            var selectAllBtn = document.getElementById('selectAllBtn');
            if (selectAllBtn) {{
                selectAllBtn.addEventListener('click', function() {{
                    // Only select visible checkboxes
                    var visibleCheckboxes = document.querySelectorAll('[data-program]:not([style*="display: none"]) .meeting-checkbox');
                    visibleCheckboxes.forEach(function(checkbox) {{
                        checkbox.checked = true;
                    }});
                    updateStartButton();
                    if ({2}) {{
                        updateSubgroups();
                    }}
                }});
            }}
            
            // Set up Deselect All button
            var deselectAllBtn = document.getElementById('deselectAllBtn');
            if (deselectAllBtn) {{
                deselectAllBtn.addEventListener('click', function() {{
                    checkboxes.forEach(function(checkbox) {{
                        checkbox.checked = false;
                    }});
                    updateStartButton();
                    if ({2}) {{
                        updateSubgroups();
                    }}
                }});
            }}
            
            // Initial check
            updateStartButton();
            if ({2}) {{
                // Don't auto-update on load since no meetings are selected
                // updateSubgroups();
            }}
        }});
    </script>
</div>
</body>
</html>
""".format(CSS_PREFIX, org_subgroups_json, "true" if ENABLE_SUBGROUP_FILTERING else "false")

def create_pagination(page, total_pages, script_name, meeting_ids, view_mode, alpha_filter, search_term, alpha_block_count=DEFAULT_ALPHA_BLOCK_COUNT):
    """Create a pagination control with scoped CSS"""
    # ::STEP:: Generate pagination controls
    if total_pages <= 1:
        return ""
        
    # Create hidden inputs for meeting IDs
    meeting_id_inputs = ""
    for meeting_id in meeting_ids:
        meeting_id_inputs += '<input type="hidden" name="meeting_id" value="{0}">'.format(meeting_id)
    
    # Get email template selection
    email_template = getattr(model.Data, 'email_template', 'none')
    email_template_input = '<input type="hidden" name="email_template" value="{0}">'.format(email_template)
    
    # Add alpha block count input
    alpha_block_count_input = '<input type="hidden" name="alpha_block_count" value="{0}">'.format(alpha_block_count)
    
    # Build pagination HTML with scoped styling
    pagination = ['<div class="{0}-pagination">'.format(CSS_PREFIX)]
    
    # Previous button
    if page > 1:
        pagination.append(
            '<form method="post" action="" onsubmit="this.action = getPyScriptAddress(); showLoading();" style="display:inline-block; margin:0 2px;">'
            '<input type="hidden" name="step" value="check_in">'
            '<input type="hidden" name="page" value="{1}">'
            '<input type="hidden" name="view_mode" value="{2}">'
            '<input type="hidden" name="alpha_filter" value="{3}">'
            '<input type="hidden" name="search_term" value="{4}">'
            '{5}'
            '{6}'
            '{8}'
            '<button type="submit" class="{7}-pagination-btn">&laquo;</button>'
            '</form>'.format(
                script_name, page - 1, view_mode, alpha_filter, search_term, 
                meeting_id_inputs, email_template_input, CSS_PREFIX, alpha_block_count_input
            )
        )
    else:
        pagination.append(
            '<span class="{0}-pagination-btn {0}-pagination-disabled">&laquo;</span>'.format(CSS_PREFIX)
        )
        
    # Page numbers
    start_page = max(1, page - 2)
    end_page = min(total_pages, page + 2)
    
    # Always show page 1
    if start_page > 1:
        pagination.append(
            '<form method="post" action="" onsubmit="this.action = getPyScriptAddress(); showLoading();" style="display:inline-block; margin:0 2px;">'
            '<input type="hidden" name="step" value="check_in">'
            '<input type="hidden" name="page" value="1">'
            '<input type="hidden" name="view_mode" value="{1}">'
            '<input type="hidden" name="alpha_filter" value="{2}">'
            '<input type="hidden" name="search_term" value="{3}">'
            '{4}'
            '{5}'
            '{7}'
            '<button type="submit" class="{6}-pagination-btn">1</button>'
            '</form>'.format(
                script_name, view_mode, alpha_filter, search_term, 
                meeting_id_inputs, email_template_input, CSS_PREFIX, alpha_block_count_input
            )
        )
        if start_page > 2:
            pagination.append(
                '<span class="{0}-pagination-btn">...</span>'.format(CSS_PREFIX)
            )
            
    # Page links
    for i in range(start_page, end_page + 1):
        if i == page:
            pagination.append(
                '<span class="{0}-pagination-btn {0}-pagination-current">{1}</span>'.format(CSS_PREFIX, i)
            )
        else:
            pagination.append(
                '<form method="post" action="" onsubmit="this.action = getPyScriptAddress(); showLoading();" style="display:inline-block; margin:0 2px;">'
                '<input type="hidden" name="step" value="check_in">'
                '<input type="hidden" name="page" value="{1}">'
                '<input type="hidden" name="view_mode" value="{2}">'
                '<input type="hidden" name="alpha_filter" value="{3}">'
                '<input type="hidden" name="search_term" value="{4}">'
                '{5}'
                '{6}'
                '{8}'
                '<button type="submit" class="{7}-pagination-btn">{1}</button>'
                '</form>'.format(
                    script_name, i, view_mode, alpha_filter, search_term, 
                    meeting_id_inputs, email_template_input, CSS_PREFIX, alpha_block_count_input
                )
            )
            
    # Always show last page
    if end_page < total_pages:
        if end_page < total_pages - 1:
            pagination.append(
                '<span class="{0}-pagination-btn">...</span>'.format(CSS_PREFIX)
            )
        pagination.append(
            '<form method="post" action="" onsubmit="this.action = getPyScriptAddress(); showLoading();" style="display:inline-block; margin:0 2px;">'
            '<input type="hidden" name="step" value="check_in">'
            '<input type="hidden" name="page" value="{1}">'
            '<input type="hidden" name="view_mode" value="{2}">'
            '<input type="hidden" name="alpha_filter" value="{3}">'
            '<input type="hidden" name="search_term" value="{4}">'
            '{5}'
            '{6}'
            '{8}'
            '<button type="submit" class="{7}-pagination-btn">{1}</button>'
            '</form>'.format(
                script_name, total_pages, view_mode, alpha_filter, search_term, 
                meeting_id_inputs, email_template_input, CSS_PREFIX, alpha_block_count_input
            )
        )
        
    # Next button
    if page < total_pages:
        pagination.append(
            '<form method="post" action="" onsubmit="this.action = getPyScriptAddress(); showLoading();" style="display:inline-block; margin:0 2px;">'
            '<input type="hidden" name="step" value="check_in">'
            '<input type="hidden" name="page" value="{1}">'
            '<input type="hidden" name="view_mode" value="{2}">'
            '<input type="hidden" name="alpha_filter" value="{3}">'
            '<input type="hidden" name="search_term" value="{4}">'
            '{5}'
            '{6}'
            '{8}'
            '<button type="submit" class="{7}-pagination-btn">&raquo;</button>'
            '</form>'.format(
                script_name, page + 1, view_mode, alpha_filter, search_term, 
                meeting_id_inputs, email_template_input, CSS_PREFIX, alpha_block_count_input
            )
        )
    else:
        pagination.append(
            '<span class="{0}-pagination-btn {0}-pagination-disabled">&raquo;</span>'.format(CSS_PREFIX)
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
        
        # Validate parameters
        if not person_id or not meeting_id:
            print '{{"success":false,"reason":"missing_parameters"}}'
            return True
            
        # Clean any comma-separated values
        if ',' in str(person_id):
            person_id = str(person_id).split(',')[0].strip()
            
        if ',' in str(meeting_id):
            meeting_id = str(meeting_id).split(',')[0].strip()
        
        # Convert to integers
        person_id_int = int(person_id)
        meeting_id_int = int(meeting_id)
        
        # Perform the check-in with email notification
        success = check_in_manager.check_in_person_with_email(
            person_id_int, 
            meeting_id_int, 
            person_name, 
            email_template
        )
        
        # Force refresh of statistics
        check_in_manager.last_check_in_time = check_in_manager.model.DateTime
        
        # Process any queued emails if batch mode
        if PARENT_EMAIL_DELAY:
            sent_count = check_in_manager.email_manager.send_queued_emails()
        
        if success:
            print '{{"success":true}}'
        else:
            print '{{"success":false,"reason":"check_in_failed"}}'
            
        return True
            
    except Exception as e:
        print "<!-- ERROR AJAX: {0} -->".format(str(e))
        print '{{"success":false,"reason":"general_error"}}'
        return True

def render_fastlane_check_in(check_in_manager):
    """Render the FastLane check-in page with enhanced features and scoped CSS"""
    # ::STEP:: Display check-in UI with subgroup and balance support
    script_name = get_script_name()
    
    # Get parameters
    meeting_ids = []
    if hasattr(check_in_manager.model.Data, 'meeting_id'):
        if isinstance(check_in_manager.model.Data.meeting_id, list):
            meeting_ids = [str(m) for m in check_in_manager.model.Data.meeting_id]
        else:
            if ',' in str(check_in_manager.model.Data.meeting_id):
                meeting_ids = str(check_in_manager.model.Data.meeting_id).split(',')
            else:
                meeting_ids = [str(check_in_manager.model.Data.meeting_id)]
            
    # Get other parameters
    alpha_filter = getattr(check_in_manager.model.Data, 'alpha_filter', 'All')
    search_term = getattr(check_in_manager.model.Data, 'search_term', '')
    view_mode = getattr(check_in_manager.model.Data, 'view_mode', 'not_checked_in')
    email_template = getattr(check_in_manager.model.Data, 'email_template', 'none')
    
    # Get alpha block count
    alpha_block_count = getattr(check_in_manager.model.Data, 'alpha_block_count', DEFAULT_ALPHA_BLOCK_COUNT)
    try:
        alpha_block_count = int(alpha_block_count)
        if alpha_block_count not in ALPHA_BLOCKS_CONFIG:
            alpha_block_count = DEFAULT_ALPHA_BLOCK_COUNT
    except:
        alpha_block_count = DEFAULT_ALPHA_BLOCK_COUNT
    
    # Handle page
    try:
        page_value = getattr(check_in_manager.model.Data, 'page', 1)
        current_page = int(page_value) if page_value else 1
    except (ValueError, TypeError):
        current_page = 1

    # Force a stats refresh
    check_in_manager.last_check_in_time = check_in_manager.model.DateTime
    
    # Get check-in statistics with forced refresh
    stats = check_in_manager.get_check_in_stats(meeting_ids)
    
    # Create inputs for meeting IDs
    meeting_id_inputs = ""
    for meeting_id in meeting_ids:
        meeting_id_inputs += '<input type="hidden" name="meeting_id" value="{0}">'.format(meeting_id)
    
    # Add email template input
    email_template_input = '<input type="hidden" name="email_template" value="{0}">'.format(email_template)
    
    # Add alpha block count input
    alpha_block_count_input = '<input type="hidden" name="alpha_block_count" value="{0}">'.format(alpha_block_count)
    
    # Add subgroup inputs if enabled
    subgroup_inputs = ""
    if ENABLE_SUBGROUP_FILTERING:
        enabled_subgroups = check_in_manager.get_enabled_subgroups(meeting_ids)
        for subgroup in enabled_subgroups:
            subgroup_inputs += '<input type="hidden" name="enabled_subgroups" value="{0}">'.format(subgroup)
    
    # Process form submission and prepare flash message
    flash_message = ""
    flash_name = ""
    if hasattr(check_in_manager.model.Data, 'action'):
        action = check_in_manager.model.Data.action
        
        if action == 'single_direct_check_in':
            person_id = getattr(check_in_manager.model.Data, 'person_id', None)
            meeting_id = getattr(check_in_manager.model.Data, 'meeting_id', None)
            person_name = getattr(check_in_manager.model.Data, 'person_name', "")
            
            if person_id and meeting_id:
                if ',' in str(meeting_id):
                    meeting_id = str(meeting_id).split(',')[0].strip()
                
                try:
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
                        
                        if PARENT_EMAIL_DELAY:
                            sent_count = check_in_manager.email_manager.send_queued_emails()
                            if sent_count > 0:
                                flash_message += " (parent notified)"
                except Exception as e:
                    flash_message = "Error checking in: {0}".format(str(e))
        
        elif action == 'remove_check_in' and hasattr(check_in_manager.model.Data, 'person_id'):
            person_id = check_in_manager.model.Data.person_id
            meeting_id = getattr(check_in_manager.model.Data, 'meeting_id', None)
            person_name = getattr(check_in_manager.model.Data, 'person_name', "")
            
            if person_id and meeting_id:
                try:
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
    pagination = create_pagination(current_page, total_pages, script_name, meeting_ids, view_mode, alpha_filter, search_term, alpha_block_count)
    
    alpha_blocks = ALPHA_BLOCKS_CONFIG.get(alpha_block_count, ALPHA_BLOCKS_CONFIG[DEFAULT_ALPHA_BLOCK_COUNT])
    
    alpha_filters_html = """
    <div class="{0}-alpha-filters">
        <div class="{0}-alpha-title">Filter:</div>
        <div class="{0}-alpha-buttons">
    """.format(CSS_PREFIX)
    
    for alpha_block in alpha_blocks:
        is_active = alpha_block == alpha_filter
        active_class = '{0}-alpha-active'.format(CSS_PREFIX) if is_active else ''
        
        alpha_filters_html += """
        <form method="post" action="" onsubmit="this.action = getPyScriptAddress(); showLoading();" style="display:inline-block; margin:2px;">
            <input type="hidden" name="step" value="check_in">
            <input type="hidden" name="view_mode" value="{1}">
            <input type="hidden" name="alpha_filter" value="{2}">
            <input type="hidden" name="search_term" value="{3}">
            <input type="hidden" name="alpha_block_count" value="{9}">
            {4}
            {5}
            {6}
            <button type="submit" class="{7}-alpha-btn {8}">{2}</button>
        </form>
        """.format(
            script_name,        # {0}
            view_mode,          # {1}
            alpha_block,        # {2}
            search_term,        # {3}
            meeting_id_inputs,  # {4}
            email_template_input, # {5}
            subgroup_inputs,    # {6}
            CSS_PREFIX,         # {7}
            active_class,       # {8}
            alpha_block_count   # {9}
        )
    
    alpha_filters_html += """
        </div>
    </div>
    """.format(CSS_PREFIX)
    
    # Generate the people list HTML with scoped CSS
    people_list_html = []
    
    people_list_html.append("""
    <div class="{0}-people-grid">
    """.format(CSS_PREFIX))
    
    for person in people:
        # Get the organizations this person is a member of
        person_org_names = []
        for org_id in person.org_ids:
            for meeting in check_in_manager.all_meetings_today:
                if str(meeting.org_id) == str(org_id):
                    if meeting.org_name not in person_org_names:
                        person_org_names.append(meeting.org_name)
                    break
        
        # Create a representation of the organizations this person belongs to
        org_display = ""
        if person_org_names:
            org_display = '<div class="{0}-person-details">{1}</div>'.format(CSS_PREFIX, ', '.join(person_org_names))
        
        # Show age indicator for minors
        age_indicator = ""
        if hasattr(person, 'age') and person.age is not None and person.age < 18:
            age_indicator = '<span class="{0}-age-indicator">(Age: {1})</span>'.format(CSS_PREFIX, person.age)
        
        # Show balance indicator if enabled and person has balance
        balance_indicator = ""
        if (ENABLE_BALANCE_DISPLAY and hasattr(person, 'balance') and 
            person.balance > BALANCE_WARNING_THRESHOLD):
            balance_indicator = '<span class="{0}-balance-indicator">${1:.2f}</span>'.format(CSS_PREFIX, person.balance)
        
        # Show subgroups if enabled and person has subgroups
        subgroup_display = ""
        if (ENABLE_SUBGROUP_FILTERING and hasattr(person, 'subgroups') and 
            person.subgroups):
            all_subgroups = []
            for org_id, subgroups in person.subgroups.items():
                all_subgroups.extend(subgroups)
            if all_subgroups:
                displayed_subgroups = list(set(all_subgroups))[:SUBGROUP_DISPLAY_LIMIT]
                more_text = ""
                if len(all_subgroups) > SUBGROUP_DISPLAY_LIMIT:
                    more_text = " (+{0} more)".format(len(all_subgroups) - SUBGROUP_DISPLAY_LIMIT)
                subgroup_display = '<div class="{0}-subgroup-display">Groups: {1}{2}</div>'.format(
                    CSS_PREFIX, ', '.join(displayed_subgroups), more_text)
        
        # For disambiguation in case of duplicate names
        unique_identifier = ""
        if person.people_id:
            people_id_str = str(person.people_id)
            last_digits = people_id_str[-4:] if len(people_id_str) >= 4 else people_id_str
            unique_identifier = '<span class="{0}-unique-id"> (ID: {1})</span>'.format(CSS_PREFIX, last_digits)
        
        # Determine background color based on balance
        card_classes = "{0}-person-card".format(CSS_PREFIX)
        if (ENABLE_BALANCE_DISPLAY and hasattr(person, 'balance') and 
            person.balance > BALANCE_WARNING_THRESHOLD):
            card_classes += " {0}-balance-warning".format(CSS_PREFIX)
                
        if view_mode == "checked_in":
            # For checked-in people, show remove button
            attended_meeting_ids = []
            
            try:
                sql = """
                    SELECT DISTINCT MeetingId
                    FROM Attend
                    WHERE PeopleId = {0}
                    AND CONVERT(date, MeetingDate) = CONVERT(date, GETDATE())
                    AND AttendanceFlag = 1
                """.format(person.people_id)
                
                results = check_in_manager.q.QuerySql(sql)
                
                for result in results:
                    attended_meeting_id = str(result.MeetingId)
                    if attended_meeting_id in meeting_ids:
                        attended_meeting_ids.append(attended_meeting_id)
            except Exception as e:
                print "<!-- DEBUG: Error getting attended meeting IDs: {0} -->".format(str(e))
            
            if not attended_meeting_ids and meeting_ids:
                attended_meeting_ids = [meeting_ids[0]]
                
            item_html = """
            <div id="person-{0}" class="{1}">
                <div class="{9}-person-card-flex">
                    <div class="{9}-person-info">
                        <div class="{9}-person-name">{2}{3}{4}{5}</div>
                        {6}
                        {7}
                    </div>
            """.format(
                person.people_id,       # {0}
                card_classes,           # {1}
                person.name,            # {2}
                unique_identifier,      # {3}
                age_indicator,          # {4}
                balance_indicator,      # {5}
                org_display,            # {6}
                subgroup_display,       # {7}
                "",                     # {8}
                CSS_PREFIX              # {9}
            )
            
            meeting_id = attended_meeting_ids[0] if attended_meeting_ids else (meeting_ids[0] if meeting_ids else "")
                
            item_html += """
                <form method="post" action="" onsubmit="this.action = getPyScriptAddress(); showLoading();" class="{10}-person-buttons">
                    <input type="hidden" name="step" value="check_in">
                    <input type="hidden" name="action" value="remove_check_in">
                    <input type="hidden" name="person_id" value="{0}">
                    <input type="hidden" name="person_name" value="{1}">
                    <input type="hidden" name="meeting_id" value="{6}">
                    <input type="hidden" name="view_mode" value="checked_in">
                    <input type="hidden" name="alpha_filter" value="{7}">
                    <input type="hidden" name="search_term" value="{8}">
                    <input type="hidden" name="page" value="{9}">
                    {11}
                    {12}
                    {13}
                    <button type="submit" class="{10}-btn {10}-btn-sm {10}-btn-danger">Remove</button>
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
                CSS_PREFIX,             # {10}
                email_template_input,   # {11}
                subgroup_inputs,        # {12}
                alpha_block_count_input # {13}
            )
            
            item_html += """
                </div>
            </div>
            """

        else:
            # For not-checked-in people, show check-in button
            item_html = """
            <div id="person-{0}" class="{1}">
                <div class="{7}-person-card-flex">
                    <div class="{7}-person-info">
                        <div class="{7}-person-name">{2}{3}{4}{5}</div>
                        {6}
                        {8}
                    </div>
                    <div class="{7}-person-buttons">
            """.format(
                person.people_id,       # {0}
                card_classes,           # {1}
                person.name,            # {2}
                unique_identifier,      # {3}
                age_indicator,          # {4}
                balance_indicator,      # {5}
                org_display,            # {6}
                CSS_PREFIX,             # {7}
                subgroup_display        # {8}
            )
        
            person_meeting_ids = []
            
            for org_id in person.org_ids:
                for meeting in check_in_manager.all_meetings_today:
                    if (str(meeting.org_id) == str(org_id) and 
                        str(meeting.meeting_id) in meeting_ids):
                        person_meeting_ids.append(str(meeting.meeting_id))
            
            if not person_meeting_ids and meeting_ids:
                person_meeting_ids = [meeting_ids[0]]
            
            for meeting_id in person_meeting_ids:
                meeting_name = "Check In"
                for meeting in check_in_manager.all_meetings_today:
                    if str(meeting.meeting_id) == str(meeting_id):
                        if len(person_meeting_ids) > 1:
                            meeting_name = meeting.org_name.split(' ')[0]
                        break
                
                item_html += """
                <script type="text/javascript">
                    // Define helper functions in global scope
                    function getPyScriptAddress() {{
                        var path = window.location.pathname;
                        return path.replace("/PyScript/", "/PyScriptForm/");
                    }}
                    
                    function showLoading() {{
                        var loadingEl = document.getElementById('loadingIndicator');
                        if (loadingEl) {{
                            loadingEl.style.display = 'block';
                        }}
                    }}
                </script>
                <button onclick="(function(personId, meetingId, personName, emailTemplate) {{
                    var personCard = document.getElementById('person-' + personId);
                    if (personCard) {{
                        personCard.classList.add('{5}-processing');
                    }}
                    
                    var miniFlash = document.createElement('div');
                    miniFlash.className = '{5}-mini-flash';
                    miniFlash.innerHTML = '<span>Processing...</span>';
                    document.body.appendChild(miniFlash);
                    
                    var formData = new FormData();
                    formData.append('step', 'check_in');
                    formData.append('action', 'ajax_check_in');
                    formData.append('person_id', personId);
                    formData.append('person_name', personName);
                    formData.append('meeting_id', meetingId);
                    formData.append('email_template', emailTemplate);
                    
                    var xhr = new XMLHttpRequest();
                    xhr.open('POST', getPyScriptAddress(), true);
                    
                    xhr.onload = function() {{
                        if (xhr.status === 200) {{
                            if (personCard) {{
                                personCard.style.display = 'none';
                            }}
                            
                            miniFlash.innerHTML = '<span>' + personName + ' checked in!</span>';
                            
                            var checkedInEl = document.getElementById('stat-checked-in');
                            var notCheckedInEl = document.getElementById('stat-not-checked-in');
                            
                            if (checkedInEl && notCheckedInEl) {{
                                var checkedIn = parseInt(checkedInEl.innerText || '0');
                                var notCheckedIn = parseInt(notCheckedInEl.innerText || '0');
                                
                                checkedInEl.innerText = (checkedIn + 1).toString();
                                notCheckedInEl.innerText = Math.max(0, notCheckedIn - 1).toString();
                            }}
                        }} else {{
                            miniFlash.style.backgroundColor = 'rgba(220,53,69,0.9)';
                            miniFlash.innerHTML = '<span>Error checking in ' + personName + '</span>';
                            
                            if (personCard) {{
                                personCard.classList.remove('{5}-processing');
                            }}
                        }}
                        
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
                }})('{0}', '{6}', '{2}', '{10}'); return false;" class="{5}-btn {5}-btn-sm {5}-btn-success">{9}</button>
                """.format(
                    person.people_id,  # {0}
                    "",                # {1}
                    person.name.replace("'", "\\'").replace('"', '\\"'),  # {2}
                    "",                # {3}
                    script_name,       # {4}
                    CSS_PREFIX,        # {5}
                    meeting_id,        # {6}
                    alpha_filter,      # {7}
                    search_term,       # {8}
                    meeting_name,      # {9}
                    email_template     # {10}
                )
            
            item_html += """
                    </div>
                </div>
            </div>
            """
        
        people_list_html.append(item_html)
    
    people_list_html.append("</div>")
    
    # Create the compact stats bar with scoped CSS
    stats_bar = """
    <div class="{0}-stats-bar">
        <div class="{0}-stat-item">
            <span class="{0}-stat-number {0}-stat-checked" id="stat-checked-in">{1}</span>
            <div>Checked In</div>
        </div>
        <div class="{0}-stat-item">
            <span class="{0}-stat-number {0}-stat-not-checked" id="stat-not-checked-in">{2}</span>
            <div>Not Checked In</div>
        </div>
        <div class="{0}-stat-item">
            <span class="{0}-stat-number {0}-stat-total" id="stat-total">{3}</span>
            <div>Total</div>
        </div>
    </div>
    """.format(CSS_PREFIX, stats["checked_in"], stats["not_checked_in"], stats["total"])
    
    # Create flash message for 2-second display with scoped CSS
    flash_html = ""
    if flash_message and flash_name:
        flash_html = """
        <div id="flash-message" class="{0}-flash-message">
            <strong>{1}</strong> {2}
        </div>
        <script>
            setTimeout(function() {{
                var flashMessage = document.getElementById('flash-message');
                if (flashMessage) {{
                    flashMessage.style.opacity = '0';
                    setTimeout(function() {{
                        flashMessage.style.display = 'none';
                    }}, 500);
                }}
            }}, 2000);
        </script>
        """.format(CSS_PREFIX, flash_name, flash_message)
    
    # Create subgroup filter display if enabled
    subgroup_filter_display = ""
    if ENABLE_SUBGROUP_FILTERING:
        enabled_subgroups = check_in_manager.get_enabled_subgroups(meeting_ids)
        if enabled_subgroups:
            subgroup_filter_display = '<div class="{0}-notification-indicator"><strong>Subgroup Filter:</strong> {1}</div>'.format(
                CSS_PREFIX, ', '.join(enabled_subgroups))
    
    # HTML with enhanced features and scoped CSS
    print """<!DOCTYPE html>
<html>
<head>
    <title>FastLane Check-In</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    {0}
</head>
<body>
    <div class="{1}-container" style="padding: 10px;">
        <div class="{1}-loading" id="loadingIndicator">
            <div class="{1}-loading-spinner"></div>
            <div>Processing...</div>
        </div>
        
        <!-- Flash message for check-in confirmation -->
        {2}
        
        <!-- Updated FastLane header with tagline -->
        <div class="{1}-header">
            <div class="{1}-header-flex">
                <div>
                    <h2 class="{1}-header-title">FastLane Check-In</h2>
                    <div class="{1}-header-tagline">Life moves pretty fast. So do our lines!</div>
                </div>
                <a href="#" onclick="window.location.href = window.location.pathname.replace('/PyScriptForm/', '/PyScript/') + '?step=choose_meetings'; return false;" class="{1}-back-link">Back</a>
            </div>
        </div>
        
        <!-- Compact stats bar -->
        {4}
        
        <!-- Main content -->
        <div class="{1}-content-panel">
            <!-- Header with toggle in the title bar -->
            <div class="{1}-content-header">
                <h3 style="margin:0; font-size:16px; flex-grow:1;">{5}</h3>
                
                <!-- Toggle switch for view mode -->
                <div style="display:flex; align-items:center;">
                    <form method="post" action="" onsubmit="this.action = getPyScriptAddress(); showLoading();" style="margin:0; padding:0;">
                        <input type="hidden" name="step" value="check_in">
                        <input type="hidden" name="view_mode" value="{7}">
                        <input type="hidden" name="alpha_filter" value="{6}">
                        <input type="hidden" name="search_term" value="{8}">
                        {9}
                        {13}
                        {15}
                        {18}
                        <button type="submit" class="{1}-toggle-button">
                            Switch to {10}
                        </button>
                    </form>
                </div>
            </div>
            
            <div class="{1}-content-body">
                <!-- Search box -->
                <form method="post" action="" onsubmit="this.action = getPyScriptAddress(); showLoading();" class="{1}-search-container">
                    <input type="hidden" name="step" value="check_in">
                    <input type="hidden" name="view_mode" value="{5}">
                    {9}
                    {13}
                    {15}
                    {18}
                    <div class="{1}-search-flex">
                        <input type="text" name="search_term" value="{8}" placeholder="Search by name..." class="{1}-search-input">
                        <button type="submit" class="{1}-search-btn">
                            Search
                        </button>
                    </div>
                </form>
                
                <!-- Alpha filters - compact -->
                {11}
                
                <!-- Parent email notification indicator -->
                {12}
                
                <!-- Subgroup filter indicator -->
                {14}
                
                <!-- People grid layout with fixed alignment -->
                {16}
                
                <!-- Pagination -->
                <div style="text-align:center;">
                    {17}
                </div>
            </div>
        </div>
        
        <script>
            function showLoading() {{
                document.getElementById('loadingIndicator').style.display = 'block';
            }}
        </script>
    </div>
</body>
</html>
""".format(
        get_scoped_css(),         # {0}
        CSS_PREFIX,               # {1}
        flash_html,               # {2}
        script_name,              # {3}
        stats_bar,                # {4}
        "People to Check In" if view_mode == "not_checked_in" else "Checked In People", # {5}
        alpha_filter,             # {6}
        "checked_in" if (view_mode == "not_checked_in" or view_mode == "") else "not_checked_in", # {7}
        search_term,              # {8}
        meeting_id_inputs,        # {9}
        "Checked In List" if (view_mode == "not_checked_in" or view_mode == "") else "Check-In Page", # {10}
        alpha_filters_html,       # {11}
        '<div class="{0}-notification-indicator"><strong>Parent Notifications:</strong> {1}</div>'.format(CSS_PREFIX, email_template if email_template != 'none' else 'Disabled') if email_template != 'none' else '', # {12}
        email_template_input,     # {13}
        subgroup_filter_display,  # {14}
        subgroup_inputs,          # {15}
        "\n".join(people_list_html), # {16}
        pagination,                # {17}
        alpha_block_count_input   # {18}
    )
    
    # Send any queued parent emails if in batch mode
    if PARENT_EMAIL_DELAY and check_in_manager.email_manager.email_queue:
        sent_count = check_in_manager.email_manager.send_queued_emails()
        if sent_count > 0:
            print "<!-- DEBUG: Sent {0} parent notification emails -->".format(sent_count)
    
    return True

# ::START:: Updated Main Function
def main():
    """Main entry point for the FastLane Check-In system"""
    
    # CRITICAL: Handle AJAX subgroup request FIRST with no other output
    step = getattr(model.Data, 'step', '')
    if not step and hasattr(model, 'Request') and hasattr(model.Request, 'QueryString'):
        # Check GET parameters
        query_string = str(model.Request.QueryString)
        if 'step=get_subgroups' in query_string:
            step = 'get_subgroups'
    
    if step == 'get_subgroups':
        # For now, just return empty subgroups embedded in the page
        # This ensures it works even if TouchPoint wraps it in HTML
        print '<div id="subgroups-json" style="display:none;">{"subgroups":[]}</div>'
        return
    
    # Continue with rest of the application...
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

        # ::STEP:: Handle AJAX requests
        if clean_action == 'ajax_check_in':
            ajax_handled = process_ajax_check_in(check_in_manager)
            if ajax_handled:
                return
        
        # ::STEP:: Process other form actions
        if clean_action == 'single_direct_check_in':
            if clean_person_id and clean_meeting_id:
                try:
                    person_name = getattr(model.Data, 'person_name', '')
                    result = check_in_manager.check_in_person_with_email(
                        int(clean_person_id), 
                        int(clean_meeting_id), 
                        person_name,
                        email_template
                    )
                    check_in_manager.last_check_in_time = check_in_manager.model.DateTime
                    
                    if PARENT_EMAIL_DELAY:
                        check_in_manager.email_manager.send_queued_emails()
                except Exception as e:
                    print "<div style='color:red;'>Error: " + str(e) + "</div>"
        
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
                    
                    check_in_manager.last_check_in_time = check_in_manager.model.DateTime
                    
                    if not success:
                        print "<div style='color:red;'>Some check-in removals failed</div>"
                except Exception as e:
                    print "<div style='color:red;'>Error: " + str(e) + "</div>"
        
        # ::STEP:: Render appropriate page
        if clean_step == 'check_in':
            render_fastlane_check_in(check_in_manager)
        else:
            render_meeting_selection(check_in_manager)
            
    except Exception as e:
        # ::STEP:: Handle errors gracefully
        import traceback
        print """
        <div style="padding: 20px; font-family: Arial, sans-serif;">
            <h2 style="color: #d9534f;">FastLane Check-In Error</h2>
            <div style="background-color: #f2dede; border: 1px solid #ebccd1; color: #a94442; padding: 15px; border-radius: 4px; margin: 10px 0;">
                <p><strong>An error occurred:</strong> {0}</p>
            </div>
            <details style="margin: 15px 0;">
                <summary style="cursor: pointer; font-weight: bold; color: #337ab7;">Show Technical Details</summary>
                <pre style="background-color: #f5f5f5; padding: 10px; border-radius: 4px; overflow-x: auto; margin-top: 10px;">{1}</pre>
            </details>
            <div style="margin-top: 20px;">
                <a href="/PyScript/{2}" style="background-color: #337ab7; color: white; padding: 8px 16px; text-decoration: none; border-radius: 4px;">Back to Check-In</a>
            </div>
        </div>
        """.format(str(e), traceback.format_exc(), get_script_name())

# ::START:: Script Entry Point
if __name__ == "__main__":
    main()
else:
    main()
